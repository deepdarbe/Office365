<#
.SYNOPSIS
    Aktif O365 kullanicilarini, tanimli kurallara gore hedef gruplara
    otomatik olarak uye yapar.

.DESCRIPTION
    "group-mapping.json" dosyasindaki kurallari kullanir:
        - defaultGroups : Her aktif kullanicinin uye olmasi gereken gruplar
        - rules         : Field/operator/value eslesmesine gore ek gruplar

    Operator destegi: equals, notEquals, startsWith, endsWith, contains, regex

    Eger kullanici (UPN) son calismadan bu yana state dosyasinda yoksa
    "yeni kullanici" olarak isaretlenir ve uyelik atanir. Bu sayede yeni
    eklenen kullanicilar otomatik olarak ilgili gruplara dahil olur.

    Hassas veri yazmamak icin state ve log dosyalari "data/" klasoru altinda
    tutulur ve .gitignore tarafindan haric tutulur.

.PARAMETER MappingPath
    Mapping JSON dosyasinin yolu. Varsayilan: ./group-mapping.json

.PARAMETER WhatIf
    Degisiklik yapmadan kuru calisma yapar.

.PARAMETER OnlyNewUsers
    Sadece state dosyasinda olmayan (yeni) kullanicilara uygular.

.EXAMPLE
    . .\Connect-O365Services.ps1; Connect-O365
    .\Sync-UserGroupMembership.ps1 -OnlyNewUsers
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $MappingPath = (Join-Path $PSScriptRoot 'group-mapping.json'),
    [switch] $OnlyNewUsers
)

if (-not (Get-MgContext)) {
    throw "Microsoft Graph oturumu yok. Once Connect-O365 calistirin."
}

if (-not (Test-Path $MappingPath)) {
    throw "Mapping dosyasi bulunamadi: $MappingPath`n  template'i kopyalayin: group-mapping.template.json -> group-mapping.json"
}

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$statePath = Join-Path $dataDir 'known-users.json'
$logPath   = Join-Path $dataDir ("sync-{0:yyyyMMdd-HHmm}.log" -f (Get-Date))

function Write-Log {
    param([string] $Message, [string] $Level = 'INFO')
    $line = "{0:yyyy-MM-dd HH:mm:ss} [{1}] {2}" -f (Get-Date), $Level, $Message
    Add-Content -Path $logPath -Value $line
    $color = switch ($Level) { 'WARN' { 'Yellow' } 'ERROR' { 'Red' } 'OK' { 'Green' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
}

function Test-RuleMatch {
    param($User, $Rule)

    $field    = $Rule.match.field
    $op       = $Rule.match.operator
    $expected = [string]$Rule.match.value
    $actual   = [string]$User.$field

    switch ($op) {
        'equals'     { return ($actual -eq $expected) }
        'notEquals'  { return ($actual -ne $expected) }
        'startsWith' { return $actual.StartsWith($expected, [StringComparison]::OrdinalIgnoreCase) }
        'endsWith'   { return $actual.EndsWith($expected, [StringComparison]::OrdinalIgnoreCase) }
        'contains'   { return $actual -like "*$expected*" }
        'regex'      { return $actual -match $expected }
        default      { return $false }
    }
}

# 1) Mapping ve state yukle
$mapping = Get-Content $MappingPath -Raw | ConvertFrom-Json
$known   = @{}
if (Test-Path $statePath) {
    $raw = Get-Content $statePath -Raw | ConvertFrom-Json
    foreach ($entry in $raw) { $known[$entry.UserPrincipalName] = $entry }
}

Write-Log "Sync basladi. Mapping: $MappingPath"

# 2) Aktif kullanicilari getir
$users = & (Join-Path $PSScriptRoot 'Get-ActiveUsers.ps1')

# 3) Hedef gruplari onceden cek (isim -> id)
$groupCache = @{}
function Get-GroupByName {
    param([string] $Name)
    if ($groupCache.ContainsKey($Name)) { return $groupCache[$Name] }
    $g = Get-MgGroup -Filter "displayName eq '$($Name -replace "'", "''")'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
    $groupCache[$Name] = $g
    return $g
}

$processed = 0
$added     = 0
$skippedExisting = 0

foreach ($user in $users) {
    $isNew = -not $known.ContainsKey($user.UserPrincipalName)
    if ($OnlyNewUsers -and -not $isNew) { continue }

    $targetGroups = New-Object System.Collections.Generic.HashSet[string]
    foreach ($g in $mapping.defaultGroups) { [void]$targetGroups.Add($g) }
    foreach ($rule in $mapping.rules) {
        if (Test-RuleMatch -User $user -Rule $rule) {
            foreach ($g in $rule.groups) { [void]$targetGroups.Add($g) }
        }
    }

    if ($targetGroups.Count -eq 0) {
        Write-Log "Kural eslesmedi: $($user.UserPrincipalName)" 'WARN'
        $known[$user.UserPrincipalName] = [pscustomobject]@{
            UserPrincipalName = $user.UserPrincipalName
            FirstSeen         = (Get-Date).ToString('o')
        }
        continue
    }

    foreach ($groupName in $targetGroups) {
        $group = Get-GroupByName -Name $groupName
        if (-not $group) {
            Write-Log "Grup bulunamadi: '$groupName' (kullanici: $($user.UserPrincipalName))" 'ERROR'
            continue
        }

        $existing = Get-MgGroupMember -GroupId $group.Id -All -ErrorAction SilentlyContinue |
            Where-Object { $_.Id -eq $user.Id }

        if ($existing) {
            $skippedExisting++
            continue
        }

        $target = "{0} -> {1}" -f $user.UserPrincipalName, $groupName
        if ($PSCmdlet.ShouldProcess($target, 'Add-MgGroupMember')) {
            try {
                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction Stop
                Write-Log "Eklendi: $target" 'OK'
                $added++
            } catch {
                Write-Log "HATA: $target -> $($_.Exception.Message)" 'ERROR'
            }
        }
    }

    $known[$user.UserPrincipalName] = [pscustomobject]@{
        UserPrincipalName = $user.UserPrincipalName
        FirstSeen         = if ($isNew) { (Get-Date).ToString('o') } else { $known[$user.UserPrincipalName].FirstSeen }
        LastSync          = (Get-Date).ToString('o')
    }
    $processed++
}

# 4) State'i yaz
$known.Values | ConvertTo-Json -Depth 4 | Set-Content -Path $statePath -Encoding UTF8

Write-Log ("Bitti. Islenen: {0}, Eklenen uyelik: {1}, Mevcut atlanmis: {2}" -f $processed, $added, $skippedExisting) 'OK'
