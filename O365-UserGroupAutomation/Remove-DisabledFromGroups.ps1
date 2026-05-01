<#
.SYNOPSIS
    Belirtilen gruplarda disable (AccountEnabled=false) ya da silinmis
    kullanicilari tespit eder ve cikarir.

.DESCRIPTION
    Kuru calisma (-WhatIf) ile once liste uretir, sonra onayinizla
    cikarma yapar. Grup tipine gore (Distribution / MailEnabledSecurity
    / Microsoft365 / Security) dogru cmdlet'i kullanir.

    Cikti CSV'leri data/ altina yazilir (gitignore'lu).

.PARAMETER GroupNamePattern
    Wildcard ile grup adi deseni. Varsayilan: '*HERKES*'

.PARAMETER ListOnly
    Sadece disable uyeleri listeler, cikarmaz (kuru calisma).

.EXAMPLE
    # Once liste
    .\Remove-DisabledFromGroups.ps1 -ListOnly

    # Onayladiktan sonra gercek temizlik
    .\Remove-DisabledFromGroups.ps1 -GroupNamePattern '*HERKES*'
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string] $GroupNamePattern = '*HERKES*',
    [switch] $ListOnly
)

if (-not (Get-MgContext)) { throw "Once Connect-O365 calistirin." }
try { $null = Get-ConnectionInformation -ErrorAction Stop }
catch { throw "Exchange Online oturumu yok. Connect-ExchangeOnline calistirin." }

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd-HHmm')

Write-Host "[*] Grup arama: $GroupNamePattern" -ForegroundColor Cyan
$groups = Get-MgGroup -All | Where-Object { $_.DisplayName -like $GroupNamePattern }
Write-Host ("    Eslesen grup: {0}" -f $groups.Count) -ForegroundColor DarkGray

if (-not $groups) { return }

# Kullanici cache'i — birkac kez ayni kisiyi sorgulamamak icin
$userCache = @{}
function Get-UserStatus {
    param([string] $UserId)
    if ($userCache.ContainsKey($UserId)) { return $userCache[$UserId] }
    try {
        $u = Get-MgUser -UserId $UserId -Property Id,UserPrincipalName,DisplayName,AccountEnabled,AssignedLicenses -ErrorAction Stop
        $status = [pscustomobject]@{
            Id            = $u.Id
            UPN           = $u.UserPrincipalName
            DisplayName   = $u.DisplayName
            Enabled       = [bool]$u.AccountEnabled
            HasLicense    = ($u.AssignedLicenses.Count -gt 0)
            Exists        = $true
        }
    } catch {
        $status = [pscustomobject]@{
            Id            = $UserId
            UPN           = $null
            DisplayName   = $null
            Enabled       = $false
            HasLicense    = $false
            Exists        = $false
        }
    }
    $userCache[$UserId] = $status
    return $status
}

$flagged = New-Object System.Collections.Generic.List[object]

foreach ($g in $groups) {
    $type = if ($g.GroupTypes -contains 'Unified') { 'Microsoft365' }
            elseif ($g.SecurityEnabled -and $g.MailEnabled) { 'MailEnabledSecurity' }
            elseif ($g.SecurityEnabled) { 'Security' }
            elseif ($g.MailEnabled) { 'Distribution' }
            else { 'Other' }

    Write-Host ("[>] {0,-25} ({1})" -f $g.DisplayName, $type) -ForegroundColor Cyan
    $members = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue
    foreach ($m in $members) {
        $s = Get-UserStatus -UserId $m.Id
        if (-not $s.Exists -or -not $s.Enabled) {
            $reason = if (-not $s.Exists) { 'Silinmis/erisilemiyor' }
                      elseif (-not $s.Enabled) { 'Disabled' }
                      else { '?' }
            $flagged.Add([pscustomobject]@{
                GroupName   = $g.DisplayName
                GroupId     = $g.Id
                GroupType   = $type
                GroupMail   = $g.Mail
                MemberId    = $m.Id
                UPN         = $s.UPN
                DisplayName = $s.DisplayName
                Reason      = $reason
            })
        }
    }
}

$reportPath = Join-Path $dataDir "disabled-in-groups-$stamp.csv"
$flagged | Export-Csv -Path $reportPath -NoTypeInformation -Encoding utf8BOM

Write-Host ""
Write-Host ("[OK] Tespit edilen disable/silinmis uyelik: {0}" -f $flagged.Count) -ForegroundColor Yellow
Write-Host "     Rapor: $reportPath"

if ($ListOnly -or $flagged.Count -eq 0) {
    Write-Host ""
    Write-Host "ListOnly aktif - hicbir cikarma yapilmadi."
    return
}

Write-Host ""
$confirm = Read-Host "Bu $($flagged.Count) uyeligi gruplardan cikarmak icin 'EVET' yazin"
if ($confirm -ne 'EVET') {
    Write-Host "Iptal edildi."
    return
}

$removed = 0; $failed = 0
foreach ($row in $flagged) {
    $target = "$($row.UPN ?? $row.MemberId) from $($row.GroupName)"
    if ($PSCmdlet.ShouldProcess($target, "Remove member")) {
        try {
            switch ($row.GroupType) {
                { $_ -in 'Distribution','MailEnabledSecurity' } {
                    if (-not $row.UPN) { throw "UPN yok, EXO ile cikarilamiyor" }
                    Remove-DistributionGroupMember -Identity $row.GroupMail -Member $row.UPN -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
                }
                'Microsoft365' {
                    if (-not $row.UPN) { throw "UPN yok" }
                    Remove-UnifiedGroupLinks -Identity $row.GroupId -LinkType Members -Links $row.UPN -Confirm:$false -ErrorAction Stop
                }
                default {
                    Remove-MgGroupMemberByRef -GroupId $row.GroupId -DirectoryObjectId $row.MemberId -ErrorAction Stop
                }
            }
            Write-Host "[-] $target" -ForegroundColor Yellow
            $removed++
        } catch {
            Write-Host "[X] $target : $($_.Exception.Message)" -ForegroundColor Red
            $failed++
        }
    }
}

Write-Host ""
Write-Host ("Bitti. Cikarilan: {0}, Hata: {1}" -f $removed, $failed) -ForegroundColor Green
