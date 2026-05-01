<#
.SYNOPSIS
    Tenant'tan tek seferde kullanici, grup ve uyelik snapshot'i alir.

.DESCRIPTION
    Bir komutla tum aktif kullanicilari, tum gruplari ve her grubun uye
    UPN listesini cekip "data/tenant-snapshot-<tarih>.json" olarak yazar.
    New-MappingFromExistingGroups.ps1 bu dosyayi okuyarak otomatik mapping
    onerir.

    Hicbir veri ekrana basilmaz — sadece dosyaya yazar (copy-paste sorunu yok).

.EXAMPLE
    . .\Connect-O365Services.ps1; Connect-O365
    .\Get-TenantSnapshot.ps1
#>

[CmdletBinding()]
param(
    [string] $OutputPath
)

if (-not (Get-MgContext)) { throw "Once Connect-O365 calistirin." }

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
if (-not $OutputPath) {
    $OutputPath = Join-Path $dataDir ("tenant-snapshot-{0:yyyyMMdd-HHmm}.json" -f (Get-Date))
}

Write-Host "[1/3] Aktif kullanicilar cekiliyor..." -ForegroundColor Cyan
$users = Get-MgUser -All -ConsistencyLevel eventual `
    -Property Id,UserPrincipalName,DisplayName,Mail,AccountEnabled,AssignedLicenses,UsageLocation |
    Where-Object { $_.AccountEnabled -and $_.AssignedLicenses.Count -gt 0 } |
    ForEach-Object {
        [pscustomobject]@{
            Id    = $_.Id
            UPN   = $_.UserPrincipalName
            Name  = $_.DisplayName
            Mail  = $_.Mail
            Loc   = $_.UsageLocation
        }
    }
Write-Host ("       {0} aktif kullanici" -f $users.Count) -ForegroundColor DarkGray

Write-Host "[2/3] Gruplar cekiliyor..." -ForegroundColor Cyan
$groups = Get-MgGroup -All
Write-Host ("       {0} grup" -f $groups.Count) -ForegroundColor DarkGray

Write-Host "[3/3] Uyelikler cekiliyor..." -ForegroundColor Cyan
$userIdToUpn = @{}
foreach ($u in $users) { $userIdToUpn[$u.Id] = $u.UPN }

$groupData = foreach ($g in $groups) {
    $members = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue
    $upns = foreach ($m in $members) {
        if ($userIdToUpn.ContainsKey($m.Id)) { $userIdToUpn[$m.Id] }
        else { $m.AdditionalProperties.userPrincipalName }
    }
    $type = if ($g.GroupTypes -contains 'Unified') { 'Microsoft365' }
            elseif ($g.SecurityEnabled -and $g.MailEnabled) { 'MailEnabledSecurity' }
            elseif ($g.SecurityEnabled) { 'Security' }
            elseif ($g.MailEnabled) { 'Distribution' }
            else { 'Other' }
    [pscustomobject]@{
        Id           = $g.Id
        DisplayName  = $g.DisplayName
        Mail         = $g.Mail
        Type         = $type
        MemberCount  = ($upns | Where-Object { $_ }).Count
        MemberUpns   = @($upns | Where-Object { $_ })
    }
}

$snapshot = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString('o')
    TenantId    = (Get-MgContext).TenantId
    Users       = $users
    Groups      = $groupData
}

$snapshot | ConvertTo-Json -Depth 6 | Set-Content -Path $OutputPath -Encoding UTF8
Write-Host ""
Write-Host "[OK] Snapshot yazildi: $OutputPath" -ForegroundColor Green
Write-Host "     Sonraki adim: .\New-MappingFromExistingGroups.ps1 -SnapshotPath '$OutputPath'" -ForegroundColor Yellow
