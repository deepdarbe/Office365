<#
.SYNOPSIS
    Tum aktif kullanicilar ve tum gruplari tek bir matris CSV'sinde verir
    (rows = kullanici, columns = grup, X = uye).

.DESCRIPTION
    Get-TenantSnapshot.ps1 ciktisini okur (yoksa otomatik olarak yeni
    snapshot olusturur). Iki cikti dosyasi yazar:

      data/membership-matrix-<stamp>.csv
        Excel'de pivot/filter icin ideal — kullanici basi bir satir,
        her grup icin ayri kolon (1 / bos), ek olarak GroupCount sutunu.

      data/membership-flat-<stamp>.csv
        Her (kullanici, grup) ikilisi ayri satir; basit raporlama icin.

    Hicbir veri ekrana basilmaz; tum cikti data/ altina yazilir
    (gitignore'lu).

.PARAMETER SnapshotPath
    Kullanilacak snapshot. Verilmezse en son snapshot kullanilir, o da
    yoksa yenisi alinir.

.PARAMETER GroupNamePattern
    Sadece eslesen gruplari matrise dahil et (orn. '*HERKES*').
    Varsayilan: tum mail-enabled gruplar.

.PARAMETER IncludeAllGroups
    Security-only gruplarini da dahil eder (varsayilan: yalniz mail-enabled).

.EXAMPLE
    .\Get-MembershipMatrix.ps1
    .\Get-MembershipMatrix.ps1 -GroupNamePattern '*HERKES*'
#>

[CmdletBinding()]
param(
    [string] $SnapshotPath,
    [string] $GroupNamePattern,
    [switch] $IncludeAllGroups
)

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd-HHmm')

if (-not $SnapshotPath) {
    $latest = Get-ChildItem (Join-Path $dataDir 'tenant-snapshot-*.json') -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($latest) {
        $SnapshotPath = $latest.FullName
        Write-Host "[*] Mevcut snapshot kullanilacak: $($latest.Name)" -ForegroundColor DarkGray
    } else {
        Write-Host "[*] Snapshot yok, yenisi aliniyor..." -ForegroundColor Yellow
        & (Join-Path $PSScriptRoot 'Get-TenantSnapshot.ps1')
        $latest = Get-ChildItem (Join-Path $dataDir 'tenant-snapshot-*.json') |
                  Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $SnapshotPath = $latest.FullName
    }
}

$snap = Get-Content $SnapshotPath -Raw | ConvertFrom-Json

# Grup filtresi
$groups = $snap.Groups
if (-not $IncludeAllGroups) {
    $groups = $groups | Where-Object { $_.Type -in 'Distribution','MailEnabledSecurity','Microsoft365' }
}
if ($GroupNamePattern) {
    $groups = $groups | Where-Object { $_.DisplayName -like $GroupNamePattern }
}
$groups = @($groups | Sort-Object DisplayName)

Write-Host ("[*] Matris hazirlaniyor: {0} kullanici x {1} grup" -f $snap.Users.Count, $groups.Count) -ForegroundColor Cyan

# UPN -> grup adlari
$userToGroups = @{}
foreach ($u in $snap.Users) { $userToGroups[$u.UPN] = New-Object System.Collections.Generic.HashSet[string] ([System.StringComparer]::OrdinalIgnoreCase) }
foreach ($g in $groups) {
    foreach ($upn in $g.MemberUpns) {
        if ($upn -and $userToGroups.ContainsKey($upn)) {
            [void]$userToGroups[$upn].Add($g.DisplayName)
        }
    }
}

# Wide matrix
$matrix = foreach ($u in $snap.Users | Sort-Object UPN) {
    $row = [ordered]@{
        UPN         = $u.UPN
        DisplayName = $u.Name
        Domain      = ($u.UPN -split '@')[1]
        GroupCount  = $userToGroups[$u.UPN].Count
    }
    foreach ($g in $groups) {
        $row[$g.DisplayName] = if ($userToGroups[$u.UPN].Contains($g.DisplayName)) { 'X' } else { '' }
    }
    [pscustomobject]$row
}

$matrixFile = Join-Path $dataDir "membership-matrix-$stamp.csv"
$matrix | Export-Csv -Path $matrixFile -NoTypeInformation -Encoding utf8BOM

# Long/flat
$flat = foreach ($u in $snap.Users | Sort-Object UPN) {
    $list = @($userToGroups[$u.UPN])
    if ($list.Count -eq 0) {
        [pscustomobject]@{ UPN = $u.UPN; DisplayName = $u.Name; Domain = ($u.UPN -split '@')[1]; GroupName = '(none)' }
    } else {
        foreach ($gname in $list | Sort-Object) {
            [pscustomobject]@{ UPN = $u.UPN; DisplayName = $u.Name; Domain = ($u.UPN -split '@')[1]; GroupName = $gname }
        }
    }
}
$flatFile = Join-Path $dataDir "membership-flat-$stamp.csv"
$flat | Export-Csv -Path $flatFile -NoTypeInformation -Encoding utf8BOM

# Ozet
$noGroup       = ($matrix | Where-Object GroupCount -eq 0).Count
$avgGroupCount = if ($matrix.Count -gt 0) { [math]::Round((($matrix | Measure-Object GroupCount -Average).Average),1) } else { 0 }

Write-Host ""
Write-Host "=== Ozet ===" -ForegroundColor Yellow
Write-Host ("  Toplam aktif kullanici  : {0}" -f $matrix.Count)
Write-Host ("  Hicbir gruba uye olmayan: {0}" -f $noGroup)
Write-Host ("  Ortalama uyelik         : {0}" -f $avgGroupCount)
Write-Host ("  Matrise dahil edilen grup: {0}" -f $groups.Count)
Write-Host ""
Write-Host "Yazildi:" -ForegroundColor Green
Write-Host "  Matris (wide) : $matrixFile"
Write-Host "  Flat (long)   : $flatFile"
Write-Host ""
Write-Host "Acmak icin:" -ForegroundColor DarkGray
Write-Host "  ii '$matrixFile'"
