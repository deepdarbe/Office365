<#
.SYNOPSIS
    Hicbir gruba (veya sadece tanimladiginiz default gruplara) uye olmayan
    aktif kullanicilari listeler.

.DESCRIPTION
    Get-TenantSnapshot.ps1 ciktisini okur. Her aktif kullaniciyi tarar,
    uye oldugu gruplari bulur ve "orphan" kullanicilari iki dosyaya yazar:

      data/orphan-users.csv         : Hic gruba uye olmayanlar
      data/user-group-matrix.csv    : Her kullanici + uye oldugu grup sayisi
                                       + grup adlari (tum kullanicilar)
      data/orphan-bulk-template.csv : Toplu ekleme icin hazir CSV sablonu
                                       (GroupName,UserPrincipalName,Action)

    -IgnoreGroups parametresi ile "All Company" gibi default gruplari
    saymadan orphan kabul edilir.

    Hicbir veri ekrana basilmaz, sadece dosyaya yazar.

.EXAMPLE
    .\Get-OrphanUsers.ps1 -SnapshotPath .\data\tenant-snapshot-...json `
        -IgnoreGroups 'All Company','NBR HERKES','NETA HERKES','TURBOFIN HERKES'
#>

[CmdletBinding()]
param(
    [string]   $SnapshotPath,
    [string[]] $IgnoreGroups = @()
)

if (-not $SnapshotPath) {
    $dataDir = Join-Path $PSScriptRoot 'data'
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

if (-not (Test-Path $SnapshotPath)) { throw "Snapshot bulunamadi: $SnapshotPath" }

$dataDir = Split-Path $SnapshotPath -Parent
$snap = Get-Content $SnapshotPath -Raw | ConvertFrom-Json

# UPN -> grup adlari listesi
$userGroups = @{}
foreach ($u in $snap.Users) { $userGroups[$u.UPN] = New-Object System.Collections.Generic.List[string] }

foreach ($g in $snap.Groups) {
    foreach ($upn in $g.MemberUpns) {
        if ($upn -and $userGroups.ContainsKey($upn)) {
            $userGroups[$upn].Add($g.DisplayName)
        }
    }
}

$ignoreSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$IgnoreGroups, [System.StringComparer]::OrdinalIgnoreCase)

$matrix = foreach ($u in $snap.Users) {
    $groups = $userGroups[$u.UPN]
    $effective = @($groups | Where-Object { -not $ignoreSet.Contains($_) })
    [pscustomobject]@{
        UPN              = $u.UPN
        DisplayName      = $u.Name
        GroupCount       = $groups.Count
        EffectiveCount   = $effective.Count
        IsOrphan         = ($effective.Count -eq 0)
        Groups           = ($groups -join '; ')
    }
}

$matrixPath = Join-Path $dataDir 'user-group-matrix.csv'
$orphanPath = Join-Path $dataDir 'orphan-users.csv'
$bulkPath   = Join-Path $dataDir 'orphan-bulk-template.csv'

$matrix | Sort-Object EffectiveCount, UPN |
    Export-Csv -Path $matrixPath -NoTypeInformation -Encoding UTF8

$orphans = $matrix | Where-Object IsOrphan
$orphans | Select-Object UPN, DisplayName, GroupCount, Groups |
    Export-Csv -Path $orphanPath -NoTypeInformation -Encoding UTF8

# Toplu ekleme sablonu — Manage-Groups.ps1 -Action BulkFromCsv ile uyumlu
$orphans | ForEach-Object {
    [pscustomobject]@{
        GroupName         = ''
        UserPrincipalName = $_.UPN
        Action            = 'Add'
    }
} | Export-Csv -Path $bulkPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host ("[OK] Toplam aktif kullanici  : {0}" -f $matrix.Count) -ForegroundColor Cyan
Write-Host ("     Orphan (gruba uye degil): {0}" -f $orphans.Count) -ForegroundColor Yellow
if ($IgnoreGroups.Count -gt 0) {
    Write-Host ("     Goz ardi edilen gruplar : {0}" -f ($IgnoreGroups -join ', ')) -ForegroundColor DarkGray
}
Write-Host ""
Write-Host "Yazildi:" -ForegroundColor Green
Write-Host "  $matrixPath"
Write-Host "  $orphanPath"
Write-Host "  $bulkPath   <- Excel'de acin, GroupName sutununu doldurun"
Write-Host ""
Write-Host "Sonraki adim:" -ForegroundColor Yellow
Write-Host "  1) $bulkPath dosyasini Excel'de acin"
Write-Host "  2) Her satir icin GroupName sutununa hedef grup adini yazin"
Write-Host "     (Bos satirlari silebilirsiniz - sadece eklenecekleri birakin)"
Write-Host "  3) Toplu uygula:"
Write-Host "     .\Manage-Groups.ps1 -Action BulkFromCsv -CsvPath '$bulkPath'"
