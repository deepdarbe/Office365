<#
.SYNOPSIS
    O365 tenant uzerindeki aktif (enabled, lisansli) kullanicilari listeler.

.DESCRIPTION
    Microsoft Graph uzerinden tum kullanicilari ceker; varsayilan olarak
    yalnizca AccountEnabled=true ve en az bir lisansa sahip kullanicilari
    dondurur. Sonucu pipeline'a dokebilir veya CSV'ye kaydedebilir.

    Cikti GitHub'a yuklenmemelidir; varsayilan CSV yolu .gitignore'lu
    "data/" klasorudur.

.PARAMETER IncludeUnlicensed
    Lisansiz fakat aktif hesaplari da dahil eder.

.PARAMETER ExportCsv
    Sonucu CSV olarak diske yazar.

.PARAMETER OutputPath
    CSV cikti yolu (varsayilan: .\data\active-users-YYYYMMDD.csv).

.EXAMPLE
    . .\Connect-O365Services.ps1; Connect-O365
    .\Get-ActiveUsers.ps1 -ExportCsv
#>

[CmdletBinding()]
param(
    [switch] $IncludeUnlicensed,
    [switch] $ExportCsv,
    [string] $OutputPath
)

if (-not (Get-MgContext)) {
    throw "Microsoft Graph oturumu yok. Once Connect-O365 calistirin."
}

Write-Host "[*] Kullanicilar Graph uzerinden cekiliyor..." -ForegroundColor Cyan

$properties = @(
    'Id','UserPrincipalName','DisplayName','GivenName','Surname',
    'Mail','JobTitle','Department','OfficeLocation','City','Country',
    'AccountEnabled','AssignedLicenses','CreatedDateTime'
)

$allUsers = Get-MgUser -All -Property $properties -ConsistencyLevel eventual `
    | Select-Object $properties

$active = $allUsers | Where-Object { $_.AccountEnabled -eq $true }

if (-not $IncludeUnlicensed) {
    $active = $active | Where-Object { $_.AssignedLicenses.Count -gt 0 }
}

$result = $active | ForEach-Object {
    [pscustomobject]@{
        UserPrincipalName = $_.UserPrincipalName
        DisplayName       = $_.DisplayName
        GivenName         = $_.GivenName
        Surname           = $_.Surname
        Mail              = $_.Mail
        JobTitle          = $_.JobTitle
        Department        = $_.Department
        OfficeLocation    = $_.OfficeLocation
        City              = $_.City
        Country           = $_.Country
        LicenseCount      = $_.AssignedLicenses.Count
        CreatedDateTime   = $_.CreatedDateTime
        Id                = $_.Id
    }
}

Write-Host ("    Bulundu: {0} aktif kullanici" -f $result.Count) -ForegroundColor Green

if ($ExportCsv) {
    if (-not $OutputPath) {
        $dataDir = Join-Path $PSScriptRoot 'data'
        if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
        $OutputPath = Join-Path $dataDir ("active-users-{0:yyyyMMdd-HHmm}.csv" -f (Get-Date))
    }
    $result | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "    CSV: $OutputPath" -ForegroundColor DarkGray
}

return $result
