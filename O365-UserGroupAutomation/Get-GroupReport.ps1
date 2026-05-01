<#
.SYNOPSIS
    O365 gruplarini analiz eder, listeler ve rapor (CSV / HTML) uretir.

.DESCRIPTION
    - Tum gruplari (Microsoft 365, Security, Distribution, Mail-enabled Security)
      ozet halinde dokumler: tip, uye sayisi, sahip sayisi, mail, olusturulma.
    - -Detailed verildiginde her grubun uye listesini ek CSV olarak yazar.
    - -GroupName ile tek bir grubu hedefleyebilirsiniz.

    Cikti varsayilan olarak data/ altina yazilir; bu klasor .gitignore'ludur.

.PARAMETER Detailed
    Her grup icin uye listesini ayri CSV olarak yazar.

.PARAMETER GroupName
    Yalnizca belirtilen grubu raporlar (wildcard destekler).

.PARAMETER Html
    Ozet raporu HTML olarak da uretir.

.EXAMPLE
    .\Get-GroupReport.ps1 -Detailed -Html
    .\Get-GroupReport.ps1 -GroupName 'NBR *'
#>

[CmdletBinding()]
param(
    [switch] $Detailed,
    [string] $GroupName,
    [switch] $Html
)

if (-not (Get-MgContext)) {
    throw "Microsoft Graph oturumu yok. Once Connect-O365 calistirin."
}

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd-HHmm')

Write-Host "[*] Gruplar cekiliyor..." -ForegroundColor Cyan

$groups = if ($GroupName) {
    Get-MgGroup -All -Filter "startswith(displayName,'$($GroupName.TrimEnd('*'))')" -ConsistencyLevel eventual -CountVariable c
} else {
    Get-MgGroup -All
}

Write-Host ("    Bulundu: {0} grup" -f $groups.Count) -ForegroundColor Green

$summary = foreach ($g in $groups) {
    $members = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue
    $owners  = Get-MgGroupOwner  -GroupId $g.Id -All -ErrorAction SilentlyContinue

    $type = if ($g.GroupTypes -contains 'Unified') { 'Microsoft365' }
            elseif ($g.SecurityEnabled -and $g.MailEnabled) { 'MailEnabledSecurity' }
            elseif ($g.SecurityEnabled) { 'Security' }
            elseif ($g.MailEnabled) { 'Distribution' }
            else { 'Other' }

    [pscustomobject]@{
        DisplayName     = $g.DisplayName
        Type            = $type
        Mail            = $g.Mail
        MemberCount     = ($members | Measure-Object).Count
        OwnerCount      = ($owners | Measure-Object).Count
        Visibility      = $g.Visibility
        CreatedDateTime = $g.CreatedDateTime
        Id              = $g.Id
    }

    if ($Detailed) {
        $detailPath = Join-Path $dataDir ("group-members-{0}-{1}.csv" -f ($g.DisplayName -replace '[^\w\-]+','_'), $stamp)
        $members | Select-Object @{N='GroupName';E={$g.DisplayName}}, Id, @{N='UPN';E={$_.AdditionalProperties.userPrincipalName}}, @{N='DisplayName';E={$_.AdditionalProperties.displayName}} |
            Export-Csv -Path $detailPath -NoTypeInformation -Encoding UTF8
    }
}

$summaryPath = Join-Path $dataDir ("group-summary-{0}.csv" -f $stamp)
$summary | Sort-Object DisplayName | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
Write-Host "    Ozet CSV: $summaryPath" -ForegroundColor DarkGray

if ($Html) {
    $htmlPath = Join-Path $dataDir ("group-summary-{0}.html" -f $stamp)
    $style = "<style>body{font-family:Segoe UI,Arial;font-size:13px} table{border-collapse:collapse} th,td{border:1px solid #ccc;padding:4px 8px} th{background:#f2f2f2}</style>"
    $summary | Sort-Object DisplayName |
        ConvertTo-Html -Title "O365 Group Report" -PreContent "<h2>O365 Group Report - $stamp</h2>" -Head $style |
        Set-Content -Path $htmlPath -Encoding UTF8
    Write-Host "    HTML: $htmlPath" -ForegroundColor DarkGray
}

return $summary
