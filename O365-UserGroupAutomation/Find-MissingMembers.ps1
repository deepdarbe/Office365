<#
.SYNOPSIS
    Bir gruba uye olmasi beklenen ama olmayan aktif kullanicilari bulur ve
    Action='Add' onceden doldurulmus bir CSV uretir.

.DESCRIPTION
    Beklenen uyelik kriterleri:
        -ExpectedDomain  : Bu domain'den tum aktif lisansli kullanicilar
                            (orn. 'nbr.com.tr')
        -IncludeShared   : Shared mailbox'lari da beklenenler listesine ekler
                            (varsayilan: dahil edilmez)
        -ExcludeUpns     : Tek tek haric tutulacak UPN'ler (servis hesaplari)

    Cikti CSV'si BulkFromCsv ile uyumlu:
        GroupName, UserPrincipalName, Action='Add', Reason

.EXAMPLE
    .\Find-MissingMembers.ps1 -GroupName 'NBR HERKES' -ExpectedDomain 'nbr.com.tr' `
        -ExcludeUpns 'logords@nbr.com.tr','partneradmin@nbr.com.tr'

    .\Find-MissingMembers.ps1 -GroupName 'NETA HERKES'     -ExpectedDomain 'netaekipman.com.tr'
    .\Find-MissingMembers.ps1 -GroupName 'TURBOFIN HERKES' -ExpectedDomain 'turbofin.com.tr'
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]   $GroupName,
    [Parameter(Mandatory)] [string]   $ExpectedDomain,
    [string[]] $ExcludeUpns = @(),
    [switch]   $IncludeShared
)

if (-not (Get-MgContext)) { throw "Once Connect-O365 calistirin." }

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd-HHmm')

$g = Get-MgGroup -Filter "displayName eq '$($GroupName -replace "'", "''")'" -ConsistencyLevel eventual -CountVariable c | Select-Object -First 1
if (-not $g) { throw "Grup bulunamadi: $GroupName" }

Write-Host "[*] '$GroupName' grubu uyeleri ve $ExpectedDomain aktif kullanicilari karsilastiriliyor..." -ForegroundColor Cyan

$members = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue
$memberUpns = $members | ForEach-Object { $_.AdditionalProperties['userPrincipalName'] } | Where-Object { $_ }
$memberSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$memberUpns, [System.StringComparer]::OrdinalIgnoreCase)

$excludeSet = [System.Collections.Generic.HashSet[string]]::new([string[]]$ExcludeUpns, [System.StringComparer]::OrdinalIgnoreCase)

$users = Get-MgUser -All -ConsistencyLevel eventual `
    -Property Id,UserPrincipalName,DisplayName,AccountEnabled,AssignedLicenses,UserType |
    Where-Object {
        $_.AccountEnabled -and
        $_.AssignedLicenses.Count -gt 0 -and
        $_.UserType -ne 'Guest' -and
        ($_.UserPrincipalName -split '@')[1] -eq $ExpectedDomain
    }

# Shared mailbox dahil edilecekse Exchange'den de cek
$sharedUpns = @()
if ($IncludeShared) {
    try {
        $null = Get-ConnectionInformation -ErrorAction Stop
        $sharedUpns = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited |
            Where-Object { $_.PrimarySmtpAddress -like "*@$ExpectedDomain" } |
            ForEach-Object { $_.PrimarySmtpAddress }
        Write-Host ("[*] {0} shared mailbox da listeye eklendi" -f $sharedUpns.Count) -ForegroundColor DarkGray
    } catch {
        Write-Warning "Exchange Online bagli degil, shared mailbox kontrolu atlandi"
    }
}

$expected = @($users.UserPrincipalName) + @($sharedUpns) | Sort-Object -Unique

$missing = foreach ($upn in $expected) {
    if ($excludeSet.Contains($upn)) { continue }
    if ($memberSet.Contains($upn))  { continue }

    $user = $users | Where-Object UserPrincipalName -eq $upn | Select-Object -First 1
    [pscustomobject]@{
        GroupName         = $GroupName
        UserPrincipalName = $upn
        Action            = 'Add'
        DisplayName       = if ($user) { $user.DisplayName } else { '(shared)' }
        Reason            = if ($user) { 'Active+licensed, missing from group' }
                            else       { 'Shared mailbox, missing from group' }
    }
}

$missing = @($missing)

$outputFile = Join-Path $dataDir ("missing-{0}-{1}.csv" -f ($GroupName -replace '[^\w-]+','_'), $stamp)
$missing | Export-Csv -Path $outputFile -NoTypeInformation -Encoding utf8BOM

Write-Host ""
Write-Host ("=== '$GroupName' icin eksik uye sayisi: {0} ===" -f $missing.Count) -ForegroundColor Yellow
Write-Host "Cikti: $outputFile" -ForegroundColor Green
if ($missing.Count -gt 0) {
    Write-Host ""
    Write-Host "Sonraki adim:" -ForegroundColor Yellow
    Write-Host "  1) Gozden gecirin: ii '$outputFile'"
    Write-Host "  2) Eklenmemesi gerekenleri silin"
    Write-Host "  3) Toplu ekle:    .\Manage-Groups.ps1 -Action BulkFromCsv -CsvPath '$outputFile'"
}
