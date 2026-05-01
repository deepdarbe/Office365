<#
.SYNOPSIS
    Snapshot'taki mevcut grup uyeliklerinden mapping kurallari onerir.

.DESCRIPTION
    Get-TenantSnapshot.ps1 ciktisini okur. Her mail-enabled grup icin uye
    UPN'lerine bakar, ortak desen tespit eder ve "group-mapping.suggested.json"
    dosyasini yazar. Tespit ettigi kural turleri:

      - Tum uyeler ayni mail domain'inden ise:  Mail endsWith "@domain"
      - Tum UPN'lerin local-part'i ayni on ek ile basliyorsa (>=3 char):
                                                UserPrincipalName startsWith "<onek>"
      - Aksi halde grup "manuel-inceleme" olarak isaretlenir.

    Yalnizca dosyaya yazar, hicbir degisiklik yapmaz. Cikti dosyasini gozden
    gecirip "group-mapping.json" olarak adlandirip Sync scripti calistirilir.

.EXAMPLE
    .\New-MappingFromExistingGroups.ps1 -SnapshotPath .\data\tenant-snapshot-...json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string] $SnapshotPath,
    [string] $OutputPath,
    [int]    $MinMembers = 2,
    [double] $PatternThreshold = 0.85
)

if (-not (Test-Path $SnapshotPath)) { throw "Snapshot bulunamadi: $SnapshotPath" }

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
if (-not $OutputPath) {
    $OutputPath = Join-Path $dataDir 'group-mapping.suggested.json'
}

$snap = Get-Content $SnapshotPath -Raw | ConvertFrom-Json

function Get-LongestCommonPrefix {
    param([string[]] $Strings)
    if (-not $Strings -or $Strings.Count -eq 0) { return '' }
    $sorted = $Strings | Sort-Object
    $first = $sorted[0]; $last = $sorted[-1]
    $i = 0
    while ($i -lt $first.Length -and $i -lt $last.Length -and $first[$i] -eq $last[$i]) { $i++ }
    return $first.Substring(0, $i)
}

$rules        = New-Object System.Collections.Generic.List[object]
$manualReview = New-Object System.Collections.Generic.List[object]

foreach ($g in $snap.Groups) {
    if (-not $g.MemberUpns -or $g.MemberUpns.Count -lt $MinMembers) { continue }
    if ($g.Type -notin @('Distribution','MailEnabledSecurity','Microsoft365')) { continue }

    $upns    = @($g.MemberUpns | Where-Object { $_ })
    $domains = $upns | ForEach-Object { ($_ -split '@')[1] } | Sort-Object -Unique
    $locals  = $upns | ForEach-Object { ($_ -split '@')[0].ToLower() }

    $rule = $null

    if ($domains.Count -eq 1) {
        $rule = [pscustomobject]@{
            name   = $g.DisplayName
            match  = [pscustomobject]@{ field = 'Mail'; operator = 'endsWith'; value = "@$($domains[0])" }
            groups = @($g.DisplayName)
            note   = "Tum {0} uye '@{1}' domain'inden" -f $upns.Count, $domains[0]
        }
    }
    else {
        $prefix = Get-LongestCommonPrefix -Strings $locals
        if ($prefix.Length -ge 3) {
            $matchCount = ($locals | Where-Object { $_.StartsWith($prefix) }).Count
            $ratio = $matchCount / $locals.Count
            if ($ratio -ge $PatternThreshold) {
                $rule = [pscustomobject]@{
                    name   = $g.DisplayName
                    match  = [pscustomobject]@{ field = 'UserPrincipalName'; operator = 'startsWith'; value = $prefix }
                    groups = @($g.DisplayName)
                    note   = "Uyelerin %{0:N0}'i '{1}' on eki ile basliyor ({2}/{3})" -f ($ratio*100), $prefix, $matchCount, $locals.Count
                }
            }
        }
    }

    if ($rule) {
        $rules.Add($rule)
    } else {
        $manualReview.Add([pscustomobject]@{
            group       = $g.DisplayName
            type        = $g.Type
            memberCount = $upns.Count
            domains     = $domains
            sampleUpns  = ($upns | Select-Object -First 5)
            reason      = "Ortak desen tespit edilemedi"
        })
    }
}

$output = [pscustomobject]@{
    _comment      = "Otomatik onerilen mapping. Inceleyin, gerekirse duzenleyin, sonra 'group-mapping.json' olarak kopyalayin."
    generatedAt   = (Get-Date).ToString('o')
    defaultGroups = @()
    rules         = $rules
    manualReview  = $manualReview
}

$output | ConvertTo-Json -Depth 6 | Set-Content -Path $OutputPath -Encoding UTF8

Write-Host ""
Write-Host "[OK] Onerilen mapping yazildi: $OutputPath" -ForegroundColor Green
Write-Host ("     Otomatik kural: {0}, Manuel inceleme gereken grup: {1}" -f $rules.Count, $manualReview.Count) -ForegroundColor DarkGray
Write-Host ""
Write-Host "Sonraki adimlar:" -ForegroundColor Yellow
Write-Host "  1) $OutputPath dosyasini gozden gecirin"
Write-Host "  2) Beğendiginizde: Copy-Item '$OutputPath' '.\group-mapping.json'"
Write-Host "  3) Test:           .\Sync-UserGroupMembership.ps1 -WhatIf"
