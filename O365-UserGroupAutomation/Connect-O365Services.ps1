<#
.SYNOPSIS
    Office 365 / Microsoft 365 yonetim modullerine baglanti yardimcisi.

.DESCRIPTION
    Microsoft Graph ve Exchange Online PowerShell modullerini yukler ve
    interaktif (modern auth / MFA destekli) oturum acar. Diger scriptler
    bu dosyayi dot-source ederek "Connect-O365" fonksiyonunu kullanir.

    Kullanim:
        . .\Connect-O365Services.ps1
        Connect-O365 -TenantId "contoso.onmicrosoft.com"

.NOTES
    Gerekli moduller:
        Install-Module Microsoft.Graph -Scope CurrentUser
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
#>

function Ensure-Module {
    param(
        [Parameter(Mandatory)] [string] $Name,
        [string] $MinimumVersion
    )

    $installed = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending | Select-Object -First 1

    if (-not $installed) {
        Write-Host "[*] $Name modulu bulunamadi, yukleniyor..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    elseif ($MinimumVersion -and $installed.Version -lt [version]$MinimumVersion) {
        Write-Host "[*] $Name guncelleniyor (>= $MinimumVersion)..." -ForegroundColor Yellow
        Update-Module -Name $Name -Force
    }

    Import-Module -Name $Name -ErrorAction Stop
}

function Connect-O365 {
    [CmdletBinding()]
    param(
        [string] $TenantId,
        [string] $UserPrincipalName,
        [switch] $SkipExchange,
        [switch] $SkipGraph
    )

    if (-not $SkipGraph) {
        Ensure-Module -Name Microsoft.Graph -MinimumVersion '2.0.0'

        $scopes = @(
            'User.Read.All',
            'Group.ReadWrite.All',
            'GroupMember.ReadWrite.All',
            'Directory.Read.All'
        )

        $params = @{ Scopes = $scopes; NoWelcome = $true }
        if ($TenantId) { $params['TenantId'] = $TenantId }

        Write-Host "[*] Microsoft Graph'e baglaniliyor..." -ForegroundColor Cyan
        Connect-MgGraph @params | Out-Null

        $ctx = Get-MgContext
        Write-Host ("    OK -> Tenant: {0} | Hesap: {1}" -f $ctx.TenantId, $ctx.Account) -ForegroundColor Green
    }

    if (-not $SkipExchange) {
        Ensure-Module -Name ExchangeOnlineManagement

        Write-Host "[*] Exchange Online'a baglaniliyor..." -ForegroundColor Cyan
        $exoParams = @{ ShowBanner = $false }
        if ($UserPrincipalName) { $exoParams['UserPrincipalName'] = $UserPrincipalName }
        Connect-ExchangeOnline @exoParams
        Write-Host "    OK -> Exchange Online oturumu acildi" -ForegroundColor Green
    }
}

function Disconnect-O365 {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    Write-Host "[*] O365 oturumlari kapatildi." -ForegroundColor DarkGray
}
