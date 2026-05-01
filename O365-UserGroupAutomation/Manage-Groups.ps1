<#
.SYNOPSIS
    O365 gruplarini olusturma / duzenleme / silme yardimcisi.

.DESCRIPTION
    Microsoft Graph uzerinden grup olusturur, ad / aciklama / mail nick
    ozelliklerini gunceller, uye ekler/cikarir veya grubu siler.

    -Action degerleri:
        Create        : Yeni grup olusturur (Microsoft365, Security, MailEnabledSecurity, Distribution)
        Update        : Mevcut grubu gunceller (DisplayName / Description / MailNickname)
        AddMember     : -UserUpn ile uye ekler
        RemoveMember  : -UserUpn ile uye cikarir
        Delete        : Grubu siler (onay ister)
        BulkFromCsv   : CSV'den toplu uye ekler/cikarir (CSV: GroupName,UserPrincipalName,Action)

.EXAMPLE
    .\Manage-Groups.ps1 -Action Create -GroupName 'NBR BURSA' -MailNickname 'nbr-bursa' -Description 'NBR Bursa toplu mail grubu' -GroupType MailEnabledSecurity
    .\Manage-Groups.ps1 -Action AddMember -GroupName 'NBR BURSA' -UserUpn 'ali@contoso.com'
    .\Manage-Groups.ps1 -Action BulkFromCsv -CsvPath .\data\bulk.csv
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Create','Update','AddMember','RemoveMember','Delete','BulkFromCsv')]
    [string] $Action,

    [string] $GroupName,
    [string] $NewDisplayName,
    [string] $Description,
    [string] $MailNickname,

    [ValidateSet('Microsoft365','Security','MailEnabledSecurity','Distribution')]
    [string] $GroupType = 'Microsoft365',

    [string] $UserUpn,
    [string] $CsvPath
)

if (-not (Get-MgContext)) {
    throw "Microsoft Graph oturumu yok. Once Connect-O365 calistirin."
}

function Resolve-Group {
    param([Parameter(Mandatory)][string] $Name)
    $g = Get-MgGroup -Filter "displayName eq '$($Name -replace "'", "''")'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $g) { throw "Grup bulunamadi: $Name" }
    $type = if ($g.GroupTypes -contains 'Unified') { 'Microsoft365' }
            elseif ($g.SecurityEnabled -and $g.MailEnabled) { 'MailEnabledSecurity' }
            elseif ($g.SecurityEnabled) { 'Security' }
            elseif ($g.MailEnabled) { 'Distribution' }
            else { 'Other' }
    Add-Member -InputObject $g -NotePropertyName ResolvedType -NotePropertyValue $type -Force
    return $g
}

function Resolve-User {
    param([Parameter(Mandatory)][string] $Upn)
    $u = Get-MgUser -Filter "userPrincipalName eq '$($Upn -replace "'", "''")'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $u) { throw "Kullanici bulunamadi: $Upn" }
    return $u
}

function Test-ExoConnected {
    try { $null = Get-ConnectionInformation -ErrorAction Stop; return $true }
    catch { return $false }
}

function Add-MemberSmart {
    param($Group, $User)
    switch ($Group.ResolvedType) {
        { $_ -in 'Distribution','MailEnabledSecurity' } {
            if (-not (Test-ExoConnected)) { throw "Exchange Online oturumu yok. Connect-ExchangeOnline calistirin." }
            Add-DistributionGroupMember -Identity $Group.Mail -Member $User.UserPrincipalName -BypassSecurityGroupManagerCheck -ErrorAction Stop
        }
        'Microsoft365' {
            if (-not (Test-ExoConnected)) { throw "Exchange Online oturumu yok. Connect-ExchangeOnline calistirin." }
            Add-UnifiedGroupLinks -Identity $Group.Id -LinkType Members -Links $User.UserPrincipalName -ErrorAction Stop
        }
        default {
            New-MgGroupMember -GroupId $Group.Id -DirectoryObjectId $User.Id -ErrorAction Stop
        }
    }
}

function Remove-MemberSmart {
    param($Group, $User)
    switch ($Group.ResolvedType) {
        { $_ -in 'Distribution','MailEnabledSecurity' } {
            if (-not (Test-ExoConnected)) { throw "Exchange Online oturumu yok. Connect-ExchangeOnline calistirin." }
            Remove-DistributionGroupMember -Identity $Group.Mail -Member $User.UserPrincipalName -BypassSecurityGroupManagerCheck -Confirm:$false -ErrorAction Stop
        }
        'Microsoft365' {
            if (-not (Test-ExoConnected)) { throw "Exchange Online oturumu yok. Connect-ExchangeOnline calistirin." }
            Remove-UnifiedGroupLinks -Identity $Group.Id -LinkType Members -Links $User.UserPrincipalName -Confirm:$false -ErrorAction Stop
        }
        default {
            Remove-MgGroupMemberByRef -GroupId $Group.Id -DirectoryObjectId $User.Id -ErrorAction Stop
        }
    }
}

switch ($Action) {

    'Create' {
        if (-not $GroupName)    { throw "-GroupName zorunlu" }
        if (-not $MailNickname) { $MailNickname = ($GroupName -replace '[^a-zA-Z0-9]','').ToLower() }

        $body = @{
            DisplayName  = $GroupName
            Description  = $Description
            MailNickname = $MailNickname
        }

        switch ($GroupType) {
            'Microsoft365'        { $body.GroupTypes = @('Unified'); $body.MailEnabled = $true;  $body.SecurityEnabled = $false }
            'Security'            { $body.GroupTypes = @();          $body.MailEnabled = $false; $body.SecurityEnabled = $true  }
            'MailEnabledSecurity' { $body.GroupTypes = @();          $body.MailEnabled = $true;  $body.SecurityEnabled = $true  }
            'Distribution'        { $body.GroupTypes = @();          $body.MailEnabled = $true;  $body.SecurityEnabled = $false }
        }

        if ($PSCmdlet.ShouldProcess($GroupName, "Create $GroupType group")) {
            $g = New-MgGroup -BodyParameter $body
            Write-Host "[OK] Grup olusturuldu: $($g.DisplayName) ($($g.Id))" -ForegroundColor Green
        }
    }

    'Update' {
        $g = Resolve-Group -Name $GroupName
        $patch = @{}
        if ($NewDisplayName) { $patch.DisplayName  = $NewDisplayName }
        if ($Description)    { $patch.Description  = $Description }
        if ($MailNickname)   { $patch.MailNickname = $MailNickname }
        if ($patch.Count -eq 0) { throw "Guncellenecek alan belirtilmedi" }

        if ($PSCmdlet.ShouldProcess($g.DisplayName, "Update group")) {
            Update-MgGroup -GroupId $g.Id -BodyParameter $patch
            Write-Host "[OK] Grup guncellendi: $($g.DisplayName)" -ForegroundColor Green
        }
    }

    'AddMember' {
        $g = Resolve-Group -Name $GroupName
        $u = Resolve-User  -Upn  $UserUpn
        if ($PSCmdlet.ShouldProcess("$UserUpn -> $GroupName", "Add member")) {
            try {
                Add-MemberSmart -Group $g -User $u
                Write-Host "[OK] Eklendi ($($g.ResolvedType)): $UserUpn -> $GroupName" -ForegroundColor Green
            } catch {
                Write-Host "[X] EKLENMEDI: $UserUpn -> $GroupName : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    'RemoveMember' {
        $g = Resolve-Group -Name $GroupName
        $u = Resolve-User  -Upn  $UserUpn
        if ($PSCmdlet.ShouldProcess("$UserUpn from $GroupName", "Remove member")) {
            try {
                Remove-MemberSmart -Group $g -User $u
                Write-Host "[OK] Cikarildi ($($g.ResolvedType)): $UserUpn from $GroupName" -ForegroundColor Yellow
            } catch {
                Write-Host "[X] CIKARILMADI: $UserUpn from $GroupName : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    'Delete' {
        $g = Resolve-Group -Name $GroupName
        if ($PSCmdlet.ShouldProcess($g.DisplayName, "DELETE group")) {
            Remove-MgGroup -GroupId $g.Id -Confirm:$false
            Write-Host "[OK] Silindi: $($g.DisplayName)" -ForegroundColor Yellow
        }
    }

    'BulkFromCsv' {
        if (-not $CsvPath -or -not (Test-Path $CsvPath)) { throw "CSV bulunamadi: $CsvPath" }
        $rows = Import-Csv -Path $CsvPath
        foreach ($r in $rows) {
            try {
                $g = Resolve-Group -Name $r.GroupName
                $u = Resolve-User  -Upn  $r.UserPrincipalName
                switch ($r.Action) {
                    'Add'    { Add-MemberSmart    -Group $g -User $u; Write-Host "[+] ($($g.ResolvedType)) $($r.UserPrincipalName) -> $($r.GroupName)" -ForegroundColor Green }
                    'Remove' { Remove-MemberSmart -Group $g -User $u; Write-Host "[-] ($($g.ResolvedType)) $($r.UserPrincipalName) from $($r.GroupName)" -ForegroundColor Yellow }
                    default  { Write-Host "[!] Bilinmeyen Action: $($r.Action)" -ForegroundColor Red }
                }
            } catch {
                Write-Host "[X] $($r.UserPrincipalName) / $($r.GroupName): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
