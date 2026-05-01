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
    return $g
}

function Resolve-User {
    param([Parameter(Mandatory)][string] $Upn)
    $u = Get-MgUser -Filter "userPrincipalName eq '$($Upn -replace "'", "''")'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $u) { throw "Kullanici bulunamadi: $Upn" }
    return $u
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
            New-MgGroupMember -GroupId $g.Id -DirectoryObjectId $u.Id
            Write-Host "[OK] Eklendi: $UserUpn -> $GroupName" -ForegroundColor Green
        }
    }

    'RemoveMember' {
        $g = Resolve-Group -Name $GroupName
        $u = Resolve-User  -Upn  $UserUpn
        if ($PSCmdlet.ShouldProcess("$UserUpn from $GroupName", "Remove member")) {
            Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $u.Id
            Write-Host "[OK] Cikarildi: $UserUpn from $GroupName" -ForegroundColor Yellow
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
                    'Add'    { New-MgGroupMember         -GroupId $g.Id -DirectoryObjectId $u.Id; Write-Host "[+] $($r.UserPrincipalName) -> $($r.GroupName)" -ForegroundColor Green }
                    'Remove' { Remove-MgGroupMemberByRef -GroupId $g.Id -DirectoryObjectId $u.Id; Write-Host "[-] $($r.UserPrincipalName) from $($r.GroupName)" -ForegroundColor Yellow }
                    default  { Write-Host "[!] Bilinmeyen Action: $($r.Action)" -ForegroundColor Red }
                }
            } catch {
                Write-Host "[X] $($r.UserPrincipalName) / $($r.GroupName): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }
}
