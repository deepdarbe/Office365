<#
.SYNOPSIS
    Bir grubun (ya da grup desenınin) tum uyelerini kategorize eder ve
    cikarilmasi onerilenler icin Action='Remove' onceden doldurulmus
    bir cleanup CSV'si uretir.

.DESCRIPTION
    Her uye icin Graph + Exchange Online sorgusu yapar:
        - AccountEnabled
        - AssignedLicenses
        - UserType (Member/Guest)
        - RecipientTypeDetails (UserMailbox/SharedMailbox/...)
        - DisplayName

    Onerilen aksiyon mantigi:
        Disabled                              -> Remove
        DisplayName "Ayr*Personel*"           -> Remove (ayrilan personel)
        SharedMailbox                         -> KEEP (toplu maile dahil olmali)
        Guest                                 -> KEEP (bilerek davet edilmis)
        Lisanssiz UserMailbox                 -> REVIEW (insan karari)
        ActiveLicensed                        -> KEEP
        Mail contact / nested group           -> KEEP

    Cikti:
        data/audit-<group>-<stamp>.csv     : tam detay rapor
        data/cleanup-<group>-<stamp>.csv   : BulkFromCsv ile uyumlu;
                                              Remove onerileri pre-filled,
                                              digerleri Action bos.

    Cleanup CSV'sini Excel'de gozden gecirin, kalmasini istediginizi
    silin/Action bosaltin, sonra:
        .\Manage-Groups.ps1 -Action BulkFromCsv -CsvPath <cleanup csv>

.PARAMETER GroupName
    Tek grup adi (orn. 'NBR HERKES').

.PARAMETER GroupNamePattern
    Wildcard desen (orn. '*HERKES*'). GroupName ile birlikte kullanilmaz.

.EXAMPLE
    .\Audit-GroupMembers.ps1 -GroupName 'NBR HERKES'
    .\Audit-GroupMembers.ps1 -GroupNamePattern '*HERKES*'
#>

[CmdletBinding(DefaultParameterSetName='Single')]
param(
    [Parameter(ParameterSetName='Single',  Mandatory)] [string] $GroupName,
    [Parameter(ParameterSetName='Pattern', Mandatory)] [string] $GroupNamePattern,
    [switch] $RemoveSharedMailboxes,
    [switch] $RemoveUnlicensed
)

if (-not (Get-MgContext)) { throw "Once Connect-O365 calistirin." }
try { $null = Get-ConnectionInformation -ErrorAction Stop }
catch { throw "Exchange Online oturumu yok. Connect-ExchangeOnline calistirin." }

$dataDir = Join-Path $PSScriptRoot 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir | Out-Null }
$stamp = (Get-Date).ToString('yyyyMMdd-HHmm')

$groups = if ($PSCmdlet.ParameterSetName -eq 'Single') {
    @(Get-MgGroup -Filter "displayName eq '$($GroupName -replace "'", "''")'" -ConsistencyLevel eventual -CountVariable c | Select-Object -First 1)
} else {
    Get-MgGroup -All | Where-Object { $_.DisplayName -like $GroupNamePattern }
}
$groups = @($groups | Where-Object { $_ })
if ($groups.Count -eq 0) { throw "Eslesen grup bulunamadi" }

Write-Host ("[*] {0} grup denetlenecek" -f $groups.Count) -ForegroundColor Cyan

$userCache    = @{}
$mailboxCache = @{}

function Get-UserDetail {
    param([string] $Id)
    if ($userCache.ContainsKey($Id)) { return $userCache[$Id] }
    try {
        $u = Get-MgUser -UserId $Id -Property Id,UserPrincipalName,DisplayName,AccountEnabled,AssignedLicenses,UserType -ErrorAction Stop
        $r = [pscustomobject]@{
            UPN         = $u.UserPrincipalName
            Name        = $u.DisplayName
            Enabled     = [bool]$u.AccountEnabled
            Licensed    = ($u.AssignedLicenses.Count -gt 0)
            UserType    = $u.UserType
            Exists      = $true
        }
    } catch {
        $r = [pscustomobject]@{ UPN=$null; Name=$null; Enabled=$false; Licensed=$false; UserType=$null; Exists=$false }
    }
    $userCache[$Id] = $r
    return $r
}

function Get-MailboxType {
    param([string] $Upn)
    if (-not $Upn) { return $null }
    if ($mailboxCache.ContainsKey($Upn)) { return $mailboxCache[$Upn] }
    try {
        $m = Get-Mailbox -Identity $Upn -ErrorAction Stop
        $mailboxCache[$Upn] = $m.RecipientTypeDetails
    } catch {
        $mailboxCache[$Upn] = $null
    }
    return $mailboxCache[$Upn]
}

$audit = New-Object System.Collections.Generic.List[object]

foreach ($g in $groups) {
    $type = if ($g.GroupTypes -contains 'Unified') { 'Microsoft365' }
            elseif ($g.SecurityEnabled -and $g.MailEnabled) { 'MailEnabledSecurity' }
            elseif ($g.SecurityEnabled) { 'Security' }
            elseif ($g.MailEnabled) { 'Distribution' }
            else { 'Other' }
    Write-Host ("[>] {0,-30} ({1})" -f $g.DisplayName, $type) -ForegroundColor Cyan

    $members = Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue
    foreach ($m in $members) {
        $odt = $m.AdditionalProperties['@odata.type']
        if ($odt -ne '#microsoft.graph.user') {
            $audit.Add([pscustomobject]@{
                GroupName         = $g.DisplayName
                GroupType         = $type
                UserPrincipalName = $m.AdditionalProperties['userPrincipalName']
                DisplayName       = $m.AdditionalProperties['displayName']
                ObjectType        = ($odt -replace '#microsoft.graph.','')
                Enabled           = $null
                Licensed          = $null
                UserType          = $null
                RecipientType     = $null
                Category          = ($odt -replace '#microsoft.graph.','')
                Suggestion        = 'KEEP (non-user object)'
                Action            = ''
            })
            continue
        }

        $u = Get-UserDetail -Id $m.Id
        $rt = if ($u.UPN) { Get-MailboxType -Upn $u.UPN } else { $null }

        $category = if (-not $u.Exists)            { 'Deleted' }
                    elseif (-not $u.Enabled)       { 'Disabled' }
                    elseif ($u.UserType -eq 'Guest') { 'Guest' }
                    elseif (-not $u.Licensed)      { 'Unlicensed' }
                    else                            { 'ActiveLicensed' }

        $suggestion = switch ($category) {
            'Deleted'        { 'REMOVE (deleted)' }
            'Disabled'       { 'REMOVE (disabled)' }
            'Guest'          { 'KEEP (guest)' }
            'ActiveLicensed' { 'KEEP (active+licensed)' }
            'Unlicensed' {
                if ($rt -eq 'SharedMailbox') {
                    if ($RemoveSharedMailboxes -and $type -in @('Distribution','MailEnabledSecurity')) { 'REMOVE (shared mailbox in dist group)' }
                    else { 'KEEP (shared mailbox)' }
                }
                elseif ($u.Name -match '^Ayr[ıi]l[ae]n.*Personel') { 'REMOVE (ayrilan personel)' }
                elseif ($RemoveUnlicensed)                         { 'REMOVE (unlicensed)' }
                elseif ($rt -eq 'UserMailbox')                     { 'REVIEW (unlicensed user mailbox)' }
                else                                                { 'REVIEW (unlicensed)' }
            }
            default          { 'REVIEW' }
        }

        $action = if ($suggestion -like 'REMOVE*') { 'Remove' } else { '' }

        $audit.Add([pscustomobject]@{
            GroupName         = $g.DisplayName
            GroupType         = $type
            UserPrincipalName = $u.UPN
            DisplayName       = $u.Name
            ObjectType        = 'user'
            Enabled           = $u.Enabled
            Licensed          = $u.Licensed
            UserType          = $u.UserType
            RecipientType     = $rt
            Category          = $category
            Suggestion        = $suggestion
            Action            = $action
        })
    }
}

# Tam audit
$auditFile = if ($PSCmdlet.ParameterSetName -eq 'Single') {
    Join-Path $dataDir ("audit-{0}-{1}.csv" -f ($GroupName -replace '[^\w-]+','_'), $stamp)
} else {
    Join-Path $dataDir ("audit-pattern-{0}.csv" -f $stamp)
}
$audit | Export-Csv -Path $auditFile -NoTypeInformation -Encoding utf8BOM

# Cleanup CSV (BulkFromCsv uyumlu — sadece Action prefilled olanlar otomatik islenecek)
$cleanupFile = if ($PSCmdlet.ParameterSetName -eq 'Single') {
    Join-Path $dataDir ("cleanup-{0}-{1}.csv" -f ($GroupName -replace '[^\w-]+','_'), $stamp)
} else {
    Join-Path $dataDir ("cleanup-pattern-{0}.csv" -f $stamp)
}
$audit | Select-Object GroupName, UserPrincipalName, Action, Suggestion, Category, RecipientType, DisplayName |
    Export-Csv -Path $cleanupFile -NoTypeInformation -Encoding utf8BOM

# Ozet
Write-Host ""
Write-Host "=== Ozet ===" -ForegroundColor Yellow
$audit | Group-Object Suggestion | Sort-Object Count -Desc |
    ForEach-Object { Write-Host ("  {0,-40} : {1}" -f $_.Name, $_.Count) }

Write-Host ""
Write-Host "Yazildi:" -ForegroundColor Green
Write-Host "  Tam rapor : $auditFile"
Write-Host "  Cleanup   : $cleanupFile"
Write-Host ""
Write-Host "Sonraki adimlar:" -ForegroundColor Yellow
Write-Host "  1) Cleanup CSV'sini acin: ii '$cleanupFile'"
Write-Host "  2) Action='Remove' satirlarini gozden gecirin (gerekirse Action bosaltin)"
Write-Host "  3) Bos Action satirlarinda eklemek istediginiz olursa Action='Add' yazin"
Write-Host "  4) Calistir: .\Manage-Groups.ps1 -Action BulkFromCsv -CsvPath '$cleanupFile'"
