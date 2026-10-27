<#PSScriptInfo

.VERSION 1.0.0

.GUID 684eb397-a5bb-4a9d-a330-5d0ef0291bb4

.AUTHOR mailpassed.com

.COMPANYNAME mailpassed.com

.COPYRIGHT (c) 2026 mailpassed.com. MIT License.

.TAGS EWS ExchangeWebServices ExchangeOnline Exchange Microsoft365 Office365 MicrosoftGraph AllowList EwsAllowedAppIDs Report

.LICENSEURI https://github.com/Donrulezz/Find-EwsApps/blob/main/LICENSE

.PROJECTURI https://mailpassed.com/ews

.ICONURI

.EXTERNALMODULEDEPENDENCIES ExchangeOnlineManagement, Microsoft.Graph.Authentication

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
1.0.0 (27 Oct 2026): first release. Tenant EWS state, allow list, EWS usage export, app names, vendor readiness and a merged allow-list command.

#>

<#
.SYNOPSIS
    Shows which apps still use Exchange Web Services (EWS) in Exchange Online, whether they are blocked now,
    and which of them stop working when Microsoft shuts EWS down for good on April 1, 2027.

.DESCRIPTION
    READ-ONLY: this script changes nothing in your tenant. It only runs Get-* cmdlets and Graph GET requests.

      1. Tenant state (always): EwsEnabled, the EwsAllowedAppIDs allow list, the older user-agent policy
         (EwsApplicationAccessPolicy / EwsAllowList) and how many mailboxes have EWS turned off.
      2. -UsageCsv: the CSV you export from Microsoft 365 admin center > Reports > Usage > Exchange > EWS usage
         (Application ID, SOAP Action, Call Volume, Last Activity). This is what tells you which apps really call EWS.
      3. -ResolveNames: looks up each app ID in Microsoft Entra (Microsoft Graph, Application.Read.All) to get its
         name and publisher.

    Every app is matched against a list of known apps with the vendor's own statement about Microsoft Graph
    (Apple Mail, Outlook, Teams Rooms, Veeam, Mimecast, AvePoint, CodeTwo, Exclaimer, MigrationWiz and more).

    It writes EwsApps-<date>.csv and EwsApps-<date>.html and prints a Set-OrganizationConfig -EwsAllowedAppIDs
    command that merges your current list with the apps in use. The command is printed, never run: review it,
    because setting the property replaces the whole list.

.EXAMPLE
    .\Find-EwsApps.ps1
    Tenant state only: is EWS on, off or limited to an allow list, and what is on the list.

.EXAMPLE
    .\Find-EwsApps.ps1 -UsageCsv .\EWSWeeklyUsage.csv -ResolveNames
    Everything: every app that called EWS, its name, whether it is blocked now and what happens on April 1, 2027.

.NOTES
    Version 1.0, 27 Oct 2026. Free to use and share (MIT license).
    Requires: ExchangeOnlineManagement module (Install-Module ExchangeOnlineManagement), a recent version so that
    Get-OrganizationConfig has -RetrieveEwsOperationAccessPolicy. For -ResolveNames: Microsoft.Graph.Authentication.
    Timeline: EwsEnabled Null is switched to False from October 2026 (rollout until late April 2027); EWS is turned
    off for good on April 1, 2027 with no exceptions. Guides per app: https://mailpassed.com/ews
#>
[CmdletBinding()]
param(
    [string]$UsageCsv,
    [switch]$ResolveNames,
    [switch]$SkipMailboxes,
    [string]$OutputFolder = (Get-Location).Path,
    [switch]$SkipConnect
)

$GuideUrl = 'https://mailpassed.com/ews'
$MicrosoftTenantId = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'
$TeamsAppId = 'cc15fd57-2c6c-4117-a88c-83b1d56b4bbe'
$GuidPattern = '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'

function New-AppInfo($Name, $Ready, $Status, $Action, $Source) {
    # Ready: $true = the vendor has a Graph version, $false = still needs EWS, $null = depends or unknown.
    return [pscustomobject]@{ Name = $Name; Ready = $Ready; Status = $Status; Action = $Action; Source = $Source }
}

function Get-KnownApps {
    # Fixed app IDs with the vendor's or Microsoft's own statement (checked 26 Oct 2026).
    $apple = New-AppInfo 'Apple Mail / Calendar (macOS)' $false 'Still uses EWS. Apple says Microsoft Graph support comes in a future macOS 27 update.' 'Keep it on the allow list until Apple ships Graph, or move users to the new Outlook for Mac.' 'https://support.apple.com/guide/deployment/dep158966b23'
    $office = New-AppInfo 'Microsoft Office (Outlook desktop, add-ins, classic Outlook for Mac, Skype for Business hybrid)' $null 'Outlook for Windows is ready from build 16.0.20430.20092. Classic Outlook for Mac and Skype for Business Server hybrid still need this app ID.' 'Update Outlook for Windows. Keep the ID on the list while you still run classic Outlook for Mac or Skype for Business hybrid.' 'https://mc.merill.net/message/MC1485116'
    $teams = New-AppInfo 'Microsoft Teams Rooms and Teams devices' $true 'Ready from the device builds listed in MC1486282.' 'Update the devices. If you use EnforceAllowList, also allow the user agent Teams CalendarSkypeSpaces/1.0a$*+' 'https://mc.merill.net/message/MC1486282'
    $excel = New-AppInfo 'Microsoft Power Query for Excel (Exchange Online connector)' $false 'The Power Query Exchange connector retires in April 2027.' 'Move the queries to the Microsoft Graph Mail connector (Microsoft 365 Apps or Office 2024).' 'https://support.microsoft.com/excel/power-query-exchange-connector-changes'
    $avepoint = New-AppInfo 'AvePoint Cloud Backup / Fly' $true 'Graph by default since the August 2026 releases. Archive, group and public folder mailboxes still use EWS.' 'Update. Keep it on the list only if you back up archive, group or public folder mailboxes.' 'https://www.avepoint.com/blog/manage/october-2026-updates-microsoft-ews-retirement'
    $hornet = New-AppInfo 'Hornetsecurity 365 Total Backup' $true 'Mailbox backup over Graph since the 15 Sep 2026 release.' 'Switch the backup to Graph, then remove the app from the list.' 'https://www.hornetsecurity.com/en/release-notes/365-total-backup-2026-09-15/'
    $quest = New-AppInfo 'Quest On Demand Migration' $false 'Still uses EWS; Quest is moving to Graph in stages.' 'Keep it on the list while migrations run; follow Quest KB 4380720.' 'https://support.quest.com/on-demand-migration/kb/4380720/new-email-migration-permissions'
    $c2mig = New-AppInfo 'CodeTwo Office 365 Migration / Backup' $true 'Versions before 3.3.x (migration) and 2.6.x (backup) use EWS.' 'Update to the current version.' 'https://www.codetwo.com/kb/ews-end-of-service-vs-codetwo/'
    $c2sig = New-AppInfo 'CodeTwo Email Signatures 365' $true 'Sent Items Update has a Graph version.' 'Switch Sent Items Update to the Graph version.' 'https://www.codetwo.com/kb/ews-end-of-service-vs-codetwo/'
    return @{
        'f8d98a96-0999-43f5-8af3-69971c7bb423' = $apple
        'd3590ed6-52b3-4102-aeff-aad2292ab01c' = $office
        'cc15fd57-2c6c-4117-a88c-83b1d56b4bbe' = $teams
        'a672d62c-fc7b-4e81-a576-e60dc46e951d' = $excel
        'c21d796d-56a5-4305-a7df-047371fa9fd7' = $avepoint
        'b14c93d6-47fe-4ace-afcd-006fc1fbfabb' = $avepoint
        '34be87c0-bd08-47ab-8ecf-38f6340592f3' = $avepoint
        '832b88db-632e-452e-90f0-b4c769cc787d' = $hornet
        '8de29203-c58a-45e2-ac4f-7a3784b66933' = $quest
        '5adae0aa-b031-421d-a64c-d8041e864be5' = $quest
        '38a32382-73cb-4813-b6d2-96ed716ddecc' = $quest
        '61f709b4-914c-43c2-a25f-bcfe4c90fa7a' = $c2mig
        '3153c72b-35a8-465b-84d8-dfb56651836a' = $c2sig
    }
}

function Get-NameRules {
    # Vendors that register a separate app in every tenant: matched on the Entra display name or publisher.
    return @(
        @{ P = 'Mimecast'; I = (New-AppInfo 'Mimecast (Sync & Recover / Continuity)' $false 'Mimecast says its EWS apps must be on the allow list.' 'Keep it on the list and follow the Mimecast article.' 'https://mimecastsupport.zendesk.com/hc/en-us/articles/55196651141395') }
        @{ P = 'Veeam'; I = (New-AppInfo 'Veeam Backup for Microsoft 365' $false 'Still uses EWS; Veeam waits for Graph delta sync.' 'Keep it on the list and watch Veeam KB4820.' 'https://www.veeam.com/kb4820') }
        @{ P = 'NAKIVO'; I = (New-AppInfo 'NAKIVO Backup & Replication' $true 'Version 11.2.3+ can run Graph only.' 'Update and set system.o365.mailbox.api.mode=graphonly.' 'https://helpcenter.nakivo.com/Knowledge-Base/Content/Best-Practices/Everything-You-Should-Know-About-NAKIVO-and-EWS-API-Deprecation.htm') }
        @{ P = 'MigrationWiz|BitTitan'; I = (New-AppInfo 'BitTitan MigrationWiz' $false 'Still uses EWS.' 'Keep your MigrationWiz Modern Auth app ID on the list while migrations run.' 'https://www.bittitan.com/?p=28334') }
        @{ P = 'Datto'; I = (New-AppInfo 'Datto SaaS Protection' $true 'Moved to Graph; customers had to reauthorize.' 'Reauthorize the Datto app, then remove it from the list.' 'https://saasprotection.datto.com/help/M365/Content/Troubleshooting/General/ews-to-graph.html') }
        @{ P = 'Barracuda'; I = (New-AppInfo 'Barracuda' $null 'Cloud-to-Cloud Backup is on Graph (Reauthorize); Message Archiver / Cloud Archiving still uses EWS.' 'Check which Barracuda product this is.' 'https://documentation.campus.barracuda.com/wiki/spaces/BCCB/pages/560955513') }
        @{ P = 'Exclaimer'; I = (New-AppInfo 'Exclaimer' $true 'Sent Items Update moved to Graph.' 'Reauthorize Exclaimer.' 'https://support.exclaimer.com/hc/en-gb/articles/38864610625181-Exchange-Web-Services-EWS-deprecation') }
        @{ P = 'CodeTwo'; I = (New-AppInfo 'CodeTwo' $true 'Graph versions are available.' 'Update the CodeTwo product.' 'https://www.codetwo.com/kb/ews-end-of-service-vs-codetwo/') }
        @{ P = 'Commvault'; I = (New-AppInfo 'Commvault' $true 'Graph from version 11.44; public folders stay on EWS.' 'Move the Exchange Online protection to Graph.' 'https://documentation.commvault.com/11.44/software/move_your_exchange_online_protection_from_ews_to_microsoft_graph.html') }
        @{ P = 'Druva'; I = (New-AppInfo 'Druva' $null 'Mail, calendar, contacts and tasks on Graph; archive mailboxes and public folders still EWS.' 'Keep it on the list only for archive or public folder backups.' 'https://help.druva.com/en/articles/16160529-exchange-online-ews-retirement-status-and-configuration-guide') }
        @{ P = 'Keepit'; I = (New-AppInfo 'Keepit' $true 'Connectors created before release 10.19 must reauthenticate.' 'Reauthenticate the connector.' 'https://www.keepit.com/help/microsoft-365-category/ews-to-microsoft-graph-migration/') }
        @{ P = 'SysCloud'; I = (New-AppInfo 'SysCloud' $null 'Mailboxes on Graph; in-place archive still pending.' 'Ask SysCloud about archive mailboxes.' 'https://help.syscloud.com/microsoft-is-retiring-exchange-web-services-ews-what-it-means-for-your-syscloud-backup-') }
        @{ P = 'CloudAlly'; I = (New-AppInfo 'CloudAlly' $true 'Graph since May 2026 (reauthorize prompt).' 'Reauthorize CloudAlly.' 'https://support.cloudally.com/hc/en-us/articles/27275914165148-Transition-to-Microsoft-Graph-API') }
        @{ P = 'eM Client'; I = (New-AppInfo 'eM Client' $true 'Graph in version 11; version 10 and older use EWS.' 'Update to eM Client 11.' 'https://www.emclient.com/blogpost.aspx?id=831') }
        @{ P = 'Thunderbird'; I = (New-AppInfo 'Thunderbird' $true 'Graph in release 154 (mail only, no calendar).' 'Update to Thunderbird 154 or later.' 'https://blog.thunderbird.net/2026/09/thunderbird-desktop-new-protocol-support-microsoft-graph-api/') }
        @{ P = 'Salesforce'; I = (New-AppInfo 'Salesforce' $null 'Einstein Activity Capture uses Graph; Lightning Sync retires in April 2027 with no Graph version.' 'Move Lightning Sync users to Einstein Activity Capture or Outlook integration.' 'https://salesforcedictionary.com/blogs/salesforce-ews-retirement-microsoft-graph-migration-2026') }
        @{ P = 'Unity Connection|Cisco Unity'; I = (New-AppInfo 'Cisco Unity Connection' $false 'Graph only in 15 SU4 and later; 14.x gets no fix.' 'Upgrade to Unity Connection 15 SU4+.' 'https://www.cisco.com/c/en/us/support/docs/field-notices/743/fn74365.html') }
        @{ P = 'DocuWare'; I = (New-AppInfo 'DocuWare Connect to Mail' $true 'Moving to Graph.' 'Update Connect to Mail.' 'https://start.docuware.com/blog/product-news/connect-to-mail-ews-to-microsoft-graph-api') }
        @{ P = 'MailStore'; I = (New-AppInfo 'MailStore Server / SPE' $true 'Version 26.2 adds an EWS-to-Graph conversion tool.' 'Update and convert the profiles.' 'https://www.mailstore.com/en/blog/microsoft-is-discontinuing-ews-for-m365/') }
        @{ P = 'Spark|Readdle'; I = (New-AppInfo 'Spark (Readdle)' $false 'Uses EWS; no Graph statement found.' 'Ask Readdle, or move users to another client.' 'https://support.readdle.com/spark/troubleshooting/cannot-add-an-exchange-or-office-365-account') }
        @{ P = 'Skype for Business'; I = (New-AppInfo 'Skype for Business Server hybrid' $false 'Microsoft says a Graph update is upcoming.' 'Keep the IDs on the list (see the Microsoft page).' 'https://learn.microsoft.com/en-us/skypeforbusiness/hybrid/prepare-for-ews-retirement') }
        @{ P = 'Zoho'; I = (New-AppInfo 'Zoho CRM Exchange extension' $null 'Zoho asks to move from Basic auth EWS to OAuth; Graph is not mentioned.' 'Ask Zoho about Graph before April 2027.' 'https://help.zoho.com/portal/en/kb/crm/extensions/overview/articles/migrate-the-microsoft-exchange-for-zoho-crm-extension-from-ews-to-oauth') }
    )
}

function Split-AppIdList($Value) {
    # EwsAllowedAppIDs comes back as one comma-separated string (or a collection); keep valid GUIDs, lower case, once.
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($part in (@($Value) -join ',') -split '[,;\s]+') {
        $id = $part.Trim().ToLowerInvariant()
        if ($id -match $GuidPattern -and -not $out.Contains($id)) { $out.Add($id) }
    }
    return , $out.ToArray()
}

function Get-TenantEwsState($EwsEnabled, $AllowList) {
    $n = @($AllowList).Count
    if ($null -eq $EwsEnabled -or "$EwsEnabled" -eq '') {
        return [pscustomobject]@{ Mode = 'Null'; Summary = 'EwsEnabled is not set (Null): EWS works for every app until Microsoft switches your tenant to False (rollout October 2026 to April 2027). The allow list has no effect yet.' }
    }
    if ("$EwsEnabled" -eq 'False') {
        return [pscustomobject]@{ Mode = 'False'; Summary = 'EwsEnabled is False: EWS is blocked for every app in your tenant.' }
    }
    if ($n -eq 0) {
        return [pscustomobject]@{ Mode = 'True-empty'; Summary = 'EwsEnabled is True but the allow list is empty: every EWS app is blocked (since 10 Oct 2026 the list is required when EwsEnabled is True).' }
    }
    return [pscustomobject]@{ Mode = 'True-list'; Summary = "EwsEnabled is True: only the $n app(s) on the allow list can use EWS." }
}

function Find-CsvColumn($Headers, [string[]]$Patterns) {
    foreach ($p in $Patterns) {
        $hit = $Headers | Where-Object { $_ -match $p } | Select-Object -First 1
        if ($hit) { return $hit }
    }
    return $null
}

function ConvertFrom-EwsUsageCsv($Rows) {
    # Admin center export: Application ID, SOAP Action, Call Volume, Last Activity (UTC). Header names vary, so detect them.
    $result = @{}
    if (-not $Rows) { return $result }
    $headers = $Rows[0].PSObject.Properties.Name
    $idCol = Find-CsvColumn $headers @('^App(lication)? ?ID$', 'App.*Id', 'Application')
    $actCol = Find-CsvColumn $headers @('SOAP', 'Action', 'Operation')
    $volCol = Find-CsvColumn $headers @('Volume', 'Calls', 'Count')
    $dateCol = Find-CsvColumn $headers @('Last Activity', 'Activity', 'Date')
    if (-not $idCol) {
        Write-Warning "EWS usage CSV columns not recognised (found: $($headers -join ', ')). Skipping it."
        return $result
    }
    Write-Verbose "Usage CSV: id='$idCol' action='$actCol' volume='$volCol' date='$dateCol'"
    foreach ($r in $Rows) {
        $id = "$($r.$idCol)".Trim().ToLowerInvariant()
        if ($id -notmatch $GuidPattern) { continue }
        if (-not $result.ContainsKey($id)) { $result[$id] = [pscustomobject]@{ Calls = 0; Actions = @{}; Last = $null } }
        $u = $result[$id]
        $n = 0
        if ($volCol) { [void][long]::TryParse(("$($r.$volCol)" -replace '[^0-9]', ''), [ref]$n) }
        $u.Calls += $n
        if ($actCol -and "$($r.$actCol)".Trim()) {
            $a = "$($r.$actCol)".Trim()
            if ($u.Actions.ContainsKey($a)) { $u.Actions[$a] += $n } else { $u.Actions[$a] = $n }
        }
        $when = [datetime]::MinValue
        if ($dateCol -and [datetime]::TryParse("$($r.$dateCol)", [ref]$when)) {
            if ($null -eq $u.Last -or $when -gt $u.Last) { $u.Last = $when }
        }
    }
    return $result
}

function Get-AppInfo($Id, $Names, $Known, $Rules) {
    if ($Known.ContainsKey($Id)) { return $Known[$Id] }
    $display = ''; $publisher = ''; $owner = ''
    if ($Names -and $Names.ContainsKey($Id)) { $display = $Names[$Id].DisplayName; $publisher = $Names[$Id].Publisher; $owner = $Names[$Id].Owner }
    foreach ($rule in $Rules) {
        if ("$display $publisher" -match $rule.P) { return (New-AppInfo "$($rule.I.Name) ($display)" $rule.I.Ready $rule.I.Status $rule.I.Action $rule.I.Source) }
    }
    if ($display) {
        $label = $display
        if ($owner -eq $MicrosoftTenantId) { $label = "Microsoft: $display" }
        elseif ($publisher) { $label = "$display ($publisher)" }
        return (New-AppInfo $label $null 'No vendor statement in this script.' 'Ask the vendor whether it uses Microsoft Graph before April 1, 2027.' '')
    }
    return (New-AppInfo 'Unknown app (run with -ResolveNames, or look the ID up in Entra > Enterprise applications)' $null 'Unknown.' 'Find the app, then ask its vendor about Microsoft Graph.' '')
}

function Get-GapHints($Actions) {
    # SOAP actions that map to Graph features Microsoft still lists as gaps (Learn deprecation page, 23 Oct 2026).
    $hints = New-Object System.Collections.Generic.List[string]
    foreach ($a in $Actions.Keys) {
        $h = $null
        if ($a -match '^(ExportItems|UploadItems)$') { $h = 'import/export (Graph target Q4 2026)' }
        elseif ($a -match 'UserConfiguration$') { $h = 'user configuration objects (Graph target Q4 2026)' }
        elseif ($a -eq 'MarkAllItemsAsRead') { $h = 'Mark All Items As Read (Graph target Q4 2026)' }
        elseif ($a -match '^(GetDelegate|AddDelegate|UpdateDelegate|RemoveDelegate)$') { $h = 'delegate management (check Graph support)' }
        if ($h -and -not $hints.Contains($h)) { $hints.Add($h) }
    }
    return ($hints -join '; ')
}

function Get-EwsVerdict($Row, $State) {
    $used = ($Row.Calls -gt 0)
    if ($State.Mode -eq 'False') {
        if ($used) { return 'BLOCKED NOW: EWS is off in your tenant (EwsEnabled False)' }
        return 'Blocked (EWS off in your tenant), no recent calls'
    }
    if ($State.Mode -eq 'True-empty') {
        if ($used) { return 'BLOCKED NOW: the allow list is empty' }
        return 'Blocked (allow list empty), no recent calls'
    }
    if ($State.Mode -eq 'True-list' -and $used -and -not $Row.OnAllowList) { return 'BLOCKED NOW: in use but not on the allow list' }
    if ($State.Mode -eq 'Null' -and $used) { return 'BREAKS when Microsoft sets EwsEnabled to False: put it on an allow list now' }
    if ($Row.OnAllowList -and -not $used) { return 'On the allow list but no calls in the report: candidate to remove (check jobs that run less than monthly)' }
    return 'Allowed for now'
}

function Get-After2027($Info) {
    if ($Info.Ready -eq $true) { return 'Vendor has a Graph version: update it, then remove it from the list' }
    if ($Info.Ready -eq $false) { return 'STOPS April 1, 2027: no Graph version yet, plan the replacement' }
    return 'STOPS April 1, 2027 unless it moves to Graph: check with the vendor'
}

function Merge-EwsFindings($State, $AllowList, $Usage, $Names) {
    $known = Get-KnownApps
    $rules = Get-NameRules
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($id in @($AllowList)) { if ($id -and -not $ids.Contains($id)) { $ids.Add($id) } }
    if ($Usage) { foreach ($id in $Usage.Keys) { if (-not $ids.Contains($id)) { $ids.Add($id) } } }
    $rows = foreach ($id in $ids) {
        $u = $null
        if ($Usage -and $Usage.ContainsKey($id)) { $u = $Usage[$id] }
        $info = Get-AppInfo $id $Names $known $rules
        $calls = 0; $last = $null; $actions = @{}
        if ($u) { $calls = $u.Calls; $last = $u.Last; $actions = $u.Actions }
        $top = ($actions.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 5 | ForEach-Object { "$($_.Key) ($($_.Value))" }) -join ', '
        $row = [pscustomobject]@{
            AppId = $id; App = $info.Name; OnAllowList = (@($AllowList) -contains $id); Calls = $calls; LastActivity = $last
            TopSoapActions = $top; GraphGaps = (Get-GapHints $actions); Verdict = ''; After2027 = (Get-After2027 $info)
            VendorStatus = $info.Status; NextStep = $info.Action; Source = $info.Source
        }
        $row.Verdict = Get-EwsVerdict $row $State
        $row
    }
    return @($rows | Sort-Object @{ Expression = { Get-VerdictRank $_.Verdict } }, @{ Expression = { $_.Calls }; Descending = $true }, AppId)
}

function Get-VerdictRank([string]$Verdict) {
    # Case-sensitive on purpose: "BLOCKED NOW" (urgent) and "Blocked (...)" (no calls) differ only by case.
    if ($Verdict -clike 'BLOCKED NOW*') { return 0 }
    if ($Verdict -clike 'BREAKS*') { return 1 }
    if ($Verdict -clike 'Allowed*') { return 2 }
    if ($Verdict -clike 'On the*') { return 3 }
    return 4
}

function New-AllowListCommand($AllowList, $Rows) {
    # Current list plus every app that called EWS. Printed for review only: the cmdlet replaces the whole list.
    $ids = New-Object System.Collections.Generic.List[string]
    foreach ($id in @($AllowList)) { if (-not $ids.Contains($id)) { $ids.Add($id) } }
    foreach ($r in $Rows) { if ($r.Calls -gt 0 -and -not $ids.Contains($r.AppId)) { $ids.Add($r.AppId) } }
    if ($ids.Count -eq 0) { return $null }
    return 'Set-OrganizationConfig -EwsAllowedAppIDs "' + ($ids -join ',') + '"'
}

function Get-PolicyWarnings($OrgConfig, $AllowList, $Rows) {
    $warn = New-Object System.Collections.Generic.List[string]
    $teamsInPlay = (@($AllowList) -contains $TeamsAppId) -or (@($Rows | Where-Object { $_.AppId -eq $TeamsAppId -and $_.Calls -gt 0 }).Count -gt 0)
    $agents = (@($OrgConfig.EwsAllowList) -join ' ')
    if ("$($OrgConfig.EwsApplicationAccessPolicy)" -eq 'EnforceAllowList' -and $teamsInPlay -and $agents -notmatch 'Teams CalendarSkypeSpaces') {
        $warn.Add('EwsApplicationAccessPolicy is EnforceAllowList but EwsAllowList lacks the Teams user agent: add ''Teams CalendarSkypeSpaces/1.0a$*+'' or Teams calendars are blocked (both checks must pass).')
    }
    if ("$($OrgConfig.EwsApplicationAccessPolicy)" -eq 'EnforceBlockList') {
        $warn.Add('EwsApplicationAccessPolicy is EnforceBlockList: user agents on EwsBlockList stay blocked even if their app ID is on the allow list.')
    }
    return , $warn.ToArray()
}

function ConvertTo-EwsHtmlReport($Rows, $State, $AllowList, $Command, $Warnings, $MailboxNote, $Sources) {
    $enc = { param($t) [System.Net.WebUtility]::HtmlEncode("$t") }
    $now = @($Rows | Where-Object { $_.Verdict -like 'BLOCKED NOW*' }).Count
    $soon = @($Rows | Where-Object { $_.Verdict -like 'BREAKS*' }).Count
    $stop = @($Rows | Where-Object { $_.Calls -gt 0 -and $_.After2027 -like 'STOPS*' }).Count
    $yesNo = { param($v) if ($v) { 'Yes' } else { 'No' } }
    $trs = foreach ($r in $Rows) {
        $cls = 'ok'
        if ($r.Verdict -like 'BLOCKED NOW*') { $cls = 'bad' } elseif ($r.Verdict -like 'BREAKS*' -or ($r.Calls -gt 0 -and $r.After2027 -like 'STOPS*')) { $cls = 'warn' }
        $last = ''
        if ($r.LastActivity) { $last = $r.LastActivity.ToString('yyyy-MM-dd') }
        $src = ''
        if ($r.Source) { $src = "<a href=""$(& $enc $r.Source)"">source</a>" }
        "<tr class='$cls'><td>$(& $enc $r.App)<br><span class='muted'>$(& $enc $r.AppId)</span></td><td>$(& $yesNo $r.OnAllowList)</td>" +
        "<td>$(& $enc $r.Calls) $last<br><span class='muted'>$(& $enc $r.TopSoapActions)</span></td><td>$(& $enc $r.Verdict)</td>" +
        "<td>$(& $enc $r.After2027)<br><span class='muted'>$(& $enc $r.GraphGaps)</span></td><td>$(& $enc $r.VendorStatus) $(& $enc $r.NextStep) $src</td></tr>"
    }
    $warnHtml = ($Warnings | ForEach-Object { "<p class='warnbox'>$(& $enc $_)</p>" }) -join "`n"
    $cmdHtml = ''
    if ($Command) { $cmdHtml = "<h2>Allow-list command (review first, not run)</h2><p class='muted'>Setting EwsAllowedAppIDs replaces the whole list. Remove apps you don't recognise, then run it in Exchange Online PowerShell. Changes take up to 24 hours.</p><pre>$(& $enc $Command)</pre>" }
    $date = (Get-Date).ToString('d MMM yyyy')
    return @"
<!doctype html><html><head><meta charset="utf-8"><title>EWS apps - $date</title><style>
body{font-family:Segoe UI,Arial,sans-serif;margin:28px;color:#0f172a;font-size:14px}table{border-collapse:collapse;width:100%}
td,th{border:1px solid #e2e8f0;padding:6px 8px;text-align:left;vertical-align:top}th{background:#f1f5f9}
tr.bad td{background:#fef2f2}tr.warn td{background:#fffbeb}.muted{color:#64748b;font-size:12px}
.warnbox{background:#fffbeb;border:1px solid #fcd34d;padding:8px}pre{background:#f1f5f9;padding:10px;white-space:pre-wrap;word-break:break-all}</style></head><body>
<h1>Exchange Web Services (EWS) apps in Exchange Online</h1>
<p class="muted">$date &middot; sources: $(& $enc ($Sources -join ', ')) &middot; read-only scan</p>
<p><b>$(& $enc $State.Summary)</b></p>
<p><b>$now</b> app(s) blocked now, <b>$soon</b> app(s) that break when Microsoft switches EwsEnabled to False, <b>$stop</b> app(s) in use that
stop on April 1, 2027 unless they move to Microsoft Graph. Apps on the allow list: <b>$(@($AllowList).Count)</b>. $(& $enc $MailboxNote)</p>
$warnHtml
<p>Guides per app and what to do before April 1, 2027: <a href="$GuideUrl">$GuideUrl</a></p>
<table><tr><th>App</th><th>On allow list</th><th>Calls / last / top SOAP actions</th><th>Now</th><th>April 1, 2027</th><th>Vendor statement and next step</th></tr>
$($trs -join "`n")
</table>
$cmdHtml
</body></html>
"@
}

function Invoke-Main {
    $ErrorActionPreference = 'Stop'
    $sources = New-Object System.Collections.Generic.List[string]
    if (-not $SkipConnect) {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            throw 'Install the ExchangeOnlineManagement module first: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
        }
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -ShowBanner:$false
    }
    Write-Host 'Reading the EWS settings of your tenant...'
    $org = Get-OrganizationConfig
    $allow = @()
    try {
        $allow = Split-AppIdList (Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy).EwsAllowedAppIDs
    } catch {
        Write-Warning ('Could not read EwsAllowedAppIDs (' + $_.Exception.Message + '). Update ExchangeOnlineManagement ' +
            '(Update-Module ExchangeOnlineManagement) and use a role that can read the organization configuration.')
    }
    $state = Get-TenantEwsState $org.EwsEnabled $allow
    $sources.Add('tenant settings')

    $mailboxNote = ''
    if (-not $SkipMailboxes) {
        Write-Host 'Counting mailboxes with EWS turned off...'
        $off = @(Get-EXOCASMailbox -ResultSize Unlimited -Properties EwsEnabled | Where-Object { "$($_.EwsEnabled)" -eq 'False' }).Count
        $mailboxNote = "Mailboxes with EWS turned off individually: $off."
    }

    $usage = $null
    if ($UsageCsv) {
        $usage = ConvertFrom-EwsUsageCsv (Import-Csv $UsageCsv)
        $sources.Add("EWS usage report ($($usage.Count) apps)")
    } else {
        Write-Warning ('No -UsageCsv given: showing the allow list only. Export Microsoft 365 admin center > Reports > Usage > ' +
            'Exchange > EWS usage (90 days) and run again with -UsageCsv <file> to see every app that calls EWS.')
    }

    $names = @{}
    $ids = @($allow)
    if ($usage) { $ids += @($usage.Keys) }
    $ids = @($ids | Select-Object -Unique)
    if ($ResolveNames -and $ids.Count -gt 0) {
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            throw 'Install the Microsoft Graph module first: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
        }
        Import-Module Microsoft.Graph.Authentication
        if (-not $SkipConnect) { Connect-MgGraph -Scopes 'Application.Read.All' -NoWelcome }
        Write-Host "Looking up $($ids.Count) app ID(s) in Microsoft Entra..."
        foreach ($id in $ids) {
            $uri = 'v1.0/servicePrincipals?$filter=' + [uri]::EscapeDataString("appId eq '$id'") + '&$select=appId,displayName,publisherName,appOwnerOrganizationId'
            $sp = (Invoke-MgGraphRequest -Method GET -Uri $uri).value | Select-Object -First 1
            if ($sp) { $names[$id] = [pscustomobject]@{ DisplayName = $sp.displayName; Publisher = $sp.publisherName; Owner = "$($sp.appOwnerOrganizationId)" } }
        }
        $sources.Add("Entra app names ($($names.Count))")
    }

    $rows = Merge-EwsFindings $state $allow $usage $names
    $command = New-AllowListCommand $allow $rows
    $warnings = Get-PolicyWarnings $org $allow $rows
    $stamp = (Get-Date).ToString('yyyy-MM-dd')
    $csv = Join-Path $OutputFolder "EwsApps-$stamp.csv"
    $html = Join-Path $OutputFolder "EwsApps-$stamp.html"
    $rows | Select-Object AppId, App, OnAllowList, Calls, LastActivity, TopSoapActions, GraphGaps, Verdict, After2027, VendorStatus, NextStep, Source |
        Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    ConvertTo-EwsHtmlReport $rows $state $allow $command $warnings $mailboxNote $sources | Out-File -FilePath $html -Encoding utf8
    Write-Host ''
    Write-Host $state.Summary
    foreach ($w in $warnings) { Write-Warning $w }
    $need = @($rows | Where-Object { $_.Verdict -like 'BLOCKED NOW*' -or $_.Verdict -like 'BREAKS*' }).Count
    Write-Host "$need app(s) need action now. Report: $html"
    if ($command) {
        Write-Host ''
        Write-Host 'Allow-list command (NOT run; review it first, it replaces the whole list):'
        Write-Host $command
    }
    Write-Host "Guides per app: $GuideUrl"
}

# Dot-sourcing (". .\Find-EwsApps.ps1") only loads the functions, for testing.
if ($MyInvocation.InvocationName -ne '.') { Invoke-Main }
