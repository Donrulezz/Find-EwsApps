$ErrorActionPreference = 'Stop'
# Offline tests: sample data only, no tenant needed. Run: powershell -File .\tests\Test-Logic.ps1
$script = Join-Path $PSScriptRoot '..\Find-EwsApps.ps1'
$errs = $null; $tokens = $null
[void][System.Management.Automation.Language.Parser]::ParseFile($script, [ref]$tokens, [ref]$errs)
if ($errs.Count) { $errs | ForEach-Object { Write-Output "PARSE ERROR: $($_.Message) at line $($_.Extent.StartLineNumber)" }; exit 1 }
Write-Output 'parse ok'
$nonAscii = (Get-Content $script -Raw).ToCharArray() | Where-Object { [int]$_ -gt 127 }
Write-Output "non-ascii chars: $(@($nonAscii).Count)"

. $script
$fail = 0
function Check($name, $cond) { if ($cond) { Write-Output "PASS $name" } else { Write-Output "FAIL $name"; $script:fail++ } }

$apple = 'f8d98a96-0999-43f5-8af3-69971c7bb423'
$teams = 'cc15fd57-2c6c-4117-a88c-83b1d56b4bbe'
$veeam = '11111111-2222-3333-4444-555555555555'   # per-tenant app: only the Entra name says "Veeam"
$old = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'     # on the list, no calls
$mystery = '99999999-8888-7777-6666-555555555555'  # unknown, not resolved

# Allow list parsing
$list = Split-AppIdList "$($apple.ToUpper()), $old ,not-a-guid,$apple"
Check 'allow list: valid, lower case, deduped' ($list.Count -eq 2 -and $list[0] -eq $apple -and $list[1] -eq $old)
Check 'allow list: empty input' ((Split-AppIdList $null).Count -eq 0)

# Tenant state
Check 'state: null' ((Get-TenantEwsState $null @()).Mode -eq 'Null')
Check 'state: false' ((Get-TenantEwsState $false $list).Mode -eq 'False')
Check 'state: true + empty list' ((Get-TenantEwsState $true @()).Mode -eq 'True-empty')
Check 'state: true + list' ((Get-TenantEwsState $true $list).Mode -eq 'True-list')

# Usage CSV (admin center export shape, one row per app and SOAP action)
$csvPath = Join-Path $env:TEMP 'ews-usage-test.csv'
@"
Application ID,SOAP Action,Call Volume,Last Activity Date (UTC)
$apple,FindItem,500,2026-10-20
$apple,GetItem,300,2026-10-22
$veeam,ExportItems,1200,2026-10-21
$veeam,GetFolder,40,2026-10-19
$teams,GetUserAvailability,50,2026-10-18
$mystery,UpdateUserConfiguration,7,2026-10-10
"@ | Set-Content -Path $csvPath -Encoding UTF8
$usage = ConvertFrom-EwsUsageCsv (Import-Csv $csvPath)
Check 'usage: 4 apps' ($usage.Count -eq 4)
Check 'usage: calls summed per app' ($usage[$apple].Calls -eq 800 -and $usage[$veeam].Calls -eq 1240)
Check 'usage: last activity is the latest date' ($usage[$apple].Last.ToString('yyyy-MM-dd') -eq '2026-10-22')

$names = @{ $veeam = [pscustomobject]@{ DisplayName = 'Veeam Backup for Microsoft 365 - contoso'; Publisher = 'Veeam'; Owner = 'x' } }

# Tenant with an allow list (apple + old)
$state = Get-TenantEwsState $true $list
$rows = Merge-EwsFindings $state $list $usage $names
$rows | Format-Table App, OnAllowList, Calls, Verdict, After2027 -AutoSize | Out-String -Width 260 | Write-Output
$by = @{}; foreach ($r in $rows) { $by[$r.AppId] = $r }
Check 'rows: union of list and usage' ($rows.Count -eq 5)
Check 'apple: allowed now' ($by[$apple].Verdict -eq 'Allowed for now')
Check 'apple: stops 2027, no Graph yet' ($by[$apple].After2027 -like 'STOPS*no Graph version yet*')
Check 'veeam: name rule + blocked now' ($by[$veeam].App -like 'Veeam*' -and $by[$veeam].Verdict -like 'BLOCKED NOW: in use but not on the allow list')
Check 'veeam: import/export gap hint' ($by[$veeam].GraphGaps -like '*import/export*')
Check 'teams: known app, Graph version' ($by[$teams].After2027 -like 'Vendor has a Graph version*')
Check 'old: candidate to remove' ($by[$old].Verdict -like 'On the allow list but no calls*')
Check 'mystery: unknown + user configuration gap' ($by[$mystery].App -like 'Unknown app*' -and $by[$mystery].GraphGaps -like '*user configuration*')
Check 'sort: blocked first, biggest caller first' ($rows[0].AppId -eq $veeam -and $rows[0].Verdict -clike 'BLOCKED NOW*')

# Null tenant: everything in use breaks when Microsoft flips it
$rowsNull = Merge-EwsFindings (Get-TenantEwsState $null @()) @() $usage $names
Check 'null: used apps BREAKS' (@($rowsNull | Where-Object { $_.Verdict -like 'BREAKS*' }).Count -eq 4)

# EWS off: used apps blocked now, ranked above the rest
$rowsOff = Merge-EwsFindings (Get-TenantEwsState $false $list) $list $usage $names
$byOff = @{}; foreach ($r in $rowsOff) { $byOff[$r.AppId] = $r }
Check 'off: apple blocked now' ($byOff[$apple].Verdict -clike 'BLOCKED NOW*')
Check 'off: old app blocked, no calls, ranked last' ($byOff[$old].Verdict -clike 'Blocked (EWS off*' -and $rowsOff[-1].AppId -eq $old)

# Allow-list command: current list first, then every app in use, once each
$cmd = New-AllowListCommand $list $rows
Check 'command merges and dedupes' ($cmd -eq ('Set-OrganizationConfig -EwsAllowedAppIDs "' + (@($apple, $old, $veeam, $teams, $mystery) -join ',') + '"'))
Check 'command: nothing to allow' ($null -eq (New-AllowListCommand @() @()))

# Older user-agent policy
$org = [pscustomobject]@{ EwsApplicationAccessPolicy = 'EnforceAllowList'; EwsAllowList = @('Outlook*') }
$w = Get-PolicyWarnings $org $list $rows
Check 'policy: Teams user agent missing' ($w.Count -eq 1 -and $w[0] -like '*Teams CalendarSkypeSpaces*')
$org2 = [pscustomobject]@{ EwsApplicationAccessPolicy = 'EnforceAllowList'; EwsAllowList = @('Teams CalendarSkypeSpaces/1.0a$*+') }
Check 'policy: Teams user agent present' ((Get-PolicyWarnings $org2 $list $rows).Count -eq 0)

# HTML
$html = ConvertTo-EwsHtmlReport $rows $state $list $cmd $w 'Mailboxes with EWS turned off individually: 3.' @('test')
Check 'html: summary, guide link, command, warning' ($html -match 'only the 2 app' -and $html -match 'mailpassed.com/ews' -and $html -match 'EwsAllowedAppIDs' -and $html -match 'CalendarSkypeSpaces')
$out = Join-Path $env:TEMP 'ews-apps-test.html'; $html | Out-File $out -Encoding utf8
Write-Output "html written: $out"
Write-Output "FAILURES: $fail"
