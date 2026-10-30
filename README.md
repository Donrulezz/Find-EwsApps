# Find-EwsApps

**See which apps still use Exchange Web Services (EWS) in Exchange Online, which of them are blocked right now,
and which stop working when Microsoft shuts EWS down for good on April 1, 2027.**

A single read-only PowerShell script. It changes nothing in your tenant: it only runs `Get-*` cmdlets and Graph GET
requests, then writes an HTML + CSV report with a verdict and the vendor's own Microsoft Graph statement for every app.

## Why now

Microsoft is retiring EWS in Exchange Online (Exchange Team blog, "Exchange Online EWS, your time is almost up"):

- **From October 2026** tenants whose `EwsEnabled` is still empty (Null) are switched to `False`, which blocks EWS for
  every app. The rollout runs until late April 2027. To keep EWS for some apps you set `EwsEnabled` to `True` and list the
  allowed app IDs in `EwsAllowedAppIDs`; since 10 Oct 2026 an empty list with `True` blocks everything.
- **On April 1, 2027** EWS is turned off permanently, with no exceptions. Only apps that moved to Microsoft Graph keep working.

The hard parts: the admin center report shows app IDs, not names; the allow list has to be rebuilt in full on every
change; and you need to know which vendors already have a Graph version. This script does those three things.

## What it reads

| Source | Switch | What you learn | Needs |
|---|---|---|---|
| Tenant settings | always | `EwsEnabled`, the `EwsAllowedAppIDs` allow list, the older user-agent policy, mailboxes with EWS off | `ExchangeOnlineManagement` (recent version), a role that can read the organization config |
| EWS usage export | `-UsageCsv <file>` | Every app that called EWS: calls, SOAP actions, last activity | Export: Microsoft 365 admin center → Reports → Usage → Exchange → **EWS usage** (90 days) |
| Entra app names | `-ResolveNames` | Name and publisher of each app ID | `Microsoft.Graph.Authentication`, `Application.Read.All` |

## Quick start

```powershell
# Once
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser   # only for -ResolveNames

# Download Find-EwsApps.ps1 from this repo, read it, then:
Unblock-File .\Find-EwsApps.ps1

# Tenant state and allow list only
.\Find-EwsApps.ps1

# Everything
.\Find-EwsApps.ps1 -UsageCsv .\EWSWeeklyUsage.csv -ResolveNames
```

Or install it from the [PowerShell Gallery](https://www.powershellgallery.com/packages/Get-EwsAppReadiness), where the same
script is published as **Get-EwsAppReadiness** (the name Find-EwsApps was already taken there):

```powershell
Install-Script -Name Get-EwsAppReadiness -Scope CurrentUser
Get-EwsAppReadiness.ps1 -UsageCsv .\EWSWeeklyUsage.csv -ResolveNames
```

Output: `EwsApps-<date>.html` and `.csv` in the current folder (`-OutputFolder` to change it), plus a printed
`Set-OrganizationConfig -EwsAllowedAppIDs "..."` command that merges your current list with every app in use.
**The command is never run by the script.** Review it first: setting the property replaces the whole list, and
changes take up to 24 hours. Works in Windows PowerShell 5.1 and PowerShell 7.

## Verdicts

| Now | Meaning |
|---|---|
| **BLOCKED NOW** | The app calls EWS but your tenant blocks it (EWS off, empty list, or not on the list) |
| **BREAKS when Microsoft sets EwsEnabled to False** | Your tenant is still Null; put the app on an allow list before Microsoft flips it |
| **Allowed for now** | On the allow list |
| **On the allow list but no calls** | Candidate to remove (but check jobs that run less than monthly: Microsoft's own auto-list only looks at 60 days) |

| April 1, 2027 | Meaning |
|---|---|
| **Vendor has a Graph version** | Update or reauthorize the app, then remove it from the list |
| **STOPS: no Graph version yet** | The vendor still needs EWS: plan the replacement now |
| **STOPS unless it moves to Graph** | No vendor statement in the script: ask the vendor |

The report also flags SOAP actions that map to Graph features Microsoft still lists as gaps (import/export, user
configuration objects, Mark All Items As Read) and warns when the older user-agent policy would still block an app
(for example Teams calendars under `EnforceAllowList`).

## Known apps

Vendor statements were checked on 26 Oct 2026 (sources are linked in the report): Apple Mail/Calendar on macOS,
Microsoft Office (Outlook, classic Outlook for Mac, Skype for Business hybrid), Teams Rooms and devices, the Excel/Power BI
Exchange connector, AvePoint, Hornetsecurity, Quest On Demand Migration, CodeTwo, plus name rules for apps that are
registered per tenant (Veeam, Mimecast, NAKIVO, MigrationWiz, Datto, Barracuda, Exclaimer, Commvault, Druva, Keepit,
SysCloud, CloudAlly, eM Client, Thunderbird, Salesforce, Cisco Unity Connection, DocuWare, MailStore, Spark, Zoho).
A vendor changed its plans? Open an issue with the link and it will be updated within a day.

## Tests

Offline tests with sample data: `powershell -File .\tests\Test-Logic.ps1`. The script has been tested with sample data,
not yet on many real tenants. If something fails on yours, please open an issue with the error message.

## Then fix each app

What to do per app before April 1, 2027: **https://mailpassed.com/ews**

Maintained by [mailpassed.com](https://mailpassed.com), free email-authentication tools (SPF, DKIM, DMARC, MTA-STS).
Also from us: [Find-SmtpAuthSenders](https://github.com/Donrulezz/Find-SmtpAuthSenders) for the SMTP AUTH Basic
authentication change at the end of December 2026.

## License

MIT. See [LICENSE](LICENSE).

## Need it done?
Many apps to sort out, or no time to build the allow list? The [EWS Shutdown Kit ($29)](https://donruler.gumroad.com/l/ews-shutdown-kit)
walks through the allow list without breaking anything, lists 16 apps and vendor groups with their app IDs, maps EWS calls to
Microsoft Graph for in-house scripts, and has ready-to-send vendor messages.
The script stays free.
