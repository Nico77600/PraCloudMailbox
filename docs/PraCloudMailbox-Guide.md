---
title: PRA Cloud Mailbox
subtitle: Developer guide
version: 1.0.0
author: Nicolas Fabert
updated: 2026-10-08
---

# PRA Cloud Mailbox — Developer guide

> Exchange disaster recovery, **scenario 2**: Active Directory, Exchange and Entra Connect are lost at the same time. Users and shared mailboxes get a mailbox in **Exchange Online on their existing identities** — the Teams storage of each user becomes his mailbox, so Teams keeps working and nothing is lost — and **everything is rolled back** when the on-premises infrastructure is rebuilt, without losing the mails received in the cloud.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.

> [!NOTE]
> This is the **developer guide**: the scenario, the procedure validated in the lab and the states of a mailbox, the prerequisites in detail (app registration, Security & Compliance, retention), the configuration, every action step by step, the reports, the architecture, the databases, the tests and how to evolve the tool. For the prerequisites and the everyday commands only, read the [user guide](PraCloudMailbox-UserGuide.md).

```cards
database | Collected before the disaster | `Collect` reads the mailboxes, their addresses and the **shared mailbox permissions** from Exchange into a SQLite snapshot, every day: after the disaster nothing on-premises can be read.
check | Preview first, then one confirmation | Without `-Mode Apply` nothing is changed. `Check` is always read-only and says, object by object, whether `Convert` can run.
refresh | Reversible, journaled | Every change is written to a journal under a **batch ID** with the original state of each object; `Recover -Batch` puts it back and resumes where it stopped.
shield | No mail lost | The Teams storage becomes the cloud mailbox; at `Recover` it stays under an **eDiscovery case hold**, and the shared mailboxes become **inactive mailboxes** under the hold, then under retention.
```

## Quick start

```steps
Install | Copy the folder on an Exchange server and on the cloud admin server, unblock the files, install the two PowerShell modules on the cloud admin server (chapter 5).
Configure | Fill in `config\PraCloudMailbox.config.psd1`: scope, Exchange, tenant, app registration, licences, case hold policy, Entra Connect (chapter 6).
Collect every day | `.\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force` in Windows PowerShell 5.1 on the Exchange server (scheduled task), then copy `data\PraCloudMailbox.db` to the cloud admin server (chapter 7).
Check | `.\Invoke-PraCloudMailbox.ps1 -Action Check` in PowerShell 7 on the cloud admin server: every object must be ready (chapter 8).
Disaster | `.\Invoke-PraCloudMailbox.ps1 -Action Convert`, then `-Mode Apply`: note the batch ID (chapter 9).
Infrastructure rebuilt | `.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch <ID>`, then `-Mode Apply` (chapter 10).
```

> [!CAUTION]
> Convert and Recover change production objects in Entra ID and Exchange Online. Run the preview, read the plan, test the whole cycle in a lab with your own tenant settings, and keep the journal (`data\PraCloudMailbox-journal.db`) with the logs: it is the record of what was changed.

# Part I · Understand

<!-- icon: target -->
## 1. Purpose

In a hybrid organisation, a mailbox hosted on-premises is only a *mail user* for Exchange Online. **Scenario 1** — the Exchange servers are lost but Active Directory and Entra Connect work — is covered by [PRA Remote Mailbox](https://github.com/Nico77600/PraRemoteMailbox): it rewrites the Exchange attributes in AD and lets Entra Connect provision new cloud mailboxes.

**Scenario 2** is worse: Active Directory, Exchange **and** Entra Connect are lost (ransomware, loss of the site). The users can still sign in to Microsoft 365 — their Entra ID objects and password hashes are in the cloud — and Teams keeps working, but their mailboxes are unreachable and nothing can be written in AD any more. *PRA* stands for *Plan de Reprise d'Activité*, the disaster recovery plan.

PRA Cloud Mailbox gives every user and shared mailbox in scope a mailbox in Exchange Online **on the same identity** (same objectId, same SMTP and X500 addresses), entirely from the cloud, and rolls everything back once AD, Exchange and Entra Connect are rebuilt. No data is migrated: the cloud mailboxes start with what the cloud already holds — for a user, his Teams storage.

Why a snapshot: after the disaster nothing on-premises can be read. The list of the objects, their GUIDs, their X500 addresses and the permissions of the shared mailboxes (FullAccess with AutoMapping, SendAs, SendOnBehalf) must therefore be collected **before**, regularly, and stored in a file that is copied to the cloud admin server.

<!-- icon: flow -->
## 2. How it works

```flow
database | Collect | Exchange server, every day: SQLite snapshot
arrow | copy |
check | Check | cloud admin server: is every object ready?
arrow | disaster |
cloud | Convert | users and shared mailboxes to Exchange Online
arrow | rebuilt |
refresh | Recover | back to AD and the on-premises mailboxes
```

| Action | When | Where | Edition | Writes |
|---|---|---|---|---|
| `Collect` | before the disaster, every day | Exchange server (or the management tools, or a remote session) | Windows PowerShell 5.1 | the SQLite snapshot only (never Exchange) |
| `Check` | any time, and before Convert | cloud admin server | PowerShell 7.4+ | nothing |
| `Convert` | disaster | cloud admin server | PowerShell 7.4+ | Entra ID (source of authority, usageLocation, licences), Exchange Online (shared mailboxes, permissions, tag), the journal |
| `Recover` | infrastructure rebuilt | cloud admin server | PowerShell 7.4+ | Entra ID (source of authority back to AD, licences, shared identities deleted), case hold, Entra Connect cycles, the journal |

- **One script, one configuration file**, the same on both sides. The scope is the configuration (an OU, the members of a group, a CSV list, or every mailbox), one object (`-Identity`) or a kind (`-Scope UsersOnly | SharedOnly`).
- **Preview by default.** Without `-Mode Apply` nothing is changed. With `-Mode Apply`, Convert and Recover show the plan and ask **one** confirmation (`-Force` for unattended runs).
- **One batch ID** links Convert and Recover: Convert prints it with the exact next command; `-Action Recover -Batch <ID>` rolls that batch back. Running the same command again resumes what is not finished.
- **Every run** prints a banner, numbered steps and a summary card, and writes a log, a transcript and a CSV + HTML report. Exit codes: `0` done, `1` failed (or an object is not ready), `2` done, next step required.

<!-- icon: checklist -->
## 3. The validated procedure

Validated step by step in a lab (Exchange Server SE, Entra Connect, Microsoft 365 E5 tenant) from 2 to 8 October 2026 (appendix B). Convert and Recover implement exactly this order; every step checks the current state first, so a step already done is not repeated.

### 3.1 Users

| Step | Action | Result |
|---|---|---|
| Convert 1 | `Set-MailUser -ExchangeGuid` empty | The MailUser no longer points to the on-premises mailbox |
| Convert 2 | Source of authority to the cloud (`onPremisesSyncBehavior.isCloudManaged = true`) | Entra Connect ignores the object (20 to 30 s) |
| Convert 3 | usageLocation set from the cloud when needed | The value synchronised from AD can be refused for a licence right after the transfer |
| Convert 4 | Exchange plan: licence group (cloud, or synchronised group switched to the cloud), Kiosk plan of the Teams licence, or direct licence | The **Teams storage (ComponentShared) becomes the primary mailbox** in 45 s to 2 min. Same objectId, SMTP and X500 addresses; the Teams chats and their compliance copies are kept |
| Recover 1 | Case hold lifted if present, source of authority back to AD, one delta cycle | Still a cloud mailbox (the plan is still there): no address gap |
| Recover 2 | Exchange plan removed (licences put back exactly as before Convert) | MailUser with the on-premises ExchangeGuid in 2 to 70 s; the cloud mailbox goes back to ComponentShared **with its content** |
| Recover 3 | eDiscovery case hold on the user, checked in `InPlaceHolds` | The mails received during the disaster stay protected and searchable (about 90 s) |
| End | Licence group back to AD, once no converted user of any batch is left | |

A hold on the cloud mailbox (eDiscovery, Litigation or retention) **blocks Recover 2**: Exchange refuses to give the mailbox back to the on-premises object. The tool lifts its own case hold for the switch and puts it back right after; any other hold stops the user with a message (remove it first). For the same reason a user never carries the retention tag.

### 3.2 Shared mailboxes

| Step | Action | Result |
|---|---|---|
| Convert 1 | `Set-MailUser -ExchangeGuid` empty, source of authority to the cloud | |
| Convert 2 | usageLocation set from the cloud | |
| Convert 3 | Temporary licence (Exchange plan only) | Mailbox in 25 to 70 s |
| Convert 4 | `Set-Mailbox -Type Shared`, **wait until it is a SharedMailbox**, tag (`CustomAttribute1 = Converted`), permissions of the snapshot, then the licence put back as before | Removing the licence before the conversion is effective disables the mailbox at once. One unit is enough: the shared mailboxes are converted one after another |
| Recover 1 | eDiscovery case hold, checked in `InPlaceHolds` | About 90 s |
| Recover 2 | Entra Connect scheduler paused, Entra ID identity deleted | **The mailbox becomes inactive** in 16 to 48 s, with its content, the hold and the tag |
| Recover 3 | Identity deleted permanently (Entra ID recycle bin) | Otherwise the next synchronisation **restores** the deleted object (same immutableId) |
| Recover 4 | Scheduler resumed (always, even after an error), one delta cycle | Entra Connect creates a **new object** (same immutableId, new objectId): MailUser with the on-premises GUID, mail flows on-premises |
| Later | The retention policy (adaptive scope `IsInactiveMailbox` + tag) picks up the inactive mailbox; the case hold entry can then be removed | Chapter 11 |

### 3.3 States of a mailbox

| | Before the disaster | After Convert | After Recover |
|---|---|---|---|
| **User** — Active Directory | mailbox on-premises | lost | rebuilt (same objectGUID), mailbox on-premises |
| **User** — Entra ID | synchronised from AD | **cloud-managed** | synchronised from AD again |
| **User** — Exchange Online | MailUser (on-premises GUID) + Teams storage (ComponentShared) | **UserMailbox** = the former Teams storage | MailUser (on-premises GUID) + ComponentShared **with the disaster mails, under the case hold** |
| **User** — mail delivered to | on-premises | Exchange Online | on-premises |
| **Shared** — Active Directory | shared mailbox on-premises | lost | rebuilt, shared mailbox on-premises |
| **Shared** — Entra ID | synchronised from AD | **cloud-managed** | **new object** created by Entra Connect (same immutableId) |
| **Shared** — Exchange Online | MailUser (on-premises GUID) | **SharedMailbox**, tagged, permissions granted | new MailUser (on-premises GUID); the cloud mailbox is **inactive** under the case hold, then under retention |
| **Shared** — mail delivered to | on-premises | Exchange Online | on-premises |

The mails received in the cloud during the disaster are **never deleted**: those of the users stay in their ComponentShared storage under the case hold, those of the shared mailboxes in their inactive mailboxes. They are found with eDiscovery (chapter 11). Copying them back into the on-premises mailboxes is not part of the tool.

### 3.4 What the tool never does

- It never writes to Exchange on-premises or to Active Directory (Collect only reads; everything else happens in the cloud).
- It never deletes a user mailbox, and never runs `Set-User -PermanentlyClearPreviousMailboxInfo` (it empties the mails of the cloud mailbox 3 to 6 hours later).
- It never deletes a shared mailbox identity before the case hold is stamped on the mailbox (`InPlaceHolds`): otherwise the mailbox would be soft-deleted, not inactive, and purged after 30 days.
- It never lifts a hold from an object that is already back on-premises.
- It never converts a room or equipment mailbox (`KIND_NOT_SUPPORTED`: their booking settings are not collected and they were not validated in the lab).

# Part II · Set up

<!-- icon: checklist -->
## 4. Prerequisites

### 4.1 Collect (Exchange side)

| Item | Requirement |
|---|---|
| Computer | An Exchange server or a computer with the Exchange management tools (`Exchange.ConnectionMode = 'Local'`), or any domain computer with PowerShell remoting to an Exchange server (`'Remote'`, Kerberos) |
| PowerShell | **Windows PowerShell 5.1** 64-bit (`powershell.exe`), where the Exchange cmdlets live |
| Account | Allowed to read recipients and permissions: *Organization Management*, or a role group with *View-Only Organization Management* and *Active Directory Permissions* (`Get-ADPermission`). The AutoMapping list (`msExchDelegateListLink`) is read in AD with the same account |
| Software | Nothing to install: the SQLite engine is in `lib\sqlite` |

### 4.2 Cloud actions (cloud admin server)

| Item | Requirement |
|---|---|
| Computer | Any Windows computer that reaches Microsoft 365 — it must **not depend on the on-premises AD** (it is lost) |
| PowerShell | **PowerShell 7.4** or later (`pwsh.exe`) |
| Modules | `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` 3.10 or later |
| App registration | Certificate sign-in; the certificate with its private key in `Cert:\CurrentUser\My` or `Cert:\LocalMachine\My` (chapter 4.3) |
| Licences | Group mode: a licence group whose licence includes an Exchange Online plan, and enough free units for the users. Kiosk mode: the Exchange Kiosk plan inside the licence the users already have (e.g. Teams Enterprise). Direct mode: free units of `Licensing.Users.SkuPartNumber`. Shared mailboxes: **one** free unit of `Licensing.Shared.SkuPartNumber` (any SKU with an Exchange Online plan) |
| Holds | An eDiscovery case with a case hold policy (`Retention.HoldPolicy`, chapter 4.4), and a retention policy for the inactive shared mailboxes |
| Entra Connect | Recover only: Entra Connect **2.5.76.0** or later on the rebuilt server (it keeps the cloud source of authority of the objects), reachable as described in chapter 6.1 |

> [!IMPORTANT]
> Recover needs the on-premises objects with **their original objectGUID** (AD restored from a backup, not recreated): the immutableId links each Entra ID object to its AD account.

### 4.3 App registration

Create an app registration for the tool (Entra admin center › *App registrations* › *New registration*, single tenant), upload the public part of its certificate, then:

| Where | What |
|---|---|
| Microsoft Graph, application permissions | `User.ReadWrite.All` (usageLocation, deletion of a shared identity), `User-OnPremisesSyncBehavior.ReadWrite.All` (source of authority), `LicenseAssignment.ReadWrite.All` (licences), `Organization.Read.All` (licence counters). Group mode: `GroupMember.ReadWrite.All`, and `Group-OnPremisesSyncBehavior.ReadWrite.All` when the licence group is synchronised from AD |
| Office 365 Exchange Online, application permission | `Exchange.ManageAsApp` |
| Admin consent | *Grant admin consent* for all of them |
| Entra ID role | **Exchange Administrator** assigned to the application |
| Microsoft Purview | The service principal in the **eDiscovery Manager** role group (case hold of Recover) |

Check reads the Graph permissions from the token and names the missing ones (`APP_PERMISSION`). The Purview role group is set once by an administrator, in Security & Compliance PowerShell (the *ObjectId* is the one of the **Enterprise application**, not of the app registration):

```powershell
Connect-IPPSSession -UserPrincipalName admin@contoso.com
New-ServicePrincipal -AppId <application-client-id> -ObjectId <enterprise-application-object-id> -DisplayName 'PRA Cloud Mailbox'
Add-RoleGroupMember -Identity eDiscoveryManager -Member <enterprise-application-object-id>
```

### 4.4 Case hold and retention

Created once, before the disaster, by a compliance administrator:

```powershell
Connect-IPPSSession -UserPrincipalName admin@contoso.com

# eDiscovery case hold used by Recover (users and shared mailboxes): Retention.HoldPolicy = 'PRA-Recover-Hold'
New-ComplianceCase -Name 'PRA Recover'
New-CaseHoldPolicy -Name 'PRA-Recover-Hold' -Case 'PRA Recover' -Enabled $true
New-CaseHoldRule -Name 'PRA-Recover-Hold-Rule' -Policy 'PRA-Recover-Hold'      # no query: everything is held

# Retention of the inactive shared mailboxes (tag written by Convert: Retention.TagAttribute / TagValue)
New-AdaptiveScope -Name 'PRA-Converted-Inactive' -LocationType User -RawQuery '(IsInactiveMailbox -eq "True") -and (CustomAttribute1 -eq "Converted")'
New-RetentionCompliancePolicy -Name 'PRA-Converted-Retention' -AdaptiveScopeLocation 'PRA-Converted-Inactive' -Applications 'User:Exchange' -Enabled $true
New-RetentionComplianceRule -Name 'PRA-Converted-Retention-Rule' -Policy 'PRA-Converted-Retention' -RetentionDuration Unlimited -RetentionComplianceAction Keep
```

- The case hold protects at once (about 90 s per object); the adaptive scope needs **up to 5 days** to pick up a new inactive mailbox. The case hold covers the meantime.
- An eDiscovery case hold holds **at most 1,000 mailboxes**: beyond that, roll back in waves with one configuration file per wave and its own `Retention.HoldPolicy`.
- The adaptive scope never catches a mailbox that is already deleted without a hold: this is why Recover puts the case hold **before** deleting a shared identity.
- Retention policies applied to **every** mailbox of the organisation would also hold the users' cloud mailboxes and block their Recover: Check warns about them (`ORG_HOLD`).

<!-- icon: download -->
## 5. Installation

1. Copy the folder (release zip, or `git clone`) to `C:\PRA\PraCloudMailbox` on the Exchange server **and** on the cloud admin server, and unblock the files.
2. On the cloud admin server, install the modules (an administrator PowerShell 7):

```powershell
Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force
Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.0 -Scope AllUsers -Force
```

`-Force` also updates or reinstalls a module that is already installed. If an older version still conflicts, close every PowerShell window, open a new one as administrator, run `Uninstall-Module <ModuleName> -AllVersions -Force`, then the `Install-Module` command again.

3. Import the certificate of the app registration (with its private key) in `Cert:\LocalMachine\My` or in `Cert:\CurrentUser\My` of the account that runs the tool.
4. Fill in the configuration (chapter 6) and copy it to the other side: **the same file** is used by Collect and by the cloud actions.

| Folder | Content |
|---|---|
| `Invoke-PraCloudMailbox.ps1` | the tool (the only script to run) |
| `module\` | `PRA2.Common` (console, log, configuration, reports), `PRA2.Store` (SQLite), `PRA2.Collect` (Exchange, 5.1), `PRA2.Cloud` (Graph, Exchange Online, holds, Entra Connect) |
| `lib\sqlite\` | System.Data.SQLite for Windows PowerShell 5.1 (`net46`) and PowerShell 7 (`netstandard2.0`), x64 |
| `config\` | the configuration, `Targets.sample.csv` (scope `Csv`), `EntraConnect.sample.ps1` (Entra Connect `Script` mode) |
| `templates\` | the HTML report template |
| `data\`, `logs\`, `reports\` | created at run time: snapshots and journal, logs and transcripts, CSV and HTML reports |

<!-- icon: settings -->
## 6. Configuration

`config\PraCloudMailbox.config.psd1` is a PowerShell data file. An unknown setting is refused (spelling mistakes are detected); relative paths are relative to the tool folder.

| Section | Settings |
|---|---|
| `Environment` | Label used in the backup file names |
| `Exchange` | `ConnectionMode` (`Local`, `Remote`), `Server` (Remote), `DomainController` (the same DC for every read of a run) |
| `Scope` | `Mode` (`Auto`, `OU`, `Group`, `Csv`), `SearchBase`, `GroupDN`, `CsvPath` (column `Identity`: UPN, SMTP, sAMAccountName, DN or GUID), `IncludeUsers`, `IncludeShared`, `IncludeRoom`, `IncludeEquipment`, `ExcludeSamAccountNames` |
| `Collect` | `SharedPermissions`, `ExpandGroupTrustees` (groups used as trustees: members stored and granted one by one), `ExcludeTrustees` (wildcards), `MailboxStatistics`, `Contacts`, `ContactsSearchBase`, `DistributionGroups`, `DistributionGroupsSearchBase`, `DynamicDistributionGroups` |
| `Store` | `Path` (snapshots), `KeepSnapshots`, `BackupFolder` (a consistent copy after each Collect), `MaxSnapshotAgeDays` (Check warns beyond), `JournalPath` (journal of Convert and Recover, cloud side) |
| `Cloud` | `TenantId`, `Organization`, `AppId`, `CertificateThumbprint`, `DefaultUsageLocation` (two letters, for the objects without usageLocation) |
| `Polling` | `IntervalSeconds`, `MailboxTimeoutMinutes`, `SyncTimeoutMinutes`, `HoldTimeoutMinutes` |
| `Licensing` | `Users.Mode` (`Group`, `Kiosk`, `Direct`), `Users.GroupId` (object ID of the licence group), `Users.SkuPartNumber` (Direct), `Shared.SkuPartNumber` (temporary licence) |
| `Retention` | `TagAttribute`, `TagValue` (written on the shared mailboxes), `HoldPolicy` (case hold policy of Recover) |
| `EntraConnect` | `Mode` (`Remoting`, `Script`, `Manual`), `Server`, `ScriptPath`, `MinVersion` (chapter 6.1) |
| `Logging`, `Report` | Folders; `Report.Enabled` |

System mailboxes (HealthMailbox, SystemMailbox, DiscoverySearchMailbox...) and system or administration trustees (NT AUTHORITY, SIDs, Exchange and domain admin groups) are always excluded.

**Choosing `Licensing.Users.Mode`**

| Mode | What Convert does | When |
|---|---|---|
| `Group` | adds the user to `GroupId`. A group synchronised from AD is switched to cloud management first, once; it goes back to AD with the last user of the journal | a licence group already exists (the usual case) |
| `Kiosk` | enables the Exchange Kiosk plan inside the licence the user already has (no extra unit; 2 GB mailbox, no Outlook desktop) | users with Teams Enterprise or another licence that carries `EXCHANGE_S_DESKLESS` |
| `Direct` | assigns `SkuPartNumber` to the user, or enables its Exchange plan when the user already holds the SKU with Exchange disabled (typical hybrid user: no new unit) | no licence group |

In every mode Recover puts the licences back **exactly** as they were (direct assignments and disabled plans are recorded by Convert).

### 6.1 Entra Connect

Recover needs, on the rebuilt Entra Connect server: delta cycles, and the scheduler paused while shared identities are deleted (a cycle at the wrong moment restores a deleted object).

| `EntraConnect.Mode` | How | Requirements |
|---|---|---|
| `Remoting` | `Invoke-Command -ComputerName <Server>` with the current account, ADSync cmdlets | WinRM to the server, account in `ADSyncAdmins` |
| `Script` | `<ScriptPath> -Operation Pause \| Resume \| Delta` | any access path (another account, a jump host, Azure Run Command...). Output: one line of state, logged. `exit 0` = done, any other exit code or an exception = failed |
| `Manual` | the tool asks the operator, who confirms | interactive console (not with `-Force`) |

`config\EntraConnect.sample.ps1` implements the `Script` contract with PowerShell remoting and a credential saved with `Export-Clixml`: copy it and adapt the server name.

# Part III · Use

<!-- icon: database -->
## 7. Collect

```powershell
# Windows PowerShell 5.1 on the Exchange server
.\Invoke-PraCloudMailbox.ps1 -Action Collect                            # preview: what would be collected
.\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force         # writes a snapshot
.\Invoke-PraCloudMailbox.ps1 -Action Collect -Identity compta@contoso.com   # one object (preview)
```

Steps: Exchange connection, mailboxes in scope, shared mailbox permissions, contacts and distribution groups, saving the snapshot. A snapshot is written with the status *Running* and marked *Complete* at the end: an interrupted Collect never replaces the last good snapshot. A snapshot of one object (`-Identity -Mode Apply`) asks for a confirmation, because it becomes the last snapshot.

**Scheduling**: a daily scheduled task running `powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PRA\PraCloudMailbox\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force` under the collect account, then a copy of `data\PraCloudMailbox.db` (or of `data\backup`) to the cloud admin server. Check warns when the last snapshot is older than `Store.MaxSnapshotAgeDays`.

<!-- icon: check -->
## 8. Check

```powershell
# PowerShell 7 on the cloud admin server
.\Invoke-PraCloudMailbox.ps1 -Action Check
.\Invoke-PraCloudMailbox.ps1 -Action Check -Scope SharedOnly
.\Invoke-PraCloudMailbox.ps1 -Action Check -Identity compta@contoso.com
.\Invoke-PraCloudMailbox.ps1 -Action Check -Snapshot 12
```

![Check: tenant prerequisites, then one line per object](images/pra-console-check.png)

Steps: snapshot, Microsoft 365 connection, tenant prerequisites, objects. An object with an **Error** is not ready (exit code 1) and Convert would refuse it; warnings do not change the exit code. Run it after each change of the configuration, of the licences or of the scope, and regularly (for example weekly) to catch drift.

| Code | Level | Meaning |
|---|---|---|
| `APP_PERMISSION` | Error | A Graph application permission is missing |
| `LICENCE_GROUP` | Error | Licence group missing or without an Exchange plan |
| `LICENCE_CAPACITY` | Error / Ok | Free units for the objects that will take a **new** unit (a user who already holds the SKU, Kiosk mode and shared mailboxes already converted take none) |
| `LICENCE_GROUP_SOA` | Info | Licence group synchronised from AD: switched to the cloud at Convert, back to AD with the last user |
| `SHARED_LICENCE` | Error / Ok | Temporary licence of the shared mailboxes: SKU, Exchange plan, one free unit |
| `ENTRA_NOT_FOUND` | Error | No Entra ID object (neither by onPremisesImmutableId nor by UPN) |
| `ENTRA_MATCHED_BY_UPN` | Warn | Found by UPN only: check the source anchor |
| `ENTRA_NOT_SYNCED` | Error | Cloud-only object |
| `ALREADY_CLOUD_MANAGED` | Warn | Source of authority already in the cloud: a Convert already started |
| `PROVISIONING_ERROR` | Warn | Unresolved provisioning errors in Entra ID |
| `EXO_NOT_FOUND` | Error | No Exchange Online recipient |
| `EXO_MAILUSER` | Ok | MailUser with the on-premises ExchangeGuid |
| `GUID_CLEARED` / `GUID_MISMATCH` | Warn | ExchangeGuid empty (a Convert started) or different from the snapshot |
| `ALREADY_CLOUD_MAILBOX` | Error | Already a cloud mailbox |
| `KIND_NOT_SUPPORTED` | Error | Room or equipment mailbox |
| `TEAMS_STORAGE` / `NO_TEAMS_STORAGE` | Info | The ComponentShared storage will be promoted, or a new empty mailbox created |
| `COMPONENT_SHARED` | Info | Shared mailbox with a ComponentShared storage: the temporary licence promotes it, with its content |
| `EXCHANGE_PLAN_PRESENT` | Error | An Exchange Online mailbox plan is already enabled (by plan ID; Foundation does not count): clearing the ExchangeGuid would not create the mailbox while the plan stays. Remove the plan before the disaster or exclude the object |
| `ORG_HOLD` | Warn | Retention policies applied to every mailbox: they would also hold the users' cloud mailboxes and block their Recover |
| `NO_USAGE_LOCATION` | Info / Error | usageLocation empty: set from `Cloud.DefaultUsageLocation` at Convert, or an error when that setting is empty |
| `KIOSK_AVAILABLE` / `NO_KIOSK_PLAN` | Ok / Error | Kiosk mode: the Exchange Kiosk plan is in the user's licences, or not |
| `HOLD_PRESENT` | Warn (user) / Info (shared) | Hold on the object (it would block the Recover of a user) |
| `TAG_ALREADY_SET`, `USER_TAGGED` | Warn | Retention tag already set; a user must not carry it |
| `TRUSTEES_OK` / `TRUSTEE_UNKNOWN` | Ok / Warn | Shared mailbox permissions: every trustee (or group member) is known in Exchange Online, or not |

<!-- icon: cloud -->
## 9. Convert

```powershell
.\Invoke-PraCloudMailbox.ps1 -Action Convert                                   # preview: the plan, nothing is changed
.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply                       # one confirmation, then the batch
.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Scope UsersOnly -Force
.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Batch 5b21030d       # resume a batch
```

![Convert -Mode Apply: users, then shared mailboxes one after another](images/pra-console-convert.png)

Steps: snapshot, Microsoft 365 connection, readiness (the rules of Check), plan and confirmation, users (source of authority and Exchange plan), users (cloud mailboxes), shared mailboxes.

- **Refused objects**: an object with a Check error is not converted; the others are. Licence capacity is checked first: not enough free units = nothing is changed.
- **Batch**: Apply creates a batch in the journal (chapter 14) with the original state of every object — source of authority, ExchangeGuid, usageLocation, recipient type, tag, holds, direct licences with their disabled plans. Its ID is printed at the end with the next command, and is the only input of Recover.
- **Users**: all the identity steps first (GUID cleared, source of authority, usageLocation, plan), then the tool polls every `Polling.IntervalSeconds` until each one is a UserMailbox (`MailboxTimeoutMinutes`). A user still waiting at the timeout is *Pending*: run the same command with `-Batch` later.
- **Licence group synchronised from AD** (Group mode): its source of authority goes to the cloud once, before the first user; the journal records it.
- **Shared mailboxes**: one after another, so one unit of `Licensing.Shared.SkuPartNumber` is enough. Before each one, the tool checks that a unit is free: a shared mailbox that failed keeps its temporary licence and the next ones are not started (never left half-converted). The licence assignment is retried for 45 s (it can be refused just after the usageLocation change). The permissions of the snapshot are granted and read back: FullAccess with AutoMapping as collected, SendAs, SendOnBehalf; group trustees as their members.
- **Resume** (`-Batch`): only the objects not *Done* are taken again; objects that already hold their licence do not count in the capacity check.

> [!TIP]
> Converting in **waves** (for example the managers first) is supported: each Convert creates its own batch, and each batch is rolled back on its own. The licence group goes back to AD only with the last converted user of all the batches of the journal.

<!-- icon: refresh -->
## 10. Recover

```powershell
.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 5b21030d                   # preview
.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 5b21030d -Mode Apply
.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 5b21030d -Mode Apply -Identity compta@contoso.com
```

![Recover -Mode Apply: users back on-premises, shared mailboxes inactive and recreated by Entra Connect](images/pra-console-recover.png)

Run it once AD, Exchange and Entra Connect are rebuilt and synchronising. Steps: snapshot, Microsoft 365 and Security & Compliance connections, prerequisites (case hold policy, Entra Connect), plan and confirmation, users (source of authority back to AD), users (back on-premises), shared mailboxes, licence group. Each object goes the way its **current state** allows:

| State found | What Recover does |
|---|---|
| User with a cloud mailbox | Case hold lifted if present (any other hold stops the user: remove it first), source of authority back to AD, one delta cycle (skipped when already synchronised), Exchange plan removed (licences as before Convert), wait for a MailUser with the on-premises ExchangeGuid, case hold added and checked. The cloud mailbox stays as ComponentShared with its content, under the hold |
| Already back on-premises (synchronised MailUser, on-premises GUID) | Case hold added if missing and checked; **never lifted** |
| SharedMailbox | Case hold checked in `InPlaceHolds` (otherwise nothing is deleted), scheduler paused, identity deleted, inactive mailbox checked, identity deleted permanently, scheduler resumed (always, even after an error), one delta cycle, new object checked as a MailUser with the on-premises GUID |
| Shared identity deleted by an earlier run | Recycle bin checked (purged if needed), delta cycle, recreated object checked |
| Shared mailbox whose Convert stopped half-way (MailUser or UserMailbox) | Rolled back like a user (temporary licence put back as before); never deleted |

- One scheduler pause and one delta cycle serve all the shared mailboxes of the batch.
- **Licence group**: given back to AD only when no converted user of any Convert batch of the journal is left.
- A Recover batch is linked to its Convert batch; running Recover again resumes it. Once complete, Recover refuses to run again on the same Convert batch.

<!-- icon: shield -->
## 11. After the Recover

| What | When | How |
|---|---|---|
| Disaster mails of the users | any time | In their ComponentShared storage, under the case hold. Search them with eDiscovery on the user; Microsoft documents `IncludeUserAppContent` for the cloud storage of users whose mailbox is on-premises (not tested in the lab) |
| Inactive shared mailboxes | once the retention policy has stamped them (up to 5 days) | `Get-Mailbox -InactiveMailboxOnly` shows the retention policy in `InPlaceHolds`; then remove the entry of the deleted identity (its old objectId) from the case hold policy |
| Case hold entries of the users | never, as long as the disaster mails must be kept | Removing the case hold releases the ComponentShared storage to its normal life |
| Journal and logs | after every run | Keep `data\PraCloudMailbox-journal.db`, `logs\` and `reports\`: they record what was changed |

> [!WARNING]
> `Invoke-HoldRemovalAction` is refused while the case hold policy exists: remove the location from the policy instead (`Set-CaseHoldPolicy -RemoveExchangeLocation <old objectId>`), and only once another hold (the retention policy) is in `InPlaceHolds`, otherwise the inactive mailbox is purged.

<!-- icon: chart -->
## 12. Reading the results

Each run writes:

| File | Content |
|---|---|
| `logs\PRA2_<date>_<run>.log` | every step with its time; last line `RESULT: PASS` or `RESULT: FAIL` |
| `logs\PRA2_<date>_<run>.transcript.txt` | the PowerShell transcript of the run |
| `reports\PRA2_<date>_<run>.csv` | one row per object |
| `reports\PRA2_<date>_<run>.html` | the same, self-contained (no external resource): result, counters, next step, filterable table |

![HTML report of a Convert](images/pra-report-convert.png)

| Column | Content |
|---|---|
| `Identity`, `Kind`, `PrimarySmtpAddress`, `ObjectGuid`, `ExchangeGuid` | the object, from the snapshot |
| `Action` | what the run did or plans |
| `Entra`, `ExchangeOnline`, `TeamsStorage`, `Licence`, `Holds`, `Permissions` | the codes or states per area (Check codes: chapter 8) |
| `FinalStatus` | `Planned`, `AlreadyDone`, `Success`, `Error`, `Pending`, `Skipped` |
| `Detail`, `Warnings` | the explanation, and what to look at |

Exit codes: `0` done, `1` failed (or an object is not ready for Check), `2` done, next step required (an object is *Pending*: run the same command again later).

# Part IV · Maintain

<!-- icon: layers -->
## 13. Architecture

| File | Role |
|---|---|
| `Invoke-PraCloudMailbox.ps1` | entry script: parameters, edition check, the run context, the four actions (`Invoke-PraCollect`, `Invoke-PraCheck`, `Invoke-PraConvert`, `Invoke-PraRecover`) and their helpers (licence need, assessment of the current state, journal) |
| `module\PRA2.Common.psm1` | console theme, banner, steps, summary card, log, configuration (`Import-PraConfiguration`), audit (log + transcript), results and reports (`Complete-PraRun`) |
| `module\PRA2.Store.psm1` | SQLite (System.Data.SQLite in `lib\sqlite`): the snapshot database and the journal, schema version 1 |
| `module\PRA2.Collect.psm1` | Collect (Windows PowerShell 5.1): Exchange connection, scope, mailboxes, permissions, groups expanded, contacts and distribution groups |
| `module\PRA2.Cloud.psm1` | cloud (PowerShell 7): connections, Graph with retries, tenant facts, readiness rules (`Get-Pra2Readiness`, pure functions), source of authority, licences, permissions, case hold, Entra Connect operations |
| `templates\Report.template.html` | the HTML report |

The run state is one hashtable, the **context**, created by the entry script and passed to every function: `Action`, `Mode`, `RunId`, `Config`, `StepIndex` / `StepTotal`, `Rows` (one result row per object), `Issues` (run errors), `WarningList`, `BatchId`, `SnapshotLabel`, `NextSteps`, `Journal` (open journal connection), `LogFile`, `TranscriptPath`, `ExitCode`.

Conventions: `Set-StrictMode -Version Latest` everywhere; generic lists created with `::new()`; never `@(...)[0]` on a result that can be empty (use `Select-Object -First 1`); files in UTF-8 **with BOM** and CRLF (Windows PowerShell 5.1 reads them); one `Write-Host` per console line (the 5.1 transcript writes each `-NoNewline` piece on its own line). The console uses console colours, not ANSI sequences, so that the transcript and the log stay clean; `PRA_ICONS = Emoji | Symbols | Ascii` forces an icon style and `PRA_ANSI = 1` writes the colours as ANSI sequences (screenshots of the guides).

<!-- icon: database -->
## 14. The databases

Two SQLite files, schema version 1, never mixed up: each one records its kind and the tool refuses the other.

**Snapshots** (`Store.Path`, written by Collect): `snapshot` (one row per Collect: status, dates, server, scope, counters), `mailbox`, `permission`, `group_member` (members of the groups used as trustees), `contact`, `distribution_group`, `dl_member`, `collect_issue`. Every data row carries its `snapshot_id`; deleting a snapshot deletes its rows.

**Journal** (`Store.JournalPath`, cloud side, written by Convert and Recover; a new copy of the snapshot file never touches it): `batch` (one row per Convert or Recover: action, snapshot, linked Convert batch, status, account, counters), `batch_item` (one row per object: Entra ID object, original state, last step, status, message), `batch_event` (every step with its time and outcome), `tenant_change` (licence group switched to the cloud, and when it was given back).

```powershell
Import-Module .\module\PRA2.Store.psm1
$db = Open-Pra2Store -Path .\data\PraCloudMailbox.db -Root . -ReadOnly
$snapshot = Get-Pra2Snapshot $db
Read-Pra2Table $db permission $snapshot.id | Format-Table access_right, trustee, trustee_upn, auto_mapping
Close-Pra2Store $db

$journal = Open-Pra2Store -Path .\data\PraCloudMailbox-journal.db -Root . -Kind Journal -ReadOnly
Get-Pra2BatchItem $journal 5b21030d | Format-Table identity, kind, step, status, message
Close-Pra2Store $journal
```

<!-- icon: beaker -->
## 15. Tests

```powershell
powershell.exe -NoProfile -File .\tests\Invoke-TestGate.ps1   # Windows PowerShell 5.1: Collect, store, rules
pwsh -NoProfile -File .\tests\Invoke-TestGate.ps1             # PowerShell 7: everything, Convert and Recover end to end
```

Pester 5 and PSScriptAnalyzer; offline, nothing reaches a server or a tenant. The gate runs every `*.Tests.ps1`, then the analyzer, and writes its evidence (Pester XML, analyzer CSV, summary) in `tests\evidence\gate`.

- `PraCloudMailbox.Tests.ps1`: Collect against synthetic Exchange cmdlets, the store, the configuration, the readiness rules (pure functions), the licence helpers, the Entra Connect modes, the console and the reports, and hygiene rules on the code (StrictMode patterns, encodings).
- `PraCloudMailbox.EndToEnd.Tests.ps1` (PowerShell 7): the real Convert and Recover code against `FakeTenant.ps1`, an in-memory tenant that reproduces what the lab showed — Teams storage promoted, mailbox disabled when the licence goes too early, a hold that blocks the switch back, inactive mailbox when a held identity is deleted, deleted object restored by Entra Connect unless purged, case hold "failed to be deployed" — with faults to inject. The scenarios cover the order of the steps, resume, refusals, waves of batches, partial failures and the rollback of every state. Each scenario was checked by reintroducing the bug it guards against.

The repository tools (not in the package, PowerShell 7.4 and Microsoft Edge):

```powershell
.\tools\New-DocumentationImages.ps1    # screenshots of the guides: the real tool against the in-memory tenant
.\tools\Build-Documentation.ps1        # the HTML guides (user guide, developer guide)
.\tools\New-ReadmeImages.ps1           # the graphics of the README (light and dark)
.\tools\New-PraCloudPackage.ps1        # the release folder: run-time files and the guides only
```

<!-- icon: wrench -->
## 16. Evolving the tool

- **A new Check rule**: a finding in `Get-Pra2Readiness` (code, level, area, message), a row in the table of chapter 8, a unit test. A rule that refuses an object must be an *Error*: Convert refuses exactly the objects with an error.
- **A new step in Convert or Recover**: assess the current state first (a resumed run must not repeat it), write it to the journal (`Write-PraJournal`), add the behaviour to `FakeTenant.ps1` and an end-to-end scenario, then validate it in a lab before a release.
- **Lab first**: every behaviour of Exchange Online and Entra ID the tool relies on was measured in a lab (appendix B). A change of order or a new operation needs the same proof.

# Appendices

<!-- icon: lifebuoy -->
## Appendix A - Troubleshooting

| Symptom | Cause and what to do |
|---|---|
| Convert refused: *Tenant prerequisites not met* | Run Check: missing Graph permission, licence group, free units, temporary licence. Nothing was changed |
| A user stays *Pending* at Convert | Exchange Online is still provisioning: run `Convert -Mode Apply -Batch <ID>` again later. An `EXCHANGE_PLAN_PRESENT` object is refused for this reason |
| Licence assignment *400 Bad Request* right after the usageLocation | Retried for 45 s by the tool; if it persists, check `Cloud.DefaultUsageLocation` and the SKU |
| *other hold(s) on the cloud mailbox block the rollback* (Recover, user) | A hold other than the case hold of the tool (Litigation, retention policy, another eDiscovery hold) blocks the switch: remove it, run Recover again. `ORG_HOLD` in Check warns about organisation-wide policies |
| *Policy ... failed to be deployed* on the case hold | Seen in the lab while the change is recorded: the tool reads the policy back and goes on; the stamp in `InPlaceHolds` is checked anyway. A policy in *Error* can be redeployed with `Set-CaseHoldPolicy -RetryDistribution` |
| *the case hold is not stamped on the shared mailbox: identity NOT deleted* | The hold did not reach the mailbox within `HoldTimeoutMinutes`: run Recover again later. The identity is never deleted without the hold |
| *back on-premises; case hold not stamped yet* (user *Pending*) | The user is back; the hold is still being applied: run Recover again later to check it |
| A deleted shared object came back with its old objectId | A synchronisation ran before the permanent deletion (scheduler not paused): the tool pauses it; in `Manual` mode, pause it before confirming |
| Recreated shared object not found | Entra Connect did not create it yet (or the AD object is missing): run a delta cycle, then Recover again |
| *Collect runs in Windows PowerShell 5.1* / *runs in PowerShell 7.4* | Each action runs in its edition: Collect in `powershell.exe`, the cloud actions in `pwsh.exe` |

<!-- icon: beaker -->
## Appendix B - Lab validation

Lab: four Exchange Server SE servers in a DAG, Active Directory, Entra Connect 2.6.3.0, a Microsoft 365 E5 tenant with Teams. Each behaviour was measured by hand first, then with the tool.

| Test | Date | What it proved |
|---|---|---|
| T1 to T8 | 2 Oct | Source of authority transfer; GUID cleared + licence = cloud mailbox; shared mailbox on a new object; mail flow; a hold blocks the user rollback |
| T13 to T16 | 2 Oct | Licence by group (cloud management of a synchronised group), Kiosk plan inside the Teams licence, gap-free rollback order |
| A, B, C, H | 2-3 Oct | The cloud mailbox of a Teams user is its Teams storage; there is no supported way to delete it; clearing the history empties its mails; a hold blocks the rollback |
| TC | 6-7 Oct | Teams chats kept (client and compliance copies) through the whole cycle |
| S1, S2 | 7 Oct | A shared mailbox deleted with a hold already in place becomes inactive; the adaptive scope never catches a mailbox already deleted without a hold |
| SH | 7 Oct | Shared mailbox on the same synchronised object: temporary licence, then identity deleted under an eDiscovery hold = inactive mailbox; new object recreated by one delta cycle |
| Tool, users | 7-8 Oct | Licence group (cloud): Convert, cloud mailbox = Teams storage; Recover, MailUser with the on-premises GUID 2 s after the plan removal, case hold stamped in 90 s. Licence group synchronised from AD: switched to the cloud, cloud mailbox in 1 min 39; Recover in 8 min 48, group back to AD last |
| Tool, Direct and Kiosk | 8 Oct | Direct mode on a user who already holds the SKU with Exchange disabled: no new unit, Exchange enabled inside the existing assignment (Graph needs the dependent plans disabled too), Recover put the assignment back with exactly the same disabled plans. Kiosk mode with the Teams licence from a synchronised group: direct assignment of the same SKU with Exchange Kiosk (no new unit), cloud mailbox in 1 min 56; Recover removed only the direct assignment |
| Tool, shared mailboxes | 7-8 Oct | Three shared mailboxes in one batch with **one** licence unit (10 min: mailbox 24-46 s, SharedMailbox 27 s each, 10/10 permissions). Recover in 14 min 33: hold stamped in 90 s each, one scheduler pause and one delta cycle for the batch, three new objects |
| Tool, waves and remoting | 8 Oct | Two users converted in two batches through a synchronised licence group: the Recover of the first wave kept the group in the cloud, the second gave it back to AD. `EntraConnect.Mode = Remoting` from a domain server: pause 6 s, resume 3 s, delta waited to the end |

<!-- icon: shield -->
## Appendix C - Security and data

- The app registration has high privileges (users, licences, Exchange Administrator, eDiscovery Manager): keep its certificate on the cloud admin server only, with a non-exportable private key if possible, and give the server the protection of a tier 0 asset.
- The snapshot holds names, addresses, GUIDs and permissions; the journal holds the original state of every converted object; the logs and reports hold names and results. Store them as such, and keep the journal: it is the only record of what was changed.
- The tool never writes a secret: the cloud sign-in uses the certificate; the Entra Connect sample uses a credential protected with DPAPI (`Export-Clixml`), readable only by the same account on the same computer.

<!-- icon: tag -->
## Appendix D - Versions

MAJOR.MINOR.PATCH: MAJOR for an incompatible change of the configuration, of the databases or of the procedure, MINOR for a feature, PATCH for a fix. See `CHANGELOG.md`.
