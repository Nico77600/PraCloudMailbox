---
title: PRA Cloud Mailbox
subtitle: User guide
version: 1.3.0 (unreleased)
author: Nicolas Fabert
updated: 2026-10-12
---

# PRA Cloud Mailbox — User guide

> What must be in place before the disaster, then one command per moment: **every day** (collect), **every week** (check), **the disaster** (convert), **the infrastructure is back** (recover) and **the days after**. Every command can also be run from a window (chapter 7). The procedure in detail, the prerequisites of the app registration, the configuration, the databases and the internals are in the [developer guide](PraCloudMailbox-Guide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.

```cards
checklist | Prerequisites | Chapter 1: two computers, two editions of PowerShell, an app registration, the licences and the holds, then the one-time setup.
database | Before the disaster | Chapter 2: collect every day, check every week.
cloud | Disaster | Chapter 3: preview, convert, note the batch ID.
refresh | Infrastructure rebuilt | Chapters 4 and 5: roll the batch back, then what is left to do.
compare | The window | Chapter 7: the same commands from a window, with a preview before every change.
```

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Exchange side (Collect) | An Exchange server (or the management tools), **Windows PowerShell 5.1**, an account that reads recipients and permissions (*View-Only Organization Management* and *Active Directory Permissions*, or *Organization Management*) |
| Cloud admin server | A Windows computer that does **not depend on the on-premises AD**, **PowerShell 7.4** or later, the modules `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` 3.10+ |
| App registration | Certificate sign-in, Microsoft Graph application permissions, `Exchange.ManageAsApp` with the Exchange Administrator role, eDiscovery Manager in Microsoft Purview: [developer guide, chapter 4.3](PraCloudMailbox-Guide.md#43-app-registration) |
| Licences | A licence group with an Exchange Online plan and free units for the users (or the Kiosk plan of their Teams licence, or direct licences), and free units for the shared mailboxes: one per shared mailbox converted at the same time (`Licensing.Shared.Parallel`; one unit is enough, they then go one after another) |
| Holds | An eDiscovery case hold policy and a retention policy for the inactive shared mailboxes: [developer guide, chapter 4.4](PraCloudMailbox-Guide.md#44-case-hold-and-retention) |
| After the disaster | Active Directory restored from a backup (same objectGUID), Exchange, and **Entra Connect 2.5.76.0** or later |

### 1.1 One-time setup

```steps
Copy the tool | Download `PraCloudMailbox-<version>.zip` from the latest release, extract it to `C:\PRA\PraCloudMailbox` on the Exchange server and on the cloud admin server, and unblock the files.
Install the modules | On the cloud admin server, in an administrator PowerShell 7: `Install-Module Microsoft.Graph.Authentication -Scope AllUsers -Force` and `Install-Module ExchangeOnlineManagement -MinimumVersion 3.10.0 -Scope AllUsers -Force`.
Import the certificate | The certificate of the app registration, with its private key, in `Cert:\LocalMachine\My` of the cloud admin server.
Configure | `notepad .\config\PraCloudMailbox.config.psd1`: the scope (an OU, a group or a CSV list), the tenant, the app registration, the licence group, the temporary licence of the shared mailboxes, the case hold policy and the way to reach Entra Connect. Copy the same file to the other computer.
First collect | On the Exchange server: `.\Invoke-PraCloudMailbox.ps1 -Action Collect` (preview), then `-Mode Apply -Force`. Copy `data\PraCloudMailbox.db` to the cloud admin server.
First check | On the cloud admin server: `.\Invoke-PraCloudMailbox.ps1 -Action Check`. Fix every red line until all the objects are ready.
```

### 1.2 Native rooms and equipment excluded from synchronisation

**Development feature, not in release 1.2.0.** Enable `Scope.IncludeRoom` / `IncludeEquipment` as needed, `Scope.ConvertRooms = $true` and `Scope.ResourceMode = 'CloudResource'`. The default remains off; the legacy `SynchronizedShared` mode is only for synchronised resources.

Before the disaster, run a **new Collect** in this mode. It reads capacity, custom resource properties, Office, postal address/country, phone, notes and room-list memberships, plus supported booking controls and SMTP references for booking delegates/policies, working hours and time zone. An unreadable property/membership or unresolved booking trustee stops Collect instead of silently omitting settings or substituting permissive defaults. Earlier resource snapshots without this metadata are insufficient.

**Convert creates the lists first, then the rooms, then their memberships.** No manual creation of missing lists is required: Collect captures their identities, SMTP/X500 addresses, visibility, delivery and membership controls as well as which rooms belong to them. Run a new Collect if your earlier room-list snapshot lacks these properties. Lists shared by selected rooms are created only once as true cloud-only **RoomLists**, with creation and immutable identities recorded for resume. Only lists needed by the selected native rooms are created; unselected rooms and empty/unreferenced lists are not recreated automatically.

Compatible existing cloud-only lists can be reused without overwriting their settings or other members. Conflicting addresses/aliases, synchronised/ordinary groups, replaced identities and failed list creation/configuration block the dependent rooms. The tool never converts unrelated groups. Equipment is not added to room lists. Check and Convert preview remain read-only and show missing lists as planned creations when the required Exchange commands/RBAC and collected controls are available.

Exchange-only RoomLists are supported without an Entra group or additional licence: their Exchange GUID is pinned immediately, and any directory ID exposed later is also pinned and must not change or disappear. Native mailbox IDs are recorded before sign-in changes; verification waits within the configured mailbox timeout for disabled sign-in to propagate. Resume the same partial batch after read/propagation failures. An attempted creation with no verifiable exact object is blocked rather than repeated.

Unfinished native rooms/equipment are hidden and booking is explicitly disabled immediately after their creation IDs are recorded. The captured booking policy is enabled only after required descriptive properties, work hours, delivery controls, final safety checks and memberships are verified. Failure leaves a Partial batch and attempts verified restrictive cleanup; a reported cleanup failure requires immediate review of the pinned resource. Resume the same batch after fixing the blocker. Already completed resources and unrelated objects are not relocked. `Set-User` failures or inconsistent Graph identity results remain blockers: do not skip descriptive fields, substitute Graph metadata or bypass the duplicate guard.

Allowed/rejected invitation senders or groups, authenticated-sender requirements and message moderation are also preserved. Unresolved restrictions or moderators block creation rather than opening reservation to additional users. These are invitation controls, not FullAccess or send permissions.

New native journals use **`protectionMode = 'None'`**. Standard rooms/equipment [up to 50 GB need no mailbox licence](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits#mailbox-storage-limits): the tool assigns no licence, joins no licensing group and does not set usageLocation for licensing. Native-only Check/Preview/Convert/Recover need no enterprise SKU, free licence capacity, resource licensing group or Security & Compliance connection. Bigger mailboxes, archive/hold requirements and other licensed features are outside this new lifecycle. Leave `Resources.LicensingGroupId` empty for new resources; it remains only for compatibility with old protected journals.

Convert refuses existing cloud identities, mailboxes, contacts and conflicting addresses rather than overwriting them. It creates a new **RoomMailbox** or **EquipmentMailbox**, with a separate cloud GUID, disabled sign-in, addresses/X500, descriptive/location properties, room-list membership, directory visibility and supported booking restrictions/delegates. **FullAccess, SendAs, SendOnBehalf and calendar-folder ACLs are not imported for native resources**, including general rights in earlier snapshots; legacy synchronised-resource permissions are unchanged. Directory visibility is not a replacement for per-user address-book segmentation or calendar viewing ACLs. **It does not import the native calendar or existing reservations**: this is service continuity, not content recovery.

Before Recover, restore the native resources and verify that each `alias@Resources.OnPremisesRoutingDomain` address and the hybrid connector deliver to the correct native mailbox. Configure that domain and explicitly set `Resources.OnPremisesReady = $true` only after this operational check. This flag is an operator assertion, not an automatic test of the on-premises infrastructure.

**New native Recover is destructive: NO hold, NO inactive copy; cloud-created bookings/messages are NOT retained or merged.** Export/reconcile before reopening native booking. `Resources.AllowUnprotectedDeletion` defaults to `$false`; explicitly set it to `$true` only after acknowledging that loss. Preview and the single Apply confirmation repeat this warning; `-Force` does not bypass the configuration guard. Unexpected licences or active/organisation holds block operations; no protection is automatically removed.

Recover follows the recorded lifecycle even if the current resource mode changes. It deletes only the verified journal-owned cloud account and verifies the exact deleted identity and non-inactive soft-deleted mailbox. Only that freshly proven tombstone (both recorded cloud IDs, exact resource type, valid `WhenSoftDeleted` and journaled Verified deletion) is excluded from the routing collision check: its retained proxies alone do not automatically block attempting a routing **MailContact** with the original addresses. Foreign active/soft-deleted recipients still block. Exchange contact creation/configuration uniqueness errors remain explicit Errors in a Partial batch, never Done; a created contact is pinned before configuration for safe resume. Deletion/soft-delete or foreign-address delays stay Pending: resume the same Recover batch once the finding resolves, never purge to accelerate it. The tombstone and its proxies remain unchanged, with no automatic second deletion. Recover does not transfer authority, pause Entra Connect or change the native resource.

An ambiguous/interrupted deletion response or a restored identity never authorises an automatic second deletion. Recover stops if the same immutable identity is still or again active after the request started; wait for deletion evidence or investigate the restore rather than recreating it or deleting journal state. Packaged configuration resets both native-resource consents to `$false`.

**Old native journals without a protection mode stay protected (`CaseHold`), never downgraded.** They still require their recorded enterprise licence group, specific case hold and exact inactive-copy verification. Cloud bookings/messages in those old protected batches remain in that held inactive copy, not automatically merged. Current configuration cannot convert an old journal to the new unprotected contract.

Room lists are retained on Recover; groups are not automatically deleted and routing contacts are not added to RoomLists. Review room discovery and any later list cleanup separately. The Convert journal identifies the lists that were created versus reused.

Use the usual Check, Convert preview/Apply and Recover preview/Apply commands, or the same buttons in the window. The native path was exercised with synthetic resources, booking responses and recovery contacts; see developer guide appendix B for the coverage and limits. A full live cycle in your own tenant is still required before production use, including client discovery and delivery through the restored mail route.

<!-- icon: database -->
## 2. Before the disaster

### 2.1 Every day: collect

A scheduled task on the Exchange server, under the collect account:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File C:\PRA\PraCloudMailbox\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force
```

Then copy `data\PraCloudMailbox.db` (or the last file of `data\backup`) to the cloud admin server, for example with a second task. **After the disaster nothing on-premises can be read**: the last copy is what the tool knows about the mailboxes and the permissions of the shared mailboxes.

### 2.2 Every week: check

```powershell
.\Invoke-PraCloudMailbox.ps1 -Action Check                          # every object of the last snapshot
.\Invoke-PraCloudMailbox.ps1 -Action Check -Identity compta@contoso.com
```

Read-only. Every object must be green: an object in error would be refused by Convert. Typical fixes: a free licence unit, a missing Graph permission, an Exchange Online plan already enabled on a user (`EXCHANGE_PLAN_PRESENT`), a user without country (`NO_USAGE_LOCATION`), an organisation-wide retention policy (`ORG_HOLD`). Check also warns when the last snapshot is too old.

![Check: the tenant prerequisites, then one line per object](images/pra-console-check.png)

<!-- icon: cloud -->
## 3. Disaster: convert

AD, Exchange and Entra Connect are lost. On the cloud admin server, in PowerShell 7:

```powershell
.\Invoke-PraCloudMailbox.ps1 -Action Convert                         # preview: the plan, nothing is changed
.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply             # one confirmation, then the batch
```

- The users get their mailbox first (their Teams storage becomes their mailbox, in one or two minutes), then the shared mailboxes, in waves (up to `Licensing.Shared.Parallel` at the same time, 100 by default, one temporary licence each), with their permissions.
- **Note the batch ID** printed in the final card: it is the only input of the rollback. It is also in the report and in the journal.
- A user still *Pending* at the end: run the same command with `-Batch <ID>` a little later; only what is not finished is taken again.
- One part only: `-Scope UsersOnly` or `-Scope SharedOnly`; one object: `-Identity compta@contoso.com`; a wave: `-IdentityPath .\wave1.txt` (one identity per line, or a CSV file with an `Identity` column). Each Convert is a batch of its own.

![Convert -Mode Apply: the batch ID and the next command in the final card](images/pra-console-convert.png)

<!-- icon: refresh -->
## 4. Infrastructure rebuilt: recover

AD (restored from a backup), Exchange and Entra Connect work again. On the cloud admin server:

```powershell
.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 5b21030d                   # preview
.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 5b21030d -Mode Apply       # roll the batch back
```

- **Users**: they are synchronised from AD again and their mail goes back to their on-premises mailbox. Their cloud mailbox — with the mails received during the disaster — stays in Microsoft 365 under the eDiscovery case hold.
- **Shared mailboxes**: their cloud mailbox becomes an **inactive mailbox** under the case hold (nothing is lost), and Entra Connect creates a new object linked to the on-premises shared mailbox.
- Something not finished (*Pending*, or an error fixed in the meantime): run the same command again.
- Several Convert batches: roll each one back with its own ID. A batch can also go back in waves: `-Batch 5b21030d -IdentityPath .\wave1.txt`, then the rest with the same `-Batch`.

![Recover -Mode Apply: users back on-premises, shared mailboxes inactive and recreated](images/pra-console-recover.png)

<!-- icon: shield -->
## 5. The days after

| What | How |
|---|---|
| Find a mail received during the disaster | eDiscovery search on the user (his cloud storage) or on the inactive shared mailbox |
| Inactive shared mailboxes | After a few days the retention policy holds them too (`Get-Mailbox -InactiveMailboxOnly`, `InPlaceHolds`); their entry in the case hold policy can then be removed — never before |
| Keep the records | `data\PraCloudMailbox-journal.db`, `logs\` and `reports\`: what was changed, when, by whom |
| Collect again | Once Exchange is back, the daily collect starts again for the next disaster |

<!-- icon: file -->
## 6. Results

Every run prints a banner, numbered steps and a summary card with the result, the batch ID, the report, the log, the duration and the next command.

| File | Content |
|---|---|
| `reports\PRA2_<date>_<run>.html` | the report: result, counters, next step, one row per object (filter box) |
| `reports\PRA2_<date>_<run>.csv` | the same rows, for Excel |
| `logs\PRA2_<date>_<run>.log` | every step with its time; last line `RESULT: PASS` or `RESULT: FAIL` |
| `logs\PRA2_<date>_<run>.transcript.txt` | the PowerShell transcript |

![HTML report of a Check](images/pra-report-check.png)

| Exit code | Meaning |
|---|---|
| `0` | done |
| `1` | failed, or (Check) an object is not ready |
| `2` | done, next step required: an object is *Pending*, run the same command again later |

Something unexpected: [developer guide, appendix A - Troubleshooting](PraCloudMailbox-Guide.md#appendix-a---troubleshooting).

<!-- icon: compare -->
## 7. The window

```powershell
pwsh -File .\Invoke-PraCloudMailbox.ps1 -Gui              # cloud admin server, PowerShell 7
powershell.exe -File .\Invoke-PraCloudMailbox.ps1 -Gui    # Exchange server, Windows PowerShell 5.1
```

The window changes nothing by itself: each button runs the command of this guide in its own PowerShell — Windows PowerShell 5.1 for Collect, PowerShell 7 for the others — with the same log, transcript, journal and report. The **Activity** panel follows the run step by step, then shows its result, the batch ID, the next step, the report and the log.

![The window: the next step, the configuration, the last snapshot and Check, the batches](images/pra-gui-overview.png)

| Page | What it shows | What it runs |
|---|---|---|
| Overview | The next step, the configuration (certificate found, valid until), the last snapshot and Check, the batches, this computer (PowerShell 7, modules) | The button of the next step |
| 1 Collect | The scheduled task to copy, the snapshots of the database | Preview, *Collect now* (after a yes) |
| 2 Check | The counters and one row per object of the last Check; filter by text, status and type | Check: every object, the users or the shared mailboxes |
| 3 Convert | Every object of the snapshot with its last Check and its last batch | Preview, then *Convert…* |
| 4 Recover | The Convert batches, then the objects of the batch chosen | Preview, then *Recover…* |

```steps
Preview | *Convert…* and *Recover…* stay disabled until a preview has run on **exactly** the same selection: same snapshot or batch, same objects ticked. Any change asks for a new preview.
Confirm | Type `CONVERT` or `RECOVER`: the dialog shows the objects and the command line that will run.
Waves | Tick the objects of a wave — or filter, then *Tick the ready ones shown*. The window writes the list to `data\gui\wave-<action>-<time>.csv` and runs the command with `-IdentityPath`: the log names the file.
Stop | *Stop after the current object*: no new object is started and the run ends normally (journal, report). *Preview the resume* (Convert) or Recover again takes up what is left.
```

- With Entra Connect in *Manual* mode, the run asks its questions in a Yes / No dialog (run the synchronisation, then answer).
- The window cannot be closed while a run is going.
- The files of each run are in `data\gui` (`run-*.events.jsonl`, the output of the PowerShell); they are deleted after 30 days, the wave lists are kept.

![Convert: every object with its last Check and its batch; the activity of the Convert and its batch ID](images/pra-gui-convert.png)
