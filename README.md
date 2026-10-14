<p align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-banner-dark.png">
    <img alt="PRA Cloud Mailbox: scenario 2 disaster recovery from the cloud, with four actions and one journaled batch; user/shared mail is protected, native-resource bookings require export before Recover" src="docs/images/readme-banner-light.png">
  </picture>
</p>

<p align="center">
  <a href="#how-it-works"><b>How it works</b></a> &nbsp;&middot;&nbsp;
  <a href="#convert-and-recover"><b>Convert and Recover</b></a> &nbsp;&middot;&nbsp;
  <a href="#shared-mailbox-permissions"><b>Shared mailboxes</b></a> &nbsp;&middot;&nbsp;
  <a href="#wpf-window"><b>WPF window</b></a> &nbsp;&middot;&nbsp;
  <a href="#reports"><b>Reports</b></a> &nbsp;&middot;&nbsp;
  <a href="#quick-start"><b>Quick start</b></a> &nbsp;&middot;&nbsp;
  <a href="docs/PraCloudMailbox-UserGuide.md"><b>User guide</b></a> &nbsp;&middot;&nbsp;
  <a href="docs/PraCloudMailbox-Guide.md"><b>Developer guide</b></a>
</p>

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.
>
> The `Install-Module` commands in this documentation use `-Force`, so they also update or reinstall a module that is already installed. If an older version still conflicts, close every PowerShell window, open a new one (as administrator for `-Scope AllUsers`), run `Uninstall-Module <ModuleName> -AllVersions -Force`, then run the `Install-Module` command again.

## Why

In a hybrid organisation, a mailbox hosted on-premises is only a *mail user* for Exchange Online. When **Active Directory, Exchange and Entra Connect** are lost at the same time — ransomware, loss of the site — the users can still sign in to Microsoft 365 and Teams still works, but nobody has a mailbox any more, and nothing can be written in AD to fix it. *PRA* stands for *Plan de Reprise d'Activité*, the disaster recovery plan; this is its **scenario 2** (scenario 1, Exchange servers lost with AD alive, is [PRA Remote Mailbox](https://github.com/Nico77600/PraRemoteMailbox)).

Done by hand, giving each user a cloud mailbox means changing the source of authority of every identity, licensing it in the right order, recreating the shared mailboxes with their permissions — which can no longer be read on-premises — and, months later, undoing all of it **without deleting the mails received in the cloud**. This tool collects what it needs before the disaster, then does the conversion and the rollback in the order validated in a lab, verified at each step and journaled.

**Author:** Nicolas Fabert · **Version:** 1.3.0 (unreleased; [changelog](CHANGELOG.md)).

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-principles-dark.png">
  <img alt="Design principles: collected before the disaster, preview first, journaled recovery, user/shared mail protected; unprotected native-resource bookings are not preserved" src="docs/images/readme-principles-light.png">
</picture>

## How it works

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-how-it-works-dark.png">
  <img alt="One configuration and one entry script: Exchange snapshot, cloud operations, journal and reports; four actions: Collect, Check, Convert, Recover" src="docs/images/readme-how-it-works-light.png">
</picture>

- **Collected before.** `Collect` runs every day on an Exchange server (Windows PowerShell 5.1) and stores the mailboxes, their GUIDs and X500 addresses, and the **permissions of the shared mailboxes** (FullAccess with AutoMapping, SendAs, SendOnBehalf, groups expanded) in one SQLite file.
- **Same identity, Teams storage promoted.** `Convert` (PowerShell 7, certificate of an app registration) moves the source of authority of each identity to the cloud and gives it an Exchange plan: the **Teams storage of the user becomes his mailbox**, with his chats. Shared mailboxes are converted in waves, one temporary licence unit each (`Licensing.Shared.Parallel`, 100 by default), given back after each wave.
- **Preview by default, one batch ID.** Without `-Mode Apply` nothing is changed. Apply asks one confirmation, journals the original state of every object, and prints the batch ID with the exact next command.
- **Back without losing a mail.** `Recover -Batch <ID>` gives the identities back to AD in waves (one scheduler pause and one synchronisation per wave, not per object); the users' cloud data stays under an **eDiscovery case hold** (several policies in the same case once 1,000 mailboxes is reached), the cloud shared mailboxes become **inactive mailboxes**, and every licence is put back exactly as it was.

## WPF window

Open the window with `pwsh -File .\Invoke-PraCloudMailbox.ps1 -Gui` on the cloud admin server, or `powershell.exe -File .\Invoke-PraCloudMailbox.ps1 -Gui` on the Exchange server. PowerShell 7.5+ uses the Fluent theme; Windows PowerShell 5.1 uses compatible WPF controls.

The window starts no operation automatically. Each action runs the existing engine in a separate process: **Windows PowerShell 5.1 for Collect, PowerShell 7 for cloud actions**. Select a snapshot or batch, filter the objects and tick a wave. Preview is mandatory before Apply, which requires typing the action name; changed selections invalidate Preview.

The activity panel shows progress, results, logs/reports and the next command. Copy the snapshot and journal between the two computers; the window does not transfer them. See the [user guide, chapter 7](docs/PraCloudMailbox-UserGuide.md#7-the-window).

![WPF window - the four actions, selection, batches and activity](docs/images/pra-gui-overview.png)

## Convert and Recover

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-flows-dark.png">
  <img alt="Users, Convert: GUID cleared, cloud authority, Exchange plan, cloud mailbox. Users, Recover: AD authority, plan removed, on-premises, case hold. Shared mailboxes, Convert: temporary licence, SharedMailbox, tag, permissions. Shared mailboxes, Recover: case hold, identity deleted, inactive mailbox, new object" src="docs/images/readme-flows-light.png">
</picture>

| Object | Convert | Recover |
|---|---|---|
| User | Cloud authority, Exchange plan, Teams storage promoted on the same identity | AD authority and original licences restored; cloud data under a case hold |
| Shared mailbox | Temporary licence per parallel object, SharedMailbox type, collected permissions | Held cloud mailbox becomes inactive; Entra Connect recreates the on-premises identity |
| Native room/equipment (opt-in) | New cloud-only resource, required RoomLists, booking settings; no contents or general ACLs imported | Owned resource deleted without hold; routing contact to the restored on-premises mailbox |

The states of user/shared mailboxes through the cycle:

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-states-dark.png">
  <img alt="User: mailbox on-premises, then cloud mailbox (the Teams storage), then on-premises again with the cloud mailbox under the case hold. Shared: on-premises, then cloud shared mailbox, then on-premises again with the cloud mailbox inactive under the hold, then retention" src="docs/images/readme-states-light.png">
</picture>

> [!CAUTION]
> The on-premises mailbox contents are **not migrated**. User/shared cloud data is protected during Recover. New native resources follow a separate **destructive** lifecycle: no hold, no inactive copy, and cloud-created bookings/messages are **not preserved or merged**. Export/reconcile them before Recover.

### Native rooms and equipment

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-resources-dark.png">
  <img alt="Native resource continuity: collect metadata and room lists, create real cloud resources, export/reconcile cloud bookings, then recover through owned routing contacts without hold or content merge" src="docs/images/readme-resources-light.png">
</picture>

**1.3.0 development addition, not in release 1.2.0.** For rooms/equipment excluded from Entra synchronisation, enable `Scope.ResourceMode = 'CloudResource'` and `Scope.ConvertRooms = $true`, then run a **new Collect**. Convert creates actual RoomMailbox/EquipmentMailbox resources with disabled sign-in and verified booking/delivery controls. Standard resources up to 50 GB require no mailbox licence; unfinished resources remain hidden with booking disabled.

Recover requires **both** verified on-premises routing (`Resources.OnPremisesReady`) and explicit loss acknowledgement (`Resources.AllowUnprotectedDeletion`); both ship as `$false`. Ownership, collisions, unexpected licences or holds block the operation. No purge, proxy removal or automatic second deletion; old protected journals keep their original contract, and RoomLists are retained.

Cloud actions require a **fresh, unconnected PowerShell process** so SDK retries cannot silently repeat an unprotected deletion. Setup, recovery and transport details: [user guide §1.2](docs/PraCloudMailbox-UserGuide.md#12-native-rooms-and-equipment-excluded-from-synchronisation), [developer guide §6.2](docs/PraCloudMailbox-Guide.md#62-unsynchronised-native-resource-lifecycle-development). Lab coverage and remaining client-discovery/delivery limits: [appendix B](docs/PraCloudMailbox-Guide.md#appendix-b---lab-validation). Test the whole cycle in your tenant before production.

## Shared mailbox permissions

<picture>
  <source media="(prefers-color-scheme: dark)" srcset="docs/images/readme-permissions-dark.png">
  <img alt="Shared mailbox permissions: collected from Exchange before the disaster, groups expanded to users, saved in SQLite, granted and verified in Exchange Online; FullAccess with AutoMapping, SendAs and SendOnBehalf" src="docs/images/readme-permissions-light.png">
</picture>

Unlike scenario 1, Active Directory is also unavailable during the disaster: permissions must be **collected beforehand** from Exchange. The snapshot keeps FullAccess with AutoMapping, SendAs and SendOnBehalf, with group trustees expanded to users. Convert grants and reads back the collected permissions in Exchange Online; see the [developer guide, chapter 9](docs/PraCloudMailbox-Guide.md#9-convert). General mailbox and calendar-folder ACLs are **not imported for native resources**.

## Reports

<table>
  <tr>
    <td width="50%" valign="top"><a href="docs/images/pra-console-convert.png"><img alt="Convert -Mode Apply in the console" src="docs/images/pra-console-convert.png"></a><br><sub><b>Convert</b> &middot; readiness, plan, users then shared mailboxes; summary card with the batch ID and the next command</sub></td>
    <td width="50%" valign="top"><a href="docs/images/pra-console-recover.png"><img alt="Recover -Mode Apply in the console" src="docs/images/pra-console-recover.png"></a><br><sub><b>Recover</b> &middot; users back on-premises under the case hold, shared mailboxes inactive and recreated by Entra Connect</sub></td>
  </tr>
  <tr>
    <td width="50%" valign="top"><a href="docs/images/pra-console-check.png"><img alt="Check in the console" src="docs/images/pra-console-check.png"></a><br><sub><b>Check</b> &middot; read-only: tenant prerequisites, then whether each object is ready for Convert</sub></td>
    <td width="50%" valign="top"><a href="docs/images/pra-report-convert.png"><img alt="HTML report of a Convert" src="docs/images/pra-report-convert.png"></a><br><sub><b>HTML report</b> &middot; result, counters, next step, one row per object with its Entra ID, Exchange Online, licence, hold and permission status</sub></td>
  </tr>
</table>

Every run also writes a log, a PowerShell transcript and a CSV; Convert and Recover write the journal of the batch. Exit codes: 0 = done, 1 = failed (or an object is not ready), 2 = done, next step required.

## Requirements

| Item | Requirement |
|---|---|
| Scenario | Hybrid Exchange organisation synchronised by **Entra Connect**. After the disaster: AD restored from a backup (same objectGUID), Exchange, and Entra Connect **2.5.76.0** or later |
| Collect | An Exchange server (or the management tools), **Windows PowerShell 5.1**, an account that reads recipients and permissions. Nothing to install (SQLite is bundled) |
| Cloud actions | A cloud admin server that does not depend on the on-premises AD, **PowerShell 7.4+**, `Microsoft.Graph.Authentication` and `ExchangeOnlineManagement` 3.10+ |
| App registration | Certificate; Graph application permissions (users, source of authority, licences, organisation, group members); `Exchange.ManageAsApp` with Exchange Administrator; eDiscovery Manager in Microsoft Purview |
| Licences and holds | A licence group with an Exchange Online plan (or the Kiosk plan of the Teams licence, or direct licences), free units for the shared mailboxes (one per mailbox of a wave), an eDiscovery case hold policy, a retention policy for the inactive shared mailboxes |
| Console | Windows Terminal for emoji and colours; the classic console shows symbols. The window: Windows 10/11 or Windows Server 2016+, Fluent theme with PowerShell 7.5+ |

## Quick start

```powershell
git clone https://github.com/Nico77600/PraCloudMailbox.git
cd PraCloudMailbox
notepad .\config\PraCloudMailbox.config.psd1        # scope, tenant, app registration, licences, case hold, Entra Connect

# Before the disaster - Exchange server, Windows PowerShell 5.1 (scheduled task, every day)
.\Invoke-PraCloudMailbox.ps1 -Action Collect                        # preview: what would be collected
.\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force     # writes a snapshot to data\PraCloudMailbox.db

# Cloud admin server, PowerShell 7 - copy data\PraCloudMailbox.db there first
.\Invoke-PraCloudMailbox.ps1 -Action Check                          # read-only: is every object ready?
.\Invoke-PraCloudMailbox.ps1 -Action Convert                        # disaster: preview of the plan
.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply            # prints the batch ID
.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 5b21030d -Mode Apply    # infrastructure rebuilt: roll the batch back

# Or the window, on either computer
.\Invoke-PraCloudMailbox.ps1 -Gui
```

One command per moment — every day, every week, the disaster, the rollback, the days after: see the [user guide](docs/PraCloudMailbox-UserGuide.md). The environment values of the configuration are empty in this repository. The zip of each [release](https://github.com/Nico77600/PraCloudMailbox/releases) contains only the files needed to run, with both guides in HTML; `.\tools\New-PraCloudPackage.ps1` builds the same package from the repository.

## Documentation

| Guide | Content |
|---|---|
| **[User guide](docs/PraCloudMailbox-UserGuide.md)** | For the people who run the tool: **prerequisites** and **one command per moment** — collect every day, check every week, convert at the disaster, recover once the infrastructure is rebuilt, the days after — and how to read the results. |
| **[Developer guide](docs/PraCloudMailbox-Guide.md)** | Everything else: the scenario, the procedure validated in the lab and the states of a mailbox, the app registration, the case hold and the retention policy, the configuration in detail, each action step by step, the Check codes, the reports, the architecture, the databases, the tests, how to evolve the tool, troubleshooting and the lab results. |

Both guides also exist as a single HTML file with a light and a dark theme (`docs/PraCloudMailbox-UserGuide.html`, `docs/PraCloudMailbox-Guide.html`): download them and open them locally, or use the copies in the release zip. `tools\Build-Documentation.ps1` and `tools\New-ReadmeImages.ps1` rebuild the guides and the themed README graphics.

## Security

Obtain the tool from this repository and verify the owner's Authenticode signatures and timestamps on the seven signed PowerShell files. Editable configuration, DLLs and other assets are not covered by those signatures. Release attestations, when present, prove ZIP provenance, not the safety of an operation. Report sensitive findings **privately** as described in [SECURITY.md](SECURITY.md); contribution and publication rules are in [AGENTS.md](AGENTS.md).

## Tests

```powershell
# Pester 5+ and PSScriptAnalyzer; no Exchange, no tenant needed
powershell.exe -NoProfile -File .\tests\Invoke-TestGate.ps1     # Windows PowerShell 5.1: Collect, store, rules
pwsh -NoProfile -File .\tests\Invoke-TestGate.ps1               # PowerShell 7: everything, Convert and Recover end to end
```

The end-to-end tests run the real Convert and Recover code against an in-memory tenant that reproduces what the lab showed (Teams storage promoted, a hold that blocks the switch back, inactive mailbox when a held identity is deleted, deleted object restored by Entra Connect unless purged...), with faults to inject: order of the steps, resume, refusals, waves, stops and resumes, partial failures and the rollback of every state; the window is tested in both editions, with a real run in a child PowerShell. The tool was also validated on a lab of Exchange Server SE servers with Entra Connect and a Microsoft 365 E5 tenant — licence groups (cloud and synchronised), Kiosk and direct licences, shared mailboxes, waves, Entra Connect remoting (developer guide, appendix B).

## License

[MIT](LICENSE). Third-party components: [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md) (System.Data.SQLite, public domain).

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

The tool changes production objects in Entra ID and Exchange Online: run the preview, review the plan, keep the journal and test the whole cycle in a lab before using it in a real disaster.
