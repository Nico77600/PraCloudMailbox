# Changelog

## 1.0.0 - 2026-10-08

First public release. Every action and every licensing mode of 0.2.1 was validated in the lab with the tool; this
release adds the publication material and two small fixes found while preparing it.

### Added
- **Documentation in two guides**, each in Markdown and as a single HTML file with a light and a dark theme: the
  **user guide** (prerequisites, one command per moment: every day, every week, the disaster, the rollback, the
  days after) and the **developer guide** (scenario, validated procedure and states of a mailbox, app registration,
  case hold and retention policy with the commands, configuration, every action, Check codes, reports,
  architecture, databases, tests, troubleshooting, lab results). New README with graphics.
- `config\EntraConnect.sample.ps1`: the contract of `EntraConnect.Mode = 'Script'` (PowerShell remoting with a
  credential saved with `Export-Clixml`), and `config\Targets.sample.csv` for `Scope.Mode = 'Csv'`.
- Repository tools (PowerShell 7.4, not in the package): `tools\Build-Documentation.ps1` (HTML guides),
  `tools\New-DocumentationImages.ps1` (screenshots rendered by the real tool against the in-memory tenant),
  `tools\New-ReadmeImages.ps1` (README graphics) and `tools\New-PraCloudPackage.ps1` (release folder and zip, with
  the environment values of the configuration emptied and checked).
- `LICENSE` (MIT) and `THIRD-PARTY-NOTICES.md` (System.Data.SQLite 1.0.119.0, public domain, checked by SHA-256
  against nuget.org).

### Changed
- The configuration of the repository is generic: every environment value is empty and `EntraConnect.Mode` is
  `Manual` until the administrator chooses `Remoting` or `Script`.

### Fixed
- Check: an object without usageLocation is an **error** (`NO_USAGE_LOCATION`) when `Cloud.DefaultUsageLocation` is
  empty — Convert refused it anyway, since a licence needs a country. With a default country it stays an
  information, which now names the country.
- Entra Connect `Manual` mode: the question no longer ends with "on ." when `EntraConnect.Server` is empty.
## 0.2.1 - 2026-10-07

Hardening after an independent code review (2 High, 3 Medium, 3 Low findings, all fixed) and end-to-end tests.

### Fixed
- **Recover run again could lift the case hold of users already back on-premises** (High): when a hold was slow to
  be stamped, the next run started the user again and removed him from the case hold. Recover now classifies every
  object by its current state (cloud mailbox, already back, shared, shared already deleted, half-converted shared);
  a hold is lifted only from a cloud mailbox, never from an object that is back. A complete Recover batch is not
  run again.
- **Licence group given back to AD while users of another wave still needed it** (High): it now goes back only
  when no converted user of any Convert batch of the journal is left, whichever batch switched it.
- **`EXCHANGE_PLAN_PRESENT`** was a warning in Check but a refusal in Convert: it is now an error in Check, with the
  real reason (test T2c), and the mailbox plans are recognised by plan ID.
- **Shared mailbox deleted by an earlier run** could never finish ("not a SharedMailbox: nothing deleted"): the
  next run checks the recycle bin and the recreated object. A shared mailbox whose Convert stopped half-way is rolled
  back like a user instead of being left as it was.
- **Kiosk and Direct licences**: the original direct licences are recorded by Convert; the plan is enabled inside the
  existing assignment and Recover puts the assignment back exactly as it was (it removed the whole SKU, Teams
  included). Batches written by 0.2.0 have no recorded licences: the licence added by Convert is removed as before.
- **Shared mailboxes**: no shared mailbox is started without a free unit of the temporary licence (one that failed
  keeps it; the next ones were left half-converted).
- `EXCHANGE_S_FOUNDATION` no longer counts as a mailbox plan.
- Entra Connect `Script` mode: a script that ends without `exit` no longer fails under StrictMode, and an exit code
  of an earlier call is no longer reused.
- Room and equipment mailboxes are refused by Check and Convert (`KIND_NOT_SUPPORTED`): their booking settings are
  not collected and they were not validated in the lab.
- Case hold policy: the tenant can answer "Policy ... failed to be deployed" to `Set-CaseHoldPolicy` (and to
  `New-CaseHoldPolicy`) while recording the change (lab 7-8 Oct). The tool reads the policy back and goes on when the
  change is there (the stamp in `InPlaceHolds` is checked anyway). The hold of a user is set by object ID, like the
  one of a shared mailbox.
- Licence capacity (Check and Convert) counts only the objects that will take a **new** unit: a user who already
  holds the SKU of his Exchange plan (hybrid user with Exchange disabled), Kiosk mode and shared mailboxes already
  converted take none (lab 8 Oct: Convert was refused with the SKU at 25/25 although nothing new was needed).

### Added
- Check: `ORG_HOLD` warning when retention policies apply to every mailbox (they would block the users' Recover).
- `tests\FakeTenant.ps1` and `PraCloudMailbox.EndToEnd.Tests.ps1`: 21 scenarios running the real Convert and Recover
  code against an in-memory tenant that reproduces the lab behaviours, with fault injection. Every fix above has a
  scenario that fails when the fix is removed (checked by mutation).
- Test gate: PowerShell 7 99 passed (+2 Windows PowerShell only), Windows PowerShell 5.1 81 passed (the 21
  end-to-end scenarios run in PowerShell 7 only); PSScriptAnalyzer clean. A second review of the fixes confirmed the
  8 findings fixed and found one more (a shared mailbox already converted was refused on resume when no unit was
  free, although it needs none), fixed with its scenario.

## 0.2.0 - 2026-10-07

Convert and Recover: the procedure validated in the lab, run by the tool.

- **Convert**: users (ExchangeGuid cleared, source of authority to the cloud, usageLocation set from the cloud,
  Exchange plan by licence group, Kiosk plan or direct licence; the Teams storage becomes the mailbox) in waves,
  then shared mailboxes one at a time (temporary licence, SharedMailbox, retention tag, permissions of the
  snapshot granted and verified, licence removed). A licence group synchronised from AD is switched to the cloud
  once and recorded. Objects with a Check error or an Exchange plan already enabled are refused.
- **Recover**: users (case hold lifted if present, source of authority back to AD, delta cycle, plan removed,
  MailUser with the on-premises GUID, case hold added and checked), shared mailboxes (case hold checked before
  anything is deleted, scheduler paused, identity deleted, inactive mailbox, identity purged, scheduler resumed,
  delta cycle, new object checked), licence group back to AD after the last user.
- **Journal** (`Store.JournalPath`, separate SQLite file): batches, original state of every object, every step,
  tenant changes. `-Batch` resumes a Convert or a Recover; every step checks the current state first.
- **Entra Connect** operations by `Remoting`, `Script` (operator script) or
  `Manual`. New settings: `Cloud.DefaultUsageLocation`, `Polling`, `EntraConnect.Mode/ScriptPath`.
- Offline test gate: Windows PowerShell 5.1 73 tests, PowerShell 7 71 tests (+2 Desktop only); PSScriptAnalyzer clean.
- Lab, 7 October 2026: full cycle of a user (licence group) and of a shared mailbox (temporary E5) with the tool;
  Collect 0.2.0 on an Exchange SE server with `powershell.exe -File`, then Check on it.
- Lab, 8 October 2026 (0.2.1): every mode with the tool. Remoting mode; three shared mailboxes in one batch with one
  E5 unit (Convert and Recover); licence group synchronised from AD (switched to the cloud and back); Direct mode on a
  user who already holds the SKU with Exchange disabled (licence restored exactly); Kiosk mode with the Teams licence
  from a group (only the direct assignment removed at Recover); two waves sharing the licence group (kept in the
  cloud until the last one is back). Collect 0.2.1 on the Exchange SE server.

### Fixed during the lab run
- `@()` on a generic list created by `New-Object` throws "Argument types do not match" (both editions): lists are
  created with `::new()`. `@(...)[0]` on an empty result throws under StrictMode: `Select-Object -First 1`. Tests
  block both patterns.
- `Connect-IPPSSession` loaded its own `Get-Recipient` and `Get-User`, which hid the Exchange Online ones and
  returned a stale recipient type (Recover waited forever for the MailUser): the session now loads the case hold
  cmdlets only.
- A licence assigned just after the usageLocation change can be refused (400): retried for 45 s. Graph errors now
  carry the Graph code and message.
- Resume: the licence capacity check counts only the objects that do not hold their licence yet; the snapshot count
  shows the objects of the batch; no delta cycle when the users are already synchronised; a Done object no longer
  keeps the message of an earlier failure. The Entra Connect script output is logged on one line.
- Check suggests Convert as the next step. `COMPONENT_SHARED` explains that a shared object recreated by Recover
  has a new, empty storage, never linked to the inactive mailbox (lab: three recreated objects).

## 0.1.0 - 2026-10-07

First version of the scenario 2 tool: preparation of the recovery.

- **Collect** (Windows PowerShell 5.1, Exchange on-premises): mailboxes in scope (Auto, OU, Group, Csv; users,
  shared, rooms, equipment), identities, GUIDs, addresses (X500 kept), custom attributes, optional statistics;
  shared mailbox permissions (FullAccess + AutoMapping from msExchDelegateListLink, SendAs, SendOnBehalf),
  trustees resolved to UPN/SMTP, groups expanded recursively, system and admin trustees excluded; optional mail
  contacts and distribution groups (members, owners, delivery settings, dynamic groups). Writes one snapshot per
  run to SQLite (Running then Complete), prunes old snapshots, writes a consistent backup copy (VACUUM INTO).
  A snapshot of one object (-Identity) asks for a confirmation (or -Force).
- **Check** (PowerShell 7.4+, certificate sign-in): snapshot age, Graph application permissions, licence
  capacity (licence group / Kiosk / direct, temporary licence of shared mailboxes), licence group managed from AD
  or the cloud, and per object: Entra ID object (by onPremisesImmutableId, then UPN), source of authority,
  provisioning errors, Exchange Online recipient and ExchangeGuid, ComponentShared storage, Exchange plan
  already enabled, usageLocation, Kiosk plan, holds, retention tag, shared mailbox trustees known in Exchange Online.
- Same console, log, transcript and HTML report as PRA Remote Mailbox 2.0.0; common module runs in both editions.
- Offline test gate: Windows PowerShell 5.1 56 tests, PowerShell 7 54 tests (+2 Desktop only); PSScriptAnalyzer clean.
- Lab: Collect on an Exchange SE server and Check against the lab tenant, 7 October 2026.

### Fixed during the lab run
- Windows PowerShell 5.1 leaves `$PSScriptRoot` empty in `param()` default values when a script runs with
  `-File` (scheduled task): `-ConfigPath` is now resolved in the script body. A test blocks the pattern.
