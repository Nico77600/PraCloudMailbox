# Changelog

## 1.2.0 - 2026-10-09

Volume: 15,000 users, 5,000 shared mailboxes.

### Changed
- The cloud state of every object is read in bulk before anything is changed (Check, Convert, Recover): Entra ID users
  page by page, source of authority by Graph batches, Exchange Online by filters of 50 objects. About 20 minutes for
  21,000 objects instead of 15 hours one by one; polls of many objects (mailboxes, synchronisation) also by sets.
- Convert goes phase by phase: users in waves of 500 (Graph batches, one wait per wave until Exchange Online no
  longer sees them as synchronised), shared mailboxes in waves of `Licensing.Shared.Parallel` (100 by default), one
  temporary unit each and never more than the free units; the permissions are granted once the temporary licence
  is given back. A stop from the window is taken between two waves.
- Recover goes phase by phase too, one wave at a time: users (and shared mailboxes whose Convert stopped half-way)
  in waves of 500, shared mailboxes in waves of `Licensing.Shared.Parallel`. Every phase of a wave runs for the
  whole wave at once (case hold, source of authority, Exchange plan removal, MailUser wait) with a single wait
  instead of one per object; a shared wave pauses the Entra Connect scheduler once for all its identity deletions
  and runs a single delta cycle, instead of once per mailbox. A stop from the window is honoured between waves (a
  wave already running is always finished).
- The case hold used by Recover can now hold more than 1,000 mailboxes: when the configured policy
  (`Retention.HoldPolicy`) is full, Recover creates `<HoldPolicy>-02`, `-03`... in the same eDiscovery case (with a
  rule that holds everything) and keeps using them transparently. A sibling policy being deleted is left out.
- Convert and Recover report their progress wave by wave: a console line and a 'progress' event (window: a second
  progress bar and a line under the step one, N of total objects of the current phase, with a rough time left once
  at least one wave is done) instead of only the final per-object result lines.
- The *Plan and confirmation* step of Convert and Recover shows a rough duration estimate from the number of waves
  (about 2 minutes per user wave, about 3 minutes per shared-mailbox pipeline round, lab order of magnitude - a real
  tenant can be much slower some evenings).
- Room and equipment mailboxes can be converted like shared mailboxes (`Scope.ConvertRooms`, default off): they
  still need `Scope.IncludeRoom` / `IncludeEquipment` at Collect. Their booking settings (capacity, auto-accept,
  booking policies) are not collected or recreated; the on-premises object is authoritative and comes back
  unchanged at Recover.

### Added
- `Licensing.Shared.Parallel` (1 to 1000, default 100).
- `Retention.HoldPolicyLimit` (1 to 1000, default 1000): mailboxes per case hold policy before Recover opens the
  next one in the series.
- `Scope.ConvertRooms` (default `$false`).

### Fixed
- The step progress bar of the window stayed at 0% through most of a run: `[Math]::Min`/`Max` silently truncated
  the fraction to an integer when compared to the whole-number literals `0` and `1`.

### Lab-validated 9 Oct
- Shared wave of 3 and user wave of 2 recovered in a real tenant with the new wave pipeline: a single scheduler
  pause and a single delta cycle for the whole shared wave, a single delta cycle for the whole user wave.
- Case hold overflow confirmed live: with `Retention.HoldPolicyLimit` lowered to force it, the configured policy
  filled up and a sibling policy was created automatically in the same case and correctly held the rest.
## 1.1.0 - 2026-10-08

A window for every action, and waves of objects. The engine is unchanged when it runs without the window.

### Added
- **The window** (`Invoke-PraCloudMailbox.ps1 -Gui`, Windows PowerShell 5.1 and PowerShell 7): the next step, the
  configuration (certificate found and valid until), the last snapshot and Check, the batches and this computer at a
  glance; a page per action with its table of objects (filter by text, status, type, 21,000 objects); Convert and
  Recover by waves of ticked objects. Every action runs the same script in its own PowerShell (5.1 for Collect, 7
  for the others) with its log, transcript, journal and report, and the activity panel follows it step by step.
  *Convert…* and *Recover…* are enabled only after a preview of exactly the same selection, and ask for a typed
  confirmation (`CONVERT`, `RECOVER`). *Stop after the current object* ends the run cleanly (resumable). The
  questions of Entra Connect `Manual` mode are asked in the window. Fluent theme with PowerShell 7.5 and later.
- `-IdentityPath`: Check, Convert and Recover on a list of objects (one identity per line, or a CSV file with an
  `Identity` column); the identities found nowhere are listed as a warning.
- `PRA2.Common`: events of the run for the window (`PRA_EVENT_FILE`), stop request (`PRA_STOP_FILE`), questions
  to the operator in the window or in the console (`Request-PraOperator`).
- Tests: `PraCloudMailbox.Gui.Tests.ps1` (both editions, a real child run included); end-to-end scenarios for waves
  with `-IdentityPath`, a stop of Convert and of Recover followed by their resume, and the events of a run.

### Changed
- Entra Connect `Manual` mode asks through `Request-PraOperator`: in the window when the run belongs to it,
  otherwise in the console as before.
- The package contains `module\PRA2.Gui.psm1` and `module\PRA2.Gui.xaml`.

### Fixed
- Microsoft Graph calls are sent again when the connection failed without an answer (timeout, reset): seen in the
  lab just after a long delta cycle, it stopped a Recover (cleanly; running it again resumed the batch).
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
