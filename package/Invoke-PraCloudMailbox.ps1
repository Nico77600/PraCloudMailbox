<#
.SYNOPSIS
    PRA Cloud Mailbox - Exchange disaster recovery, scenario 2: Active Directory, Exchange and Entra Connect
    are lost; users and shared mailboxes get a mailbox in Exchange Online, and everything is rolled back
    when the on-premises infrastructure is rebuilt.

.DESCRIPTION
    Five actions:

      Collect   Before the disaster (PowerShell 7.5, from any domain-joined server that reaches Exchange.Server with
                Kerberos). Reads the mailboxes in
                scope, their addresses and GUIDs, the shared mailbox permissions (FullAccess, SendAs,
                SendOnBehalf) and optionally the contacts and distribution groups, and stores them as a
                snapshot in the SQLite database (Store.Path). Run it regularly; copy the database to the
                cloud admin server.
      Check     On the cloud admin server (PowerShell 7.5). Read-only: for every object of the last snapshot,
                checks its Entra ID object and its Exchange Online recipient, the licences and the app
                permissions, and tells whether Convert can run. With -Batch it follows a batch instead:
                it checks that its objects are in the state the batch was meant to give (-Expect) and waits
                for them, without changing anything.

      Convert   Disaster (PowerShell 7.5). Users: ExchangeGuid cleared, source of authority to the cloud, Exchange
                plan (licence group, Kiosk plan or direct licence): the Teams storage becomes the mailbox.
                Shared mailboxes, in waves (Licensing.Shared.Parallel, never using the Licensing.Shared.KeepFree
                units): temporary licence, shared, retention tag, licence given back, permissions from the snapshot.
                Every step is journaled (Store.JournalPath) under a batch ID.
      Recover   Infrastructure rebuilt (PowerShell 7.5). Users: source of authority back to AD, delta sync, Exchange
                plan removed; the cloud mailbox stays as ComponentShared under the case hold (Retention.HoldPolicy).
                Shared mailboxes: case hold, scheduler paused, identity deleted (the mailbox becomes inactive),
                deleted permanently, recreated from AD by a delta sync.
      Schedule  Before the disaster (PowerShell 7.5, in an elevated window, on the computer that runs Collect). Registers the daily Collect as a Windows scheduled task,
                after proving that its account can run it: a domain user (not SYSTEM, a computer or a managed service account), accepted by Exchange (a remote session, role
                group View-Only Organization Management) and, on a domain controller, allowed to log on as a batch job. Preview checks and shows what would be registered;
                Apply registers it (a task of the same name is replaced).

    Opt-in unsynchronised native rooms/equipment (CloudResource): Collect captures metadata, room lists,
    booking/delivery controls and work hours, never old contents or general mailbox rights. Convert creates
    required lists first, then each unlicensed, disabled-sign-in resource and its membership. New journals
    use protectionMode None: Recover has NO hold/NO inactive copy and does NOT preserve cloud-created
    reservations/messages. Export/reconcile and explicitly set Resources.AllowUnprotectedDeletion and
    OnPremisesReady to $true before deleting only the journal-owned cloud identity. Soft-deleted addresses
    remain Pending; no purge or hold/licence removal. Ambiguous deletion requests and restored identities
    are never automatically deleted again. Old protected journals and legacy synced paths are unchanged.

    Which objects: the whole snapshot, or -Identity (one object), -CsvPath (a list), -SearchBase (an OU) or -Group (a group of
    Active Directory captured by Collect, else of Microsoft Entra ID); -Scope and -MaxObjects narrow any of them. A member of a
    group that is not a mailbox is skipped, and the run says why.

    Every run starts in Preview. Collect -Mode Apply writes the snapshot (nothing is ever written to
    Exchange). Check is always read-only. Convert and Recover -Mode Apply ask one confirmation (or none
    with -Force); an interrupted batch is resumed with the same -Batch. Every change to Microsoft Entra ID,
    Exchange Online, the case hold or Entra Connect goes through ShouldProcess (the script has a High
    ConfirmImpact): -WhatIf shows a Preview, -Confirm asks before each change, -Force and -Confirm:$false
    ask nothing. A Preview never writes.

    The letter case of what you type never matters (parameter values, identities, file paths, batch IDs).

.PARAMETER Action
    Collect, Check, Convert, Recover or Schedule. Without it nothing runs: the script prints how to use it and ends with exit code 1
    (it never waits for an answer, so a scheduled task cannot hang).

.PARAMETER Mode
    Preview (default): read and report. Apply: Collect writes the snapshot to the database; Schedule registers the daily task.

.PARAMETER Identity
    One mailbox only (UPN, SMTP address, sAMAccountName, DN or GUID) instead of the configured scope. Check, Convert and Recover
    stop, before any change, when it is not in the snapshot (or in the batch).

.PARAMETER CsvPath
    Check, Convert and Recover: a wave - the objects listed in this file, either a CSV file with an Identity column (comma or
    semicolon, the header in any letter case) or one identity per line (UPN, SMTP address, sAMAccountName, DN or GUID; # starts a
    comment). The run stops, before any change, when an object is listed twice (the message gives the line numbers), when a line
    has no valid identity, or when an object is not in the snapshot (or in the batch). -IdentityPath is the former name of this
    parameter and still works.

.PARAMETER SearchBase
    Check, Convert and Recover: the objects of one organizational unit of the snapshot, sub-OUs included, given as a distinguished
    name (OU=Paris,DC=contoso,DC=com) or as a canonical name (contoso.com/Paris). The run stops, before any change, when the unit
    holds no object of the snapshot (or of the batch).

.PARAMETER Group
    Check, Convert and Recover: the members of one group, nested groups included (name, sAMAccountName, distinguished name, SMTP
    address or object ID). The group is looked up first among the groups that Collect captured while Active Directory was alive
    (Collect.TargetGroups, Scope.GroupDN), then in Microsoft Entra ID at the time of the run. A member that is not a mailbox (a
    computer, a contact...), that is not in the snapshot, or that the other options leave out is skipped, and the run says why.
    For Recover the group selects a wave inside the Convert batch.

.PARAMETER MaxObjects
    Check, Convert and Recover: the first N objects of the selection, in the order of the snapshot (0, the default, means no
    limit). Use it to work wave by wave: after a Convert, run the same command again for the next wave.

.PARAMETER Scope
    All (default), UsersOnly or SharedOnly (shared, room and equipment mailboxes).

.PARAMETER Snapshot
    Check and Convert: snapshot ID (default: the last complete snapshot).

.PARAMETER Batch
    Convert: resume this batch (-Identity, -CsvPath, -SearchBase, -Group, -Scope and -MaxObjects narrow the objects resumed). Recover:
    the Convert batch to roll back (printed at the end of Convert); an interrupted rollback goes on with the same -Batch (the ID of the rollback
    batch is accepted too). Check: follow this batch (see -Expect).

.PARAMETER Expect
    Check -Batch: the state the objects of the batch should be in. Provisioned (the default for a Convert batch): the cloud mailbox
    is there, with its type, plan and permissions. Deprovisioned (the default for a Recover batch): the object is back under Active
    Directory. Retained: the cloud data is protected as the journal says (case hold, inactive mailbox, native resources). Any
    letter case is accepted.

.PARAMETER Once
    Check -Batch: look once and report. Without it, Check waits (Polling settings of the configuration) until the objects are in
    the expected state or the time is over.

.PARAMETER ResolveOnly
    Check -Group: only says which group the name means and how many members it has, then stops. It looks the group up exactly like a real
    run (the snapshot first, then Microsoft Entra ID), prints the line of the run ("Group X (captured by Collect, snapshot N): M
    member(s) including nested groups, K skipped"), changes nothing and writes no report. Exit code 0 when the group is found, 1 when it
    is not found or the name matches several groups.

.PARAMETER AcceptUnverifiedSnapshot
    Check, Convert and Recover. Snapshot check: every Collect writes a .sha256 file beside the snapshot file (and beside its backup copy), and these actions compare the snapshot
    file with it before they read it. When the file was changed after its last Collect, the run stops, with nothing changed: restore the file from your external backup. Give this
    switch only when you know why the file changed and accept it as it is: the run then goes on, says so in the banner, the log, the summary and the report, and the next steps
    keep the switch. A snapshot without a .sha256 file (an old snapshot, or a copy made without it) stops a Convert or a Recover in Apply mode in the same way, because the tool
    cannot prove the file is the one Collect wrote; a Preview and a Check only warn. A batch also keeps the run and the content of the snapshot it was made from: a snapshot that
    is not that one stops the run too, and this switch accepts it. The journal is written by Convert and Recover themselves and is not covered.

.PARAMETER Credential
    Schedule: the domain user account that runs the daily Collect, with its password (Get-Credential). Without it, a Preview does not test the account and an Apply asks for it,
    which needs a window: with -NonInteractive, give -Credential. Windows stores the password for the task; it is never written to the log, the transcript or the report.

.PARAMETER At
    Schedule: the time of the daily task, on 24 hours (02:00). Default: Collect.Schedule.At of the configuration file.

.PARAMETER ConfigPath
    Configuration file. Default: config\PraCloudMailbox.config.psd1 next to this script.

.PARAMETER BaseConfigPath
    For the window: -ConfigPath is then the configuration generated for this run only (the settings changed in the window) and
    -BaseConfigPath is the configuration file itself, never modified. The settings that differ are listed in the banner, the log, the
    summary and the report, and the next steps show the original file.

.PARAMETER Force
    No confirmation prompt (the same as -Confirm:$false); never bypasses native-resource deletion consent or routing guards.
    Collect asks one only before writing a snapshot of ONE object (-Identity -Mode Apply).

.PARAMETER PassThru
    Also returns the result of the run as an object (Status, ExitCode, counters, file paths).

.PARAMETER Gui
    Opens the window: the state of the configuration, the snapshot and the batches, and every action with its
    preview, a typed confirmation and the progress. It runs this script for each action in a PowerShell 7.5
    window of its own, so the window and the command line do exactly the same.

.EXAMPLE
    .\Invoke-PraCloudMailbox.ps1 -Gui
    Opens the window (PowerShell 7.5 or later).

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Collect
    Preview of what would be collected from Exchange.Server (nothing is written).

.EXAMPLE
    .\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force
    Writes a new snapshot (scheduled task).

.EXAMPLE
    .\Invoke-PraCloudMailbox.ps1 -Action Schedule -Mode Apply
    Elevated PowerShell 7.5 on the Collect computer: asks for the account and its password, checks them, then registers the daily Collect task.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check
    PowerShell 7.5 or later on the cloud admin server: readiness of every object of the last snapshot.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply
    Disaster: converts every ready object of the last snapshot; prints the batch ID.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch 1a2b3c4d
    Infrastructure rebuilt: rolls back the Convert batch 1a2b3c4d.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch 1a2b3c4d -CsvPath .\data\wave1.csv
    Rolls back only the objects of the batch listed in wave1.csv (a wave).

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Group PRA-Wave1 -MaxObjects 200
    Converts the first 200 mailboxes of the group PRA-Wave1 (the group captured by Collect, else the Microsoft Entra group).

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -SearchBase 'contoso.com/Paris'
    Converts the objects of the organizational unit Paris and of its sub-units.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check -Group PRA-Wave1 -ResolveOnly
    Says which group PRA-Wave1 is and how many members it has, without touching anything.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check -Batch 1a2b3c4d -Expect Provisioned
    Follows the Convert batch 1a2b3c4d: waits until every object has its cloud mailbox, and changes nothing.

.NOTES
    Author     : Nicolas Fabert
    Version    : 2.1.0
    Requires   : PowerShell 7.5 or later for every action and for the window.
                 Collect: a domain-joined computer that reaches Exchange.Server with Kerberos.
                 Schedule: the same computer, in an elevated window.
                 Check, Convert, Recover: Microsoft.Graph.Authentication 2.41.1 or later, ExchangeOnlineManagement 3.10.1 or later.
    Exit codes : 0 = done, 1 = failed (or an object is not ready), 2 = done, next step required.
    Documentation : docs\PraCloudMailbox-UserGuide.md (user guide), docs\PraCloudMailbox-Guide.md (administrator guide)
#>
#Requires -Version 7.5
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High', DefaultParameterSetName = 'Run')]
param(
    [Parameter(ParameterSetName = 'Run')][ValidateSet('Collect','Check','Convert','Recover','Schedule')][string]$Action,
    [Parameter(ParameterSetName = 'Run')][ValidateSet('Preview','Apply')][string]$Mode = 'Preview',
    [Parameter(ParameterSetName = 'Run')][string]$Identity,
    [Parameter(ParameterSetName = 'Run')][Alias('IdentityPath')][string]$CsvPath,
    [Parameter(ParameterSetName = 'Run')][string]$SearchBase,
    [Parameter(ParameterSetName = 'Run')][string]$Group,
    [Parameter(ParameterSetName = 'Run')][ValidateRange(0, [int]::MaxValue)][int]$MaxObjects = 0,
    [Parameter(ParameterSetName = 'Run')][ValidateSet('All','UsersOnly','SharedOnly')][string]$Scope = 'All',
    [Parameter(ParameterSetName = 'Run')][ValidateRange(0, [long]::MaxValue)][long]$Snapshot = 0,
    [Parameter(ParameterSetName = 'Run')][string]$Batch,
    [Parameter(ParameterSetName = 'Run')][ValidateSet('Provisioned','Deprovisioned','Retained')][string]$Expect,
    [Parameter(ParameterSetName = 'Run')][switch]$Once,
[Parameter(ParameterSetName = 'Run')][switch]$ResolveOnly,
    [Parameter(ParameterSetName = 'Run')][switch]$AcceptUnverifiedSnapshot,
    [Parameter(ParameterSetName = 'Run')][pscredential]$Credential,
    [Parameter(ParameterSetName = 'Run')][string]$At,
    [string]$ConfigPath,
    [Parameter(ParameterSetName = 'Run')][string]$BaseConfigPath,
    [Parameter(ParameterSetName = 'Run')][switch]$Force,
    [Parameter(ParameterSetName = 'Run')][switch]$PassThru,
    [Parameter(Mandatory, ParameterSetName = 'Gui')][switch]$Gui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
# >>> PRA package integrity bootstrap: identical in every runtime PowerShell file (checked by the tests).
# It authenticates module\Test-PackageIntegrity.ps1 before running it; there is no unsigned or bypass path.
& {
    param([Parameter(Mandatory)][string]$PackageRoot)
    $verifier = [IO.Path]::GetFullPath([IO.Path]::Combine($PackageRoot, 'module', 'Test-PackageIntegrity.ps1'))
    $cursor = $verifier
    while ($cursor) {
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if (-not [string]::Equals([string]$signature.Status, 'Valid', [StringComparison]::Ordinal) -or $null -eq $signature.SignerCertificate -or
        -not [string]::Equals([string]$signature.SignerCertificate.Thumbprint, '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6', [StringComparison]::Ordinal) -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
    }
    & $verifier -Root $PackageRoot
} $PSScriptRoot
# <<< PRA package integrity bootstrap
# The options typed on the command line (a function sees its own $PSBoundParameters, not the one of the script).
$script:BoundOptions = [string[]]@($PSBoundParameters.Keys)
# Not a parameter default: the default path is computed here, and an explicit -ConfigPath is left as it is.
$defaultConfigPath = Join-Path $PSScriptRoot 'config\PraCloudMailbox.config.psd1'
if (-not $ConfigPath) { $ConfigPath = $defaultConfigPath }
# ValidateSet accepts any letter case (-Action convert) but binds the value AS TYPED. Every ValidateSet parameter is rewritten to its official
# spelling before anything reads it, so a run typed in lower case behaves and shows exactly like the official spelling. ValidateSet itself uses the
# culture comparison, which ignores invisible characters ("Apply" followed by U+200B passes): the lookup below compares by character code and
# refuses any value that has no match.
foreach ($parameterInfo in $MyInvocation.MyCommand.Parameters.Values) {
    if ($parameterInfo.ParameterType -ne [string]) { continue }
    $validSet = @($parameterInfo.Attributes | Where-Object { $_ -is [Management.Automation.ValidateSetAttribute] })
    if ($validSet.Count -ne 1) { continue }
    $typed = [string](Get-Variable -Name $parameterInfo.Name -ValueOnly)
    if (-not $typed) { continue }
    $official = @($validSet[0].ValidValues | Where-Object { [string]::Equals($_, $typed, [StringComparison]::OrdinalIgnoreCase) })
    if ($official.Count -ne 1) { throw ("Parameter -{0}: '{1}' is not one of the accepted values ({2})." -f $parameterInfo.Name, $typed, ($validSet[0].ValidValues -join ', ')) }
    Set-Variable -Name $parameterInfo.Name -Value ([string]$official[0])
    if ($PSBoundParameters.ContainsKey($parameterInfo.Name)) { $PSBoundParameters[$parameterInfo.Name] = [string]$official[0] }
}
# -Action is not Mandatory on purpose: PowerShell would ask for it with a prompt, and a scheduled or unattended run (no console, no keyboard) would wait for
# an answer forever. A missing -Action stops at once, with the usage.
if (-not $Action -and -not $Gui) {
    Write-Error -ErrorAction Continue -Message ('-Action is missing: nothing was done. Give the task: -Action Collect, Check, Convert, Recover or Schedule (Preview is the default mode). Example: .\Invoke-PraCloudMailbox.ps1 -Action Check. For the window: -Gui. Help: Get-Help .\Invoke-PraCloudMailbox.ps1 -Full.')
    exit 1
}

if ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion -lt [version]'7.5') {
    Write-Error 'PRA Cloud Mailbox runs in PowerShell 7.5 or later (pwsh.exe), for every action and for the window. Windows PowerShell 5.1 is not supported: open PowerShell 7 and run the same command.' -ErrorAction Continue
    exit 1
}

# The window. It runs this script for each action, in a child PowerShell 7.
if ($Gui) {
    Import-Module (Join-Path $PSScriptRoot 'PraCloudMailbox.psd1') -Force -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot 'module\PRA.Gui.psm1') -Force -ErrorAction Stop
    try { Show-PcmGui -Root $PSScriptRoot -ConfigPath $ConfigPath -Version (Get-PcmToolVersion); exit 0 }
    catch { Write-Error ('The window could not open: {0}' -f $_.Exception.Message) -ErrorAction Continue; exit 1 }
}

Import-Module (Join-Path $PSScriptRoot 'PraCloudMailbox.psd1') -Force -ErrorAction Stop
if ($Action -in @('Collect', 'Schedule')) { Import-Module (Join-Path $PSScriptRoot 'module\PRA.Collect.psm1') -Force -ErrorAction Stop }
else { Import-Module (Join-Path $PSScriptRoot 'module\PRA.Cloud.psm1') -Force -ErrorAction Stop }
# The warnings of the Microsoft modules called from this script go to the log, not to the console (Get-PcmModuleQuietDefault).
$PcmEntryModuleWarnings = [System.Collections.ArrayList]::new()
if ($Action -notin @('Collect', 'Schedule')) { $PSDefaultParameterValues = Get-PcmModuleQuietDefault -VariableName 'PcmEntryModuleWarnings' }
$toolVersion = Get-PcmToolVersion

# -ConfigPath is repeated in the commands proposed as next steps only when it is not the default file. Run from the window, -ConfigPath is the configuration
# generated for that run: the next steps show the original file (-BaseConfigPath) instead.
$script:ConfigPathBound = $false
if ($PSBoundParameters.ContainsKey('ConfigPath') -or $PSBoundParameters.ContainsKey('BaseConfigPath')) {
    try { $script:ConfigPathBound = -not (Test-PcmSameText (Resolve-PcmPath $(if ($BaseConfigPath) { $BaseConfigPath } else { $ConfigPath }) 'ConfigPath' $PSScriptRoot) (Resolve-PcmPath $defaultConfigPath 'ConfigPath' $PSScriptRoot)) }
    catch { $script:ConfigPathBound = $true }
}
# -Force acts like -Confirm:$false: the single confirmation of the Apply and the confirmation of each change are not asked.
# An explicit -Confirm (or -Confirm:$false) is the operator's own choice: the single confirmation is not asked either.
$script:ConfirmBound = $PSBoundParameters.ContainsKey('Confirm')
if ($Force -and -not $script:ConfirmBound) { $ConfirmPreference = 'None' }

$effectiveMode = if ($WhatIfPreference -or $Action -eq 'Check') { 'Preview' } else { $Mode }
$caller = $PSCmdlet
function New-PcmApprovalCallback {
    <# Callback asked before every change to Entra ID, Exchange Online, the case hold or Entra Connect (ShouldProcess: -WhatIf, -Confirm). A Preview never writes. #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only builds a callback; the callback itself calls ShouldProcess of the entry script.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'ShouldProcess is called on the cmdlet of the entry script, which supports it.')]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Both parameters are captured by the closure (GetNewClosure).')]
    param($Cmdlet, [hashtable]$RunContext)
    # Created in this function so that the closure only captures these two variables. ShouldProcess reads $ConfirmPreference from the calling
    # scope: it is taken from the run context ('None' with -Force, or after the one confirmation of the run).
    return {
        param($Target, $Operation)
        if (-not (Test-PcmSameText $RunContext.Mode 'Apply')) { throw ("A Preview never writes: '{0}' on {1} was not done." -f $Operation, $Target) }
        $ConfirmPreference = $RunContext.ConfirmPreference
        $Cmdlet.ShouldProcess([string]$Target, [string]$Operation)
    }.GetNewClosure()
}
$context = @{
    Root = $PSScriptRoot; Version = $toolVersion; RunId = ((Get-Date -Format 'yyyyMMdd_HHmmss') + '-' + [guid]::NewGuid().ToString('N'))
    StartTime = Get-Date; Action = $Action; Mode = $effectiveMode; Phase = ''
    CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; StepIndex = 0; StepTotal = 0; Warnings = 0
    Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]')
    BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
    Excluded = (New-Object 'Collections.Generic.List[string]')
    LogFolder = (Join-Path $PSScriptRoot 'logs'); ReportFolder = (Join-Path $PSScriptRoot 'reports')
    LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $false
    Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''; Config = @{}
    VerboseEnabled = ($VerbosePreference -eq 'Continue' -or $DebugPreference -eq 'Continue')
    BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null
    Approval = $null; ConfirmPreference = [string]$ConfirmPreference; FallbackNext = @(); NextOnFailure = $false; NoChangeMade = $false; Cancelled = $false
    SkippedMembers = (New-Object 'Collections.Generic.List[string]'); ObjectsText = ''; ConfigOverrides = @(); GroupTarget = $null; LeftForLater = 0
    CloudDescription = ''; GlobalReportPath = ''; SnapshotCheck = ''; AcceptUnverifiedSnapshot = [bool]$AcceptUnverifiedSnapshot
}
$context.Approval = New-PcmApprovalCallback $caller $context
# The next steps proposed when the time limit of the calls to Microsoft 365 has to end the run at once (a call that cannot be interrupted): asked by the time limit, never by the run.
$context.FallbackProvider = { Get-PcmFallbackNext -AfterTimeout }

function Skip-PcmStep {
    <# Prints a step that does not run in this execution (numbering stays in the announced order). #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Reason, [string]$Icon = 'Skip')
    Write-PcmStep -Context $context -Title $Title -Icon $Icon
    Write-PcmItem -Context $context -Status Skip -Text $Reason
    Write-PcmStepState -Context $context -State skipped
}

function Get-PcmPlanTitle {
    <# The titles of the steps of this run, in order, announced to the window before the first step (they are the titles that Write-PcmStep is called with). #>
    switch ($Action) {
        'Collect' { return @('Exchange on-premises connection', 'Mailboxes in scope', 'Shared mailbox permissions', 'Contacts and distribution groups', 'Target groups', 'Saving the snapshot') }
        'Schedule' { return @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task') }
        'Check' {
            if ($ResolveOnly) { return @('Collect snapshot') }
            if ($Batch) { return @('Batch journal', 'Collect snapshot', 'Microsoft 365 connection', 'Follow-up of the batch') }
            return @('Collect snapshot', 'Microsoft 365 connection', 'Tenant prerequisites', 'Objects')
        }
        'Convert' { return @('Collect snapshot', 'Microsoft 365 connection', 'Readiness', 'Plan and confirmation', 'Users: source of authority and Exchange plan', 'Users: cloud mailboxes', 'Shared mailboxes') }
        default { return @('Collect snapshot', 'Microsoft 365 connection', 'Prerequisites', 'Plan and confirmation', 'Users: back on-premises', 'Shared mailboxes', 'Licence group') }
    }
}

function New-PcmRow {
    <# One result row (CSV / HTML report). #>
    param([Parameter(Mandatory)][object]$Record)
    return [ordered]@{
        Identity = $(if ($Record.user_principal_name) { [string]$Record.user_principal_name } else { [string]$Record.primary_smtp_address })
        Kind = [string]$Record.kind; PrimarySmtpAddress = [string]$Record.primary_smtp_address; ObjectGuid = [string]$Record.object_guid
        ExchangeGuid = [string]$Record.exchange_guid; Action = $Action; Entra = ''; ExchangeOnline = ''; TeamsStorage = ''; Licence = ''; Holds = ''
        Permissions = ''; FinalStatus = 'Planned'; Detail = ''; Warnings = ''
    }
}

function Get-PcmScopeText {
    param([hashtable]$Config)
    $dot = [char]0x00B7
    if ($Identity) { return "one object: $Identity" }
    $s = $Config.Scope
    $base = switch ($s.Mode) { 'OU' { $s.SearchBase } 'Group' { "members of $($s.GroupDN)" } 'Csv' { "CSV $($s.CsvPath)" } default { if ($s.SearchBase) { $s.SearchBase } else { 'whole organisation' } } }
    $kinds = @(); if ($s.IncludeUsers) { $kinds += 'users' }; if ($s.IncludeShared) { $kinds += 'shared' }; if ($s.IncludeRoom) { $kinds += 'rooms' }; if ($s.IncludeEquipment) { $kinds += 'equipment' }
    if ($Scope -eq 'UsersOnly') { $kinds = @('users only') } elseif ($Scope -eq 'SharedOnly') { $kinds = @('shared mailboxes only') }
    return ('{0} {1} {2}' -f $base, $dot, ($kinds -join ' + '))
}

function ConvertTo-PcmIdentityKey {
    <# What identities are compared on: a GUID in any notation and letter case becomes the lower-case hyphenated form, any other text is kept (the comparison ignores its letter case). #>
    param([AllowNull()][AllowEmptyString()][string]$Value)
    $guid = [guid]::Empty
    if ($Value -and [guid]::TryParse($Value.Trim(), [ref]$guid)) { return $guid.ToString() }
    return [string]$Value
}

function Read-PcmIdentityFile {
    <#
    -CsvPath: one identity per line (# starts a comment, a blank line is ignored), or a CSV file with an Identity column (any letter case, anywhere among the columns,
    delimiter , or ;). A file saved by Excel or Notepad as "ANSI" is read with the ANSI code page of the computer, so that its accents are kept. The whole list is refused, with the line numbers, when an object is listed twice (letter case ignored) or when a line has no
    valid identity: the run must not start on a list that is not what the operator meant.
    #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "-CsvPath: file not found: $Path" }
    $lines = @([string](Read-PcmCsvText -Path $Path) -split '\r?\n')
    $entries = [System.Collections.Generic.List[object]]::new()
    $problems = [System.Collections.Generic.List[string]]::new()
    $header = ''; $delimiter = ','
    for ($index = 0; $index -lt $lines.Count; $index++) {
        $text = ([string]$lines[$index]).Trim()
        if (-not $text -or $text.StartsWith('#')) { continue }
        $number = $index + 1
        if (-not $header -and -not $entries.Count -and (Get-PcmCsvDelimiter $text)) {
            $header = $text; $delimiter = Get-PcmCsvDelimiter $text
            continue
        }
        $value = $text
        if ($header) {
            try { $value = ([string]((@($header, $text) | ConvertFrom-Csv -Delimiter $delimiter | Select-Object -First 1).Identity)).Trim() }
            catch { $value = '' }
        }
        if (-not $value -or $value -match '[\x00-\x1f]') { $problems.Add(('line {0} has no valid identity' -f $number)); continue }
        $entries.Add([pscustomobject]@{ Line = $number; Value = $value })
    }
    $firstLine = @{}
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $entries) {
        $key = ConvertTo-PcmIdentityKey $entry.Value
        if ($firstLine.ContainsKey($key)) { $problems.Add(('line {0} lists {1} again (already on line {2})' -f $entry.Line, $entry.Value, $firstLine[$key])); continue }
        $firstLine[$key] = $entry.Line
        $list.Add($entry.Value)
    }
    if ($problems.Count) {
        $shown = @($problems | Select-Object -First 10)
        throw ("-CsvPath: {0} cannot be used, nothing was done. Each object must be listed once and every line must hold an identity: {1}{2}." -f (Split-Path -Path $Path -Leaf), ($shown -join '; '), $(if ($problems.Count -gt 10) { ('; and {0} more' -f ($problems.Count - 10)) } else { '' }))
    }
    if (-not $list.Count) { throw "-CsvPath: no identity in $Path." }
    return $list.ToArray()
}

function Get-PcmIdentityList {
    <# The objects named one by one: -CsvPath (read by the main block into the context) or -Identity; none = no list. #>
    $stored = Get-PcmValue $context 'Identities' $null
    $list = @(@($stored) | Where-Object { $_ })
    if (-not $list.Count -and $Identity) { $list = @($Identity) }
    return $list
}

function Select-PcmIdentity {
    <# Objects whose keys match the identity list (letter case ignored, a GUID in any notation). An identity found nowhere refuses the run: nothing is changed. #>
    param([AllowEmptyCollection()][object[]]$Items = @(), [Parameter(Mandatory)][scriptblock]$Keys, [Parameter(Mandatory)][string]$Where)
    $wanted = @(Get-PcmIdentityList)
    if (-not $wanted.Count) { return $Items }
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($one in $wanted) { [void]$set.Add((ConvertTo-PcmIdentityKey $one)) }
    $found = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $selected = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Items) {
        $hit = $false
        foreach ($key in @(& $Keys $item)) {
            $normal = ConvertTo-PcmIdentityKey ([string]$key)
            if ($normal -and $set.Contains($normal)) { [void]$found.Add($normal); $hit = $true }
        }
        if ($hit) { $selected.Add($item) }
    }
    $missing = @($wanted | Where-Object { -not $found.Contains((ConvertTo-PcmIdentityKey $_)) })
    if ($missing.Count) {
        $context.NoChangeMade = $true
        $shown = (@($missing | Select-Object -First 10)) -join ', '
        if ($wanted.Count -eq 1) { throw "$($wanted[0]) is not in $Where. Nothing was changed." }
        throw ('{0} of the {1} identities asked for {2} not in {3}: {4}{5}. Nothing was changed: correct the list, or ask for the objects that exist.' -f $missing.Count, $wanted.Count, $(if ($missing.Count -eq 1) { 'is' } else { 'are' }), $Where, $shown, $(if ($missing.Count -gt 10) { ', ...' } else { '' }))
    }
    return $selected.ToArray()
}

function Get-PcmTargetKind {
    <# Which source says the objects of this run: Identity, Csv (-CsvPath), SearchBase, Group, or None (the whole snapshot). #>
    if ($Identity) { return 'Identity' }
    if ($CsvPath) { return 'Csv' }
    if ($SearchBase) { return 'SearchBase' }
    if ($Group) { return 'Group' }
    return 'None'
}

function Connect-PcmCloudSession {
    <# Signs in to Microsoft 365 once per run (a group read from Microsoft Entra ID needs it before the plan is made). Later steps reuse the sign-in. #>
    $known = [string](Get-PcmValue $context 'CloudDescription' '')
    if ($known) { return $known }
    try { $how = Connect-PcmCloud -Context $context } catch { $context.CloudConnectFailed = $true; throw }
    $context.CloudDescription = [string]$how
    return [string]$how
}

function Test-PcmRunOption {
    <# The options of the run that cannot go together or that do not suit the action: refused in plain words before anything is read or changed. #>
    $named = @(@(@('-Identity', $Identity), @('-CsvPath', $CsvPath), @('-SearchBase', $SearchBase), @('-Group', $Group)) | Where-Object { $_[1] })
    if ($named.Count -gt 1) {
        $context.NoChangeMade = $true; $context.OptionRefusal = 'Selectors'
        throw ('Choose one way to say which objects: {0} cannot be used together. -Identity names one object, -CsvPath a list, -SearchBase an organizational unit and -Group a group. -Scope and -MaxObjects can go with any of them.' -f ((@($named | ForEach-Object { $_[0] })) -join ' and '))
    }
    foreach ($name in @('CsvPath', 'SearchBase', 'Group')) {
        if (($script:BoundOptions -contains $name) -and [string]::IsNullOrWhiteSpace([string](Get-Variable -Name $name -ValueOnly))) { throw "-${name} is empty: give a value, or leave the option out." }
        $text = [string](Get-Variable -Name $name -ValueOnly)
        if ($text -match '[\x00-\x1f]') { throw "-${name} holds a control character (a pasted line break is the usual cause): retype it." }
    }
    if ($Action -eq 'Collect') {
        if ($CsvPath -or $SearchBase -or $Group -or $MaxObjects -gt 0) { throw '-CsvPath, -SearchBase, -Group and -MaxObjects are for Check, Convert and Recover. Collect reads the scope of the configuration (Scope.Mode, Collect.TargetGroups), or one object with -Identity.' }
        if ($Expect -or $Once) { throw '-Expect and -Once are for Check with -Batch.' }
    }
    if (Test-PcmSameText $Action 'Schedule') {
        $notUsed = @(@('Identity', 'CsvPath', 'SearchBase', 'Group', 'MaxObjects', 'Scope', 'Snapshot', 'Batch', 'Expect', 'Once', 'ResolveOnly', 'AcceptUnverifiedSnapshot') | Where-Object { $script:BoundOptions -contains $_ })
        if ($notUsed.Count) { throw ('-{0} {1} not used by Schedule, which only registers the daily Collect task. Schedule takes -Mode, -Credential, -At, -ConfigPath and -Force.' -f ($notUsed -join ', -'), $(if ($notUsed.Count -gt 1) { 'are' } else { 'is' })) }
    }
    elseif (($script:BoundOptions -contains 'Credential') -or ($script:BoundOptions -contains 'At')) { throw '-Credential and -At are for Schedule, which registers the daily Collect task.' }
    if ($ResolveOnly -and -not ((Test-PcmSameText $Action 'Check') -and $Group -and -not $Batch)) { throw '-ResolveOnly is for Check with -Group (Check -Group <name> -ResolveOnly): it only says which group is meant and how many members it has.' }
    if (($Expect -or $Once) -and -not ((Test-PcmSameText $Action 'Check') -and $Batch)) { throw '-Expect and -Once are for Check with -Batch: they follow a batch that Convert or Recover started (Check -Batch <id>).' }
}

function Add-PcmSkippedMember {
    <# An object of the group (or of the list) that the run does not take, with the reason: one line on the console for the first ones, every one in the log and in the report. #>
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Reason)
    $list = $context.SkippedMembers
    [void]$list.Add(('{0} ({1})' -f $Name, $Reason))
    $text = '{0}: not taken, {1}' -f $Name, $Reason
    if ($list.Count -le 25) { Write-PcmItem -Context $context -Status Skip -Text $text }
    else {
        Write-PcmLog -Context $context -Message $text -Level Detail
        if ($list.Count -eq 26) { Write-PcmItem -Context $context -Status Info -Text 'More members are skipped: the log and the report name them all.' }
    }
}

function Get-PcmGroupMemberLabel {
    param([Parameter(Mandatory)][object]$Member)
    foreach ($name in @('Upn', 'Smtp', 'Name', 'Dn', 'Guid', 'Id')) { $value = [string](Get-PcmValue $Member $name ''); if ($value) { return $value } }
    return '(unnamed object)'
}

function Resolve-PcmGroupTarget {
    <#
    The group of -Group: among the groups that Collect captured in the snapshot (name, sAMAccountName, distinguished name, SMTP address, display name or GUID, the
    letter case ignored), and only when it is not there, in Microsoft Entra ID at this moment. Several matches refuse the run with the list. A snapshot written before
    the capture existed has no group: the group is read from Microsoft Entra ID.
    #>
    param([Parameter(Mandatory)][object]$Snapshot, [Parameter(Mandatory)][object]$Found, [AllowEmptyCollection()][object[]]$CapturedMembers = @(), [AllowEmptyCollection()][object[]]$Records = @())
    $value = ([string]$Group).Trim()
    if ($Found.Groups.Count -gt 1) {
        $context.NoChangeMade = $true
        $list = (@($Found.Groups | Select-Object -First 10 | ForEach-Object { '{0} ({1})' -f $_.name, $(if ($_.distinguished_name) { $_.distinguished_name } else { $_.group_guid }) }) -join '; ')
        throw ('{0} groups of snapshot {1} match "{2}": {3}. Nothing was changed: give the distinguished name or the object ID of the group you want.' -f $Found.Groups.Count, $Snapshot.id, $value, $list)
    }
    if ($Found.Groups.Count -eq 1) {
        $captured = $Found.Groups[0]
        $members = @(foreach ($row in $CapturedMembers) {
                [pscustomobject]@{ Class = [string]$row.member_class; Name = [string]$row.member_name; Guid = [string]$row.member_guid; Upn = [string]$row.member_upn
                    Smtp = [string]$row.member_smtp; Dn = [string]$row.member_dn; ImmutableId = ''; Id = '' }
            })
        $name = [string]$captured.name
        if (-not $name) { $name = [string]$captured.sam_account_name }
        if (-not $name) { $name = $value }
        return [pscustomobject]@{ Name = $name; Origin = ('captured by Collect, snapshot {0}' -f $Snapshot.id); From = 'Snapshot'
            Members = @($members | Where-Object { $_.Class -ne 'Group' }); Nested = @($members | Where-Object { $_.Class -eq 'Group' }).Count }
    }
    $context.CurrentOperation = 'Resolve-Group'
    if ($Found.HasTables) { Write-PcmItem -Context $context -Status Info -Icon Cloud -Text ('Group {0} is not among the groups captured in snapshot {1}: reading it from Microsoft Entra ID now.' -f $value, $Snapshot.id) }
    else { Write-PcmItem -Context $context -Status Info -Icon Cloud -Text ('Snapshot {0} was written before groups were captured: group {1} is read from Microsoft Entra ID now.' -f $Snapshot.id, $value) }
    $context.NativeResourceScope = @($Records | Where-Object { Test-PcmNativeResource -Config $context.Config -Record $_ }).Count -gt 0
    $how = Connect-PcmCloudSession
    Write-PcmItem -Context $context -Status Ok -Text $how
    if (-not (Test-PcmGroupReadRole -Roles @(Get-PcmGraphRole))) {
        $context.NoChangeMade = $true
        throw 'This sign-in cannot read groups in Microsoft Entra ID: none of GroupMember.Read.All, Group.Read.All, GroupMember.ReadWrite.All, Group.ReadWrite.All or Directory.Read.All is granted to it. Grant one of them, or capture the group with Collect while Active Directory is alive. Nothing was changed.'
    }
    $groups = @(Find-PcmEntraGroup -Value $value)
    if (-not $groups.Count) {
        $context.NoChangeMade = $true
        throw ('Group {0} is not in snapshot {1} nor in Microsoft Entra. Nothing was changed: check the name (the display name, the e-mail address, the mail nickname or the object ID), or add the group to Collect.TargetGroups and collect again.' -f $value, $Snapshot.id)
    }
    if ($groups.Count -gt 1) {
        $context.NoChangeMade = $true
        $list = (@($groups | Select-Object -First 10 | ForEach-Object { '{0} ({1}, {2})' -f $_.displayName, $(if ($_.mail) { $_.mail } else { 'no e-mail address' }), $_.id }) -join '; ')
        throw ('{0} Microsoft Entra groups match "{1}": {2}. Nothing was changed: give the object ID of the group you want.' -f $groups.Count, $value, $list)
    }
    $found = $groups[0]
    $all = @(Get-PcmEntraGroupMember -GroupId ([string]$found.id))
    return [pscustomobject]@{ Name = [string]$found.displayName; Origin = 'Microsoft Entra group, read now'; From = 'Entra'
        Members = @($all | Where-Object { $_.Class -ne 'Group' } | ForEach-Object {
                [pscustomobject]@{ Class = $_.Class; Name = $_.Name; Guid = ''; Upn = $_.Upn; Smtp = $_.Smtp; Dn = ''; ImmutableId = $_.ImmutableId; Id = $_.Id } })
        Nested = @($all | Where-Object { $_.Class -eq 'Group' }).Count }
}

function Get-PcmKindExclusion {
    <# Why a kind of mailbox is left out by -Scope or by Scope.Include* of the configuration (empty text: it is taken). #>
    param([Parameter(Mandatory)][string]$Kind)
    if ((Test-PcmSameText $Scope 'UsersOnly') -and -not (Test-PcmSameText $Kind 'User')) { return 'a {0} mailbox, left out by -Scope UsersOnly' -f $Kind.ToLowerInvariant() }
    if ((Test-PcmSameText $Scope 'SharedOnly') -and (Test-PcmSameText $Kind 'User')) { return 'a user mailbox, left out by -Scope SharedOnly' }
    $switches = @{ User = 'IncludeUsers'; Shared = 'IncludeShared'; Room = 'IncludeRoom'; Equipment = 'IncludeEquipment' }
    $setting = $switches[$Kind]
    if ($setting -and -not [bool](Get-PcmValue $context.Config.Scope $setting $true)) { return 'a {0} mailbox, left out by Scope.{1} = $false' -f $Kind.ToLowerInvariant(), $setting }
    return ''
}

function Get-PcmNotInSnapshotReason {
    param([Parameter(Mandatory)][object]$Member, [Parameter(Mandatory)][string]$Where)
    switch -Regex ([string]$Member.Class) {
        '^Computer$' { return 'a computer, not a mailbox' }
        '^Contact$' { return 'a contact, not a mailbox' }
        '^Device$' { return 'a device, not a mailbox' }
        '^ServicePrincipal$' { return 'an application, not a mailbox' }
        '^(User|Unknown|)$' { return ('not in {0}: no mailbox, or outside the scope of the Collect' -f $Where) }
        default { return ('a {0} object, not a mailbox' -f ([string]$Member.Class).ToLowerInvariant()) }
    }
}

function Select-PcmTarget {
    <#
    The objects chosen by the source of the run (-Identity, -CsvPath, -SearchBase or -Group; none = every object), then -Scope. Items are snapshot mailboxes or the
    items of a batch (-RecordOf gives the snapshot mailbox of an item); -AllRecords are all the mailboxes of the snapshot, to tell a member that is not in the batch from
    one that is not in the snapshot at all; -Known holds a reason for objects of the batch that are not left to do. A list or an OU that gives no object, or a group
    without a usable member, refuses the run: nothing is changed.
    #>
    param([AllowEmptyCollection()][object[]]$Items = @(), [Parameter(Mandatory)][scriptblock]$Keys, [Parameter(Mandatory)][scriptblock]$RecordOf, [Parameter(Mandatory)][string]$Where,
        [AllowEmptyCollection()][object[]]$AllRecords = @(), [hashtable]$Known = @{}, [AllowEmptyCollection()][object[]]$KnownItems = @(), [switch]$InBatch, [switch]$SkipConverted)
    $kind = Get-PcmTargetKind
    $selected = @($Items)
    if ($kind -in @('Identity', 'Csv') -and @($KnownItems).Count) {
        # A list given again to resume a batch names objects that are already done: they are skipped with their reason, never refused.
        $selected = @()
        foreach ($item in @(Select-PcmIdentity -Items (@($Items) + @($KnownItems)) -Keys $Keys -Where $Where)) {
            $guid = [string]$item.object_guid
            if (-not $Known.ContainsKey($guid)) { $selected += $item; continue }
            $record = & $RecordOf $item
            $label = if ($record -and [string]$record.user_principal_name) { [string]$record.user_principal_name } elseif ($record -and [string]$record.primary_smtp_address) { [string]$record.primary_smtp_address } else { [string]$item.identity }
            Add-PcmSkippedMember -Name $label -Reason ([string]$Known[$guid])
        }
        if (-not $selected.Count) {
            $context.NoChangeMade = $true
            throw ('Every object asked for is already done ({0}): nothing is left to do in {1}. Nothing was changed.' -f [string]($Known.Values | Select-Object -First 1), $Where)
        }
    }
    elseif ($kind -in @('Identity', 'Csv')) { $selected = @(Select-PcmIdentity -Items $Items -Keys $Keys -Where $Where) }
    elseif ($kind -eq 'SearchBase') {
        $selected = @($Items | Where-Object {
                $record = & $RecordOf $_
                $record -and (Test-PcmInSearchBase -SearchBase $SearchBase -DistinguishedName ([string](Get-PcmValue $record 'distinguished_name' '')) -OrganizationalUnit ([string](Get-PcmValue $record 'organizational_unit' '')))
            })
        if (-not $selected.Count) {
            $context.NoChangeMade = $true
            throw ('-SearchBase {0} gives no object in {1}. Nothing was changed: check the spelling (a distinguished name such as OU=Paris,DC=contoso,DC=com, or a canonical name such as contoso.com/Paris).' -f $SearchBase, $Where)
        }
    }
    elseif ($kind -eq 'Group') {
        $target = $context.GroupTarget
        if (-not $target) { throw 'The group was not resolved before the objects were chosen.' }
        $index = New-PcmRecordIndex -Records $AllRecords
        $inList = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($item in $Items) { $record = & $RecordOf $item; if ($record) { [void]$inList.Add((ConvertTo-PcmGuidKey ([string]$record.object_guid))) } }
        $chosen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $handled = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $excluded = @(@(Get-PcmValue $context.Config.Scope 'ExcludeSamAccountNames' @()) | ForEach-Object { [string]$_ })
        foreach ($member in @($target.Members)) {
            $label = Get-PcmGroupMemberLabel -Member $member
            $record = Find-PcmRecord -Index $index -Guid ([string]$member.Guid) -ImmutableId ([string]$member.ImmutableId) -Upn ([string]$member.Upn) -Smtp ([string]$member.Smtp)
            if (-not $record) { Add-PcmSkippedMember -Name $label -Reason (Get-PcmNotInSnapshotReason -Member $member -Where ('snapshot {0}' -f (Get-PcmValue $context 'SnapshotId' ''))); continue }
            $key = ConvertTo-PcmGuidKey ([string]$record.object_guid)
            if (-not $handled.Add($key)) { continue }
            $label = if ([string]$record.user_principal_name) { [string]$record.user_principal_name } elseif ([string]$record.primary_smtp_address) { [string]$record.primary_smtp_address } else { $label }
            $reason = ''
            $sam = [string](Get-PcmValue $record 'sam_account_name' '')
            if ((Test-PcmSystemMailboxName ([string](Get-PcmValue $record 'name' ''))) -or ($sam -and @($excluded | Where-Object { Test-PcmSameText $_ $sam }).Count)) {
                $reason = 'excluded (a system mailbox, or listed in Scope.ExcludeSamAccountNames)'
            }
            if (-not $reason) { $reason = Get-PcmKindExclusion -Kind ([string]$record.kind) }
            if (-not $reason -and -not $inList.Contains($key)) {
                $reason = if ($Known.ContainsKey($key)) { [string]$Known[$key] } elseif ($InBatch) { 'not in this batch' } else { 'not selected' }
            }
            if (-not $reason -and $SkipConverted -and $context.Journal) {
                $converted = Find-PcmConvertedBatch -ObjectGuid ([string]$record.object_guid)
                if ($converted) { $reason = 'already converted by Convert batch {0}' -f $converted.id }
            }
            if ($reason) { Add-PcmSkippedMember -Name $label -Reason $reason; continue }
            [void]$chosen.Add($key)
        }
        $selected = @($Items | Where-Object { $record = & $RecordOf $_; $record -and $chosen.Contains((ConvertTo-PcmGuidKey ([string]$record.object_guid))) })
        if (-not $selected.Count -and -not $ResolveOnly) {
            $context.NoChangeMade = $true
            throw ('Group {0}: none of its {1} member(s) can be taken in {2} (see the reasons above). Nothing was changed.' -f $target.Name, @($target.Members).Count, $Where)
        }
    }
    if ($Scope -eq 'UsersOnly') { $selected = @($selected | Where-Object { Test-PcmSameText ([string](Get-PcmValue $_ 'kind' '')) 'User' }) }
    elseif ($Scope -eq 'SharedOnly') { $selected = @($selected | Where-Object { -not (Test-PcmSameText ([string](Get-PcmValue $_ 'kind' '')) 'User') }) }
    return $selected
}

function Limit-PcmTarget {
    <# -MaxObjects: the first N objects, in the order of the list. For a new Convert the objects already converted by an earlier batch do not count: the same command goes on with the next wave. #>
    param([AllowEmptyCollection()][object[]]$Items = @(), [switch]$SkipConverted)
    $context.LeftForLater = 0
    $max = [int]$MaxObjects
    if ($max -le 0 -or @($Items).Count -eq 0) { return @($Items) }
    if ($SkipConverted -and $context.Journal -and (Get-PcmTargetKind) -in @('None', 'SearchBase')) {
        $taken = [System.Collections.Generic.List[object]]::new(); $left = 0; $already = 0
        foreach ($item in $Items) {
            if (Find-PcmConvertedBatch -ObjectGuid ([string]$item.object_guid)) { $already++; continue }
            if ($taken.Count -lt $max) { $taken.Add($item) } else { $left++ }
        }
        if ($already) { Write-PcmItem -Context $context -Status Info -Text ('{0} object(s) already converted by an earlier batch are left out.' -f $already) }
        $context.LeftForLater = $left
        return $taken.ToArray()
    }
    $all = @($Items)
    if ($all.Count -gt $max) { $context.LeftForLater = $all.Count - $max }
    return @($all | Select-Object -First $max)
}

function Write-PcmTargetHeader {
    <# The line that says which objects the run takes and where they come from (group, OU, list), the skipped members and the -MaxObjects limit; also kept for the summary and the report. #>
    param([int]$Count = 0)
    $dot = [char]0x00B7
    $kind = Get-PcmTargetKind
    $text = ''
    $show = $MaxObjects -gt 0
    switch ($kind) {
        'Group' {
            $show = $true
            $target = $context.GroupTarget
            $text = 'Group {0} ({1}): {2} member(s) including nested groups{3}{4}' -f $target.Name, $target.Origin, @($target.Members).Count,
                $(if ($target.Nested) { ', {0} nested group(s)' -f $target.Nested } else { '' }), $(if ($context.SkippedMembers.Count) { ', {0} skipped' -f $context.SkippedMembers.Count } else { '' })
        }
        'SearchBase' { $show = $true; $text = 'OU {0}: {1} object(s)' -f $SearchBase, $Count }
        'Csv' { $text = 'List {0}: {1} object(s)' -f (Split-Path -Path $CsvPath -Leaf), @($context.Identities).Count }
        'Identity' { $text = "One object: $Identity" }
    }
    if ($MaxObjects -gt 0) {
        $limit = 'limited to the first {0} object(s) (-MaxObjects){1}' -f $MaxObjects, $(if ($context.LeftForLater) { ', {0} left for a later run' -f $context.LeftForLater } else { '' })
        $text = if ($text) { "$text $dot $limit" } else { $limit }
    }
    if (-not $text) { return }
    $context.ObjectsText = $text
    $targetEvent = @{ text = $text; selection = $kind; count = $Count }
    if ($kind -eq 'Group' -and $context.GroupTarget) {
        $targetEvent.name = [string]$context.GroupTarget.Name; $targetEvent.source = [string]$context.GroupTarget.From; $targetEvent.origin = [string]$context.GroupTarget.Origin
        $targetEvent.members = @($context.GroupTarget.Members).Count; $targetEvent.nested = [int]$context.GroupTarget.Nested; $targetEvent.skipped = $context.SkippedMembers.Count
    }
    Write-PcmEvent 'target' $targetEvent
    if ($show) { Write-PcmItem -Context $context -Status Info -Icon Target -Text $text }
}

function ConvertTo-PcmLiteral {
    <# A value as a single-quoted PowerShell literal: the apostrophe AND the typographic quotes that PowerShell also reads as quotes are doubled. #>
    param([AllowEmptyString()][string]$Value)
    $escaped = $Value.Replace("'", "''")
    foreach ($code in @(0x2018, 0x2019, 0x201A, 0x201B)) { $quote = [string][char]$code; $escaped = $escaped.Replace($quote, $quote + $quote) }
    return "'" + $escaped + "'"
}

function Get-PcmCommandLine {
    <#
    The command line proposed as a next step: this script as it is run from the package folder (.\Invoke-PraCloudMailbox.ps1), with the options of this run that still apply
    (-KeepFilters: -Identity, -CsvPath, -SearchBase, -Group, -Scope, -MaxObjects and -Snapshot of the run, as typed). Every free-text value is quoted as a literal. -ConfigPath is repeated
    only when it is not the default file. Add a new filter of the run to the list below.
    #>
    param([Parameter(Mandatory)][string]$ActionName, [string]$ModeName = '', [string]$BatchId = '', [switch]$KeepFilters, [switch]$OtherServer, [string]$ExpectName = '', [switch]$NoOnce,
        [ValidateSet('', 'Identity', 'SearchBase', 'CsvPath', 'Group', 'None')][string]$Only = '')
    $root = [string](Get-PcmValue $context 'Root' '')
    $parts = @('.\Invoke-PraCloudMailbox.ps1', "-Action $ActionName")
    if (Test-PcmSameText $ModeName 'Apply') { $parts += '-Mode Apply' }
    if ($BatchId) { $parts += $(if ($BatchId -match '^[0-9a-f]+$') { "-Batch $BatchId" } else { '-Batch ' + (ConvertTo-PcmLiteral $BatchId) }) }
    if ($KeepFilters) {
        $filters = [ordered]@{ Identity = $Identity; CsvPath = $CsvPath; SearchBase = $SearchBase; Group = $Group }
        # With -Only, the one object selector kept (-Identity, -SearchBase, -CsvPath or -Group; None keeps no selector): the others are left out.
        foreach ($name in $filters.Keys) { if ($filters[$name] -and (-not $Only -or (Test-PcmSameText $Only $name))) { $parts += ('-{0} {1}' -f $name, (ConvertTo-PcmLiteral ([string]$filters[$name]))) } }
        if (-not (Test-PcmSameText $Scope 'All')) { $parts += "-Scope $Scope" }
        if ($MaxObjects -gt 0) { $parts += "-MaxObjects $MaxObjects" }
        if ($Snapshot -gt 0 -and -not $BatchId -and -not (Test-PcmSameText $ActionName 'Recover')) { $parts += "-Snapshot $Snapshot" }
        if ([bool](Get-PcmValue $context 'AcceptUnverifiedSnapshot' $false) -and -not (Test-PcmSameText $ActionName 'Collect')) { $parts += '-AcceptUnverifiedSnapshot' }
    }
    # Follow-up of a batch (Check -Batch): the state expected, and the single look when it was asked for.
    if ((Test-PcmSameText $ActionName 'Check') -and $BatchId) {
        $expectation = if ($ExpectName) { $ExpectName } elseif ($KeepFilters -and $Expect) { $Expect } else { '' }
        if ($expectation) { $parts += "-Expect $expectation" }
        if ($KeepFilters -and $Once -and -not $NoOnce) { $parts += '-Once' }
    }
    # The time of a Schedule, when it was typed.
    if ((Test-PcmSameText $ActionName 'Schedule') -and $At) { $parts += $(if ($At -match '^[0-9]{1,2}:[0-9]{2}$') { "-At $At" } else { '-At ' + (ConvertTo-PcmLiteral $At) }) }
    # On another server (the cloud admin server, after a Collect) the configuration file is that server's own: -ConfigPath is not repeated.
    if ($script:ConfigPathBound -and -not $OtherServer) { $parts += ('-ConfigPath ' + (ConvertTo-PcmLiteral (Resolve-PcmPath $(if ($BaseConfigPath) { $BaseConfigPath } else { $ConfigPath }) 'ConfigPath' $(if ($root) { $root } else { (Get-Location).Path })))) }
    return ($parts -join ' ')
}

function Get-PcmFallbackNext {
    <#
    The next step of a run that ends in failure when no step set a more specific one: what the failed run kept, whether anything was changed, and the exact
    command to run again (the command of this run as typed; a batch already started is continued by its ID). -AfterTimeout: the run stopped because Microsoft 365 did not answer.
    #>
    param([switch]$AfterTimeout)
    $fix = if ($AfterTimeout) { 'Check the service health of Microsoft 365 (https://status.cloud.microsoft)' } else { 'Fix the problem named above' }
    $isCheck = Test-PcmSameText $Action 'Check'
    $isCollect = Test-PcmSameText $Action 'Collect'
    $applied = (Test-PcmSameText $effectiveMode 'Apply') -and -not $isCheck
    $typedBatch = if ($Batch) { $Batch } else { '' }
    $modeName = if ($isCheck) { '' } else { $effectiveMode }
    $same = Get-PcmCommandLine -ActionName $Action -ModeName $modeName -BatchId $typedBatch -KeepFilters
    if ($isCheck -and [bool](Get-Variable -Name ResolveOnly -ValueOnly -ErrorAction SilentlyContinue)) { $same += ' -ResolveOnly' }
    if ((Get-PcmValue $context 'OptionRefusal' '') -eq 'Selectors') {
        $given = @(@('Identity', $Identity), @('CsvPath', $CsvPath), @('SearchBase', $SearchBase), @('Group', $Group) | Where-Object { $_[1] } | ForEach-Object { $_[0] })
        $steps = @(('These options select the objects in different ways and cannot be combined ({0}): keep one of them. Run the command with the one you want:' -f (($given | ForEach-Object { '-' + $_ }) -join ', ')))
        foreach ($name in $given) {
            $steps += ('Only -{0} (remove {1}):' -f $name, ((@($given | Where-Object { -not (Test-PcmSameText $_ $name) }) | ForEach-Object { '-' + $_ }) -join ', '))
            $steps += (Get-PcmCommandLine -ActionName $Action -ModeName $modeName -BatchId $typedBatch -KeepFilters -Only $name)
        }
        return $steps
    }
    if ([bool](Get-PcmValue $context 'Cancelled' $false)) { return @('Nothing was written. To run it, give the same command again and answer yes to the confirmation:', $same) }
    if ([bool](Get-PcmValue $context 'ConfirmationUnavailable' $false)) { return @('Nothing was written. This console cannot ask for the confirmation: add -Force to the command, or run it in an interactive console and answer yes:', ($same + ' -Force')) }
    if ($isCheck) { return @("Check only reads: nothing was changed. $fix, then run the same command again:", $same) }
    if (-not $applied) { return @("This was a Preview: nothing was changed. $fix, then run the same command again:", $same) }
    if (Test-PcmSameText $Action 'Schedule') { return @('The task was not registered, or it was left as it was. Fix the problem named above, then run the same command again:', $same) }
    if ($isCollect) { return @('Nothing was changed on Exchange (Collect only reads). No new snapshot became the last complete one if the run stopped before the end. Fix the problem named above, then run the same command again:', $same) }
    $batchId = [string]$context.BatchId
    if ($batchId) {
        $resume = if (Test-PcmSameText $Action 'Convert') { Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -BatchId $batchId -KeepFilters }
            else { Get-PcmCommandLine -ActionName 'Recover' -ModeName 'Apply' -BatchId $(if ($typedBatch) { $typedBatch } else { $batchId }) -KeepFilters }
        return @(('What this run already did is kept in the journal (batch {0}) and is never repeated. {1}, then continue the batch:' -f $batchId, $fix), $resume)
    }
    # A Recover that continues a batch which already exists and stopped before it took the batch back (the sign-in, for example): the batch and its journal are intact.
    $existingBatch = [string](Get-PcmValue $context 'ResumeBatchId' '')
    if ($existingBatch) { return @(('The Recover batch {0} already exists and is kept in the journal: what it already did is never repeated, and the command below continues it. {1}, then run it:' -f $existingBatch, $fix), $same) }
    if ([bool](Get-PcmValue $context 'NoChangeMade' $false)) { return @("Nothing was changed. $fix, then run the same command again:", $same) }
    return @("The run stopped before it wrote its journal: nothing was changed in the tenant. $fix, then run the same command again:", $same)
}

function Test-PcmCloudModuleAvailable {
    <#
    Before any change, checks the cloud modules of this computer: Microsoft.Graph.Authentication and ExchangeOnlineManagement at the minimum version
    or later (Get-PcmCloudModuleReport). The cloud actions all sign in, the Preview and Check included: the run stops here, with nothing changed, and the
    message gives the versions found, the minimum and the exact Install-Module command (one module per command).
    #>
    $context.CurrentOperation = 'Check-CloudModules'
    $report = Get-PcmCloudModuleReport -MinExo ([string](Get-PcmValue (Get-PcmValue $context.Config 'Exo' @{}) 'MinModuleVersion' '3.10.1'))
    if ($report.Ok) {
        Write-PcmItem -Context $context -Status Ok -Icon Cloud -Text ('Cloud modules: Microsoft.Graph.Authentication {0}, ExchangeOnlineManagement {1}' -f $report.Graph.Version, $report.Exo.Version)
        return
    }
    $context.NoChangeMade = $true; $context.CurrentIdentity = ''
    throw ('The cloud part cannot run on this computer. {0} Nothing was changed.' -f ($report.Messages -join ' '))
}

function Test-PcmStop {
    <#
    Stop asked by the window (PRA_STOP_FILE): no new object is started. The caller leaves the objects not started as
    Pending (Set-PcmRowStopped); the run then ends normally (scheduler resumed, journal, report) and can be resumed.
    #>
    if (-not (Test-PcmStopRequest)) { return $false }
    if (-not (Get-PcmValue $context 'StopNoted' $false)) {
        $context['StopNoted'] = $true
        Write-PcmLog -Context $context -Message 'Stop requested by the operator: no new object is started; the objects not started stay as they are (Pending). Run the same command again to go on.' -Level Warning
    }
    return $true
}

function Set-PcmRowStopped {
    <# An object not started because of a stop request: Pending, without a console line each (there can be thousands). #>
    param([Parameter(Mandatory)][object]$Row)
    $Row.FinalStatus = 'Pending'
    $Row.Detail = 'not started: stop requested by the operator; run the same command again to go on'
}

# =================================================================================================
# Collect - Exchange on-premises, before the disaster (remote session to Exchange.Server).
# =================================================================================================
function Get-PcmCollectSizeProblem {
    <#
    Before a snapshot is written: is its size acceptable? A snapshot without any mailbox is refused, and so is one that holds more than Collect.MaxDropPercent percent fewer
    mailboxes than the last complete snapshot of the same scope (a full one, not a one-object snapshot): it would become the last complete snapshot and push the good ones out of
    Store.KeepSnapshots. Collect.MaxDropPercent = 100 accepts any size. Returns the plain message for the operator, or '' when the size is acceptable.
    #>
    param([Parameter(Mandatory)][int]$Count, [Parameter(Mandatory)][string]$ScopeJson)
    $config = $context.Config
    $limit = [int](Get-PcmValue $config.Collect 'MaxDropPercent' 50)
    if ($limit -ge 100) { return '' }
    $previous = $null
    if ([IO.File]::Exists($config.Store.Path)) {
        $connection = $null
        try {
            $connection = Open-PcmStore -Path $config.Store.Path -Root $context.Root -ReadOnly
            $now = $ScopeJson | ConvertFrom-PcmJson
            $oneObject = [bool][string](Get-PcmValue $now 'Identity' '')
            foreach ($row in @(if (-not $oneObject) { Invoke-PcmSql $connection "SELECT id, scope_json, counts_json FROM snapshot WHERE status = 'Complete' ORDER BY id DESC LIMIT 50" -As Rows })) {
                $then = $null; try { $then = [string]$row.scope_json | ConvertFrom-PcmJson } catch { $then = $null }
                if (-not $then -or [string](Get-PcmValue $then 'Identity' '')) { continue }
                $same = $true
                foreach ($name in @('Mode', 'SearchBase', 'GroupDN', 'CsvPath', 'Filter', 'IncludeUsers', 'IncludeShared', 'IncludeRoom', 'IncludeEquipment')) {
                    if (-not (Test-PcmSameText ([string](Get-PcmValue $then $name '')) ([string](Get-PcmValue $now $name '')))) { $same = $false; break }
                }
                if (-not $same) { continue }
                $mailboxes = [int](Get-PcmValue ([string]$row.counts_json | ConvertFrom-PcmJson) 'mailbox' 0)
                $previous = [pscustomobject]@{ Id = [long]$row.id; Count = $mailboxes }
                break
            }
        } catch { $previous = $null }
        finally { Close-PcmStore $connection }
    }
    $last = if ($previous) { "The last complete snapshot is snapshot $($previous.Id), with $($previous.Count) mailbox(es), and it stays the last complete one. " } else { '' }
    if ($Count -eq 0) {
        return ("Collect found no mailbox in the scope ({0}), so it wrote nothing: an empty snapshot would become the last complete snapshot and Check, Convert and Recover would have nothing to work from. {1}Check Exchange.Server, the scope in the configuration file (Scope.Mode and its SearchBase, GroupDN or CsvPath, and Scope.Include*) and that the account can read the mailboxes, then run Collect again. If the scope really holds no mailbox, set Collect.MaxDropPercent to 100 for this run." -f (Get-PcmScopeText $config), $last)
    }
    if ($previous -and $previous.Count -gt 0) {
        $drop = ($previous.Count - $Count) * 100.0 / $previous.Count
        if ($drop -gt $limit) {
            return ("Collect found {0} mailbox(es), {1:0}% fewer than the last complete snapshot of the same scope (snapshot {2}, {3} mailbox(es)): more than Collect.MaxDropPercent ({4}%), so it wrote nothing and that snapshot stays the last complete one. If mailboxes were really moved or the scope was narrowed on purpose, set Collect.MaxDropPercent to a higher value in the configuration file (100 accepts any size) and run Collect again; otherwise find out why Exchange returned fewer mailboxes." -f $Count, $drop, $previous.Id, $previous.Count, $limit)
        }
    }
    return ''
}

function Invoke-PcmCollect {
    $config = $context.Config
    $collect = $config.Collect
    $dot = [char]0x00B7

    Write-PcmStep -Context $context -Title 'Exchange on-premises connection' -Icon Directory
    $context.CurrentOperation = 'Connect-Exchange'
    $how = Connect-PcmExchange -Context $context
    Write-PcmItem -Context $context -Status Ok -Text "Connected: $how"
    $dc = [string]$config.Exchange.DomainController
    Write-PcmItem -Context $context -Status Info -Text $(if ($dc) { "Domain controller for every read: $dc" } else { 'Domain controller: chosen by Exchange (set Exchange.DomainController to pin one)' })

    Write-PcmStep -Context $context -Title 'Mailboxes in scope' -Icon Target
    $context.CurrentOperation = 'Get-ScopeMailbox'
    $mailboxes = @(Get-PcmScopeMailbox -Context $context -Identity $Identity -Filter $Scope)
    $records = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($mailbox in $mailboxes) {
        $index++
        $context.CurrentIdentity = [string]$mailbox.PrimarySmtpAddress
        $statistics = $null
        if ($collect.MailboxStatistics) {
            try { $statistics = Get-MailboxStatistics -Identity ([string]$mailbox.DistinguishedName) -WarningAction SilentlyContinue -ErrorAction Stop }
            catch { Write-PcmLog -Context $context -Message ("Statistics of {0}: {1}" -f $mailbox.PrimarySmtpAddress, $_.Exception.Message) -Level Detail }
        }
        $recordParameters = @{ Mailbox = $mailbox; Statistics = $statistics }
        if ([string]$mailbox.RecipientTypeDetails -in @('RoomMailbox','EquipmentMailbox') -and
            [string](Get-PcmValue $config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource') {
            $context.CurrentOperation = 'Get-ResourceSettings'
            try {
                $resourceSettings = Get-PcmResourceSettings -Context $context -Mailbox $mailbox
                if (-not $resourceSettings) { throw 'No resource settings were returned.' }
                $recordParameters.ResourceSettings = $resourceSettings
            } catch {
                throw ("Collect stopped before writing the snapshot: resource settings of {0} could not be collected. Booking settings must not be silently omitted or broadened. {1}" -f $mailbox.PrimarySmtpAddress, $_.Exception.Message)
            }
        }
        [void]$records.Add((ConvertTo-PcmMailboxRecord @recordParameters))
    }
    $context.CurrentIdentity = ''
    $byKind = $records | Group-Object { $_.kind }
    $text = if ($records.Count) { ($byKind | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name.ToLowerInvariant() }) -join " $dot " } else { 'none' }
    Write-PcmItem -Context $context -Status $(if ($records.Count) { 'Ok' } else { 'Warn' }) -Icon $(if ($records.Count) { 'People' } else { '' }) -Text ("{0} mailbox(es): {1}" -f $records.Count, $text)
    foreach ($excluded in $context.Excluded) { Write-PcmLog -Context $context -Message ("Excluded: $excluded") -Level Sub }
    $noAddress = @($records | Where-Object { -not $_.primary_smtp_address })
    if ($noAddress.Count) { Write-PcmLog -Context $context -Message ('{0} mailbox(es) without primary SMTP address.' -f $noAddress.Count) -Level Warning }

    $permissions = [System.Collections.Generic.List[object]]::new()
    $members = [System.Collections.Generic.List[object]]::new()
    $permissionText = @{}
    $sharedRecords = @($records | Where-Object { $_.kind -ne 'User' -and -not (
                $_.kind -in @('Room','Equipment') -and [string](Get-PcmValue $config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource') })
    if ($collect.SharedPermissions -and $sharedRecords.Count) {
        Write-PcmStep -Context $context -Title 'Shared mailbox permissions' -Icon Key
        $context.CurrentOperation = 'Get-SharedPermission'
        $permissionMailboxGuids = @{}
        foreach ($record in $sharedRecords) { $permissionMailboxGuids[[string]$record.object_guid] = $true }
        foreach ($mailbox in @($mailboxes | Where-Object { $permissionMailboxGuids.ContainsKey([string]$_.Guid) })) {
            $context.CurrentIdentity = [string]$mailbox.PrimarySmtpAddress
            $result = Get-PcmSharedPermission -Context $context -Mailbox $mailbox
            foreach ($row in $result.Permissions) { [void]$permissions.Add($row) }
            foreach ($row in $result.Members) { [void]$members.Add($row) }
            foreach ($warning in $result.Warnings) { Write-PcmLog -Context $context -Message $warning -Level Warning }
            $mine = @($result.Permissions)
            $counts = @{ FullAccess = 0; SendAs = 0; SendOnBehalf = 0 }
            foreach ($row in $mine) { $counts[$row.access_right]++ }
            $permissionText[[string]$mailbox.Guid] = 'FullAccess {0} {3} SendAs {1} {3} SendOnBehalf {2}' -f $counts.FullAccess, $counts.SendAs, $counts.SendOnBehalf, $dot
            $groups = @($mine | Where-Object { $_.trustee_kind -eq 'Group' }).Count
            Write-PcmLog -Context $context -Message ('{0}: {1}{2}' -f $mailbox.PrimarySmtpAddress, $permissionText[[string]$mailbox.Guid], $(if ($groups) { " $dot $groups group(s) expanded" } else { '' })) -Level Sub
        }
        $context.CurrentIdentity = ''
        $unresolved = @($permissions | Where-Object { -not $_.resolved }).Count
        Write-PcmItem -Context $context -Status $(if ($unresolved) { 'Warn' } else { 'Ok' }) -Text ('{0} permission(s) on {1} shared mailbox(es), {2} group member(s) expanded{3}' -f $permissions.Count, $sharedRecords.Count, $members.Count, $(if ($unresolved) { " $dot $unresolved trustee(s) not resolved" } else { '' }))
    } else {
        Skip-PcmStep -Title 'Shared mailbox permissions' -Reason $(if (-not $collect.SharedPermissions) { 'Collect.SharedPermissions = $false.' } else { 'No shared mailbox in scope.' }) -Icon Key
    }

    $contacts = @(); $groupResult = @{ Groups = @(); Members = @() }
    if (($collect.Contacts -or $collect.DistributionGroups) -and -not $Identity) {
        Write-PcmStep -Context $context -Title 'Contacts and distribution groups' -Icon Mail
        if ($collect.Contacts) {
            $context.CurrentOperation = 'Get-MailContact'
            $contacts = @(Get-PcmContact -Context $context)
            Write-PcmItem -Context $context -Status Ok -Text ('{0} mail contact(s)' -f $contacts.Count)
        }
        if ($collect.DistributionGroups) {
            $context.CurrentOperation = 'Get-DistributionGroup'
            $groupResult = Get-PcmDistributionGroup -Context $context
            Write-PcmItem -Context $context -Status Ok -Text ('{0} distribution group(s), {1} membership(s)' -f @($groupResult.Groups).Count, @($groupResult.Members).Count)
        }
    } else {
        Skip-PcmStep -Title 'Contacts and distribution groups' -Reason $(if ($Identity) { 'One object (-Identity): not collected.' } else { 'Collect.Contacts and Collect.DistributionGroups = $false.' }) -Icon Mail
    }

    $targetGroups = @{ Groups = @(); Members = @(); Warnings = @() }
    $wantedGroups = @(@(Get-PcmValue $collect 'TargetGroups' @()) | ForEach-Object { $_ } | Where-Object { $_ }).Count + $(if ((Test-PcmSameText (Get-PcmValue $config.Scope 'Mode' '') 'Group') -and (Get-PcmValue $config.Scope 'GroupDN' '')) { 1 } else { 0 })
    if ($wantedGroups -and -not $Identity) {
        Write-PcmStep -Context $context -Title 'Target groups' -Icon Directory
        $context.CurrentOperation = 'Get-TargetGroup'
        $targetGroups = Get-PcmTargetGroup -Context $context
        foreach ($warning in @($targetGroups.Warnings)) { Write-PcmLog -Context $context -Message $warning -Level Warning }
        Write-PcmItem -Context $context -Status $(if (@($targetGroups.Groups).Count) { 'Ok' } else { 'Warn' }) -Text ('{0} of {1} target group(s) captured, {2} member(s) in all (nested groups included, computers and contacts too: a run skips what is not a mailbox)' -f @($targetGroups.Groups).Count, $wantedGroups, @($targetGroups.Members).Count)
    } else {
        Skip-PcmStep -Title 'Target groups' -Reason $(if ($Identity) { 'One object (-Identity): not collected.' } else { 'Collect.TargetGroups is empty and Scope.Mode is not Group.' }) -Icon Directory
    }

    foreach ($record in $records) {
        $row = New-PcmRow ([pscustomobject]$record)
        $row.Permissions = [string]$permissionText[[string]$record.object_guid]
        $row.Detail = 'X500 {0} {1} {2} address(es)' -f $(if ($record.legacy_exchange_dn) { 'kept' } else { 'missing' }), $dot, @(($record.email_addresses_json | ConvertFrom-PcmJson) | ForEach-Object { $_ }).Count
        [void]$context.Rows.Add($row)
    }

    $counts = @{ mailbox = $records.Count; permission = $permissions.Count; group_member = $members.Count; contact = $contacts.Count
        distribution_group = @($groupResult.Groups).Count; dl_member = @($groupResult.Members).Count; target_group = @($targetGroups.Groups).Count; target_member = @($targetGroups.Members).Count; excluded = $context.Excluded.Count }
    $scopeJson = @{ Mode = $config.Scope.Mode; SearchBase = $config.Scope.SearchBase; GroupDN = $config.Scope.GroupDN; CsvPath = $config.Scope.CsvPath
        Identity = $Identity; Filter = $Scope; IncludeUsers = $config.Scope.IncludeUsers; IncludeShared = $config.Scope.IncludeShared
        IncludeRoom = $config.Scope.IncludeRoom; IncludeEquipment = $config.Scope.IncludeEquipment } | ConvertTo-Json -Compress
    $sizeProblem = Get-PcmCollectSizeProblem -Count $records.Count -ScopeJson $scopeJson
    if ($context.Mode -ne 'Apply') {
        if ($sizeProblem) { Write-PcmLog -Context $context -Message ('In Apply mode this Collect would be refused: {0}' -f $sizeProblem) -Level Warning }
        Skip-PcmStep -Title 'Saving the snapshot' -Reason 'Preview: not written (add -Mode Apply).' -Icon Backup
        $context.NextSteps = @((Get-PcmCommandLine -ActionName 'Collect' -ModeName 'Apply' -KeepFilters))
        return
    }
    Write-PcmStep -Context $context -Title 'Saving the snapshot' -Icon Backup
    if ($sizeProblem) { $context.NoChangeMade = $true; $context.CurrentOperation = 'Check-SnapshotSize'; throw $sizeProblem }
    if ($Identity) {
        # A one-object snapshot becomes the last complete snapshot: Check and Convert would only see this object.
        if (-not $Force -and -not ($script:ConfirmBound -and (Test-PcmSameText $ConfirmPreference 'None'))) {
            $question = "Write a snapshot of ONE object ($Identity)? It becomes the last snapshot: Check and Convert will only see this object until the next full Collect."
            $context.CurrentOperation = 'Confirm-Apply'
            Write-Host ''
            $answer = $false
            try { $answer = $caller.ShouldContinue($question, 'PRA Cloud Mailbox - Collect') }
            catch { $context.ConfirmationUnavailable = $true; throw "This Collect needs a confirmation and this console cannot ask for one ($($_.Exception.Message)). Add -Force." }
            if (-not $answer) { $context.Cancelled = $true; throw 'Cancelled by the operator: nothing was written.' }
        }
        Write-PcmLog -Context $context -Message 'Snapshot of ONE object (-Identity): Check and Convert use the last complete snapshot, so a full Collect must follow.' -Level Warning
    }
    $context.CurrentOperation = 'Write-Snapshot'
    # The hash file of the snapshot that is there now stays until the new one is written; when this Collect fails after it touched the database, the hash file is written again for
    # the database as it is (the last complete snapshot is intact) provided that it matched the file before.
    $hashedBefore = [IO.File]::Exists($config.Store.Path) -and (Test-PcmSameText (Test-PcmSnapshotHash -Path $config.Store.Path).State 'Verified')
    $connection = Open-PcmStore -Path $config.Store.Path -Root $context.Root
    $backup = ''
    $finished = $false
    try {
        $id = New-PcmSnapshot $connection @{ RunId = $context.RunId; Environment = $config.Environment; Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            ToolVersion = $toolVersion; ExchangeServer = $context.Server; ScopeJson = $scopeJson }
        try {
            $null = Add-PcmRow $connection mailbox $id $records.ToArray()
            $null = Add-PcmRow $connection permission $id $permissions.ToArray()
            $null = Add-PcmRow $connection group_member $id $members.ToArray()
            $null = Add-PcmRow $connection contact $id $contacts
            $null = Add-PcmRow $connection distribution_group $id @($groupResult.Groups)
            $null = Add-PcmRow $connection dl_member $id @($groupResult.Members)
            $null = Add-PcmRow $connection target_group $id @($targetGroups.Groups)
            $null = Add-PcmRow $connection target_member $id @($targetGroups.Members)
            $issues = @(@(Get-PcmValue $context 'WarningList' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ } | ForEach-Object { [ordered]@{ level = 'Warning'; object = ''; operation = [string]$_.Phase; message = [string]$_.Message } })
            $null = Add-PcmRow $connection collect_issue $id $issues
            Set-PcmSnapshotStatus $connection $id Complete $counts
        } catch {
            try { Set-PcmSnapshotStatus $connection $id Failed $counts $_.Exception.Message } catch { $null = $_ }
            throw
        }
        $context.SnapshotLabel = '{0} {1} {2}' -f $id, $dot, (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Write-PcmItem -Context $context -Status Ok -Icon Batch -Text ('Snapshot {0} written to {1}' -f $id, $config.Store.Path)
        $inUse = Get-PcmSnapshotInUse -JournalPath $config.Store.JournalPath -Root $context.Root
        $pruned = 0
        if ($inUse.Readable) {
            $kept = @()
            $pruned = Remove-PcmOldSnapshot $connection $config.Store.KeepSnapshots -Protect $inUse.Items -Kept ([ref]$kept)
            if (@($kept).Count) { Write-PcmItem -Context $context -Status Info -Text ('{0} old snapshot(s) kept although older than Store.KeepSnapshots = {1}: a Convert batch that is not rolled back yet needs them (snapshot {2})' -f @($kept).Count, $config.Store.KeepSnapshots, (@($kept) -join ', ')) }
        } else {
            Write-PcmLog -Context $context -Message ('The journal {0} could not be read ({1}): no old snapshot is deleted in this run, because a batch may still need it.' -f $config.Store.JournalPath, $inUse.Message) -Level Warning
        }
        if ($pruned) { Write-PcmItem -Context $context -Status Info -Text ('{0} old snapshot(s) deleted (Store.KeepSnapshots = {1})' -f $pruned, $config.Store.KeepSnapshots) }
        $backup = Backup-PcmStore $connection $config.Store.BackupFolder ('PraCloudMailbox_{0}_snapshot{1}_{2}.db' -f $config.Environment, $id, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        [void]$context.BackupFiles.Add($backup)
        Write-PcmItem -Context $context -Status Ok -Icon Backup -Text "Copy of the database: $backup"
        $finished = $true
    } finally {
        Close-PcmStore $connection
        if (-not $finished -and $hashedBefore) {
            try { $null = Write-PcmHashFile -Path $config.Store.Path }
            catch { try { Remove-PcmHashFile -Path $config.Store.Path } catch { $null = $_ } }
        }
    }
    $context.CurrentOperation = 'Write-SnapshotHash'
    foreach ($hashed in @($backup, $config.Store.Path) | Where-Object { $_ }) {
        try {
            $null = Write-PcmHashFile -Path $hashed
            Write-PcmItem -Context $context -Status Ok -Icon Backup -Text ('SHA-256 file written: {0}.sha256' -f $hashed)
        } catch {
            try { Remove-PcmHashFile -Path $hashed } catch { $null = $_ }
            Write-PcmLog -Context $context -Message ('The SHA-256 file of {0} could not be written ({1}): the check and the next steps cannot prove that this snapshot file is unchanged.' -f $hashed, $_.Exception.Message) -Level Warning
        }
    }
    foreach ($row in $context.Rows) { $row.FinalStatus = 'Success' }
    $context.NextSteps = @("Copy $($config.Store.Path) to the cloud admin server (or let the backup folder be replicated), then, in the folder of the tool there:", (Get-PcmCommandLine -ActionName 'Check' -OtherServer))
}

# =================================================================================================
# Schedule - registers the daily Collect as a Windows scheduled task, after proving that its account can run it.
# =================================================================================================
function Set-PcmScheduleRow {
    <# One row of the summary card of a Schedule: the rows always come in the same order, whatever the step that sets them. #>
    param([Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Icon, [Parameter(Mandatory)][object]$Value)
    if (-not $context.ContainsKey('ScheduleValues')) { $context.ScheduleValues = @{} }
    $context.ScheduleValues[$Key] = @($Icon, $Value)
    $rows = [ordered]@{}
    foreach ($name in @('Task', 'Account', 'Runs at', 'Command', 'Exchange', 'Batch right', 'Replaces')) { if ($context.ScheduleValues.ContainsKey($name)) { $rows[$name] = $context.ScheduleValues[$name] } }
    $context.SummaryRows = $rows
}

function Invoke-PcmSchedule {
    $config = $context.Config
    $apply = Test-PcmSameText $context.Mode 'Apply'
    $dot = [char]0x00B7
    $setting = $context['ScheduleSetting']
    $context.NoChangeMade = $true
    $commandText = @([string]$setting.Execute, [string]$setting.Argument)
    Set-PcmScheduleRow -Key 'Task' -Icon 'Plan' -Value ('{0} (root folder)' -f $setting.TaskName)
    Set-PcmScheduleRow -Key 'Runs at' -Icon 'Clock' -Value ('every day at {0}' -f $setting.At)
    Set-PcmScheduleRow -Key 'Command' -Icon 'Write' -Value $commandText

    Write-PcmStep -Context $context -Title 'Configuration' -Icon Config
    $context.CurrentOperation = 'Check-Configuration'
    $server = ([string]$config.Exchange.Server).Trim()
    if (-not $server) { throw 'Exchange.Server is empty. The daily Collect reads Exchange through a remote session: set Exchange.Server to the host name of an Exchange server (for example EX01.contoso.com) in the configuration file, then run this again.' }
    Write-PcmItem -Context $context -Status Ok -Text ('Exchange.Server {0}; task "{1}", every day at {2}' -f $server, $setting.TaskName, $setting.At)
    $window = Test-PcmScheduleWindow
    if ($window.Elevated) { Write-PcmItem -Context $context -Status Ok -Text 'This window is elevated: a scheduled task can be registered.' }
    elseif ($apply) { throw 'Registering a scheduled task needs an elevated window. Close this window, open PowerShell 7 with "Run as administrator", and run the same command again. Nothing was changed.' }
    else { Write-PcmItem -Context $context -Status Warn -Text 'This window is not elevated. The Preview goes on, but -Mode Apply needs an elevated window (PowerShell 7, "Run as administrator").' }

    Write-PcmStep -Context $context -Title 'Account' -Icon People
    $context.CurrentOperation = 'Check-Account'
    $credential = $Credential
    if (-not $credential -and $apply) {
        if (-not $window.Interactive) { throw 'No account was given, and this PowerShell was started with -NonInteractive, so it cannot ask for one. Give -Credential, for example: $credential = Get-Credential, then add -Credential $credential to the command.' }
        Write-PcmItem -Context $context -Status Info -Text 'Asking for the account that runs the task (a domain user in View-Only Organization Management) and its password.'
        try { $credential = Get-Credential -Message 'Account that runs the daily Collect: a domain user in the Exchange role group View-Only Organization Management' }
        catch { throw ('The account and its password could not be asked for: {0} Give -Credential instead.' -f $_.Exception.Message.Trim()) }
        if (-not $credential) { $context.Cancelled = $true; throw 'Cancelled by the operator: no account was given, nothing was changed.' }
    }
    $account = $null
    if ($credential) {
        if (-not ([string]$credential.UserName).Trim()) { throw 'The account is empty. Give the account as DOMAIN\name or name@domain.' }
        $account = Get-PcmScheduleAccount -Name ([string]$credential.UserName)
        if (-not $account.Ok) { Set-PcmScheduleRow -Key 'Account' -Icon 'People' -Value ('refused: {0}' -f $account.Typed); throw $account.Message }
        Write-PcmItem -Context $context -Status Ok -Text ('Account {0}: a domain user.' -f $account.Name)
        Set-PcmScheduleRow -Key 'Account' -Icon 'People' -Value ([string]$account.Name)
    } else {
        Write-PcmItem -Context $context -Status Info -Text 'Not tested: no -Credential given. Give -Credential to check the account in the Preview; the Apply asks for it.'
        Set-PcmScheduleRow -Key 'Account' -Icon 'People' -Value 'not given (the Apply asks for it)'
    }

    Write-PcmStep -Context $context -Title 'Exchange access' -Icon Directory
    $context.CurrentOperation = 'Test-ExchangeAccess'
    if ($credential) {
        $proof = Test-PcmExchangeAccess -Server $server -Credential $credential
        if (-not $proof.Ok) { Set-PcmScheduleRow -Key 'Exchange' -Icon 'Directory' -Value 'refused'; throw $proof.Message }
        Write-PcmItem -Context $context -Status Ok -Text $proof.Message
        Set-PcmScheduleRow -Key 'Exchange' -Icon 'Directory' -Value ('accepted by {0} (Kerberos)' -f $server)
    } else {
        Write-PcmItem -Context $context -Status Info -Text 'Not tested: no -Credential given, so no session to Exchange was opened.'
        Set-PcmScheduleRow -Key 'Exchange' -Icon 'Directory' -Value 'not tested (no -Credential)'
    }

    $context.CurrentOperation = 'Test-BatchLogonRight'
    if (-not (Test-PcmDomainController)) {
        Skip-PcmStep -Title 'Batch logon right' -Reason 'This computer is not a domain controller: the account needs no extra right here.' -Icon Key
        Set-PcmScheduleRow -Key 'Batch right' -Icon 'Key' -Value 'not needed (not a domain controller)'
    } else {
        Write-PcmStep -Context $context -Title 'Batch logon right' -Icon Key
        if (-not $account) {
            Write-PcmItem -Context $context -Status Info -Text 'Not tested: no account to check. On a domain controller the account needs the right "Log on as a batch job".'
            Set-PcmScheduleRow -Key 'Batch right' -Icon 'Key' -Value 'not tested (no -Credential)'
        } else {
            $right = Test-PcmBatchLogonRight -Account $account -Folder $context.LogFolder
            switch ($right.State) {
                'Granted' { Write-PcmItem -Context $context -Status Ok -Text $right.Message; Set-PcmScheduleRow -Key 'Batch right' -Icon 'Key' -Value 'granted' }
                'Unknown' { Write-PcmItem -Context $context -Status Warn -Text $right.Message; Set-PcmScheduleRow -Key 'Batch right' -Icon 'Key' -Value 'not known: check it by hand' }
                default { Set-PcmScheduleRow -Key 'Batch right' -Icon 'Key' -Value ([string]$right.State).ToLowerInvariant(); throw $right.Message }
            }
        }
    }

    Write-PcmStep -Context $context -Title 'Scheduled task' -Icon Plan
    $context.CurrentOperation = 'Read-Task'
    $existing = Get-PcmScheduledCollectTask -Setting $setting
    $who = if ($account) { [string]$account.Name } else { 'the account given at the Apply' }
    Write-PcmItem -Context $context -Status Info -Icon Clock -Text ('Task "{0}" in the root folder, every day at {1}, as {2} (Windows keeps its password to get a Kerberos ticket)' -f $setting.TaskName, $setting.At, $who)
    Write-PcmItem -Context $context -Status Info -Icon Write -Text ('Command: "{0}" {1}' -f $setting.Execute, $setting.Argument)
    $replaceText = ''
    if ($existing) {
        $replaceText = 'account {0}, daily at {1}, command {2}' -f $existing.Account, $(if ($existing.At) { $existing.At } else { 'unknown time' }), $existing.Command
        Set-PcmScheduleRow -Key 'Replaces' -Icon 'Info' -Value $replaceText
        Write-PcmItem -Context $context -Status Info -Text ('A task named "{0}" exists and will be replaced. It is now: {1}' -f $existing.Name, $replaceText)
    }
    $row = New-PcmRow ([pscustomobject]@{ user_principal_name = ''; primary_smtp_address = [string]$setting.TaskName; kind = 'ScheduledTask'; object_guid = ''; exchange_guid = '' })
    $row.Detail = 'every day at {0} {1} {2}{3}' -f $setting.At, $dot, $who, $(if ($existing) { ' ' + $dot + ' replaces the task that exists' } else { '' })
    if (-not $apply) {
        Write-PcmItem -Context $context -Status Skip -Text 'Preview: not registered (add -Mode Apply).'
        [void]$context.Rows.Add($row)
        $context.NextSteps = @('To register the task, run the same command in Apply mode (it asks for the account and its password):', (Get-PcmCommandLine -ActionName 'Schedule' -ModeName 'Apply'))
        return
    }
    $question = 'Register the daily Collect as the scheduled task "{0}" on this computer, running as {1} at {2}?{3}' -f $setting.TaskName, $who, $setting.At, $(if ($existing) { ' The task of this name that exists now is replaced.' } else { '' })
    Confirm-PcmApply -Question $question
    if (-not (& $context.Approval ('scheduled task "{0}"' -f $setting.TaskName) 'Register the daily Collect task')) { $context.Cancelled = $true; throw 'Cancelled by the operator: nothing was changed.' }
    $context.CurrentOperation = 'Register-Task'
    $context.NoChangeMade = $false
    Register-PcmCollectTask -Setting $setting -UserName ([string]$account.Name) -Credential $credential
    $context.CurrentOperation = 'Read-Task'
    $registered = Get-PcmScheduledCollectTask -Setting $setting
    if (-not $registered) { throw ('The task "{0}" was registered, but Windows does not list it. Open the Task Scheduler and look for it.' -f $setting.TaskName) }
    Write-PcmItem -Context $context -Status Ok -Text ('Task "{0}" {1}: every day at {2}, as {3}.' -f $registered.Name, $(if ($existing) { 'replaced' } else { 'registered' }), $setting.At, $account.Name)
    $row.FinalStatus = 'Success'
    [void]$context.Rows.Add($row)
    $context.NextSteps = @('Run the task once now, in this elevated window:', ("Start-ScheduledTask -TaskName '{0}'" -f $setting.TaskName.Replace("'", "''")),
        'When it has finished, check the snapshot, then look at the light Daily Collect on the Setup page of the console:', (Get-PcmCommandLine -ActionName 'Check'))
}

# =================================================================================================
# Snapshot of the Collect (Check, Convert, Recover).
# =================================================================================================
function Confirm-PcmSnapshotHash {
    <#
    Before the snapshot file is read: compares it with the .sha256 that Collect wrote beside it. A file changed after its Collect stops the run with nothing changed (unless
    -AcceptUnverifiedSnapshot); a file without a .sha256 stops a Convert or a Recover in Apply mode in the same way, and only warns in a Preview, a Check or a Collect. The journal is not covered: Convert and Recover write it themselves.
    #>
    $config = $context.Config
    $path = $config.Store.Path
    if (-not [IO.File]::Exists($path)) { return }
    $state = Test-PcmSnapshotHash -Path $path -BackupFolder $config.Store.BackupFolder
    $accept = [bool](Get-PcmValue $context 'AcceptUnverifiedSnapshot' $false)
    switch ($state.State) {
        'Verified' {
            $context.SnapshotCheck = if ($accept) { 'skipped on request (-AcceptUnverifiedSnapshot); the file matches its .sha256' } else { 'verified: the file matches its .sha256 written by Collect' }
            Write-PcmItem -Context $context -Status Ok -Text 'Snapshot file verified: it matches the .sha256 written by Collect.'
        }
        'NotVerifiable' {
            if (-not $accept) { $context.SnapshotCheck = 'not verifiable: no .sha256 beside the file' }
            if (-not $accept -and (Test-PcmSameText $context.Mode 'Apply') -and ((Test-PcmSameText $context.Action 'Convert') -or (Test-PcmSameText $context.Action 'Recover'))) {
                $context.NoChangeMade = $true; $context.CurrentIdentity = ''
                throw ($state.Message + ' Nothing was changed. If you know that this file is the one Collect wrote, run the same command again with -AcceptUnverifiedSnapshot.')
            }
            Write-PcmLog -Context $context -Message $state.Message -Level Warning
        }
        default {
            if (-not $accept) {
                $context.SnapshotCheck = 'failed: the file does not match its .sha256'
                $context.NoChangeMade = $true; $context.CurrentIdentity = ''
                throw $state.Message
            }
            $context.SnapshotCheck = 'skipped on request (-AcceptUnverifiedSnapshot); the file does not match its .sha256'
            Write-PcmLog -Context $context -Message $(if ($state.Expected) { 'Snapshot check skipped on request (-AcceptUnverifiedSnapshot): the file does not match its .sha256.' } else { 'Snapshot check skipped on request (-AcceptUnverifiedSnapshot): the .sha256 file holds no hash.' }) -Level Warning
        }
    }
}

function Read-PcmSnapshotData {
    <#
    Loads one snapshot (default: the last complete one), resolves -Group, applies the source of the run (-Identity, -CsvPath, -SearchBase, -Group), -Scope and -MaxObjects
    (or keeps -ObjectGuid only), prints the step. -Known gives the reason why an object of the batch is not left to do; -SkipConverted (a new Convert) leaves out the
    objects that an earlier batch converted.
    #>
    param([long]$Id = 0, [string]$Verb = 'check', [switch]$All, [AllowEmptyCollection()][string[]]$ObjectGuid, [switch]$NoCount, [string]$Where = '', [hashtable]$Known = @{}, [switch]$SkipConverted, [AllowNull()][object]$Batch = $null)
    $config = $context.Config
    $dot = [char]0x00B7
    Write-PcmStep -Context $context -Title 'Collect snapshot' -Icon Batch
    $context.CurrentOperation = 'Read-Snapshot'
    Confirm-PcmSnapshotHash
    $batchRun = [string](Get-PcmValue $Batch 'snapshot_run_id' '')
    $batchHash = [string](Get-PcmValue $Batch 'snapshot_hash' '')
    $connection = Open-PcmStore -Path $config.Store.Path -Root $context.Root -ReadOnly
    try {
        $snap = if ($batchRun) { Get-PcmSnapshot $connection -RunId $batchRun } else { Get-PcmSnapshot $connection $Id }
        if (-not $snap) {
            $context.NoChangeMade = $true; $context.CurrentIdentity = ''
            $copies = "Every Collect also writes a copy of the database in Store.BackupFolder ($($config.Store.BackupFolder)), named PraCloudMailbox_<Environment>_snapshot<number>_<date>.db and kept as it is: copy the file of the snapshot you need over $($config.Store.Path) (with its .sha256 file) and run the same command again."
            throw $(if ($Batch) { "Batch $(Get-PcmValue $Batch 'id' '') was made from snapshot $Id$(if ($batchRun) { " (Collect run $batchRun)" }), which is not in $($config.Store.Path): Collect keeps only the last Store.KeepSnapshots ($($config.Store.KeepSnapshots)) snapshots, or this database is another one. Nothing was changed. $copies" }
                elseif ($Id) { "Snapshot $Id not found in $($config.Store.Path) (Collect keeps the last Store.KeepSnapshots = $($config.Store.KeepSnapshots) snapshots). $copies" }
                else { "No complete snapshot in $($config.Store.Path): run Collect -Mode Apply (it reads Exchange.Server) or copy the snapshot database here." })
        }
        if ($snap.status -ne 'Complete') { throw "Snapshot $($snap.id) is $($snap.status), not Complete." }
        $records = @(Read-PcmTable $connection mailbox $snap.id)
        $permissions = @(Read-PcmTable $connection permission $snap.id)
        $members = @(Read-PcmTable $connection group_member $snap.id)
        $found = $null; $capturedMembers = @()
        if ($Group) {
            $found = Find-PcmTargetGroup $connection $snap.id $Group
            if ($found.Groups.Count -eq 1) {
                $capturedMembers = @(Invoke-PcmSql $connection 'SELECT * FROM target_member WHERE snapshot_id = @id AND group_guid = @g COLLATE NOCASE ORDER BY rowid' @{ id = $snap.id; g = [string]$found.Groups[0].group_guid } -As Rows)
            }
        }
    } finally { Close-PcmStore $connection }
    $context.SnapshotId = [string]$snap.id
    $context.SnapshotRunId = [string]$snap.run_id
    $context.SnapshotHash = Get-PcmSnapshotDigest -Mailbox $records -Permission $permissions -Member $members
    if ($Batch -and $batchHash -and -not (Test-PcmSameText $batchHash $context.SnapshotHash)) {
        if (-not [bool](Get-PcmValue $context 'AcceptUnverifiedSnapshot' $false)) {
            $context.NoChangeMade = $true; $context.CurrentIdentity = ''
            throw ("Snapshot {0} of {1} is not the snapshot that batch {2} was made from: the mailboxes, permissions or group members it holds are not the ones that Convert read (digest {3}, the journal says {4}). Nothing was changed. Restore the right database from Store.BackupFolder ({5}), or, if you know why it differs and accept it, run the same command again with -AcceptUnverifiedSnapshot." -f $snap.id, $config.Store.Path, (Get-PcmValue $Batch 'id' ''), $context.SnapshotHash, $batchHash, $config.Store.BackupFolder)
        }
        Write-PcmLog -Context $context -Message ('Snapshot {0} is not the snapshot that batch {1} was made from; the run goes on because of -AcceptUnverifiedSnapshot.' -f $snap.id, (Get-PcmValue $Batch 'id' '')) -Level Warning
    } elseif ($Batch -and -not $batchHash) {
        Write-PcmLog -Context $context -Message ('Batch {0} was made before batches kept the content of their snapshot: the tool cannot prove that snapshot {1} is the one it was made from.' -f (Get-PcmValue $Batch 'id' ''), $snap.id) -Level Warning
    }
    $taken = [datetime]::Parse([string]$snap.finished_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    $age = (Get-Date).ToUniversalTime() - $taken.ToUniversalTime()
    $context.SnapshotLabel = '{0} {1} {2:yyyy-MM-dd HH:mm} UTC {1} {3}' -f $snap.id, $dot, $taken.ToUniversalTime(), $snap.exchange_server
    $context.Server = [string]$snap.exchange_server
    $scopeInfo = $null; try { $scopeInfo = [string]$snap.scope_json | ConvertFrom-PcmJson } catch { $scopeInfo = $null }
    $partial = $scopeInfo -and [string](Get-PcmValue $scopeInfo 'Identity' '')
    Write-PcmItem -Context $context -Status Ok -Text ('Snapshot {0} of {1:yyyy-MM-dd HH:mm} UTC from {2} ({3} mailbox(es), {4} permission(s))' -f $snap.id, $taken.ToUniversalTime(), $snap.exchange_server, $records.Count, $permissions.Count)
    if ($partial) { Write-PcmLog -Context $context -Message ('Snapshot {0} holds one object only ({1}): not a full Collect.' -f $snap.id, $scopeInfo.Identity) -Level Warning }
    if ($age.TotalDays -gt $config.Store.MaxSnapshotAgeDays) { Write-PcmLog -Context $context -Message ('Snapshot {0:0} day(s) old (Store.MaxSnapshotAgeDays = {1}): permissions and addresses may have changed since.' -f $age.TotalDays, $config.Store.MaxSnapshotAgeDays) -Level Warning }
    if ($Group) { $context.GroupTarget = Resolve-PcmGroupTarget -Snapshot $snap -Found $found -CapturedMembers $capturedMembers -Records $records }
    # -ObjectGuid keeps the objects of a batch still to resume; the source of the run, -Scope and -MaxObjects then narrow them like they narrow a snapshot.
    $allRecords = @($records)
    $inBatch = $PSBoundParameters.ContainsKey('ObjectGuid')
    $knownRecords = @()
    if ($inBatch) {
        $wanted = [System.Collections.Generic.HashSet[string]]::new([string[]]@($ObjectGuid), [StringComparer]::OrdinalIgnoreCase)
        $knownRecords = @($records | Where-Object { -not $wanted.Contains([string]$_.object_guid) -and $Known.ContainsKey([string]$_.object_guid) })
        $records = @($records | Where-Object { $wanted.Contains([string]$_.object_guid) })
    }
    if (-not $All) {
        $records = @(Select-PcmTarget -Items $records -Where $(if ($Where) { $Where } else { "snapshot $($snap.id)" }) -AllRecords $allRecords -Known $Known -KnownItems $knownRecords -InBatch:$inBatch -SkipConverted:$SkipConverted `
                -RecordOf { param($r) $r } -Keys { param($r) @([string]$r.user_principal_name, [string]$r.primary_smtp_address, [string]$r.sam_account_name, [string]$r.object_guid, [string]$r.distinguished_name) })
        $records = @(Limit-PcmTarget -Items $records -SkipConverted:$SkipConverted)
        Write-PcmTargetHeader -Count $records.Count
    }
    $users = @($records | Where-Object { $_.kind -eq 'User' }).Count
    if (-not $NoCount) { Write-PcmItem -Context $context -Status Info -Icon Target -Text ('{0} object(s) to {3}: {1} user(s), {2} shared/room/equipment' -f $records.Count, $users, ($records.Count - $users), $Verb) }
    return @{ Snapshot = $snap; Records = $records; Permissions = $permissions; Members = $members; Users = $users }
}

# =================================================================================================
# Check - cloud admin server, read-only (PowerShell 7).
# =================================================================================================
function Invoke-PcmCheck {
    $config = $context.Config

    $data = Read-PcmSnapshotData -Id $Snapshot -Verb 'check'
    $records = @($data.Records); $permissions = @($data.Permissions); $members = @($data.Members)
    $context.NativeResourceScope = @($records | Where-Object { Test-PcmNativeResource -Config $config -Record $_ }).Count -gt 0
    $nativeOnly = $records.Count -gt 0 -and @($records | Where-Object { -not (Test-PcmNativeResource -Config $config -Record $_) }).Count -eq 0

    Write-PcmStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    $how = Connect-PcmCloudSession
    Write-PcmItem -Context $context -Status Ok -Text $how
    if ($context.ContainsKey('CertificateWarning')) { Write-PcmLog -Context $context -Message $context['CertificateWarning'] -Level Warning }

    Write-PcmStep -Context $context -Title 'Tenant prerequisites' -Icon Key
    $context.CurrentOperation = 'Tenant-Facts'
    $skipUserGroup = @($records | Where-Object { $_.kind -eq 'User' }).Count -eq 0 -and @($records | Where-Object { Test-PcmNativeResource -Config $config -Record $_ }).Count -gt 0
    $tenant = Get-PcmTenantFact -Context $context -SkipUserGroup:$skipUserGroup -NativeOnly:$nativeOnly
    $roles = @(Get-PcmGraphRole)
    # The state of every object first: the licence units really needed depend on what each one already holds.
    if (Test-PcmStop) { throw 'Check stopped by the operator before reading the cloud state.' }
    $states = Read-PcmCloudStates -Records $records
    $need = Get-PcmLicenceNeed -Entries @($records | ForEach-Object { $s = $states[[string]$_.object_guid]; [pscustomobject]@{ Record = $_; State = $(if ($s -is [System.Management.Automation.ErrorRecord]) { $null } else { $s }); Step = '' } }) -Tenant $tenant
    $tenantFindings = @(Test-PcmTenantReadiness -Config $config -Roles $roles -Tenant $tenant -UserCount $need.Users -SharedCount $need.Shared -ResourceCount $need.Resources -SkipUserLicensing:$need.SkipUserLicensing -NativeOnly:$nativeOnly -SignIn $config.Cloud.SignIn)
    $tenantFindings += @(Get-PcmResourceLicenceFinding -Need $need -Tenant $tenant)
    if (Test-PcmSameText $config.Cloud.SignIn 'Interactive') {
        $tenantFindings += @(Test-PcmAdminRole -Action Convert -Roles (Get-PcmAdminRole))
        $tenantFindings += [pscustomobject]@{ Level = 'Info'; Code = 'ADMIN_ROLE'; Message = 'Not readable from Microsoft Graph, to check by hand for Convert and Recover: an Exchange Online role that changes recipients and mailbox permissions (Exchange Administrator), and, for the case hold of Recover, membership of the eDiscovery Manager role group of Microsoft Purview (administrator of the case).' }
    }
    foreach ($finding in $tenantFindings) {
        switch ($finding.Level) {
            'Error' { Write-PcmLog -Context $context -Message $finding.Message -Level Error }
            'Warn' { Write-PcmLog -Context $context -Message $finding.Message -Level Warning }
            'Ok' { Write-PcmItem -Context $context -Status Ok -Text $finding.Message }
            default { Write-PcmItem -Context $context -Status Info -Text $finding.Message }
        }
    }
    if (@($records | Where-Object { $_.kind -ne 'User' }).Count) {
        $waveSource = if ([int](Get-PcmValue (Get-PcmValue $config 'Recover' @{}) 'WaveSize' 0) -ge 1) { 'Recover.WaveSize' } else { 'Recover.WaveSize not set: same as Licensing.Shared.Parallel' }
        Write-PcmItem -Context $context -Status Info -Text ('Recover rolls back the shared mailboxes in waves of {0} ({1}); it uses no licence unit.' -f (Get-PcmRecoverWaveSize $config), $waveSource)
    }
    Write-PcmItem -Context $context -Status Info -Text ("Recover needs Entra Connect {0} or later on the rebuilt server (it keeps the cloud source of authority of the objects); this version is not readable from the cloud." -f $config.EntraConnect.MinVersion)

    Write-PcmStep -Context $context -Title 'Objects' -Icon People
    $context.CurrentOperation = 'Object-Readiness'
    $lookup = New-PcmTrusteeLookup -States $states -Records $records
    $permissionIndex = Get-PcmPermissionIndex -Permissions $permissions -Members $members
    $ready = 0
    foreach ($record in $records) {
        $row = New-PcmRow $record
        $context.CurrentIdentity = $row.Identity
        try {
            $state = $states[[string]$record.object_guid]
            if ($state -is [System.Management.Automation.ErrorRecord]) { throw $state.Exception }
            $findings = @(Get-PcmReadiness -Config $config -Record $record -State $state -Skus @($tenant.Skus))
            if (Get-PcmValue $state 'NativeResource' $false) {
                try { Test-PcmResourceConfigurationReadiness -Record $record -AllowRoomListCreation -RoomListPlan @((Read-PcmResourceSettings $record).RoomLists) }
                catch { $findings += [pscustomobject]@{ Level = 'Error'; Field = 'ExchangeOnline'; Code = 'RESOURCE_CONFIGURATION_REQUIRED'; Message = $_.Exception.Message } }
            } elseif ($record.kind -ne 'User') {
                $subset = Get-PcmPermissionSubset -Index $permissionIndex -MailboxGuid ([string]$record.object_guid)
                $findings += @(Get-PcmTrusteeFinding -Record $record -Permissions $subset.Permissions -Members $subset.Members -Lookup $lookup)
            }
        } catch {
            $findings = @([pscustomobject]@{ Level = 'Error'; Field = 'Entra'; Code = 'CHECK_FAILED'; Message = "Check failed: $($_.Exception.Message)" })
        }
        foreach ($field in @('Entra','ExchangeOnline','TeamsStorage','Licence','Holds','Permissions')) {
            $row[$field] = (@($findings | Where-Object { $_.Field -eq $field -and $_.Level -ne 'Info' } | ForEach-Object { $_.Code }) +
                @($findings | Where-Object { $_.Field -eq $field -and $_.Level -eq 'Info' } | ForEach-Object { $_.Code })) -join ', '
        }
        $errors = @($findings | Where-Object { $_.Level -eq 'Error' })
        $warns = @($findings | Where-Object { $_.Level -eq 'Warn' })
        $row.Detail = (@($findings | Where-Object { $_.Level -in @('Error','Info') -or $_.Field -eq 'Permissions' } | ForEach-Object { $_.Message })) -join ' | '
        $row.Warnings = (@($warns | ForEach-Object { $_.Message })) -join ' | '
        $row.FinalStatus = if ($errors.Count) { 'Error' } else { 'Success' }
        if (-not $errors.Count) { $ready++ }
        $status = if ($errors.Count) { 'Fail' } elseif ($warns.Count) { 'Warn' } else { 'Ok' }
        $kind = if ($record.kind -eq 'User') { 'user' } else { [string]$record.kind.ToLowerInvariant() }
        $main = if ($errors.Count) { $errors[0] } elseif ($warns.Count) { $warns[0] } else { $findings | Where-Object { $_.Level -eq 'Info' } | Select-Object -First 1 }
        $summary = if ($main) { $main.Message } else { 'ready' }
        Write-PcmItem -Context $context -Status $(if ($status -eq 'Warn') { 'Warn' } elseif ($status -eq 'Fail') { 'Fail' } else { 'Ok' }) -Text ('{0} ({1}): {2}' -f $row.Identity, $kind, $summary)
        foreach ($finding in @($findings | Where-Object { $_.Level -ne 'Ok' -and -not [object]::ReferenceEquals($_, $main) })) { Write-PcmLog -Context $context -Message ('{0}: {1}' -f $finding.Code, $finding.Message) -Level Sub }
        [void]$context.Rows.Add($row)
    }
    $context.CurrentIdentity = ''
    Write-PcmItem -Context $context -Status $(if ($ready -eq $records.Count) { 'Ok' } else { 'Info' }) -Icon Target -Text ('{0} of {1} object(s) ready for Convert' -f $ready, $records.Count)
    $context.CheckFinished = $true
    $context.NextSteps = @('Disaster: run the Preview of Convert first (nothing is changed), then the same command with -Mode Apply:', (Get-PcmCommandLine -ActionName 'Convert' -KeepFilters))
}

function Invoke-PcmResolveGroup {
    <#
    Check -Group <name> -ResolveOnly: which group the name means (the snapshot first, then Microsoft Entra ID, exactly like a real run) and how many members it has, with the
    members that a run would skip. Nothing is changed and no report is written. A group that is not found, or a name that matches several groups, ends the run with exit code 1.
    #>
    $null = Read-PcmSnapshotData -Id $Snapshot -Verb 'check'
    $context.CurrentOperation = 'Resolve-Group'
    Write-PcmItem -Context $context -Status Ok -Icon Target -Text ('Group {0} found. Nothing was changed.' -f $context.GroupTarget.Name)
}

# =================================================================================================
# Check -Batch - follow-up of a batch, read-only (PowerShell 7).
# =================================================================================================
function Get-PcmJsonField {
    <# One field of the result_json of a journal row ('' when there is none). #>
    param([AllowNull()][object]$Item, [Parameter(Mandatory)][string]$Name)
    if (-not $Item) { return $null }
    $text = [string](Get-PcmValue $Item 'result_json' '')
    if (-not $text) { return $null }
    try { $value = $text | ConvertFrom-PcmJson } catch { return $null }
    if ($null -eq $value -or -not $value.PSObject.Properties[$Name]) { return $null }
    return $value.$Name
}

function Invoke-PcmCheckBatch {
    <#
    Check -Batch <id> [-Expect Provisioned|Deprovisioned|Retained] [-Once]: are the objects of the batch in the state that Convert or Recover was meant to give? The
    verifications are the ones Convert and Recover wait for (Test-PcmBatchExpectation, the read-back of the permissions, the case hold, the inactive mailbox). Without -Once the
    check waits, with the Polling settings, until every object is in the expected state or the time is over. It writes nothing to the tenant.
    #>
    $config = $context.Config
    $dot = [char]0x00B7
    Write-PcmStep -Context $context -Title 'Batch journal' -Icon Batch
    $context.CurrentOperation = 'Read-Batch'
    $context.Journal = Open-PcmStore -Path $config.Store.JournalPath -Root $context.Root -Kind Journal -ReadOnly
    $given = Get-PcmBatch $context.Journal -Id $Batch
    if (-not $given) { throw "Batch $Batch not found in $($config.Store.JournalPath)." }
    $isRecover = Test-PcmSameText $given.action 'Recover'
    $convert = if ($isRecover) { Get-PcmBatch $context.Journal -Id ([string]$given.convert_batch) } else { $given }
    if (-not $convert) { throw ('The Convert batch {0} of Recover batch {1} is not in the journal.' -f $given.convert_batch, $given.id) }
    $expect = if ($Expect) { $Expect } elseif ($isRecover) { 'Deprovisioned' } else { 'Provisioned' }
    if ($isRecover -and $expect -eq 'Provisioned') { throw ('Batch {0} is a Recover batch: its objects are rolled back, not converted. Use -Expect Deprovisioned or -Expect Retained.' -f $given.id) }
    $context.BatchId = [string]$given.id
    $context.ConvertBatchId = [string]$convert.id
    $history = Get-PcmBatchHistory $context.Journal ([string]$convert.id)
    $actionOf = @{}; foreach ($batchRow in @($history.Batches)) { $actionOf[[string]$batchRow.id] = [string]$batchRow.action }
    $recoverByGuid = @{}
    foreach ($row in @($history.Items | Where-Object { $actionOf[[string]$_.batch_id] -eq 'Recover' })) { $recoverByGuid[[string]$row.object_guid] = $row }
    # Every object of the batch is followed, also the ones that Convert never started or that failed at their first step.
    $items = @($history.Items | Where-Object { [string]$_.batch_id -eq [string]$convert.id -and $_.status -ne 'Skipped' })
    $hasRecover = @($history.Batches | Where-Object { Test-PcmSameText ([string]$_.action) 'Recover' }).Count -gt 0
    if ($isRecover) {
        $wanted = @(@($history.Items | Where-Object { [string]$_.batch_id -eq [string]$given.id }) | ForEach-Object { [string]$_.object_guid })
        $items = @($items | Where-Object { [string]$_.object_guid -in $wanted })
    }
    Write-PcmItem -Context $context -Status Info -Icon Batch -Text ('{0} batch {1} of {2}: {3} object(s) in the journal, state expected: {4}{5}' -f $given.action, $given.id, (Format-PcmUtcText $given.created_utc), $items.Count, $expect, $(if ($Once) { ' (one look)' } else { '' }))
    $data = Read-PcmSnapshotData -Id ([long]$convert.snapshot_id) -All -NoCount -Batch $convert
    $byGuid = @{}; foreach ($record in $data.Records) { $byGuid[[string]$record.object_guid] = $record }
    $items = @(Select-PcmTarget -Items $items -Where "the objects of batch $($convert.id)" -AllRecords @($data.Records) -InBatch -RecordOf { param($t) $byGuid[[string]$t.object_guid] } -Keys {
            param($t) $r = $byGuid[[string]$t.object_guid]
            @([string]$t.identity, [string]$t.object_guid) + $(if ($r) { @([string]$r.primary_smtp_address, [string]$r.user_principal_name, [string]$r.sam_account_name) } else { @() }) })
    $items = @(Limit-PcmTarget -Items $items)
    Write-PcmTargetHeader -Count $items.Count
    if (-not $items.Count) { throw ('Batch {0} holds no object to follow with these options. Nothing to check.' -f $given.id) }
    $nativeItems = @($items | Where-Object { Get-PcmResourceIntent -Item $_ })
    $legacyItems = @($items | Where-Object { -not (Get-PcmResourceIntent -Item $_) })
    $context.NativeResourceScope = $nativeItems.Count -gt 0
    $context.ResourceJournal = @{}

    Write-PcmStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    $how = Connect-PcmCloudSession
    Write-PcmItem -Context $context -Status Ok -Text $how
    $holdSet = $null
    if ($expect -eq 'Retained' -and $config.Retention.HoldPolicy -and @($legacyItems | Where-Object { Test-PcmSameText $_.kind 'User' }).Count) {
        try { Connect-PcmCompliance -Context $context; $holdSet = Get-PcmHoldPolicySet -Policy $config.Retention.HoldPolicy -Limit $config.Retention.HoldPolicyLimit }
        catch { Write-PcmLog -Context $context -Message ('The case hold policies could not be read ({0}): any case hold of Microsoft Purview on the mailbox is taken for the hold of the tool.' -f $_.Exception.Message) -Level Warning }
    }

    Write-PcmStep -Context $context -Title 'Follow-up of the batch' -Icon Target
    $context.CurrentOperation = 'Follow-Batch'
    $skus = @((Invoke-PcmGraph -Uri 'v1.0/subscribedSkus').value)
    $mailboxPlans = @(Get-PcmMailboxPlanId -Skus $skus)
    $legacyRecords = @($legacyItems | ForEach-Object { $byGuid[[string]$_.object_guid] } | Where-Object { $_ })
    $permissionIndex = Get-PcmPermissionIndex -Permissions @($data.Permissions) -Members @($data.Members)
    $liveLimit = 200
    $sharedCount = @($legacyItems | Where-Object { -not (Test-PcmSameText $_.kind 'User') }).Count
    if ($expect -eq 'Provisioned' -and $sharedCount -gt $liveLimit) {
        Write-PcmItem -Context $context -Status Info -Text ('{0} shared mailboxes: their permissions are checked from the journal, not one by one in Exchange Online (the live read-back is done up to {1} shared mailboxes).' -f $sharedCount, $liveLimit)
    }
    $evaluate = {
        $states = if ($legacyRecords.Count) { Get-PcmCloudStateSet -Context $context -Records $legacyRecords } else { @{} }
        $results = [ordered]@{}
        foreach ($item in $items) {
            $guid = [string]$item.object_guid
            $record = $byGuid[$guid]
            $intent = Get-PcmResourceIntent -Item $item
            $facts = @{}
            $state = $null
            if (-not [string](Get-PcmValue $item 'step' '') -and -not $intent) {
                $seen = if ($record) { $states[$guid] } else { $null }
                $stage = if ($seen -and $seen -isnot [System.Management.Automation.ErrorRecord]) { [string](Get-PcmValue $seen 'RecipientType' '') } else { '' }
                $why = [string](Get-PcmValue $item 'message' '')
                $text = if ($expect -eq 'Provisioned') { 'not converted: Convert never started this object, nothing was changed on it' }
                    else { 'not rolled back: Convert never started this object, so Recover has nothing to roll back for it' }
                $results[$guid] = [pscustomobject]@{ Met = $false; Stage = $stage; Detail = ($text + $(if ($why) { " (journal: $why)" } else { '' })) }
                continue
            }
            if ($intent) {
                $facts.Native = $true
                $facts.NativeMode = Get-PcmResourceProtectionMode $intent
                $facts.CloudType = [string]$intent.cloudType
                $cloud = if ([string]$intent.cloudId) { Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id" -f [string]$intent.cloudId) -AllowNotFound } else { $null }
                $facts.NativeExists = [bool]$cloud
                $facts.NativeType = if ($cloud) { [string](Get-PcmExoStateSet -Id @([string]$intent.cloudId))[[string]$intent.cloudId].Type } else { '' }
                if ($expect -eq 'Retained' -and $facts.NativeMode -ne 'None' -and [string]$intent.cloudGuid) { $facts.InactiveFound = [bool](Get-Mailbox -InactiveMailboxOnly -Identity ([string]$intent.cloudGuid) -ErrorAction SilentlyContinue) }
            } else {
                $state = if ($record) { $states[$guid] } else { $null }
                $user = if ($state -isnot [System.Management.Automation.ErrorRecord] -and $state) { Get-PcmValue $state 'User' $null } else { $null }
                if ($user) {
                    $enabled = @($user.assignedPlans | Where-Object { $_ -and [string]$_.capabilityStatus -eq 'Enabled' -and [string]$_.servicePlanId -in $mailboxPlans })
                    $facts.HasPlan = [bool]$enabled.Count
                }
                if ($expect -eq 'Provisioned' -and -not (Test-PcmSameText $item.kind 'User')) {
                    $failed = Get-PcmJsonField $item 'failed'
                    $facts.PermissionsFailed = @($failed | ForEach-Object { $_ } | Where-Object { $_ })
                    if ($user -and $sharedCount -le $liveLimit -and (Get-PcmValue $state 'RecipientType' '') -eq 'SharedMailbox') {
                        $subset = Get-PcmPermissionSubset -Index $permissionIndex -MailboxGuid $guid
                        $missing = [System.Collections.Generic.List[string]]::new()
                        foreach ($grant in @(Get-PcmSharedGrant -Record $record -Permissions $subset.Permissions -Members $subset.Members)) {
                            if (-not (Test-PcmSharedGrant -Mailbox ([string]$user.id) -Grant $grant)) { $missing.Add(('{0} {1}' -f $grant.Right, $grant.Trustee)) }
                        }
                        $facts.PermissionsMissing = $missing.ToArray()
                    }
                }
                if ($expect -eq 'Retained') {
                    if (Test-PcmSameText $item.kind 'User') {
                        $holds = @(if ($state -isnot [System.Management.Automation.ErrorRecord] -and $state) { Get-PcmValue $state 'Holds' @() | ForEach-Object { $_ } })
                        $facts.HoldStamped = if ($holdSet) { Test-PcmToolHold -Set $holdSet -Holds @($holds) } else { [bool]@($holds | Where-Object { ([string]$_).StartsWith('UniH', [StringComparison]::OrdinalIgnoreCase) }).Count }
                    } else {
                        $inactive = Get-PcmJsonField $recoverByGuid[$guid] 'inactiveGuid'
                        if ($inactive) { $facts.InactiveFound = [bool](Get-Mailbox -InactiveMailboxOnly -Identity ([string]$inactive) -ErrorAction SilentlyContinue) }
                    }
                }
            }
            $results[$guid] = Test-PcmBatchExpectation -Expect $expect -Item $item -Record $record -State $state -Config $config -Facts $facts
        }
        return $results
    }
    $timeout = switch ($expect) { 'Provisioned' { Get-PcmWaitSeconds 'MailboxTimeoutMinutes' } 'Deprovisioned' { Get-PcmWaitSeconds 'SyncTimeoutMinutes' } default { Get-PcmWaitSeconds 'HoldTimeoutMinutes' } }
    $watch = @{ Results = $null; Passes = 0 }
    $poll = {
        $watch.Results = & $evaluate
        $watch.Passes++
        (@($watch.Results.Values | Where-Object { -not $_.Met }).Count -eq 0) -or (Test-PcmStop)
    }
    if ($Once) { $null = & $poll }
    else {
        Write-PcmItem -Context $context -Status Info -Icon Clock -Text ('Waiting until every object is {0}, for {1} at most (the Polling settings of the configuration); -Once looks one time only.' -f $expect.ToLowerInvariant(), (Format-PcmDuration $timeout))
        $null = Wait-PcmCondition -Phase 'Waiting for the objects to reach the expected state' -TimeoutSeconds $timeout -IntervalSeconds ([int]$config.Polling.IntervalSeconds) -Test $poll
    }
    $followUp = [System.Collections.Generic.List[object]]::new()
    $met = 0
    foreach ($item in $items) {
        $guid = [string]$item.object_guid
        $record = $byGuid[$guid]
        $row = New-PcmRow $(if ($record) { $record } else { [pscustomobject]@{ user_principal_name = $item.identity; kind = $item.kind; primary_smtp_address = ''; object_guid = $item.object_guid; exchange_guid = '' } })
        $row.Action = 'Check'
        [void]$context.Rows.Add($row)
        $verdict = $watch.Results[$guid]
        $row.ExchangeOnline = [string]$verdict.Stage
        $detail = 'expected {0}: {1}' -f $expect, $verdict.Detail
        if ($verdict.Met) { $met++; Set-PcmRowResult -Row $row -Status Success -Detail $detail }
        else { Set-PcmRowResult -Row $row -Status Error -Detail $detail }
        $followUp.Add([pscustomobject]@{ Identity = $row.Identity; Expect = $expect; Met = [bool]$verdict.Met; Detail = [string]$verdict.Detail })
    }
    $context.FollowUp = $followUp.ToArray()
    Write-PcmItem -Context $context -Status $(if ($met -eq $items.Count) { 'Ok' } else { 'Warn' }) -Icon Target -Text ('{0} of {1} object(s) {2}{3}' -f $met, $items.Count, $expect.ToLowerInvariant(), $(if ($watch.Passes -gt 1) { " $dot $($watch.Passes) look(s)" } else { '' }))
    $context.CheckFinished = $true
    $context.NextOnFailure = $true
    $next = [System.Collections.Generic.List[string]]::new()
    $batchOpen = $expect -eq 'Provisioned' -and -not $isRecover -and $met -eq $items.Count -and -not (Test-PcmSameText ([string]$convert.status) 'Complete')
    if ($batchOpen) {
        Write-PcmLog -Context $context -Message ('The {0} object(s) followed are provisioned, but Convert batch {1} is {2}, not Complete: it still has work to resume (a room list, or an object that is not in the journal as done).' -f $items.Count, $convert.id, ([string]$convert.status).ToLowerInvariant()) -Level Error
        $next.Add('Resume the batch first (nothing is converted twice), then look again:')
        $next.Add((Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -BatchId ([string]$convert.id)))
    } elseif ($met -eq $items.Count) {
        switch ($expect) {
            'Provisioned' {
                $next.Add(('Every object of batch {0} is provisioned. When AD, Exchange and Entra Connect are rebuilt, roll the batch back (Preview first, then -Mode Apply):' -f $convert.id))
                $next.Add((Get-PcmCommandLine -ActionName 'Recover' -BatchId ([string]$convert.id)))
            }
            'Deprovisioned' {
                $next.Add('Every object is back under Active Directory. To check that the cloud data is kept as the journal says:')
                $next.Add((Get-PcmCommandLine -ActionName 'Check' -BatchId ([string]$given.id) -ExpectName 'Retained' -KeepFilters))
            }
            default { $next.Add('The cloud data of the batch is protected as the journal says. Nothing more to do.') }
        }
    } else {
        $next.Add(('{0} object(s) are not {1} yet (see the detail of each one above). Wait a few minutes, or fix what is named, then look again{2}:' -f ($items.Count - $met), $expect.ToLowerInvariant(), $(if ($Once) { ' (without -Once the command waits for you)' } else { '' })))
        $next.Add((Get-PcmCommandLine -ActionName 'Check' -BatchId ([string]$given.id) -ExpectName $expect -KeepFilters -NoOnce))
        if ($expect -eq 'Provisioned' -and $hasRecover) {
            $next.Add(('Recover already rolled back objects of Convert batch {0}, so this batch is not resumed. Convert again the objects that are not converted (Preview first, then -Mode Apply), choosing them with -Identity, -CsvPath or -Group:' -f $convert.id))
            $next.Add((Get-PcmCommandLine -ActionName 'Convert'))
        } elseif ($expect -eq 'Provisioned') {
            $next.Add('An object that Convert did not finish is resumed with the same batch (nothing is converted twice):')
            $next.Add((Get-PcmCommandLine -ActionName 'Convert' -BatchId ([string]$convert.id)))
        } elseif ($expect -eq 'Deprovisioned') {
            $next.Add('An object that is not back yet is rolled back (or its rollback is resumed) with the same batch:')
            $next.Add((Get-PcmCommandLine -ActionName 'Recover' -BatchId ([string]$convert.id)))
        }
    }
    $context.NextSteps = $next.ToArray()
}

# =================================================================================================
# Convert and Recover - cloud admin server (PowerShell 7). Procedure described in the administrator guide.
# =================================================================================================
function Confirm-PcmApply {
    <#
    One confirmation for the whole run, after the plan is shown (none with -Force, -Confirm or -Confirm:$false). After it, the individual changes do not ask
    again (unless -Confirm is given explicitly: then each change asks).
    #>
    param([Parameter(Mandatory)][string]$Question)
    if (-not (Test-PcmSameText $context.Mode 'Apply') -or $Force -or $script:ConfirmBound) { return }
    $context.CurrentOperation = 'Confirm-Apply'; $context.CurrentIdentity = ''
    Write-Host ''
    $answer = $false
    try { $answer = $caller.ShouldContinue($Question, 'PRA Cloud Mailbox - Apply') }
    catch { $context.ConfirmationUnavailable = $true; throw "Apply needs a confirmation and this console cannot ask for one ($($_.Exception.Message)). Add -Force." }
    if (-not $answer) { $context.Cancelled = $true; $context.NoChangeMade = $true; throw 'Cancelled by the operator: nothing was changed.' }
    $context.ConfirmPreference = 'None'
    Write-PcmLog -Context $context -Message ('Operator confirmed: ' + $Question) -Level Detail
}

function Write-PcmJournal {
    <# One event of the batch (journal + log) and, optionally, the new state of the object. #>
    param([string]$ObjectGuid = '', [Parameter(Mandatory)][string]$Step, [ValidateSet('Ok','Fail','Info')][string]$Outcome = 'Ok', [string]$Detail = '', [hashtable]$Item)
    if ($context.Journal -and $context.BatchId) {
        Add-PcmBatchEvent $context.Journal $context.BatchId -ObjectGuid $ObjectGuid -Step $Step -Outcome $Outcome -Detail $Detail
        # A resumed object that ends Done must not keep the error message of an earlier run.
        if ($Item -and $Item['status'] -eq 'Done' -and -not $Item.ContainsKey('message')) { $Item['message'] = '' }
        if ($Item -and $ObjectGuid) { Set-PcmBatchItem $context.Journal $context.BatchId $ObjectGuid $Item }
    }
    Write-PcmLog -Context $context -Message ('[{0}] {1} {2}{3}' -f $context.BatchId, $Step, $Outcome, $(if ($Detail) { ': ' + $Detail } else { '' })) -Level Detail
}

function Get-PcmWaitSeconds {
    param([Parameter(Mandatory)][ValidateSet('MailboxTimeoutMinutes','SyncTimeoutMinutes','HoldTimeoutMinutes')][string]$Name)
    return [int]$context.Config.Polling[$Name] * 60
}

function Get-PcmOriginalLicence {
    <# Direct licences recorded by Convert in original_json (none when the batch did not record them). #>
    param([AllowNull()][object]$Item)
    if (-not $Item -or -not $Item.original_json) { return @() }
    try { $original = [string]$Item.original_json | ConvertFrom-PcmJson } catch { return @() }
    if (-not $original.PSObject.Properties['licences']) { return @() }
    return @($original.licences | ForEach-Object { $_ } | Where-Object { $_ })
}

function Get-PcmLicenceNeed {
    <#
    Objects that will take a NEW licence unit: users without the SKU of their Exchange plan (Group, Direct; Kiosk uses
    the licence they have), shared mailboxes without the temporary SKU that are not converted yet. An object that already
    holds the SKU (hybrid user with Exchange disabled, resume) takes no unit: a user consumes one unit per SKU.
    Entries: Record, State (Get-PcmCloudState) and Step (journal step on resume, else '').
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$Tenant)
    $config = $context.Config
    $userSkus = @(switch ($config.Licensing.Users.Mode) {
            'Group' { if ($Tenant.LicenceGroup) { @($Tenant.LicenceGroup.assignedLicenses) | ForEach-Object { [string]$_.skuId } } }
            'Direct' { $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Users.SkuPartNumber } | ForEach-Object { [string]$_.skuId } }
        })
    $sharedSkus = @($Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | ForEach-Object { [string]$_.skuId })
    $users = 0; $shared = 0; $resources = 0; $resourceGroups = @{}
    foreach ($entry in $Entries) {
        $user = if ($entry.State) { $entry.State.User } else { $null }
        $held = @(if ($user -and $user.PSObject.Properties['assignedLicenses']) { $user.assignedLicenses | ForEach-Object { [string]$_.skuId } })
        if ((Get-PcmValue $entry.State 'NativeResource' $false) -or (Test-PcmNativeResource -Config $config -Record $entry.Record)) {
            $intent = if ($context.ContainsKey('ResourceJournal')) { Get-PcmResourceIntent -Item $context.ResourceJournal[[string]$entry.Record.object_guid] } else { $null }
            if ((Get-PcmResourceProtectionMode $intent) -eq 'None') { continue }
            $groupId = [string]$intent.licensingGroupId
            if (-not $resourceGroups.ContainsKey($groupId)) { $resourceGroups[$groupId] = 0 }
            $licensed = $user -and @($user.licenseAssignmentStates | Where-Object { [string]$_.assignedByGroup -eq $groupId -and [string]$_.state -eq 'Active' }).Count
            if (-not $licensed) { $resources++; $resourceGroups[$groupId]++ }
        } elseif ($entry.Record.kind -eq 'User') {
            if ($config.Licensing.Users.Mode -eq 'Kiosk' -or @($held | Where-Object { $_ -in $userSkus }).Count) { continue }
            $users++
        } else {
            if ($entry.State -and $entry.State.RecipientType -eq 'SharedMailbox') { continue }
            if ([string]$entry.Step -in @('Shared', 'Tagged', 'LicenceRemoved', 'Done') -or @($held | Where-Object { $_ -in $sharedSkus }).Count) { continue }
            $shared++
        }
    }
    return [pscustomobject]@{ Users = $users; Shared = $shared; Resources = $resources; ResourceGroups = $resourceGroups
        SkipUserLicensing = (@($Entries | Where-Object { $_.Record.kind -eq 'User' }).Count -eq 0 -and @($Entries | Where-Object {
                    (Get-PcmValue $_.State 'NativeResource' $false) -or (Test-PcmNativeResource -Config $config -Record $_.Record)
                }).Count -gt 0) }
}

function Get-PcmResourceLicenceFinding {
    param([object]$Need, [object]$Tenant)
    $groups = @{}; $demand = @{}
    $userSkus = @(switch ($context.Config.Licensing.Users.Mode) {
            Group { if ($Tenant.LicenceGroup) { $Tenant.LicenceGroup.assignedLicenses | ForEach-Object { [string]$_.skuId } } }
            Direct { $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$context.Config.Licensing.Users.SkuPartNumber } | ForEach-Object { [string]$_.skuId } }
        })
    foreach ($id in $Need.ResourceGroups.Keys) {
        try {
            if (-not $id) { throw 'Resources.LicensingGroupId is empty: native resources need enterprise licensing throughout their active lifetime.' }
            $group = Get-PcmResourceGroup -Id $id -Skus @($Tenant.Skus)
            $groups[$id] = $group
            foreach ($assignment in $group.Group.assignedLicenses) {
                $skuId = [string]$assignment.skuId
                if (-not $demand.ContainsKey($skuId)) { $demand[$skuId] = $(if ($skuId -in $userSkus) { $Need.Users } else { 0 }) }
                $demand[$skuId] += $Need.ResourceGroups[$id]
            }
        } catch { [pscustomobject]@{ Level = 'Error'; Code = 'RESOURCE_LICENCE'; Message = $_.Exception.Message } }
    }
    foreach ($skuId in $demand.Keys) {
        $sku = $Tenant.Skus | Where-Object { [string]$_.skuId -eq $skuId } | Select-Object -First 1
        if ($demand[$skuId] -gt 0 -and (Get-PcmFreeUnit $sku) -lt $demand[$skuId]) {
            [pscustomobject]@{ Level = 'Error'; Code = 'RESOURCE_LICENCE'; Message = "Resource licensing capacity insufficient: $($sku.skuPartNumber) needs $($demand[$skuId]) new permanent unit(s) across selected users and resources." }
        }
    }
    foreach ($id in $groups.Keys) {
        [pscustomobject]@{ Level = 'Ok'; Code = 'RESOURCE_LICENCE'; Message = "Resource licence group ${id}: enterprise plan checked ($($Need.ResourceGroups[$id]) new permanent unit(s)); resources do not recycle temporary shared units." }
    }
}

function Get-PcmObjectAssessment {
    <# Cloud state (read in bulk beforehand, or now) + readiness findings of one snapshot object (Check rules). #>
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][object]$Tenant, [object[]]$Permissions = @(), [object[]]$Members = @(), [scriptblock]$Lookup, [object]$State)
    if ($State -is [System.Management.Automation.ErrorRecord]) { throw $State }
    $state = if ($State) { $State } else { Get-PcmCloudState -Context $context -Record $Record }
    $findings = @(Get-PcmReadiness -Config $context.Config -Record $Record -State $state -Skus @($Tenant.Skus))
    if (Get-PcmValue $state 'NativeResource' $false) {
        if ($context.Mode -eq 'Preview') {
            try { Test-PcmResourceConfigurationReadiness -Record $Record -AllowRoomListCreation -RoomListPlan @((Read-PcmResourceSettings $Record).RoomLists) -RoomListIntents (Get-PcmValue $context 'RoomListIntents' @{}) }
            catch { $findings += [pscustomobject]@{ Level = 'Error'; Field = 'ExchangeOnline'; Code = 'RESOURCE_CONFIGURATION_REQUIRED'; Message = $_.Exception.Message } }
        }
    } elseif ($Record.kind -ne 'User' -and $Lookup) { $findings += @(Get-PcmTrusteeFinding -Record $Record -Permissions $Permissions -Members $Members -Lookup $Lookup) }
    return [pscustomobject]@{ State = $state; Findings = $findings; Errors = @($findings | Where-Object Level -eq 'Error'); Warnings = @($findings | Where-Object Level -eq 'Warn') }
}

function Get-PcmPermissionIndex {
    <# Permissions by mailbox and group members by group: each shared mailbox reads only its own (5,000 shared mailboxes). #>
    param([AllowEmptyCollection()][object[]]$Permissions = @(), [AllowEmptyCollection()][object[]]$Members = @())
    $byMailbox = @{}; $byGroup = @{}
    foreach ($permission in $Permissions) {
        $key = [string]$permission.mailbox_guid
        if (-not $byMailbox.ContainsKey($key)) { $byMailbox[$key] = [System.Collections.Generic.List[object]]::new() }
        $byMailbox[$key].Add($permission)
    }
    foreach ($member in $Members) {
        $key = [string]$member.group_guid
        if (-not $byGroup.ContainsKey($key)) { $byGroup[$key] = [System.Collections.Generic.List[object]]::new() }
        $byGroup[$key].Add($member)
    }
    return [pscustomobject]@{ ByMailbox = $byMailbox; ByGroup = $byGroup }
}

function Get-PcmPermissionSubset {
    <# The permissions of one mailbox and the members of the groups they name. #>
    param([Parameter(Mandatory)][object]$Index, [Parameter(Mandatory)][string]$MailboxGuid)
    $permissions = @(if ($Index.ByMailbox.ContainsKey($MailboxGuid)) { $Index.ByMailbox[$MailboxGuid] })
    $groups = @($permissions | Where-Object { [string]$_.trustee_kind -eq 'Group' } | ForEach-Object { [string]$_.trustee_guid } | Select-Object -Unique)
    $members = @(foreach ($group in $groups) { if ($Index.ByGroup.ContainsKey($group)) { $Index.ByGroup[$group] } })
    return [pscustomobject]@{ Permissions = $permissions; Members = $members }
}

function New-PcmTrusteeLookup {
    <#
    Is a trustee known in Exchange Online? The objects of the snapshot that have a recipient answer from the states
    read in bulk; any other identity is asked once (Get-Recipient) and remembered.
    #>
    param([hashtable]$States = @{}, [AllowEmptyCollection()][object[]]$Records = @())
    $known = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($record in $Records) {
        $state = $States[[string]$record.object_guid]
        if (-not $state -or $state -is [System.Management.Automation.ErrorRecord] -or -not $state.RecipientType) { continue }
        foreach ($name in @([string]$record.user_principal_name, [string]$record.primary_smtp_address, $(if ($state.User) { [string]$state.User.userPrincipalName } else { '' }))) { if ($name) { [void]$known.Add($name) } }
    }
    $cache = @{}
    return {
        param($identity)
        $key = ([string]$identity).ToLowerInvariant()
        if ($known.Contains($key)) { return $true }
        if (-not $cache.ContainsKey($key)) { $cache[$key] = [bool](Get-Recipient -Identity $identity -ErrorAction SilentlyContinue) }
        return $cache[$key]
    }.GetNewClosure()
}

function Read-PcmCloudStates {
    <# The cloud state of every object in bulk (Get-PcmCloudStateSet), with its duration in the console. #>
    param([AllowEmptyCollection()][object[]]$Records = @())
    if (-not $Records.Count) { return @{} }
    $context.CurrentOperation = 'Read-States'
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $states = Get-PcmCloudStateSet -Context $context -Records $Records
    Write-PcmItem -Context $context -Status Info -Text ('Cloud state of {0} object(s) read in {1}' -f $Records.Count, (Format-PcmDuration $clock.Elapsed.TotalSeconds))
    return $states
}

function Wait-PcmExoStateSet {
    <#
    Polls the Exchange Online state of many objects together (Get-PcmExoStateSet) until -Test is true for each or the
    timeout; -OnReady runs for each object as soon as it is ready. Returns the entries still not ready.
    Entries need an EntraId property.
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][scriptblock]$Test, [scriptblock]$OnReady, [Parameter(Mandatory)][int]$TimeoutSeconds, [switch]$StopOnRequest, [string]$Phase = '')
    $waiting = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $Entries) { $waiting.Add($entry) }
    if (-not $waiting.Count) { return @() }
    # The poll runs in the scope of Wait-PcmCondition, whose own parameter is $Test: other names here.
    $readyTest = $Test; $readyAction = $OnReady; $honourStop = [bool]$StopOnRequest
    $readyClock = [Diagnostics.Stopwatch]::StartNew()
    $null = Wait-PcmCondition -Phase $Phase -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
        $states = Get-PcmExoStateSet -Id ([string[]]@($waiting | ForEach-Object { $_.EntraId })) -TagAttribute ([string]$context.Config.Retention.TagAttribute)
        foreach ($one in @($waiting)) {
            $state = $states[[string]$one.EntraId]
            if ($state -and (& $readyTest $state $one)) { [void]$waiting.Remove($one); if ($readyAction) { & $readyAction $one $state ([int]$readyClock.Elapsed.TotalSeconds) } }
        }
        (-not $waiting.Count) -or ($honourStop -and (Test-PcmStop))
    }
    return $waiting.ToArray()
}

function Find-PcmRecordRow {
    param([Parameter(Mandatory)][string]$ObjectGuid)
    return ($context.Rows | Where-Object { $_.ObjectGuid -eq $ObjectGuid } | Select-Object -First 1)
}

function Set-PcmRowResult {
    <# Final status and detail of one report row, plus the console line. #>
    param([Parameter(Mandatory)][object]$Row, [Parameter(Mandatory)][ValidateSet('Success','Error','Pending','Skipped','Planned')][string]$Status, [string]$Detail = '', [string]$Field, [string]$Value)
    $Row.FinalStatus = $Status
    $Detail = ConvertTo-PcmPlainMessage $Detail
    if ($Detail) { $Row.Detail = $Detail }
    if ($Field) { $Row[$Field] = $Value }
    $item = @{ Success = 'Ok'; Error = 'Fail'; Pending = 'Warn'; Skipped = 'Skip'; Planned = 'Info' }[$Status]
    if ($Status -eq 'Error') {
        # The error belongs to this object, whatever object the step was on: its row and its recorded error are one failure.
        $previousIdentity = [string](Get-PcmValue $context 'CurrentIdentity' '')
        $context.CurrentIdentity = [string]$Row.Identity
        try { Write-PcmLog -Context $context -Message ('{0}: {1}' -f $Row.Identity, $Detail) -Level Error }
        finally { $context.CurrentIdentity = $previousIdentity }
    }
    else { Write-PcmItem -Context $context -Status $item -Text ('{0}: {1}' -f $Row.Identity, $Detail) }
}

function Stop-PcmEntry {
    <# An object that fails in a phase: journal, report row, and it leaves the next phases. #>
    param([Parameter(Mandatory)][object]$Entry, [Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Message)
    $text = (ConvertTo-PcmPlainMessage $Message) -replace '\s+', ' '
    Write-PcmJournal -ObjectGuid ([string]$Entry.Record.object_guid) -Step $Step -Outcome Fail -Detail $text -Item @{ status = 'Failed'; message = $text }
    Set-PcmRowResult -Row $Entry.Row -Status Error -Detail "Convert stopped: $text"
}

function Invoke-PcmIdentityStartSet {
    <#
    Common start of a conversion, phase by phase for many objects: ExchangeGuid cleared, source of authority in the cloud (Graph batches), ONE wait for all until
    Exchange Online no longer sees them as synchronised (typically 23 to 45 s), usageLocation set from the cloud
    (Graph batches). An object that fails is reported and left out. Returns the entries that went through.
    Entries: Record, Row, EntraId, State (Get-PcmCloudState).
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @())
    $config = $context.Config
    if (-not $Entries.Count) { return @() }
    $exo = Get-PcmExoStateSet -Id ([string[]]@($Entries | ForEach-Object { $_.EntraId })) -TagAttribute $config.Retention.TagAttribute
    $active = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $Entries) {
        $context.CurrentIdentity = $entry.Row.Identity
        $state = $exo[$entry.EntraId]
        try {
            if ($state -and $state.Type -eq 'MailUser' -and $state.ExchangeGuid -ne [guid]::Empty.ToString()) {
                Confirm-PcmWrite -Target $entry.EntraId -Operation 'Exchange Online: clear the ExchangeGuid of the synchronised user'
                Set-MailUser -Identity $entry.EntraId -ExchangeGuid ([guid]::Empty) -ErrorAction Stop
                Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'GuidCleared' -Detail ("ExchangeGuid {0} cleared" -f $state.ExchangeGuid) -Item @{ step = 'GuidCleared'; status = 'Running' }
            }
            $active.Add($entry)
        } catch { Stop-PcmEntry -Entry $entry -Step 'GuidCleared' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    $behaviors = Get-PcmSyncBehaviorSet -UserId ([string[]]@($active | ForEach-Object { $_.EntraId }))
    $toCloud = @($active | Where-Object { $behaviors[$_.EntraId] -ne $true })
    $answers = Invoke-PcmGraphBatch -Requests @($toCloud | ForEach-Object { @{ Method = 'PATCH'; Url = ('/users/{0}/onPremisesSyncBehavior' -f $_.EntraId); Body = @{ isCloudManaged = $true } } })
    for ($i = 0; $i -lt $toCloud.Count; $i++) {
        if ($answers[$i].Ok) { Write-PcmJournal -ObjectGuid ([string]$toCloud[$i].Record.object_guid) -Step 'SoaCloud' -Detail 'isCloudManaged = true' -Item @{ step = 'SoaCloud'; status = 'Running' } }
        else { Stop-PcmEntry -Entry $toCloud[$i] -Step 'SoaCloud' -Message ('source of authority not transferred ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$active.Remove($toCloud[$i]) }
    }
    $seconds = @{}
    $late = @(Wait-PcmExoStateSet -Phase 'Waiting for the identities to become cloud-only' -Entries $active.ToArray() -TimeoutSeconds (Get-PcmWaitSeconds 'SyncTimeoutMinutes') -Test { param($state) $state.IsDirSynced -eq $false } -OnReady {
            $seconds[$args[0].EntraId] = $args[2] })
    foreach ($entry in $late) { Stop-PcmEntry -Entry $entry -Step 'SoaCloud' -Message 'Exchange Online still sees the object as synchronised from AD after the source of authority transfer.'; [void]$active.Remove($entry) }
    # Set from the cloud even when present: the value synchronised from AD is refused right after the transfer.
    $countries = @{}
    foreach ($entry in @($active)) {
        $current = if ($entry.State -and $entry.State.User) { [string]$entry.State.User.usageLocation } else { '' }
        $country = if ($current) { $current } else { [string]$config.Cloud.DefaultUsageLocation }
        if (-not $country) { Stop-PcmEntry -Entry $entry -Step 'UsageLocation' -Message 'usageLocation is empty and Cloud.DefaultUsageLocation is not set (a licence needs a country).'; [void]$active.Remove($entry); continue }
        $countries[$entry.EntraId] = $country
    }
    $set = $active.ToArray()
    $answers = Invoke-PcmGraphBatch -Requests @($set | ForEach-Object { @{ Method = 'PATCH'; Url = ('/users/{0}' -f $_.EntraId); Body = @{ usageLocation = $countries[$_.EntraId] } } })
    for ($i = 0; $i -lt $set.Count; $i++) {
        $entry = $set[$i]
        if ($answers[$i].Ok) {
            Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'UsageLocation' -Detail ("usageLocation = {0} (cloud), Exchange Online no longer synchronised after {1} s" -f $countries[$entry.EntraId], $seconds[$entry.EntraId]) -Item @{ step = 'UsageLocation'; status = 'Running' }
        } else { Stop-PcmEntry -Entry $entry -Step 'UsageLocation' -Message ('usageLocation not set ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$active.Remove($entry) }
    }
    return $active.ToArray()
}

function Invoke-PcmUserPlanSet {
    <#
    Users: identity start of the wave (Invoke-PcmIdentityStartSet), then the Exchange plan: membership of the licence
    group by Graph batches (already a member is not an error), or the Kiosk / direct plan one user at a time.
    Returns the users whose plan is assigned (their mailboxes are awaited together afterwards).
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$Tenant)
    $config = $context.Config
    $users = $config.Licensing.Users
    $started = @(Invoke-PcmIdentityStartSet -Entries $Entries)
    $done = [System.Collections.Generic.List[object]]::new()
    $assigned = {
        param($entry, $detail)
        Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'PlanAssigned' -Detail $detail -Item @{ step = 'PlanAssigned'; status = 'Running' }
        Write-PcmItem -Context $context -Status Ok -Text ('{0}: source of authority in the cloud, Exchange plan assigned' -f $entry.Row.Identity)
        $done.Add($entry)
    }
    if ($users.Mode -eq 'Group') {
        # Members already (resume) are not added again.
        $checks = Invoke-PcmGraphBatch -Requests @($started | ForEach-Object { @{ Method = 'POST'; Url = ('/users/{0}/checkMemberGroups' -f $_.EntraId); Body = @{ groupIds = @($users.GroupId) } } })
        $toAdd = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $started.Count; $i++) {
            if ($checks[$i].Ok -and @((Get-PcmValue $checks[$i].Body 'value' @()) | ForEach-Object { $_ }) -contains $users.GroupId) { & $assigned $started[$i] "member of the licence group $($users.GroupId)" }
            else { $toAdd.Add($started[$i]) }
        }
        $answers = Invoke-PcmGraphBatch -Requests @($toAdd | ForEach-Object { @{ Method = 'POST'; Url = ('/groups/{0}/members/$ref' -f $users.GroupId); Body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($_.EntraId)" } } })
        for ($i = 0; $i -lt $toAdd.Count; $i++) {
            if ($answers[$i].Ok -or $answers[$i].Error -match 'already exist') { & $assigned $toAdd[$i] "member of the licence group $($users.GroupId)" }
            else { Stop-PcmEntry -Entry $toAdd[$i] -Step 'ConvertUser' -Message ('not added to the licence group ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error) }
        }
        return $done.ToArray()
    }
    foreach ($entry in $started) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            if ($users.Mode -eq 'Kiosk') {
                $kiosk = Get-PcmKioskAssignment -Licence (Get-PcmUserLicence -UserId $entry.EntraId) -Skus @($Tenant.Skus)
                if (-not $kiosk) { throw 'Kiosk mode: no licence of this user contains EXCHANGE_S_DESKLESS.' }
                # Kiosk enabled in the existing direct assignment (other plans unchanged), or a direct assignment next to the group one.
                $how = Enable-PcmLicencePlan -UserId $entry.EntraId -SkuId $kiosk.SkuId -PlanIds @($kiosk.KioskPlanId) -DisabledPlans $kiosk.DisabledPlans
                & $assigned $entry "$($kiosk.SkuName): Exchange Kiosk $($how.ToLowerInvariant())"
            } else {
                $sku = $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$users.SkuPartNumber } | Select-Object -First 1
                if (-not $sku) { throw "Licensing.Users.SkuPartNumber $($users.SkuPartNumber) not found in the tenant." }
                $how = Enable-PcmLicencePlan -UserId $entry.EntraId -SkuId $sku.skuId -PlanIds @(Get-PcmMailboxPlanId -Skus @($sku)) -DisabledPlans (Get-PcmDisabledPlanForExchangeOnly -Sku $sku)
                & $assigned $entry "$($sku.skuPartNumber): Exchange plan $($how.ToLowerInvariant())"
            }
        } catch { Stop-PcmEntry -Entry $entry -Step 'ConvertUser' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    return $done.ToArray()
}

function Select-PcmSharedWave {
    <#
    From the next shared mailboxes of a Convert, the ones that can start now: an object that needs no licence unit (already a shared mailbox, or already holding the temporary unit) can
    always start, the others only while a unit is free. -NeedsUnit says by object ID whether an object needs a new unit. Returns Wave (to start) and Waiting (needs a unit, none is free).
    #>
    param([AllowEmptyCollection()][object[]]$Candidates = @(), [Parameter(Mandatory)][hashtable]$NeedsUnit, [int]$Free = 0)
    $wave = [System.Collections.Generic.List[object]]::new()
    $waiting = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $Candidates) {
        if (-not $NeedsUnit[[string]$entry.EntraId]) { $wave.Add($entry); continue }
        if ($Free -ge 1) { $Free--; $wave.Add($entry) } else { $waiting.Add($entry) }
    }
    return [pscustomobject]@{ Wave = $wave.ToArray(); Waiting = $waiting.ToArray() }
}

function Invoke-PcmConvertSharedWave {
    <#
    A wave of shared mailboxes, phase by phase (one temporary unit each): identity start, temporary licence, ONE wait
    for all the mailboxes, Type Shared asked again until it is effective (the licence must stay until then),
    tag, licence given back as it was, ONE pause of 60 s for the wave, every object checked still SharedMailbox, then
    the permissions of the snapshot (a shared mailbox needs no licence for them). An object that fails keeps its
    temporary licence and is reported; the others go on.
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$Tenant, [Parameter(Mandatory)][object]$PermissionIndex, [hashtable]$Originals = @{})
    $config = $context.Config
    $tagAttribute = [string]$config.Retention.TagAttribute
    $sku = $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
    if (-not $sku) { throw "Licensing.Shared.SkuPartNumber $($config.Licensing.Shared.SkuPartNumber) not found in the tenant." }
    $ids = { param($set) [string[]]@($set | ForEach-Object { $_.EntraId }) }
    $exo = Get-PcmExoStateSet -Id (& $ids $Entries) -TagAttribute $tagAttribute
    $already = @($Entries | Where-Object { $exo[$_.EntraId] -and $exo[$_.EntraId].Type -eq 'SharedMailbox' })
    $toConvert = @($Entries | Where-Object { $_ -notin $already })
    $licensed = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @(Invoke-PcmIdentityStartSet -Entries $toConvert)) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            $how = Enable-PcmLicencePlan -UserId $entry.EntraId -SkuId $sku.skuId -PlanIds @(Get-PcmMailboxPlanId -Skus @($sku)) -DisabledPlans (Get-PcmDisabledPlanForExchangeOnly -Sku $sku)
            if ($how -ne 'Present') { Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'LicenceAssigned' -Detail "temporary $($sku.skuPartNumber) (Exchange plan only, $($how.ToLowerInvariant()))" -Item @{ step = 'LicenceAssigned'; status = 'Running' } }
            $licensed.Add($entry)
        } catch { Stop-PcmEntry -Entry $entry -Step 'LicenceAssigned' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    $late = @(Wait-PcmExoStateSet -Phase 'Waiting for the shared mailboxes to appear' -Entries $licensed.ToArray() -TimeoutSeconds (Get-PcmWaitSeconds 'MailboxTimeoutMinutes') -Test { param($state) $state.Type -in @('UserMailbox', 'SharedMailbox') } -OnReady {
            Write-PcmJournal -ObjectGuid ([string]$args[0].Record.object_guid) -Step 'MailboxCreated' -Detail "cloud mailbox after $($args[2]) s" })
    foreach ($entry in $late) { Stop-PcmEntry -Entry $entry -Step 'MailboxCreated' -Message 'No cloud mailbox after the temporary licence (the licence is kept: run Convert -Batch again later).'; [void]$licensed.Remove($entry) }
    # Type Shared for all, asked again every 15 s until it is effective (10 min).
    $pending = [System.Collections.Generic.List[object]]::new($licensed)
    $ready = [System.Collections.Generic.List[object]]::new($already)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while ($pending.Count) {
        $states = Get-PcmExoStateSet -Id (& $ids $pending) -TagAttribute $tagAttribute
        foreach ($entry in @($pending)) {
            if ($states[$entry.EntraId] -and $states[$entry.EntraId].Type -eq 'SharedMailbox') {
                Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'Shared' -Detail ("SharedMailbox after {0} s" -f [int]$clock.Elapsed.TotalSeconds) -Item @{ step = 'Shared'; status = 'Running' }
                [void]$pending.Remove($entry); $ready.Add($entry)
            } else {
                Confirm-PcmWrite -Target $entry.EntraId -Operation 'Exchange Online: turn the mailbox into a shared mailbox'
                try { Set-Mailbox -Identity $entry.EntraId -Type Shared -ErrorAction Stop } catch { $null = $_ }
            }
        }
        if (-not $pending.Count -or $clock.Elapsed.TotalSeconds -ge 600) { break }
        Start-Sleep -Seconds 15
    }
    foreach ($entry in $pending) { Stop-PcmEntry -Entry $entry -Step 'Shared' -Message 'The mailbox is not a SharedMailbox after 10 min (the licence is kept).' }
    # Tag, then the temporary licence back to what each object had before Convert.
    $states = Get-PcmExoStateSet -Id (& $ids $ready) -TagAttribute $tagAttribute
    $back = [System.Collections.Generic.List[object]]::new()
    $removed = 0
    foreach ($entry in $ready) {
        $context.CurrentIdentity = $entry.Row.Identity
        $guid = [string]$entry.Record.object_guid
        try {
            if (-not $states[$entry.EntraId] -or $states[$entry.EntraId].Tag -ne $config.Retention.TagValue) {
                $tagParameters = @{ Identity = $entry.EntraId; ErrorAction = 'Stop' }; $tagParameters[$tagAttribute] = $config.Retention.TagValue
                Confirm-PcmWrite -Target $entry.EntraId -Operation ('Exchange Online: set {0} to the retention tag value' -f $tagAttribute)
                Set-Mailbox @tagParameters
            }
            Write-PcmJournal -ObjectGuid $guid -Step 'Tagged' -Detail "$tagAttribute = $($config.Retention.TagValue)" -Item @{ step = 'Tagged'; status = 'Running' }
            $how = Restore-PcmLicence -UserId $entry.EntraId -SkuId $sku.skuId -Original @($Originals[$guid])
            if ($how -in @('Removed', 'Restored')) { $removed++; Write-PcmJournal -ObjectGuid $guid -Step 'LicenceRemoved' -Detail "temporary $($sku.skuPartNumber) $($how.ToLowerInvariant())" -Item @{ step = 'LicenceRemoved'; status = 'Running' } }
            $back.Add($entry)
        } catch { Stop-PcmEntry -Entry $entry -Step 'Tagged' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    # One pause for the wave: a mailbox converted too early would be disabled once its licence is gone.
    if ($removed) { Start-Sleep -Seconds 60 }
    $final = Get-PcmExoStateSet -Id (& $ids $back) -TagAttribute $tagAttribute
    foreach ($entry in $back) {
        $context.CurrentIdentity = $entry.Row.Identity
        $guid = [string]$entry.Record.object_guid
        try {
            $state = $final[$entry.EntraId]
            if (-not $state -or $state.Type -ne 'SharedMailbox') { throw "After the licence removal the object is $(if ($state) { $state.Type } else { 'not found' }), not a SharedMailbox." }
            $subset = Get-PcmPermissionSubset -Index $PermissionIndex -MailboxGuid $guid
            $grants = @(Get-PcmSharedGrant -Record $entry.Record -Permissions $subset.Permissions -Members $subset.Members)
            $granted = 0; $failed = [System.Collections.Generic.List[string]]::new()
            foreach ($grant in $grants) {
                try { $outcome = Grant-PcmSharedPermission -Mailbox $entry.EntraId -Grant $grant; $granted++; Write-PcmJournal -ObjectGuid $guid -Step 'Permission' -Detail ('{0} {1} ({2}, {3})' -f $grant.Right, $grant.Trustee, $outcome, $grant.Source) }
                catch { [void]$failed.Add(('{0} {1}: {2}' -f $grant.Right, $grant.Trustee, ($_.Exception.Message -replace '\s+', ' '))); Write-PcmJournal -ObjectGuid $guid -Step 'Permission' -Outcome Fail -Detail $failed[$failed.Count - 1] }
            }
            $permissionText = '{0}/{1} permission(s) granted' -f $granted, $grants.Count
            Write-PcmJournal -ObjectGuid $guid -Step 'Done' -Detail $permissionText -Item @{ step = 'Done'; status = $(if ($failed.Count) { 'Running' } else { 'Done' }); result_json = (@{ cloudGuid = $state.ExchangeGuid; failed = $failed.ToArray() } | ConvertTo-Json -Compress) }
            $entry.Row.Permissions = $permissionText
            Set-PcmRowResult -Row $entry.Row -Status $(if ($failed.Count) { 'Pending' } else { 'Success' }) -Detail $(if ($failed.Count) { 'shared mailbox ready; permissions not granted: ' + ($failed -join '; ') } else { "shared mailbox $($state.ExchangeGuid), tagged, $permissionText" }) -Field 'ExchangeOnline' -Value 'SharedMailbox'
        } catch { Stop-PcmEntry -Entry $entry -Step 'ConvertShared' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
}


function Write-PcmResourceJournal {
    param([object]$Entry, [string]$Step, [string]$Status = 'Running', [string]$Detail = '', [hashtable]$Extra = @{})
    $guid = [string]$(if ($Entry.PSObject.Properties['Record']) { $Entry.Record.object_guid } else { $Entry.Item.object_guid })
    $result = @{ lifecycle = 'CloudResource'; cloudId = [string]$Entry.Intent.cloudId; cloudGuid = [string]$Entry.Intent.cloudGuid
        cloudType = [string]$Entry.Intent.cloudType; ownedName = [string]$Entry.Intent.creationName; markerAttribute = [string]$Entry.Intent.markerAttribute; token = [string]$Entry.Intent.token }
    foreach ($key in $Extra.Keys) { $result[$key] = $Extra[$key] }
    Write-PcmJournal -ObjectGuid $guid -Step $Step -Detail $Detail -Item @{ step = $Step; status = $Status; entra_id = [string]$Entry.Intent.cloudId
        original_json = ($Entry.Intent | ConvertTo-Json -Depth 15 -Compress); result_json = ($result | ConvertTo-Json -Depth 15 -Compress) }
}

function ConvertTo-PcmRoomListSignature {
    param([object]$List)
    $values = [ordered]@{}
    foreach ($field in @('SchemaVersion','SourceGuid','Name','DisplayName','Alias','PrimarySmtpAddress','LegacyExchangeDN',
            'HiddenFromAddressListsEnabled','MemberJoinRestriction','MemberDepartRestriction')) { $values[$field] = $List.$field }
    $values.EmailAddresses = @($List.EmailAddresses | Sort-Object -Unique)
    $delivery = [ordered]@{}
    foreach ($field in @('RequireSenderAuthenticationEnabled','ModerationEnabled','SendModerationNotifications')) { $delivery[$field] = $List.Delivery.$field }
    foreach ($field in @('AcceptMessagesOnlyFromSendersOrMembers','RejectMessagesFromSendersOrMembers','ModeratedBy','BypassModerationFromSendersOrMembers')) {
        $delivery[$field] = @($List.Delivery.$field | Sort-Object -Unique)
    }
    $values.Delivery = $delivery
    return ConvertTo-Json -InputObject $values -Depth 10 -Compress
}

function Get-PcmRoomListPlan {
    param([object[]]$Entries, [object[]]$Records)
    $lists = @{}
    $keys = @{}
    if (-not $Entries.Count) { return @() }
    $mailboxKeys = @{}
    foreach ($record in $Records) {
        foreach ($key in @(Get-PcmResourceKeys $record)) { $mailboxKeys[$key] = [string]$record.primary_smtp_address }
    }
    foreach ($entry in $Entries) {
        foreach ($list in (Read-PcmResourceSettings $entry.Record).RoomLists) {
            $source = [string]$list.SourceGuid
            if ($lists.ContainsKey($source)) {
                if ((ConvertTo-PcmRoomListSignature $lists[$source]) -ine (ConvertTo-PcmRoomListSignature $list)) { throw "Conflicting room-list snapshot metadata for $source; nothing is changed." }
                continue
            }
            foreach ($key in @(Get-PcmResourceKeys (ConvertTo-PcmRoomListRecord $list))) {
                if ($keys.ContainsKey($key)) { throw "Competing room lists claim $key; nothing is changed." }
                if ($mailboxKeys.ContainsKey($key)) { throw "Room list $($list.PrimarySmtpAddress) conflicts with selected mailbox $($mailboxKeys[$key]); nothing is changed." }
                $keys[$key] = $source
            }
            $lists[$source] = $list
        }
    }
    return @($lists.Values | Sort-Object PrimarySmtpAddress)
}

function Write-PcmRoomListJournal {
    param([object]$Intent, [string]$Step, [string]$Status = 'Running', [string]$Detail = '')
    $Intent.status = $Status
    Set-PcmTenantChange -Connection $context.Journal -BatchId $context.BatchId -Kind 'NativeRoomList' -TargetId ([string]$Intent.sourceGuid) `
        -OriginalJson (ConvertTo-Json -InputObject $Intent -Depth 15 -Compress) -UpdateOriginal
    Write-PcmJournal -ObjectGuid ("RoomList:{0}" -f $Intent.sourceGuid) -Step $Step -Outcome $(if ($Status -eq 'Failed') { 'Fail' } else { 'Ok' }) -Detail $Detail
}

function Write-PcmRoomListIdentityPin {
    param([object]$Intent, [object]$Group)
    $guid = [string](Get-PcmValue $Group 'Guid' '')
    $id = [string](Get-PcmValue $Group 'ExternalDirectoryObjectId' '')
    if (-not $guid -or $guid -eq [guid]::Empty.ToString() -or $guid -eq [string]$Intent.sourceGuid -or
        ($Intent.cloudGuid -and [string]$Intent.cloudGuid -ne $guid) -or ($Intent.cloudId -and [string]$Intent.cloudId -ne $id)) {
        throw 'Room-list immutable creation identity is missing/inconsistent; no configuration.'
    }
    $Intent.cloudGuid = $guid
    if ($id) { $Intent.cloudId = $id }
    Write-PcmRoomListJournal -Intent $Intent -Step 'RoomListIdentityPinned' -Detail "Exchange GUID $guid; directory ID recorded only when exposed"
}

function Write-PcmResourceIdentityPin {
    param([object]$Entry, [object]$Mailbox)
    $intent = $Entry.Intent
    $guid = [string](Get-PcmValue $Mailbox 'ExchangeGuid' '')
    $id = [string](Get-PcmValue $Mailbox 'ExternalDirectoryObjectId' '')
    if (-not $id -or -not $guid -or $guid -eq [guid]::Empty.ToString() -or $guid -eq [string]$Entry.Record.exchange_guid -or
        ($intent.cloudId -and [string]$intent.cloudId -ne $id) -or ($intent.cloudGuid -and [string]$intent.cloudGuid -ne $guid)) {
        throw 'Resource immutable creation identity is missing/inconsistent; no configuration.'
    }
    $intent.cloudId = $id; $intent.cloudGuid = $guid
    $Entry.EntraId = $id
    Write-PcmResourceJournal -Entry $Entry -Step 'ResourceIdentityPinned' -Detail "cloud ID $id, ExchangeGuid $guid recorded before further reads or mutations"
}

function Invoke-PcmConvertRoomLists {
    param([object[]]$Lists, [switch]$Unprotected)
    foreach ($list in $Lists) {
        if (Test-PcmStop) { break }
        $context.CurrentIdentity = [string]$list.PrimarySmtpAddress
        $intent = $context.RoomListIntents[[string]$list.SourceGuid]
        try {
            if ($Unprotected) { Assert-PcmResourceOrganizationUnheld }
            if ($intent) {
                if ((ConvertTo-PcmRoomListSignature $intent.list) -ine (ConvertTo-PcmRoomListSignature $list)) { throw 'Room-list metadata differs from its creation journal; nothing is changed.' }
            } else {
                $existing = Get-PcmResourceCloudRoomList -List $list -AllowMissing
                $token = [guid]::NewGuid().ToString('N')
                $intent = [pscustomobject]@{ sourceGuid = [string]$list.SourceGuid; lifecycle = 'CloudRoomList'; list = $list
                    operation = $(if ($existing) { 'Reuse' } else { 'Create' }); status = 'Planned'; creationAttempted = $false
                    token = $token; creationName = ('PRA-L-{0}-{1}-{2}' -f $context.BatchId.Substring(0,8), $token.Substring(0,8), $token)
                    cloudGuid = $(if ($existing) { [string]$existing.Guid } else { '' }); cloudId = $(if ($existing) { [string](Get-PcmValue $existing 'ExternalDirectoryObjectId' '') } else { '' }) }
                $context.RoomListIntents[[string]$list.SourceGuid] = $intent
                Write-PcmRoomListJournal -Intent $intent -Step 'RoomListIntent' -Detail "$($intent.operation) $($list.PrimarySmtpAddress); selected native rooms only"
            }
            $group = Get-PcmResourceCloudRoomList -List $list -Intent $intent -AllowMissing
            $wasDone = $intent.status -eq 'Done'
            if ($intent.operation -eq 'Create' -and $intent.status -ne 'Done') {
                Test-PcmRoomListCreationReadiness -List $list
                if (-not $group) {
                    if ($intent.creationAttempted) { throw 'Room-list creation was already attempted but its exact object is not visible; resume only after visibility is resolved, never issue a second create.' }
                    Confirm-PcmWrite -Target ([string]$list.PrimarySmtpAddress) -Operation 'Exchange Online: create the room list'
                    $intent.creationAttempted = $true
                    Write-PcmRoomListJournal -Intent $intent -Step 'RoomListCreationIntent' -Detail "create $($intent.creationName); token recorded before New-DistributionGroup"
                    $created = @(New-DistributionGroup -Name $intent.creationName -Alias $list.Alias -DisplayName $list.DisplayName `
                        -PrimarySmtpAddress $list.PrimarySmtpAddress -RoomList -IgnoreNamingPolicy `
                        -MemberJoinRestriction Closed -MemberDepartRestriction Closed -RequireSenderAuthenticationEnabled $true -ErrorAction Stop)
                    if ($created.Count -ne 1) { throw 'Room-list creation response is missing/ambiguous; never repeat creation.' }
                    $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force
                    Write-PcmRoomListIdentityPin -Intent $intent -Group $created[0]
                    $group = Get-PcmResourceCloudRoomList -List $list -Intent $intent
                }
                if (-not $intent.cloudGuid) { $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force }
                Write-PcmRoomListIdentityPin -Intent $intent -Group $group
                Set-PcmRoomListConfiguration -List $list -Intent $intent -Group $group
                $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $false -Force
                $group = Get-PcmResourceCloudRoomList -List $list -Intent $intent
            }
            if (-not $group) { throw 'Room list is not visible/verified; dependent rooms are not created.' }
            Write-PcmRoomListIdentityPin -Intent $intent -Group $group
            if ($wasDone) { $intent.status = 'Done' }
            $null = Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop
            Write-PcmRoomListJournal -Intent $intent -Step 'RoomListReady' -Status Done -Detail "$($intent.operation) RoomList $($list.PrimarySmtpAddress), cloud GUID $($group.Guid); members added only after room provisioning"
            Write-PcmItem -Context $context -Status Ok -Text "Room list ready: $($list.PrimarySmtpAddress) ($($intent.operation))"
        } catch {
            $message = $_.Exception.Message
            if ($intent) { Write-PcmRoomListJournal -Intent $intent -Step 'ConvertRoomList' -Status Failed -Detail $message }
            Write-PcmLog -Context $context -Message ("Room list {0} failed; dependent rooms are blocked: {1}" -f $list.PrimarySmtpAddress, $message) -Level Error
        }
    }
    $context.CurrentIdentity = ''
}

function Invoke-PcmConvertResourceWave {
    param([object[]]$Entries, [object[]]$NativeEntries = @(), [object[]]$NativeRecords = @())
    $intents = @{}
    foreach ($native in $NativeEntries) { $intents[[string]$native.Record.object_guid] = $native.Intent }
    foreach ($entry in $Entries) {
        $context.CurrentIdentity = $entry.Row.Identity
        $provisioningId = ''
        try {
            $intent = $entry.Intent
            if (-not $intent -or $intent.lifecycle -ne 'CloudResource') { throw 'Resource creation intent is missing; do not guess an existing cloud identity.' }
            $protected = (Get-PcmResourceProtectionMode $intent) -eq 'CaseHold'
            $owned = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
            if ($owned) {
                $provisioningId = $owned.Id
                Set-PcmResourceProvisioningDisabled -Identity $provisioningId
            }
            if (-not $protected) { Assert-PcmResourceOrganizationUnheld }
            Test-PcmResourceConfigurationReadiness -Record $entry.Record -FutureResources $NativeRecords -ResourceIntents $intents -RoomListIntents $context.RoomListIntents
            if (-not $owned) {
                if ($intent.cloudId -or $intent.cloudGuid) { throw 'Journal-owned resource is missing; use Recover, never replace its identity.' }
                if (Get-PcmValue $intent 'creationAttempted' $false) { throw 'Resource creation was already attempted but its exact object is not visible; never issue a second create.' }
                $collisions = @(Get-PcmResourceCollisions -Record $entry.Record -Inventory (Get-PcmResourceInventory -Records @($entry.Record)))
                if ($collisions.Count) { throw ($collisions -join '; ') }
                if ($protected) {
                    $skus = @((Invoke-PcmGraph -Uri 'v1.0/subscribedSkus').value)
                    $null = Get-PcmResourceGroup -Id $intent.licensingGroupId -Skus $skus -NewUnits 1
                }
                # Persist BEFORE New-Mailbox. The creation Name contains the journal's unpredictable token.
                Confirm-PcmWrite -Target ([string]$entry.Record.primary_smtp_address) -Operation ('Exchange Online: create the cloud {0} mailbox' -f ([string]$entry.Record.kind).ToLowerInvariant())
                $intent | Add-Member -NotePropertyName creationAttempted -NotePropertyValue $true -Force
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceCreationIntent' -Detail "create $($intent.creationName) ($($intent.cloudType)); ownership token recorded before creation"
                $parameters = @{ Name = [string]$intent.creationName; Alias = [string](Get-PcmValue $entry.Record 'alias' (($entry.Record.primary_smtp_address -split '@')[0]))
                    DisplayName = [string](Get-PcmValue $entry.Record 'display_name' $entry.Row.Identity); PrimarySmtpAddress = [string]$entry.Record.primary_smtp_address; ErrorAction = 'Stop' }
                if (-not $parameters.Alias) { $parameters.Alias = ($entry.Record.primary_smtp_address -split '@')[0] }
                if (-not $parameters.DisplayName) { $parameters.DisplayName = $entry.Row.Identity }
                $parameters[[string]$entry.Record.kind] = $true
                $created = @(New-Mailbox @parameters)
                if ($created.Count -ne 1) { throw 'Resource creation response is missing/ambiguous; never repeat creation.' }
                $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force
                Write-PcmResourceIdentityPin -Entry $entry -Mailbox $created[0]
                $provisioningId = [string]$intent.cloudId
                Set-PcmResourceProvisioningDisabled -Identity $provisioningId
                $owned = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
                if (-not $owned) { throw 'Resource creation accepted but mailbox not visible yet; resume the same batch, do not create by another name.' }
            }
            if (-not $intent.cloudId -or -not $intent.cloudGuid) { $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force }
            Write-PcmResourceIdentityPin -Entry $entry -Mailbox $owned.Mailbox
            if (-not $protected) { Assert-PcmUnprotectedResource -Owned $owned }
            $identitySettings = @{ accountEnabled = $false }
            if ($protected) { $identitySettings.usageLocation = [string]$context.Config.Cloud.DefaultUsageLocation }
            $null = Invoke-PcmGraph -Method PATCH -Uri ("v1.0/users/{0}" -f $owned.Id) -Body $identitySettings
            $disabled = Wait-PcmCondition -TimeoutSeconds (Get-PcmWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                $now = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
                if (-not $now) { throw 'Pinned resource disappeared while verifying disabled sign-in; nothing is adopted.' }
                return ($now.User.accountEnabled -eq $false)
            }
            if (-not $disabled.Ok) { throw 'Resource sign-in disable verification timed out; IDs retained, resume the same batch.' }
            $marker = @{ Identity = $owned.Id; ErrorAction = 'Stop' }; $marker[[string]$intent.markerAttribute] = [string]$intent.token
            Confirm-PcmWrite -Target ([string]$owned.Id) -Operation 'Exchange Online: stamp the ownership marker on the resource mailbox'
            Set-Mailbox @marker
            $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $false -Force
            $null = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
            Write-PcmResourceJournal -Entry $entry -Step 'ResourceOwned' -Detail "owned cloud-only $($intent.cloudType), ID $($owned.Id), ExchangeGuid $($owned.Guid), sign-in disabled"
            if ($protected) {
            $skus = @((Invoke-PcmGraph -Uri 'v1.0/subscribedSkus').value)
            $member = Test-PcmGroupMember -GroupId $intent.licensingGroupId -UserId $owned.Id
            $group = Get-PcmResourceGroup -Id $intent.licensingGroupId -Skus $skus -NewUnits $(if ($member) { 0 } else { 1 })
            if (-not $member) { Set-PcmGroupMember -GroupId $group.Id -UserId $owned.Id -Operation Add }
            $licensed = Wait-PcmCondition -TimeoutSeconds (Get-PcmWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                $now = Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id,assignedPlans,licenseAssignmentStates" -f $owned.Id)
                Test-PcmResourceLicence -User $now -Group $group
            }
            if (-not $licensed.Ok) {
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceLicencePending' -Detail 'Enterprise group plan not enabled yet; licence membership and data retained. Resume this batch.'
                Set-PcmRowResult -Row $entry.Row -Status Pending -Detail 'Resource enterprise licensing is not enabled yet; data retained, resume this Convert batch.'
                continue
            }
            }
            $calendar = Set-PcmResourceConfiguration -Record $entry.Record -Owned $owned -Intent $intent
            $verified = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
            if ($verified.User.accountEnabled -ne $false) { throw 'Final resource disabled sign-in verification failed.' }
            if ($protected) {
                if (-not (Test-PcmResourceLicence -User $verified.User -Group $group)) { throw 'Final resource enterprise licensing verification failed.' }
            } else { Assert-PcmUnprotectedResource -Owned $verified }
            $collisions = @(Get-PcmResourceCollisions -Record $entry.Record -Inventory (Get-PcmResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId)
            if ($collisions.Count) { throw ($collisions -join '; ') }
            Set-PcmResourceRoomListMembership -Record $entry.Record -Owned $verified -RoomListIntents $context.RoomListIntents
            Enable-PcmResourceBooking -Record $entry.Record -Owned $verified -Calendar $calendar
            $protectionText = if ($protected) { 'legacy enterprise licence retained throughout active lifetime' } else { 'no dedicated licence or hold; Recover has no inactive copy and does NOT preserve cloud reservations/messages' }
            Write-PcmResourceJournal -Entry $entry -Step 'ResourceReady' -Status Done -Detail "Native resource configured; $protectionText."
            Set-PcmRowResult -Row $entry.Row -Status Success -Field ExchangeOnline -Value $intent.cloudType -Detail "native $($intent.cloudType) $($owned.Guid), disabled sign-in, descriptive properties/room lists and booking restrictions/work hours restored, $protectionText; no general mailbox permissions or old on-premises content restored"
        } catch {
            $message = $_.Exception.Message
            if ($provisioningId) {
                try {
                    Set-PcmResourceProvisioningDisabled -Identity $provisioningId
                    $message += ' | Owned resource booking disabled and hidden; resume the same batch after resolving the failure.'
                } catch { $message += " | FAIL-CLOSED CLEANUP FAILED: $($_.Exception.Message); verify the pinned resource booking and visibility immediately." }
            }
            Stop-PcmEntry -Entry $entry -Step 'ConvertResource' -Message $message
        }
    }
    $context.CurrentIdentity = ''
}

function Find-PcmConvertedBatch {
    <#
    The Convert batch that converted an object and that no Recover batch rolled back for it yet, or $null. The journal is read once per run (one query for every object, the latest
    batch of each object kept) and the answers come from that index.
    #>
    param([Parameter(Mandatory)][string]$ObjectGuid)
    if ($null -eq (Get-PcmValue $context 'ConvertedIndex' $null)) {
        $index = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($row in @(Invoke-PcmSql $context.Journal @"
SELECT ci.object_guid AS object_guid, cb.id AS id, cb.created_utc AS created_utc FROM batch_item ci JOIN batch cb ON cb.id = ci.batch_id
WHERE cb.action = 'Convert' AND ci.step IS NOT NULL AND ci.step <> '' AND ci.status <> 'Skipped'
  AND NOT EXISTS (SELECT 1 FROM batch_item ri JOIN batch rb ON rb.id = ri.batch_id
                  WHERE rb.action = 'Recover' AND rb.convert_batch = cb.id COLLATE NOCASE AND ri.object_guid = ci.object_guid COLLATE NOCASE AND ri.status = 'Done')
ORDER BY cb.created_utc ASC
"@ -As Rows)) { $index[[string]$row.object_guid] = $row }
        $context.ConvertedIndex = $index
    }
    $found = $null
    if ($context.ConvertedIndex.TryGetValue($ObjectGuid, [ref]$found)) { return $found }
    return $null
}

function Assert-PcmNotAlreadyConverted {
    <# An object asked for by name that an earlier Convert batch converted (and Recover did not roll back) is not converted again: the run stops and names that batch. #>
    param([AllowEmptyCollection()][object[]]$Records = @())
    foreach ($record in $Records) {
        $converted = Find-PcmConvertedBatch -ObjectGuid ([string]$record.object_guid)
        if (-not $converted) { continue }
        $name = if ($record.user_principal_name) { [string]$record.user_principal_name } else { [string]$record.primary_smtp_address }
        $context.NoChangeMade = $true; $context.NextOnFailure = $true; $context.CurrentIdentity = ''
        $context.NextSteps = @('To continue that batch (what is not finished is resumed, nothing is converted twice):', (Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -BatchId ([string]$converted.id) -KeepFilters),
            'When AD, Exchange and Entra Connect are rebuilt, to roll it back:', (Get-PcmCommandLine -ActionName 'Recover' -BatchId ([string]$converted.id)))
        throw ('{0} is already converted by Convert batch {1}: nothing was changed.' -f $name, $converted.id)
    }
}

function Invoke-PcmConvert {
    $config = $context.Config
    $dot = [char]0x00B7
    $context.Journal = Open-PcmStore -Path $config.Store.JournalPath -Root $context.Root -Kind Journal
    $resume = $null; $snapshotId = $Snapshot; $resumeItems = @{}
    $read = @{ Verb = 'convert' }
    if ($Batch) {
        $context.CurrentOperation = 'Read-Batch'
        $resume = Get-PcmBatch $context.Journal -Id $Batch
        if (-not $resume -or $resume.action -ne 'Convert') { throw "Convert batch $Batch not found in $($config.Store.JournalPath)." }
        $rollbacks = @(Invoke-PcmSql $context.Journal "SELECT id, status FROM batch WHERE action = 'Recover' AND convert_batch = @c COLLATE NOCASE ORDER BY created_utc" @{ c = [string]$resume.id } -As Rows)
        if ($rollbacks.Count) {
            $context.NoChangeMade = $true; $context.NextOnFailure = $true; $context.CurrentIdentity = ''
            $context.NextSteps = @(('To convert the objects that Convert batch {0} did not finish, run a new Convert for them (Preview first, then -Mode Apply), choosing them with -Identity, -CsvPath or -Group:' -f $resume.id), (Get-PcmCommandLine -ActionName 'Convert'))
            throw ('Convert batch {0} cannot be resumed: Recover batch {1} ({2}) already rolled back objects of it, and an object converted now could not be rolled back with it. Nothing was changed.' -f $resume.id, $rollbacks[-1].id, ([string]$rollbacks[-1].status).ToLowerInvariant())
        }
        $snapshotId = [long]$resume.snapshot_id
        $read.Batch = $resume
        $context.BatchId = [string]$resume.id
        foreach ($item in @(Get-PcmBatchItem $context.Journal $resume.id)) { $resumeItems[[string]$item.object_guid] = $item }
        $read.ObjectGuid = [string[]]@($resumeItems.Values | Where-Object { $_.status -notin @('Done','Skipped') } | ForEach-Object { [string]$_.object_guid })
        $read.Where = "the objects of Convert batch $($resume.id) still to resume"
        $finished = @{}
        foreach ($item in @($resumeItems.Values | Where-Object { $_.status -in @('Done','Skipped') })) { $finished[[string]$item.object_guid] = ('already finished in Convert batch {0}' -f $resume.id) }
        $read.Known = $finished
    } else { $read.SkipConverted = $true }
    $data = Read-PcmSnapshotData -Id $snapshotId @read
    $records = @($data.Records); $permissions = @($data.Permissions); $members = @($data.Members)
    if (-not $resume -and @(Get-PcmIdentityList).Count) { Assert-PcmNotAlreadyConverted -Records $records }
    $context.ResourceJournal = $resumeItems
    $context.RoomListIntents = @{}
    if ($resume) {
        foreach ($change in @(Get-PcmTenantChange $context.Journal $resume.id | Where-Object kind -eq 'NativeRoomList')) {
            $intent = [string]$change.original_json | ConvertFrom-PcmJson
            if ([string]$intent.lifecycle -ne 'CloudRoomList' -or [string]$intent.sourceGuid -ne [string]$change.target_id) { throw 'Invalid room-list creation journal.' }
            $context.RoomListIntents[[string]$change.target_id] = $intent
        }
    }
    $context.NativeResourceScope = @($records | Where-Object {
            (Test-PcmNativeResource -Config $config -Record $_) -or (Get-PcmResourceIntent -Item $resumeItems[[string]$_.object_guid])
        }).Count -gt 0
    $nativeOnly = $records.Count -gt 0 -and @($records | Where-Object {
            $intent = Get-PcmResourceIntent -Item $resumeItems[[string]$_.object_guid]
            -not ((Test-PcmNativeResource -Config $config -Record $_) -or $intent) -or (Get-PcmResourceProtectionMode $intent) -ne 'None'
        }).Count -eq 0
    if ($resume) { Write-PcmItem -Context $context -Status Info -Icon Batch -Text ('Resuming batch {0}: {1} object(s) not finished' -f $resume.id, $records.Count) }

    Write-PcmStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    $how = Connect-PcmCloudSession
    Write-PcmItem -Context $context -Status Ok -Text $how

    Write-PcmStep -Context $context -Title 'Readiness' -Icon Key
    $context.CurrentOperation = 'Readiness'
    $skipUserGroup = @($records | Where-Object { $_.kind -eq 'User' }).Count -eq 0 -and @($records | Where-Object {
            (Test-PcmNativeResource -Config $config -Record $_) -or (Get-PcmResourceIntent -Item $resumeItems[[string]$_.object_guid])
        }).Count -gt 0
    $tenant = Get-PcmTenantFact -Context $context -SkipUserGroup:$skipUserGroup -NativeOnly:$nativeOnly
    $states = Read-PcmCloudStates -Records $records
    $lookup = New-PcmTrusteeLookup -States $states -Records $records
    $permissionIndex = Get-PcmPermissionIndex -Permissions $permissions -Members $members
    # Read-only first: the state of every object, then the licence units really needed, then the decision.
    $assessed = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        if (Test-PcmStop) { throw 'Stopped by the operator during the readiness check: nothing was changed.' }
        $context.CurrentIdentity = if ($record.user_principal_name) { [string]$record.user_principal_name } else { [string]$record.primary_smtp_address }
        $subset = Get-PcmPermissionSubset -Index $permissionIndex -MailboxGuid ([string]$record.object_guid)
        $assessment = Get-PcmObjectAssessment -Record $record -Tenant $tenant -Permissions $subset.Permissions -Members $subset.Members -Lookup $lookup -State $states[[string]$record.object_guid]
        $step = if ($resume -and $resumeItems.ContainsKey([string]$record.object_guid)) { [string](Get-PcmValue $resumeItems[[string]$record.object_guid] 'step' '') } else { '' }
        [void]$assessed.Add([pscustomobject]@{ Record = $record; Assessment = $assessment; State = $assessment.State; Step = $step })
    }
    $context.CurrentIdentity = ''
    $need = Get-PcmLicenceNeed -Entries @($assessed) -Tenant $tenant
    $tenantFindings = @(Test-PcmTenantReadiness -Config $config -Roles @(Get-PcmGraphRole) -Tenant $tenant -UserCount $need.Users -SharedCount $need.Shared -ResourceCount $need.Resources -SkipUserLicensing:$need.SkipUserLicensing -NativeOnly:$nativeOnly -SignIn $config.Cloud.SignIn)
    $tenantFindings += @(Get-PcmResourceLicenceFinding -Need $need -Tenant $tenant)
    foreach ($finding in $tenantFindings | Where-Object { $_.Level -eq 'Error' }) { Write-PcmLog -Context $context -Message $finding.Message -Level Error }
    if (@($tenantFindings | Where-Object Level -eq 'Error').Count) { throw 'Tenant prerequisites not met (see above): nothing was changed.' }
    # The shared mailboxes hold their temporary unit one wave at a time, with the units that are left after the users of this run take theirs (and after KeepFree).
    $sharedUnits = if ($need.Shared) { Get-PcmSharedUnitCount -Config $config -Tenant $tenant -UserCount $(if ($need.SkipUserLicensing) { 0 } else { [int]$need.Users }) } else { $null }
    $sharedAtOnce = if ($sharedUnits) { [Math]::Min([int]$sharedUnits.Wave, [int]$need.Shared) } else { [Math]::Min([Math]::Max(1, [int]$config.Licensing.Shared.Parallel), [int]$need.Shared) }
    Write-PcmItem -Context $context -Status Ok -Text ('Tenant prerequisites met (permissions, licences: {0} new unit(s) for users{1})' -f $need.Users, $(if ($need.Shared) { ', up to {0} at a time for the temporary licence of the shared mailboxes' -f $sharedAtOnce } else { '' }))
    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $assessed) {
        $record = $entry.Record; $assessment = $entry.Assessment
        $row = New-PcmRow $record
        [void]$context.Rows.Add($row)
        $context.CurrentIdentity = $row.Identity
        $blocking = @($assessment.Errors) + @($assessment.Findings | Where-Object { $_.Code -eq 'EXCHANGE_PLAN_PRESENT' -and $_.Level -ne 'Error' })
        if ($resume) { $blocking = @($blocking | Where-Object { $_.Code -notin @('ALREADY_CLOUD_MAILBOX','EXCHANGE_PLAN_PRESENT') }) }
        $row.Warnings = (@($assessment.Warnings | ForEach-Object { $_.Message })) -join ' | '
        if ($blocking.Count) { Set-PcmRowResult -Row $row -Status Error -Detail ('Not converted: ' + (($blocking | ForEach-Object { $_.Message }) -join ' | ')); continue }
        [void]$plan.Add([pscustomobject]@{ Record = $record; Row = $row; EntraId = $(if ($assessment.State.User) { [string]$assessment.State.User.id } else { '' }); State = $assessment.State
                Intent = $(if ($resumeItems.ContainsKey([string]$record.object_guid)) { Get-PcmResourceIntent -Item $resumeItems[[string]$record.object_guid] } else { $null }) })
    }
    $context.CurrentIdentity = ''
    $planUsers = @($plan | Where-Object { $_.Record.kind -eq 'User' })
    $planResources = @($plan | Where-Object { Get-PcmValue $_.State 'NativeResource' $false })
    $planShared = @($plan | Where-Object { $_.Record.kind -ne 'User' -and $_ -notin $planResources })
    $roomLists = @(Get-PcmRoomListPlan -Entries $planResources -Records $records)
    Write-PcmStepCount -Context $context -Titles @('Users: source of authority and Exchange plan', 'Users: cloud mailboxes') -Count $planUsers.Count
    Write-PcmStepCount -Context $context -Titles @('Shared mailboxes') -Count ($planShared.Count + $planResources.Count)
    Write-PcmItem -Context $context -Status Info -Icon Target -Text ('{0} object(s) ready: {1} user(s), {2} shared, {5} native resources {3} {4} blocked' -f $plan.Count, $planUsers.Count, $planShared.Count, $dot, ($records.Count - $plan.Count), $planResources.Count)

    Write-PcmStep -Context $context -Title 'Plan and confirmation' -Icon Plan
    $userSkuName = Get-PcmSkuSpelling -Name ([string]$config.Licensing.Users.SkuPartNumber) -Tenant $tenant
    $sharedSkuName = Get-PcmSkuSpelling -Name ([string]$config.Licensing.Shared.SkuPartNumber) -Tenant $tenant
    $modeText = @{ Group = "licence group $($config.Licensing.Users.GroupId)"; Kiosk = 'Exchange Kiosk plan of the user licence'; Direct = "$userSkuName (Exchange only)" }[$config.Licensing.Users.Mode]
    Write-PcmItem -Context $context -Status Info -Text ("Users: ExchangeGuid cleared, source of authority to the cloud, usageLocation, Exchange plan by $modeText; the Teams storage becomes the mailbox")
    if ($planShared.Count) { Write-PcmItem -Context $context -Status Info -Text ("Shared mailboxes, in waves of up to $($config.Licensing.Shared.Parallel) (one temporary unit each$(if ([int](Get-PcmValue $config.Licensing.Shared 'KeepFree' 0) -gt 0) { ", $($config.Licensing.Shared.KeepFree) unit(s) kept free" })): temporary $sharedSkuName, shared, $($config.Retention.TagAttribute) = $($config.Retention.TagValue), licence given back, permissions from the snapshot") }
    foreach ($entry in $planResources) {
        $protectionText = if ((Get-PcmResourceProtectionMode $entry.Intent) -eq 'CaseHold') { 'legacy enterprise licence/case-hold lifecycle retained' } else { 'no dedicated licence or hold (standard resources up to 50 GB); Recover has no inactive copy, cloud reservations/messages NOT preserved' }
        Write-PcmItem -Context $context -Status Info -Text ("Native $($entry.Record.kind): New-Mailbox, disabled cloud-only identity, independent GUID, $protectionText; addresses: " + ((Get-PcmResourceAddresses $entry.Record) -join ', ') + '; booking/work hours restored, no on-premises content restored.')
    }
    foreach ($list in $roomLists) { Write-PcmItem -Context $context -Status Info -Text "RoomList $($list.PrimarySmtpAddress): create if missing, or reuse verified cloud-only list; before its rooms, then add only selected room members." }
    $convertEstimate = Get-PcmDurationEstimate -UserWaves ([Math]::Ceiling($planUsers.Count / [Math]::Max(1, [int](Get-PcmValue $context 'UserWaveSize' 500)))) -SharedRounds ([Math]::Ceiling($planShared.Count / [Math]::Max(1, [int]$config.Licensing.Shared.Parallel)))
    if ($convertEstimate.Text) { Write-PcmItem -Context $context -Status Info -Icon Clock -Text ('Rough estimate: about {0} (typical tenant; a slow Microsoft 365 evening can take much longer)' -f $convertEstimate.Text) }
    foreach ($entry in $plan) { $entry.Row.FinalStatus = 'Planned'; Write-PcmLog -Context $context -Message ('Planned: {0} ({1})' -f $entry.Row.Identity, $entry.Record.kind) -Level Sub }
    if ($context.Mode -ne 'Apply') {
        Skip-PcmStep -Title 'Users: source of authority and Exchange plan' -Reason 'Preview: not run.' -Icon Write
        Skip-PcmStep -Title 'Users: cloud mailboxes' -Reason 'Preview: not run.' -Icon Clock
        Skip-PcmStep -Title 'Shared mailboxes' -Reason 'Preview: not run.' -Icon Mail
        $context.NextSteps = @('To make the changes, run the same command in Apply mode (the confirmation is asked once):', (Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -BatchId $(if ($resume) { [string]$resume.id } else { '' }) -KeepFilters))
        return
    }
    if (-not $plan.Count) { throw 'No object can be converted (see the reasons above).' }
    Confirm-PcmApply ('Convert {0} object(s) to Exchange Online ({1} user(s), {2} shared, {3} native resources)?' -f $plan.Count, $planUsers.Count, $planShared.Count, $planResources.Count)
    if (-not $resume) {
        $context.BatchId = New-PcmBatch $context.Journal Convert @{ SnapshotId = [long]$data.Snapshot.id; SnapshotRunId = [string]$data.Snapshot.run_id; SnapshotHash = [string]$context.SnapshotHash; Environment = $config.Environment; Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            ToolVersion = $toolVersion; ScopeJson = (@{ Identity = $Identity; CsvPath = $CsvPath; SearchBase = $SearchBase; Group = $Group; MaxObjects = $MaxObjects; Identities = @(Get-PcmIdentityList).Count; Scope = $Scope; UsersMode = $config.Licensing.Users.Mode } | ConvertTo-Json -Compress) }
        foreach ($entry in $plan) {
            if ($entry -in $planResources) {
                $token = [guid]::NewGuid().ToString('N')
                $attribute = if ($config.Retention.TagAttribute -eq 'CustomAttribute15') { 'CustomAttribute14' } else { 'CustomAttribute15' }
                $entry.Intent = [pscustomobject]@{ lifecycle = 'CloudResource'; protectionMode = 'None'; token = $token; markerAttribute = $attribute
                    creationName = ('PRA-R-{0}-{1}-{2}' -f $context.BatchId.Substring(0,8), ([string]$entry.Record.object_guid).Substring(0,8), $token)
                    creationAttempted = $false; cloudId = ''; cloudGuid = ''; cloudType = "$($entry.Record.kind)Mailbox" }
                Set-PcmBatchItem $context.Journal $context.BatchId ([string]$entry.Record.object_guid) @{ kind = [string]$entry.Record.kind; identity = $entry.Row.Identity
                    status = 'Planned'; step = 'ResourceIntent'; original_json = ($entry.Intent | ConvertTo-Json -Depth 15 -Compress) }
                continue
            }
            $original = @{ entraId = $entry.EntraId; isCloudManaged = $entry.State.IsCloudManaged; exchangeGuid = [string]$entry.Record.exchange_guid
                usageLocation = [string]$entry.State.User.usageLocation; recipientType = $entry.State.RecipientType; tag = $entry.State.TagValue; holds = @($entry.State.Holds)
                licences = @(Get-PcmDirectLicence $entry.State.User) }
            Set-PcmBatchItem $context.Journal $context.BatchId ([string]$entry.Record.object_guid) @{ kind = [string]$entry.Record.kind; identity = $entry.Row.Identity
                entra_id = $entry.EntraId; status = 'Planned'; original_json = ($original | ConvertTo-Json -Compress -Depth 6) }
        }
    }
    $originals = @{}; foreach ($item in @(Get-PcmBatchItem $context.Journal $context.BatchId)) { $originals[[string]$item.object_guid] = @(Get-PcmOriginalLicence $item) }
    Write-PcmItem -Context $context -Status Ok -Icon Batch -Text ("Batch {0} in {1}" -f $context.BatchId, $config.Store.JournalPath)

    Write-PcmStep -Context $context -Title 'Users: source of authority and Exchange plan' -Icon Write
    $started = [System.Collections.Generic.List[object]]::new()
    if ($planUsers.Count -and $config.Licensing.Users.Mode -eq 'Group' -and $tenant.LicenceGroup -and $tenant.LicenceGroup.onPremisesSyncEnabled -and -not $tenant.LicenceGroupCloudManaged) {
        $context.CurrentOperation = 'Group-SOA'
        Set-PcmSourceOfAuthority -Id $config.Licensing.Users.GroupId -Cloud $true -Type groups
        Set-PcmTenantChange $context.Journal $context.BatchId 'GroupSoa' $config.Licensing.Users.GroupId -OriginalJson '{"isCloudManaged":false}'
        Write-PcmJournal -Step 'GroupSoaCloud' -Detail "licence group $($tenant.LicenceGroup.displayName) managed in the cloud"
        Write-PcmItem -Context $context -Status Ok -Text "Licence group $($tenant.LicenceGroup.displayName) now managed in the cloud (given back to AD by Recover)"
    }
    $stopped = $false
    # Waves of 500 users, phase by phase (a stop is taken between two waves); their mailboxes are awaited together below.
    $waveSize = [Math]::Max(1, [int](Get-PcmValue $context 'UserWaveSize' 500))
    $userClock = [Diagnostics.Stopwatch]::StartNew()
    for ($start = 0; $start -lt $planUsers.Count; $start += $waveSize) {
        $wave = @($planUsers[$start..([Math]::Min($planUsers.Count, $start + $waveSize) - 1)])
        if ($stopped -or (Test-PcmStop)) { $stopped = $true; foreach ($entry in $wave) { Set-PcmRowStopped $entry.Row }; continue }
        $context.CurrentOperation = 'Convert-User'
        if ($planUsers.Count -gt $waveSize) { Write-PcmProgress -Context $context -Phase 'Users' -Done $start -Total $planUsers.Count -ElapsedSeconds $userClock.Elapsed.TotalSeconds }
        foreach ($entry in @(Invoke-PcmUserPlanSet -Entries $wave -Tenant $tenant)) { [void]$started.Add($entry) }
    }
    if (-not $planUsers.Count) { Write-PcmItem -Context $context -Status Skip -Text 'No user in this batch.' }

    Write-PcmStep -Context $context -Title 'Users: cloud mailboxes' -Icon Clock
    $context.CurrentIdentity = ''; $context.CurrentOperation = 'Wait-Mailbox'
    # Every user waiting is polled together (filters of 50 objects): thousands of users per round.
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $waiting = @(Wait-PcmExoStateSet -Phase 'Waiting for the cloud mailboxes of the users' -Entries @($started) -TimeoutSeconds (Get-PcmWaitSeconds 'MailboxTimeoutMinutes') -StopOnRequest -Test { param($state) $state.Type -eq 'UserMailbox' } -OnReady {
            param($entry, $state, $seconds)
            Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'MailboxReady' -Detail ("UserMailbox {0} after {1}" -f $state.ExchangeGuid, (Format-PcmDuration $seconds)) -Item @{ step = 'MailboxReady'; status = 'Done'; result_json = (@{ cloudGuid = $state.ExchangeGuid } | ConvertTo-Json -Compress) }
            Set-PcmRowResult -Row $entry.Row -Status Success -Detail ('cloud mailbox {0} (Teams storage promoted when present)' -f $state.ExchangeGuid) -Field 'ExchangeOnline' -Value 'UserMailbox'
        })
    foreach ($entry in $waiting) { Set-PcmRowResult -Row $entry.Row -Status Pending -Detail ('no cloud mailbox yet after {0}: run Convert -Batch {1} again later' -f (Format-PcmDuration $clock.Elapsed.TotalSeconds), $context.BatchId) }
    if (-not $started.Count) { Write-PcmItem -Context $context -Status Skip -Text 'Nothing to wait for.' }

    Write-PcmStep -Context $context -Title 'Shared mailboxes' -Icon Mail
    $sharedSku = $tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
    # Waves of Licensing.Shared.Parallel shared mailboxes, one temporary unit each, never more than the free units minus
    # Licensing.Shared.KeepFree, read again before each wave: a shared mailbox that failed keeps its temporary licence,
    # so no object is started without a unit (it would stop half-way,
    # GUID cleared, source of authority in the cloud). An object already a SharedMailbox, or already holding the unit
    # (resume), needs none. A stop is taken between two waves.
    $parallel = [Math]::Max(1, [int]$config.Licensing.Shared.Parallel)
    $keepFree = [Math]::Max(0, [int](Get-PcmValue $config.Licensing.Shared 'KeepFree' 0))
    $queue = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $planShared) { $queue.Add($entry) }
    $wavesRun = 0; $unitWaited = $false
    $sharedClock = [Diagnostics.Stopwatch]::StartNew()
    $readUnits = {
        $freeUnits = 0
        if ($sharedSku) {
            $now = @((Invoke-PcmGraph -Uri 'v1.0/subscribedSkus').value) | Where-Object { [string]$_.skuId -eq [string]$sharedSku.skuId } | Select-Object -First 1
            if ($now) { $freeUnits = Get-PcmFreeUnit $now }
        }
        $freeUnits
    }
    while ($queue.Count) {
        if ($stopped -or (Test-PcmStop)) { $stopped = $true; foreach ($entry in $queue) { Set-PcmRowStopped $entry.Row }; $queue.Clear(); break }
        $context.CurrentOperation = 'Convert-Shared'
        $candidates = @($queue.GetRange(0, [Math]::Min($parallel, $queue.Count)))
        $exo = Get-PcmExoStateSet -Id ([string[]]@($candidates | ForEach-Object { $_.EntraId })) -TagAttribute $config.Retention.TagAttribute
        $freeNow = & $readUnits
        $needsUnit = @{}
        foreach ($entry in $candidates) {
            $isShared = $exo[$entry.EntraId] -and $exo[$entry.EntraId].Type -eq 'SharedMailbox'
            $holds = $sharedSku -and @(Get-PcmDirectLicence $entry.State.User | Where-Object { $_.skuId -eq [string]$sharedSku.skuId }).Count
            $needsUnit[[string]$entry.EntraId] = -not $isShared -and -not $holds
        }
        $pick = Select-PcmSharedWave -Candidates $candidates -NeedsUnit $needsUnit -Free ($freeNow - $keepFree)
        if (-not $pick.Wave.Count) {
            # The units that an earlier wave of this run gave back can take a moment to show in the tenant: wait for them once, with the polling settings, before giving up.
            if ($wavesRun -gt 0 -and -not $unitWaited) {
                $unitWaited = $true
                $context.CurrentOperation = 'Wait-Unit'
                $null = Wait-PcmCondition -Phase 'Waiting for a licence unit to be free' -TimeoutSeconds ([Math]::Min(300, (Get-PcmWaitSeconds 'MailboxTimeoutMinutes'))) -IntervalSeconds ([int]$config.Polling.IntervalSeconds) -Test { ((& $readUnits) - $keepFree) -ge 1 }
                continue
            }
            foreach ($entry in $pick.Waiting) {
                $message = if ($keepFree -gt 0) { "not started: $($sharedSku.skuPartNumber) has $freeNow free unit(s) and $keepFree must stay free (Licensing.Shared.KeepFree), so none can be used for the temporary licence; nothing was changed on this object. Free a unit or lower Licensing.Shared.KeepFree, then run Convert -Batch $($context.BatchId) again." }
                else { "not started: no free unit of $($sharedSku.skuPartNumber) (a shared mailbox that failed keeps its temporary licence until it is finished); nothing was changed on this object. Free a unit, then run Convert -Batch $($context.BatchId) again." }
                Write-PcmJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'NoFreeUnit' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
                Set-PcmRowResult -Row $entry.Row -Status Error -Detail $message
                [void]$queue.Remove($entry)
            }
            continue
        }
        Write-PcmProgress -Context $context -Phase 'Shared mailboxes' -Done ($planShared.Count - $queue.Count) -Total $planShared.Count -ElapsedSeconds $sharedClock.Elapsed.TotalSeconds
        foreach ($entry in $pick.Wave) { [void]$queue.Remove($entry) }
        $wavesRun++
        Invoke-PcmConvertSharedWave -Entries $pick.Wave -Tenant $tenant -PermissionIndex $permissionIndex -Originals $originals
    }
    if (-not $planShared.Count) { Write-PcmItem -Context $context -Status Skip -Text 'No shared mailbox in this batch.' }
    if (-not $stopped -and -not (Test-PcmStop) -and $roomLists.Count) {
        $context.CurrentOperation = 'Convert-RoomLists'
        Invoke-PcmConvertRoomLists -Lists $roomLists -Unprotected:(@($planResources | Where-Object { (Get-PcmResourceProtectionMode $_.Intent) -eq 'None' }).Count -gt 0)
    }
    $resourceParallel = [Math]::Max(1, [int](Get-PcmValue (Get-PcmValue $config 'Resources' @{}) 'Parallel' 100))
    for ($start = 0; $start -lt $planResources.Count; $start += $resourceParallel) {
        $wave = @($planResources[$start..([Math]::Min($planResources.Count, $start + $resourceParallel) - 1)])
        if ($stopped -or (Test-PcmStop)) { $stopped = $true; foreach ($entry in $wave) { Set-PcmRowStopped $entry.Row }; continue }
        Invoke-PcmConvertResourceWave -Entries $wave -NativeEntries $planResources -NativeRecords @($records | Where-Object { Get-PcmValue $states[[string]$_.object_guid] 'NativeResource' $false })
    }
    $context.CurrentIdentity = ''
    $items = @(Get-PcmBatchItem $context.Journal $context.BatchId)
    $counts = @{ done = @($items | Where-Object status -eq 'Done').Count; failed = @($items | Where-Object status -eq 'Failed').Count; running = @($items | Where-Object status -in @('Running','Planned')).Count }
    $pendingLists = @($context.RoomListIntents.Values | Where-Object status -ne 'Done').Count
    $counts.roomListsDone = @($context.RoomListIntents.Values | Where-Object status -eq 'Done').Count
    $counts.roomListsPending = $pendingLists
    Set-PcmBatch $context.Journal $context.BatchId $(if ($counts.done -eq $items.Count -and -not $pendingLists) { 'Complete' } else { 'Partial' }) $counts
    $context.NextOnFailure = $true
    $next = [System.Collections.Generic.List[string]]::new()
    if ($counts.done -ne $items.Count -or $pendingLists) {
        $next.Add('Some objects or lists are not finished. Fix what is named above, then resume them (nothing is converted twice):')
        $next.Add((Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -BatchId $context.BatchId -KeepFilters))
    }
    if ($context.LeftForLater -gt 0) {
        $next.Add(('{0} more object(s) were left for a later run (-MaxObjects). For the next wave, run the same selection again (Preview first, then -Mode Apply):' -f $context.LeftForLater))
        $next.Add((Get-PcmCommandLine -ActionName 'Convert' -KeepFilters))
    }
    $next.Add('When AD, Exchange and Entra Connect are rebuilt, roll this batch back (Preview first, then -Mode Apply):')
    $next.Add((Get-PcmCommandLine -ActionName 'Recover' -BatchId $context.BatchId))
    $context.NextSteps = $next.ToArray()
}

function Stop-PcmRecoverEntry {
    <# An object of Recover that fails in a phase: journal, report row, and it leaves the next phases. #>
    param([Parameter(Mandatory)][object]$Entry, [Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Message)
    $text = $Message -replace '\s+', ' '
    Write-PcmJournal -ObjectGuid ([string]$Entry.Item.object_guid) -Step $Step -Outcome Fail -Detail $text -Item @{ status = 'Failed'; message = $text }
    Set-PcmRowResult -Row $Entry.Row -Status Error -Detail "rollback stopped: $text"
}

function Add-PcmRecoverHold {
    <#
    Case hold for many objects: added where missing (groups of 100 locations per call, new policies of the tool when the
    ones in use are full), then ONE wait for all until a case hold of the tool is stamped (HoldTimeoutMinutes).
    Returns Stamped, Late and Failed (entries) and Seconds (object ID -> seconds).
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$HoldSet)
    $result = [pscustomobject]@{ Stamped = @(); Late = @(); Failed = @(); Seconds = @{} }
    if (-not $Entries.Count) { return $result }
    $ids = [string[]]@($Entries | ForEach-Object { $_.EntraId })
    $now = Get-PcmExoStateSet -Id $ids -TagAttribute $context.Config.Retention.TagAttribute
    $missing = [string[]]@($Entries | Where-Object { -not ($now[$_.EntraId] -and (Test-PcmToolHold -Set $HoldSet -Holds @($now[$_.EntraId].Holds))) } | ForEach-Object { $_.EntraId })
    $added = Add-PcmHoldLocationSet -Set $HoldSet -Id $missing
    $failed = @($Entries | Where-Object { $added.Failed.ContainsKey($_.EntraId) })
    foreach ($entry in $failed) { Stop-PcmRecoverEntry -Entry $entry -Step 'Hold' -Message $added.Failed[$entry.EntraId] }
    $active = @($Entries | Where-Object { -not $added.Failed.ContainsKey($_.EntraId) })
    $seconds = $result.Seconds
    $late = @(Wait-PcmExoStateSet -Phase 'Waiting for the case hold to apply' -Entries $active -TimeoutSeconds (Get-PcmWaitSeconds 'HoldTimeoutMinutes') -Test { param($state) Test-PcmToolHold -Set $HoldSet -Holds @($state.Holds) } -OnReady {
            $seconds[$args[0].EntraId] = $args[2] })
    $result.Late = $late
    $result.Stamped = @($active | Where-Object { $_ -notin $late })
    $result.Failed = $failed
    return $result
}

function Get-PcmHoldTagOf {
    <# The tag (UniH + GUID) of the policy of the tool that holds an object, or ''. #>
    param([Parameter(Mandatory)][object]$HoldSet, [Parameter(Mandatory)][string]$Id)
    $policy = $HoldSet.Policies | Where-Object { $_.Locations.Contains($Id) } | Select-Object -First 1
    if ($policy) { return $policy.Tag }
    return ''
}

function Invoke-PcmRecoverUserWave {
    <#
    A wave of users (and of shared mailboxes whose Convert stopped half-way, rolled back like users), phase by phase:
    case hold lifted where the cloud mailbox carries it (it blocks the switch back), source of authority back to AD
    (Graph batches), one delta cycle, ONE wait for the synchronisation, Exchange plan removed, ONE wait until each one
    is a MailUser with its on-premises GUID, then the case hold and ONE wait for the stamps. Objects already back
    (Path Back) only get the case hold. An object that fails is reported and left out; the wave is always finished.
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$HoldSet)
    $config = $context.Config
    $ids = { param($set) [string[]]@($set | ForEach-Object { $_.EntraId }) }
    $users = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($Entries | Where-Object { $_.Path -eq 'User' })) { $users.Add($entry) }
    $back = @($Entries | Where-Object { $_.Path -eq 'Back' })
    # 1. The case hold blocks the switch back of a cloud mailbox: lifted first (it is added again at the end).
    $lift = @($users | Where-Object { $_.Exo.Type -eq 'UserMailbox' -and (Test-PcmToolHold -Set $HoldSet -Holds @($_.Exo.Holds)) })
    if ($lift.Count) {
        $removed = Remove-PcmHoldLocationSet -Set $HoldSet -Id (& $ids $lift)
        foreach ($entry in $lift) {
            if ($removed.Failed.ContainsKey($entry.EntraId)) { Stop-PcmRecoverEntry -Entry $entry -Step 'HoldLifted' -Message $removed.Failed[$entry.EntraId]; [void]$users.Remove($entry) }
            else { Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'HoldLifted' -Detail 'case hold lifted for the rollback' }
        }
    }
    # 2. Source of authority back to AD.
    $behaviors = Get-PcmSyncBehaviorSet -UserId (& $ids $users)
    $toAd = @($users | Where-Object { $behaviors[$_.EntraId] -eq $true })
    $answers = Invoke-PcmGraphBatch -Requests @($toAd | ForEach-Object { @{ Method = 'PATCH'; Url = ('/users/{0}/onPremisesSyncBehavior' -f $_.EntraId); Body = @{ isCloudManaged = $false } } })
    for ($i = 0; $i -lt $toAd.Count; $i++) {
        if ($answers[$i].Ok) { Write-PcmJournal -ObjectGuid ([string]$toAd[$i].Item.object_guid) -Step 'SoaAd' -Detail 'isCloudManaged = false' -Item @{ step = 'SoaAd' } }
        else { Stop-PcmRecoverEntry -Entry $toAd[$i] -Step 'SoaAd' -Message ('source of authority not given back to AD ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$users.Remove($toAd[$i]) }
    }
    # 3. One delta cycle (none when every user is synchronised already), one wait for all.
    if ($users.Count) {
        $context.CurrentOperation = 'EntraConnect-Delta'
        $now = Get-PcmExoStateSet -Id (& $ids $users)
        if (@($users | Where-Object { $s = $now[$_.EntraId]; -not $s -or $s.IsDirSynced -ne $true }).Count) {
            $state = Invoke-PcmEntraConnect -Context $context -Operation Delta -Caller $caller
            Write-PcmItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect delta cycle: $state"
        } else { Write-PcmItem -Context $context -Status Info -Icon Sync -Text 'Already synchronised from AD: no delta cycle needed.' }
        $late = @(Wait-PcmExoStateSet -Phase 'Waiting for Entra Connect to bring the users back' -Entries $users.ToArray() -TimeoutSeconds (Get-PcmWaitSeconds 'SyncTimeoutMinutes') -Test { param($state) $state.IsDirSynced -eq $true } -OnReady {
                Write-PcmItem -Context $context -Status Ok -Text ('{0}: synchronised from AD again' -f $args[0].Row.Identity) })
        foreach ($entry in $late) { Set-PcmRowResult -Row $entry.Row -Status Pending -Detail 'not synchronised from AD yet: run Entra Connect, then Recover -Batch again'; [void]$users.Remove($entry) }
    }
    $context.CurrentOperation = 'Users-Plan'
    # 4. A lifted case hold must be gone from the cloud mailbox before the plan is removed.
    $lifted = @($users | Where-Object { $_ -in $lift })
    if ($lifted.Count) {
        $late = @(Wait-PcmExoStateSet -Phase 'Waiting for the case hold to be lifted' -Entries $lifted -TimeoutSeconds (Get-PcmWaitSeconds 'HoldTimeoutMinutes') -Test { param($state) -not (Test-PcmToolHold -Set $HoldSet -Holds @($state.Holds)) })
        foreach ($entry in $late) { Stop-PcmRecoverEntry -Entry $entry -Step 'RecoverUser' -Message 'the case hold is still stamped on the cloud mailbox (it would block the rollback)'; [void]$users.Remove($entry) }
    }
    # 5. Exchange plan removed: licence group by Graph batches, Kiosk / direct / temporary licence one at a time.
    $removedText = @{}
    $tenantSkus = $null
    $groupUsers = @($users | Where-Object { $_.Item.kind -eq 'User' -and $config.Licensing.Users.Mode -eq 'Group' })
    $answers = Invoke-PcmGraphBatch -Requests @($groupUsers | ForEach-Object { @{ Method = 'DELETE'; Url = ('/groups/{0}/members/{1}/$ref' -f $config.Licensing.Users.GroupId, $_.EntraId) } })
    for ($i = 0; $i -lt $groupUsers.Count; $i++) {
        if ($answers[$i].Ok -or $answers[$i].Status -eq 404) { $removedText[$groupUsers[$i].EntraId] = 'removed from the licence group' }
        else { Stop-PcmRecoverEntry -Entry $groupUsers[$i] -Step 'PlanRemoved' -Message ('not removed from the licence group ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$users.Remove($groupUsers[$i]) }
    }
    foreach ($entry in @($users | Where-Object { $_ -notin $groupUsers })) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            if (-not $tenantSkus) { $tenantSkus = @((Get-PcmTenantFact -Context $context).Skus) }
            $removedText[$entry.EntraId] = if ($entry.Item.kind -ne 'User') {
                # Shared mailbox not fully converted: its temporary licence goes back to what it was before Convert.
                $sku = $tenantSkus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
                if ($sku) { 'temporary licence ' + (Restore-PcmLicence -UserId $entry.EntraId -SkuId $sku.skuId -Original $entry.Original).ToLowerInvariant() } else { 'no temporary licence' }
            } elseif ($config.Licensing.Users.Mode -eq 'Kiosk') {
                $kiosk = Get-PcmKioskAssignment -Licence (Get-PcmUserLicence -UserId $entry.EntraId) -Skus $tenantSkus
                if ($kiosk) { "$($kiosk.SkuName) " + (Restore-PcmLicence -UserId $entry.EntraId -SkuId $kiosk.SkuId -Original $entry.Original).ToLowerInvariant() } else { 'no Kiosk licence' }
            } else {
                $sku = $tenantSkus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Users.SkuPartNumber } | Select-Object -First 1
                if ($sku) { "$($sku.skuPartNumber) " + (Restore-PcmLicence -UserId $entry.EntraId -SkuId $sku.skuId -Original $entry.Original).ToLowerInvariant() } else { 'no direct licence' }
            }
        } catch { Stop-PcmRecoverEntry -Entry $entry -Step 'PlanRemoved' -Message $_.Exception.Message; [void]$users.Remove($entry) }
    }
    $context.CurrentIdentity = ''
    foreach ($entry in $users) { Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'PlanRemoved' -Detail "Exchange plan removed: $($removedText[$entry.EntraId])" -Item @{ step = 'PlanRemoved' } }
    # 6. One wait for all: MailUser with the on-premises GUID.
    $late = @(Wait-PcmExoStateSet -Phase 'Waiting for the on-premises mailboxes of the users' -Entries $users.ToArray() -TimeoutSeconds (Get-PcmWaitSeconds 'MailboxTimeoutMinutes') -Test { param($state, $entry) $state.Type -eq 'MailUser' -and $state.ExchangeGuid -eq [string]$entry.Record.exchange_guid } -OnReady {
            Write-PcmJournal -ObjectGuid ([string]$args[0].Item.object_guid) -Step 'MailUser' -Detail "MailUser with the on-premises GUID after $($args[2]) s" -Item @{ step = 'MailUser' } })
    $nowAt = if ($late.Count) { Get-PcmExoStateSet -Id (& $ids $late) -TagAttribute $config.Retention.TagAttribute } else { @{} }
    foreach ($entry in $late) {
        $shows = $nowAt[$entry.EntraId]
        $held = $shows -and (Test-PcmToolHold -Set $HoldSet -Holds @($shows.Holds))
        $why = if ($shows -and $shows.Type) { " (Exchange Online shows $($shows.Type))" } else { '' }
        $hold = if ($held) { '' } else { '; the cloud mailbox is not under the case hold yet (it is added once the user is back: run Recover -Batch again to finish)' }
        Stop-PcmRecoverEntry -Entry $entry -Step 'RecoverUser' -Message ('not a MailUser with the on-premises ExchangeGuid {0} yet{1}{2}' -f $entry.Record.exchange_guid, $why, $hold)
        [void]$users.Remove($entry)
    }
    # 7. Case hold, by object ID (the policy keeps it right when the UPN and the SMTP address differ), one wait.
    $ready = @($users) + @($back)
    $hold = Add-PcmRecoverHold -Entries $ready -HoldSet $HoldSet
    foreach ($entry in $hold.Stamped) {
        $guid = [string]$entry.Item.object_guid
        $entry.Row.ExchangeOnline = 'MailUser'
        $text = if ($entry.Item.kind -eq 'User') { 'back on-premises (MailUser, on-premises GUID); the cloud mailbox stays as ComponentShared under the case hold' }
        else { 'shared mailbox not fully converted, rolled back like a user (MailUser, on-premises GUID); its cloud data, if any, stays under the case hold' }
        Write-PcmJournal -ObjectGuid $guid -Step 'Hold' -Detail "case hold stamped after $($hold.Seconds[$entry.EntraId]) s" -Item @{ step = 'Hold'; status = 'Done' }
        Set-PcmRowResult -Row $entry.Row -Status Success -Detail $text -Field 'Holds' -Value (Get-PcmHoldTagOf -HoldSet $HoldSet -Id $entry.EntraId)
    }
    foreach ($entry in $hold.Late) {
        $entry.Row.ExchangeOnline = 'MailUser'
        Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Hold' -Outcome Info -Detail 'case hold added, not stamped yet' -Item @{ step = 'HoldPending' }
        Set-PcmRowResult -Row $entry.Row -Status Pending -Detail "back on-premises; case hold not stamped yet: run Recover -Batch $Batch again to check it (the hold is never removed from an object that is back)"
    }
}

function Invoke-PcmRecoverObjectSync {
    <#
    The shared identities that a delta cycle did not recreate are synchronised one by one on the Entra Connect server (Invoke-PcmEntraConnectObject, by the distinguished name of
    the snapshot). Each answer is a step of the journal, so that a resumed Recover does not repeat a synchronisation that was done. Returns the entries synchronised by this run.
    #>
    param([Parameter(Mandatory)][object[]]$Entries)
    $todo = @($Entries | Where-Object { $_.TargetedSync -ne 'Done' })
    if (-not $todo.Count) { return @() }
    $context.CurrentOperation = 'EntraConnect-SyncObject'
    $withDn = @($todo | Where-Object { [string]$_.Record.distinguished_name })
    $answers = @()
    if ($withDn.Count) {
        try { $answers = @(Invoke-PcmEntraConnectObject -Context $context -ObjectDn ([string[]]@($withDn | ForEach-Object { [string]$_.Record.distinguished_name })) -Caller $caller) }
        catch { $failure = $_.Exception.Message; $answers = @($withDn | ForEach-Object { [pscustomobject]@{ Dn = [string]$_.Record.distinguished_name; Ok = $false; Detail = $failure } }) }
    }
    $synced = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $todo) {
        $dn = [string]$entry.Record.distinguished_name
        $answer = $answers | Where-Object { Test-PcmSameText ([string]$_.Dn) $dn } | Select-Object -First 1
        $ok = [bool]($answer -and $answer.Ok)
        $detail = if (-not $dn) { 'the snapshot holds no distinguished name for this object' } elseif ($answer) { [string]$answer.Detail } else { 'no answer from Entra Connect' }
        $entry.TargetedSync = if ($ok) { 'Done' } else { 'Failed' }
        $entry.TargetedSyncDetail = $detail
        if (-not $ok -and $detail -match 'no endpoint listening' -and ((Test-PcmSameText $context.Config.EntraConnect.Mode 'Remoting') -or (Test-PcmSameText $context.Config.EntraConnect.Mode 'Script'))) { $context.FullSyncOnServerOnly = $true }
        Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'TargetedSync' -Outcome $(if ($ok) { 'Ok' } else { 'Fail' }) -Detail $detail -Item @{ step = 'TargetedSync'; result_json = (@{ cloudGuid = [string]$entry.CloudGuid; state = [string]$entry.TargetedSync } | ConvertTo-Json -Compress) }
        Write-PcmItem -Context $context -Status $(if ($ok) { 'Ok' } else { 'Warn' }) -Icon Sync -Text ('{0}: Entra Connect, targeted synchronisation of the object: {1}' -f $entry.Row.Identity, $detail)
        if ($ok) { $synced.Add($entry) }
    }
    return @($synced)
}

function Invoke-PcmRecoverSharedWave {
    <#
    A wave of shared mailboxes, phase by phase: case hold and ONE wait for the stamps (an identity is never deleted
    without it), Entra Connect scheduler paused, identities deleted (Graph batches), every mailbox checked inactive,
    identities deleted permanently from the recycle bin (Graph batches), scheduler resumed (always, even after an
    error), ONE delta cycle, then ONE wait until Entra Connect has recreated every object from AD (MailUser with the
    on-premises GUID). The scheduler is paused only during the deletions of the wave.
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$HoldSet)
    $config = $context.Config
    $interval = $config.Polling.IntervalSeconds
    $fresh = @($Entries | Where-Object { $_.Path -eq 'Shared' })
    $toPurge = [System.Collections.Generic.List[object]]::new()
    if ($fresh.Count) {
        $hold = Add-PcmRecoverHold -Entries $fresh -HoldSet $HoldSet
        foreach ($entry in $hold.Late) { Stop-PcmRecoverEntry -Entry $entry -Step 'RecoverShared' -Message 'the case hold is not stamped on the shared mailbox: identity NOT deleted (it would not become inactive)' }
        foreach ($entry in $hold.Stamped) { Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Hold' -Detail "case hold stamped after $($hold.Seconds[$entry.EntraId]) s" -Item @{ step = 'Hold' } }
        $fresh = @($hold.Stamped)
    }
    $deleted = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @($Entries | Where-Object { $_.Path -eq 'SharedDeleted' })) {
        if (Get-PcmValue $entry 'AwaitInactive' $false) {
            Write-PcmItem -Context $context -Status Info -Text ('{0}: identity deleted by an earlier run: waiting until the mailbox {1} is inactive' -f $entry.Row.Identity, $entry.CloudGuid)
            $deleted.Add($entry)
            continue
        }
        Write-PcmItem -Context $context -Status Info -Text ('{0}: identity deleted by an earlier run (mailbox {1} inactive): recycle bin and recreated object checked' -f $entry.Row.Identity, $entry.CloudGuid)
        $toPurge.Add($entry)
    }
    $paused = $false
    # Dot-sourced: it sets $paused of this function.
    $pause = {
        if (-not $paused) {
            $paused = $true
            $pauseState = Invoke-PcmEntraConnect -Context $context -Operation Pause -Caller $caller
            Write-PcmItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler paused: $pauseState"
        }
    }
    try {
        if ($fresh.Count) {
            . $pause
            foreach ($entry in $fresh) { $entry.CloudGuid = $entry.Exo.ExchangeGuid }
            # Written before the identities are deleted: a run that stops during or after the deletion still knows which cloud mailbox to wait for.
            foreach ($entry in $fresh) { Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Deleting' -Outcome Info -Detail ('identity about to be deleted, cloud mailbox {0}' -f $entry.CloudGuid) -Item @{ step = 'Deleting'; result_json = (@{ cloudGuid = [string]$entry.CloudGuid } | ConvertTo-Json -Compress) } }
            $answers = Invoke-PcmGraphBatch -Requests @($fresh | ForEach-Object { @{ Method = 'DELETE'; Url = ('/users/{0}' -f $_.EntraId) } })
            for ($i = 0; $i -lt $fresh.Count; $i++) {
                if ($answers[$i].Ok -or $answers[$i].Status -eq 404) {
                    Write-PcmJournal -ObjectGuid ([string]$fresh[$i].Item.object_guid) -Step 'Deleted' -Detail ('identity deleted, cloud mailbox {0}' -f $fresh[$i].CloudGuid) -Item @{ step = 'Deleted'; result_json = (@{ cloudGuid = [string]$fresh[$i].CloudGuid } | ConvertTo-Json -Compress) }
                    $deleted.Add($fresh[$i])
                }
                else { Stop-PcmRecoverEntry -Entry $fresh[$i] -Step 'RecoverShared' -Message ('identity not deleted ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error) }
            }
        }
        if ($deleted.Count) {
            # Each mailbox inactive (no filter on ExchangeGuid in Exchange Online: one call per mailbox and round).
            $clock = [Diagnostics.Stopwatch]::StartNew()
            $null = Wait-PcmCondition -Phase 'Waiting for the inactive mailboxes' -TimeoutSeconds 900 -IntervalSeconds 15 -Test {
                foreach ($entry in @($deleted)) {
                    if (Get-Mailbox -InactiveMailboxOnly -Identity $entry.CloudGuid -ErrorAction SilentlyContinue) {
                        Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Inactive' -Detail ("identity deleted, mailbox {0} inactive after {1} s" -f $entry.CloudGuid, [int]$clock.Elapsed.TotalSeconds) -Item @{ step = 'Inactive'; result_json = (@{ inactiveGuid = $entry.CloudGuid } | ConvertTo-Json -Compress) }
                        [void]$deleted.Remove($entry); $toPurge.Add($entry)
                    }
                }
                -not $deleted.Count
            }
            foreach ($entry in $deleted) { Stop-PcmRecoverEntry -Entry $entry -Step 'RecoverShared' -Message ("identity deleted but mailbox {0} not inactive after 15 min: check it before anything else (Entra ID recycle bin, 30 days), then run Recover -Batch again" -f $entry.CloudGuid) }
        }
        if ($toPurge.Count) {
            $binUrl = { param($entry) '/directory/deletedItems/{0}' -f $entry.EntraId }
            $answers = Invoke-PcmGraphBatch -Requests @($toPurge | ForEach-Object { @{ Method = 'GET'; Url = ((& $binUrl $_) + '?$select=id') } })
            $binned = [System.Collections.Generic.List[object]]::new()
            $unread = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $toPurge.Count; $i++) {
                if ($answers[$i].Ok) { $binned.Add($toPurge[$i]) }
                elseif ($answers[$i].Status -ne 404) {
                    Stop-PcmRecoverEntry -Entry $toPurge[$i] -Step 'RecoverShared' -Message ('the Entra ID recycle bin could not be read ({0}): {1}. The identity is not confirmed as deleted permanently and nothing more was changed on it: run Recover -Batch again' -f $answers[$i].Status, $answers[$i].Error)
                    $unread.Add($toPurge[$i])
                }
            }
            foreach ($entry in $unread) { [void]$toPurge.Remove($entry) }
            if ($binned.Count) {
                . $pause
                $answers = Invoke-PcmGraphBatch -Requests @($binned | ForEach-Object { @{ Method = 'DELETE'; Url = (& $binUrl $_) } })
                for ($i = $binned.Count - 1; $i -ge 0; $i--) {
                    if (-not $answers[$i].Ok -and $answers[$i].Status -ne 404) {
                        Stop-PcmRecoverEntry -Entry $binned[$i] -Step 'RecoverShared' -Message ('identity not deleted permanently ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error)
                        [void]$toPurge.Remove($binned[$i]); $binned.RemoveAt($i)
                    }
                }
                $null = Wait-PcmCondition -TimeoutSeconds 300 -IntervalSeconds 10 -Test {
                    $answers = Invoke-PcmGraphBatch -Requests @($binned | ForEach-Object { @{ Method = 'GET'; Url = ((& $binUrl $_) + '?$select=id') } })
                    for ($i = $binned.Count - 1; $i -ge 0; $i--) { if (-not $answers[$i].Ok -and $answers[$i].Status -eq 404) { $binned.RemoveAt($i) } }
                    -not $binned.Count
                }
                foreach ($entry in $binned) { Stop-PcmRecoverEntry -Entry $entry -Step 'RecoverShared' -Message 'identity not confirmed as deleted permanently: it is still in the Entra ID recycle bin, or the recycle bin could not be read (the next sync would restore it)'; [void]$toPurge.Remove($entry) }
            }
            foreach ($entry in $toPurge) {
                Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Purged' -Detail 'identity deleted permanently' -Item @{ step = 'Purged' }
                Write-PcmItem -Context $context -Status Ok -Text ('{0}: mailbox inactive under the case hold, identity deleted permanently' -f $entry.Row.Identity)
            }
        }
    } finally {
        if ($paused) {
            try { $state = Invoke-PcmEntraConnect -Context $context -Operation Resume -Caller $caller; Write-PcmItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler resumed: $state" }
            catch { Write-PcmLog -Context $context -Message "Entra Connect scheduler NOT resumed: $($_.Exception.Message). Resume it by hand (Set-ADSyncScheduler -SyncCycleEnabled `$true)." -Level Error }
        }
    }
    if (-not $toPurge.Count) { return }
    # One delta cycle, then the new objects of the wave together (Graph batches by immutableId, Exchange Online by filters).
    $context.CurrentOperation = 'EntraConnect-Delta'
    $state = Invoke-PcmEntraConnect -Context $context -Operation Delta -Caller $caller
    Write-PcmItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect delta cycle: $state"
    $waiting = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $toPurge) { $waiting.Add($entry) }
    $recreated = {
        $answers = Invoke-PcmGraphBatch -Requests @($waiting | ForEach-Object {
                @{ Method = 'GET'; Url = ('/users?$filter={0}&$select=id' -f [uri]::EscapeDataString(("onPremisesImmutableId eq '{0}'" -f ([string]$_.Record.immutable_id).Replace("'", "''")))) } })
        $found = @{}
        for ($i = 0; $i -lt $waiting.Count; $i++) {
            $values = @(if ($answers[$i].Ok) { @(Get-PcmValue $answers[$i].Body 'value' @()) | ForEach-Object { $_ } })
            if ($values.Count -eq 1 -and [string]$values[0].id -ne $waiting[$i].EntraId) { $found[$waiting[$i].EntraId] = [string]$values[0].id }
        }
        $states = Get-PcmExoStateSet -Id ([string[]]@($found.Values))
        foreach ($entry in @($waiting)) {
            if (-not $found.ContainsKey($entry.EntraId)) { continue }
            $newId = $found[$entry.EntraId]; $s = $states[$newId]
            if ($s -and $s.Type -eq 'MailUser' -and $s.ExchangeGuid -eq [string]$entry.Record.exchange_guid) {
                Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Recreated' -Detail ("new object {0} from AD, MailUser with the on-premises GUID" -f $newId) -Item @{ step = 'Recreated'; status = 'Done'; entra_id = $newId }
                Set-PcmRowResult -Row $entry.Row -Status Success -Detail ('recreated from AD ({0}); the cloud mailbox {1} is inactive under the case hold' -f $newId, $entry.CloudGuid) -Field 'ExchangeOnline' -Value 'MailUser'
                [void]$waiting.Remove($entry)
            }
        }
        -not $waiting.Count
    }
    $null = Wait-PcmCondition -Phase 'Waiting for Entra Connect to recreate the identities' -TimeoutSeconds (Get-PcmWaitSeconds 'SyncTimeoutMinutes') -IntervalSeconds $interval -Test $recreated
    if ($waiting.Count) {
        # A delta cycle does not recreate an identity that Entra Connect still holds as exported: each object that is still missing is synchronised on its own, then the wait goes on.
        $synced = @(Invoke-PcmRecoverObjectSync -Entries @($waiting))
        if ($synced.Count) {
            $null = Wait-PcmCondition -Phase 'Waiting for Entra Connect to recreate the identities' -TimeoutSeconds ([Math]::Min(600, (Get-PcmWaitSeconds 'SyncTimeoutMinutes'))) -IntervalSeconds $interval -Test $recreated
        }
    }
    foreach ($entry in $waiting) {
        $tried = if ($entry.TargetedSyncDetail) { " (targeted synchronisation of the object: $($entry.TargetedSyncDetail))" } else { '' }
        Set-PcmRowResult -Row $entry.Row -Status Pending -Detail "identity deleted permanently, mailbox inactive; Entra Connect did not recreate it${tried}: run a full synchronisation on the Entra Connect server, then run Recover -Batch $Batch again"
        $context.FullSyncNeeded = $true
    }
}

function Get-PcmOwnedResourceContact {
    param([object]$Intent)
    $name = [string](Get-PcmValue $Intent 'contactName' '')
    if (-not $name) { return $null }
    $found = @(Get-Recipient -Filter ("Name -eq '{0}'" -f $name.Replace("'", "''")) -ResultSize Unlimited -ErrorAction Stop)
    if (-not $found.Count) { return $null }
    if ($found.Count -ne 1 -or [string]$found[0].Name -cne $name -or [string]$found[0].RecipientTypeDetails -ne 'MailContact') { throw 'Routing contact creation identity is ambiguous or not journal-owned.' }
    $id = [string]$found[0].Guid
    $persisted = [string](Get-PcmValue $Intent 'contactId' '')
    if (-not $id -or ($persisted -and $persisted -ne $id)) { throw 'Routing contact immutable ID differs from the journal.' }
    $contact = Get-MailContact -Identity $id -ErrorAction Stop
    if (-not [bool](Get-PcmValue $Intent 'contactCreationAttempted' $false)) { throw 'Routing contact creation was never journaled as attempted; pre-existing contact is not adopted.' }
    $marker = [string](Get-PcmValue $contact ([string]$Intent.markerAttribute) '')
    $markerPending = Get-PcmValue $Intent 'contactMarkerPending' $false
    if ($markerPending -isnot [bool]) { throw 'Routing contact marker progress is invalid; no contact is adopted.' }
    if ($marker -cne [string]$Intent.token -and ($marker -or ($persisted -and -not $markerPending))) { throw 'Routing contact ownership marker differs from the journal.' }
    return $contact
}

function Invoke-PcmRecoverResourceWave {
    param([object[]]$Entries, [object]$HoldSet)
    foreach ($entry in $Entries) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            $intent = $entry.Intent
            $protected = (Get-PcmResourceProtectionMode $intent) -eq 'CaseHold'
            if (-not $protected -and (Get-PcmValue $context.Config.Resources 'AllowUnprotectedDeletion' $false) -ne $true) {
                throw 'Resources.AllowUnprotectedDeletion = $true is required: explicitly acknowledge no hold/no inactive copy and loss of cloud-created reservations/messages after export/reconciliation. No mutation.'
            }
            if (-not $intent.cloudId -or -not $intent.cloudGuid -or $intent.cloudGuid -eq [string]$entry.Record.exchange_guid) { throw 'Native resource immutable IDs were not persisted. Resume Convert first to prove ownership; no deletion.' }
            $route = Get-PcmResourceRoute -Config $context.Config -Record $entry.Record -RequireReady
            $oldRoute = [string](Get-PcmValue $intent 'route' '')
            if ($oldRoute -and $oldRoute -ine $route) { throw "Routing domain changed mid-Recover: keep the journaled destination $oldRoute until this batch is complete." }
            $intent | Add-Member -NotePropertyName route -NotePropertyValue $route -Force
            $intent | Add-Member -NotePropertyName contactName -NotePropertyValue (([string]$intent.creationName) -replace '^PRA-R-', 'PRA-C-') -Force
            $contact = Get-PcmOwnedResourceContact -Intent $intent
            $live = Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound
            if ($live) {
                if (-not $protected) { Assert-PcmUnprotectedDeletionRepeatable -Intent $intent }
                $null = Get-PcmResourceRoute -Config $context.Config -Record $entry.Record -RequireReady
                $owned = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
                if (-not $owned -or $owned.User.accountEnabled -ne $false) { throw 'Native resource ownership/disabled identity cannot be verified; no deletion.' }
                if (-not $protected) {
                    Assert-PcmUnprotectedResource -Owned $owned -CheckOrganization
                    $collisions = @(Get-PcmResourceCollisions -Record $entry.Record -Inventory (Get-PcmResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId)
                    if ($collisions.Count) { throw ($collisions -join '; ') }
                    Confirm-PcmWrite -Target ([string]$intent.cloudId) -Operation 'Entra ID: delete the native resource identity (no hold, no inactive copy)'
                    $intent | Add-Member -NotePropertyName unprotectedDeletionAttempted -NotePropertyValue $true -Force
                    $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Intent' -Force
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceDeleteIntent' -Detail "Unprotected deletion acknowledged: exact $($intent.cloudId)/$($intent.cloudGuid); NO hold, NO inactive copy, cloud reservations/messages NOT preserved; route $route."
                    $final = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
                    if (-not $final -or $final.User.accountEnabled -ne $false) { throw 'Final resource ownership/disabled identity verification failed; no deletion.' }
                    Assert-PcmUnprotectedResource -Owned $final -CheckOrganization
                    $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Issuing' -Force
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceDeleteIssuing' -Detail 'Deletion request starting; ambiguous/interrupted response cannot authorise another deletion of a restored identity.'
                    try {
                        $null = Invoke-PcmGraph -Method DELETE -Uri ("v1.0/users/{0}" -f $intent.cloudId) -AllowNotFound -NoRetry -Approved
                    } catch {
                        $failure = $_
                        $code = 0
                        for ($exception = $failure.Exception; $exception; $exception = $exception.InnerException) {
                            $direct = $exception.PSObject.Properties['StatusCode']
                            $response = $exception.PSObject.Properties['Response']
                            if ($direct -and $null -ne $direct.Value) { $code = [int]$direct.Value; break }
                            if ($response -and $null -ne $response.Value -and $response.Value.PSObject.Properties['StatusCode']) {
                                $code = [int]$response.Value.StatusCode; break
                            }
                        }
                        if ($code -in @(400, 401, 403, 409, 412, 422, 429)) {
                            $intent.unprotectedDeletionProgress = 'Intent'
                            Write-PcmResourceJournal -Entry $entry -Step 'ResourceDeleteRejected' -Detail "DELETE rejected with HTTP $code; no accepted deletion is recorded."
                        }
                        throw $failure
                    }
                    $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Issued' -Force
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceIdentityDeleted' -Detail 'Only journal-owned unprotected cloud identity delete issued; NO inactive copy/data preservation guaranteed. Deletion still to be verified; no purge.'
                } else {
                $group = Get-PcmResourceGroup -Id $intent.licensingGroupId -Skus @((Invoke-PcmGraph -Uri 'v1.0/subscribedSkus').value)
                if (-not (Test-PcmResourceLicence -User $owned.User -Group $group)) { throw 'Native resource enterprise licensing is not active before the hold; no deletion.' }
                $collisions = @(Get-PcmResourceCollisions -Record $entry.Record -Inventory (Get-PcmResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId)
                if ($collisions.Count) { throw ($collisions -join '; ') }
                $added = Add-PcmHoldLocationSet -Set $HoldSet -Id @([string]$intent.cloudId)
                if ($added.Failed.ContainsKey([string]$intent.cloudId)) { throw $added.Failed[[string]$intent.cloudId] }
                $tag = Get-PcmHoldTagOf -HoldSet $HoldSet -Id $intent.cloudId
                if (-not $tag) { throw 'No specific tool hold policy contains this native resource; no deletion.' }
                $intent | Add-Member -NotePropertyName holdTag -NotePropertyValue $tag -Force
                $held = Wait-PcmCondition -TimeoutSeconds (Get-PcmWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                    $now = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
                    if (-not $now -or $now.User.accountEnabled -ne $false -or -not (Test-PcmResourceLicence -User $now.User -Group $group)) { throw 'Resource ownership/disabled sign-in/enterprise licence changed while waiting for hold; no deletion.' }
                    @($now.Mailbox.InPlaceHolds | ForEach-Object { [string]$_ }) -contains $tag
                }
                if (-not $held.Ok) {
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceHoldPending' -Detail 'Specific tool hold not stamped yet; resource, identity and licence retained.'
                    Set-PcmRowResult -Row $entry.Row -Status Pending -Detail 'Specific tool hold not stamped; identity NOT deleted. Retry Recover after hold deployment.'
                    continue
                }
                # Record the exact IDs/GUID and expected hold before the irreversible side effect.
                Confirm-PcmWrite -Target ([string]$intent.cloudId) -Operation 'Entra ID: delete the native resource identity (the mailbox stays inactive under the case hold)'
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceDeleteIntent' -Detail "verified $($intent.cloudId)/$($intent.cloudGuid), hold $tag, enterprise licence active; route $route"
                $final = Get-PcmOwnedResource -Intent $intent -Record $entry.Record
                if (-not $final -or $final.User.accountEnabled -ne $false -or -not (Test-PcmResourceLicence -User $final.User -Group $group) -or
                    @($final.Mailbox.InPlaceHolds | ForEach-Object { [string]$_ }) -notcontains $tag) { throw 'Final ownership/licence/hold verification failed immediately before deletion.' }
                $null = Invoke-PcmGraph -Method DELETE -Uri ("v1.0/users/{0}" -f $intent.cloudId) -AllowNotFound -Approved
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceIdentityDeleted' -Detail 'Only the journal-owned cloud identity deleted; recycle bin and mailbox data retained.'
                }
            }
            if ($protected) {
            $expectedHold = [string](Get-PcmValue $intent 'holdTag' '')
            if (-not $expectedHold -or -not $HoldSet.Tags.Contains($expectedHold)) { throw 'Expected tool hold was not journaled before deletion; verify the inactive mailbox manually. No routing contact is created.' }
            $inactive = Wait-PcmCondition -TimeoutSeconds (Get-PcmWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                # A filtered zero-result is absence, not a swallowed cmdlet/service failure.
                $found = @(Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop | Where-Object { [string]$_.ExchangeGuid -eq [string]$intent.cloudGuid })
                if (-not $found.Count) { return $false }
                if ($found.Count -ne 1 -or [string]$found[0].ExchangeGuid -ne [string]$intent.cloudGuid -or
                    @($found[0].InPlaceHolds | ForEach-Object { [string]$_ }) -notcontains $expectedHold) { throw 'Inactive resource GUID/specific tool hold verification failed; no contact, no purge.' }
                return $found[0]
            }
            if (-not $inactive.Ok) {
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceInactivePending' -Detail "Identity deleted; exact mailbox $($intent.cloudGuid) not yet inactive with expected hold. Recycle bin retained."
                Set-PcmRowResult -Row $entry.Row -Status Pending -Detail "Identity deleted; wait for exact inactive GUID $($intent.cloudGuid) with $expectedHold, then resume Recover. No mailbox data purged."
                continue
            }
            if (Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound) { throw 'Deleted resource identity is active again; routing contact creation refused.' }
            Write-PcmResourceJournal -Entry $entry -Step 'ResourceInactiveVerified' -Detail "Exact mailbox $($intent.cloudGuid) inactive under $expectedHold; mailbox and recycle-bin identity never purged."
            $collisions = @(Get-PcmResourceCollisions -Record $entry.Record -Inventory (Get-PcmResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId -OwnedGuid $intent.cloudGuid -OwnedContactId $(if ($contact) { [string]$contact.Guid } else { '' }))
            } else {
                $deleted = Wait-PcmCondition -TimeoutSeconds (Get-PcmWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                    Test-PcmUnprotectedResourceDeleted -Intent $intent
                }
                if (-not $deleted.Ok) {
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceDeletionPending' -Detail "Exact $($intent.cloudId)/$($intent.cloudGuid) not yet verified deleted/soft-deleted. No inactive copy or purge; resume."
                    Set-PcmRowResult -Row $entry.Row -Status Pending -Detail "Wait for journal-owned identity $($intent.cloudId) and soft-deleted non-inactive mailbox $($intent.cloudGuid), then resume Recover. Cloud reservations/messages NOT preserved; no purge."
                    continue
                }
                $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Verified' -Force
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceDeletionVerified' -Detail 'Exact identity deleted and mailbox soft-deleted, not inactive; foreign collisions and Exchange contact writes still to be checked. No purge/data preservation.'
                $inventory = Get-PcmResourceInventory -Records @($entry.Record)
                $collisions = @(Get-PcmUnprotectedRecoverCollisions -Intent $intent -Record $entry.Record -Inventory $inventory -OwnedContactId $(if ($contact) { [string]$contact.Guid } else { '' }))
                if ($collisions.Count) {
                    $reason = "Address release pending: $($collisions -join '; ') No automatic purge; resolve safely and resume this Recover batch."
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceAddressReleasePending' -Detail $reason
                    Set-PcmRowResult -Row $entry.Row -Status Pending -Detail $reason
                    continue
                }
            }
            if ($collisions.Count) { throw ($collisions -join '; ') }
            if (Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound) { throw 'Deleted resource identity is active again; routing contact creation/configuration refused.' }
            if (-not $contact) {
                Confirm-PcmWrite -Target ([string]$intent.contactName) -Operation 'Exchange Online: create the routing contact of the resource'
                $intent | Add-Member -NotePropertyName contactCreationAttempted -NotePropertyValue $true -Force
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceContactIntent' -Detail "create owned contact $($intent.contactName); addresses $((Get-PcmResourceAddresses $entry.Record) -join ', '), destination $route"
                $alias = [string](Get-PcmValue $entry.Record 'alias' '')
                if (-not $alias) { $alias = ([string]$entry.Record.primary_smtp_address -split '@')[0] }
                $created = New-MailContact -Name $intent.contactName -Alias $alias -ExternalEmailAddress $route -ErrorAction Stop
                $createdId = [string](Get-PcmValue $created 'Guid' '')
                if ($createdId) {
                    $intent | Add-Member -NotePropertyName contactId -NotePropertyValue $createdId -Force
                    $intent | Add-Member -NotePropertyName contactMarkerPending -NotePropertyValue $true -Force
                    Write-PcmResourceJournal -Entry $entry -Step 'ResourceContactCreated' -Detail "Returned routing contact immutable ID $createdId pinned before configuration."
                }
                $contact = Get-PcmOwnedResourceContact -Intent $intent
                if (-not $contact) { throw 'Routing contact creation accepted but not visible; resume this batch by its journaled name.' }
            }
            if (-not (Get-PcmValue $intent 'contactId' '')) {
                $intent | Add-Member -NotePropertyName contactId -NotePropertyValue ([string]$contact.Guid) -Force
                $intent | Add-Member -NotePropertyName contactMarkerPending -NotePropertyValue $true -Force
                Write-PcmResourceJournal -Entry $entry -Step 'ResourceContactCreated' -Detail "Journal-owned routing contact immutable ID $($contact.Guid) pinned before configuration."
            }
            $parameters = @{ Identity = [string]$contact.Guid; ExternalEmailAddress = $route; EmailAddresses = @(Get-PcmResourceAddresses $entry.Record)
                HiddenFromAddressListsEnabled = [bool](Get-PcmValue $entry.Record 'hidden_from_address_lists' $false); ErrorAction = 'Stop' }
            $parameters[[string]$intent.markerAttribute] = [string]$intent.token
            Confirm-PcmWrite -Target ([string]$contact.Guid) -Operation 'Exchange Online: configure the routing contact of the resource'
            Set-MailContact @parameters
            $intent | Add-Member -NotePropertyName contactMarkerPending -NotePropertyValue $false -Force
            $contact = Get-PcmOwnedResourceContact -Intent $intent
            if (([string]$contact.ExternalEmailAddress -replace '^smtp:', '') -ine $route -or
                [string]$contact.PrimarySmtpAddress -ine [string]$entry.Record.primary_smtp_address -or
                [bool]$contact.HiddenFromAddressListsEnabled -ne $parameters.HiddenFromAddressListsEnabled) { throw 'Routing contact destination/primary/hidden verification failed.' }
            foreach ($address in @(Get-PcmResourceAddresses $entry.Record)) {
                if ($address -notin @($contact.EmailAddresses | ForEach-Object { [string]$_ })) { throw "Routing contact address not applied: $address." }
            }
            $result = @{ contactId = [string]$contact.Guid; route = $route }
            $protectionText = if ($protected) { $result.inactiveGuid = $intent.cloudGuid; 'inactive cloud data and deleted identity retained under tool hold' } else {
                $result.deletedGuid = $intent.cloudGuid; 'NO hold/NO inactive copy; cloud-created reservations/messages NOT preserved, soft-deleted object retained without purge'
            }
            Write-PcmResourceJournal -Entry $entry -Step 'ResourceRouted' -Status Done -Detail "Owned contact $($contact.Guid) routes to $route; $protectionText; no calendar merge." -Extra $result
            Set-PcmRowResult -Row $entry.Row -Status Success -Field ExchangeOnline -Value MailContact -Detail "native resource routes to $route via owned contact; $protectionText. On-premises objects untouched; cloud calendars are not merged."
        } catch { Stop-PcmRecoverEntry -Entry $entry -Step 'RecoverResource' -Message ($_.Exception.Message + ' Resolve the finding and resume the same Recover batch. No automatic purge or hold removal. Unprotected deletion does NOT preserve cloud reservations/messages.') }
    }
    $context.CurrentIdentity = ''
}

function Get-PcmBatchUsersMode {
    <# Licensing.Users.Mode that Convert recorded in a batch ('' when the batch did not record it). #>
    param([AllowNull()][object]$Row)
    $text = [string](Get-PcmValue $Row 'scope_json' '')
    if (-not $text) { return '' }
    try { $scope = $text | ConvertFrom-PcmJson } catch { return '' }
    return [string](Get-PcmValue $scope 'UsersMode' '')
}

function Invoke-PcmRecover {
    $config = $context.Config
    if (-not $Batch) { throw '-Batch is required: the ID of the Convert batch to roll back (printed at the end of Convert).' }
    $context.Journal = Open-PcmStore -Path $config.Store.JournalPath -Root $context.Root -Kind Journal
    $context.CurrentOperation = 'Read-Batch'
    $convert = Get-PcmBatch $context.Journal -Id $Batch
    if ($convert -and $convert.action -eq 'Recover' -and $convert.convert_batch) {
        $given = $convert
        $convert = Get-PcmBatch $context.Journal -Id ([string]$given.convert_batch)
        if ($convert) { Write-PcmItem -Context $context -Status Info -Icon Batch -Text ('Batch {0} is a rollback of Convert batch {1}: the rollback goes on with Convert batch {1}.' -f $given.id, $convert.id) }
    }
    if (-not $convert -or $convert.action -ne 'Convert') { throw "Convert batch $Batch not found in $($config.Store.JournalPath)." }
    $existing = @(Invoke-PcmSql $context.Journal "SELECT * FROM batch WHERE action = 'Recover' AND convert_batch = @c COLLATE NOCASE ORDER BY created_utc DESC" @{ c = $convert.id } -As Rows)
    $resume = $existing | Where-Object { $_.status -in @('Running','Partial') } | Select-Object -First 1
    $context.ConvertBatchId = [string]$convert.id
    if ($resume) { $context.ResumeBatchId = [string]$resume.id }
    $finished = $existing | Where-Object { $_.status -eq 'Complete' } | Select-Object -First 1
    if (-not $resume -and $finished) { throw "Convert batch $($convert.id) was already rolled back by Recover batch $($finished.id) (complete): nothing to do. Check shows the current state of the objects." }
    $data = Read-PcmSnapshotData -Id ([long]$convert.snapshot_id) -Verb 'roll back' -All -NoCount -Batch $convert
    $byGuid = @{}; foreach ($record in $data.Records) { $byGuid[[string]$record.object_guid] = $record }
    $convertItems = @(Get-PcmBatchItem $context.Journal $convert.id | Where-Object { $_.status -ne 'Skipped' -and $_.step -and ($_.entra_id -or (Get-PcmResourceIntent -Item $_)) })
    $done = @{}; $previous = @{}
    if ($resume) { foreach ($item in @(Get-PcmBatchItem $context.Journal $resume.id)) { $previous[[string]$item.object_guid] = $item; if ($item.status -eq 'Done') { $done[[string]$item.object_guid] = $true } } }
    $targets = @($convertItems | Where-Object { -not $done.ContainsKey([string]$_.object_guid) })
    $rolledBack = @{}
    foreach ($guid in @($done.Keys)) { $rolledBack[[string]$guid] = 'already rolled back by this Recover batch' }
    $doneItems = @($convertItems | Where-Object { $done.ContainsKey([string]$_.object_guid) })
    $targets = @(Select-PcmTarget -Items $targets -Where "the objects of batch $($convert.id) still to roll back" -AllRecords @($data.Records) -Known $rolledBack -KnownItems $doneItems -InBatch `
            -RecordOf { param($t) $byGuid[[string]$t.object_guid] } -Keys {
            param($t) $r = $byGuid[[string]$t.object_guid]
            @([string]$t.identity, [string]$t.object_guid) + $(if ($r) { @([string]$r.primary_smtp_address, [string]$r.user_principal_name, [string]$r.sam_account_name) } else { @() }) })
    $targets = @(Limit-PcmTarget -Items $targets)
    Write-PcmTargetHeader -Count $targets.Count
    # The Exchange plan of the users is taken back the way Convert gave it: another Licensing.Users.Mode now would remove it from the wrong place, so the run stops before any change.
    $batchMode = Get-PcmBatchUsersMode -Row $convert
    if ($batchMode -and -not (Test-PcmSameText $batchMode ([string]$config.Licensing.Users.Mode)) -and @($targets | Where-Object { -not (Get-PcmResourceIntent -Item $_) -and (Test-PcmSameText ([string]$_.kind) 'User') }).Count) {
        $context.NoChangeMade = $true; $context.NextOnFailure = $true; $context.CurrentIdentity = ''
        $context.NextSteps = @(('Set Licensing.Users.Mode back to {0} in the configuration (with the same GroupId or SkuPartNumber as at Convert), then run the same command again:' -f $batchMode), (Get-PcmCommandLine -ActionName 'Recover' -ModeName $effectiveMode -BatchId ([string]$convert.id) -KeepFilters))
        throw ('Convert batch {0} gave the Exchange plan of its users with Licensing.Users.Mode = {1}, and the configuration now says {2}: the rollback would take the plan away from the wrong place. Nothing was changed.' -f $convert.id, $batchMode, $config.Licensing.Users.Mode)
    }
    $context.NativeResourceScope = @($targets | Where-Object { Get-PcmResourceIntent -Item $_ }).Count -gt 0
    $requiresHold = @($targets | Where-Object {
            $intent = Get-PcmResourceIntent -Item $_
            -not $intent -or (Get-PcmResourceProtectionMode $intent) -eq 'CaseHold'
        }).Count -gt 0
    if ($requiresHold -and -not $config.Retention.HoldPolicy) { throw 'Retention.HoldPolicy is empty: protected/user/shared Recover requires that eDiscovery case hold.' }
    Write-PcmItem -Context $context -Status Info -Icon Batch -Text ('Convert batch {0} of {1}: {2} object(s) to roll back{3}' -f $convert.id, (Format-PcmUtcText $convert.created_utc), $targets.Count, $(if ($resume) { " (resuming Recover $($resume.id))" } else { '' }))

    Write-PcmStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-PcmCloudSession; if ($requiresHold) { Connect-PcmCompliance -Context $context } } catch { $context.CloudConnectFailed = $true; throw }
    Write-PcmItem -Context $context -Status Ok -Text ($how + $(if ($requiresHold) { ', Security & Compliance (case hold)' } else { '; unprotected native lifecycle, no compliance connection' }))

    Write-PcmStep -Context $context -Title 'Prerequisites' -Icon Key
    $context.CurrentOperation = 'Prerequisites'
    $holdSet = $null
    if ($requiresHold) {
    $holdSet = Get-PcmHoldPolicySet -Policy $config.Retention.HoldPolicy -Limit $config.Retention.HoldPolicyLimit
    Write-PcmItem -Context $context -Status Ok -Text ('Case hold: {0} ({1} mailbox(es) held, {2} at most per policy; full policies are followed by {3}-02, -03... in the same case)' -f
        (@($holdSet.Policies | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join ', '), (@($holdSet.Policies | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum, $holdSet.Limit, $holdSet.Base)
    Write-PcmItem -Context $context -Status Info -Text ("Entra Connect {0} or later, operations by {1}{2}" -f $config.EntraConnect.MinVersion, $config.EntraConnect.Mode, $(if ($config.EntraConnect.Server) { " on $($config.EntraConnect.Server)" } else { '' }))
    }
    # Each object goes the way its current state allows; a step already done is never undone (a hold is never lifted
    # from an object that is back on-premises).
    #   User          user, or a shared mailbox whose Convert stopped half-way: back to AD like a user
    #   Back          already a synchronised MailUser with the on-premises GUID: case hold only
    #   Shared        SharedMailbox: hold, identity deleted, purged, recreated by Entra Connect
    #   SharedDeleted identity already deleted by an earlier run: purge check, recreated object
    $plan = [System.Collections.Generic.List[object]]::new()
    $resourceBlocked = [System.Collections.Generic.List[object]]::new()
    if (Test-PcmStop) { throw 'Stopped by the operator before reading the state of the objects: nothing was changed.' }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $exoStates = Get-PcmExoStateSet -Id ([string[]]@($targets | Where-Object { -not (Get-PcmResourceIntent -Item $_) } | ForEach-Object { [string]$_.entra_id })) -TagAttribute $config.Retention.TagAttribute
    if ($targets.Count) { Write-PcmItem -Context $context -Status Info -Text ('Exchange Online state of {0} object(s) read in {1}' -f $targets.Count, (Format-PcmDuration $clock.Elapsed.TotalSeconds)) }
    foreach ($item in $targets) {
        $guid = [string]$item.object_guid
        $record = $byGuid[$guid]
        $row = New-PcmRow $(if ($record) { $record } else { [pscustomobject]@{ user_principal_name = $item.identity; kind = $item.kind; primary_smtp_address = ''; object_guid = $item.object_guid; exchange_guid = '' } })
        $row.Action = 'Recover'
        [void]$context.Rows.Add($row)
        if (-not $record) { Set-PcmRowResult -Row $row -Status Error -Detail "not in snapshot $($convert.snapshot_id)"; continue }
        $intent = Get-PcmResourceIntent -Item $item
        if ($intent) {
            try {
                if ($previous.ContainsKey($guid)) {
                    $oldIntent = Get-PcmResourceIntent -Item $previous[$guid]
                    if ($oldIntent -and (Get-PcmResourceProtectionMode $oldIntent) -cne (Get-PcmResourceProtectionMode $intent)) { throw 'Recover protection mode differs from original Convert journal; no downgrade or deletion.' }
                    if ($oldIntent) {
                        foreach ($field in @('token','creationName','markerAttribute','cloudId','cloudGuid','cloudType','licensingGroupId')) {
                            $before = [string](Get-PcmValue $oldIntent $field ''); $after = [string](Get-PcmValue $intent $field '')
                            $sameField = if ($field -in @('cloudId','cloudGuid','licensingGroupId')) { Test-PcmSameText $before $after } else { $before -ceq $after }
                            if (-not $sameField) { throw "Recover ownership $field differs from original Convert journal; no identity is adopted or deleted." }
                        }
                    }
                    if ($oldIntent) { $intent = $oldIntent }
                }
                $protected = (Get-PcmResourceProtectionMode $intent) -eq 'CaseHold'
                if (-not $protected) {
                    $row.Warnings = 'NO hold, NO inactive copy: cloud-created reservations/messages NOT preserved or merged. Export/reconcile before explicitly allowing unprotected deletion.'
                    Write-PcmItem -Context $context -Status Warn -Text "$($row.Identity): $($row.Warnings)"
                    if ((Get-PcmValue $config.Resources 'AllowUnprotectedDeletion' $false) -ne $true) { throw 'Resources.AllowUnprotectedDeletion = $true required after export/reconciliation and explicit acknowledgement of cloud reservations/messages loss (no hold/no inactive copy). No mutation.' }
                }
                $route = Get-PcmResourceRoute -Config $config -Record $record -RequireReady
                if (-not $intent.cloudId -or -not $intent.cloudGuid) { throw 'Native resource immutable ownership IDs not persisted: resume Convert first; no deletion.' }
                $live = Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound
                if ($live) {
                    if (-not $protected) { Assert-PcmUnprotectedDeletionRepeatable -Intent $intent }
                    $owned = Get-PcmOwnedResource -Intent $intent -Record $record
                    if (-not $owned -or $owned.User.accountEnabled -ne $false) { throw 'Native resource cloud-only/disabled ownership verification failed.' }
                    $null = Get-PcmResourceRoute -Config $config -Record $record -RequireReady
                    if (-not $protected) {
                        Assert-PcmUnprotectedResource -Owned $owned -CheckOrganization
                    }
                }
                $nativeEntry = [pscustomobject]@{ Item = $item; Record = $record; Row = $row; EntraId = [string]$intent.cloudId; Exo = $null
                    Path = 'CloudResource'; CloudGuid = [string]$intent.cloudGuid; Original = @(); Intent = $intent }
                [void]$plan.Add($nativeEntry)
                $protectionText = if ($protected) { 'legacy hold/enterprise licence verification and exact inactive copy' } else { 'NO hold, NO inactive copy, cloud reservations/messages NOT preserved; verify deletion/soft-deleted state and address release without purge' }
                Write-PcmItem -Context $context -Status Info -Text "Native resource: $protectionText; delete only owned $($intent.cloudId)/$($intent.cloudGuid), contact to $route; addresses: $((Get-PcmResourceAddresses $record) -join ', '). No sync/SOA, no calendar merge."
            } catch {
                Set-PcmRowResult -Row $row -Status Error -Detail $_.Exception.Message
                $resourceBlocked.Add([pscustomobject]@{ Item = $item; Record = $record; Row = $row; Intent = $intent })
            }
            continue
        }
        $entry = [pscustomobject]@{ Item = $item; Record = $record; Row = $row; EntraId = [string]$item.entra_id; Exo = $null; Path = ''; CloudGuid = ''; AwaitInactive = $false; Original = @(Get-PcmOriginalLicence $item); TargetedSync = ''; TargetedSyncDetail = '' }
        $exo = $exoStates[$entry.EntraId]
        if (-not $exo) { $exo = Get-PcmExoState -Identity $entry.EntraId -TagAttribute $config.Retention.TagAttribute }
        $entry.Exo = $exo
        $row.ExchangeOnline = $exo.Type; $row.Holds = ($exo.Holds -join ', ')
        $lastStep = if ($previous.ContainsKey($guid)) { [string](Get-PcmValue $previous[$guid] 'step' '') } else { '' }
        if ($item.kind -ne 'User' -and $lastStep -in @('Deleting', 'Deleted', 'Inactive', 'Purged', 'TargetedSync') -and -not $exo.Type) {
            $result = $null; try { $result = [string](Get-PcmValue $previous[$guid] 'result_json' '') | ConvertFrom-PcmJson } catch { $result = $null }
            $entry.Path = 'SharedDeleted'
            $entry.CloudGuid = [string](Get-PcmValue $result 'inactiveGuid' '')
            if (-not $entry.CloudGuid) { $entry.CloudGuid = [string](Get-PcmValue $result 'cloudGuid' '') }
            $entry.AwaitInactive = $lastStep -in @('Deleting', 'Deleted')
            if ($lastStep -eq 'TargetedSync' -and (Test-PcmSameText ([string](Get-PcmValue $result 'state' '')) 'Done')) { $entry.TargetedSync = 'Done'; $entry.TargetedSyncDetail = 'done by an earlier run' }
            $row.ExchangeOnline = if ($entry.AwaitInactive) { 'identity deleted' } else { 'inactive' }
            [void]$plan.Add($entry); continue
        }
        if (-not $exo.Type) { Set-PcmRowResult -Row $row -Status Error -Detail ("no Exchange Online recipient for {0}: the identity was deleted or replaced outside this batch. Check it in Entra ID and Exchange Online; nothing was changed." -f $entry.EntraId); continue }
        if ($exo.Type -eq 'MailUser' -and $exo.ExchangeGuid -eq [string]$record.exchange_guid -and $exo.IsDirSynced -eq $true) { $entry.Path = 'Back'; [void]$plan.Add($entry); continue }
        if ($item.kind -ne 'User' -and $exo.Type -eq 'SharedMailbox') { $entry.Path = 'Shared'; [void]$plan.Add($entry); continue }
        $other = @($exo.Holds | Where-Object { $_ -and -not $holdSet.Tags.Contains([string]$_) })
        if ($exo.Type -eq 'UserMailbox' -and $other.Count) { Set-PcmRowResult -Row $row -Status Error -Detail ('other hold(s) on the cloud mailbox block the rollback: {0}. Remove them, then run Recover again.' -f ($other -join ', ')); continue }
        $entry.Path = 'User'
        [void]$plan.Add($entry)
    }
    $planUsers = @($plan | Where-Object { $_.Path -eq 'User' })
    $planBack = @($plan | Where-Object { $_.Path -eq 'Back' })
    $planShared = @($plan | Where-Object { $_.Path -in @('Shared', 'SharedDeleted') })
    $planResources = @($plan | Where-Object { $_.Path -eq 'CloudResource' })
    $partial = @($planUsers | Where-Object { $_.Item.kind -ne 'User' })
    Write-PcmStepCount -Context $context -Titles @('Users: back on-premises') -Count @($plan | Where-Object { $_.Path -in @('User', 'Back') }).Count
    Write-PcmStepCount -Context $context -Titles @('Shared mailboxes') -Count ($planShared.Count + $planResources.Count)
    Write-PcmItem -Context $context -Status Info -Icon Target -Text ('{0} user(s) and {1} shared mailbox(es) to roll back{2}{3}' -f @($planUsers + $planBack | Where-Object { $_.Item.kind -eq 'User' }).Count, ($planShared.Count + $partial.Count + @($planBack | Where-Object { $_.Item.kind -ne 'User' }).Count),
        $(if ($planBack.Count) { "; $($planBack.Count) already back on-premises (case hold only)" } else { '' }), $(if ($partial.Count) { "; $($partial.Count) shared mailbox(es) not fully converted, rolled back like users" } else { '' }))

    Write-PcmStep -Context $context -Title 'Plan and confirmation' -Icon Plan
    Write-PcmItem -Context $context -Status Info -Text ('Users, in waves of {0}: source of authority back to AD, delta sync, Exchange plan removed (licences as before Convert), MailUser with the on-premises GUID, cloud mailbox kept as ComponentShared under the case hold' -f [Math]::Max(1, [int](Get-PcmValue $context 'UserWaveSize' 500)))
    if ($planShared.Count) { Write-PcmItem -Context $context -Status Info -Text ('Shared mailboxes, in waves of {0} (Recover.WaveSize, or Licensing.Shared.Parallel when it is not set): case hold, Entra Connect scheduler paused, identities deleted (mailboxes inactive), deleted permanently, scheduler resumed, recreated from AD by a delta sync' -f (Get-PcmRecoverWaveSize $config)) }
    $recoverEstimate = Get-PcmDurationEstimate -UserWaves ([Math]::Ceiling(@($planUsers + $planBack).Count / [Math]::Max(1, [int](Get-PcmValue $context 'UserWaveSize' 500)))) -SharedRounds ([Math]::Ceiling($planShared.Count / (Get-PcmRecoverWaveSize $config)))
    if ($recoverEstimate.Text) { Write-PcmItem -Context $context -Status Info -Icon Clock -Text ('Rough estimate: about {0} (typical tenant; a slow Microsoft 365 evening can take much longer)' -f $recoverEstimate.Text) }
    foreach ($entry in $plan) { $entry.Row.FinalStatus = 'Planned'; Write-PcmLog -Context $context -Message ('Planned: {0} ({1})' -f $entry.Row.Identity, $entry.Path) -Level Sub }
    if ($context.Mode -ne 'Apply') {
        foreach ($title in @('Users: back on-premises', 'Shared mailboxes', 'Licence group')) { Skip-PcmStep -Title $title -Reason 'Preview: not run.' -Icon Write }
        $context.NextSteps = @('To make the changes, run the same command in Apply mode (the confirmation is asked once):', (Get-PcmCommandLine -ActionName 'Recover' -ModeName 'Apply' -BatchId ([string]$convert.id) -KeepFilters))
        return
    }
    if (-not $plan.Count) { throw 'Nothing to roll back (see above).' }
    $lossWarning = if (@($planResources | Where-Object { (Get-PcmResourceProtectionMode $_.Intent) -eq 'None' }).Count) { ' UNPROTECTED native resources: NO hold, NO inactive copy; cloud-created reservations/messages are NOT preserved or merged. Deletion may destroy them. Acknowledge after export/reconciliation.' } else { '' }
    Confirm-PcmApply (('Roll back {0} object(s) of Convert batch {1}?' -f $plan.Count, $convert.id) + $lossWarning)
    if ($resume) { $context.BatchId = [string]$resume.id }
    else {
        $context.BatchId = New-PcmBatch $context.Journal Recover @{ ConvertBatch = $convert.id; SnapshotId = [long]$convert.snapshot_id; SnapshotRunId = [string]$context.SnapshotRunId; SnapshotHash = [string]$context.SnapshotHash; Environment = $config.Environment
            Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name; ToolVersion = $toolVersion; ScopeJson = (@{ Identity = $Identity; CsvPath = $CsvPath; SearchBase = $SearchBase; Group = $Group; MaxObjects = $MaxObjects; Identities = @(Get-PcmIdentityList).Count; Scope = $Scope } | ConvertTo-Json -Compress) }
    }
    foreach ($entry in $plan) {
        $values = @{ kind = [string]$entry.Item.kind; identity = $entry.Row.Identity; entra_id = $entry.EntraId; status = 'Running' }
        if ($entry.Path -eq 'CloudResource') { $values.original_json = ($entry.Intent | ConvertTo-Json -Depth 15 -Compress) }
        Set-PcmBatchItem $context.Journal $context.BatchId ([string]$entry.Item.object_guid) $values
    }
    foreach ($entry in $resourceBlocked) {
        Write-PcmJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'ResourcePreconditionFailed' -Outcome Fail -Detail $entry.Row.Detail -Item @{
            kind = [string]$entry.Item.kind; identity = $entry.Row.Identity; entra_id = [string]$entry.Intent.cloudId; status = 'Failed'; step = 'ResourcePreconditionFailed'
            original_json = ($entry.Intent | ConvertTo-Json -Depth 15 -Compress); message = $entry.Row.Detail }
    }
    Write-PcmItem -Context $context -Status Ok -Icon Batch -Text ("Recover batch {0} in {1}" -f $context.BatchId, $config.Store.JournalPath)
    $stopped = $false
    # Waves: every phase of a wave runs for all its objects (one wait per phase); the operator's stop is honoured
    # between waves, a wave in progress is always finished.
    $runWaves = {
        param([object[]]$Entries, [int]$Size, [string]$Label, [scriptblock]$Run)
        $clock = [Diagnostics.Stopwatch]::StartNew()
        for ($start = 0; $start -lt $Entries.Count; $start += $Size) {
            $wave = @($Entries[$start..([Math]::Min($Entries.Count, $start + $Size) - 1)])
            if ($stopped -or (Test-PcmStop)) { $stopped = $true; foreach ($one in $wave) { Set-PcmRowStopped $one.Row }; continue }
            if ($Entries.Count -gt $Size) { Write-PcmProgress -Context $context -Phase $Label -Done $start -Total $Entries.Count -ElapsedSeconds $clock.Elapsed.TotalSeconds }
            & $Run $wave
        }
    }

    Write-PcmStep -Context $context -Title 'Users: back on-premises' -Icon Sync
    $context.CurrentOperation = 'Recover-Users'
    $userEntries = @($plan | Where-Object { $_.Path -in @('User', 'Back') })
    . $runWaves -Entries $userEntries -Size ([Math]::Max(1, [int](Get-PcmValue $context 'UserWaveSize' 500))) -Label 'Users' -Run { param($wave) Invoke-PcmRecoverUserWave -Entries $wave -HoldSet $holdSet }
    if (-not $userEntries.Count) { Write-PcmItem -Context $context -Status Skip -Text 'No user to bring back.' }

    Write-PcmStep -Context $context -Title 'Shared mailboxes' -Icon Mail
    $context.CurrentOperation = 'Recover-Shared'
    # The Entra Connect scheduler is paused only while the identities of one wave are deleted.
    . $runWaves -Entries $planShared -Size (Get-PcmRecoverWaveSize $config) -Label 'Shared mailboxes' -Run { param($wave) Invoke-PcmRecoverSharedWave -Entries $wave -HoldSet $holdSet }
    if (-not $planShared.Count) { Write-PcmItem -Context $context -Status Skip -Text 'No shared mailbox in this batch.' }
    . $runWaves -Entries $planResources -Size ([Math]::Max(1, [int](Get-PcmValue (Get-PcmValue $config 'Resources' @{}) 'Parallel' 100))) -Label 'Native resources' -Run { param($wave) Invoke-PcmRecoverResourceWave -Entries $wave -HoldSet $holdSet }
    $context.CurrentIdentity = ''
    Write-PcmStep -Context $context -Title 'Licence group' -Icon Key
    $context.CurrentOperation = 'Licence-Group'
    # The group serves every Convert batch of this journal: it goes back to AD only when no converted user is left,
    # whichever batch switched it (later waves record nothing, the group was already in the cloud).
    $changes = @(if (@($targets | Where-Object { -not (Get-PcmResourceIntent -Item $_) }).Count) {
            Invoke-PcmSql $context.Journal "SELECT * FROM tenant_change WHERE kind = 'GroupSoa' AND restored = 0" -As Rows
        })
    if (-not $changes.Count) { Write-PcmItem -Context $context -Status Skip -Text 'No licence group switched to the cloud by this tool.' }
    else {
        $pending = @(Invoke-PcmSql $context.Journal @"
SELECT DISTINCT ci.identity AS identity FROM batch_item ci JOIN batch cb ON cb.id = ci.batch_id
WHERE cb.action = 'Convert' AND ci.kind = 'User' AND ci.step IS NOT NULL
  AND NOT EXISTS (SELECT 1 FROM batch_item ri JOIN batch rb ON rb.id = ri.batch_id
                  WHERE rb.action = 'Recover' AND rb.convert_batch = cb.id COLLATE NOCASE AND ri.object_guid = ci.object_guid COLLATE NOCASE AND ri.status = 'Done')
"@ -As Rows)
        if ($pending.Count) {
            Write-PcmItem -Context $context -Status Warn -Text ('Licence group kept in the cloud: {0} converted user(s) not back yet ({1}); it goes back to AD with the last one.' -f $pending.Count, ((@($pending | Select-Object -First 5 | ForEach-Object { [string]$_.identity })) -join ', '))
        } else {
            foreach ($change in $changes) {
                Set-PcmSourceOfAuthority -Id $change.target_id -Cloud $false -Type groups
                Set-PcmTenantChange $context.Journal $change.batch_id 'GroupSoa' $change.target_id -Restored
                Write-PcmJournal -Step 'GroupSoaAd' -Detail "licence group $($change.target_id) back to AD"
                Write-PcmItem -Context $context -Status Ok -Text "Licence group $($change.target_id) managed by AD again (next Entra Connect cycle)"
            }
        }
    }
    $items = @(Get-PcmBatchItem $context.Journal $context.BatchId)
    $counts = @{ done = @($items | Where-Object status -eq 'Done').Count; failed = @($items | Where-Object status -eq 'Failed').Count; running = @($items | Where-Object status -eq 'Running').Count }
    $complete = $counts.done -eq @($convertItems).Count
    Set-PcmBatch $context.Journal $context.BatchId $(if ($complete) { 'Complete' } else { 'Partial' }) $counts
    $context.NextOnFailure = $true
    $context.NextSteps = @(if (-not $complete) {
            if (Get-PcmValue $context 'FullSyncNeeded' $false) {
                $server = if ([string](Get-PcmValue $config.EntraConnect 'Server' '')) { [string]$config.EntraConnect.Server } else { 'the Entra Connect server' }
                if (Get-PcmValue $context 'FullSyncOnServerOnly' $false) {
                    "Entra Connect did not recreate every deleted identity: the synchronisation of each object cannot be reached from a remote session. This command works only in a session opened ON the Entra Connect server ($server) itself, not through remoting or a script run remotely. Run a full synchronisation there (it can take a long while), in PowerShell on that server:"
                } else {
                    "Entra Connect did not recreate every deleted identity, even after it synchronised each object on its own. Run a full synchronisation on $server (it can take a long while), in PowerShell on that server:"
                }
                'Start-ADSyncSyncCycle -PolicyType Initial'
                'Then resume the rollback (nothing is rolled back twice):'
            } else { 'Some objects are not back yet. Fix what is named above, then resume them (nothing is rolled back twice):' }
            (Get-PcmCommandLine -ActionName 'Recover' -ModeName 'Apply' -BatchId ([string]$convert.id) -KeepFilters)
        })
    $createdLists = @(Get-PcmTenantChange $context.Journal $convert.id | Where-Object kind -eq 'NativeRoomList' |
        ForEach-Object { [string]$_.original_json | ConvertFrom-PcmJson } | Where-Object { $_.operation -eq 'Create' -and $_.creationAttempted })
    if ($createdLists.Count) {
        Write-PcmItem -Context $context -Status Info -Text "$($createdLists.Count) journaled cloud room list(s) retained: Recover never deletes groups or inserts routing contacts into RoomLists; review room discovery separately."
    }
}
function Get-PcmSourceText {
    <# Which objects the run took, in words (journal of the batch, global report). #>
    $dot = [char]0x00B7
    $text = switch (Get-PcmTargetKind) {
        'Identity' { "one object: $Identity" }
        'Csv' { 'list {0} ({1} object(s))' -f (Split-Path -Path $CsvPath -Leaf), @($context.Identities).Count }
        'SearchBase' { "OU $SearchBase" }
        'Group' { 'group {0}' -f $(if ($context.GroupTarget) { '{0} ({1})' -f $context.GroupTarget.Name, $context.GroupTarget.Origin } else { $Group }) }
        default { 'every object of the snapshot' }
    }
    if (-not (Test-PcmSameText $Scope 'All')) { $text += " $dot $Scope" }
    if ($MaxObjects -gt 0) { $text += " $dot first $MaxObjects" }
    return $text
}

function Write-PcmBatchRecord {
    <#
    At the end of a run on a batch: a Convert or Recover in Apply mode adds the run to the journal (when, who, which objects, which settings, how it ended), then every run
    that touches a batch (those and Check -Batch) writes the global report of the whole batch, and its path goes to the next steps. A failure here never changes the result of the run.
    #>
    $batchId = [string](Get-PcmValue $context 'BatchId' '')
    if (-not $batchId -or -not $context.Journal) { return }
    $isFollowUp = Test-PcmSameText $Action 'Check'
    if (-not $isFollowUp -and -not (Test-PcmSameText $effectiveMode 'Apply')) { return }
    $outcome = Get-PcmOutcome $context
    if (-not $isFollowUp) {
        Add-PcmBatchRun $context.Journal $batchId @{ RunId = [string]$context.RunId; Action = $Action; Mode = [string]$effectiveMode; Started = $context.StartTime.ToUniversalTime()
            Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name; ToolVersion = $toolVersion; Source = (Get-PcmSourceText); Overrides = (@($context.ConfigOverrides) -join ' | ')
            Status = [string]$outcome.Status; ExitCode = [int]$outcome.ExitCode; Done = [int]$outcome.Counts.Success; Pending = [int]$outcome.Counts.Pending; Errors = [int]$outcome.Counts.Error
            Skipped = [int]$outcome.Counts.Skipped; LogFile = [string]$context.LogFile }
    }
    if (-not $isFollowUp -and -not [bool](Get-PcmValue $context 'RunReachedEnd' $false)) {
        $open = Get-PcmBatch $context.Journal -Id $batchId
        if ($open -and (Test-PcmSameText ([string]$open.status) 'Running')) {
            $items = @(Get-PcmBatchItem $context.Journal $batchId)
            Set-PcmBatch $context.Journal $batchId 'Partial' @{ done = @($items | Where-Object status -eq 'Done').Count; failed = @($items | Where-Object status -eq 'Failed').Count; running = @($items | Where-Object status -in @('Running', 'Planned')).Count } -Note 'The run ended before it finished: resume the batch with the same command.'
        }
    }
    $convertId = [string](Get-PcmValue $context 'ConvertBatchId' '')
    if (-not $convertId) {
        $row = Get-PcmBatch $context.Journal -Id $batchId
        $convertId = if ($row -and (Test-PcmSameText $row.action 'Recover')) { [string]$row.convert_batch } else { $batchId }
    }
    $history = Get-PcmBatchHistory $context.Journal $convertId
    $report = Write-PcmGlobalReport -Context $context -History $history -ConvertBatchId $convertId -FollowUp @(@(Get-PcmValue $context 'FollowUp' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ })
    if ($report.Html) {
        $context.GlobalReportPath = [string]$report.Html
        $lines = @(('Global report of batch {0} (every run and every object of the batch):' -f $convertId), [string]$report.Html)
        if ($report.EventsCsv) { $lines += @(('Events of batch {0} (CSV, one line per event of every object):' -f $convertId), [string]$report.EventsCsv) }
        $context.NextSteps = @(@($context.NextSteps) + $lines)
        $context.FallbackNext = @(@($context.FallbackNext) + $lines)
    }
}

# =================================================================================================
# Main
# =================================================================================================
try {
    $context.RunReachedEnd = $false
    $context.CurrentOperation = 'Read-Configuration'
    $context.Config = Import-PcmConfiguration -Path $ConfigPath -Root $PSScriptRoot
    $config = $context.Config
    # Run from the window: -ConfigPath is the configuration generated for this run, -BaseConfigPath the configuration file (never modified). What differs is shown everywhere.
    if ($BaseConfigPath) {
        $baseConfig = Import-PcmConfiguration -Path $BaseConfigPath -Root $PSScriptRoot
        $context.ConfigOverrides = @(Get-PcmConfigurationDifference -Base $baseConfig -Effective $config | ForEach-Object { $_.Text })
    }
    $context.LogFolder = $config.Logging.Folder
    $context.ReportFolder = $config.Report.Folder
    $context.NoReport = (-not $config.Report.Enabled) -or [bool]$ResolveOnly

    Initialize-PcmAudit $context

    $context.CurrentOperation = 'Check-Options'
    Test-PcmRunOption
    if ($CsvPath) {
        $context.Identities = @(Read-PcmIdentityFile -Path $CsvPath)
        Write-PcmLog -Context $context -Message ('Identity list {0}: {1} object(s)' -f $CsvPath, $context.Identities.Count) -Level Detail
    }

    $dot = [char]0x00B7
    if (Test-PcmSameText $Action 'Schedule') {
        # The task runs this script with the configuration file that was asked for (the original file when the window generated one for this run), unless it is the default one.
        $taskConfig = if ($script:ConfigPathBound) { Resolve-PcmPath $(if ($BaseConfigPath) { $BaseConfigPath } else { $ConfigPath }) 'ConfigPath' $PSScriptRoot } else { '' }
        $context['ScheduleSetting'] = Get-PcmScheduleSetting -Config $config -At $At -ScriptPath (Join-Path $PSScriptRoot 'Invoke-PraCloudMailbox.ps1') -ConfigPath $taskConfig
    }
    $modeText = if ($Action -eq 'Check') { 'Read-only' } elseif ($effectiveMode -eq 'Apply') { $(if ($Action -eq 'Collect') { 'Apply (the snapshot is written)' } elseif ($Action -eq 'Schedule') { 'Apply (the task is registered)' } else { 'Apply (changes are made)' }) } else { 'Preview (nothing is changed)' }
    $banner = [ordered]@{ 'Action' = @('Mode', $Action); 'Mode' = @($(if ($effectiveMode -eq 'Apply') { 'Write' } else { 'Plan' }), $modeText) }
    $objects = if ($CsvPath) { "list $(Split-Path $CsvPath -Leaf) ($(@($context.Identities).Count) object(s))" } elseif ($Identity) { "one object: $Identity" } elseif ($SearchBase) { "OU $SearchBase" } elseif ($Group) { "group $Group" } else { '' }
    if ($MaxObjects -gt 0 -and $objects) { $objects += " $dot first $MaxObjects" } elseif ($MaxObjects -gt 0) { $objects = "first $MaxObjects objects" }
    if ($Action -eq 'Schedule') { $banner['Task'] = @('Target', ('{0} {1} every day at {2}' -f $context.ScheduleSetting.TaskName, $dot, $context.ScheduleSetting.At)) }
    else { $banner['Scope'] = @('Target', $(if ($Action -eq 'Collect') { Get-PcmScopeText $config } elseif ($Batch) { "batch $Batch" + $(if ($objects) { " $dot $objects" } else { '' }) } else { $(if ($objects) { $objects } else { "snapshot $(if ($Snapshot) { $Snapshot } else { 'last complete' })" + $(if ($Scope -ne 'All') { " $dot $Scope" } else { '' }) }) })) }
    if ($Action -notin @('Collect', 'Schedule') -and $config.Cloud.Organization) { $banner['Tenant'] = @('Cloud', $config.Cloud.Organization) }
    if ($Action -ne 'Schedule') { $banner['Database'] = @('Folder', $config.Store.Path) }
    $skipSnapshotCheck = $AcceptUnverifiedSnapshot -and -not (Test-PcmSameText $Action 'Collect')
    if ($skipSnapshotCheck) {
        $context.SnapshotCheck = 'skipped on request (-AcceptUnverifiedSnapshot)'
        $withoutHash = [IO.File]::Exists($config.Store.Path) -and -not [IO.File]::Exists($config.Store.Path + '.sha256')
        $banner['Integrity'] = @('Key', $(if ($withoutHash) { 'Skipped on request (-AcceptUnverifiedSnapshot): this snapshot cannot be verified because it has no .sha256 file; this run goes on because -AcceptUnverifiedSnapshot was given' } else { 'Skipped on request (-AcceptUnverifiedSnapshot): a snapshot file changed after its Collect does not stop this run' }))
    }
    $banner['Config'] = @('Config', ('{0} {1} Environment {2}' -f (Split-Path $config._Path -Leaf), $dot, $config.Environment))
    if (@($context.ConfigOverrides).Count) { $banner['Settings'] = @('Key', ('{0} setting(s) differ from {1}: {2}{3}' -f @($context.ConfigOverrides).Count, (Split-Path $BaseConfigPath -Leaf), ((@($context.ConfigOverrides) | Select-Object -First 3) -join '; '), $(if (@($context.ConfigOverrides).Count -gt 3) { '; ...' } else { '' }))) }
    $banner['Log'] = @('Log', $context.LogFile)
    Write-PcmBanner -Context $context -Title 'PRA Cloud Mailbox' -Subtitle "Exchange DR $dot scenario 2 $dot on-premises lost $([char]0x2192) Exchange Online" -Details $banner
    foreach ($difference in @($context.ConfigOverrides)) { Write-PcmLog -Context $context -Message ('Run setting changed from the configuration file: {0}' -f $difference) -Level Detail }
    if ($skipSnapshotCheck) { Write-PcmLog -Context $context -Message 'Snapshot check: skipped on request (-AcceptUnverifiedSnapshot).' -Level Detail }

    Write-PcmPlan -Context $context -Titles (Get-PcmPlanTitle)
    if ($Action -notin @('Collect', 'Schedule')) { Test-PcmCloudModuleAvailable }
    switch ($Action) { 'Collect' { Invoke-PcmCollect } 'Schedule' { Invoke-PcmSchedule } 'Check' { if ($ResolveOnly) { Invoke-PcmResolveGroup } elseif ($Batch) { Invoke-PcmCheckBatch } else { Invoke-PcmCheck } } 'Convert' { Invoke-PcmConvert } 'Recover' { Invoke-PcmRecover } }
    $context.RunReachedEnd = $true
}
catch {
    $failure = $_; $failureOperation = $context.CurrentOperation
    # The audit starts here when the run stopped before reaching it (for example on a configuration that cannot be read); an audit that already failed is reported once.
    if (-not $context.TranscriptStarted -and -not $context.LogFile -and $context.LogFolder -and -not [bool](Get-PcmValue $context 'AuditAttempted' $false)) {
        try { Initialize-PcmAudit $context }
        catch {
            $retryMessage = $_.Exception.Message
            if (-not @($context.Issues.ToArray() | Where-Object { ([string](Get-PcmValue $_ 'Message' '')).Contains($retryMessage) }).Count) { $context.Issues.Add([pscustomobject]@{ Message = "Audit unavailable: $retryMessage"; Source = 'Audit' }) }
        }
    }
    $context.CurrentOperation = $failureOperation
    try {
        $known = @($context.Issues.ToArray() | Where-Object { ([string](Get-PcmValue $_ 'Message' '')).Contains($failure.Exception.Message) }).Count -gt 0
        $stopMessage = "Stopped at step '{0}' (operation {1}{2}): {3}" -f $context.CurrentPhase, $context.CurrentOperation, $(if ($context.CurrentIdentity) { ', object ' + $context.CurrentIdentity } else { '' }), $failure.Exception.Message
        Write-PcmLog -Context $context -Message $stopMessage -Level $(if ($known -or [bool](Get-PcmValue $context 'Cancelled' $false)) { 'Detail' } else { 'Error' })
        Write-PcmLog -Context $context -Message ("ErrorId={0} | {1} | {2}" -f $failure.FullyQualifiedErrorId, $failure.InvocationInfo.PositionMessage, $failure.ScriptStackTrace) -Level Debug
    }
    catch { $context.Issues.Add([pscustomobject]@{ Message = $failure.Exception.Message; Source = 'Execution' }) }
    $context.ExitCode = 1
}
finally {
    # A call to Microsoft 365 that did not answer in time stopped the run: say it once, with its message, and end with exit code 1.
    $callTimeout = ''
    if ($Action -ne 'Collect' -and (Get-Command Get-PcmCallTimeout -ErrorAction SilentlyContinue)) { $callTimeout = [string](Get-PcmCallTimeout) }
    if ($callTimeout) {
        try {
            if (-not @($context.Issues.ToArray() | Where-Object { ([string](Get-PcmValue $_ 'Message' '')).Contains($callTimeout) }).Count) { Write-PcmLog -Context $context -Message $callTimeout -Level Error }
            Write-PcmJournal -Step 'MicrosoftTimeout' -Outcome Fail -Detail $callTimeout
        } catch { $null = $_ }
        $context.ExitCode = 1
    }
    # A run that ends without reaching the end of the main block and without an error (Ctrl+C, a closed window, a stop of the host) is never a success.
    $interrupted = $false
    if (-not [bool](Get-PcmValue $context 'RunReachedEnd' $false) -and [int]$context.ExitCode -ne 1) {
        $interrupted = $true
        try {
            $where = "Stopped at step '{0}' (operation {1}{2}): the run was interrupted before it finished (Ctrl+C, a closed window or a stop of the host)." -f $context.CurrentPhase, $context.CurrentOperation, $(if ($context.CurrentIdentity) { ', object ' + $context.CurrentIdentity } else { '' })
            Write-PcmLog -Context $context -Message $where -Level Error
        } catch { $null = $_ }
        $context.ExitCode = 1
    }
    if ($Action -eq 'Collect' -and (Get-Command Disconnect-PcmExchange -ErrorAction SilentlyContinue)) {
        try { Disconnect-PcmExchange -Context $context } catch { $context.Issues.Add([pscustomobject]@{ Message = "Exchange session: $($_.Exception.Message)"; Source = 'Exchange' }) }
    }
    if ($Action -notin @('Collect', 'Schedule') -and (Get-Command Disconnect-PcmCloud -ErrorAction SilentlyContinue)) {
        try { Add-PcmModuleWarning -Warning @($PcmEntryModuleWarnings); $PcmEntryModuleWarnings.Clear(); Disconnect-PcmCloud -Context $context } catch { $context.Issues.Add([pscustomobject]@{ Message = "Cloud sign-out: $($_.Exception.Message)"; Source = 'Cloud' }) }
    }
    # The time limit of the calls goes with the run, whatever stopped it.
    if ($Action -ne 'Collect' -and (Get-Command Stop-PcmCallGuard -ErrorAction SilentlyContinue)) { try { Stop-PcmCallGuard } catch { $null = $_ } }
    # What a failed run proposes when no step set a more specific next step: what it kept, and the command to run again.
    try { $context.FallbackNext = @(Get-PcmFallbackNext -AfterTimeout:([bool]$callTimeout)) }
    catch { $context.FallbackNext = @('Fix the problem named above, then run the same command again with the same options.') }
    try { Write-PcmBatchRecord } catch { try { Write-PcmLog -Context $context -Message ('The record of this run in the journal and the global report could not be written: {0}' -f $_.Exception.Message) -Level Warning } catch { $null = $_ } }
    if ($context.Journal) { try { Close-PcmStore $context.Journal } catch { $null = $_ } }
    $result = Complete-PcmRun -Context $context
    # A run whose pipeline was stopped by the time limit gives its exit code to the process only through the host.
    if ($callTimeout -or $interrupted) { try { $Host.SetShouldExit([int]$result.ExitCode) } catch { $null = $_ } }
    if ($PassThru) { $result }
    exit $result.ExitCode
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZdz/RZH/IxXkp
# wWhRD8uPM1aE61BNUMuHT2C2Wv/CwqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDrJF1Py9/y
# FR1Y1+9/JPHtNcxiuM1InAYec5GI0NiiXDANBgkqhkiG9w0BAQEFAASCAgBgNc2a
# Zf3XXa0D/tUP+QfkaW0dEIIaYKrQFV0swbvGVD06M4IGniv6wXsMS8MxNJiK8gmi
# gq0UWlmvs1IC5ORxSrBEvSpCQdTg/G+i7sEAtD3vGZpyIZz4hxmypoFAtN+lUOrB
# 5Zrawl1Bvnz6UM+PankJBtUHP9I0P6SBSWwWIy0pzxfqwqe26tvW9NxSryfUt5ac
# ZHgtqjkpaF+78ByVZ2K2SaANFpznV+XsCVBFwX1MJqPmUy0QwhnKpP4FVYr9vshk
# Epj1fSNOXhrmTO7/1mZAxdsl1w5PvlYDlzosZPZBWOrLCsW+B1niIk6JX/YJ79G7
# rYVI+YtEEoE+jTWstiZBpBhQn9q3inz1F3vmcNmMiSGNhaOC0RshdG3dSOJ3YNg6
# mQq6Hpg9Yno0Vi6haUpPckveEASdDya1xkJ6ig71B2lzMjB3XBeEan2ZuSlLiOHY
# XnpVoR8x/KHdVJJ90ztGbj/6NakcwkA+GIaGv/88RVv+8NSiawxjXLLKvvcBFGZc
# VQnhYLa8lC78Kw1LwFyS7uVVllqIqE0itwSuNEDF1ag8W1x+3dNbymG/fF7QE3wc
# 3cVbpBoXktWnhwsjf6jVmL+brQ0EAWPSnVtMuq5ICxmrJmqvfvqTTZLbyb+pzWyG
# X/hk6c3uL4QraKtGTxoT+ShD5WKSTbNWqqbK+KGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIO019iMMwVkiXRyKrq2X6qCoZgR5Rl7ldu6SAWra2ku6AggvrXRzIh/v
# 2hgPMjAyNjExMTUyMDA2MDZaMAMCAQECECeseSyF5mrejzdhaAdnfGKgghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTIwMDYwNlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIGSm
# jtOhtq3OV+OIlbeiQnpCYSXhnFGu06Fxa7UnmG0xMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAkYUaLkkUgmmEzV1vWUGy0u0xzZCEJPN0Rk5zYgpnR6ICIGNX
# M5bRw+Tx/A915nIIsOpI8R2fxH0jMgggB1KZP1MT
# SIG # End signature block
