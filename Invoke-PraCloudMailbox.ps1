<#
.SYNOPSIS
    PRA Cloud Mailbox - Exchange disaster recovery, scenario 2: Active Directory, Exchange and Entra Connect
    are lost; users and shared mailboxes get a mailbox in Exchange Online, and everything is rolled back
    when the on-premises infrastructure is rebuilt.

.DESCRIPTION
    Four actions:

      Collect   Before the disaster, on the Exchange side (Windows PowerShell 5.1). Reads the mailboxes in
                scope, their addresses and GUIDs, the shared mailbox permissions (FullAccess, SendAs,
                SendOnBehalf) and optionally the contacts and distribution groups, and stores them as a
                snapshot in the SQLite database (Store.Path). Run it regularly; copy the database to the
                cloud admin server.
      Check     On the cloud admin server (PowerShell 7). Read-only: for every object of the last snapshot,
                checks its Entra ID object and its Exchange Online recipient, the licences and the app
                permissions, and tells whether Convert can run.

      Convert   Disaster (PowerShell 7). Users: ExchangeGuid cleared, source of authority to the cloud, Exchange
                plan (licence group, Kiosk plan or direct licence): the Teams storage becomes the mailbox.
                Shared mailboxes, in waves (Licensing.Shared.Parallel): temporary licence, shared, retention tag,
                licence given back, permissions from the snapshot. Every step is journaled (Store.JournalPath) under a batch ID.
      Recover   Infrastructure rebuilt (PowerShell 7). Users: source of authority back to AD, delta sync, Exchange
                plan removed; the cloud mailbox stays as ComponentShared under the case hold (Retention.HoldPolicy).
                Shared mailboxes: case hold, scheduler paused, identity deleted (the mailbox becomes inactive),
                deleted permanently, recreated from AD by a delta sync.

    Opt-in unsynchronised native rooms/equipment (CloudResource): Collect captures metadata, room lists,
    booking/delivery controls and work hours, never old contents or general mailbox rights. Convert creates
    required lists first, then each unlicensed, disabled-sign-in resource and its membership. New journals
    use protectionMode None: Recover has NO hold/NO inactive copy and does NOT preserve cloud-created
    reservations/messages. Export/reconcile and explicitly set Resources.AllowUnprotectedDeletion and
    OnPremisesReady to $true before deleting only the journal-owned cloud identity. Soft-deleted addresses
    remain Pending; no purge or hold/licence removal. Ambiguous deletion requests and restored identities
    are never automatically deleted again. Old protected journals and legacy synced paths are unchanged.

    Every run starts in Preview. Collect -Mode Apply writes the snapshot (nothing is ever written to
    Exchange). Check is always read-only. Convert and Recover -Mode Apply ask one confirmation (or none
    with -Force); an interrupted batch is resumed with the same -Batch.

.PARAMETER Action
    Collect, Check, Convert or Recover.

.PARAMETER Mode
    Preview (default): read and report. Apply: Collect writes the snapshot to the database.

.PARAMETER Identity
    One mailbox only (UPN, SMTP address, sAMAccountName, DN or GUID) instead of the configured scope.

.PARAMETER IdentityPath
    Check, Convert and Recover: a wave - the objects listed in this file, one per line (UPN, SMTP address,
    sAMAccountName, DN or GUID; # starts a comment), or a CSV file with an Identity column.

.PARAMETER Scope
    All (default), UsersOnly or SharedOnly (shared, room and equipment mailboxes).

.PARAMETER Snapshot
    Check and Convert: snapshot ID (default: the last complete snapshot).

.PARAMETER Batch
    Convert: resume this batch. Recover: the Convert batch to roll back (printed at the end of Convert).

.PARAMETER ConfigPath
    Configuration file. Default: config\PraCloudMailbox.config.psd1 next to this script.

.PARAMETER Force
    No confirmation prompt; never bypasses native-resource deletion consent or routing guards.
    Collect asks one only before writing a snapshot of ONE object (-Identity -Mode Apply).

.PARAMETER PassThru
    Also returns the result of the run as an object (Status, ExitCode, counters, file paths).

.PARAMETER Gui
    Opens the window: the state of the configuration, the snapshot and the batches, and every action with its
    preview, a typed confirmation and the progress. It runs this script for each action (Windows PowerShell 5.1
    for Collect, PowerShell 7 for the others), so the window and the command line do exactly the same.

.EXAMPLE
    .\Invoke-PraCloudMailbox.ps1 -Gui
    Opens the window (Windows PowerShell 5.1 on an Exchange server, or PowerShell 7 on the cloud admin server).

.EXAMPLE
    .\Invoke-PraCloudMailbox.ps1 -Action Collect
    Windows PowerShell 5.1 on an Exchange server: preview of what would be collected.

.EXAMPLE
    .\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply -Force
    Writes a new snapshot (scheduled task).

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check
    PowerShell 7 on the cloud admin server: readiness of every object of the last snapshot.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply
    Disaster: converts every ready object of the last snapshot; prints the batch ID.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch 1a2b3c4d
    Infrastructure rebuilt: rolls back the Convert batch 1a2b3c4d.

.EXAMPLE
    pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch 1a2b3c4d -IdentityPath .\data\wave1.txt
    Rolls back only the objects of the batch listed in wave1.txt (a wave).

.NOTES
    Author     : Nicolas Fabert
    Version    : 1.3.0 (UNRELEASED)
    Requires   : Collect: Windows PowerShell 5.1 and the Exchange cmdlets (local or remote).
                 Check, Convert, Recover: PowerShell 7.4+, Microsoft.Graph.Authentication, ExchangeOnlineManagement 3.10+.
    Exit codes : 0 = done, 1 = failed (or an object is not ready), 2 = done, next step required.
    Documentation : docs\PraCloudMailbox-UserGuide.md (user guide), docs\PraCloudMailbox-Guide.md (developer guide)
#>
#Requires -Version 5.1
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Force', Justification = 'Read by Invoke-PraCollect (script scope).')]
[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Run')]
param(
    [Parameter(Mandatory, ParameterSetName = 'Run')][ValidateSet('Collect','Check','Convert','Recover')][string]$Action,
    [Parameter(ParameterSetName = 'Run')][ValidateSet('Preview','Apply')][string]$Mode = 'Preview',
    [Parameter(ParameterSetName = 'Run')][string]$Identity,
    [Parameter(ParameterSetName = 'Run')][string]$IdentityPath,
    [Parameter(ParameterSetName = 'Run')][ValidateSet('All','UsersOnly','SharedOnly')][string]$Scope = 'All',
    [Parameter(ParameterSetName = 'Run')][ValidateRange(0, [long]::MaxValue)][long]$Snapshot = 0,
    [Parameter(ParameterSetName = 'Run')][string]$Batch,
    [string]$ConfigPath,
    [Parameter(ParameterSetName = 'Run')][switch]$Force,
    [Parameter(ParameterSetName = 'Run')][switch]$PassThru,
    [Parameter(Mandatory, ParameterSetName = 'Gui')][switch]$Gui
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$toolVersion = '1.3.0'
# Not a parameter default: Windows PowerShell 5.1 leaves $PSScriptRoot empty in param() defaults when run with -File (scheduled task).
if (-not $ConfigPath) { $ConfigPath = Join-Path $PSScriptRoot 'config\PraCloudMailbox.config.psd1' }

# The window: any edition (it runs this script for each action, in the edition of the action).
if ($Gui) {
    Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Common.psm1') -Force -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Store.psm1') -Force -ErrorAction Stop
    Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Gui.psm1') -Force -ErrorAction Stop
    try { Show-PraGui -Root $PSScriptRoot -ConfigPath $ConfigPath -Version $toolVersion; exit 0 }
    catch { Write-Error ('The window could not open: {0}' -f $_.Exception.Message) -ErrorAction Continue; exit 1 }
}

# Each action runs in one edition: Collect needs the Exchange cmdlets (Windows PowerShell 5.1), the cloud actions PowerShell 7.
if ($Action -eq 'Collect' -and $PSVersionTable.PSEdition -ne 'Desktop') {
    Write-Error 'Collect runs in Windows PowerShell 5.1 (powershell.exe), where the Exchange cmdlets live. Open Windows PowerShell and run the same command.'
    exit 1
}
if ($Action -ne 'Collect' -and ($PSVersionTable.PSEdition -ne 'Core' -or $PSVersionTable.PSVersion -lt [version]'7.4')) {
    Write-Error "$Action runs in PowerShell 7.4 or later (pwsh.exe). Open PowerShell 7 and run the same command."
    exit 1
}
Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Common.psm1') -Force -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Store.psm1') -Force -ErrorAction Stop
if ($Action -eq 'Collect') { Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Collect.psm1') -Force -ErrorAction Stop }
else { Import-Module (Join-Path $PSScriptRoot 'module\PRA2.Cloud.psm1') -Force -ErrorAction Stop }

$effectiveMode = if ($WhatIfPreference -or $Action -eq 'Check') { 'Preview' } else { $Mode }
$caller = $PSCmdlet
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
}

function Skip-PraStep {
    <# Prints a step that does not run in this execution (numbering stays in the announced order). #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Reason, [string]$Icon = 'Skip')
    Write-PraStep -Context $context -Title $Title -Icon $Icon
    Write-PraItem -Context $context -Status Skip -Text $Reason
}

function New-PraRow {
    <# One result row (CSV / HTML report). #>
    param([Parameter(Mandatory)][object]$Record)
    return [ordered]@{
        Identity = $(if ($Record.user_principal_name) { [string]$Record.user_principal_name } else { [string]$Record.primary_smtp_address })
        Kind = [string]$Record.kind; PrimarySmtpAddress = [string]$Record.primary_smtp_address; ObjectGuid = [string]$Record.object_guid
        ExchangeGuid = [string]$Record.exchange_guid; Action = $Action; Entra = ''; ExchangeOnline = ''; TeamsStorage = ''; Licence = ''; Holds = ''
        Permissions = ''; FinalStatus = 'Planned'; Detail = ''; Warnings = ''
    }
}

function Get-PraScopeText {
    param([hashtable]$Config)
    $dot = [char]0x00B7
    if ($Identity) { return "one object: $Identity" }
    $s = $Config.Scope
    $base = switch ($s.Mode) { 'OU' { $s.SearchBase } 'Group' { "members of $($s.GroupDN)" } 'Csv' { "CSV $($s.CsvPath)" } default { if ($s.SearchBase) { $s.SearchBase } else { 'whole organisation' } } }
    $kinds = @(); if ($s.IncludeUsers) { $kinds += 'users' }; if ($s.IncludeShared) { $kinds += 'shared' }; if ($s.IncludeRoom) { $kinds += 'rooms' }; if ($s.IncludeEquipment) { $kinds += 'equipment' }
    if ($Scope -eq 'UsersOnly') { $kinds = @('users only') } elseif ($Scope -eq 'SharedOnly') { $kinds = @('shared mailboxes only') }
    return ('{0} {1} {2}' -f $base, $dot, ($kinds -join ' + '))
}

function Read-PraIdentityFile {
    <# -IdentityPath: one identity per line (# starts a comment), or a CSV file with an Identity column. Duplicates removed. #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "-IdentityPath: file not found: $Path" }
    $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
    if ($lines.Count -and $lines[0] -match '^"?Identity"?\s*([,;]|$)') {
        $delimiter = if ($lines[0] -match ';') { ';' } else { ',' }
        $lines = @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -Encoding UTF8 | ForEach-Object { ([string]$_.Identity).Trim() } | Where-Object { $_ })
    }
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($line in $lines) { if ($seen.Add($line)) { $list.Add($line) } }
    if (-not $list.Count) { throw "-IdentityPath: no identity in $Path." }
    return $list.ToArray()
}

function Get-PraIdentityList {
    <# The objects asked for: -IdentityPath (read by the main block into the context) or -Identity; none = the whole scope. #>
    $stored = Get-PraValue $context 'Identities' $null
    $list = @(@($stored) | Where-Object { $_ })
    if (-not $list.Count -and $Identity) { $list = @($Identity) }
    return $list
}

function Get-PraIdentityArgument {
    <# The identity part of a command line shown as the next step. #>
    if ($IdentityPath) { return (' -IdentityPath "{0}"' -f $IdentityPath) }
    if ($Identity) { return " -Identity $Identity" }
    return ''
}

function Select-PraIdentity {
    <# Objects whose keys match the identity list (case-insensitive); the identities found nowhere are reported. #>
    param([AllowEmptyCollection()][object[]]$Items = @(), [Parameter(Mandatory)][scriptblock]$Keys, [Parameter(Mandatory)][string]$Where)
    $wanted = @(Get-PraIdentityList)
    if (-not $wanted.Count) { return $Items }
    $set = [System.Collections.Generic.HashSet[string]]::new([string[]]$wanted, [StringComparer]::OrdinalIgnoreCase)
    $found = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $selected = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Items) {
        $hit = $false
        foreach ($key in @(& $Keys $item)) { if ($key -and $set.Contains([string]$key)) { [void]$found.Add([string]$key); $hit = $true } }
        if ($hit) { $selected.Add($item) }
    }
    if (-not $selected.Count) { throw $(if ($wanted.Count -eq 1) { "$($wanted[0]) is not in $Where." } else { "None of the $($wanted.Count) identities of the list is in $Where." }) }
    $missing = @($wanted | Where-Object { -not $found.Contains($_) })
    if ($missing.Count) { Write-PraLog -Context $context -Message ('{0} identity(ies) of the list not in {1}: {2}{3}' -f $missing.Count, $Where, ((@($missing | Select-Object -First 10)) -join ', '), $(if ($missing.Count -gt 10) { ', ...' } else { '' })) -Level Warning }
    return $selected.ToArray()
}

function Test-PraStop {
    <#
    Stop asked by the window (PRA_STOP_FILE): no new object is started. The caller leaves the objects not started as
    Pending (Set-PraRowStopped); the run then ends normally (scheduler resumed, journal, report) and can be resumed.
    #>
    if (-not (Test-PraStopRequest)) { return $false }
    if (-not (Get-PraValue $context 'StopNoted' $false)) {
        $context['StopNoted'] = $true
        Write-PraLog -Context $context -Message 'Stop requested by the operator: no new object is started; the objects not started stay as they are (Pending). Run the same command again to go on.' -Level Warning
    }
    return $true
}

function Set-PraRowStopped {
    <# An object not started because of a stop request: Pending, without a console line each (there can be thousands). #>
    param([Parameter(Mandatory)][object]$Row)
    $Row.FinalStatus = 'Pending'
    $Row.Detail = 'not started: stop requested by the operator; run the same command again to go on'
}

# =================================================================================================
# Collect - Exchange on-premises, before the disaster (Windows PowerShell 5.1).
# =================================================================================================
function Invoke-PraCollect {
    $config = $context.Config
    $collect = $config.Collect
    $dot = [char]0x00B7

    Write-PraStep -Context $context -Title 'Exchange on-premises connection' -Icon Directory
    $context.CurrentOperation = 'Connect-Exchange'
    $how = Connect-Pra2Exchange -Context $context
    Write-PraItem -Context $context -Status Ok -Text "Connected: $how"
    $dc = [string]$config.Exchange.DomainController
    Write-PraItem -Context $context -Status Info -Text $(if ($dc) { "Domain controller for every read: $dc" } else { 'Domain controller: chosen by Exchange (set Exchange.DomainController to pin one)' })

    Write-PraStep -Context $context -Title 'Mailboxes in scope' -Icon Target
    $context.CurrentOperation = 'Get-ScopeMailbox'
    $mailboxes = @(Get-Pra2ScopeMailbox -Context $context -Identity $Identity -Filter $Scope)
    $records = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($mailbox in $mailboxes) {
        $index++
        $context.CurrentIdentity = [string]$mailbox.PrimarySmtpAddress
        $statistics = $null
        if ($collect.MailboxStatistics) {
            try { $statistics = Get-MailboxStatistics -Identity ([string]$mailbox.DistinguishedName) -WarningAction SilentlyContinue -ErrorAction Stop }
            catch { Write-PraLog -Context $context -Message ("Statistics of {0}: {1}" -f $mailbox.PrimarySmtpAddress, $_.Exception.Message) -Level Detail }
        }
        $recordParameters = @{ Mailbox = $mailbox; Statistics = $statistics }
        if ([string]$mailbox.RecipientTypeDetails -in @('RoomMailbox','EquipmentMailbox') -and
            [string](Get-PraValue $config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource') {
            $context.CurrentOperation = 'Get-ResourceSettings'
            try {
                $resourceSettings = Get-Pra2ResourceSettings -Context $context -Mailbox $mailbox
                if (-not $resourceSettings) { throw 'No resource settings were returned.' }
                $recordParameters.ResourceSettings = $resourceSettings
            } catch {
                throw ("Collect stopped before writing the snapshot: resource settings of {0} could not be collected. Booking settings must not be silently omitted or broadened. {1}" -f $mailbox.PrimarySmtpAddress, $_.Exception.Message)
            }
        }
        [void]$records.Add((ConvertTo-Pra2MailboxRecord @recordParameters))
    }
    $context.CurrentIdentity = ''
    $byKind = $records | Group-Object { $_.kind }
    $text = if ($records.Count) { ($byKind | ForEach-Object { '{0} {1}' -f $_.Count, $_.Name.ToLowerInvariant() }) -join " $dot " } else { 'none' }
    Write-PraItem -Context $context -Status $(if ($records.Count) { 'Ok' } else { 'Warn' }) -Icon $(if ($records.Count) { 'People' } else { '' }) -Text ("{0} mailbox(es): {1}" -f $records.Count, $text)
    foreach ($excluded in $context.Excluded) { Write-PraLog -Context $context -Message ("Excluded: $excluded") -Level Sub }
    $noAddress = @($records | Where-Object { -not $_.primary_smtp_address })
    if ($noAddress.Count) { Write-PraLog -Context $context -Message ('{0} mailbox(es) without primary SMTP address.' -f $noAddress.Count) -Level Warning }

    $permissions = [System.Collections.Generic.List[object]]::new()
    $members = [System.Collections.Generic.List[object]]::new()
    $permissionText = @{}
    $sharedRecords = @($records | Where-Object { $_.kind -ne 'User' -and -not (
                $_.kind -in @('Room','Equipment') -and [string](Get-PraValue $config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource') })
    if ($collect.SharedPermissions -and $sharedRecords.Count) {
        Write-PraStep -Context $context -Title 'Shared mailbox permissions' -Icon Key
        $context.CurrentOperation = 'Get-SharedPermission'
        $permissionMailboxGuids = @{}
        foreach ($record in $sharedRecords) { $permissionMailboxGuids[[string]$record.object_guid] = $true }
        foreach ($mailbox in @($mailboxes | Where-Object { $permissionMailboxGuids.ContainsKey([string]$_.Guid) })) {
            $context.CurrentIdentity = [string]$mailbox.PrimarySmtpAddress
            $result = Get-Pra2SharedPermission -Context $context -Mailbox $mailbox
            foreach ($row in $result.Permissions) { [void]$permissions.Add($row) }
            foreach ($row in $result.Members) { [void]$members.Add($row) }
            foreach ($warning in $result.Warnings) { Write-PraLog -Context $context -Message $warning -Level Warning }
            $mine = @($result.Permissions)
            $counts = @{ FullAccess = 0; SendAs = 0; SendOnBehalf = 0 }
            foreach ($row in $mine) { $counts[$row.access_right]++ }
            $permissionText[[string]$mailbox.Guid] = 'FullAccess {0} {3} SendAs {1} {3} SendOnBehalf {2}' -f $counts.FullAccess, $counts.SendAs, $counts.SendOnBehalf, $dot
            $groups = @($mine | Where-Object { $_.trustee_kind -eq 'Group' }).Count
            Write-PraLog -Context $context -Message ('{0}: {1}{2}' -f $mailbox.PrimarySmtpAddress, $permissionText[[string]$mailbox.Guid], $(if ($groups) { " $dot $groups group(s) expanded" } else { '' })) -Level Sub
        }
        $context.CurrentIdentity = ''
        $unresolved = @($permissions | Where-Object { -not $_.resolved }).Count
        Write-PraItem -Context $context -Status $(if ($unresolved) { 'Warn' } else { 'Ok' }) -Text ('{0} permission(s) on {1} shared mailbox(es), {2} group member(s) expanded{3}' -f $permissions.Count, $sharedRecords.Count, $members.Count, $(if ($unresolved) { " $dot $unresolved trustee(s) not resolved" } else { '' }))
    } else {
        Skip-PraStep -Title 'Shared mailbox permissions' -Reason $(if (-not $collect.SharedPermissions) { 'Collect.SharedPermissions = $false.' } else { 'No shared mailbox in scope.' }) -Icon Key
    }

    $contacts = @(); $groupResult = @{ Groups = @(); Members = @() }
    if (($collect.Contacts -or $collect.DistributionGroups) -and -not $Identity) {
        Write-PraStep -Context $context -Title 'Contacts and distribution groups' -Icon Mail
        if ($collect.Contacts) {
            $context.CurrentOperation = 'Get-MailContact'
            $contacts = @(Get-Pra2Contact -Context $context)
            Write-PraItem -Context $context -Status Ok -Text ('{0} mail contact(s)' -f $contacts.Count)
        }
        if ($collect.DistributionGroups) {
            $context.CurrentOperation = 'Get-DistributionGroup'
            $groupResult = Get-Pra2DistributionGroup -Context $context
            Write-PraItem -Context $context -Status Ok -Text ('{0} distribution group(s), {1} membership(s)' -f @($groupResult.Groups).Count, @($groupResult.Members).Count)
        }
    } else {
        Skip-PraStep -Title 'Contacts and distribution groups' -Reason $(if ($Identity) { 'One object (-Identity): not collected.' } else { 'Collect.Contacts and Collect.DistributionGroups = $false.' }) -Icon Mail
    }

    foreach ($record in $records) {
        $row = New-PraRow ([pscustomobject]$record)
        $row.Permissions = [string]$permissionText[[string]$record.object_guid]
        $row.Detail = 'X500 {0} {1} {2} address(es)' -f $(if ($record.legacy_exchange_dn) { 'kept' } else { 'missing' }), $dot, @(($record.email_addresses_json | ConvertFrom-Json) | ForEach-Object { $_ }).Count
        [void]$context.Rows.Add($row)
    }

    $counts = @{ mailbox = $records.Count; permission = $permissions.Count; group_member = $members.Count; contact = $contacts.Count
        distribution_group = @($groupResult.Groups).Count; dl_member = @($groupResult.Members).Count; excluded = $context.Excluded.Count }
    if ($context.Mode -ne 'Apply') {
        Skip-PraStep -Title 'Saving the snapshot' -Reason 'Preview: not written (add -Mode Apply).' -Icon Backup
        $context.NextSteps = @(('.\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply' + $(if ($Identity) { " -Identity $Identity" } else { '' })))
        return
    }
    Write-PraStep -Context $context -Title 'Saving the snapshot' -Icon Backup
    if ($Identity) {
        # A one-object snapshot becomes the last complete snapshot: Check and Convert would only see this object.
        if (-not $Force) {
            $question = "Write a snapshot of ONE object ($Identity)? It becomes the last snapshot: Check and Convert will only see this object until the next full Collect."
            $answer = $false
            try { $answer = $caller.ShouldContinue($question, 'PRA Cloud Mailbox - Collect') }
            catch { throw "This Collect needs a confirmation and this console cannot ask for one ($($_.Exception.Message)). Add -Force." }
            if (-not $answer) { throw 'Cancelled by the operator: nothing was written.' }
        }
        Write-PraLog -Context $context -Message 'Snapshot of ONE object (-Identity): Check and Convert use the last complete snapshot, so a full Collect must follow.' -Level Warning
    }
    $context.CurrentOperation = 'Write-Snapshot'
    $connection = Open-Pra2Store -Path $config.Store.Path -Root $context.Root
    try {
        $scopeJson = @{ Mode = $config.Scope.Mode; SearchBase = $config.Scope.SearchBase; GroupDN = $config.Scope.GroupDN; CsvPath = $config.Scope.CsvPath
            Identity = $Identity; Filter = $Scope; IncludeUsers = $config.Scope.IncludeUsers; IncludeShared = $config.Scope.IncludeShared
            IncludeRoom = $config.Scope.IncludeRoom; IncludeEquipment = $config.Scope.IncludeEquipment } | ConvertTo-Json -Compress
        $id = New-Pra2Snapshot $connection @{ RunId = $context.RunId; Environment = $config.Environment; Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            ToolVersion = $toolVersion; ExchangeServer = $context.Server; ScopeJson = $scopeJson }
        try {
            $null = Add-Pra2Row $connection mailbox $id $records.ToArray()
            $null = Add-Pra2Row $connection permission $id $permissions.ToArray()
            $null = Add-Pra2Row $connection group_member $id $members.ToArray()
            $null = Add-Pra2Row $connection contact $id $contacts
            $null = Add-Pra2Row $connection distribution_group $id @($groupResult.Groups)
            $null = Add-Pra2Row $connection dl_member $id @($groupResult.Members)
            $issues = @(@(Get-PraValue $context 'WarningList' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ } | ForEach-Object { [ordered]@{ level = 'Warning'; object = ''; operation = [string]$_.Phase; message = [string]$_.Message } })
            $null = Add-Pra2Row $connection collect_issue $id $issues
            Set-Pra2SnapshotStatus $connection $id Complete $counts
        } catch {
            try { Set-Pra2SnapshotStatus $connection $id Failed $counts $_.Exception.Message } catch { $null = $_ }
            throw
        }
        $context.SnapshotLabel = '{0} {1} {2}' -f $id, $dot, (Get-Date -Format 'yyyy-MM-dd HH:mm')
        Write-PraItem -Context $context -Status Ok -Icon Batch -Text ('Snapshot {0} written to {1}' -f $id, $config.Store.Path)
        $pruned = Remove-Pra2OldSnapshot $connection $config.Store.KeepSnapshots
        if ($pruned) { Write-PraItem -Context $context -Status Info -Text ('{0} old snapshot(s) deleted (Store.KeepSnapshots = {1})' -f $pruned, $config.Store.KeepSnapshots) }
        $backup = Backup-Pra2Store $connection $config.Store.BackupFolder ('PraCloudMailbox_{0}_snapshot{1}_{2}.db' -f $config.Environment, $id, (Get-Date -Format 'yyyyMMdd-HHmmss'))
        [void]$context.BackupFiles.Add($backup)
        Write-PraItem -Context $context -Status Ok -Icon Backup -Text "Copy of the database: $backup"
    } finally { Close-Pra2Store $connection }
    foreach ($row in $context.Rows) { $row.FinalStatus = 'Success' }
    $context.NextSteps = @("Copy $($config.Store.Path) to the cloud admin server (or let the backup folder be replicated), then:", 'pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check')
}

# =================================================================================================
# Snapshot of the Collect (Check, Convert, Recover).
# =================================================================================================
function Read-PraSnapshotData {
    <# Loads one snapshot (default: the last complete one), applies -Identity and -Scope (or keeps -ObjectGuid only), prints the step. #>
    param([long]$Id = 0, [string]$Verb = 'check', [switch]$All, [AllowEmptyCollection()][string[]]$ObjectGuid, [switch]$NoCount)
    $config = $context.Config
    $dot = [char]0x00B7
    Write-PraStep -Context $context -Title 'Collect snapshot' -Icon Batch
    $context.CurrentOperation = 'Read-Snapshot'
    $connection = Open-Pra2Store -Path $config.Store.Path -Root $context.Root -ReadOnly
    try {
        $snap = Get-Pra2Snapshot $connection $Id
        if (-not $snap) { throw $(if ($Id) { "Snapshot $Id not found in $($config.Store.Path)." } else { "No complete snapshot in $($config.Store.Path): run Collect -Mode Apply on the Exchange side." }) }
        if ($snap.status -ne 'Complete') { throw "Snapshot $($snap.id) is $($snap.status), not Complete." }
        $records = @(Read-Pra2Table $connection mailbox $snap.id)
        $permissions = @(Read-Pra2Table $connection permission $snap.id)
        $members = @(Read-Pra2Table $connection group_member $snap.id)
    } finally { Close-Pra2Store $connection }
    $taken = [datetime]::Parse([string]$snap.finished_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    $age = (Get-Date).ToUniversalTime() - $taken.ToUniversalTime()
    $context.SnapshotLabel = '{0} {1} {2:yyyy-MM-dd HH:mm} UTC {1} {3}' -f $snap.id, $dot, $taken.ToUniversalTime(), $snap.exchange_server
    $context.Server = [string]$snap.exchange_server
    $scopeInfo = $null; try { $scopeInfo = [string]$snap.scope_json | ConvertFrom-Json } catch { $scopeInfo = $null }
    $partial = $scopeInfo -and [string](Get-PraValue $scopeInfo 'Identity' '')
    Write-PraItem -Context $context -Status Ok -Text ('Snapshot {0} of {1:yyyy-MM-dd HH:mm} UTC from {2} ({3} mailbox(es), {4} permission(s))' -f $snap.id, $taken.ToUniversalTime(), $snap.exchange_server, $records.Count, $permissions.Count)
    if ($partial) { Write-PraLog -Context $context -Message ('Snapshot {0} holds one object only ({1}): not a full Collect.' -f $snap.id, $scopeInfo.Identity) -Level Warning }
    if ($age.TotalDays -gt $config.Store.MaxSnapshotAgeDays) { Write-PraLog -Context $context -Message ('Snapshot {0:0} day(s) old (Store.MaxSnapshotAgeDays = {1}): permissions and addresses may have changed since.' -f $age.TotalDays, $config.Store.MaxSnapshotAgeDays) -Level Warning }
    if ($PSBoundParameters.ContainsKey('ObjectGuid')) { $records = @($records | Where-Object { [string]$_.object_guid -in @($ObjectGuid) }) }
    else {
        if (-not $All) {
            $records = @(Select-PraIdentity -Items $records -Where "snapshot $($snap.id)" -Keys { param($r) @([string]$r.user_principal_name, [string]$r.primary_smtp_address, [string]$r.sam_account_name, [string]$r.object_guid, [string]$r.distinguished_name) })
        }
        if (-not $All -and $Scope -eq 'UsersOnly') { $records = @($records | Where-Object { $_.kind -eq 'User' }) }
        elseif (-not $All -and $Scope -eq 'SharedOnly') { $records = @($records | Where-Object { $_.kind -ne 'User' }) }
    }
    $users = @($records | Where-Object { $_.kind -eq 'User' }).Count
    if (-not $NoCount) { Write-PraItem -Context $context -Status Info -Icon Target -Text ('{0} object(s) to {3}: {1} user(s), {2} shared/room/equipment' -f $records.Count, $users, ($records.Count - $users), $Verb) }
    return @{ Snapshot = $snap; Records = $records; Permissions = $permissions; Members = $members; Users = $users }
}

# =================================================================================================
# Check - cloud admin server, read-only (PowerShell 7).
# =================================================================================================
function Invoke-PraCheck {
    $config = $context.Config

    $data = Read-PraSnapshotData -Id $Snapshot -Verb 'check'
    $records = @($data.Records); $permissions = @($data.Permissions); $members = @($data.Members)
    $context.NativeResourceScope = @($records | Where-Object { Test-Pra2NativeResource -Config $config -Record $_ }).Count -gt 0
    $nativeOnly = $records.Count -gt 0 -and @($records | Where-Object { -not (Test-Pra2NativeResource -Config $config -Record $_) }).Count -eq 0

    Write-PraStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-Pra2Cloud -Context $context } catch { $context.CloudConnectFailed = $true; throw }
    Write-PraItem -Context $context -Status Ok -Text $how
    if ($context.ContainsKey('CertificateWarning')) { Write-PraLog -Context $context -Message $context['CertificateWarning'] -Level Warning }

    Write-PraStep -Context $context -Title 'Tenant prerequisites' -Icon Key
    $context.CurrentOperation = 'Tenant-Facts'
    $skipUserGroup = @($records | Where-Object { $_.kind -eq 'User' }).Count -eq 0 -and @($records | Where-Object { Test-Pra2NativeResource -Config $config -Record $_ }).Count -gt 0
    $tenant = Get-Pra2TenantFact -Context $context -SkipUserGroup:$skipUserGroup -NativeOnly:$nativeOnly
    $roles = @(Get-Pra2GraphRole)
    # The state of every object first: the licence units really needed depend on what each one already holds.
    if (Test-PraStop) { throw 'Check stopped by the operator before reading the cloud state.' }
    $states = Read-PraCloudStates -Records $records
    $need = Get-PraLicenceNeed -Entries @($records | ForEach-Object { $s = $states[[string]$_.object_guid]; [pscustomobject]@{ Record = $_; State = $(if ($s -is [System.Management.Automation.ErrorRecord]) { $null } else { $s }); Step = '' } }) -Tenant $tenant
    $tenantFindings = @(Test-Pra2TenantReadiness -Config $config -Roles $roles -Tenant $tenant -UserCount $need.Users -SharedCount $need.Shared -ResourceCount $need.Resources -SkipUserLicensing:$need.SkipUserLicensing -NativeOnly:$nativeOnly)
    $tenantFindings += @(Get-PraResourceLicenceFinding -Need $need -Tenant $tenant)
    foreach ($finding in $tenantFindings) {
        switch ($finding.Level) {
            'Error' { Write-PraLog -Context $context -Message $finding.Message -Level Error }
            'Warn' { Write-PraLog -Context $context -Message $finding.Message -Level Warning }
            'Ok' { Write-PraItem -Context $context -Status Ok -Text $finding.Message }
            default { Write-PraItem -Context $context -Status Info -Text $finding.Message }
        }
    }
    Write-PraItem -Context $context -Status Info -Text ("Recover needs Entra Connect {0} or later on the rebuilt server (it keeps the cloud source of authority of the objects); this version is not readable from the cloud." -f $config.EntraConnect.MinVersion)

    Write-PraStep -Context $context -Title 'Objects' -Icon People
    $context.CurrentOperation = 'Object-Readiness'
    $lookup = New-PraTrusteeLookup -States $states -Records $records
    $permissionIndex = Get-PraPermissionIndex -Permissions $permissions -Members $members
    $ready = 0
    foreach ($record in $records) {
        $row = New-PraRow $record
        $context.CurrentIdentity = $row.Identity
        try {
            $state = $states[[string]$record.object_guid]
            if ($state -is [System.Management.Automation.ErrorRecord]) { throw $state.Exception }
            $findings = @(Get-Pra2Readiness -Config $config -Record $record -State $state -Skus @($tenant.Skus))
            if (Get-PraValue $state 'NativeResource' $false) {
                try { Test-Pra2ResourceConfigurationReadiness -Record $record -AllowRoomListCreation -RoomListPlan @((Read-Pra2ResourceSettings $record).RoomLists) }
                catch { $findings += [pscustomobject]@{ Level = 'Error'; Field = 'ExchangeOnline'; Code = 'RESOURCE_CONFIGURATION_REQUIRED'; Message = $_.Exception.Message } }
            } elseif ($record.kind -ne 'User') {
                $subset = Get-PraPermissionSubset -Index $permissionIndex -MailboxGuid ([string]$record.object_guid)
                $findings += @(Get-Pra2TrusteeFinding -Record $record -Permissions $subset.Permissions -Members $subset.Members -Lookup $lookup)
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
        Write-PraItem -Context $context -Status $(if ($status -eq 'Warn') { 'Warn' } elseif ($status -eq 'Fail') { 'Fail' } else { 'Ok' }) -Text ('{0} ({1}): {2}' -f $row.Identity, $kind, $summary)
        foreach ($finding in @($findings | Where-Object { $_.Level -ne 'Ok' -and -not [object]::ReferenceEquals($_, $main) })) { Write-PraLog -Context $context -Message ('{0}: {1}' -f $finding.Code, $finding.Message) -Level Sub }
        [void]$context.Rows.Add($row)
    }
    $context.CurrentIdentity = ''
    Write-PraItem -Context $context -Status $(if ($ready -eq $records.Count) { 'Ok' } else { 'Info' }) -Icon Target -Text ('{0} of {1} object(s) ready for Convert' -f $ready, $records.Count)
    $context.NextSteps = @('Disaster: pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert' + $(if (Get-PraIdentityArgument) { Get-PraIdentityArgument } elseif ($Scope -ne 'All') { " -Scope $Scope" } else { '' }) + '   # preview first, then -Mode Apply')
}

# =================================================================================================
# Convert and Recover - cloud admin server (PowerShell 7). Procedure validated in the lab (developer guide, chapter 3).
# =================================================================================================
function Confirm-PraApply {
    <# One confirmation for the whole run, after the plan is shown (none with -Force). #>
    param([Parameter(Mandatory)][string]$Question)
    if ($context.Mode -ne 'Apply' -or $Force) { return }
    $answer = $false
    try { $answer = $caller.ShouldContinue($Question, 'PRA Cloud Mailbox - Apply') }
    catch { throw "Apply needs a confirmation and this console cannot ask for one ($($_.Exception.Message)). Add -Force." }
    if (-not $answer) { throw 'Cancelled by the operator: nothing was changed.' }
    Write-PraLog -Context $context -Message ('Operator confirmed: ' + $Question) -Level Detail
}

function Write-PraJournal {
    <# One event of the batch (journal + log) and, optionally, the new state of the object. #>
    param([string]$ObjectGuid = '', [Parameter(Mandatory)][string]$Step, [ValidateSet('Ok','Fail','Info')][string]$Outcome = 'Ok', [string]$Detail = '', [hashtable]$Item)
    if ($context.Journal -and $context.BatchId) {
        Add-Pra2BatchEvent $context.Journal $context.BatchId -ObjectGuid $ObjectGuid -Step $Step -Outcome $Outcome -Detail $Detail
        # A resumed object that ends Done must not keep the error message of an earlier run.
        if ($Item -and $Item['status'] -eq 'Done' -and -not $Item.ContainsKey('message')) { $Item['message'] = '' }
        if ($Item -and $ObjectGuid) { Set-Pra2BatchItem $context.Journal $context.BatchId $ObjectGuid $Item }
    }
    Write-PraLog -Context $context -Message ('[{0}] {1} {2}{3}' -f $context.BatchId, $Step, $Outcome, $(if ($Detail) { ': ' + $Detail } else { '' })) -Level Detail
}

function Get-PraWaitSeconds {
    param([Parameter(Mandatory)][ValidateSet('MailboxTimeoutMinutes','SyncTimeoutMinutes','HoldTimeoutMinutes')][string]$Name)
    return [int]$context.Config.Polling[$Name] * 60
}

function Get-PraOriginalLicence {
    <# Direct licences recorded by Convert in original_json (none for a batch written by 0.2.0). #>
    param([AllowNull()][object]$Item)
    if (-not $Item -or -not $Item.original_json) { return @() }
    try { $original = [string]$Item.original_json | ConvertFrom-Json } catch { return @() }
    if (-not $original.PSObject.Properties['licences']) { return @() }
    return @($original.licences | ForEach-Object { $_ } | Where-Object { $_ })
}

function Get-PraLicenceNeed {
    <#
    Objects that will take a NEW licence unit: users without the SKU of their Exchange plan (Group, Direct; Kiosk uses
    the licence they have), shared mailboxes without the temporary SKU that are not converted yet. An object that already
    holds the SKU (hybrid user with Exchange disabled, resume) takes no unit: a user consumes one unit per SKU.
    Entries: Record, State (Get-Pra2CloudState) and Step (journal step on resume, else '').
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
        if ((Get-PraValue $entry.State 'NativeResource' $false) -or (Test-Pra2NativeResource -Config $config -Record $entry.Record)) {
            $intent = if ($context.ContainsKey('ResourceJournal')) { Get-Pra2ResourceIntent -Item $context.ResourceJournal[[string]$entry.Record.object_guid] } else { $null }
            if ((Get-Pra2ResourceProtectionMode $intent) -eq 'None') { continue }
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
                    (Get-PraValue $_.State 'NativeResource' $false) -or (Test-Pra2NativeResource -Config $config -Record $_.Record)
                }).Count -gt 0) }
}

function Get-PraResourceLicenceFinding {
    param([object]$Need, [object]$Tenant)
    $groups = @{}; $demand = @{}
    $userSkus = @(switch ($context.Config.Licensing.Users.Mode) {
            Group { if ($Tenant.LicenceGroup) { $Tenant.LicenceGroup.assignedLicenses | ForEach-Object { [string]$_.skuId } } }
            Direct { $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$context.Config.Licensing.Users.SkuPartNumber } | ForEach-Object { [string]$_.skuId } }
        })
    foreach ($id in $Need.ResourceGroups.Keys) {
        try {
            if (-not $id) { throw 'Resources.LicensingGroupId is empty: native resources need enterprise licensing throughout their active lifetime.' }
            $group = Get-Pra2ResourceGroup -Id $id -Skus @($Tenant.Skus)
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
        if ($demand[$skuId] -gt 0 -and (Get-Pra2FreeUnit $sku) -lt $demand[$skuId]) {
            [pscustomobject]@{ Level = 'Error'; Code = 'RESOURCE_LICENCE'; Message = "Resource licensing capacity insufficient: $($sku.skuPartNumber) needs $($demand[$skuId]) new permanent unit(s) across selected users and resources." }
        }
    }
    foreach ($id in $groups.Keys) {
        [pscustomobject]@{ Level = 'Ok'; Code = 'RESOURCE_LICENCE'; Message = "Resource licence group ${id}: enterprise plan checked ($($Need.ResourceGroups[$id]) new permanent unit(s)); resources do not recycle temporary shared units." }
    }
}

function Get-PraObjectAssessment {
    <# Cloud state (read in bulk beforehand, or now) + readiness findings of one snapshot object (Check rules). #>
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][object]$Tenant, [object[]]$Permissions = @(), [object[]]$Members = @(), [scriptblock]$Lookup, [object]$State)
    if ($State -is [System.Management.Automation.ErrorRecord]) { throw $State }
    $state = if ($State) { $State } else { Get-Pra2CloudState -Context $context -Record $Record }
    $findings = @(Get-Pra2Readiness -Config $context.Config -Record $Record -State $state -Skus @($Tenant.Skus))
    if (Get-PraValue $state 'NativeResource' $false) {
        if ($context.Mode -eq 'Preview') {
            try { Test-Pra2ResourceConfigurationReadiness -Record $Record -AllowRoomListCreation -RoomListPlan @((Read-Pra2ResourceSettings $Record).RoomLists) -RoomListIntents (Get-PraValue $context 'RoomListIntents' @{}) }
            catch { $findings += [pscustomobject]@{ Level = 'Error'; Field = 'ExchangeOnline'; Code = 'RESOURCE_CONFIGURATION_REQUIRED'; Message = $_.Exception.Message } }
        }
    } elseif ($Record.kind -ne 'User' -and $Lookup) { $findings += @(Get-Pra2TrusteeFinding -Record $Record -Permissions $Permissions -Members $Members -Lookup $Lookup) }
    return [pscustomobject]@{ State = $state; Findings = $findings; Errors = @($findings | Where-Object Level -eq 'Error'); Warnings = @($findings | Where-Object Level -eq 'Warn') }
}

function Get-PraPermissionIndex {
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

function Get-PraPermissionSubset {
    <# The permissions of one mailbox and the members of the groups they name. #>
    param([Parameter(Mandatory)][object]$Index, [Parameter(Mandatory)][string]$MailboxGuid)
    $permissions = @(if ($Index.ByMailbox.ContainsKey($MailboxGuid)) { $Index.ByMailbox[$MailboxGuid] })
    $groups = @($permissions | Where-Object { [string]$_.trustee_kind -eq 'Group' } | ForEach-Object { [string]$_.trustee_guid } | Select-Object -Unique)
    $members = @(foreach ($group in $groups) { if ($Index.ByGroup.ContainsKey($group)) { $Index.ByGroup[$group] } })
    return [pscustomobject]@{ Permissions = $permissions; Members = $members }
}

function New-PraTrusteeLookup {
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

function Read-PraCloudStates {
    <# The cloud state of every object in bulk (Get-Pra2CloudStateSet), with its duration in the console. #>
    param([AllowEmptyCollection()][object[]]$Records = @())
    if (-not $Records.Count) { return @{} }
    $context.CurrentOperation = 'Read-States'
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $states = Get-Pra2CloudStateSet -Context $context -Records $Records
    Write-PraItem -Context $context -Status Info -Text ('Cloud state of {0} object(s) read in {1}' -f $Records.Count, (Format-PraDuration $clock.Elapsed.TotalSeconds))
    return $states
}

function Wait-PraExoStateSet {
    <#
    Polls the Exchange Online state of many objects together (Get-Pra2ExoStateSet) until -Test is true for each or the
    timeout; -OnReady runs for each object as soon as it is ready. Returns the entries still not ready.
    Entries need an EntraId property.
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][scriptblock]$Test, [scriptblock]$OnReady, [Parameter(Mandatory)][int]$TimeoutSeconds, [switch]$StopOnRequest)
    $waiting = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $Entries) { $waiting.Add($entry) }
    if (-not $waiting.Count) { return @() }
    # The poll runs in the scope of Wait-Pra2Condition, whose own parameter is $Test: other names here.
    $readyTest = $Test; $readyAction = $OnReady; $honourStop = [bool]$StopOnRequest
    $readyClock = [Diagnostics.Stopwatch]::StartNew()
    $null = Wait-Pra2Condition -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
        $states = Get-Pra2ExoStateSet -Id ([string[]]@($waiting | ForEach-Object { $_.EntraId })) -TagAttribute ([string]$context.Config.Retention.TagAttribute)
        foreach ($one in @($waiting)) {
            $state = $states[[string]$one.EntraId]
            if ($state -and (& $readyTest $state $one)) { [void]$waiting.Remove($one); if ($readyAction) { & $readyAction $one $state ([int]$readyClock.Elapsed.TotalSeconds) } }
        }
        (-not $waiting.Count) -or ($honourStop -and (Test-PraStop))
    }
    return $waiting.ToArray()
}

function Find-PraRecordRow {
    param([Parameter(Mandatory)][string]$ObjectGuid)
    return ($context.Rows | Where-Object { $_.ObjectGuid -eq $ObjectGuid } | Select-Object -First 1)
}

function Set-PraRowResult {
    <# Final status and detail of one report row, plus the console line. #>
    param([Parameter(Mandatory)][object]$Row, [Parameter(Mandatory)][ValidateSet('Success','Error','Pending','Skipped','Planned')][string]$Status, [string]$Detail = '', [string]$Field, [string]$Value)
    $Row.FinalStatus = $Status
    if ($Detail) { $Row.Detail = $Detail }
    if ($Field) { $Row[$Field] = $Value }
    $item = @{ Success = 'Ok'; Error = 'Fail'; Pending = 'Warn'; Skipped = 'Skip'; Planned = 'Info' }[$Status]
    if ($Status -eq 'Error') { Write-PraLog -Context $context -Message ('{0}: {1}' -f $Row.Identity, $Detail) -Level Error }
    else { Write-PraItem -Context $context -Status $item -Text ('{0}: {1}' -f $Row.Identity, $Detail) }
}

function Stop-PraEntry {
    <# An object that fails in a phase: journal, report row, and it leaves the next phases. #>
    param([Parameter(Mandatory)][object]$Entry, [Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Message)
    $text = $Message -replace '\s+', ' '
    Write-PraJournal -ObjectGuid ([string]$Entry.Record.object_guid) -Step $Step -Outcome Fail -Detail $text -Item @{ status = 'Failed'; message = $text }
    Set-PraRowResult -Row $Entry.Row -Status Error -Detail "Convert stopped: $text"
}

function Invoke-PraIdentityStartSet {
    <#
    Common start of a conversion, phase by phase for many objects (the order validated in the lab, one object at a
    time before 1.2): ExchangeGuid cleared, source of authority in the cloud (Graph batches), ONE wait for all until
    Exchange Online no longer sees them as synchronised (23 to 45 s in the lab), usageLocation set from the cloud
    (Graph batches). An object that fails is reported and left out. Returns the entries that went through.
    Entries: Record, Row, EntraId, State (Get-Pra2CloudState).
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @())
    $config = $context.Config
    if (-not $Entries.Count) { return @() }
    $exo = Get-Pra2ExoStateSet -Id ([string[]]@($Entries | ForEach-Object { $_.EntraId })) -TagAttribute $config.Retention.TagAttribute
    $active = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $Entries) {
        $context.CurrentIdentity = $entry.Row.Identity
        $state = $exo[$entry.EntraId]
        try {
            if ($state -and $state.Type -eq 'MailUser' -and $state.ExchangeGuid -ne [guid]::Empty.ToString()) {
                Set-MailUser -Identity $entry.EntraId -ExchangeGuid ([guid]::Empty) -ErrorAction Stop
                Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'GuidCleared' -Detail ("ExchangeGuid {0} cleared" -f $state.ExchangeGuid) -Item @{ step = 'GuidCleared'; status = 'Running' }
            }
            $active.Add($entry)
        } catch { Stop-PraEntry -Entry $entry -Step 'GuidCleared' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    $behaviors = Get-Pra2SyncBehaviorSet -UserId ([string[]]@($active | ForEach-Object { $_.EntraId }))
    $toCloud = @($active | Where-Object { $behaviors[$_.EntraId] -ne $true })
    $answers = Invoke-Pra2GraphBatch -Requests @($toCloud | ForEach-Object { @{ Method = 'PATCH'; Url = ('/users/{0}/onPremisesSyncBehavior' -f $_.EntraId); Body = @{ isCloudManaged = $true } } })
    for ($i = 0; $i -lt $toCloud.Count; $i++) {
        if ($answers[$i].Ok) { Write-PraJournal -ObjectGuid ([string]$toCloud[$i].Record.object_guid) -Step 'SoaCloud' -Detail 'isCloudManaged = true' -Item @{ step = 'SoaCloud'; status = 'Running' } }
        else { Stop-PraEntry -Entry $toCloud[$i] -Step 'SoaCloud' -Message ('source of authority not transferred ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$active.Remove($toCloud[$i]) }
    }
    $seconds = @{}
    $late = @(Wait-PraExoStateSet -Entries $active.ToArray() -TimeoutSeconds (Get-PraWaitSeconds 'SyncTimeoutMinutes') -Test { param($state) $state.IsDirSynced -eq $false } -OnReady {
            $seconds[$args[0].EntraId] = $args[2] })
    foreach ($entry in $late) { Stop-PraEntry -Entry $entry -Step 'SoaCloud' -Message 'Exchange Online still sees the object as synchronised from AD after the source of authority transfer.'; [void]$active.Remove($entry) }
    # Set from the cloud even when present: the value synchronised from AD is refused right after the transfer (test SH2).
    $countries = @{}
    foreach ($entry in @($active)) {
        $current = if ($entry.State -and $entry.State.User) { [string]$entry.State.User.usageLocation } else { '' }
        $country = if ($current) { $current } else { [string]$config.Cloud.DefaultUsageLocation }
        if (-not $country) { Stop-PraEntry -Entry $entry -Step 'UsageLocation' -Message 'usageLocation is empty and Cloud.DefaultUsageLocation is not set (a licence needs a country).'; [void]$active.Remove($entry); continue }
        $countries[$entry.EntraId] = $country
    }
    $set = $active.ToArray()
    $answers = Invoke-Pra2GraphBatch -Requests @($set | ForEach-Object { @{ Method = 'PATCH'; Url = ('/users/{0}' -f $_.EntraId); Body = @{ usageLocation = $countries[$_.EntraId] } } })
    for ($i = 0; $i -lt $set.Count; $i++) {
        $entry = $set[$i]
        if ($answers[$i].Ok) {
            Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'UsageLocation' -Detail ("usageLocation = {0} (cloud), Exchange Online no longer synchronised after {1} s" -f $countries[$entry.EntraId], $seconds[$entry.EntraId]) -Item @{ step = 'UsageLocation'; status = 'Running' }
        } else { Stop-PraEntry -Entry $entry -Step 'UsageLocation' -Message ('usageLocation not set ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$active.Remove($entry) }
    }
    return $active.ToArray()
}

function Invoke-PraUserPlanSet {
    <#
    Users: identity start of the wave (Invoke-PraIdentityStartSet), then the Exchange plan: membership of the licence
    group by Graph batches (already a member is not an error), or the Kiosk / direct plan one user at a time.
    Returns the users whose plan is assigned (their mailboxes are awaited together afterwards).
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$Tenant)
    $config = $context.Config
    $users = $config.Licensing.Users
    $started = @(Invoke-PraIdentityStartSet -Entries $Entries)
    $done = [System.Collections.Generic.List[object]]::new()
    $assigned = {
        param($entry, $detail)
        Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'PlanAssigned' -Detail $detail -Item @{ step = 'PlanAssigned'; status = 'Running' }
        Write-PraItem -Context $context -Status Ok -Text ('{0}: source of authority in the cloud, Exchange plan assigned' -f $entry.Row.Identity)
        $done.Add($entry)
    }
    if ($users.Mode -eq 'Group') {
        # Members already (resume) are not added again.
        $checks = Invoke-Pra2GraphBatch -Requests @($started | ForEach-Object { @{ Method = 'POST'; Url = ('/users/{0}/checkMemberGroups' -f $_.EntraId); Body = @{ groupIds = @($users.GroupId) } } })
        $toAdd = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $started.Count; $i++) {
            if ($checks[$i].Ok -and @((Get-PraValue $checks[$i].Body 'value' @()) | ForEach-Object { $_ }) -contains $users.GroupId) { & $assigned $started[$i] "member of the licence group $($users.GroupId)" }
            else { $toAdd.Add($started[$i]) }
        }
        $answers = Invoke-Pra2GraphBatch -Requests @($toAdd | ForEach-Object { @{ Method = 'POST'; Url = ('/groups/{0}/members/$ref' -f $users.GroupId); Body = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($_.EntraId)" } } })
        for ($i = 0; $i -lt $toAdd.Count; $i++) {
            if ($answers[$i].Ok -or $answers[$i].Error -match 'already exist') { & $assigned $toAdd[$i] "member of the licence group $($users.GroupId)" }
            else { Stop-PraEntry -Entry $toAdd[$i] -Step 'ConvertUser' -Message ('not added to the licence group ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error) }
        }
        return $done.ToArray()
    }
    foreach ($entry in $started) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            if ($users.Mode -eq 'Kiosk') {
                $kiosk = Get-Pra2KioskAssignment -Licence (Get-Pra2UserLicence -UserId $entry.EntraId) -Skus @($Tenant.Skus)
                if (-not $kiosk) { throw 'Kiosk mode: no licence of this user contains EXCHANGE_S_DESKLESS.' }
                # Kiosk enabled in the existing direct assignment (other plans unchanged), or a direct assignment next to the group one.
                $how = Enable-Pra2LicencePlan -UserId $entry.EntraId -SkuId $kiosk.SkuId -PlanIds @($kiosk.KioskPlanId) -DisabledPlans $kiosk.DisabledPlans
                & $assigned $entry "$($kiosk.SkuName): Exchange Kiosk $($how.ToLowerInvariant())"
            } else {
                $sku = $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$users.SkuPartNumber } | Select-Object -First 1
                if (-not $sku) { throw "Licensing.Users.SkuPartNumber $($users.SkuPartNumber) not found in the tenant." }
                $how = Enable-Pra2LicencePlan -UserId $entry.EntraId -SkuId $sku.skuId -PlanIds @(Get-Pra2MailboxPlanId -Skus @($sku)) -DisabledPlans (Get-Pra2DisabledPlanForExchangeOnly -Sku $sku)
                & $assigned $entry "$($sku.skuPartNumber): Exchange plan $($how.ToLowerInvariant())"
            }
        } catch { Stop-PraEntry -Entry $entry -Step 'ConvertUser' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    return $done.ToArray()
}

function Invoke-PraConvertSharedWave {
    <#
    A wave of shared mailboxes, phase by phase (one temporary unit each): identity start, temporary licence, ONE wait
    for all the mailboxes, Type Shared asked again until it is effective (the licence must stay until then, test SH2),
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
    $exo = Get-Pra2ExoStateSet -Id (& $ids $Entries) -TagAttribute $tagAttribute
    $already = @($Entries | Where-Object { $exo[$_.EntraId] -and $exo[$_.EntraId].Type -eq 'SharedMailbox' })
    $toConvert = @($Entries | Where-Object { $_ -notin $already })
    $licensed = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in @(Invoke-PraIdentityStartSet -Entries $toConvert)) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            $how = Enable-Pra2LicencePlan -UserId $entry.EntraId -SkuId $sku.skuId -PlanIds @(Get-Pra2MailboxPlanId -Skus @($sku)) -DisabledPlans (Get-Pra2DisabledPlanForExchangeOnly -Sku $sku)
            if ($how -ne 'Present') { Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'LicenceAssigned' -Detail "temporary $($sku.skuPartNumber) (Exchange plan only, $($how.ToLowerInvariant()))" -Item @{ step = 'LicenceAssigned'; status = 'Running' } }
            $licensed.Add($entry)
        } catch { Stop-PraEntry -Entry $entry -Step 'LicenceAssigned' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    $late = @(Wait-PraExoStateSet -Entries $licensed.ToArray() -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -Test { param($state) $state.Type -in @('UserMailbox', 'SharedMailbox') } -OnReady {
            Write-PraJournal -ObjectGuid ([string]$args[0].Record.object_guid) -Step 'MailboxCreated' -Detail "cloud mailbox after $($args[2]) s" })
    foreach ($entry in $late) { Stop-PraEntry -Entry $entry -Step 'MailboxCreated' -Message 'No cloud mailbox after the temporary licence (the licence is kept: run Convert -Batch again later).'; [void]$licensed.Remove($entry) }
    # Type Shared for all, asked again every 15 s until it is effective (10 min).
    $pending = [System.Collections.Generic.List[object]]::new($licensed)
    $ready = [System.Collections.Generic.List[object]]::new($already)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while ($pending.Count) {
        $states = Get-Pra2ExoStateSet -Id (& $ids $pending) -TagAttribute $tagAttribute
        foreach ($entry in @($pending)) {
            if ($states[$entry.EntraId] -and $states[$entry.EntraId].Type -eq 'SharedMailbox') {
                Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'Shared' -Detail ("SharedMailbox after {0} s" -f [int]$clock.Elapsed.TotalSeconds) -Item @{ step = 'Shared'; status = 'Running' }
                [void]$pending.Remove($entry); $ready.Add($entry)
            } else { try { Set-Mailbox -Identity $entry.EntraId -Type Shared -ErrorAction Stop } catch { $null = $_ } }
        }
        if (-not $pending.Count -or $clock.Elapsed.TotalSeconds -ge 600) { break }
        Start-Sleep -Seconds 15
    }
    foreach ($entry in $pending) { Stop-PraEntry -Entry $entry -Step 'Shared' -Message 'The mailbox is not a SharedMailbox after 10 min (the licence is kept).' }
    # Tag, then the temporary licence back to what each object had before Convert.
    $states = Get-Pra2ExoStateSet -Id (& $ids $ready) -TagAttribute $tagAttribute
    $back = [System.Collections.Generic.List[object]]::new()
    $removed = 0
    foreach ($entry in $ready) {
        $context.CurrentIdentity = $entry.Row.Identity
        $guid = [string]$entry.Record.object_guid
        try {
            if (-not $states[$entry.EntraId] -or $states[$entry.EntraId].Tag -ne $config.Retention.TagValue) {
                $tagParameters = @{ Identity = $entry.EntraId; ErrorAction = 'Stop' }; $tagParameters[$tagAttribute] = $config.Retention.TagValue
                Set-Mailbox @tagParameters
            }
            Write-PraJournal -ObjectGuid $guid -Step 'Tagged' -Detail "$tagAttribute = $($config.Retention.TagValue)" -Item @{ step = 'Tagged'; status = 'Running' }
            $how = Restore-Pra2Licence -UserId $entry.EntraId -SkuId $sku.skuId -Original @($Originals[$guid])
            if ($how -in @('Removed', 'Restored')) { $removed++; Write-PraJournal -ObjectGuid $guid -Step 'LicenceRemoved' -Detail "temporary $($sku.skuPartNumber) $($how.ToLowerInvariant())" -Item @{ step = 'LicenceRemoved'; status = 'Running' } }
            $back.Add($entry)
        } catch { Stop-PraEntry -Entry $entry -Step 'Tagged' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
    # One pause for the wave: a mailbox converted too early would be disabled once its licence is gone (test SH2).
    if ($removed) { Start-Sleep -Seconds 60 }
    $final = Get-Pra2ExoStateSet -Id (& $ids $back) -TagAttribute $tagAttribute
    foreach ($entry in $back) {
        $context.CurrentIdentity = $entry.Row.Identity
        $guid = [string]$entry.Record.object_guid
        try {
            $state = $final[$entry.EntraId]
            if (-not $state -or $state.Type -ne 'SharedMailbox') { throw "After the licence removal the object is $(if ($state) { $state.Type } else { 'not found' }), not a SharedMailbox." }
            $subset = Get-PraPermissionSubset -Index $PermissionIndex -MailboxGuid $guid
            $grants = @(Get-Pra2SharedGrant -Record $entry.Record -Permissions $subset.Permissions -Members $subset.Members)
            $granted = 0; $failed = [System.Collections.Generic.List[string]]::new()
            foreach ($grant in $grants) {
                try { $outcome = Grant-Pra2SharedPermission -Mailbox $entry.EntraId -Grant $grant; $granted++; Write-PraJournal -ObjectGuid $guid -Step 'Permission' -Detail ('{0} {1} ({2}, {3})' -f $grant.Right, $grant.Trustee, $outcome, $grant.Source) }
                catch { [void]$failed.Add(('{0} {1}: {2}' -f $grant.Right, $grant.Trustee, ($_.Exception.Message -replace '\s+', ' '))); Write-PraJournal -ObjectGuid $guid -Step 'Permission' -Outcome Fail -Detail $failed[$failed.Count - 1] }
            }
            $permissionText = '{0}/{1} permission(s) granted' -f $granted, $grants.Count
            Write-PraJournal -ObjectGuid $guid -Step 'Done' -Detail $permissionText -Item @{ step = 'Done'; status = $(if ($failed.Count) { 'Running' } else { 'Done' }); result_json = (@{ cloudGuid = $state.ExchangeGuid; failed = $failed.ToArray() } | ConvertTo-Json -Compress) }
            $entry.Row.Permissions = $permissionText
            Set-PraRowResult -Row $entry.Row -Status $(if ($failed.Count) { 'Pending' } else { 'Success' }) -Detail $(if ($failed.Count) { 'shared mailbox ready; permissions not granted: ' + ($failed -join '; ') } else { "shared mailbox $($state.ExchangeGuid), tagged, $permissionText" }) -Field 'ExchangeOnline' -Value 'SharedMailbox'
        } catch { Stop-PraEntry -Entry $entry -Step 'ConvertShared' -Message $_.Exception.Message }
    }
    $context.CurrentIdentity = ''
}


function Write-PraResourceJournal {
    param([object]$Entry, [string]$Step, [string]$Status = 'Running', [string]$Detail = '', [hashtable]$Extra = @{})
    $guid = [string]$(if ($Entry.PSObject.Properties['Record']) { $Entry.Record.object_guid } else { $Entry.Item.object_guid })
    $result = @{ lifecycle = 'CloudResource'; cloudId = [string]$Entry.Intent.cloudId; cloudGuid = [string]$Entry.Intent.cloudGuid
        cloudType = [string]$Entry.Intent.cloudType; ownedName = [string]$Entry.Intent.creationName; markerAttribute = [string]$Entry.Intent.markerAttribute; token = [string]$Entry.Intent.token }
    foreach ($key in $Extra.Keys) { $result[$key] = $Extra[$key] }
    Write-PraJournal -ObjectGuid $guid -Step $Step -Detail $Detail -Item @{ step = $Step; status = $Status; entra_id = [string]$Entry.Intent.cloudId
        original_json = ($Entry.Intent | ConvertTo-Json -Depth 15 -Compress); result_json = ($result | ConvertTo-Json -Depth 15 -Compress) }
}

function ConvertTo-PraRoomListSignature {
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

function Get-PraRoomListPlan {
    param([object[]]$Entries, [object[]]$Records)
    $lists = @{}
    $keys = @{}
    if (-not $Entries.Count) { return @() }
    $mailboxKeys = @{}
    foreach ($record in $Records) {
        foreach ($key in @(Get-Pra2ResourceKeys $record)) { $mailboxKeys[$key] = [string]$record.primary_smtp_address }
    }
    foreach ($entry in $Entries) {
        foreach ($list in (Read-Pra2ResourceSettings $entry.Record).RoomLists) {
            $source = [string]$list.SourceGuid
            if ($lists.ContainsKey($source)) {
                if ((ConvertTo-PraRoomListSignature $lists[$source]) -cne (ConvertTo-PraRoomListSignature $list)) { throw "Conflicting room-list snapshot metadata for $source; nothing is changed." }
                continue
            }
            foreach ($key in @(Get-Pra2ResourceKeys (ConvertTo-Pra2RoomListRecord $list))) {
                if ($keys.ContainsKey($key)) { throw "Competing room lists claim $key; nothing is changed." }
                if ($mailboxKeys.ContainsKey($key)) { throw "Room list $($list.PrimarySmtpAddress) conflicts with selected mailbox $($mailboxKeys[$key]); nothing is changed." }
                $keys[$key] = $source
            }
            $lists[$source] = $list
        }
    }
    return @($lists.Values | Sort-Object PrimarySmtpAddress)
}

function Write-PraRoomListJournal {
    param([object]$Intent, [string]$Step, [string]$Status = 'Running', [string]$Detail = '')
    $Intent.status = $Status
    Set-Pra2TenantChange -Connection $context.Journal -BatchId $context.BatchId -Kind 'NativeRoomList' -TargetId ([string]$Intent.sourceGuid) `
        -OriginalJson (ConvertTo-Json -InputObject $Intent -Depth 15 -Compress) -UpdateOriginal
    Write-PraJournal -ObjectGuid ("RoomList:{0}" -f $Intent.sourceGuid) -Step $Step -Outcome $(if ($Status -eq 'Failed') { 'Fail' } else { 'Ok' }) -Detail $Detail
}

function Write-PraRoomListIdentityPin {
    param([object]$Intent, [object]$Group)
    $guid = [string](Get-PraValue $Group 'Guid' '')
    $id = [string](Get-PraValue $Group 'ExternalDirectoryObjectId' '')
    if (-not $guid -or $guid -eq [guid]::Empty.ToString() -or $guid -eq [string]$Intent.sourceGuid -or
        ($Intent.cloudGuid -and [string]$Intent.cloudGuid -ne $guid) -or ($Intent.cloudId -and [string]$Intent.cloudId -ne $id)) {
        throw 'Room-list immutable creation identity is missing/inconsistent; no configuration.'
    }
    $Intent.cloudGuid = $guid
    if ($id) { $Intent.cloudId = $id }
    Write-PraRoomListJournal -Intent $Intent -Step 'RoomListIdentityPinned' -Detail "Exchange GUID $guid; directory ID recorded only when exposed"
}

function Write-PraResourceIdentityPin {
    param([object]$Entry, [object]$Mailbox)
    $intent = $Entry.Intent
    $guid = [string](Get-PraValue $Mailbox 'ExchangeGuid' '')
    $id = [string](Get-PraValue $Mailbox 'ExternalDirectoryObjectId' '')
    if (-not $id -or -not $guid -or $guid -eq [guid]::Empty.ToString() -or $guid -eq [string]$Entry.Record.exchange_guid -or
        ($intent.cloudId -and [string]$intent.cloudId -ne $id) -or ($intent.cloudGuid -and [string]$intent.cloudGuid -ne $guid)) {
        throw 'Resource immutable creation identity is missing/inconsistent; no configuration.'
    }
    $intent.cloudId = $id; $intent.cloudGuid = $guid
    $Entry.EntraId = $id
    Write-PraResourceJournal -Entry $Entry -Step 'ResourceIdentityPinned' -Detail "cloud ID $id, ExchangeGuid $guid recorded before further reads or mutations"
}

function Invoke-PraConvertRoomLists {
    param([object[]]$Lists, [switch]$Unprotected)
    foreach ($list in $Lists) {
        if (Test-PraStop) { break }
        $context.CurrentIdentity = [string]$list.PrimarySmtpAddress
        $intent = $context.RoomListIntents[[string]$list.SourceGuid]
        try {
            if ($Unprotected) { Assert-Pra2ResourceOrganizationUnheld }
            if ($intent) {
                if ((ConvertTo-PraRoomListSignature $intent.list) -cne (ConvertTo-PraRoomListSignature $list)) { throw 'Room-list metadata differs from its creation journal; nothing is changed.' }
            } else {
                $existing = Get-Pra2ResourceCloudRoomList -List $list -AllowMissing
                $token = [guid]::NewGuid().ToString('N')
                $intent = [pscustomobject]@{ sourceGuid = [string]$list.SourceGuid; lifecycle = 'CloudRoomList'; list = $list
                    operation = $(if ($existing) { 'Reuse' } else { 'Create' }); status = 'Planned'; creationAttempted = $false
                    token = $token; creationName = ('PRA-L-{0}-{1}-{2}' -f $context.BatchId.Substring(0,8), $token.Substring(0,8), $token)
                    cloudGuid = $(if ($existing) { [string]$existing.Guid } else { '' }); cloudId = $(if ($existing) { [string](Get-PraValue $existing 'ExternalDirectoryObjectId' '') } else { '' }) }
                $context.RoomListIntents[[string]$list.SourceGuid] = $intent
                Write-PraRoomListJournal -Intent $intent -Step 'RoomListIntent' -Detail "$($intent.operation) $($list.PrimarySmtpAddress); selected native rooms only"
            }
            $group = Get-Pra2ResourceCloudRoomList -List $list -Intent $intent -AllowMissing
            $wasDone = $intent.status -eq 'Done'
            if ($intent.operation -eq 'Create' -and $intent.status -ne 'Done') {
                Test-Pra2RoomListCreationReadiness -List $list
                if (-not $group) {
                    if ($intent.creationAttempted) { throw 'Room-list creation was already attempted but its exact object is not visible; resume only after visibility is resolved, never issue a second create.' }
                    $intent.creationAttempted = $true
                    Write-PraRoomListJournal -Intent $intent -Step 'RoomListCreationIntent' -Detail "create $($intent.creationName); token recorded before New-DistributionGroup"
                    $created = @(New-DistributionGroup -Name $intent.creationName -Alias $list.Alias -DisplayName $list.DisplayName `
                        -PrimarySmtpAddress $list.PrimarySmtpAddress -RoomList -IgnoreNamingPolicy `
                        -MemberJoinRestriction Closed -MemberDepartRestriction Closed -RequireSenderAuthenticationEnabled $true -ErrorAction Stop)
                    if ($created.Count -ne 1) { throw 'Room-list creation response is missing/ambiguous; never repeat creation.' }
                    $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force
                    Write-PraRoomListIdentityPin -Intent $intent -Group $created[0]
                    $group = Get-Pra2ResourceCloudRoomList -List $list -Intent $intent
                }
                if (-not $intent.cloudGuid) { $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force }
                Write-PraRoomListIdentityPin -Intent $intent -Group $group
                Set-Pra2RoomListConfiguration -List $list -Intent $intent -Group $group
                $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $false -Force
                $group = Get-Pra2ResourceCloudRoomList -List $list -Intent $intent
            }
            if (-not $group) { throw 'Room list is not visible/verified; dependent rooms are not created.' }
            Write-PraRoomListIdentityPin -Intent $intent -Group $group
            if ($wasDone) { $intent.status = 'Done' }
            $null = Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop
            Write-PraRoomListJournal -Intent $intent -Step 'RoomListReady' -Status Done -Detail "$($intent.operation) RoomList $($list.PrimarySmtpAddress), cloud GUID $($group.Guid); members added only after room provisioning"
            Write-PraItem -Context $context -Status Ok -Text "Room list ready: $($list.PrimarySmtpAddress) ($($intent.operation))"
        } catch {
            $message = $_.Exception.Message
            if ($intent) { Write-PraRoomListJournal -Intent $intent -Step 'ConvertRoomList' -Status Failed -Detail $message }
            Write-PraLog -Context $context -Message ("Room list {0} failed; dependent rooms are blocked: {1}" -f $list.PrimarySmtpAddress, $message) -Level Error
        }
    }
    $context.CurrentIdentity = ''
}

function Invoke-PraConvertResourceWave {
    param([object[]]$Entries, [object[]]$NativeEntries = @(), [object[]]$NativeRecords = @())
    $intents = @{}
    foreach ($native in $NativeEntries) { $intents[[string]$native.Record.object_guid] = $native.Intent }
    foreach ($entry in $Entries) {
        $context.CurrentIdentity = $entry.Row.Identity
        $provisioningId = ''
        try {
            $intent = $entry.Intent
            if (-not $intent -or $intent.lifecycle -ne 'CloudResource') { throw 'Resource creation intent is missing; do not guess an existing cloud identity.' }
            $protected = (Get-Pra2ResourceProtectionMode $intent) -eq 'CaseHold'
            $owned = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
            if ($owned) {
                $provisioningId = $owned.Id
                Set-Pra2ResourceProvisioningDisabled -Identity $provisioningId
            }
            if (-not $protected) { Assert-Pra2ResourceOrganizationUnheld }
            Test-Pra2ResourceConfigurationReadiness -Record $entry.Record -FutureResources $NativeRecords -ResourceIntents $intents -RoomListIntents $context.RoomListIntents
            if (-not $owned) {
                if ($intent.cloudId -or $intent.cloudGuid) { throw 'Journal-owned resource is missing; use Recover, never replace its identity.' }
                if (Get-PraValue $intent 'creationAttempted' $false) { throw 'Resource creation was already attempted but its exact object is not visible; never issue a second create.' }
                $collisions = @(Get-Pra2ResourceCollisions -Record $entry.Record -Inventory (Get-Pra2ResourceInventory -Records @($entry.Record)))
                if ($collisions.Count) { throw ($collisions -join '; ') }
                if ($protected) {
                    $skus = @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value)
                    $null = Get-Pra2ResourceGroup -Id $intent.licensingGroupId -Skus $skus -NewUnits 1
                }
                # Persist BEFORE New-Mailbox. The creation Name contains the journal's unpredictable token.
                $intent | Add-Member -NotePropertyName creationAttempted -NotePropertyValue $true -Force
                Write-PraResourceJournal -Entry $entry -Step 'ResourceCreationIntent' -Detail "create $($intent.creationName) ($($intent.cloudType)); ownership token recorded before creation"
                $parameters = @{ Name = [string]$intent.creationName; Alias = [string](Get-PraValue $entry.Record 'alias' (($entry.Record.primary_smtp_address -split '@')[0]))
                    DisplayName = [string](Get-PraValue $entry.Record 'display_name' $entry.Row.Identity); PrimarySmtpAddress = [string]$entry.Record.primary_smtp_address; ErrorAction = 'Stop' }
                if (-not $parameters.Alias) { $parameters.Alias = ($entry.Record.primary_smtp_address -split '@')[0] }
                if (-not $parameters.DisplayName) { $parameters.DisplayName = $entry.Row.Identity }
                $parameters[[string]$entry.Record.kind] = $true
                $created = @(New-Mailbox @parameters)
                if ($created.Count -ne 1) { throw 'Resource creation response is missing/ambiguous; never repeat creation.' }
                $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force
                Write-PraResourceIdentityPin -Entry $entry -Mailbox $created[0]
                $provisioningId = [string]$intent.cloudId
                Set-Pra2ResourceProvisioningDisabled -Identity $provisioningId
                $owned = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
                if (-not $owned) { throw 'Resource creation accepted but mailbox not visible yet; resume the same batch, do not create by another name.' }
            }
            if (-not $intent.cloudId -or -not $intent.cloudGuid) { $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $true -Force }
            Write-PraResourceIdentityPin -Entry $entry -Mailbox $owned.Mailbox
            if (-not $protected) { Assert-Pra2UnprotectedResource -Owned $owned }
            $identitySettings = @{ accountEnabled = $false }
            if ($protected) { $identitySettings.usageLocation = [string]$context.Config.Cloud.DefaultUsageLocation }
            $null = Invoke-Pra2Graph -Method PATCH -Uri ("v1.0/users/{0}" -f $owned.Id) -Body $identitySettings
            $disabled = Wait-Pra2Condition -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                $now = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
                if (-not $now) { throw 'Pinned resource disappeared while verifying disabled sign-in; nothing is adopted.' }
                return ($now.User.accountEnabled -eq $false)
            }
            if (-not $disabled.Ok) { throw 'Resource sign-in disable verification timed out; IDs retained, resume the same batch.' }
            $marker = @{ Identity = $owned.Id; ErrorAction = 'Stop' }; $marker[[string]$intent.markerAttribute] = [string]$intent.token
            Set-Mailbox @marker
            $intent | Add-Member -NotePropertyName markerPending -NotePropertyValue $false -Force
            $null = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
            Write-PraResourceJournal -Entry $entry -Step 'ResourceOwned' -Detail "owned cloud-only $($intent.cloudType), ID $($owned.Id), ExchangeGuid $($owned.Guid), sign-in disabled"
            if ($protected) {
            $skus = @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value)
            $member = Test-Pra2GroupMember -GroupId $intent.licensingGroupId -UserId $owned.Id
            $group = Get-Pra2ResourceGroup -Id $intent.licensingGroupId -Skus $skus -NewUnits $(if ($member) { 0 } else { 1 })
            if (-not $member) { Set-Pra2GroupMember -GroupId $group.Id -UserId $owned.Id -Operation Add }
            $licensed = Wait-Pra2Condition -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                $now = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id,assignedPlans,licenseAssignmentStates" -f $owned.Id)
                Test-Pra2ResourceLicence -User $now -Group $group
            }
            if (-not $licensed.Ok) {
                Write-PraResourceJournal -Entry $entry -Step 'ResourceLicencePending' -Detail 'Enterprise group plan not enabled yet; licence membership and data retained. Resume this batch.'
                Set-PraRowResult -Row $entry.Row -Status Pending -Detail 'Resource enterprise licensing is not enabled yet; data retained, resume this Convert batch.'
                continue
            }
            }
            $calendar = Set-Pra2ResourceConfiguration -Record $entry.Record -Owned $owned -Intent $intent
            $verified = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
            if ($verified.User.accountEnabled -ne $false) { throw 'Final resource disabled sign-in verification failed.' }
            if ($protected) {
                if (-not (Test-Pra2ResourceLicence -User $verified.User -Group $group)) { throw 'Final resource enterprise licensing verification failed.' }
            } else { Assert-Pra2UnprotectedResource -Owned $verified }
            $collisions = @(Get-Pra2ResourceCollisions -Record $entry.Record -Inventory (Get-Pra2ResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId)
            if ($collisions.Count) { throw ($collisions -join '; ') }
            Set-Pra2ResourceRoomListMembership -Record $entry.Record -Owned $verified -RoomListIntents $context.RoomListIntents
            Enable-Pra2ResourceBooking -Record $entry.Record -Owned $verified -Calendar $calendar
            $protectionText = if ($protected) { 'legacy enterprise licence retained throughout active lifetime' } else { 'no dedicated licence or hold; Recover has no inactive copy and does NOT preserve cloud reservations/messages' }
            Write-PraResourceJournal -Entry $entry -Step 'ResourceReady' -Status Done -Detail "Native resource configured; $protectionText."
            Set-PraRowResult -Row $entry.Row -Status Success -Field ExchangeOnline -Value $intent.cloudType -Detail "native $($intent.cloudType) $($owned.Guid), disabled sign-in, descriptive properties/room lists and booking restrictions/work hours restored, $protectionText; no general mailbox permissions or old on-premises content restored"
        } catch {
            $message = $_.Exception.Message
            if ($provisioningId) {
                try {
                    Set-Pra2ResourceProvisioningDisabled -Identity $provisioningId
                    $message += ' | Owned resource booking disabled and hidden; resume the same batch after resolving the failure.'
                } catch { $message += " | FAIL-CLOSED CLEANUP FAILED: $($_.Exception.Message); verify the pinned resource booking and visibility immediately." }
            }
            Stop-PraEntry -Entry $entry -Step 'ConvertResource' -Message $message
        }
    }
    $context.CurrentIdentity = ''
}

function Invoke-PraConvert {
    $config = $context.Config
    $dot = [char]0x00B7
    $context.Journal = Open-Pra2Store -Path $config.Store.JournalPath -Root $context.Root -Kind Journal
    $resume = $null; $snapshotId = $Snapshot; $resumeItems = @{}
    $read = @{ Verb = 'convert' }
    if ($Batch) {
        $resume = Get-Pra2Batch $context.Journal -Id $Batch
        if (-not $resume -or $resume.action -ne 'Convert') { throw "Convert batch $Batch not found in $($config.Store.JournalPath)." }
        $snapshotId = [long]$resume.snapshot_id
        $context.BatchId = [string]$resume.id
        foreach ($item in @(Get-Pra2BatchItem $context.Journal $resume.id)) { $resumeItems[[string]$item.object_guid] = $item }
        $read.ObjectGuid = [string[]]@($resumeItems.Values | Where-Object { $_.status -notin @('Done','Skipped') } | ForEach-Object { [string]$_.object_guid })
    }
    $data = Read-PraSnapshotData -Id $snapshotId @read
    $records = @($data.Records); $permissions = @($data.Permissions); $members = @($data.Members)
    $context.ResourceJournal = $resumeItems
    $context.RoomListIntents = @{}
    if ($resume) {
        foreach ($change in @(Get-Pra2TenantChange $context.Journal $resume.id | Where-Object kind -eq 'NativeRoomList')) {
            $intent = [string]$change.original_json | ConvertFrom-Json -ErrorAction Stop
            if ([string]$intent.lifecycle -ne 'CloudRoomList' -or [string]$intent.sourceGuid -ne [string]$change.target_id) { throw 'Invalid room-list creation journal.' }
            $context.RoomListIntents[[string]$change.target_id] = $intent
        }
    }
    $context.NativeResourceScope = @($records | Where-Object {
            (Test-Pra2NativeResource -Config $config -Record $_) -or (Get-Pra2ResourceIntent -Item $resumeItems[[string]$_.object_guid])
        }).Count -gt 0
    $nativeOnly = $records.Count -gt 0 -and @($records | Where-Object {
            $intent = Get-Pra2ResourceIntent -Item $resumeItems[[string]$_.object_guid]
            -not ((Test-Pra2NativeResource -Config $config -Record $_) -or $intent) -or (Get-Pra2ResourceProtectionMode $intent) -ne 'None'
        }).Count -eq 0
    if ($resume) { Write-PraItem -Context $context -Status Info -Icon Batch -Text ('Resuming batch {0}: {1} object(s) not finished' -f $resume.id, $records.Count) }

    Write-PraStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-Pra2Cloud -Context $context } catch { $context.CloudConnectFailed = $true; throw }
    Write-PraItem -Context $context -Status Ok -Text $how

    Write-PraStep -Context $context -Title 'Readiness' -Icon Key
    $context.CurrentOperation = 'Readiness'
    $skipUserGroup = @($records | Where-Object { $_.kind -eq 'User' }).Count -eq 0 -and @($records | Where-Object {
            (Test-Pra2NativeResource -Config $config -Record $_) -or (Get-Pra2ResourceIntent -Item $resumeItems[[string]$_.object_guid])
        }).Count -gt 0
    $tenant = Get-Pra2TenantFact -Context $context -SkipUserGroup:$skipUserGroup -NativeOnly:$nativeOnly
    $states = Read-PraCloudStates -Records $records
    $lookup = New-PraTrusteeLookup -States $states -Records $records
    $permissionIndex = Get-PraPermissionIndex -Permissions $permissions -Members $members
    # Read-only first: the state of every object, then the licence units really needed, then the decision.
    $assessed = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        if (Test-PraStop) { throw 'Stopped by the operator during the readiness check: nothing was changed.' }
        $context.CurrentIdentity = if ($record.user_principal_name) { [string]$record.user_principal_name } else { [string]$record.primary_smtp_address }
        $subset = Get-PraPermissionSubset -Index $permissionIndex -MailboxGuid ([string]$record.object_guid)
        $assessment = Get-PraObjectAssessment -Record $record -Tenant $tenant -Permissions $subset.Permissions -Members $subset.Members -Lookup $lookup -State $states[[string]$record.object_guid]
        $step = if ($resume -and $resumeItems.ContainsKey([string]$record.object_guid)) { [string](Get-PraValue $resumeItems[[string]$record.object_guid] 'step' '') } else { '' }
        [void]$assessed.Add([pscustomobject]@{ Record = $record; Assessment = $assessment; State = $assessment.State; Step = $step })
    }
    $context.CurrentIdentity = ''
    $need = Get-PraLicenceNeed -Entries @($assessed) -Tenant $tenant
    $tenantFindings = @(Test-Pra2TenantReadiness -Config $config -Roles @(Get-Pra2GraphRole) -Tenant $tenant -UserCount $need.Users -SharedCount $need.Shared -ResourceCount $need.Resources -SkipUserLicensing:$need.SkipUserLicensing -NativeOnly:$nativeOnly)
    $tenantFindings += @(Get-PraResourceLicenceFinding -Need $need -Tenant $tenant)
    foreach ($finding in $tenantFindings | Where-Object { $_.Level -eq 'Error' }) { Write-PraLog -Context $context -Message $finding.Message -Level Error }
    if (@($tenantFindings | Where-Object Level -eq 'Error').Count) { throw 'Tenant prerequisites not met (see above): nothing was changed.' }
    Write-PraItem -Context $context -Status Ok -Text ('Tenant prerequisites met (permissions, licences: {0} new unit(s) for users{1})' -f $need.Users, $(if ($need.Shared) { ', up to {0} at a time for the temporary licence of the shared mailboxes' -f [Math]::Min([Math]::Max(1, [int]$config.Licensing.Shared.Parallel), $need.Shared) } else { '' }))
    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $assessed) {
        $record = $entry.Record; $assessment = $entry.Assessment
        $row = New-PraRow $record
        [void]$context.Rows.Add($row)
        $context.CurrentIdentity = $row.Identity
        $blocking = @($assessment.Errors) + @($assessment.Findings | Where-Object { $_.Code -eq 'EXCHANGE_PLAN_PRESENT' })
        if ($resume) { $blocking = @($blocking | Where-Object { $_.Code -notin @('ALREADY_CLOUD_MAILBOX','EXCHANGE_PLAN_PRESENT') }) }
        $row.Warnings = (@($assessment.Warnings | ForEach-Object { $_.Message })) -join ' | '
        if ($blocking.Count) { Set-PraRowResult -Row $row -Status Error -Detail ('Not converted: ' + (($blocking | ForEach-Object { $_.Message }) -join ' | ')); continue }
        [void]$plan.Add([pscustomobject]@{ Record = $record; Row = $row; EntraId = $(if ($assessment.State.User) { [string]$assessment.State.User.id } else { '' }); State = $assessment.State
                Intent = $(if ($resumeItems.ContainsKey([string]$record.object_guid)) { Get-Pra2ResourceIntent -Item $resumeItems[[string]$record.object_guid] } else { $null }) })
    }
    $context.CurrentIdentity = ''
    $planUsers = @($plan | Where-Object { $_.Record.kind -eq 'User' })
    $planResources = @($plan | Where-Object { Get-PraValue $_.State 'NativeResource' $false })
    $planShared = @($plan | Where-Object { $_.Record.kind -ne 'User' -and $_ -notin $planResources })
    $roomLists = @(Get-PraRoomListPlan -Entries $planResources -Records $records)
    Write-PraItem -Context $context -Status Info -Icon Target -Text ('{0} object(s) ready: {1} user(s), {2} shared, {5} native resources {3} {4} blocked' -f $plan.Count, $planUsers.Count, $planShared.Count, $dot, ($records.Count - $plan.Count), $planResources.Count)

    Write-PraStep -Context $context -Title 'Plan and confirmation' -Icon Plan
    $modeText = @{ Group = "licence group $($config.Licensing.Users.GroupId)"; Kiosk = 'Exchange Kiosk plan of the user licence'; Direct = "$($config.Licensing.Users.SkuPartNumber) (Exchange only)" }[$config.Licensing.Users.Mode]
    Write-PraItem -Context $context -Status Info -Text ("Users: ExchangeGuid cleared, source of authority to the cloud, usageLocation, Exchange plan by $modeText; the Teams storage becomes the mailbox")
    if ($planShared.Count) { Write-PraItem -Context $context -Status Info -Text ("Shared mailboxes, in waves of up to $($config.Licensing.Shared.Parallel) (one temporary unit each): temporary $($config.Licensing.Shared.SkuPartNumber), shared, $($config.Retention.TagAttribute) = $($config.Retention.TagValue), licence given back, permissions from the snapshot") }
    foreach ($entry in $planResources) {
        $protectionText = if ((Get-Pra2ResourceProtectionMode $entry.Intent) -eq 'CaseHold') { 'legacy enterprise licence/case-hold lifecycle retained' } else { 'no dedicated licence or hold (standard resources up to 50 GB); Recover has no inactive copy, cloud reservations/messages NOT preserved' }
        Write-PraItem -Context $context -Status Info -Text ("Native $($entry.Record.kind): New-Mailbox, disabled cloud-only identity, independent GUID, $protectionText; addresses: " + ((Get-Pra2ResourceAddresses $entry.Record) -join ', ') + '; booking/work hours restored, no on-premises content restored.')
    }
    foreach ($list in $roomLists) { Write-PraItem -Context $context -Status Info -Text "RoomList $($list.PrimarySmtpAddress): create if missing, or reuse verified cloud-only list; before its rooms, then add only selected room members." }
    $convertEstimate = Get-PraDurationEstimate -UserWaves ([Math]::Ceiling($planUsers.Count / [Math]::Max(1, [int](Get-PraValue $context 'UserWaveSize' 500)))) -SharedRounds ([Math]::Ceiling($planShared.Count / [Math]::Max(1, [int]$config.Licensing.Shared.Parallel)))
    if ($convertEstimate.Text) { Write-PraItem -Context $context -Status Info -Icon Clock -Text ('Rough estimate: about {0} (typical tenant; a slow Microsoft 365 evening can take much longer)' -f $convertEstimate.Text) }
    foreach ($entry in $plan) { $entry.Row.FinalStatus = 'Planned'; Write-PraLog -Context $context -Message ('Planned: {0} ({1})' -f $entry.Row.Identity, $entry.Record.kind) -Level Sub }
    if ($context.Mode -ne 'Apply') {
        Skip-PraStep -Title 'Users: source of authority and Exchange plan' -Reason 'Preview: not run.' -Icon Write
        Skip-PraStep -Title 'Users: cloud mailboxes' -Reason 'Preview: not run.' -Icon Clock
        Skip-PraStep -Title 'Shared mailboxes' -Reason 'Preview: not run.' -Icon Mail
        $context.NextSteps = @('pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply' + $(if ($Batch) { " -Batch $Batch" } else { Get-PraIdentityArgument }))
        return
    }
    if (-not $plan.Count) { throw 'No object can be converted (see the reasons above).' }
    Confirm-PraApply ('Convert {0} object(s) to Exchange Online ({1} user(s), {2} shared, {3} native resources)?' -f $plan.Count, $planUsers.Count, $planShared.Count, $planResources.Count)
    if (-not $resume) {
        $context.BatchId = New-Pra2Batch $context.Journal Convert @{ SnapshotId = [long]$data.Snapshot.id; Environment = $config.Environment; Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            ToolVersion = $toolVersion; ScopeJson = (@{ Identity = $Identity; IdentityPath = $IdentityPath; Identities = @(Get-PraIdentityList).Count; Scope = $Scope; UsersMode = $config.Licensing.Users.Mode } | ConvertTo-Json -Compress) }
        foreach ($entry in $plan) {
            if ($entry -in $planResources) {
                $token = [guid]::NewGuid().ToString('N')
                $attribute = if ($config.Retention.TagAttribute -eq 'CustomAttribute15') { 'CustomAttribute14' } else { 'CustomAttribute15' }
                $entry.Intent = [pscustomobject]@{ lifecycle = 'CloudResource'; protectionMode = 'None'; token = $token; markerAttribute = $attribute
                    creationName = ('PRA-R-{0}-{1}-{2}' -f $context.BatchId.Substring(0,8), ([string]$entry.Record.object_guid).Substring(0,8), $token)
                    creationAttempted = $false; cloudId = ''; cloudGuid = ''; cloudType = "$($entry.Record.kind)Mailbox" }
                Set-Pra2BatchItem $context.Journal $context.BatchId ([string]$entry.Record.object_guid) @{ kind = [string]$entry.Record.kind; identity = $entry.Row.Identity
                    status = 'Planned'; step = 'ResourceIntent'; original_json = ($entry.Intent | ConvertTo-Json -Depth 15 -Compress) }
                continue
            }
            $original = @{ entraId = $entry.EntraId; isCloudManaged = $entry.State.IsCloudManaged; exchangeGuid = [string]$entry.Record.exchange_guid
                usageLocation = [string]$entry.State.User.usageLocation; recipientType = $entry.State.RecipientType; tag = $entry.State.TagValue; holds = @($entry.State.Holds)
                licences = @(Get-Pra2DirectLicence $entry.State.User) }
            Set-Pra2BatchItem $context.Journal $context.BatchId ([string]$entry.Record.object_guid) @{ kind = [string]$entry.Record.kind; identity = $entry.Row.Identity
                entra_id = $entry.EntraId; status = 'Planned'; original_json = ($original | ConvertTo-Json -Compress -Depth 6) }
        }
    }
    $originals = @{}; foreach ($item in @(Get-Pra2BatchItem $context.Journal $context.BatchId)) { $originals[[string]$item.object_guid] = @(Get-PraOriginalLicence $item) }
    Write-PraItem -Context $context -Status Ok -Icon Batch -Text ("Batch {0} in {1}" -f $context.BatchId, $config.Store.JournalPath)

    Write-PraStep -Context $context -Title 'Users: source of authority and Exchange plan' -Icon Write
    $started = [System.Collections.Generic.List[object]]::new()
    if ($planUsers.Count -and $config.Licensing.Users.Mode -eq 'Group' -and $tenant.LicenceGroup -and $tenant.LicenceGroup.onPremisesSyncEnabled -and -not $tenant.LicenceGroupCloudManaged) {
        $context.CurrentOperation = 'Group-SOA'
        Set-Pra2SourceOfAuthority -Id $config.Licensing.Users.GroupId -Cloud $true -Type groups
        Set-Pra2TenantChange $context.Journal $context.BatchId 'GroupSoa' $config.Licensing.Users.GroupId -OriginalJson '{"isCloudManaged":false}'
        Write-PraJournal -Step 'GroupSoaCloud' -Detail "licence group $($tenant.LicenceGroup.displayName) managed in the cloud"
        Write-PraItem -Context $context -Status Ok -Text "Licence group $($tenant.LicenceGroup.displayName) now managed in the cloud (given back to AD by Recover)"
    }
    $stopped = $false
    # Waves of 500 users, phase by phase (a stop is taken between two waves); their mailboxes are awaited together below.
    $waveSize = [Math]::Max(1, [int](Get-PraValue $context 'UserWaveSize' 500))
    $userClock = [Diagnostics.Stopwatch]::StartNew()
    for ($start = 0; $start -lt $planUsers.Count; $start += $waveSize) {
        $wave = @($planUsers[$start..([Math]::Min($planUsers.Count, $start + $waveSize) - 1)])
        if ($stopped -or (Test-PraStop)) { $stopped = $true; foreach ($entry in $wave) { Set-PraRowStopped $entry.Row }; continue }
        $context.CurrentOperation = 'Convert-User'
        if ($planUsers.Count -gt $waveSize) { Write-PraProgress -Context $context -Phase 'Users' -Done $start -Total $planUsers.Count -ElapsedSeconds $userClock.Elapsed.TotalSeconds }
        foreach ($entry in @(Invoke-PraUserPlanSet -Entries $wave -Tenant $tenant)) { [void]$started.Add($entry) }
    }
    if (-not $planUsers.Count) { Write-PraItem -Context $context -Status Skip -Text 'No user in this batch.' }

    Write-PraStep -Context $context -Title 'Users: cloud mailboxes' -Icon Clock
    $context.CurrentIdentity = ''; $context.CurrentOperation = 'Wait-Mailbox'
    # Every user waiting is polled together (filters of 50 objects): thousands of users per round.
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $waiting = @(Wait-PraExoStateSet -Entries @($started) -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -StopOnRequest -Test { param($state) $state.Type -eq 'UserMailbox' } -OnReady {
            param($entry, $state, $seconds)
            Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'MailboxReady' -Detail ("UserMailbox {0} after {1}" -f $state.ExchangeGuid, (Format-PraDuration $seconds)) -Item @{ step = 'MailboxReady'; status = 'Done'; result_json = (@{ cloudGuid = $state.ExchangeGuid } | ConvertTo-Json -Compress) }
            Set-PraRowResult -Row $entry.Row -Status Success -Detail ('cloud mailbox {0} (Teams storage promoted when present)' -f $state.ExchangeGuid) -Field 'ExchangeOnline' -Value 'UserMailbox'
        })
    foreach ($entry in $waiting) { Set-PraRowResult -Row $entry.Row -Status Pending -Detail ('no cloud mailbox yet after {0}: run Convert -Batch {1} again later' -f (Format-PraDuration $clock.Elapsed.TotalSeconds), $context.BatchId) }
    if (-not $started.Count) { Write-PraItem -Context $context -Status Skip -Text 'Nothing to wait for.' }

    Write-PraStep -Context $context -Title 'Shared mailboxes' -Icon Mail
    $sharedSku = $tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
    # Waves of Licensing.Shared.Parallel shared mailboxes, one temporary unit each, never more than the free units: a shared
    # mailbox that failed keeps its temporary licence, so no object is started without a unit (it would stop half-way,
    # GUID cleared, source of authority in the cloud). An object already a SharedMailbox, or already holding the unit
    # (resume), needs none. A stop is taken between two waves.
    $parallel = [Math]::Max(1, [int]$config.Licensing.Shared.Parallel)
    $next = 0
    $sharedClock = [Diagnostics.Stopwatch]::StartNew()
    while ($next -lt $planShared.Count) {
        if ($stopped -or (Test-PraStop)) { $stopped = $true; for (; $next -lt $planShared.Count; $next++) { Set-PraRowStopped $planShared[$next].Row }; break }
        $context.CurrentOperation = 'Convert-Shared'
        $candidates = @($planShared[$next..([Math]::Min($planShared.Count, $next + $parallel) - 1)])
        $exo = Get-Pra2ExoStateSet -Id ([string[]]@($candidates | ForEach-Object { $_.EntraId })) -TagAttribute $config.Retention.TagAttribute
        $free = 0
        if ($sharedSku) {
            $now = @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value) | Where-Object { [string]$_.skuId -eq [string]$sharedSku.skuId } | Select-Object -First 1
            if ($now) { $free = Get-Pra2FreeUnit $now }
        }
        $wave = [System.Collections.Generic.List[object]]::new()
        foreach ($entry in $candidates) {
            $isShared = $exo[$entry.EntraId] -and $exo[$entry.EntraId].Type -eq 'SharedMailbox'
            $holds = $sharedSku -and @(Get-Pra2DirectLicence $entry.State.User | Where-Object { $_.skuId -eq [string]$sharedSku.skuId }).Count
            if (-not $isShared -and -not $holds) { if ($free -lt 1) { break }; $free-- }
            $wave.Add($entry)
        }
        if (-not $wave.Count) {
            for (; $next -lt $planShared.Count; $next++) {
                $entry = $planShared[$next]
                $message = "not started: no free unit of $($sharedSku.skuPartNumber) (a shared mailbox that failed keeps its temporary licence until it is finished); nothing was changed on this object. Free a unit, then run Convert -Batch $($context.BatchId) again."
                Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'NoFreeUnit' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
                Set-PraRowResult -Row $entry.Row -Status Error -Detail $message
            }
            break
        }
        Write-PraProgress -Context $context -Phase 'Shared mailboxes' -Done $next -Total $planShared.Count -ElapsedSeconds $sharedClock.Elapsed.TotalSeconds
        $next += $wave.Count
        Invoke-PraConvertSharedWave -Entries $wave.ToArray() -Tenant $tenant -PermissionIndex $permissionIndex -Originals $originals
    }
    if (-not $planShared.Count) { Write-PraItem -Context $context -Status Skip -Text 'No shared mailbox in this batch.' }
    if (-not $stopped -and -not (Test-PraStop) -and $roomLists.Count) {
        $context.CurrentOperation = 'Convert-RoomLists'
        Invoke-PraConvertRoomLists -Lists $roomLists -Unprotected:(@($planResources | Where-Object { (Get-Pra2ResourceProtectionMode $_.Intent) -eq 'None' }).Count -gt 0)
    }
    $resourceParallel = [Math]::Max(1, [int](Get-PraValue (Get-PraValue $config 'Resources' @{}) 'Parallel' 100))
    for ($start = 0; $start -lt $planResources.Count; $start += $resourceParallel) {
        $wave = @($planResources[$start..([Math]::Min($planResources.Count, $start + $resourceParallel) - 1)])
        if ($stopped -or (Test-PraStop)) { $stopped = $true; foreach ($entry in $wave) { Set-PraRowStopped $entry.Row }; continue }
        Invoke-PraConvertResourceWave -Entries $wave -NativeEntries $planResources -NativeRecords @($records | Where-Object { Get-PraValue $states[[string]$_.object_guid] 'NativeResource' $false })
    }
    $context.CurrentIdentity = ''
    $items = @(Get-Pra2BatchItem $context.Journal $context.BatchId)
    $counts = @{ done = @($items | Where-Object status -eq 'Done').Count; failed = @($items | Where-Object status -eq 'Failed').Count; running = @($items | Where-Object status -in @('Running','Planned')).Count }
    $pendingLists = @($context.RoomListIntents.Values | Where-Object status -ne 'Done').Count
    $counts.roomListsDone = @($context.RoomListIntents.Values | Where-Object status -eq 'Done').Count
    $counts.roomListsPending = $pendingLists
    Set-Pra2Batch $context.Journal $context.BatchId $(if ($counts.done -eq $items.Count -and -not $pendingLists) { 'Complete' } else { 'Partial' }) $counts
    $context.SnapshotLabel = $context.SnapshotLabel
    $context.NextSteps = @(
        $(if ($counts.done -ne $items.Count -or $pendingLists) { "pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Batch $($context.BatchId)   # resume the objects/lists not finished" }),
        "When AD, Exchange and Entra Connect are rebuilt: pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch $($context.BatchId)") | Where-Object { $_ }
}

function Stop-PraRecoverEntry {
    <# An object of Recover that fails in a phase: journal, report row, and it leaves the next phases. #>
    param([Parameter(Mandatory)][object]$Entry, [Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Message)
    $text = $Message -replace '\s+', ' '
    Write-PraJournal -ObjectGuid ([string]$Entry.Item.object_guid) -Step $Step -Outcome Fail -Detail $text -Item @{ status = 'Failed'; message = $text }
    Set-PraRowResult -Row $Entry.Row -Status Error -Detail "rollback stopped: $text"
}

function Add-PraRecoverHold {
    <#
    Case hold for many objects: added where missing (groups of 100 locations per call, new policies of the tool when the
    ones in use are full), then ONE wait for all until a case hold of the tool is stamped (HoldTimeoutMinutes).
    Returns Stamped, Late and Failed (entries) and Seconds (object ID -> seconds).
    #>
    param([AllowEmptyCollection()][object[]]$Entries = @(), [Parameter(Mandatory)][object]$HoldSet)
    $result = [pscustomobject]@{ Stamped = @(); Late = @(); Failed = @(); Seconds = @{} }
    if (-not $Entries.Count) { return $result }
    $ids = [string[]]@($Entries | ForEach-Object { $_.EntraId })
    $now = Get-Pra2ExoStateSet -Id $ids -TagAttribute $context.Config.Retention.TagAttribute
    $missing = [string[]]@($Entries | Where-Object { -not ($now[$_.EntraId] -and (Test-Pra2ToolHold -Set $HoldSet -Holds @($now[$_.EntraId].Holds))) } | ForEach-Object { $_.EntraId })
    $added = Add-Pra2HoldLocationSet -Set $HoldSet -Id $missing
    $failed = @($Entries | Where-Object { $added.Failed.ContainsKey($_.EntraId) })
    foreach ($entry in $failed) { Stop-PraRecoverEntry -Entry $entry -Step 'Hold' -Message $added.Failed[$entry.EntraId] }
    $active = @($Entries | Where-Object { -not $added.Failed.ContainsKey($_.EntraId) })
    $seconds = $result.Seconds
    $late = @(Wait-PraExoStateSet -Entries $active -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -Test { param($state) Test-Pra2ToolHold -Set $HoldSet -Holds @($state.Holds) } -OnReady {
            $seconds[$args[0].EntraId] = $args[2] })
    $result.Late = $late
    $result.Stamped = @($active | Where-Object { $_ -notin $late })
    $result.Failed = $failed
    return $result
}

function Get-PraHoldTagOf {
    <# The tag (UniH + GUID) of the policy of the tool that holds an object, or ''. #>
    param([Parameter(Mandatory)][object]$HoldSet, [Parameter(Mandatory)][string]$Id)
    $policy = $HoldSet.Policies | Where-Object { $_.Locations.Contains($Id) } | Select-Object -First 1
    if ($policy) { return $policy.Tag }
    return ''
}

function Invoke-PraRecoverUserWave {
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
    $lift = @($users | Where-Object { $_.Exo.Type -eq 'UserMailbox' -and (Test-Pra2ToolHold -Set $HoldSet -Holds @($_.Exo.Holds)) })
    if ($lift.Count) {
        $removed = Remove-Pra2HoldLocationSet -Set $HoldSet -Id (& $ids $lift)
        foreach ($entry in $lift) {
            if ($removed.Failed.ContainsKey($entry.EntraId)) { Stop-PraRecoverEntry -Entry $entry -Step 'HoldLifted' -Message $removed.Failed[$entry.EntraId]; [void]$users.Remove($entry) }
            else { Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'HoldLifted' -Detail 'case hold lifted for the rollback' }
        }
    }
    # 2. Source of authority back to AD.
    $behaviors = Get-Pra2SyncBehaviorSet -UserId (& $ids $users)
    $toAd = @($users | Where-Object { $behaviors[$_.EntraId] -eq $true })
    $answers = Invoke-Pra2GraphBatch -Requests @($toAd | ForEach-Object { @{ Method = 'PATCH'; Url = ('/users/{0}/onPremisesSyncBehavior' -f $_.EntraId); Body = @{ isCloudManaged = $false } } })
    for ($i = 0; $i -lt $toAd.Count; $i++) {
        if ($answers[$i].Ok) { Write-PraJournal -ObjectGuid ([string]$toAd[$i].Item.object_guid) -Step 'SoaAd' -Detail 'isCloudManaged = false' -Item @{ step = 'SoaAd' } }
        else { Stop-PraRecoverEntry -Entry $toAd[$i] -Step 'SoaAd' -Message ('source of authority not given back to AD ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$users.Remove($toAd[$i]) }
    }
    # 3. One delta cycle (none when every user is synchronised already), one wait for all.
    if ($users.Count) {
        $context.CurrentOperation = 'EntraConnect-Delta'
        $now = Get-Pra2ExoStateSet -Id (& $ids $users)
        if (@($users | Where-Object { $s = $now[$_.EntraId]; -not $s -or $s.IsDirSynced -ne $true }).Count) {
            $state = Invoke-Pra2EntraConnect -Context $context -Operation Delta -Caller $caller
            Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect delta cycle: $state"
        } else { Write-PraItem -Context $context -Status Info -Icon Sync -Text 'Already synchronised from AD: no delta cycle needed.' }
        $late = @(Wait-PraExoStateSet -Entries $users.ToArray() -TimeoutSeconds (Get-PraWaitSeconds 'SyncTimeoutMinutes') -Test { param($state) $state.IsDirSynced -eq $true } -OnReady {
                Write-PraItem -Context $context -Status Ok -Text ('{0}: synchronised from AD again' -f $args[0].Row.Identity) })
        foreach ($entry in $late) { Set-PraRowResult -Row $entry.Row -Status Pending -Detail 'not synchronised from AD yet: run Entra Connect, then Recover -Batch again'; [void]$users.Remove($entry) }
    }
    $context.CurrentOperation = 'Users-Plan'
    # 4. A lifted case hold must be gone from the cloud mailbox before the plan is removed.
    $lifted = @($users | Where-Object { $_ -in $lift })
    if ($lifted.Count) {
        $late = @(Wait-PraExoStateSet -Entries $lifted -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -Test { param($state) -not (Test-Pra2ToolHold -Set $HoldSet -Holds @($state.Holds)) })
        foreach ($entry in $late) { Stop-PraRecoverEntry -Entry $entry -Step 'RecoverUser' -Message 'the case hold is still stamped on the cloud mailbox (it would block the rollback)'; [void]$users.Remove($entry) }
    }
    # 5. Exchange plan removed: licence group by Graph batches, Kiosk / direct / temporary licence one at a time.
    $removedText = @{}
    $tenantSkus = $null
    $groupUsers = @($users | Where-Object { $_.Item.kind -eq 'User' -and $config.Licensing.Users.Mode -eq 'Group' })
    $answers = Invoke-Pra2GraphBatch -Requests @($groupUsers | ForEach-Object { @{ Method = 'DELETE'; Url = ('/groups/{0}/members/{1}/$ref' -f $config.Licensing.Users.GroupId, $_.EntraId) } })
    for ($i = 0; $i -lt $groupUsers.Count; $i++) {
        if ($answers[$i].Ok -or $answers[$i].Status -eq 404) { $removedText[$groupUsers[$i].EntraId] = 'removed from the licence group' }
        else { Stop-PraRecoverEntry -Entry $groupUsers[$i] -Step 'PlanRemoved' -Message ('not removed from the licence group ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error); [void]$users.Remove($groupUsers[$i]) }
    }
    foreach ($entry in @($users | Where-Object { $_ -notin $groupUsers })) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            if (-not $tenantSkus) { $tenantSkus = @((Get-Pra2TenantFact -Context $context).Skus) }
            $removedText[$entry.EntraId] = if ($entry.Item.kind -ne 'User') {
                # Shared mailbox not fully converted: its temporary licence goes back to what it was before Convert.
                $sku = $tenantSkus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
                if ($sku) { 'temporary licence ' + (Restore-Pra2Licence -UserId $entry.EntraId -SkuId $sku.skuId -Original $entry.Original).ToLowerInvariant() } else { 'no temporary licence' }
            } elseif ($config.Licensing.Users.Mode -eq 'Kiosk') {
                $kiosk = Get-Pra2KioskAssignment -Licence (Get-Pra2UserLicence -UserId $entry.EntraId) -Skus $tenantSkus
                if ($kiosk) { "$($kiosk.SkuName) " + (Restore-Pra2Licence -UserId $entry.EntraId -SkuId $kiosk.SkuId -Original $entry.Original).ToLowerInvariant() } else { 'no Kiosk licence' }
            } else {
                $sku = $tenantSkus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Users.SkuPartNumber } | Select-Object -First 1
                if ($sku) { "$($sku.skuPartNumber) " + (Restore-Pra2Licence -UserId $entry.EntraId -SkuId $sku.skuId -Original $entry.Original).ToLowerInvariant() } else { 'no direct licence' }
            }
        } catch { Stop-PraRecoverEntry -Entry $entry -Step 'PlanRemoved' -Message $_.Exception.Message; [void]$users.Remove($entry) }
    }
    $context.CurrentIdentity = ''
    foreach ($entry in $users) { Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'PlanRemoved' -Detail "Exchange plan removed: $($removedText[$entry.EntraId])" -Item @{ step = 'PlanRemoved' } }
    # 6. One wait for all: MailUser with the on-premises GUID.
    $late = @(Wait-PraExoStateSet -Entries $users.ToArray() -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -Test { param($state, $entry) $state.Type -eq 'MailUser' -and $state.ExchangeGuid -eq [string]$entry.Record.exchange_guid } -OnReady {
            Write-PraJournal -ObjectGuid ([string]$args[0].Item.object_guid) -Step 'MailUser' -Detail "MailUser with the on-premises GUID after $($args[2]) s" -Item @{ step = 'MailUser' } })
    foreach ($entry in $late) { Stop-PraRecoverEntry -Entry $entry -Step 'RecoverUser' -Message ('not a MailUser with the on-premises ExchangeGuid {0} yet' -f $entry.Record.exchange_guid); [void]$users.Remove($entry) }
    # 7. Case hold, by object ID (the policy keeps it right when the UPN and the SMTP address differ), one wait.
    $ready = @($users) + @($back)
    $hold = Add-PraRecoverHold -Entries $ready -HoldSet $HoldSet
    foreach ($entry in $hold.Stamped) {
        $guid = [string]$entry.Item.object_guid
        $entry.Row.ExchangeOnline = 'MailUser'
        $text = if ($entry.Item.kind -eq 'User') { 'back on-premises (MailUser, on-premises GUID); the cloud mailbox stays as ComponentShared under the case hold' }
        else { 'shared mailbox not fully converted, rolled back like a user (MailUser, on-premises GUID); its cloud data, if any, stays under the case hold' }
        Write-PraJournal -ObjectGuid $guid -Step 'Hold' -Detail "case hold stamped after $($hold.Seconds[$entry.EntraId]) s" -Item @{ step = 'Hold'; status = 'Done' }
        Set-PraRowResult -Row $entry.Row -Status Success -Detail $text -Field 'Holds' -Value (Get-PraHoldTagOf -HoldSet $HoldSet -Id $entry.EntraId)
    }
    foreach ($entry in $hold.Late) {
        $entry.Row.ExchangeOnline = 'MailUser'
        Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Hold' -Outcome Info -Detail 'case hold added, not stamped yet' -Item @{ step = 'HoldPending' }
        Set-PraRowResult -Row $entry.Row -Status Pending -Detail "back on-premises; case hold not stamped yet: run Recover -Batch $Batch again to check it (the hold is never removed from an object that is back)"
    }
}

function Invoke-PraRecoverSharedWave {
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
        $hold = Add-PraRecoverHold -Entries $fresh -HoldSet $HoldSet
        foreach ($entry in $hold.Late) { Stop-PraRecoverEntry -Entry $entry -Step 'RecoverShared' -Message 'the case hold is not stamped on the shared mailbox: identity NOT deleted (it would not become inactive)' }
        foreach ($entry in $hold.Stamped) { Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Hold' -Detail "case hold stamped after $($hold.Seconds[$entry.EntraId]) s" -Item @{ step = 'Hold' } }
        $fresh = @($hold.Stamped)
    }
    foreach ($entry in @($Entries | Where-Object { $_.Path -eq 'SharedDeleted' })) {
        Write-PraItem -Context $context -Status Info -Text ('{0}: identity deleted by an earlier run (mailbox {1} inactive): recycle bin and recreated object checked' -f $entry.Row.Identity, $entry.CloudGuid)
        $toPurge.Add($entry)
    }
    $paused = $false
    # Dot-sourced: it sets $paused of this function.
    $pause = {
        if (-not $paused) {
            $paused = $true
            $pauseState = Invoke-Pra2EntraConnect -Context $context -Operation Pause -Caller $caller
            Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler paused: $pauseState"
        }
    }
    try {
        if ($fresh.Count) {
            . $pause
            foreach ($entry in $fresh) { $entry.CloudGuid = $entry.Exo.ExchangeGuid }
            $answers = Invoke-Pra2GraphBatch -Requests @($fresh | ForEach-Object { @{ Method = 'DELETE'; Url = ('/users/{0}' -f $_.EntraId) } })
            $deleted = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $fresh.Count; $i++) {
                if ($answers[$i].Ok -or $answers[$i].Status -eq 404) { $deleted.Add($fresh[$i]) }
                else { Stop-PraRecoverEntry -Entry $fresh[$i] -Step 'RecoverShared' -Message ('identity not deleted ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error) }
            }
            # Each mailbox inactive (no filter on ExchangeGuid in Exchange Online: one call per mailbox and round).
            $clock = [Diagnostics.Stopwatch]::StartNew()
            $null = Wait-Pra2Condition -TimeoutSeconds 900 -IntervalSeconds 15 -Test {
                foreach ($entry in @($deleted)) {
                    if (Get-Mailbox -InactiveMailboxOnly -Identity $entry.CloudGuid -ErrorAction SilentlyContinue) {
                        Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Inactive' -Detail ("identity deleted, mailbox {0} inactive after {1} s" -f $entry.CloudGuid, [int]$clock.Elapsed.TotalSeconds) -Item @{ step = 'Inactive'; result_json = (@{ inactiveGuid = $entry.CloudGuid } | ConvertTo-Json -Compress) }
                        [void]$deleted.Remove($entry); $toPurge.Add($entry)
                    }
                }
                -not $deleted.Count
            }
            foreach ($entry in $deleted) { Stop-PraRecoverEntry -Entry $entry -Step 'RecoverShared' -Message ("identity deleted but mailbox {0} not inactive after 15 min: check it before anything else (Entra ID recycle bin, 30 days)" -f $entry.CloudGuid) }
        }
        if ($toPurge.Count) {
            $binUrl = { param($entry) '/directory/deletedItems/{0}' -f $entry.EntraId }
            $answers = Invoke-Pra2GraphBatch -Requests @($toPurge | ForEach-Object { @{ Method = 'GET'; Url = ((& $binUrl $_) + '?$select=id') } })
            $binned = [System.Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $toPurge.Count; $i++) { if ($answers[$i].Ok) { $binned.Add($toPurge[$i]) } }
            if ($binned.Count) {
                . $pause
                $answers = Invoke-Pra2GraphBatch -Requests @($binned | ForEach-Object { @{ Method = 'DELETE'; Url = (& $binUrl $_) } })
                for ($i = $binned.Count - 1; $i -ge 0; $i--) {
                    if (-not $answers[$i].Ok -and $answers[$i].Status -ne 404) {
                        Stop-PraRecoverEntry -Entry $binned[$i] -Step 'RecoverShared' -Message ('identity not deleted permanently ({0}): {1}' -f $answers[$i].Status, $answers[$i].Error)
                        [void]$toPurge.Remove($binned[$i]); $binned.RemoveAt($i)
                    }
                }
                $null = Wait-Pra2Condition -TimeoutSeconds 300 -IntervalSeconds 10 -Test {
                    $answers = Invoke-Pra2GraphBatch -Requests @($binned | ForEach-Object { @{ Method = 'GET'; Url = ((& $binUrl $_) + '?$select=id') } })
                    for ($i = $binned.Count - 1; $i -ge 0; $i--) { if (-not $answers[$i].Ok) { $binned.RemoveAt($i) } }
                    -not $binned.Count
                }
                foreach ($entry in $binned) { Stop-PraRecoverEntry -Entry $entry -Step 'RecoverShared' -Message 'identity still in the Entra ID recycle bin: the next sync would restore it'; [void]$toPurge.Remove($entry) }
            }
            foreach ($entry in $toPurge) {
                Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Purged' -Detail 'identity deleted permanently' -Item @{ step = 'Purged' }
                Write-PraItem -Context $context -Status Ok -Text ('{0}: mailbox inactive under the case hold, identity deleted permanently' -f $entry.Row.Identity)
            }
        }
    } finally {
        if ($paused) {
            try { $state = Invoke-Pra2EntraConnect -Context $context -Operation Resume -Caller $caller; Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler resumed: $state" }
            catch { Write-PraLog -Context $context -Message "Entra Connect scheduler NOT resumed: $($_.Exception.Message). Resume it by hand (Set-ADSyncScheduler -SyncCycleEnabled `$true)." -Level Error }
        }
    }
    if (-not $toPurge.Count) { return }
    # One delta cycle, then the new objects of the wave together (Graph batches by immutableId, Exchange Online by filters).
    $context.CurrentOperation = 'EntraConnect-Delta'
    $state = Invoke-Pra2EntraConnect -Context $context -Operation Delta -Caller $caller
    Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect delta cycle: $state"
    $waiting = [System.Collections.Generic.List[object]]::new()
    foreach ($entry in $toPurge) { $waiting.Add($entry) }
    $null = Wait-Pra2Condition -TimeoutSeconds (Get-PraWaitSeconds 'SyncTimeoutMinutes') -IntervalSeconds $interval -Test {
        $answers = Invoke-Pra2GraphBatch -Requests @($waiting | ForEach-Object {
                @{ Method = 'GET'; Url = ('/users?$filter={0}&$select=id' -f [uri]::EscapeDataString(("onPremisesImmutableId eq '{0}'" -f ([string]$_.Record.immutable_id).Replace("'", "''")))) } })
        $found = @{}
        for ($i = 0; $i -lt $waiting.Count; $i++) {
            $values = @(if ($answers[$i].Ok) { @(Get-PraValue $answers[$i].Body 'value' @()) | ForEach-Object { $_ } })
            if ($values.Count -eq 1 -and [string]$values[0].id -ne $waiting[$i].EntraId) { $found[$waiting[$i].EntraId] = [string]$values[0].id }
        }
        $states = Get-Pra2ExoStateSet -Id ([string[]]@($found.Values))
        foreach ($entry in @($waiting)) {
            if (-not $found.ContainsKey($entry.EntraId)) { continue }
            $newId = $found[$entry.EntraId]; $s = $states[$newId]
            if ($s -and $s.Type -eq 'MailUser' -and $s.ExchangeGuid -eq [string]$entry.Record.exchange_guid) {
                Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'Recreated' -Detail ("new object {0} from AD, MailUser with the on-premises GUID" -f $newId) -Item @{ step = 'Recreated'; status = 'Done'; entra_id = $newId }
                Set-PraRowResult -Row $entry.Row -Status Success -Detail ('recreated from AD ({0}); the cloud mailbox {1} is inactive under the case hold' -f $newId, $entry.CloudGuid) -Field 'ExchangeOnline' -Value 'MailUser'
                [void]$waiting.Remove($entry)
            }
        }
        -not $waiting.Count
    }
    foreach ($entry in $waiting) { Set-PraRowResult -Row $entry.Row -Status Pending -Detail "identity deleted permanently, mailbox inactive; not recreated by Entra Connect yet: run a delta sync, then Recover -Batch $Batch again" }
}

function Get-PraOwnedResourceContact {
    param([object]$Intent)
    $name = [string](Get-PraValue $Intent 'contactName' '')
    if (-not $name) { return $null }
    $found = @(Get-Recipient -Filter ("Name -eq '{0}'" -f $name.Replace("'", "''")) -ResultSize Unlimited -ErrorAction Stop)
    if (-not $found.Count) { return $null }
    if ($found.Count -ne 1 -or [string]$found[0].Name -cne $name -or [string]$found[0].RecipientTypeDetails -ne 'MailContact') { throw 'Routing contact creation identity is ambiguous or not journal-owned.' }
    $id = [string]$found[0].Guid
    $persisted = [string](Get-PraValue $Intent 'contactId' '')
    if (-not $id -or ($persisted -and $persisted -ne $id)) { throw 'Routing contact immutable ID differs from the journal.' }
    $contact = Get-MailContact -Identity $id -ErrorAction Stop
    if (-not [bool](Get-PraValue $Intent 'contactCreationAttempted' $false)) { throw 'Routing contact creation was never journaled as attempted; pre-existing contact is not adopted.' }
    $marker = [string](Get-PraValue $contact ([string]$Intent.markerAttribute) '')
    $markerPending = Get-PraValue $Intent 'contactMarkerPending' $false
    if ($markerPending -isnot [bool]) { throw 'Routing contact marker progress is invalid; no contact is adopted.' }
    if ($marker -cne [string]$Intent.token -and ($marker -or ($persisted -and -not $markerPending))) { throw 'Routing contact ownership marker differs from the journal.' }
    return $contact
}

function Invoke-PraRecoverResourceWave {
    param([object[]]$Entries, [object]$HoldSet)
    foreach ($entry in $Entries) {
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            $intent = $entry.Intent
            $protected = (Get-Pra2ResourceProtectionMode $intent) -eq 'CaseHold'
            if (-not $protected -and (Get-PraValue $context.Config.Resources 'AllowUnprotectedDeletion' $false) -ne $true) {
                throw 'Resources.AllowUnprotectedDeletion = $true is required: explicitly acknowledge no hold/no inactive copy and loss of cloud-created reservations/messages after export/reconciliation. No mutation.'
            }
            if (-not $intent.cloudId -or -not $intent.cloudGuid -or $intent.cloudGuid -eq [string]$entry.Record.exchange_guid) { throw 'Native resource immutable IDs were not persisted. Resume Convert first to prove ownership; no deletion.' }
            $route = Get-Pra2ResourceRoute -Config $context.Config -Record $entry.Record -RequireReady
            $oldRoute = [string](Get-PraValue $intent 'route' '')
            if ($oldRoute -and $oldRoute -ine $route) { throw "Routing domain changed mid-Recover: keep the journaled destination $oldRoute until this batch is complete." }
            $intent | Add-Member -NotePropertyName route -NotePropertyValue $route -Force
            $intent | Add-Member -NotePropertyName contactName -NotePropertyValue (([string]$intent.creationName) -replace '^PRA-R-', 'PRA-C-') -Force
            $contact = Get-PraOwnedResourceContact -Intent $intent
            $live = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound
            if ($live) {
                if (-not $protected) { Assert-Pra2UnprotectedDeletionRepeatable -Intent $intent }
                $null = Get-Pra2ResourceRoute -Config $context.Config -Record $entry.Record -RequireReady
                $owned = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
                if (-not $owned -or $owned.User.accountEnabled -ne $false) { throw 'Native resource ownership/disabled identity cannot be verified; no deletion.' }
                if (-not $protected) {
                    Assert-Pra2UnprotectedResource -Owned $owned -CheckOrganization
                    $collisions = @(Get-Pra2ResourceCollisions -Record $entry.Record -Inventory (Get-Pra2ResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId)
                    if ($collisions.Count) { throw ($collisions -join '; ') }
                    $intent | Add-Member -NotePropertyName unprotectedDeletionAttempted -NotePropertyValue $true -Force
                    $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Intent' -Force
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceDeleteIntent' -Detail "Unprotected deletion acknowledged: exact $($intent.cloudId)/$($intent.cloudGuid); NO hold, NO inactive copy, cloud reservations/messages NOT preserved; route $route."
                    $final = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
                    if (-not $final -or $final.User.accountEnabled -ne $false) { throw 'Final resource ownership/disabled identity verification failed; no deletion.' }
                    Assert-Pra2UnprotectedResource -Owned $final -CheckOrganization
                    $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Issuing' -Force
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceDeleteIssuing' -Detail 'Deletion request starting; ambiguous/interrupted response cannot authorise another deletion of a restored identity.'
                    try {
                        $null = Invoke-Pra2Graph -Method DELETE -Uri ("v1.0/users/{0}" -f $intent.cloudId) -AllowNotFound -NoRetry
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
                            Write-PraResourceJournal -Entry $entry -Step 'ResourceDeleteRejected' -Detail "DELETE rejected with HTTP $code; no accepted deletion is recorded."
                        }
                        throw $failure
                    }
                    $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Issued' -Force
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceIdentityDeleted' -Detail 'Only journal-owned unprotected cloud identity delete issued; NO inactive copy/data preservation guaranteed. Deletion still to be verified; no purge.'
                } else {
                $group = Get-Pra2ResourceGroup -Id $intent.licensingGroupId -Skus @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value)
                if (-not (Test-Pra2ResourceLicence -User $owned.User -Group $group)) { throw 'Native resource enterprise licensing is not active before the hold; no deletion.' }
                $collisions = @(Get-Pra2ResourceCollisions -Record $entry.Record -Inventory (Get-Pra2ResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId)
                if ($collisions.Count) { throw ($collisions -join '; ') }
                $added = Add-Pra2HoldLocationSet -Set $HoldSet -Id @([string]$intent.cloudId)
                if ($added.Failed.ContainsKey([string]$intent.cloudId)) { throw $added.Failed[[string]$intent.cloudId] }
                $tag = Get-PraHoldTagOf -HoldSet $HoldSet -Id $intent.cloudId
                if (-not $tag) { throw 'No specific tool hold policy contains this native resource; no deletion.' }
                $intent | Add-Member -NotePropertyName holdTag -NotePropertyValue $tag -Force
                $held = Wait-Pra2Condition -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                    $now = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
                    if (-not $now -or $now.User.accountEnabled -ne $false -or -not (Test-Pra2ResourceLicence -User $now.User -Group $group)) { throw 'Resource ownership/disabled sign-in/enterprise licence changed while waiting for hold; no deletion.' }
                    @($now.Mailbox.InPlaceHolds | ForEach-Object { [string]$_ }) -contains $tag
                }
                if (-not $held.Ok) {
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceHoldPending' -Detail 'Specific tool hold not stamped yet; resource, identity and licence retained.'
                    Set-PraRowResult -Row $entry.Row -Status Pending -Detail 'Specific tool hold not stamped; identity NOT deleted. Retry Recover after hold deployment.'
                    continue
                }
                # Record the exact IDs/GUID and expected hold before the irreversible side effect.
                Write-PraResourceJournal -Entry $entry -Step 'ResourceDeleteIntent' -Detail "verified $($intent.cloudId)/$($intent.cloudGuid), hold $tag, enterprise licence active; route $route"
                $final = Get-Pra2OwnedResource -Intent $intent -Record $entry.Record
                if (-not $final -or $final.User.accountEnabled -ne $false -or -not (Test-Pra2ResourceLicence -User $final.User -Group $group) -or
                    @($final.Mailbox.InPlaceHolds | ForEach-Object { [string]$_ }) -notcontains $tag) { throw 'Final ownership/licence/hold verification failed immediately before deletion.' }
                $null = Invoke-Pra2Graph -Method DELETE -Uri ("v1.0/users/{0}" -f $intent.cloudId) -AllowNotFound
                Write-PraResourceJournal -Entry $entry -Step 'ResourceIdentityDeleted' -Detail 'Only the journal-owned cloud identity deleted; recycle bin and mailbox data retained.'
                }
            }
            if ($protected) {
            $expectedHold = [string](Get-PraValue $intent 'holdTag' '')
            if (-not $expectedHold -or -not $HoldSet.Tags.Contains($expectedHold)) { throw 'Expected tool hold was not journaled before deletion; verify the inactive mailbox manually. No routing contact is created.' }
            $inactive = Wait-Pra2Condition -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                # A filtered zero-result is absence, not a swallowed cmdlet/service failure.
                $found = @(Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited -ErrorAction Stop | Where-Object { [string]$_.ExchangeGuid -eq [string]$intent.cloudGuid })
                if (-not $found.Count) { return $false }
                if ($found.Count -ne 1 -or [string]$found[0].ExchangeGuid -ne [string]$intent.cloudGuid -or
                    @($found[0].InPlaceHolds | ForEach-Object { [string]$_ }) -notcontains $expectedHold) { throw 'Inactive resource GUID/specific tool hold verification failed; no contact, no purge.' }
                return $found[0]
            }
            if (-not $inactive.Ok) {
                Write-PraResourceJournal -Entry $entry -Step 'ResourceInactivePending' -Detail "Identity deleted; exact mailbox $($intent.cloudGuid) not yet inactive with expected hold. Recycle bin retained."
                Set-PraRowResult -Row $entry.Row -Status Pending -Detail "Identity deleted; wait for exact inactive GUID $($intent.cloudGuid) with $expectedHold, then resume Recover. No mailbox data purged."
                continue
            }
            if (Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound) { throw 'Deleted resource identity is active again; routing contact creation refused.' }
            Write-PraResourceJournal -Entry $entry -Step 'ResourceInactiveVerified' -Detail "Exact mailbox $($intent.cloudGuid) inactive under $expectedHold; mailbox and recycle-bin identity never purged."
            $collisions = @(Get-Pra2ResourceCollisions -Record $entry.Record -Inventory (Get-Pra2ResourceInventory -Records @($entry.Record)) -OwnedId $intent.cloudId -OwnedGuid $intent.cloudGuid -OwnedContactId $(if ($contact) { [string]$contact.Guid } else { '' }))
            } else {
                $deleted = Wait-Pra2Condition -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $context.Config.Polling.IntervalSeconds -Test {
                    Test-Pra2UnprotectedResourceDeleted -Intent $intent
                }
                if (-not $deleted.Ok) {
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceDeletionPending' -Detail "Exact $($intent.cloudId)/$($intent.cloudGuid) not yet verified deleted/soft-deleted. No inactive copy or purge; resume."
                    Set-PraRowResult -Row $entry.Row -Status Pending -Detail "Wait for journal-owned identity $($intent.cloudId) and soft-deleted non-inactive mailbox $($intent.cloudGuid), then resume Recover. Cloud reservations/messages NOT preserved; no purge."
                    continue
                }
                $intent | Add-Member -NotePropertyName unprotectedDeletionProgress -NotePropertyValue 'Verified' -Force
                Write-PraResourceJournal -Entry $entry -Step 'ResourceDeletionVerified' -Detail 'Exact identity deleted and mailbox soft-deleted, not inactive; foreign collisions and Exchange contact writes still to be checked. No purge/data preservation.'
                $inventory = Get-Pra2ResourceInventory -Records @($entry.Record)
                $collisions = @(Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $entry.Record -Inventory $inventory -OwnedContactId $(if ($contact) { [string]$contact.Guid } else { '' }))
                if ($collisions.Count) {
                    $reason = "Address release pending: $($collisions -join '; ') No automatic purge; resolve safely and resume this Recover batch."
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceAddressReleasePending' -Detail $reason
                    Set-PraRowResult -Row $entry.Row -Status Pending -Detail $reason
                    continue
                }
            }
            if ($collisions.Count) { throw ($collisions -join '; ') }
            if (Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound) { throw 'Deleted resource identity is active again; routing contact creation/configuration refused.' }
            if (-not $contact) {
                $intent | Add-Member -NotePropertyName contactCreationAttempted -NotePropertyValue $true -Force
                Write-PraResourceJournal -Entry $entry -Step 'ResourceContactIntent' -Detail "create owned contact $($intent.contactName); addresses $((Get-Pra2ResourceAddresses $entry.Record) -join ', '), destination $route"
                $alias = [string](Get-PraValue $entry.Record 'alias' '')
                if (-not $alias) { $alias = ([string]$entry.Record.primary_smtp_address -split '@')[0] }
                $created = New-MailContact -Name $intent.contactName -Alias $alias -ExternalEmailAddress $route -ErrorAction Stop
                $createdId = [string](Get-PraValue $created 'Guid' '')
                if ($createdId) {
                    $intent | Add-Member -NotePropertyName contactId -NotePropertyValue $createdId -Force
                    $intent | Add-Member -NotePropertyName contactMarkerPending -NotePropertyValue $true -Force
                    Write-PraResourceJournal -Entry $entry -Step 'ResourceContactCreated' -Detail "Returned routing contact immutable ID $createdId pinned before configuration."
                }
                $contact = Get-PraOwnedResourceContact -Intent $intent
                if (-not $contact) { throw 'Routing contact creation accepted but not visible; resume this batch by its journaled name.' }
            }
            if (-not (Get-PraValue $intent 'contactId' '')) {
                $intent | Add-Member -NotePropertyName contactId -NotePropertyValue ([string]$contact.Guid) -Force
                $intent | Add-Member -NotePropertyName contactMarkerPending -NotePropertyValue $true -Force
                Write-PraResourceJournal -Entry $entry -Step 'ResourceContactCreated' -Detail "Journal-owned routing contact immutable ID $($contact.Guid) pinned before configuration."
            }
            $parameters = @{ Identity = [string]$contact.Guid; ExternalEmailAddress = $route; EmailAddresses = @(Get-Pra2ResourceAddresses $entry.Record)
                HiddenFromAddressListsEnabled = [bool](Get-PraValue $entry.Record 'hidden_from_address_lists' $false); ErrorAction = 'Stop' }
            $parameters[[string]$intent.markerAttribute] = [string]$intent.token
            Set-MailContact @parameters
            $intent | Add-Member -NotePropertyName contactMarkerPending -NotePropertyValue $false -Force
            $contact = Get-PraOwnedResourceContact -Intent $intent
            if (([string]$contact.ExternalEmailAddress -replace '^smtp:', '') -ine $route -or
                [string]$contact.PrimarySmtpAddress -ine [string]$entry.Record.primary_smtp_address -or
                [bool]$contact.HiddenFromAddressListsEnabled -ne $parameters.HiddenFromAddressListsEnabled) { throw 'Routing contact destination/primary/hidden verification failed.' }
            foreach ($address in @(Get-Pra2ResourceAddresses $entry.Record)) {
                if ($address -notin @($contact.EmailAddresses | ForEach-Object { [string]$_ })) { throw "Routing contact address not applied: $address." }
            }
            $result = @{ contactId = [string]$contact.Guid; route = $route }
            $protectionText = if ($protected) { $result.inactiveGuid = $intent.cloudGuid; 'inactive cloud data and deleted identity retained under tool hold' } else {
                $result.deletedGuid = $intent.cloudGuid; 'NO hold/NO inactive copy; cloud-created reservations/messages NOT preserved, soft-deleted object retained without purge'
            }
            Write-PraResourceJournal -Entry $entry -Step 'ResourceRouted' -Status Done -Detail "Owned contact $($contact.Guid) routes to $route; $protectionText; no calendar merge." -Extra $result
            Set-PraRowResult -Row $entry.Row -Status Success -Field ExchangeOnline -Value MailContact -Detail "native resource routes to $route via owned contact; $protectionText. On-premises objects untouched; cloud calendars are not merged."
        } catch { Stop-PraRecoverEntry -Entry $entry -Step 'RecoverResource' -Message ($_.Exception.Message + ' Resolve the finding and resume the same Recover batch. No automatic purge or hold removal. Unprotected deletion does NOT preserve cloud reservations/messages.') }
    }
    $context.CurrentIdentity = ''
}

function Invoke-PraRecover {
    $config = $context.Config
    if (-not $Batch) { throw '-Batch is required: the ID of the Convert batch to roll back (printed at the end of Convert).' }
    $context.Journal = Open-Pra2Store -Path $config.Store.JournalPath -Root $context.Root -Kind Journal
    $convert = Get-Pra2Batch $context.Journal -Id $Batch
    if (-not $convert -or $convert.action -ne 'Convert') { throw "Convert batch $Batch not found in $($config.Store.JournalPath)." }
    $existing = @(Invoke-Pra2Sql $context.Journal "SELECT * FROM batch WHERE action = 'Recover' AND convert_batch = @c ORDER BY created_utc DESC" @{ c = $convert.id } -As Rows)
    $resume = $existing | Where-Object { $_.status -in @('Running','Partial') } | Select-Object -First 1
    $finished = $existing | Where-Object { $_.status -eq 'Complete' } | Select-Object -First 1
    if (-not $resume -and $finished) { throw "Convert batch $($convert.id) was already rolled back by Recover batch $($finished.id) (complete): nothing to do. Check shows the current state of the objects." }
    $data = Read-PraSnapshotData -Id ([long]$convert.snapshot_id) -Verb 'roll back' -All -NoCount
    $byGuid = @{}; foreach ($record in $data.Records) { $byGuid[[string]$record.object_guid] = $record }
    $convertItems = @(Get-Pra2BatchItem $context.Journal $convert.id | Where-Object { $_.status -ne 'Skipped' -and $_.step -and ($_.entra_id -or (Get-Pra2ResourceIntent -Item $_)) })
    $done = @{}; $previous = @{}
    if ($resume) { foreach ($item in @(Get-Pra2BatchItem $context.Journal $resume.id)) { $previous[[string]$item.object_guid] = $item; if ($item.status -eq 'Done') { $done[[string]$item.object_guid] = $true } } }
    $targets = @($convertItems | Where-Object { -not $done.ContainsKey([string]$_.object_guid) })
    if (@(Get-PraIdentityList).Count) {
        $targets = @(Select-PraIdentity -Items $targets -Where "the objects of batch $($convert.id) still to roll back" -Keys {
                param($t) $r = $byGuid[[string]$t.object_guid]
                @([string]$t.identity, [string]$t.object_guid) + $(if ($r) { @([string]$r.primary_smtp_address, [string]$r.user_principal_name, [string]$r.sam_account_name) } else { @() }) })
    }
    if ($Scope -eq 'UsersOnly') { $targets = @($targets | Where-Object { $_.kind -eq 'User' }) } elseif ($Scope -eq 'SharedOnly') { $targets = @($targets | Where-Object { $_.kind -ne 'User' }) }
    $context.NativeResourceScope = @($targets | Where-Object { Get-Pra2ResourceIntent -Item $_ }).Count -gt 0
    $requiresHold = @($targets | Where-Object {
            $intent = Get-Pra2ResourceIntent -Item $_
            -not $intent -or (Get-Pra2ResourceProtectionMode $intent) -eq 'CaseHold'
        }).Count -gt 0
    if ($requiresHold -and -not $config.Retention.HoldPolicy) { throw 'Retention.HoldPolicy is empty: protected/user/shared Recover requires that eDiscovery case hold.' }
    Write-PraItem -Context $context -Status Info -Icon Batch -Text ('Convert batch {0} of {1}: {2} object(s) to roll back{3}' -f $convert.id, $convert.created_utc, $targets.Count, $(if ($resume) { " (resuming Recover $($resume.id))" } else { '' }))

    Write-PraStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-Pra2Cloud -Context $context; if ($requiresHold) { Connect-Pra2Compliance -Context $context } } catch { $context.CloudConnectFailed = $true; throw }
    Write-PraItem -Context $context -Status Ok -Text ($how + $(if ($requiresHold) { ', Security & Compliance (case hold)' } else { '; unprotected native lifecycle, no compliance connection' }))

    Write-PraStep -Context $context -Title 'Prerequisites' -Icon Key
    $context.CurrentOperation = 'Prerequisites'
    $holdSet = $null
    if ($requiresHold) {
    $holdSet = Get-Pra2HoldPolicySet -Policy $config.Retention.HoldPolicy -Limit $config.Retention.HoldPolicyLimit
    Write-PraItem -Context $context -Status Ok -Text ('Case hold: {0} ({1} mailbox(es) held, {2} at most per policy; full policies are followed by {3}-02, -03... in the same case)' -f
        (@($holdSet.Policies | ForEach-Object { '{0} {1}' -f $_.Name, $_.Count }) -join ', '), (@($holdSet.Policies | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum, $holdSet.Limit, $holdSet.Base)
    Write-PraItem -Context $context -Status Info -Text ("Entra Connect {0} or later, operations by {1}{2}" -f $config.EntraConnect.MinVersion, $config.EntraConnect.Mode, $(if ($config.EntraConnect.Server) { " on $($config.EntraConnect.Server)" } else { '' }))
    }
    # Each object goes the way its current state allows; a step already done is never undone (a hold is never lifted
    # from an object that is back on-premises).
    #   User          user, or a shared mailbox whose Convert stopped half-way: back to AD like a user
    #   Back          already a synchronised MailUser with the on-premises GUID: case hold only
    #   Shared        SharedMailbox: hold, identity deleted, purged, recreated by Entra Connect
    #   SharedDeleted identity already deleted by an earlier run: purge check, recreated object
    $plan = [System.Collections.Generic.List[object]]::new()
    $resourceBlocked = [System.Collections.Generic.List[object]]::new()
    if (Test-PraStop) { throw 'Stopped by the operator before reading the state of the objects: nothing was changed.' }
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $exoStates = Get-Pra2ExoStateSet -Id ([string[]]@($targets | Where-Object { -not (Get-Pra2ResourceIntent -Item $_) } | ForEach-Object { [string]$_.entra_id })) -TagAttribute $config.Retention.TagAttribute
    if ($targets.Count) { Write-PraItem -Context $context -Status Info -Text ('Exchange Online state of {0} object(s) read in {1}' -f $targets.Count, (Format-PraDuration $clock.Elapsed.TotalSeconds)) }
    foreach ($item in $targets) {
        $guid = [string]$item.object_guid
        $record = $byGuid[$guid]
        $row = New-PraRow $(if ($record) { $record } else { [pscustomobject]@{ user_principal_name = $item.identity; kind = $item.kind; primary_smtp_address = ''; object_guid = $item.object_guid; exchange_guid = '' } })
        $row.Action = 'Recover'
        [void]$context.Rows.Add($row)
        if (-not $record) { Set-PraRowResult -Row $row -Status Error -Detail "not in snapshot $($convert.snapshot_id)"; continue }
        $intent = Get-Pra2ResourceIntent -Item $item
        if ($intent) {
            try {
                if ($previous.ContainsKey($guid)) {
                    $oldIntent = Get-Pra2ResourceIntent -Item $previous[$guid]
                    if ($oldIntent -and (Get-Pra2ResourceProtectionMode $oldIntent) -cne (Get-Pra2ResourceProtectionMode $intent)) { throw 'Recover protection mode differs from original Convert journal; no downgrade or deletion.' }
                    if ($oldIntent) {
                        foreach ($field in @('token','creationName','markerAttribute','cloudId','cloudGuid','cloudType','licensingGroupId')) {
                            if ([string](Get-PraValue $oldIntent $field '') -cne [string](Get-PraValue $intent $field '')) { throw "Recover ownership $field differs from original Convert journal; no identity is adopted or deleted." }
                        }
                    }
                    if ($oldIntent) { $intent = $oldIntent }
                }
                $protected = (Get-Pra2ResourceProtectionMode $intent) -eq 'CaseHold'
                if (-not $protected) {
                    $row.Warnings = 'NO hold, NO inactive copy: cloud-created reservations/messages NOT preserved or merged. Export/reconcile before explicitly allowing unprotected deletion.'
                    Write-PraItem -Context $context -Status Warn -Text "$($row.Identity): $($row.Warnings)"
                    if ((Get-PraValue $config.Resources 'AllowUnprotectedDeletion' $false) -ne $true) { throw 'Resources.AllowUnprotectedDeletion = $true required after export/reconciliation and explicit acknowledgement of cloud reservations/messages loss (no hold/no inactive copy). No mutation.' }
                }
                $route = Get-Pra2ResourceRoute -Config $config -Record $record -RequireReady
                if (-not $intent.cloudId -or -not $intent.cloudGuid) { throw 'Native resource immutable ownership IDs not persisted: resume Convert first; no deletion.' }
                $live = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id" -f $intent.cloudId) -AllowNotFound
                if ($live) {
                    if (-not $protected) { Assert-Pra2UnprotectedDeletionRepeatable -Intent $intent }
                    $owned = Get-Pra2OwnedResource -Intent $intent -Record $record
                    if (-not $owned -or $owned.User.accountEnabled -ne $false) { throw 'Native resource cloud-only/disabled ownership verification failed.' }
                    $null = Get-Pra2ResourceRoute -Config $config -Record $record -RequireReady
                    if (-not $protected) {
                        Assert-Pra2UnprotectedResource -Owned $owned -CheckOrganization
                    }
                }
                $nativeEntry = [pscustomobject]@{ Item = $item; Record = $record; Row = $row; EntraId = [string]$intent.cloudId; Exo = $null
                    Path = 'CloudResource'; CloudGuid = [string]$intent.cloudGuid; Original = @(); Intent = $intent }
                [void]$plan.Add($nativeEntry)
                $protectionText = if ($protected) { 'legacy hold/enterprise licence verification and exact inactive copy' } else { 'NO hold, NO inactive copy, cloud reservations/messages NOT preserved; verify deletion/soft-deleted state and address release without purge' }
                Write-PraItem -Context $context -Status Info -Text "Native resource: $protectionText; delete only owned $($intent.cloudId)/$($intent.cloudGuid), contact to $route; addresses: $((Get-Pra2ResourceAddresses $record) -join ', '). No sync/SOA, no calendar merge."
            } catch {
                Set-PraRowResult -Row $row -Status Error -Detail $_.Exception.Message
                $resourceBlocked.Add([pscustomobject]@{ Item = $item; Record = $record; Row = $row; Intent = $intent })
            }
            continue
        }
        $entry = [pscustomobject]@{ Item = $item; Record = $record; Row = $row; EntraId = [string]$item.entra_id; Exo = $null; Path = ''; CloudGuid = ''; Original = @(Get-PraOriginalLicence $item) }
        $exo = $exoStates[$entry.EntraId]
        if (-not $exo) { $exo = Get-Pra2ExoState -Identity $entry.EntraId -TagAttribute $config.Retention.TagAttribute }
        $entry.Exo = $exo
        $row.ExchangeOnline = $exo.Type; $row.Holds = ($exo.Holds -join ', ')
        $lastStep = if ($previous.ContainsKey($guid)) { [string](Get-PraValue $previous[$guid] 'step' '') } else { '' }
        if ($item.kind -ne 'User' -and $lastStep -in @('Inactive', 'Purged') -and -not $exo.Type) {
            $result = $null; try { $result = [string](Get-PraValue $previous[$guid] 'result_json' '') | ConvertFrom-Json } catch { $result = $null }
            $entry.Path = 'SharedDeleted'; $entry.CloudGuid = [string](Get-PraValue $result 'inactiveGuid' ''); $row.ExchangeOnline = 'inactive'
            [void]$plan.Add($entry); continue
        }
        if (-not $exo.Type) { Set-PraRowResult -Row $row -Status Error -Detail ("no Exchange Online recipient for {0}: the identity was deleted or replaced outside this batch. Check it in Entra ID and Exchange Online; nothing was changed." -f $entry.EntraId); continue }
        if ($exo.Type -eq 'MailUser' -and $exo.ExchangeGuid -eq [string]$record.exchange_guid -and $exo.IsDirSynced -eq $true) { $entry.Path = 'Back'; [void]$plan.Add($entry); continue }
        if ($item.kind -ne 'User' -and $exo.Type -eq 'SharedMailbox') { $entry.Path = 'Shared'; [void]$plan.Add($entry); continue }
        $other = @($exo.Holds | Where-Object { $_ -and -not $holdSet.Tags.Contains([string]$_) })
        if ($exo.Type -eq 'UserMailbox' -and $other.Count) { Set-PraRowResult -Row $row -Status Error -Detail ('other hold(s) on the cloud mailbox block the rollback: {0}. Remove them, then run Recover again.' -f ($other -join ', ')); continue }
        $entry.Path = 'User'
        [void]$plan.Add($entry)
    }
    $planUsers = @($plan | Where-Object { $_.Path -eq 'User' })
    $planBack = @($plan | Where-Object { $_.Path -eq 'Back' })
    $planShared = @($plan | Where-Object { $_.Path -in @('Shared', 'SharedDeleted') })
    $planResources = @($plan | Where-Object { $_.Path -eq 'CloudResource' })
    $partial = @($planUsers | Where-Object { $_.Item.kind -ne 'User' })
    Write-PraItem -Context $context -Status Info -Icon Target -Text ('{0} user(s) and {1} shared mailbox(es) to roll back{2}{3}' -f @($planUsers + $planBack | Where-Object { $_.Item.kind -eq 'User' }).Count, ($planShared.Count + $partial.Count + @($planBack | Where-Object { $_.Item.kind -ne 'User' }).Count),
        $(if ($planBack.Count) { "; $($planBack.Count) already back on-premises (case hold only)" } else { '' }), $(if ($partial.Count) { "; $($partial.Count) shared mailbox(es) not fully converted, rolled back like users" } else { '' }))

    Write-PraStep -Context $context -Title 'Plan and confirmation' -Icon Plan
    Write-PraItem -Context $context -Status Info -Text ('Users, in waves of {0}: source of authority back to AD, delta sync, Exchange plan removed (licences as before Convert), MailUser with the on-premises GUID, cloud mailbox kept as ComponentShared under the case hold' -f [Math]::Max(1, [int](Get-PraValue $context 'UserWaveSize' 500)))
    if ($planShared.Count) { Write-PraItem -Context $context -Status Info -Text ('Shared mailboxes, in waves of {0} (Licensing.Shared.Parallel): case hold, Entra Connect scheduler paused, identities deleted (mailboxes inactive), deleted permanently, scheduler resumed, recreated from AD by a delta sync' -f $config.Licensing.Shared.Parallel) }
    $recoverEstimate = Get-PraDurationEstimate -UserWaves ([Math]::Ceiling(@($planUsers + $planBack).Count / [Math]::Max(1, [int](Get-PraValue $context 'UserWaveSize' 500)))) -SharedRounds ([Math]::Ceiling($planShared.Count / [Math]::Max(1, [int]$config.Licensing.Shared.Parallel)))
    if ($recoverEstimate.Text) { Write-PraItem -Context $context -Status Info -Icon Clock -Text ('Rough estimate: about {0} (typical tenant; a slow Microsoft 365 evening can take much longer)' -f $recoverEstimate.Text) }
    foreach ($entry in $plan) { $entry.Row.FinalStatus = 'Planned'; Write-PraLog -Context $context -Message ('Planned: {0} ({1})' -f $entry.Row.Identity, $entry.Path) -Level Sub }
    if ($context.Mode -ne 'Apply') {
        foreach ($title in @('Users: back on-premises', 'Shared mailboxes', 'Licence group')) { Skip-PraStep -Title $title -Reason 'Preview: not run.' -Icon Write }
        $context.NextSteps = @("pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch $Batch" + (Get-PraIdentityArgument))
        return
    }
    if (-not $plan.Count) { throw 'Nothing to roll back (see above).' }
    $lossWarning = if (@($planResources | Where-Object { (Get-Pra2ResourceProtectionMode $_.Intent) -eq 'None' }).Count) { ' UNPROTECTED native resources: NO hold, NO inactive copy; cloud-created reservations/messages are NOT preserved or merged. Deletion may destroy them. Acknowledge after export/reconciliation.' } else { '' }
    Confirm-PraApply (('Roll back {0} object(s) of Convert batch {1}?' -f $plan.Count, $convert.id) + $lossWarning)
    if ($resume) { $context.BatchId = [string]$resume.id }
    else {
        $context.BatchId = New-Pra2Batch $context.Journal Recover @{ ConvertBatch = $convert.id; SnapshotId = [long]$convert.snapshot_id; Environment = $config.Environment
            Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name; ToolVersion = $toolVersion; ScopeJson = (@{ Identity = $Identity; IdentityPath = $IdentityPath; Identities = @(Get-PraIdentityList).Count; Scope = $Scope } | ConvertTo-Json -Compress) }
    }
    foreach ($entry in $plan) {
        $values = @{ kind = [string]$entry.Item.kind; identity = $entry.Row.Identity; entra_id = $entry.EntraId; status = 'Running' }
        if ($entry.Path -eq 'CloudResource') { $values.original_json = ($entry.Intent | ConvertTo-Json -Depth 15 -Compress) }
        Set-Pra2BatchItem $context.Journal $context.BatchId ([string]$entry.Item.object_guid) $values
    }
    foreach ($entry in $resourceBlocked) {
        Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'ResourcePreconditionFailed' -Outcome Fail -Detail $entry.Row.Detail -Item @{
            kind = [string]$entry.Item.kind; identity = $entry.Row.Identity; entra_id = [string]$entry.Intent.cloudId; status = 'Failed'; step = 'ResourcePreconditionFailed'
            original_json = ($entry.Intent | ConvertTo-Json -Depth 15 -Compress); message = $entry.Row.Detail }
    }
    Write-PraItem -Context $context -Status Ok -Icon Batch -Text ("Recover batch {0} in {1}" -f $context.BatchId, $config.Store.JournalPath)
    $stopped = $false
    # Waves: every phase of a wave runs for all its objects (one wait per phase); the operator's stop is honoured
    # between waves, a wave in progress is always finished.
    $runWaves = {
        param([object[]]$Entries, [int]$Size, [string]$Label, [scriptblock]$Run)
        $clock = [Diagnostics.Stopwatch]::StartNew()
        for ($start = 0; $start -lt $Entries.Count; $start += $Size) {
            $wave = @($Entries[$start..([Math]::Min($Entries.Count, $start + $Size) - 1)])
            if ($stopped -or (Test-PraStop)) { $stopped = $true; foreach ($one in $wave) { Set-PraRowStopped $one.Row }; continue }
            if ($Entries.Count -gt $Size) { Write-PraProgress -Context $context -Phase $Label -Done $start -Total $Entries.Count -ElapsedSeconds $clock.Elapsed.TotalSeconds }
            & $Run $wave
        }
    }

    Write-PraStep -Context $context -Title 'Users: back on-premises' -Icon Sync
    $context.CurrentOperation = 'Recover-Users'
    $userEntries = @($plan | Where-Object { $_.Path -in @('User', 'Back') })
    . $runWaves -Entries $userEntries -Size ([Math]::Max(1, [int](Get-PraValue $context 'UserWaveSize' 500))) -Label 'Users' -Run { param($wave) Invoke-PraRecoverUserWave -Entries $wave -HoldSet $holdSet }
    if (-not $userEntries.Count) { Write-PraItem -Context $context -Status Skip -Text 'No user to bring back.' }

    Write-PraStep -Context $context -Title 'Shared mailboxes' -Icon Mail
    $context.CurrentOperation = 'Recover-Shared'
    # The Entra Connect scheduler is paused only while the identities of one wave are deleted.
    . $runWaves -Entries $planShared -Size $config.Licensing.Shared.Parallel -Label 'Shared mailboxes' -Run { param($wave) Invoke-PraRecoverSharedWave -Entries $wave -HoldSet $holdSet }
    if (-not $planShared.Count) { Write-PraItem -Context $context -Status Skip -Text 'No shared mailbox in this batch.' }
    . $runWaves -Entries $planResources -Size ([Math]::Max(1, [int](Get-PraValue (Get-PraValue $config 'Resources' @{}) 'Parallel' 100))) -Label 'Native resources' -Run { param($wave) Invoke-PraRecoverResourceWave -Entries $wave -HoldSet $holdSet }
    $context.CurrentIdentity = ''
    Write-PraStep -Context $context -Title 'Licence group' -Icon Key
    $context.CurrentOperation = 'Licence-Group'
    # The group serves every Convert batch of this journal: it goes back to AD only when no converted user is left,
    # whichever batch switched it (later waves record nothing, the group was already in the cloud).
    $changes = @(if (@($targets | Where-Object { -not (Get-Pra2ResourceIntent -Item $_) }).Count) {
            Invoke-Pra2Sql $context.Journal "SELECT * FROM tenant_change WHERE kind = 'GroupSoa' AND restored = 0" -As Rows
        })
    if (-not $changes.Count) { Write-PraItem -Context $context -Status Skip -Text 'No licence group switched to the cloud by this tool.' }
    else {
        $pending = @(Invoke-Pra2Sql $context.Journal @"
SELECT DISTINCT ci.identity AS identity FROM batch_item ci JOIN batch cb ON cb.id = ci.batch_id
WHERE cb.action = 'Convert' AND ci.kind = 'User' AND ci.step IS NOT NULL
  AND NOT EXISTS (SELECT 1 FROM batch_item ri JOIN batch rb ON rb.id = ri.batch_id
                  WHERE rb.action = 'Recover' AND rb.convert_batch = cb.id AND ri.object_guid = ci.object_guid AND ri.status = 'Done')
"@ -As Rows)
        if ($pending.Count) {
            Write-PraItem -Context $context -Status Warn -Text ('Licence group kept in the cloud: {0} converted user(s) not back yet ({1}); it goes back to AD with the last one.' -f $pending.Count, ((@($pending | Select-Object -First 5 | ForEach-Object { [string]$_.identity })) -join ', '))
        } else {
            foreach ($change in $changes) {
                Set-Pra2SourceOfAuthority -Id $change.target_id -Cloud $false -Type groups
                Set-Pra2TenantChange $context.Journal $change.batch_id 'GroupSoa' $change.target_id -Restored
                Write-PraJournal -Step 'GroupSoaAd' -Detail "licence group $($change.target_id) back to AD"
                Write-PraItem -Context $context -Status Ok -Text "Licence group $($change.target_id) managed by AD again (next Entra Connect cycle)"
            }
        }
    }
    $items = @(Get-Pra2BatchItem $context.Journal $context.BatchId)
    $counts = @{ done = @($items | Where-Object status -eq 'Done').Count; failed = @($items | Where-Object status -eq 'Failed').Count; running = @($items | Where-Object status -eq 'Running').Count }
    $complete = $counts.done -eq @($convertItems).Count
    Set-Pra2Batch $context.Journal $context.BatchId $(if ($complete) { 'Complete' } else { 'Partial' }) $counts
    $context.NextSteps = @($(if (-not $complete) { "pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch $($convert.id)   # resume the objects not finished" })) | Where-Object { $_ }
    $createdLists = @(Get-Pra2TenantChange $context.Journal $convert.id | Where-Object kind -eq 'NativeRoomList' |
        ForEach-Object { [string]$_.original_json | ConvertFrom-Json -ErrorAction Stop } | Where-Object { $_.operation -eq 'Create' -and $_.creationAttempted })
    if ($createdLists.Count) {
        Write-PraItem -Context $context -Status Info -Text "$($createdLists.Count) journaled cloud room list(s) retained: Recover never deletes groups or inserts routing contacts into RoomLists; review room discovery separately."
    }
}
# =================================================================================================
# Main
# =================================================================================================
try {
    $context.CurrentOperation = 'Read-Configuration'
    $context.Config = Import-PraConfiguration -Path $ConfigPath -Root $PSScriptRoot
    $config = $context.Config
    $context.LogFolder = $config.Logging.Folder
    $context.ReportFolder = $config.Report.Folder
    $context.NoReport = -not $config.Report.Enabled

    Initialize-PraAudit $context

    if ($Identity -and $IdentityPath) { throw 'Use -Identity (one object) or -IdentityPath (a list), not both.' }
    if ($IdentityPath) {
        if ($Action -eq 'Collect') { throw '-IdentityPath is for Check, Convert and Recover (Collect reads the scope of the configuration, or one object with -Identity).' }
        $context.Identities = @(Read-PraIdentityFile -Path $IdentityPath)
        Write-PraLog -Context $context -Message ('Identity list {0}: {1} object(s)' -f $IdentityPath, $context.Identities.Count) -Level Detail
    }

    $dot = [char]0x00B7
    $modeText = if ($Action -eq 'Check') { 'Read-only' } elseif ($effectiveMode -eq 'Apply') { $(if ($Action -eq 'Collect') { 'Apply (the snapshot is written)' } else { 'Apply (changes are made)' }) } else { 'Preview (nothing is changed)' }
    $banner = [ordered]@{ 'Action' = @('Mode', $Action); 'Mode' = @($(if ($effectiveMode -eq 'Apply') { 'Write' } else { 'Plan' }), $modeText) }
    $objects = if ($IdentityPath) { "list $(Split-Path $IdentityPath -Leaf) ($(@($context.Identities).Count) object(s))" } elseif ($Identity) { "one object: $Identity" } else { '' }
    $banner['Scope'] = @('Target', $(if ($Action -eq 'Collect') { Get-PraScopeText $config } elseif ($Batch) { "batch $Batch" + $(if ($objects) { " $dot $objects" } else { '' }) } else { $(if ($objects) { $objects } else { "snapshot $(if ($Snapshot) { $Snapshot } else { 'last complete' })" + $(if ($Scope -ne 'All') { " $dot $Scope" } else { '' }) }) }))
    if ($Action -ne 'Collect' -and $config.Cloud.Organization) { $banner['Tenant'] = @('Cloud', $config.Cloud.Organization) }
    $banner['Database'] = @('Folder', $config.Store.Path)
    $banner['Config'] = @('Config', ('{0} {1} Environment {2}' -f (Split-Path $config._Path -Leaf), $dot, $config.Environment))
    $banner['Log'] = @('Log', $context.LogFile)
    Write-PraBanner -Context $context -Title 'PRA Cloud Mailbox' -Subtitle "Exchange DR $dot scenario 2 $dot on-premises lost $([char]0x2192) Exchange Online" -Details $banner

    $context.StepTotal = @{ Collect = 5; Check = 4; Convert = 7; Recover = 7 }[$Action]
    switch ($Action) { 'Collect' { Invoke-PraCollect } 'Check' { Invoke-PraCheck } 'Convert' { Invoke-PraConvert } 'Recover' { Invoke-PraRecover } }
}
catch {
    $failure = $_
    try {
        $known = @($context.Issues.ToArray() | Where-Object { ([string](Get-PraValue $_ 'Message' '')).Contains($failure.Exception.Message) }).Count -gt 0
        $stopMessage = "Stopped at step '{0}' (operation {1}{2}): {3}" -f $context.CurrentPhase, $context.CurrentOperation, $(if ($context.CurrentIdentity) { ', object ' + $context.CurrentIdentity } else { '' }), $failure.Exception.Message
        Write-PraLog -Context $context -Message $stopMessage -Level $(if ($known) { 'Detail' } else { 'Error' })
        Write-PraLog -Context $context -Message ("ErrorId={0} | {1} | {2}" -f $failure.FullyQualifiedErrorId, $failure.InvocationInfo.PositionMessage, $failure.ScriptStackTrace) -Level Debug
    }
    catch { $context.Issues.Add([pscustomobject]@{ Message = $failure.Exception.Message; Source = 'Execution' }) }
    $context.ExitCode = 1
}
finally {
    if ($Action -eq 'Collect' -and (Get-Command Disconnect-Pra2Exchange -ErrorAction SilentlyContinue)) {
        try { Disconnect-Pra2Exchange -Context $context } catch { $context.Issues.Add([pscustomobject]@{ Message = "Exchange session: $($_.Exception.Message)"; Source = 'Exchange' }) }
    }
    if ($Action -ne 'Collect' -and (Get-Command Disconnect-Pra2Cloud -ErrorAction SilentlyContinue)) {
        try { Disconnect-Pra2Cloud -Context $context } catch { $context.Issues.Add([pscustomobject]@{ Message = "Cloud sign-out: $($_.Exception.Message)"; Source = 'Cloud' }) }
    }
    if ($context.Journal) { try { Close-Pra2Store $context.Journal } catch { $null = $_ } }
    $result = Complete-PraRun -Context $context
    if ($PassThru) { $result }
    exit $result.ExitCode
}

# SIG # Begin signature block
# MIIghQYJKoZIhvcNAQcCoIIgdjCCIHICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAAqi3LAvSkKrAe
# iWN1n+XPozUfrRf7wBlE4vXkAvs3J6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYISbDCCEmgCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDnzrNNUWKl
# aCMIMuZ/l92yGK2zxFMRACkjw7MHcejBNTANBgkqhkiG9w0BAQEFAASCAgA8xenb
# wzpPscVjtIYj+Xs43JEPF0jvHgodU5Hr7Msl6aEKTRxlvRW8JM9Pe1yLQ0BqiSh/
# sDFDsju0Ca9zLcfl8/CV+8zqWbq4MJF4i44VtcNn7oVirbmw4qmP0xYKrN1oEs1X
# pQdRaGDmSYpZWa9zPaZk6ldwnCc3kfJ06oki/qk8/yJo7VpY+o7/6+FaWjcCIrdu
# +7ZfqRBLmO4fxpwD0dRv9s22/Pud4eG1V44YglTUPltjH3+4aLrwIFagkdsv3ubv
# +sU4l7cny/UcuKDCIyylGrcCWRAlV8Cdq1KlKXJgYgAnB8kIApV3t3H6KP2JLWhV
# vOTvnLhnWlobZNyAVmD8tDabjuJG1ZMsmVdKVpJlL6PqKec4sn3W6i3VOdea3Y0m
# B14lGRGS5UXkLMHQVAZq/THflM7pjC9Qda6HmmfRyh5hO5FbcXfG9bK2SnRSeY/y
# qgB6MTsGQI2rCXdo56oHxcR+yq9fTkKSbtka28Wm0lLmSO5uWdASC3jJHgkne+KJ
# I9g/gjp72xcbYBSWqKfLcQb+86Fzqo01A+eA2cMILgYNpcce5Jlq8TFbjca51YZv
# zRVHYF0KLhzNpgiI6W4COG+dcnHNg2MCT5136uPdHLZ+uowoN+Fe72TEkq57X1IG
# jnpSnG0xRrwkGNcwNYclcCZve2nI8dI93fKOI6GCDykwgg8lBgorBgEEAYI3AwMB
# MYIPFTCCDxEGCSqGSIb3DQEHAqCCDwIwgg7+AgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIKJwrFWe5cP5PTqEDfOUKNnFyOSOalYY+J5Nx2PKqpIzAghR0r3c9B1d
# dxgPMjAyNjEwMTQxNTIzMzVaMAMCAQECEH3nmINPa8j7AHu0NaKH8X2gggwAMIIE
# /DCCAuSgAwIBAgIQH2sWYtIuG2xd8cDBoGAOODANBgkqhkiG9w0BAQsFADBzMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS8wLQYDVQQDDCZTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSBSMTAeFw0yNTAyMTgxNjMyMDJaFw0zNDExMTIxODUwMDVa
# MG4xCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3Rv
# bjERMA8GA1UECgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1w
# aW5nIFVuaXQgMjAyNSBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABBv7UVHH
# b+ZVluxPXlfE3M6tg0Xnq8dic+O3vPOCRpalUM1vO9A+GRzSVVjyygHhYrBw62XL
# Fh1kv7e+yRd/aTajggFaMIIBVjAfBgNVHSMEGDAWgBQMnRAljpqnG5mHQ88IfuG9
# gZD0zzBRBggrBgEFBQcBAQRFMEMwQQYIKwYBBQUHMAKGNWh0dHA6Ly9jZXJ0LnNz
# bC5jb20vU1NMLmNvbS10aW1lU3RhbXBpbmctSS1SU0EtUjEuY2VyMFEGA1UdIARK
# MEgwPAYMKwYBBAGCqTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNz
# bC5jb20vcmVwb3NpdG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwRgYDVR0fBD8wPTA7oDmgN4Y1aHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29t
# LXRpbWVTdGFtcGluZy1JLVJTQS1SMS5jcmwwHQYDVR0OBBYEFM582cAEgMUkEGoJ
# 6hyrJT0R/ajSMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAgHN1
# JIW9ZlNip07FfY9r1LJwLKTl1h/y9r1BjDgHMEz7U7J7mAy32JxA1KLyjCCJqNku
# RfhDHGFEfZtjTOru7AQMDSDiUDvCCc3jaz5py9jGvwIe5npkLbXqXhA+SqCzG719
# USigwhHkYr7v8NiXoaY0Xu03p0BiZZreZlpngmGB9+86N94FqIFx2I5Z5iuu7d4j
# 7aAOZBNxfU9jFDIIGDtWfWeCpF0Q8rLE/3DHDlzJKs5Il3i/VoOl7rTU938oSrxM
# www6GDBlKX8rGHxvkjgIb9MXkVglJSDHsNAZDeCS/inaxKHA+ChuB8K4OxIxa6k1
# e+eHBypakAY3wuN7w7PlPmhN6k5IdqZn3HZHs2VjoR841Z9wmVPFdGDdkBZ7XOmb
# 5OZUaKufsvtER0VbY2DkzaqAaItt1kc+I+FUz7PiVz2PUzpegkftRRxvrPyIL08b
# lJTnbMQ0XrqDN0rEfuTLv4XhhnKyJaXhxCUdjO/cRumMzZ35QTUMItBWy1xGp0iy
# oFVnAya2pPsVMuAI+sy1zxlOS9l5iSxvKJ8gpnpOgjqYTa7u1eyZo+4JlcGoiiR1
# 7LthMTrF1q62tO5xOpI8txUZ6gtKGtkJV6wdj+vZLcwKZ2M93GqVpOo6UBlGjx1s
# KsVZ7AIUzJZli4NJ0M2KYcrfNnlebkYKpDnSqWwwggb8MIIE5KADAgECAhBtUhhw
# h+gjTYVgANCAj5NWMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTE5MTExMzE4NTAwNVoXDTM0MTExMjE4NTAwNVowczELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCu
# URAT0vk8IKAghd7JUBxkyeH9xek0/wp/MUjoclrFXqhh/fGH91Fc+7fm0MHCE7A+
# wmOiqBj9ODrJAYGq3rm33jCnHSsCBNWAQYyoauLq8IjqsS1JlXL29qDNMMdwZ8UN
# zQS7vWZMDJ40JSGNphMGTIA2qn2bohGtgRc4p1395ESypUOaGvJ3t0FNL3BuKmb6
# YctMcQUF2sqooMzd89h0E6ujdvBDo6ZwNnWoxj7YmfWjSXg33A5GuY9ym4QZM5OE
# Vgo8ebz/B+gyhyCLNNhh4Mb/4xvCTCMVmNYrBviGgdPZYrym8Zb84TQCmSuX0JlL
# La6WK1aO6qlwISbb9bVGh866ekKblC/XRP20gAu1CjvcYciUgNTrGFg8f8AJgQPO
# Cc1/CCdaJSYwhJpSdheKOnQgESgNmYZPhFOC6IKaMAUXk5U1tjTcFCgFvvArXtK4
# azAWUOO1Y3fdldIBL6LjkzLUCYJNkFXqhsBVcPMuB0nUDWvLJfPimstjJ8lF4S6E
# CxWnlWi7OElVwTnt1GtRqeY9ydvvGLntU+FecK7DbqHDUd366UreMkSBtzevAc9a
# qoZPnjVMjvFqV1pYOjzmTiVHZtAc80bAfFe5LLfJzPI6DntNyqobpwTevQpHqPDN
# 9qqNO83r3kaw8A9j+HZiSw2AX5cGdQP0kG0vhzfgBwIDAQABo4IBgTCCAX0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTCBgwYIKwYBBQUHAQEEdzB1MFEGCCsGAQUFBzAChkVodHRwOi8vd3d3LnNzbC5j
# b20vcmVwb3NpdG9yeS9TU0xjb21Sb290Q2VydGlmaWNhdGlvbkF1dGhvcml0eVJT
# QS5jcnQwIAYIKwYBBQUHMAGGFGh0dHA6Ly9vY3Nwcy5zc2wuY29tMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwOwYDVR0fBDQwMjAwoC6gLIYq
# aHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0GA1Ud
# DgQWBBQMnRAljpqnG5mHQ88IfuG9gZD0zzAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAJIZdQ2mWkLPGQfZ8vyU+sCb8BXpRJZaL3Ez3VDlE3uZk3cP
# xPtybVfLuqaci0W6SB22JTMttCiQMnIVOsXWnIuAbD/aFTcUkTLBI3xys+wEajzX
# aXJYWACDS47BRjDtYlDW14gLJxf8W6DQoH3jHDGGy8kGJFOlDKG7/YrK7UGfHtBA
# EDVe6lyZ+FtCsrk7dD/IiL/+Q3Q6SFASJLQ2XI89ihFugdYL77CiDNXrI2MFspQG
# swXEAGpHuaQDTHUp/LdR3TyrIsLlnzoLskUGswF/KF8+kpWUiKJNC4rPWtNrxlbX
# YRGgdEdx8SMjUTDClldcrknlFxbqHsVmr9xkT2QtFmG+dEq1v5fsIK0vHaHrWjMM
# maJ9i+4qGJSD0stYfQ6v0PddT7EpGxGd867Ada6FZyHwbuQSadMb0K0P0OC2r7rw
# qBUe0BaMqTa6LWzWItgBjGcObXeMxmbQqlEz2YtAcErkZvh0WABDDE4U8GyV/32F
# daAvJgTfe9MiL2nSBioYe/g5mHUSWAay/Ip1RQmQCvmF9sNfqlhJwkjy/1U1ibUk
# TIUBX3HgymyQvqQTZLLys6pL2tCdWcjI9YuLw30rgZm8+K387L7ycUvqrmQ3ZJlu
# jHl3r1hgV76s3WwMPgKk1bAEFMj+rRXimSC+Ev30hXZdqyMdl/il5Ksd0vhGMYIC
# WDCCAlQCAQEwgYcwczELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgUjECEB9rFmLSLhtsXfHAwaBg
# DjgwCwYJYIZIAWUDBAIBoIIBYTAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# HAYJKoZIhvcNAQkFMQ8XDTI2MTAxNDE1MjMzNVowKAYJKoZIhvcNAQk0MRswGTAL
# BglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEINddWB9HniDd
# evXuue0vSyvYMS+a63Dc9P4RRHqOTHdPMIHJBgsqhkiG9w0BCRACLzGBuTCBtjCB
# szCBsAQgVCr5oWqNci5mEUl4iumUwYqaruWmXLNEolSa+Wx5x4swgYswd6R1MHMx
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjER
# MA8GA1UECgwIU1NMIENvcnAxLzAtBgNVBAMMJlNTTC5jb20gVGltZXN0YW1waW5n
# IElzc3VpbmcgUlNBIENBIFIxAhAfaxZi0i4bbF3xwMGgYA44MAoGCCqGSM49BAMC
# BEcwRQIhAM/+aSJeEoSaGN27jaX9fouFfU5tWNLHdpIOi7mRZAEPAiAnHqO8F8gc
# b//s6T79RiYbnMAzXwyGAgfqGTby9yg40w==
# SIG # End signature block
