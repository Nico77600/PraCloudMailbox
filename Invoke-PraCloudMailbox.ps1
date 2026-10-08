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
                Shared mailboxes, one at a time: temporary licence, shared, retention tag, permissions from the
                snapshot, licence removed. Every step is journaled (Store.JournalPath) under a batch ID.
      Recover   Infrastructure rebuilt (PowerShell 7). Users: source of authority back to AD, delta sync, Exchange
                plan removed; the cloud mailbox stays as ComponentShared under the case hold (Retention.HoldPolicy).
                Shared mailboxes: case hold, scheduler paused, identity deleted (the mailbox becomes inactive),
                deleted permanently, recreated from AD by a delta sync.

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
    No confirmation prompt. Collect asks one only before writing a snapshot of ONE object (-Identity -Mode Apply).

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
    Version    : 1.1.0
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
$toolVersion = '1.1.0'
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
        [void]$records.Add((ConvertTo-Pra2MailboxRecord -Mailbox $mailbox -Statistics $statistics))
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
    $sharedRecords = @($records | Where-Object { $_.kind -ne 'User' })
    if ($collect.SharedPermissions -and $sharedRecords.Count) {
        Write-PraStep -Context $context -Title 'Shared mailbox permissions' -Icon Key
        $context.CurrentOperation = 'Get-SharedPermission'
        foreach ($mailbox in @($mailboxes | Where-Object { [string]$_.RecipientTypeDetails -ne 'UserMailbox' })) {
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

    Write-PraStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-Pra2Cloud -Context $context } catch { $context.CloudConnectFailed = $true; throw }
    Write-PraItem -Context $context -Status Ok -Text $how
    if ($context.ContainsKey('CertificateWarning')) { Write-PraLog -Context $context -Message $context['CertificateWarning'] -Level Warning }

    Write-PraStep -Context $context -Title 'Tenant prerequisites' -Icon Key
    $context.CurrentOperation = 'Tenant-Facts'
    $tenant = Get-Pra2TenantFact -Context $context
    $roles = @(Get-Pra2GraphRole)
    # The state of every object first: the licence units really needed depend on what each one already holds.
    $states = @{}
    foreach ($record in $records) {
        if (Test-PraStop) { throw ('Check stopped by the operator while reading the cloud state ({0} of {1} object(s) read).' -f $states.Count, $records.Count) }
        try { $states[[string]$record.object_guid] = Get-Pra2CloudState -Context $context -Record $record } catch { $states[[string]$record.object_guid] = $_ }
    }
    $need = Get-PraLicenceNeed -Entries @($records | ForEach-Object { $s = $states[[string]$_.object_guid]; [pscustomobject]@{ Record = $_; State = $(if ($s -is [System.Management.Automation.ErrorRecord]) { $null } else { $s }); Step = '' } }) -Tenant $tenant
    $tenantFindings = Test-Pra2TenantReadiness -Config $config -Roles $roles -Tenant $tenant -UserCount $need.Users -SharedCount $need.Shared
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
    $lookupCache = @{}
    $lookup = {
        param($identity)
        $key = ([string]$identity).ToLowerInvariant()
        if (-not $lookupCache.ContainsKey($key)) { $lookupCache[$key] = [bool](Get-Recipient -Identity $identity -ErrorAction SilentlyContinue) }
        return $lookupCache[$key]
    }.GetNewClosure()
    $ready = 0
    foreach ($record in $records) {
        $row = New-PraRow $record
        $context.CurrentIdentity = $row.Identity
        try {
            $state = $states[[string]$record.object_guid]
            if ($state -is [System.Management.Automation.ErrorRecord]) { throw $state.Exception }
            $findings = @(Get-Pra2Readiness -Config $config -Record $record -State $state -Skus @($tenant.Skus))
            if ($record.kind -ne 'User') { $findings += @(Get-Pra2TrusteeFinding -Record $record -Permissions $permissions -Members $members -Lookup $lookup) }
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
    $users = 0; $shared = 0
    foreach ($entry in $Entries) {
        $user = if ($entry.State) { $entry.State.User } else { $null }
        $held = @(if ($user -and $user.PSObject.Properties['assignedLicenses']) { $user.assignedLicenses | ForEach-Object { [string]$_.skuId } })
        if ($entry.Record.kind -eq 'User') {
            if ($config.Licensing.Users.Mode -eq 'Kiosk' -or @($held | Where-Object { $_ -in $userSkus }).Count) { continue }
            $users++
        } else {
            if ($entry.State -and $entry.State.RecipientType -eq 'SharedMailbox') { continue }
            if ([string]$entry.Step -in @('Shared', 'Tagged', 'LicenceRemoved', 'Done') -or @($held | Where-Object { $_ -in $sharedSkus }).Count) { continue }
            $shared++
        }
    }
    return [pscustomobject]@{ Users = $users; Shared = $shared }
}

function Get-PraObjectAssessment {
    <# Cloud state + readiness findings of one snapshot object (Check rules). #>
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][object]$Tenant, [object[]]$Permissions = @(), [object[]]$Members = @(), [scriptblock]$Lookup)
    $state = Get-Pra2CloudState -Context $context -Record $Record
    $findings = @(Get-Pra2Readiness -Config $context.Config -Record $Record -State $state -Skus @($Tenant.Skus))
    if ($Record.kind -ne 'User' -and $Lookup) { $findings += @(Get-Pra2TrusteeFinding -Record $Record -Permissions $Permissions -Members $Members -Lookup $Lookup) }
    return [pscustomobject]@{ State = $state; Findings = $findings; Errors = @($findings | Where-Object Level -eq 'Error'); Warnings = @($findings | Where-Object Level -eq 'Warn') }
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

function Invoke-PraIdentityStart {
    <# Common start of a conversion: ExchangeGuid cleared, source of authority in the cloud, usageLocation set from the cloud. #>
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][string]$EntraId)
    $config = $context.Config
    $guid = [string]$Record.object_guid
    $exo = Get-Pra2ExoState -Identity $EntraId -TagAttribute $config.Retention.TagAttribute
    if ($exo.Type -eq 'MailUser' -and $exo.ExchangeGuid -ne [guid]::Empty.ToString()) {
        Set-MailUser -Identity $EntraId -ExchangeGuid ([guid]::Empty) -ErrorAction Stop
        Write-PraJournal -ObjectGuid $guid -Step 'GuidCleared' -Detail ("ExchangeGuid {0} cleared" -f $exo.ExchangeGuid) -Item @{ step = 'GuidCleared'; status = 'Running' }
    }
    if (-not (Get-Pra2SourceOfAuthority -Id $EntraId)) {
        Set-Pra2SourceOfAuthority -Id $EntraId -Cloud $true
        Write-PraJournal -ObjectGuid $guid -Step 'SoaCloud' -Detail 'isCloudManaged = true' -Item @{ step = 'SoaCloud'; status = 'Running' }
    }
    $wait = Wait-Pra2Condition -Test { (Get-Pra2ExoState -Identity $EntraId).IsDirSynced -eq $false } -TimeoutSeconds (Get-PraWaitSeconds 'SyncTimeoutMinutes') -IntervalSeconds $config.Polling.IntervalSeconds
    if (-not $wait.Ok) { throw 'Exchange Online still sees the object as synchronised from AD after the source of authority transfer.' }
    $licence = Get-Pra2UserLicence -UserId $EntraId
    $country = if ($licence.usageLocation) { [string]$licence.usageLocation } else { [string]$config.Cloud.DefaultUsageLocation }
    if (-not $country) { throw 'usageLocation is empty and Cloud.DefaultUsageLocation is not set (a licence needs a country).' }
    # Set from the cloud even when present: the value synchronised from AD is refused right after the transfer (test SH2).
    Set-Pra2UsageLocation -UserId $EntraId -Country $country
    Write-PraJournal -ObjectGuid $guid -Step 'UsageLocation' -Detail "usageLocation = $country (cloud), Exchange Online no longer synchronised after $($wait.Seconds) s" -Item @{ step = 'UsageLocation'; status = 'Running' }
    return Get-Pra2UserLicence -UserId $EntraId
}

function Invoke-PraConvertUserStart {
    <# User: identity start, then the Exchange plan (licence group, Kiosk or direct). #>
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][string]$EntraId, [Parameter(Mandatory)][object]$Tenant)
    $config = $context.Config
    $guid = [string]$Record.object_guid
    $licence = Invoke-PraIdentityStart -Record $Record -EntraId $EntraId
    $users = $config.Licensing.Users
    switch ($users.Mode) {
        'Group' {
            if (-not (Test-Pra2GroupMember -GroupId $users.GroupId -UserId $EntraId)) { Set-Pra2GroupMember -GroupId $users.GroupId -UserId $EntraId -Operation Add }
            $detail = "member of the licence group $($users.GroupId)"
        }
        'Kiosk' {
            $kiosk = Get-Pra2KioskAssignment -Licence $licence -Skus @($Tenant.Skus)
            if (-not $kiosk) { throw 'Kiosk mode: no licence of this user contains EXCHANGE_S_DESKLESS.' }
            # Kiosk enabled in the existing direct assignment (other plans unchanged), or a direct assignment next to the group one.
            $how = Enable-Pra2LicencePlan -UserId $EntraId -SkuId $kiosk.SkuId -PlanIds @($kiosk.KioskPlanId) -DisabledPlans $kiosk.DisabledPlans
            $detail = "$($kiosk.SkuName): Exchange Kiosk $($how.ToLowerInvariant())"
        }
        'Direct' {
            $sku = $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$users.SkuPartNumber } | Select-Object -First 1
            if (-not $sku) { throw "Licensing.Users.SkuPartNumber $($users.SkuPartNumber) not found in the tenant." }
            $how = Enable-Pra2LicencePlan -UserId $EntraId -SkuId $sku.skuId -PlanIds @(Get-Pra2MailboxPlanId -Skus @($sku)) -DisabledPlans (Get-Pra2DisabledPlanForExchangeOnly -Sku $sku)
            $detail = "$($sku.skuPartNumber): Exchange plan $($how.ToLowerInvariant())"
        }
    }
    Write-PraJournal -ObjectGuid $guid -Step 'PlanAssigned' -Detail $detail -Item @{ step = 'PlanAssigned'; status = 'Running' }
}

function Invoke-PraConvertShared {
    <# Shared mailbox, one at a time: temporary licence, shared, tag, permissions, licence put back as it was. #>
    param([Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][string]$EntraId, [Parameter(Mandatory)][object]$Tenant,
        [object[]]$Permissions = @(), [object[]]$Members = @(), [AllowEmptyCollection()][object[]]$Original = @())
    $config = $context.Config
    $guid = [string]$Record.object_guid
    $interval = $config.Polling.IntervalSeconds
    $sku = $Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
    if (-not $sku) { throw "Licensing.Shared.SkuPartNumber $($config.Licensing.Shared.SkuPartNumber) not found in the tenant." }
    $exo = Get-Pra2ExoState -Identity $EntraId -TagAttribute $config.Retention.TagAttribute
    if ($exo.Type -ne 'SharedMailbox') {
        $null = Invoke-PraIdentityStart -Record $Record -EntraId $EntraId
        $how = Enable-Pra2LicencePlan -UserId $EntraId -SkuId $sku.skuId -PlanIds @(Get-Pra2MailboxPlanId -Skus @($sku)) -DisabledPlans (Get-Pra2DisabledPlanForExchangeOnly -Sku $sku)
        if ($how -ne 'Present') { Write-PraJournal -ObjectGuid $guid -Step 'LicenceAssigned' -Detail "temporary $($sku.skuPartNumber) (Exchange plan only, $($how.ToLowerInvariant()))" -Item @{ step = 'LicenceAssigned'; status = 'Running' } }
        $wait = Wait-Pra2Condition -Test { (Get-Pra2ExoState -Identity $EntraId).Type -in @('UserMailbox','SharedMailbox') } -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $interval
        if (-not $wait.Ok) { throw 'No cloud mailbox after the temporary licence (the licence is kept: run Convert -Batch again later).' }
        Write-PraJournal -ObjectGuid $guid -Step 'MailboxCreated' -Detail "cloud mailbox after $($wait.Seconds) s"
        # The licence must not be removed before the conversion is effective: the mailbox would be disabled at once (test SH2).
        $wait = Wait-Pra2Condition -Test {
            if ((Get-Pra2ExoState -Identity $EntraId).Type -eq 'SharedMailbox') { return $true }
            try { Set-Mailbox -Identity $EntraId -Type Shared -ErrorAction Stop } catch { $null = $_ }
            return $false
        } -TimeoutSeconds 600 -IntervalSeconds 15
        if (-not $wait.Ok) { throw 'The mailbox is not a SharedMailbox after 10 min (the licence is kept).' }
        Write-PraJournal -ObjectGuid $guid -Step 'Shared' -Detail "SharedMailbox after $($wait.Seconds) s" -Item @{ step = 'Shared'; status = 'Running' }
    }
    $tagAttribute = $config.Retention.TagAttribute
    if ((Get-Pra2ExoState -Identity $EntraId -TagAttribute $tagAttribute).Tag -ne $config.Retention.TagValue) {
        $tagParameters = @{ Identity = $EntraId; ErrorAction = 'Stop' }; $tagParameters[$tagAttribute] = $config.Retention.TagValue
        Set-Mailbox @tagParameters
    }
    Write-PraJournal -ObjectGuid $guid -Step 'Tagged' -Detail "$tagAttribute = $($config.Retention.TagValue)" -Item @{ step = 'Tagged'; status = 'Running' }
    $grants = @(Get-Pra2SharedGrant -Record $Record -Permissions $Permissions -Members $Members)
    $done = 0; $failed = [System.Collections.Generic.List[string]]::new()
    foreach ($grant in $grants) {
        try { $outcome = Grant-Pra2SharedPermission -Mailbox $EntraId -Grant $grant; $done++; Write-PraJournal -ObjectGuid $guid -Step 'Permission' -Detail ('{0} {1} ({2}, {3})' -f $grant.Right, $grant.Trustee, $outcome, $grant.Source) }
        catch { [void]$failed.Add(('{0} {1}: {2}' -f $grant.Right, $grant.Trustee, ($_.Exception.Message -replace '\s+', ' '))); Write-PraJournal -ObjectGuid $guid -Step 'Permission' -Outcome Fail -Detail $failed[$failed.Count - 1] }
    }
    # The temporary licence goes back to what the object had before Convert (removed, or its original plans).
    $how = Restore-Pra2Licence -UserId $EntraId -SkuId $sku.skuId -Original $Original
    if ($how -in @('Removed', 'Restored')) {
        Write-PraJournal -ObjectGuid $guid -Step 'LicenceRemoved' -Detail "temporary $($sku.skuPartNumber) $($how.ToLowerInvariant())" -Item @{ step = 'LicenceRemoved'; status = 'Running' }
        Start-Sleep -Seconds 60
    }
    $final = Get-Pra2ExoState -Identity $EntraId -TagAttribute $tagAttribute
    if ($final.Type -ne 'SharedMailbox') { throw "After the licence removal the object is $($final.Type), not a SharedMailbox." }
    return [pscustomobject]@{ CloudGuid = $final.ExchangeGuid; Granted = $done; Total = $grants.Count; Failed = $failed.ToArray() }
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
    if ($resume) { Write-PraItem -Context $context -Status Info -Icon Batch -Text ('Resuming batch {0}: {1} object(s) not finished' -f $resume.id, $records.Count) }

    Write-PraStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-Pra2Cloud -Context $context } catch { $context.CloudConnectFailed = $true; throw }
    Write-PraItem -Context $context -Status Ok -Text $how

    Write-PraStep -Context $context -Title 'Readiness' -Icon Key
    $context.CurrentOperation = 'Readiness'
    $tenant = Get-Pra2TenantFact -Context $context
    $lookupCache = @{}
    $lookup = { param($identity) $key = ([string]$identity).ToLowerInvariant(); if (-not $lookupCache.ContainsKey($key)) { $lookupCache[$key] = [bool](Get-Recipient -Identity $identity -ErrorAction SilentlyContinue) }; $lookupCache[$key] }.GetNewClosure()
    # Read-only first: the state of every object, then the licence units really needed, then the decision.
    $assessed = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        if (Test-PraStop) { throw 'Stopped by the operator during the readiness check: nothing was changed.' }
        $context.CurrentIdentity = if ($record.user_principal_name) { [string]$record.user_principal_name } else { [string]$record.primary_smtp_address }
        $assessment = Get-PraObjectAssessment -Record $record -Tenant $tenant -Permissions $permissions -Members $members -Lookup $lookup
        $step = if ($resume -and $resumeItems.ContainsKey([string]$record.object_guid)) { [string](Get-PraValue $resumeItems[[string]$record.object_guid] 'step' '') } else { '' }
        [void]$assessed.Add([pscustomobject]@{ Record = $record; Assessment = $assessment; State = $assessment.State; Step = $step })
    }
    $context.CurrentIdentity = ''
    $need = Get-PraLicenceNeed -Entries @($assessed) -Tenant $tenant
    $tenantFindings = @(Test-Pra2TenantReadiness -Config $config -Roles @(Get-Pra2GraphRole) -Tenant $tenant -UserCount $need.Users -SharedCount $need.Shared)
    foreach ($finding in $tenantFindings | Where-Object { $_.Level -eq 'Error' }) { Write-PraLog -Context $context -Message $finding.Message -Level Error }
    if (@($tenantFindings | Where-Object Level -eq 'Error').Count) { throw 'Tenant prerequisites not met (see above): nothing was changed.' }
    Write-PraItem -Context $context -Status Ok -Text ('Tenant prerequisites met (permissions, licences: {0} new unit(s) for users, {1} for the temporary licence)' -f $need.Users, [Math]::Min(1, $need.Shared))
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
        [void]$plan.Add([pscustomobject]@{ Record = $record; Row = $row; EntraId = [string]$assessment.State.User.id; State = $assessment.State })
    }
    $context.CurrentIdentity = ''
    $planUsers = @($plan | Where-Object { $_.Record.kind -eq 'User' })
    $planShared = @($plan | Where-Object { $_.Record.kind -ne 'User' })
    Write-PraItem -Context $context -Status Info -Icon Target -Text ('{0} object(s) ready: {1} user(s), {2} shared {3} {4} blocked' -f $plan.Count, $planUsers.Count, $planShared.Count, $dot, ($records.Count - $plan.Count))

    Write-PraStep -Context $context -Title 'Plan and confirmation' -Icon Plan
    $modeText = @{ Group = "licence group $($config.Licensing.Users.GroupId)"; Kiosk = 'Exchange Kiosk plan of the user licence'; Direct = "$($config.Licensing.Users.SkuPartNumber) (Exchange only)" }[$config.Licensing.Users.Mode]
    Write-PraItem -Context $context -Status Info -Text ("Users: ExchangeGuid cleared, source of authority to the cloud, usageLocation, Exchange plan by $modeText; the Teams storage becomes the mailbox")
    if ($planShared.Count) { Write-PraItem -Context $context -Status Info -Text ("Shared mailboxes, one at a time: temporary $($config.Licensing.Shared.SkuPartNumber), shared, $($config.Retention.TagAttribute) = $($config.Retention.TagValue), permissions from the snapshot, licence removed") }
    foreach ($entry in $plan) { $entry.Row.FinalStatus = 'Planned'; Write-PraLog -Context $context -Message ('Planned: {0} ({1})' -f $entry.Row.Identity, $entry.Record.kind) -Level Sub }
    if ($context.Mode -ne 'Apply') {
        Skip-PraStep -Title 'Users: source of authority and Exchange plan' -Reason 'Preview: not run.' -Icon Write
        Skip-PraStep -Title 'Users: cloud mailboxes' -Reason 'Preview: not run.' -Icon Clock
        Skip-PraStep -Title 'Shared mailboxes' -Reason 'Preview: not run.' -Icon Mail
        $context.NextSteps = @('pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply' + $(if ($Batch) { " -Batch $Batch" } else { Get-PraIdentityArgument }))
        return
    }
    if (-not $plan.Count) { throw 'No object can be converted (see the reasons above).' }
    Confirm-PraApply ('Convert {0} object(s) to Exchange Online ({1} user(s), {2} shared)?' -f $plan.Count, $planUsers.Count, $planShared.Count)
    if (-not $resume) {
        $context.BatchId = New-Pra2Batch $context.Journal Convert @{ SnapshotId = [long]$data.Snapshot.id; Environment = $config.Environment; Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
            ToolVersion = $toolVersion; ScopeJson = (@{ Identity = $Identity; IdentityPath = $IdentityPath; Identities = @(Get-PraIdentityList).Count; Scope = $Scope; UsersMode = $config.Licensing.Users.Mode } | ConvertTo-Json -Compress) }
        foreach ($entry in $plan) {
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
    foreach ($entry in $planUsers) {
        if ($stopped -or (Test-PraStop)) { $stopped = $true; Set-PraRowStopped $entry.Row; continue }
        $context.CurrentIdentity = $entry.Row.Identity; $context.CurrentOperation = 'Convert-User'
        try { Invoke-PraConvertUserStart -Record $entry.Record -EntraId $entry.EntraId -Tenant $tenant; [void]$started.Add($entry); Write-PraItem -Context $context -Status Ok -Text ('{0}: source of authority in the cloud, Exchange plan assigned' -f $entry.Row.Identity) }
        catch {
            $message = $_.Exception.Message -replace '\s+', ' '
            Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'ConvertUser' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
            Set-PraRowResult -Row $entry.Row -Status Error -Detail "Convert stopped: $message"
        }
    }
    if (-not $planUsers.Count) { Write-PraItem -Context $context -Status Skip -Text 'No user in this batch.' }

    Write-PraStep -Context $context -Title 'Users: cloud mailboxes' -Icon Clock
    $context.CurrentIdentity = ''; $context.CurrentOperation = 'Wait-Mailbox'
    $waiting = @($started)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    while ($waiting.Count -and $clock.Elapsed.TotalSeconds -lt (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -and -not (Test-PraStop)) {
        $still = @()
        foreach ($entry in $waiting) {
            $exo = Get-Pra2ExoState -Identity $entry.EntraId
            if ($exo.Type -eq 'UserMailbox') {
                Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'MailboxReady' -Detail ("UserMailbox {0} after {1}" -f $exo.ExchangeGuid, (Format-PraDuration $clock.Elapsed.TotalSeconds)) -Item @{ step = 'MailboxReady'; status = 'Done'; result_json = (@{ cloudGuid = $exo.ExchangeGuid } | ConvertTo-Json -Compress) }
                Set-PraRowResult -Row $entry.Row -Status Success -Detail ('cloud mailbox {0} (Teams storage promoted when present)' -f $exo.ExchangeGuid) -Field 'ExchangeOnline' -Value 'UserMailbox'
            } else { $still += $entry }
        }
        $waiting = $still
        if ($waiting.Count) { Start-Sleep -Seconds $config.Polling.IntervalSeconds }
    }
    foreach ($entry in $waiting) { Set-PraRowResult -Row $entry.Row -Status Pending -Detail ('no cloud mailbox yet after {0}: run Convert -Batch {1} again later' -f (Format-PraDuration $clock.Elapsed.TotalSeconds), $context.BatchId) }
    if (-not $started.Count) { Write-PraItem -Context $context -Status Skip -Text 'Nothing to wait for.' }

    Write-PraStep -Context $context -Title 'Shared mailboxes' -Icon Mail
    $sharedSku = $tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
    foreach ($entry in $planShared) {
        if ($stopped -or (Test-PraStop)) { $stopped = $true; Set-PraRowStopped $entry.Row; continue }
        $context.CurrentIdentity = $entry.Row.Identity; $context.CurrentOperation = 'Convert-Shared'
        $objectGuid = [string]$entry.Record.object_guid
        try {
            # A shared mailbox that failed keeps the temporary licence: never start the next one without a free unit
            # (it would stop half-way, with its GUID cleared and its source of authority in the cloud). An object that is
            # already a SharedMailbox needs no unit (resume of its permissions).
            if ($sharedSku -and (Get-Pra2ExoState -Identity $entry.EntraId).Type -ne 'SharedMailbox' -and -not @(Get-Pra2DirectLicence (Get-Pra2UserLicence -UserId $entry.EntraId) | Where-Object { $_.skuId -eq [string]$sharedSku.skuId }).Count) {
                $now = @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value) | Where-Object { [string]$_.skuId -eq [string]$sharedSku.skuId } | Select-Object -First 1
                if (-not $now -or (Get-Pra2FreeUnit $now) -lt 1) {
                    $message = "not started: no free unit of $($sharedSku.skuPartNumber) (a shared mailbox that failed keeps its temporary licence until it is finished); nothing was changed on this object. Free a unit, then run Convert -Batch $($context.BatchId) again."
                    Write-PraJournal -ObjectGuid $objectGuid -Step 'NoFreeUnit' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
                    Set-PraRowResult -Row $entry.Row -Status Error -Detail $message
                    continue
                }
            }
            $result = Invoke-PraConvertShared -Record $entry.Record -EntraId $entry.EntraId -Tenant $tenant -Permissions $permissions -Members $members -Original $originals[$objectGuid]
            $permissionText = '{0}/{1} permission(s) granted' -f $result.Granted, $result.Total
            $status = if ($result.Failed.Count) { 'Pending' } else { 'Success' }
            Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'Done' -Detail $permissionText -Item @{ step = 'Done'; status = $(if ($result.Failed.Count) { 'Running' } else { 'Done' }); result_json = (@{ cloudGuid = $result.CloudGuid; failed = $result.Failed } | ConvertTo-Json -Compress) }
            $entry.Row.Permissions = $permissionText
            Set-PraRowResult -Row $entry.Row -Status $status -Detail $(if ($result.Failed.Count) { 'shared mailbox ready; permissions not granted: ' + ($result.Failed -join '; ') } else { "shared mailbox $($result.CloudGuid), tagged, $permissionText" }) -Field 'ExchangeOnline' -Value 'SharedMailbox'
        } catch {
            $message = $_.Exception.Message -replace '\s+', ' '
            Write-PraJournal -ObjectGuid ([string]$entry.Record.object_guid) -Step 'ConvertShared' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
            Set-PraRowResult -Row $entry.Row -Status Error -Detail "Convert stopped: $message"
        }
    }
    if (-not $planShared.Count) { Write-PraItem -Context $context -Status Skip -Text 'No shared mailbox in this batch.' }
    $context.CurrentIdentity = ''
    $items = @(Get-Pra2BatchItem $context.Journal $context.BatchId)
    $counts = @{ done = @($items | Where-Object status -eq 'Done').Count; failed = @($items | Where-Object status -eq 'Failed').Count; running = @($items | Where-Object status -in @('Running','Planned')).Count }
    Set-Pra2Batch $context.Journal $context.BatchId $(if ($counts.done -eq $items.Count) { 'Complete' } else { 'Partial' }) $counts
    $context.SnapshotLabel = $context.SnapshotLabel
    $context.NextSteps = @(
        $(if ($counts.done -ne $items.Count) { "pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Batch $($context.BatchId)   # resume the objects not finished" }),
        "When AD, Exchange and Entra Connect are rebuilt: pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch $($context.BatchId)") | Where-Object { $_ }
}

function Complete-PraUserRecover {
    <# Last step of a user (or of a half-converted shared mailbox) once back on-premises: case hold added if missing, then checked. #>
    param([Parameter(Mandatory)][object]$Entry, [Parameter(Mandatory)][string]$HoldTag, [Parameter(Mandatory)][string]$Done)
    $config = $context.Config
    $guid = [string]$Entry.Item.object_guid
    # By object ID: the policy keeps it, and it stays right when the UPN and the SMTP address differ.
    if (@((Get-Pra2ExoState -Identity $Entry.EntraId).Holds) -notcontains $HoldTag) { Set-Pra2HoldLocation -Policy $config.Retention.HoldPolicy -Identity $Entry.EntraId -Operation Add }
    $wait = Wait-Pra2Condition -Test { @((Get-Pra2ExoState -Identity $Entry.EntraId).Holds) -contains $HoldTag } -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $config.Polling.IntervalSeconds
    $Entry.Row.ExchangeOnline = 'MailUser'
    if ($wait.Ok) {
        Write-PraJournal -ObjectGuid $guid -Step 'Hold' -Detail "case hold stamped after $($wait.Seconds) s" -Item @{ step = 'Hold'; status = 'Done' }
        Set-PraRowResult -Row $Entry.Row -Status Success -Detail $Done -Field 'Holds' -Value $HoldTag
    } else {
        Write-PraJournal -ObjectGuid $guid -Step 'Hold' -Outcome Info -Detail 'case hold added, not stamped yet' -Item @{ step = 'HoldPending' }
        Set-PraRowResult -Row $Entry.Row -Status Pending -Detail "back on-premises; case hold not stamped yet: run Recover -Batch $Batch again to check it (the hold is never removed from an object that is back)"
    }
}

function Invoke-PraRecover {
    $config = $context.Config
    if (-not $Batch) { throw '-Batch is required: the ID of the Convert batch to roll back (printed at the end of Convert).' }
    if (-not $config.Retention.HoldPolicy) { throw 'Retention.HoldPolicy is empty: Recover puts the converted mailboxes on that eDiscovery case hold.' }
    $context.Journal = Open-Pra2Store -Path $config.Store.JournalPath -Root $context.Root -Kind Journal
    $convert = Get-Pra2Batch $context.Journal -Id $Batch
    if (-not $convert -or $convert.action -ne 'Convert') { throw "Convert batch $Batch not found in $($config.Store.JournalPath)." }
    $existing = @(Invoke-Pra2Sql $context.Journal "SELECT * FROM batch WHERE action = 'Recover' AND convert_batch = @c ORDER BY created_utc DESC" @{ c = $convert.id } -As Rows)
    $resume = $existing | Where-Object { $_.status -in @('Running','Partial') } | Select-Object -First 1
    $finished = $existing | Where-Object { $_.status -eq 'Complete' } | Select-Object -First 1
    if (-not $resume -and $finished) { throw "Convert batch $($convert.id) was already rolled back by Recover batch $($finished.id) (complete): nothing to do. Check shows the current state of the objects." }
    $data = Read-PraSnapshotData -Id ([long]$convert.snapshot_id) -Verb 'roll back' -All -NoCount
    $byGuid = @{}; foreach ($record in $data.Records) { $byGuid[[string]$record.object_guid] = $record }
    $convertItems = @(Get-Pra2BatchItem $context.Journal $convert.id | Where-Object { $_.status -ne 'Skipped' -and $_.entra_id -and $_.step })
    $done = @{}; $previous = @{}
    if ($resume) { foreach ($item in @(Get-Pra2BatchItem $context.Journal $resume.id)) { $previous[[string]$item.object_guid] = $item; if ($item.status -eq 'Done') { $done[[string]$item.object_guid] = $true } } }
    $targets = @($convertItems | Where-Object { -not $done.ContainsKey([string]$_.object_guid) })
    if (@(Get-PraIdentityList).Count) {
        $targets = @(Select-PraIdentity -Items $targets -Where "the objects of batch $($convert.id) still to roll back" -Keys {
                param($t) $r = $byGuid[[string]$t.object_guid]
                @([string]$t.identity, [string]$t.object_guid) + $(if ($r) { @([string]$r.primary_smtp_address, [string]$r.user_principal_name, [string]$r.sam_account_name) } else { @() }) })
    }
    if ($Scope -eq 'UsersOnly') { $targets = @($targets | Where-Object { $_.kind -eq 'User' }) } elseif ($Scope -eq 'SharedOnly') { $targets = @($targets | Where-Object { $_.kind -ne 'User' }) }
    Write-PraItem -Context $context -Status Info -Icon Batch -Text ('Convert batch {0} of {1}: {2} object(s) to roll back{3}' -f $convert.id, $convert.created_utc, $targets.Count, $(if ($resume) { " (resuming Recover $($resume.id))" } else { '' }))

    Write-PraStep -Context $context -Title 'Microsoft 365 connection' -Icon Cloud
    $context.CurrentOperation = 'Connect-Cloud'
    try { $how = Connect-Pra2Cloud -Context $context; Connect-Pra2Compliance -Context $context } catch { $context.CloudConnectFailed = $true; throw }
    Write-PraItem -Context $context -Status Ok -Text ($how + ', Security & Compliance (case hold)')

    Write-PraStep -Context $context -Title 'Prerequisites' -Icon Key
    $context.CurrentOperation = 'Prerequisites'
    $holdTag = Get-Pra2HoldTag -Policy $config.Retention.HoldPolicy
    Write-PraItem -Context $context -Status Ok -Text ("Case hold policy {0} ({1})" -f $config.Retention.HoldPolicy, $holdTag)
    Write-PraItem -Context $context -Status Info -Text ("Entra Connect {0} or later, operations by {1}{2}" -f $config.EntraConnect.MinVersion, $config.EntraConnect.Mode, $(if ($config.EntraConnect.Server) { " on $($config.EntraConnect.Server)" } else { '' }))
    # Each object goes the way its current state allows; a step already done is never undone (a hold is never lifted
    # from an object that is back on-premises).
    #   User          user, or a shared mailbox whose Convert stopped half-way: back to AD like a user
    #   Back          already a synchronised MailUser with the on-premises GUID: case hold only
    #   Shared        SharedMailbox: hold, identity deleted, purged, recreated by Entra Connect
    #   SharedDeleted identity already deleted by an earlier run: purge check, recreated object
    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $targets) {
        if (Test-PraStop) { throw 'Stopped by the operator while reading the state of the objects: nothing was changed.' }
        $guid = [string]$item.object_guid
        $record = $byGuid[$guid]
        $row = New-PraRow $(if ($record) { $record } else { [pscustomobject]@{ user_principal_name = $item.identity; kind = $item.kind; primary_smtp_address = ''; object_guid = $item.object_guid; exchange_guid = '' } })
        $row.Action = 'Recover'
        [void]$context.Rows.Add($row)
        if (-not $record) { Set-PraRowResult -Row $row -Status Error -Detail "not in snapshot $($convert.snapshot_id)"; continue }
        $entry = [pscustomobject]@{ Item = $item; Record = $record; Row = $row; EntraId = [string]$item.entra_id; Exo = $null; Path = ''; CloudGuid = ''; Original = @(Get-PraOriginalLicence $item) }
        $exo = Get-Pra2ExoState -Identity $entry.EntraId -TagAttribute $config.Retention.TagAttribute
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
        $other = @($exo.Holds | Where-Object { $_ -and $_ -ne $holdTag })
        if ($exo.Type -eq 'UserMailbox' -and $other.Count) { Set-PraRowResult -Row $row -Status Error -Detail ('other hold(s) on the cloud mailbox block the rollback: {0}. Remove them, then run Recover again.' -f ($other -join ', ')); continue }
        $entry.Path = 'User'
        [void]$plan.Add($entry)
    }
    $planUsers = @($plan | Where-Object { $_.Path -eq 'User' })
    $planBack = @($plan | Where-Object { $_.Path -eq 'Back' })
    $planShared = @($plan | Where-Object { $_.Path -in @('Shared', 'SharedDeleted') })
    $partial = @($planUsers | Where-Object { $_.Item.kind -ne 'User' })
    Write-PraItem -Context $context -Status Info -Icon Target -Text ('{0} user(s) and {1} shared mailbox(es) to roll back{2}{3}' -f @($planUsers + $planBack | Where-Object { $_.Item.kind -eq 'User' }).Count, ($planShared.Count + $partial.Count + @($planBack | Where-Object { $_.Item.kind -ne 'User' }).Count),
        $(if ($planBack.Count) { "; $($planBack.Count) already back on-premises (case hold only)" } else { '' }), $(if ($partial.Count) { "; $($partial.Count) shared mailbox(es) not fully converted, rolled back like users" } else { '' }))

    Write-PraStep -Context $context -Title 'Plan and confirmation' -Icon Plan
    Write-PraItem -Context $context -Status Info -Text 'Users: source of authority back to AD, delta sync, Exchange plan removed (licences as before Convert), MailUser with the on-premises GUID, cloud mailbox kept as ComponentShared under the case hold'
    if ($planShared.Count) { Write-PraItem -Context $context -Status Info -Text 'Shared mailboxes: case hold, Entra Connect scheduler paused, identity deleted (mailbox inactive), deleted permanently, recreated from AD by a delta sync' }
    foreach ($entry in $plan) { $entry.Row.FinalStatus = 'Planned'; Write-PraLog -Context $context -Message ('Planned: {0} ({1})' -f $entry.Row.Identity, $entry.Path) -Level Sub }
    if ($context.Mode -ne 'Apply') {
        foreach ($title in @('Users: source of authority back to AD', 'Users: back on-premises', 'Shared mailboxes', 'Licence group')) { Skip-PraStep -Title $title -Reason 'Preview: not run.' -Icon Write }
        $context.NextSteps = @("pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch $Batch" + (Get-PraIdentityArgument))
        return
    }
    if (-not $plan.Count) { throw 'Nothing to roll back (see above).' }
    Confirm-PraApply ('Roll back {0} object(s) of Convert batch {1}?' -f $plan.Count, $convert.id)
    if ($resume) { $context.BatchId = [string]$resume.id }
    else {
        $context.BatchId = New-Pra2Batch $context.Journal Recover @{ ConvertBatch = $convert.id; SnapshotId = [long]$convert.snapshot_id; Environment = $config.Environment
            Account = [Security.Principal.WindowsIdentity]::GetCurrent().Name; ToolVersion = $toolVersion; ScopeJson = (@{ Identity = $Identity; IdentityPath = $IdentityPath; Identities = @(Get-PraIdentityList).Count; Scope = $Scope } | ConvertTo-Json -Compress) }
    }
    foreach ($entry in $plan) { Set-Pra2BatchItem $context.Journal $context.BatchId ([string]$entry.Item.object_guid) @{ kind = [string]$entry.Item.kind; identity = $entry.Row.Identity; entra_id = $entry.EntraId; status = 'Running' } }
    Write-PraItem -Context $context -Status Ok -Icon Batch -Text ("Recover batch {0} in {1}" -f $context.BatchId, $config.Store.JournalPath)
    $interval = $config.Polling.IntervalSeconds
    $tenantSkus = $null

    Write-PraStep -Context $context -Title 'Users: source of authority back to AD' -Icon Sync
    $context.CurrentOperation = 'Users-SOA'
    $synced = [System.Collections.Generic.List[object]]::new()
    $stopped = $false
    foreach ($entry in $planUsers) {
        if ($stopped -or (Test-PraStop)) { $stopped = $true; Set-PraRowStopped $entry.Row; continue }
        $context.CurrentIdentity = $entry.Row.Identity
        try {
            # The case hold is lifted only while the object is still a cloud mailbox (it blocks the switch back).
            if ($entry.Exo.Type -eq 'UserMailbox' -and @($entry.Exo.Holds) -contains $holdTag) { Set-Pra2HoldLocation -Policy $config.Retention.HoldPolicy -Identity $entry.EntraId -Operation Remove; Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'HoldLifted' -Detail 'case hold lifted for the rollback' }
            if (Get-Pra2SourceOfAuthority -Id $entry.EntraId) { Set-Pra2SourceOfAuthority -Id $entry.EntraId -Cloud $false; Write-PraJournal -ObjectGuid ([string]$entry.Item.object_guid) -Step 'SoaAd' -Detail 'isCloudManaged = false' -Item @{ step = 'SoaAd' } }
            [void]$synced.Add($entry)
        } catch { Set-PraRowResult -Row $entry.Row -Status Error -Detail ('rollback stopped: ' + ($_.Exception.Message -replace '\s+', ' ')) }
    }
    $context.CurrentIdentity = ''
    if ($synced.Count) {
        $context.CurrentOperation = 'EntraConnect-Delta'
        # Resume: no cycle when every user is already synchronised from AD again.
        if (@($synced | Where-Object { (Get-Pra2ExoState -Identity $_.EntraId).IsDirSynced -ne $true }).Count) {
            $state = Invoke-Pra2EntraConnect -Context $context -Operation Delta -Caller $caller
            Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect delta cycle: $state"
        } else { Write-PraItem -Context $context -Status Info -Icon Sync -Text 'Already synchronised from AD: no delta cycle needed.' }
        foreach ($entry in @($synced)) {
            $wait = Wait-Pra2Condition -Test { (Get-Pra2ExoState -Identity $entry.EntraId).IsDirSynced -eq $true } -TimeoutSeconds (Get-PraWaitSeconds 'SyncTimeoutMinutes') -IntervalSeconds $interval
            if ($wait.Ok) { Write-PraItem -Context $context -Status Ok -Text ('{0}: synchronised from AD again' -f $entry.Row.Identity) }
            else { Set-PraRowResult -Row $entry.Row -Status Pending -Detail 'not synchronised from AD yet: run Entra Connect, then Recover -Batch again'; [void]$synced.Remove($entry) }
        }
    } elseif (-not $planUsers.Count) { Write-PraItem -Context $context -Status Skip -Text 'No user to switch back.' }

    Write-PraStep -Context $context -Title 'Users: back on-premises' -Icon Write
    $context.CurrentOperation = 'Users-Plan'
    $usersMode = $config.Licensing.Users.Mode
    foreach ($entry in @($synced) + @($planBack)) {
        # The users switched back to AD are always finished; an object already back (case hold only) is a new start.
        if ($entry.Path -eq 'Back' -and ($stopped -or (Test-PraStop))) { $stopped = $true; Set-PraRowStopped $entry.Row; continue }
        $context.CurrentIdentity = $entry.Row.Identity
        $guid = [string]$entry.Item.object_guid
        $isUser = $entry.Item.kind -eq 'User'
        try {
            if ($entry.Path -eq 'User') {
                if ($entry.Exo.Type -eq 'UserMailbox') {
                    $wait = Wait-Pra2Condition -Test { @((Get-Pra2ExoState -Identity $entry.EntraId).Holds) -notcontains $holdTag } -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $interval
                    if (-not $wait.Ok) { throw 'the case hold is still stamped on the cloud mailbox (it would block the rollback)' }
                }
                if (-not $tenantSkus) { $tenantSkus = @((Get-Pra2TenantFact -Context $context).Skus) }
                $removed = if (-not $isUser) {
                    # Shared mailbox not fully converted: its temporary licence goes back to what it was before Convert.
                    $sku = $tenantSkus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Shared.SkuPartNumber } | Select-Object -First 1
                    if ($sku) { 'temporary licence ' + (Restore-Pra2Licence -UserId $entry.EntraId -SkuId $sku.skuId -Original $entry.Original).ToLowerInvariant() } else { 'no temporary licence' }
                } else {
                    switch ($usersMode) {
                        'Group' { Set-Pra2GroupMember -GroupId $config.Licensing.Users.GroupId -UserId $entry.EntraId -Operation Remove; 'removed from the licence group' }
                        'Kiosk' {
                            $kiosk = Get-Pra2KioskAssignment -Licence (Get-Pra2UserLicence -UserId $entry.EntraId) -Skus $tenantSkus
                            if ($kiosk) { "$($kiosk.SkuName) " + (Restore-Pra2Licence -UserId $entry.EntraId -SkuId $kiosk.SkuId -Original $entry.Original).ToLowerInvariant() } else { 'no Kiosk licence' }
                        }
                        'Direct' {
                            $sku = $tenantSkus | Where-Object { [string]$_.skuPartNumber -eq [string]$config.Licensing.Users.SkuPartNumber } | Select-Object -First 1
                            if ($sku) { "$($sku.skuPartNumber) " + (Restore-Pra2Licence -UserId $entry.EntraId -SkuId $sku.skuId -Original $entry.Original).ToLowerInvariant() } else { 'no direct licence' }
                        }
                    }
                }
                Write-PraJournal -ObjectGuid $guid -Step 'PlanRemoved' -Detail "Exchange plan removed: $removed" -Item @{ step = 'PlanRemoved' }
                $onPrem = [string]$entry.Record.exchange_guid
                $wait = Wait-Pra2Condition -Test { $s = Get-Pra2ExoState -Identity $entry.EntraId; $s.Type -eq 'MailUser' -and $s.ExchangeGuid -eq $onPrem } -TimeoutSeconds (Get-PraWaitSeconds 'MailboxTimeoutMinutes') -IntervalSeconds $interval
                if (-not $wait.Ok) { throw "not a MailUser with the on-premises ExchangeGuid $onPrem yet" }
                Write-PraJournal -ObjectGuid $guid -Step 'MailUser' -Detail "MailUser with the on-premises GUID after $($wait.Seconds) s" -Item @{ step = 'MailUser' }
            }
            $text = if ($isUser) { 'back on-premises (MailUser, on-premises GUID); the cloud mailbox stays as ComponentShared under the case hold' }
            else { 'shared mailbox not fully converted, rolled back like a user (MailUser, on-premises GUID); its cloud data, if any, stays under the case hold' }
            Complete-PraUserRecover -Entry $entry -HoldTag $holdTag -Done $text
        } catch {
            $message = $_.Exception.Message -replace '\s+', ' '
            Write-PraJournal -ObjectGuid $guid -Step 'RecoverUser' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
            Set-PraRowResult -Row $entry.Row -Status Error -Detail "rollback stopped: $message"
        }
    }
    $context.CurrentIdentity = ''
    if (-not $synced.Count -and -not $planBack.Count) { Write-PraItem -Context $context -Status Skip -Text 'No user to bring back.' }

    Write-PraStep -Context $context -Title 'Shared mailboxes' -Icon Mail
    $context.CurrentOperation = 'Recover-Shared'
    $deleted = [System.Collections.Generic.List[object]]::new()
    $paused = $false
    try {
        foreach ($entry in $planShared) {
            if ($stopped -or (Test-PraStop)) { $stopped = $true; Set-PraRowStopped $entry.Row; continue }
            $context.CurrentIdentity = $entry.Row.Identity
            $guid = [string]$entry.Item.object_guid
            try {
                if ($entry.Path -eq 'Shared') {
                    Set-Pra2HoldLocation -Policy $config.Retention.HoldPolicy -Identity $entry.EntraId -Operation Add
                    $wait = Wait-Pra2Condition -Test { @((Get-Pra2ExoState -Identity $entry.EntraId).Holds) -contains $holdTag } -TimeoutSeconds (Get-PraWaitSeconds 'HoldTimeoutMinutes') -IntervalSeconds $interval
                    if (-not $wait.Ok) { throw 'the case hold is not stamped on the shared mailbox: identity NOT deleted (it would not become inactive)' }
                    Write-PraJournal -ObjectGuid $guid -Step 'Hold' -Detail "case hold stamped after $($wait.Seconds) s" -Item @{ step = 'Hold' }
                    if (-not $paused) {
                        $state = Invoke-Pra2EntraConnect -Context $context -Operation Pause -Caller $caller
                        $paused = $true
                        Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler paused: $state"
                    }
                    $entry.CloudGuid = $entry.Exo.ExchangeGuid
                    Remove-Pra2Identity -UserId $entry.EntraId -Operation Delete
                    $cloudGuid = $entry.CloudGuid
                    $wait = Wait-Pra2Condition -Test { [bool](Get-Mailbox -InactiveMailboxOnly -Identity $cloudGuid -ErrorAction SilentlyContinue) } -TimeoutSeconds 900 -IntervalSeconds 15
                    if (-not $wait.Ok) { throw "identity deleted but mailbox $cloudGuid not inactive after 15 min: check it before anything else (Entra ID recycle bin, 30 days)" }
                    Write-PraJournal -ObjectGuid $guid -Step 'Inactive' -Detail "identity deleted, mailbox $cloudGuid inactive after $($wait.Seconds) s" -Item @{ step = 'Inactive'; result_json = (@{ inactiveGuid = $cloudGuid } | ConvertTo-Json -Compress) }
                } else { Write-PraItem -Context $context -Status Info -Text ('{0}: identity deleted by an earlier run (mailbox {1} inactive): recycle bin and recreated object checked' -f $entry.Row.Identity, $entry.CloudGuid) }
                if (Test-Pra2DeletedIdentity -UserId $entry.EntraId) {
                    if (-not $paused) {
                        $state = Invoke-Pra2EntraConnect -Context $context -Operation Pause -Caller $caller
                        $paused = $true
                        Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler paused: $state"
                    }
                    Remove-Pra2Identity -UserId $entry.EntraId -Operation Purge
                    $wait = Wait-Pra2Condition -Test { -not (Test-Pra2DeletedIdentity -UserId $entry.EntraId) } -TimeoutSeconds 300 -IntervalSeconds 10
                    if (-not $wait.Ok) { throw 'identity still in the Entra ID recycle bin: the next sync would restore it' }
                }
                Write-PraJournal -ObjectGuid $guid -Step 'Purged' -Detail 'identity deleted permanently' -Item @{ step = 'Purged' }
                [void]$deleted.Add($entry)
                Write-PraItem -Context $context -Status Ok -Text ('{0}: mailbox inactive under the case hold, identity deleted permanently' -f $entry.Row.Identity)
            } catch {
                $message = $_.Exception.Message -replace '\s+', ' '
                Write-PraJournal -ObjectGuid $guid -Step 'RecoverShared' -Outcome Fail -Detail $message -Item @{ status = 'Failed'; message = $message }
                Set-PraRowResult -Row $entry.Row -Status Error -Detail "rollback stopped: $message"
            }
        }
    } finally {
        if ($paused) {
            try { $state = Invoke-Pra2EntraConnect -Context $context -Operation Resume -Caller $caller; Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect scheduler resumed: $state" }
            catch { Write-PraLog -Context $context -Message "Entra Connect scheduler NOT resumed: $($_.Exception.Message). Resume it by hand (Set-ADSyncScheduler -SyncCycleEnabled `$true)." -Level Error }
        }
    }
    $context.CurrentIdentity = ''
    if ($deleted.Count) {
        $context.CurrentOperation = 'EntraConnect-Delta'
        $state = Invoke-Pra2EntraConnect -Context $context -Operation Delta -Caller $caller
        Write-PraItem -Context $context -Status Ok -Icon Sync -Text "Entra Connect delta cycle: $state"
        foreach ($entry in $deleted) {
            $guid = [string]$entry.Item.object_guid
            $onPrem = [string]$entry.Record.exchange_guid; $immutable = [string]$entry.Record.immutable_id
            $wait = Wait-Pra2Condition -Test {
                $filter = [uri]::EscapeDataString(("onPremisesImmutableId eq '{0}'" -f $immutable.Replace("'", "''")))
                $found = @((Invoke-Pra2Graph -Uri ("v1.0/users?`$filter={0}&`$select=id" -f $filter)).value)
                if ($found.Count -ne 1 -or [string]$found[0].id -eq $entry.EntraId) { return $false }
                $s = Get-Pra2ExoState -Identity ([string]$found[0].id)
                if ($s.Type -eq 'MailUser' -and $s.ExchangeGuid -eq $onPrem) { return [string]$found[0].id }
                return $false
            } -TimeoutSeconds (Get-PraWaitSeconds 'SyncTimeoutMinutes') -IntervalSeconds $interval
            if ($wait.Ok) {
                Write-PraJournal -ObjectGuid $guid -Step 'Recreated' -Detail ("new object {0} from AD, MailUser with the on-premises GUID" -f $wait.Last) -Item @{ step = 'Recreated'; status = 'Done'; entra_id = [string]$wait.Last }
                Set-PraRowResult -Row $entry.Row -Status Success -Detail ('recreated from AD ({0}); the cloud mailbox {1} is inactive under the case hold' -f $wait.Last, $entry.CloudGuid) -Field 'ExchangeOnline' -Value 'MailUser'
            } else { Set-PraRowResult -Row $entry.Row -Status Pending -Detail "identity deleted permanently, mailbox inactive; not recreated by Entra Connect yet: run a delta sync, then Recover -Batch $Batch again" }
        }
    } elseif (-not $planShared.Count) { Write-PraItem -Context $context -Status Skip -Text 'No shared mailbox in this batch.' }

    Write-PraStep -Context $context -Title 'Licence group' -Icon Key
    $context.CurrentOperation = 'Licence-Group'
    # The group serves every Convert batch of this journal: it goes back to AD only when no converted user is left,
    # whichever batch switched it (later waves record nothing, the group was already in the cloud).
    $changes = @(Invoke-Pra2Sql $context.Journal "SELECT * FROM tenant_change WHERE kind = 'GroupSoa' AND restored = 0" -As Rows)
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

    $context.StepTotal = @{ Collect = 5; Check = 4; Convert = 7; Recover = 8 }[$Action]
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
