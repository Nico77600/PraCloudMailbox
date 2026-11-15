# What the console knows without running anything: the request of a run and its command line, the events of the engine, the configuration, the snapshot and the
# journal read read-only, the batches, the waves and the lookups made in the snapshot. Pure functions (no window). Dot-sourced by PRA.Gui.psm1.

#region Comparisons and values
# Ordinal comparisons, private to the console: the culture comparison ignores zero-weight characters (a string and the same string followed by U+200B compare as equal), and every
# guard of the console (typed confirmation, batch, environment) must see them as different.
function Test-PcmGuiOrdinal {
    param([AllowNull()][object]$Left, [AllowNull()][object]$Right)
    if ($null -eq $Left -or $null -eq $Right) { return ($null -eq $Left -and $null -eq $Right) }
    return [string]::Equals([string]$Left, [string]$Right, [StringComparison]::OrdinalIgnoreCase)
}

function Test-PcmGuiOrdinalIn {
    param([AllowNull()][object]$Value, [AllowNull()][AllowEmptyCollection()][object]$Set)
    foreach ($item in @($Set)) { if (Test-PcmGuiOrdinal $Value $item) { return $true } }
    return $false
}

function Get-PcmGuiOrdinalUnique {
    <# Distinct strings, first spelling and first order kept, letter case ignored (Select-Object -Unique uses the culture comparison). #>
    param([AllowNull()][AllowEmptyCollection()][object]$Items)
    $seen = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    foreach ($item in @($Items)) { if ($null -ne $item -and $seen.Add([string]$item)) { [string]$item } }
}

function Get-PcmGuiValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
    } elseif ($Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}

function Resolve-PcmGuiOfficialValue {
    <# The official spelling of a value (letter case ignored); throws when no member of the set matches. #>
    param([string]$Value, [string[]]$Set)
    foreach ($item in $Set) { if ([string]::Equals($Value, $item, [StringComparison]::OrdinalIgnoreCase)) { return $item } }
    throw ("'{0}' is not one of the accepted values ({1})." -f $Value, ($Set -join ', '))
}

function Test-PcmGuiConfirmation {
    <# The typed confirmation is the name of the action (CONVERT, RECOVER) in any letter case, spaces around it ignored. Compared by character code: an invisible character makes it differ. #>
    param([AllowNull()][AllowEmptyString()][string]$Typed, [AllowNull()][string]$Word)
    if (-not $Word -or $null -eq $Typed) { return $false }
    return [string]::Equals($Typed.Trim(), $Word.Trim(), [StringComparison]::OrdinalIgnoreCase)
}

function ConvertTo-PcmGuiPlural {
    param([int]$Count, [string]$One, [string]$Many)
    if ($Count -eq 1) { return ('{0} {1}' -f $Count, $One) }
    return ('{0} {1}' -f $Count, $Many)
}

function Format-PcmGuiTime {
    <# A UTC time of the store or the journal (ISO 8601) in local time, yyyy-MM-dd HH:mm; '' when empty. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)
    if (-not $Value) { return '' }
    $time = [datetime]::MinValue
    if (-not [datetime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$time)) { return $Value }
    if ($time.Kind -eq [DateTimeKind]::Unspecified) { $time = [datetime]::SpecifyKind($time, [DateTimeKind]::Utc) }
    return $time.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-PcmGuiDate {
    <# 'Today 14:30' or '11 Nov 2026 23:41' for the list of batches. #>
    param([AllowEmptyString()][AllowNull()][string]$Value)
    $time = [datetime]::MinValue
    if (-not $Value -or -not [datetime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$time)) { return 'date unknown' }
    if ($time.Kind -eq [DateTimeKind]::Unspecified) { $time = [datetime]::SpecifyKind($time, [DateTimeKind]::Utc) }
    $time = $time.ToLocalTime()
    if ($time.Date -eq [datetime]::Today) { return 'Today ' + $time.ToString('HH:mm', [Globalization.CultureInfo]::InvariantCulture) }
    return $time.ToString('d MMM yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-PcmGuiClock {
    param([double]$Seconds)
    $total = [int][Math]::Max(0, [Math]::Floor($Seconds))
    return ('{0:00}:{1:00}' -f [Math]::Floor($total / 60), ($total % 60))
}

function Test-PcmGuiSamePath {
    <# Windows paths ignore the letter case: the FULL paths are compared, not the typed text. #>
    param([AllowEmptyString()][string]$Left, [AllowEmptyString()][string]$Right)
    if (-not $Left.Trim() -or -not $Right.Trim()) { return $false }
    try { $a = [IO.Path]::GetFullPath($Left.Trim()).TrimEnd('\'); $b = [IO.Path]::GetFullPath($Right.Trim()).TrimEnd('\') } catch { return $false }
    return [string]::Equals($a, $b, [StringComparison]::OrdinalIgnoreCase)
}
#endregion

#region The request of a run and its command line
function New-PcmGuiRequest {
    <#
    .SYNOPSIS
        What one run asks of the engine, checked: the same rules as the command line. One source of objects at most (-Identity, -CsvPath, -SearchBase or -Group); -Scope and -MaxObjects go with any of them.
    #>
    param(
        [ValidateSet('Collect', 'Check', 'Convert', 'Recover', 'Schedule')][string]$Action = 'Check',
        [ValidateSet('Preview', 'Apply')][string]$Mode = 'Preview',
        [string]$Batch = '', [string]$Identity = '', [string]$SearchBase = '', [string]$CsvPath = '', [string]$Group = '',
        [ValidateSet('All', 'UsersOnly', 'SharedOnly')][string]$Scope = 'All',
        [int]$MaxObjects = 0, [long]$Snapshot = 0, [string]$Expect = '', [bool]$Once = $false, [bool]$ResolveOnly = $false
    )
    $Action = Resolve-PcmGuiOfficialValue $Action @('Collect', 'Check', 'Convert', 'Recover', 'Schedule')
    $Mode = Resolve-PcmGuiOfficialValue $Mode @('Preview', 'Apply')
    $Scope = Resolve-PcmGuiOfficialValue $Scope @('All', 'UsersOnly', 'SharedOnly')
    if ($Expect) { $Expect = Resolve-PcmGuiOfficialValue $Expect @('Provisioned', 'Deprovisioned', 'Retained') }
    foreach ($value in @($Batch, $Identity, $SearchBase, $CsvPath, $Group, $Expect)) {
        if ($value -match '[\x00-\x1f]') { throw 'A value contains a control character (a pasted line break is the usual cause): retype it.' }
    }
    if ($MaxObjects -lt 0) { throw 'The maximum number of objects must be zero (no limit) or more.' }
    $sources = @(@($Identity, $CsvPath, $SearchBase, $Group) | Where-Object { $_ })
    if ($sources.Count -gt 1) { throw 'Choose one way to pick the objects: one mailbox, an OU, a group or a CSV file, not several.' }
    if ($Action -eq 'Check' -and $Mode -eq 'Apply') { throw 'A check only reads: it never has an Apply.' }
    if ($Action -eq 'Recover' -and -not $Batch) { throw 'Pick the batch to roll back.' }
    if ($Action -eq 'Collect' -and ($Batch -or $SearchBase -or $Group -or $CsvPath)) { throw 'Collect reads Exchange: it takes no batch, OU, group or CSV file.' }
    if ($Action -eq 'Schedule' -and ($Batch -or $Identity -or $SearchBase -or $Group -or $CsvPath -or $MaxObjects -gt 0 -or $Snapshot -gt 0)) { throw 'Scheduling the daily Collect takes no mailbox, batch, OU, group, CSV file or snapshot.' }
    if ($Action -eq 'Schedule' -and ($Expect -or $Once -or $ResolveOnly)) { throw 'Scheduling the daily Collect takes no check option.' }
    if ($Expect -and ($Action -ne 'Check' -or -not $Batch)) { throw 'What the batch should look like is only asked when a check follows a batch.' }
    if ($Once -and $Action -ne 'Check') { throw 'Looking once is only for a check.' }
    if ($ResolveOnly -and ($Action -ne 'Check' -or -not $Group -or $Batch)) { throw 'Looking a group up is only for a check of a group.' }
    $request = [ordered]@{ Action = $Action; Mode = $Mode; Batch = $Batch; Identity = $Identity; SearchBase = $SearchBase; CsvPath = $CsvPath; Group = $Group
        Scope = $Scope; MaxObjects = $MaxObjects; Snapshot = $Snapshot; Expect = $Expect; Once = $Once }
    if ($ResolveOnly) { $request.ResolveOnly = $true }
    return $request
}

function ConvertTo-PcmGuiNativeArgument {
    <# CommandLineToArgvW: backslashes before a quote are doubled, then the text is put between quotes. No shell is used. #>
    param([AllowEmptyString()][string]$Value)
    if ($Value.Contains([string][char]0)) { throw 'NUL is not a valid argument.' }
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

function ConvertTo-PcmGuiShownArgument {
    <# A value as an operator would type it: bare when it is plain, between double quotes otherwise. #>
    param([AllowEmptyString()][string]$Value)
    if ($Value -match '^[A-Za-z0-9_.:@\\/-]+$') { return $Value }
    return '"' + $Value.Replace('"', '""') + '"'
}

function Get-PcmGuiEngine {
    <#
        The PowerShell that runs an action (every action, Collect included): PowerShell 7.5 or later. Returns Ok, Path, Version and the problem when there is none.
    #>
    $candidates = [Collections.Generic.List[string]]::new()
    if ($PSVersionTable.PSEdition -eq 'Core') { $candidates.Add((Join-Path $PSHOME 'pwsh.exe')) }
    $command = Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { $candidates.Add($command.Source) }
    $candidates.Add((Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))
    $best = $null
    foreach ($candidate in (Get-PcmGuiOrdinalUnique @($candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) }))) {
        $text = [regex]::Match([string](Get-Item -LiteralPath $candidate).VersionInfo.ProductVersion, '^\d+(\.\d+){1,3}').Value
        if (-not $text) { continue }
        if ([version]$text -ge [version]'7.5') { return [pscustomobject]@{ Ok = $true; Path = $candidate; Version = $text; Problem = '' } }
        if (-not $best) { $best = "$candidate ($text)" }
    }
    $problem = 'PowerShell 7.5 or later (pwsh.exe) not found: every action runs in PowerShell 7.5 or later (winget install Microsoft.PowerShell).'
    if ($best) { $problem = "PowerShell 7.5 or later needed, found ${best}: every action runs in PowerShell 7.5 or later." }
    return [pscustomobject]@{ Ok = $false; Path = ''; Version = ''; Problem = $problem }
}

function Get-PcmGuiCommand {
    <#
    .SYNOPSIS
        The command line of one run: the same script and parameters as by hand. Apply gets -Force: the console asked its own typed confirmation. Never -NonInteractive: an interactive
        sign-in must be allowed to open its window. Returns FilePath, Tokens, Arguments, Display (as an operator would type it) and NeedsConsoleWindow.
    .PARAMETER BaseConfigPath
        The original configuration file, when -ConfigPath is the copy generated for this run (settings changed in the window).
    #>
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)][Collections.IDictionary]$Request, [string]$Engine = '', [string]$BaseConfigPath = '', [string]$ScriptPath = '')
    if (-not $ScriptPath) { $ScriptPath = Join-Path $Root 'Invoke-PraCloudMailbox.ps1' }
    $action = [string]$Request.Action
    if (-not $Engine) { $Engine = (Get-PcmGuiEngine).Path }
    $mode = if ($Request.Contains('Mode') -and $Request.Mode) { [string]$Request.Mode } else { 'Preview' }
    $options = [Collections.Generic.List[string]]::new()
    $options.Add('-Action'); $options.Add($action)
    if ($action -ne 'Check') { $options.Add('-Mode'); $options.Add($mode) }
    if ($Request.Contains('Scope') -and $Request.Scope -and -not (Test-PcmGuiOrdinal $Request.Scope 'All')) { $options.Add('-Scope'); $options.Add([string]$Request.Scope) }
    foreach ($name in @('Batch', 'Identity', 'CsvPath', 'SearchBase', 'Group')) { if ($Request.Contains($name) -and $Request[$name]) { $options.Add('-' + $name); $options.Add([string]$Request[$name]) } }
    if ($Request.Contains('MaxObjects') -and [int]$Request.MaxObjects -gt 0) { $options.Add('-MaxObjects'); $options.Add([string][int]$Request.MaxObjects) }
    if ($Request.Contains('Snapshot') -and [long]$Request.Snapshot -gt 0) { $options.Add('-Snapshot'); $options.Add([string][long]$Request.Snapshot) }
    if ($Request.Contains('Expect') -and $Request.Expect) { $options.Add('-Expect'); $options.Add([string]$Request.Expect) }
    if ($Request.Contains('Once') -and [bool]$Request.Once) { $options.Add('-Once') }
    if ($Request.Contains('ResolveOnly') -and [bool]$Request.ResolveOnly) { $options.Add('-ResolveOnly') }
    $tokens = [Collections.Generic.List[string]]::new()
    foreach ($token in @('-NoProfile', '-STA', '-File', $ScriptPath)) { $tokens.Add($token) }
    foreach ($token in $options) { $tokens.Add($token) }
    $tokens.Add('-ConfigPath'); $tokens.Add($ConfigPath)
    if ($BaseConfigPath) { $tokens.Add('-BaseConfigPath'); $tokens.Add($BaseConfigPath) }
    if ($mode -eq 'Apply' -and $action -ne 'Check') { $tokens.Add('-Force') }
    # As an operator would type it in the package folder: the script as .\Invoke-PraCloudMailbox.ps1, the original configuration file only when it is not the default one.
    $shown = [Collections.Generic.List[string]]::new()
    $shown.Add('.\Invoke-PraCloudMailbox.ps1')
    foreach ($token in $options) { $shown.Add((ConvertTo-PcmGuiShownArgument $token)) }
    $shownConfig = if ($BaseConfigPath) { $BaseConfigPath } else { $ConfigPath }
    $defaultConfig = Join-Path $Root 'config\PraCloudMailbox.config.psd1'
    if (-not (Test-PcmGuiSamePath $shownConfig $defaultConfig)) { $shown.Add('-ConfigPath'); $shown.Add((ConvertTo-PcmGuiShownArgument $shownConfig)) }
    if ($mode -eq 'Apply' -and $action -ne 'Check') { $shown.Add('-Force') }
    return [pscustomobject]@{
        FilePath = $Engine; Tokens = $tokens.ToArray(); NeedsConsoleWindow = $true
        Arguments = (@($tokens | ForEach-Object { ConvertTo-PcmGuiShownTokenForProcess $_ }) -join ' ')
        Display = ($shown -join ' ')
    }
}

function ConvertTo-PcmGuiShownTokenForProcess {
    <# One token of the command line given to CreateProcess: plain tokens bare, everything else quoted for CommandLineToArgvW. #>
    param([AllowEmptyString()][string]$Value)
    if ($Value -match '^[A-Za-z0-9_.:-]+$') { return $Value }
    return ConvertTo-PcmGuiNativeArgument $Value
}
#endregion

#region Events of the engine
function Read-PcmGuiEvents {
    <# The events written since Position (complete lines only: the run may be writing the next one). Returns the events and the new position. #>
    param([Parameter(Mandatory)][string]$Path, [long]$Position = 0)
    $none = [pscustomobject]@{ Events = @(); Position = $Position; HasPartial = $false }
    if (-not (Test-Path -LiteralPath $Path)) { return $none }
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        if ($stream.Length -le $Position) { return $none }
        [void]$stream.Seek($Position, [IO.SeekOrigin]::Begin)
        $buffer = New-Object byte[] ([int]($stream.Length - $Position))
        $read = $stream.Read($buffer, 0, $buffer.Length)
    } finally { $stream.Dispose() }
    $last = -1
    for ($i = $read - 1; $i -ge 0; $i--) { if ($buffer[$i] -eq 10) { $last = $i; break } }
    $partial = $last -lt ($read - 1)
    if ($last -lt 0) { return [pscustomobject]@{ Events = @(); Position = $Position; HasPartial = $partial } }
    $text = [Text.Encoding]::UTF8.GetString($buffer, 0, $last + 1).TrimStart([char]0xFEFF)
    $events = [Collections.Generic.List[object]]::new()
    foreach ($line in $text.Split([char]10)) {
        if (-not $line.Trim()) { continue }
        try { $events.Add(($line | ConvertFrom-Json)) } catch { $events.Add([pscustomobject]@{ kind = 'item'; status = 'Info'; text = $line.Trim(); identity = '' }) }
    }
    return [pscustomobject]@{ Events = $events.ToArray(); Position = $Position + $last + 1; HasPartial = $partial }
}
#endregion

#region The configuration, the snapshot and the journal, read-only
function Read-PcmGuiState {
    <#
        What the console shows, read without changing anything: the configuration, the snapshots and the objects of the one in use, the batches of the journal with their objects. A file
        that cannot be read is reported in the state (ConfigError, SnapshotError, JournalError), never thrown.
    #>
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ConfigPath, [long]$SnapshotId = 0)
    $state = [pscustomobject]@{ Root = $Root; ConfigPath = $ConfigPath; Config = $null; ConfigError = ''; Snapshot = $null; Snapshots = @(); Objects = @()
        Permissions = 0; SnapshotError = ''; SnapshotExists = $false; Batches = @(); Items = @(); JournalError = ''; JournalExists = $false; Read = (Get-Date) }
    try { $state.Config = Import-PcmConfiguration -Path $ConfigPath -Root $Root }
    catch { $state.ConfigError = $_.Exception.Message; return $state }
    $config = $state.Config
    $state.SnapshotExists = Test-Path -LiteralPath $config.Store.Path
    if ($state.SnapshotExists) {
        try {
            $connection = Open-PcmStore -Path $config.Store.Path -Root $Root -ReadOnly
            try {
                $state.Snapshots = @(Get-PcmSnapshotList $connection)
                $state.Snapshot = if ($SnapshotId -gt 0) { Get-PcmSnapshot $connection -Id $SnapshotId } else { Get-PcmSnapshot $connection }
                if ($state.Snapshot) {
                    $id = [long]$state.Snapshot.id
                    $state.Objects = @(Invoke-PcmSql $connection 'SELECT object_guid, kind, user_principal_name, primary_smtp_address, sam_account_name, name, display_name, distinguished_name, organizational_unit FROM mailbox WHERE snapshot_id = @id ORDER BY rowid' @{ id = $id } -As Rows)
                    $state.Permissions = [int](Invoke-PcmSql $connection 'SELECT COUNT(*) FROM permission WHERE snapshot_id = @id' @{ id = $id } -As Scalar)
                }
            } finally { Close-PcmStore $connection }
        } catch { $state.SnapshotError = $_.Exception.Message }
    }
    $state.JournalExists = Test-Path -LiteralPath $config.Store.JournalPath
    if ($state.JournalExists) {
        try {
            $journal = Open-PcmStore -Path $config.Store.JournalPath -Root $Root -Kind Journal -ReadOnly
            try {
                $state.Batches = @(Invoke-PcmSql $journal 'SELECT * FROM batch ORDER BY created_utc DESC' -As Rows)
                $state.Items = @(Invoke-PcmSql $journal 'SELECT batch_id, object_guid, kind, identity, step, status, message FROM batch_item' -As Rows)
            } finally { Close-PcmStore $journal }
        } catch { $state.JournalError = $_.Exception.Message }
    }
    return $state
}

$script:GuiIntegrityCache = $null

function Get-PcmGuiFileStamp {
    <# Where each file is, how long it is and when it was last written (a file that is not there is named as missing): what tells that a file changed without reading it. #>
    param([Parameter(Mandatory)][string[]]$Path)
    return (@($Path | ForEach-Object {
                $file = [IO.FileInfo]::new($_)
                if ($file.Exists) { '{0}:{1}:{2}' -f $file.FullName, $file.Length, $file.LastWriteTimeUtc.Ticks } else { $_ + ':missing' }
            }) -join '|')
}

function Get-PcmGuiSnapshotInfo {
    <#
        What the Setup page says about the snapshot file: Path, Exists, Id, Taken, AgeHours, MaxAgeDays, Error, and Integrity (Verified, Not verifiable or Changed: the file against the
        .sha256 that Collect wrote beside it; empty when there is no file) with IntegrityDetail, the plain sentence for the operator.
    #>
    param([Parameter(Mandatory)]$State)
    if (-not $State.Config) { return $null }
    $info = [ordered]@{ Path = [string]$State.Config.Store.Path; Exists = [bool]$State.SnapshotExists; Id = 0; Taken = ''; AgeHours = 0.0; MaxAgeDays = [int]$State.Config.Store.MaxSnapshotAgeDays; Error = [string]$State.SnapshotError
        Integrity = ''; IntegrityDetail = '' }
    if ($info.Exists) {
        # The file is hashed once, then again only when its length or its time of last write (or those of its .sha256 file) change: the window asks for this at every keystroke.
        $backupFolder = [string]$State.Config.Store.BackupFolder
        $stamp = '{0}|{1}' -f $backupFolder, (Get-PcmGuiFileStamp -Path @($info.Path, ($info.Path + '.sha256')))
        if ($script:GuiIntegrityCache -and [string]::Equals($script:GuiIntegrityCache.Stamp, $stamp, [StringComparison]::OrdinalIgnoreCase)) {
            $info.Integrity = $script:GuiIntegrityCache.Integrity; $info.IntegrityDetail = $script:GuiIntegrityCache.Detail
        }
        else {
            try {
                $check = Test-PcmSnapshotHash -Path $info.Path -BackupFolder $backupFolder
                $info.Integrity = switch ($check.State) { 'Verified' { 'Verified' } 'Changed' { 'Changed' } default { 'Not verifiable' } }
                $info.IntegrityDetail = [string]$check.Message
            } catch { $info.Integrity = ''; $info.IntegrityDetail = '' }
            $script:GuiIntegrityCache = @{ Stamp = $stamp; Integrity = $info.Integrity; Detail = $info.IntegrityDetail }
        }
    }
    if ($State.Snapshot) {
        $info.Id = [long]$State.Snapshot.id
        $finished = [string](Get-PcmGuiValue $State.Snapshot 'finished_utc' '')
        $info.Taken = Format-PcmGuiTime $finished
        $time = [datetime]::MinValue
        if ($finished -and [datetime]::TryParse($finished, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$time)) { $info.AgeHours = ([datetime]::UtcNow - $time.ToUniversalTime()).TotalHours }
    }
    return [pscustomobject]$info
}

function Get-PcmGuiBatchSummary {
    <#
        One line per Convert batch: its objects, its status, and how far its Recover went. An object is waiting for Recover when it was converted (at least started) and no Recover batch of
        this Convert batch finished it.
    #>
    param([Parameter(Mandatory)][object]$State)
    $itemsByBatch = @{}
    foreach ($item in $State.Items) {
        $key = ([string]$item.batch_id).ToLowerInvariant()
        if (-not $itemsByBatch.ContainsKey($key)) { $itemsByBatch[$key] = [Collections.Generic.List[object]]::new() }
        $itemsByBatch[$key].Add($item)
    }
    $summary = [Collections.Generic.List[object]]::new()
    foreach ($batch in @($State.Batches | Where-Object { $_.action -eq 'Convert' })) {
        $id = [string]$batch.id
        $items = @(if ($itemsByBatch.ContainsKey($id.ToLowerInvariant())) { $itemsByBatch[$id.ToLowerInvariant()] })
        $recovers = @($State.Batches | Where-Object { $_.action -eq 'Recover' -and (Test-PcmGuiOrdinal ([string]$_.convert_batch) $id) })
        $back = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
        foreach ($recover in $recovers) {
            $rid = ([string]$recover.id).ToLowerInvariant()
            if ($itemsByBatch.ContainsKey($rid)) { foreach ($done in @($itemsByBatch[$rid] | Where-Object { $_.status -eq 'Done' })) { [void]$back.Add([string]$done.object_guid) } }
        }
        $converted = @($items | Where-Object { $_.step -and $_.status -ne 'Skipped' })
        $waiting = @($converted | Where-Object { -not $back.Contains([string]$_.object_guid) })
        # A rollback is interrupted when it did not end or failed, or when it ended with objects of its own that are not back yet. One that ended with all its objects back is a wave that finished.
        $openRecovers = [Collections.Generic.List[object]]::new(); $undone = [Collections.Generic.List[object]]::new()
        $undoneSeen = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
        foreach ($recover in $recovers) {
            $rid = ([string]$recover.id).ToLowerInvariant()
            $own = @(if ($itemsByBatch.ContainsKey($rid)) { $itemsByBatch[$rid] | Where-Object { $_.status -notin @('Done', 'Skipped') -and -not $back.Contains([string]$_.object_guid) } })
            if ($recover.status -in @('Running', 'Failed') -or ($recover.status -eq 'Partial' -and $own.Count)) {
                $openRecovers.Add($recover)
                foreach ($item in $own) { if ($undoneSeen.Add([string]$item.object_guid)) { $undone.Add($item) } }
            }
        }
        $wave = @(if ($undone.Count -and $undone.Count -lt $waiting.Count) { foreach ($item in $undone) { [pscustomobject]@{ Guid = [string]$item.object_guid; Name = [string]$item.identity } } })
        $left = @($items | Where-Object { $_.status -in @('Planned', 'Running', 'Failed') })
        $recover = if (-not $recovers.Count) { if ($converted.Count) { 'Not started' } else { '-' } }
            elseif (@($recovers | Where-Object { $_.status -eq 'Complete' }).Count -and -not $waiting.Count) { 'Complete' }
            else { 'Partial: {0} of {1} back' -f ($converted.Count - $waiting.Count), $converted.Count }
        $tone = if ($recover -eq 'Complete') { 'Success' } elseif ($recover -like 'Partial*') { 'Caution' } else { 'Secondary' }
        $summary.Add([pscustomobject]@{ Id = $id; Created = (Format-PcmGuiTime ([string]$batch.created_utc)); CreatedUtc = [string]$batch.created_utc
                Status = [string]$batch.status; Objects = $items.Count; Users = @($items | Where-Object { $_.kind -eq 'User' }).Count
                Shared = @($items | Where-Object { $_.kind -ne 'User' }).Count; Converted = $converted.Count; Waiting = $waiting.Count; Left = $left.Count
                Recover = $recover; Tone = $tone; Items = $items; Back = $back; Account = [string]$batch.account; Environment = [string](Get-PcmGuiValue $batch 'environment' '')
                SnapshotId = [string](Get-PcmGuiValue $batch 'snapshot_id' ''); Recovers = $recovers; OpenRecovers = $openRecovers.ToArray(); Undone = $undone.Count; Wave = $wave })
    }
    return $summary.ToArray()
}

function Get-PcmGuiBatchEntry {
    <#
    .SYNOPSIS
        The rows of the batch list of the Task page for a task: the batches of the journal with what can be done with them. Selectable = false greys a row.
    .PARAMETER Kind
        ContinueConvert (a Convert batch with objects left), Rollback (a Convert batch with mailboxes in the cloud), ContinueRollback (a rollback that stopped), Follow (any batch, for a check).
    #>
    param([Parameter(Mandatory)][object]$State, [Parameter(Mandatory)][ValidateSet('ContinueConvert', 'Rollback', 'ContinueRollback', 'Follow')][string]$Kind)
    $rows = [Collections.Generic.List[object]]::new()
    $row = {
        param($Title, $Sub, $State_, $Level, $Selectable, $Glyph, $BatchId, $Environment, $Action, $Wave = @())
        [Dynamic.ExpandoObject]$object = [Dynamic.ExpandoObject]::new()
        foreach ($pair in @{ Title = $Title; Sub = $Sub; State = $State_; Level = $Level; Selectable = $Selectable; Glyph = $Glyph; BatchId = $BatchId; Environment = $Environment; Action = $Action; Wave = @($Wave) }.GetEnumerator()) { $object.($pair.Key) = $pair.Value }
        return , $object
    }
    $convertGlyph = [string][char]0xE753; $recoverGlyph = [string][char]0xE7A7
    $summary = if ($Kind -eq 'Follow') { @() } else { @(Get-PcmGuiBatchSummary -State $State) }
    if ($Kind -eq 'Follow') {
        $objectCount = New-Object 'Collections.Generic.Dictionary[string,int]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
        foreach ($item in @($State.Items)) { $objectCount[[string]$item.batch_id] = 1 + $(if ($objectCount.ContainsKey([string]$item.batch_id)) { $objectCount[[string]$item.batch_id] } else { 0 }) }
        foreach ($entry in @($State.Batches | Sort-Object { [string]$_.created_utc } -Descending)) {
            $isRecover = $entry.action -eq 'Recover'
            $own = [pscustomobject]@{ Count = $(if ($objectCount.ContainsKey([string]$entry.id)) { $objectCount[[string]$entry.id] } else { 0 }) }
            $sub = 'Batch {0} · snapshot #{1} · {2}' -f $entry.id, (Get-PcmGuiValue $entry 'snapshot_id' '?'), (ConvertTo-PcmGuiPlural $own.Count 'object' 'objects')
            if ($isRecover -and $entry.convert_batch) { $sub += ' · rolls back {0}' -f $entry.convert_batch }
            $level = switch ([string]$entry.status) { 'Complete' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }
            $rows.Add((& $row ('{0} · {1}' -f $entry.action, (Format-PcmGuiDate ([string]$entry.created_utc))) $sub ([string]$entry.status) $level $true $(if ($isRecover) { $recoverGlyph } else { $convertGlyph }) ([string]$entry.id) ([string](Get-PcmGuiValue $entry 'environment' '')) ([string]$entry.action)))
        }
        return $rows.ToArray()
    }
    foreach ($batch in $summary) {
        $subBase = 'Batch {0} · snapshot #{1}' -f $batch.Id, $(if ($batch.SnapshotId) { $batch.SnapshotId } else { '?' })
        switch ($Kind) {
            'ContinueConvert' {
                $done = $batch.Objects - $batch.Left
                $selectable = $batch.Left -gt 0 -or ($batch.Status -in @('Partial', 'Running', 'Failed'))
                $reason = switch ($batch.Status) { 'Failed' { ' (it failed)' } 'Running' { ' (it did not end)' } 'Partial' { ' (it stopped)' } default { '' } }
                $sub = '{0} · {1}, {2} done, {3} left{4}' -f $subBase, (ConvertTo-PcmGuiPlural $batch.Objects 'object' 'objects'), $done, $batch.Left, $reason
                $rows.Add((& $row ('Convert · ' + (Format-PcmGuiDate $batch.CreatedUtc)) $sub $(if ($selectable) { 'Can be continued' } else { 'Complete' }) $(if ($selectable) { 'Warn' } else { 'Ok' }) $selectable $convertGlyph $batch.Id $batch.Environment 'Convert'))
            }
            'Rollback' {
                $selectable = $batch.Waiting -gt 0
                $state_ = if ($batch.Waiting -gt 0 -and $batch.Recover -like 'Partial*') { 'Partly rolled back' } elseif ($batch.Waiting -gt 0) { 'Can be rolled back' } elseif ($batch.Converted -gt 0) { 'Rolled back' } else { 'Nothing in the cloud' }
                $level = if ($state_ -eq 'Partly rolled back') { 'Warn' } elseif ($selectable) { 'Ok' } else { 'Info' }
                $sub = '{0} · {1}, {2} in the cloud' -f $subBase, (ConvertTo-PcmGuiPlural $batch.Objects 'object' 'objects'), $batch.Waiting
                if ($batch.Converted -gt $batch.Waiting) { $sub += ', {0} already rolled back' -f ($batch.Converted - $batch.Waiting) }
                $rows.Add((& $row ('Convert · ' + (Format-PcmGuiDate $batch.CreatedUtc)) $sub $state_ $level $selectable $recoverGlyph $batch.Id $batch.Environment 'Convert'))
            }
            'ContinueRollback' {
                $open = @($batch.OpenRecovers)
                if (-not $open.Count) { continue }
                $reason = switch ([string]$open[0].status) { 'Failed' { 'it failed' } 'Running' { 'it did not end' } default { 'it stopped' } }
                $sub = if ($batch.Wave.Count) { '{0} · {1} of the interrupted wave still in the cloud ({2})' -f $subBase, $batch.Wave.Count, $reason } else { '{0} · {1} still in the cloud ({2})' -f $subBase, $batch.Waiting, $reason }
                $rows.Add((& $row ('Rollback · ' + (Format-PcmGuiDate ([string]$open[0].created_utc))) $sub 'Can be continued' 'Warn' ($batch.Waiting -gt 0) $recoverGlyph $batch.Id $batch.Environment 'Recover' $batch.Wave))
            }
        }
    }
    return $rows.ToArray()
}

function Get-PcmGuiWaveRow {
    <# The mailboxes of a Convert batch that are in the cloud (converted, not rolled back), for the grid of a wave: Guid, Name, Upn, Type, Result. #>
    param([Parameter(Mandatory)][object]$State, [Parameter(Mandatory)][string]$BatchId)
    $batch = @(Get-PcmGuiBatchSummary -State $State | Where-Object { Test-PcmGuiOrdinal $_.Id $BatchId }) | Select-Object -First 1
    if (-not $batch) { return @() }
    $names = @{}
    foreach ($object in $State.Objects) { $names[([string]$object.object_guid).ToLowerInvariant()] = $object }
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($item in @($batch.Items | Where-Object { $_.step -and $_.status -ne 'Skipped' -and -not $batch.Back.Contains([string]$_.object_guid) })) {
        $object = $names[([string]$item.object_guid).ToLowerInvariant()]
        $name = if ($object -and $object.display_name) { [string]$object.display_name } elseif ($object -and $object.name) { [string]$object.name } else { [string]$item.identity }
        $rows.Add([pscustomobject]@{ Guid = [string]$item.object_guid; Name = $name; Upn = [string]$item.identity; Type = $(if ($item.kind -eq 'User') { 'User' } else { 'Shared' })
                Result = $(if ($item.status -eq 'Done') { 'Active in Exchange Online' } else { 'Conversion not finished ({0})' -f $item.status }) })
    }
    return $rows.ToArray()
}

function Write-PcmGuiWave {
    <# The objects ticked for a wave, as the CSV given to -CsvPath: a column Identity (the object ID) and a column Name for the reader. #>
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][object[]]$Rows)
    $null = New-Item -ItemType Directory -Path $Folder -Force
    $path = Join-Path $Folder ('wave-{0}-{1}.csv' -f $Name, (Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add('Identity,Name')
    foreach ($row in $Rows) { $lines.Add(('{0},"{1}"' -f $row.Guid, ([string]$row.Name).Replace('"', "'"))) }
    [IO.File]::WriteAllLines($path, $lines, (New-Object Text.UTF8Encoding($true)))
    return $path
}
#endregion

#region Lookups made in the snapshot (no engine, no cloud)
$script:GuiObjectIndex = $null

function Get-PcmGuiObjectCache {
    <#
        The mailboxes of the snapshot arranged for lookups, built once for each state that is read: ByGuid (by object ID), ByName (every text that names a mailbox: sign-in name, e-mail address,
        account name, name, display name, object ID; a list of mailboxes each), and the organizational unit and distinguished name of each mailbox. Letter case is ignored.
    #>
    param([Parameter(Mandatory)][object]$State)
    if ($script:GuiObjectIndex -and [object]::ReferenceEquals($script:GuiObjectIndex.State, $State)) { return $script:GuiObjectIndex }
    $comparer = [StringComparer]::OrdinalIgnoreCase
    $byGuid = New-Object 'Collections.Generic.Dictionary[string,object]' -ArgumentList $comparer
    $byName = New-Object 'Collections.Generic.Dictionary[string,Collections.Generic.List[object]]' -ArgumentList $comparer
    $objects = @($State.Objects)
    $units = [string[]]::new($objects.Count); $names = [string[]]::new($objects.Count)
    for ($i = 0; $i -lt $objects.Count; $i++) {
        $object = $objects[$i]
        $guid = [string]$object.object_guid
        if ($guid -and -not $byGuid.ContainsKey($guid)) { $byGuid[$guid] = $object }
        foreach ($field in 'object_guid', 'user_principal_name', 'primary_smtp_address', 'sam_account_name', 'name', 'display_name') {
            $property = $object.PSObject.Properties[$field]
            if (-not $property) { continue }
            $text = [string]$property.Value
            if (-not $text) { continue }
            if (-not $byName.ContainsKey($text)) { $byName[$text] = [Collections.Generic.List[object]]::new() }
            if (-not $byName[$text].Contains($object)) { $byName[$text].Add($object) }
        }
        $units[$i] = [string](Get-PcmGuiValue $object 'organizational_unit' '')
        $names[$i] = [string](Get-PcmGuiValue $object 'distinguished_name' '')
    }
    $script:GuiObjectIndex = @{ State = $State; Objects = $objects; ByGuid = $byGuid; ByName = $byName; Unit = $units; Dn = $names }
    return $script:GuiObjectIndex
}

function Get-PcmGuiObjectIndex {
    <# The mailboxes of the snapshot by object ID, letter case ignored. #>
    param([Parameter(Mandatory)][object]$State)
    return (Get-PcmGuiObjectCache -State $State).ByGuid
}

function Find-PcmGuiMailbox {
    <# The mailboxes of the snapshot that a text names: sign-in name, e-mail address, account name, name, display name or ID, letter case ignored. #>
    param([Parameter(Mandatory)][object]$State, [AllowEmptyString()][string]$Text)
    $key = $Text.Trim()
    if (-not $key) { return @() }
    $guid = [guid]::Empty
    if ([guid]::TryParse($key, [ref]$guid)) { $key = $guid.ToString() }
    $list = $null
    if ((Get-PcmGuiObjectCache -State $State).ByName.TryGetValue($key, [ref]$list)) { return $list.ToArray() }
    return @()
}

function Get-PcmGuiOuList {
    <# The organizational units of the snapshot with the number of mailboxes under each (sub-OUs included): Name (the path shown by Exchange, contoso.com/Paris) and Count. #>
    param([Parameter(Mandatory)][object]$State)
    $counts = New-Object 'Collections.Generic.Dictionary[string,int]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    foreach ($object in $State.Objects) {
        $path = [string](Get-PcmGuiValue $object 'organizational_unit' '')
        if (-not $path) { continue }
        $parts = $path.Split('/')
        for ($depth = 2; $depth -le $parts.Count; $depth++) {
            $name = ($parts[0..($depth - 1)]) -join '/'
            if ($counts.ContainsKey($name)) { $counts[$name]++ } else { $counts[$name] = 1 }
        }
    }
    return @($counts.Keys | Sort-Object | ForEach-Object { [pscustomobject]@{ Name = [string]$_; Count = [int]$counts[$_] } })
}

function Get-PcmGuiOuMatch {
    <# The mailboxes of the snapshot under an OU given as a path (contoso.com/Paris) or a distinguished name (OU=Paris,DC=contoso,DC=com), sub-OUs included. #>
    param([Parameter(Mandatory)][object]$State, [AllowEmptyString()][string]$Text)
    $key = $Text.Trim().TrimEnd('/')
    if (-not $key) { return @() }
    $isDn = $key -match '^(OU|DC|CN)\s*=' -and $key.Contains('=')
    $cache = Get-PcmGuiObjectCache -State $State
    $found = [Collections.Generic.List[object]]::new()
    $suffix = ',' + $key; $prefix = $key + '/'
    for ($i = 0; $i -lt $cache.Objects.Count; $i++) {
        if ($isDn) {
            $dn = $cache.Dn[$i]
            if ($dn.EndsWith($suffix, [StringComparison]::OrdinalIgnoreCase) -or [string]::Equals($dn, $key, [StringComparison]::OrdinalIgnoreCase)) { $found.Add($cache.Objects[$i]) }
            continue
        }
        $ou = $cache.Unit[$i]
        if ([string]::Equals($ou, $key, [StringComparison]::OrdinalIgnoreCase) -or $ou.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { $found.Add($cache.Objects[$i]) }
    }
    return $found.ToArray()
}

function Get-PcmGuiGroupInSnapshot {
    <#
    .SYNOPSIS
        Is this group one that Collect captured in the snapshot? Reads the groups and members of the snapshot, read-only.
    .OUTPUTS
        HasTables (the snapshot knows captured groups), Count (groups that match), Name, Members (people and mailboxes, nested groups followed), Mailboxes, Skipped, Candidates, Error.
    #>
    param([Parameter(Mandatory)][object]$State, [AllowEmptyString()][string]$Value)
    $answer = [pscustomobject]@{ HasTables = $false; Count = 0; Name = ''; Members = 0; Mailboxes = 0; Skipped = 0; Candidates = @(); Error = '' }
    if (-not $State.Config -or -not $State.Snapshot -or -not $Value.Trim()) { return $answer }
    try {
        $connection = Open-PcmStore -Path $State.Config.Store.Path -Root $State.Root -ReadOnly
        try {
            $found = Find-PcmTargetGroup $connection ([long]$State.Snapshot.id) $Value
            $answer.HasTables = [bool]$found.HasTables
            $groups = @($found.Groups)
            $answer.Count = $groups.Count
            $answer.Candidates = @($groups | ForEach-Object { [string]$_.name })
            if ($groups.Count -eq 1) {
                $group = $groups[0]
                $answer.Name = [string]$group.name
                $members = @(Invoke-PcmSql $connection 'SELECT member_guid, member_class FROM target_member WHERE snapshot_id = @id AND group_guid = @g COLLATE NOCASE' @{ id = [long]$State.Snapshot.id; g = [string]$group.group_guid } -As Rows)
                $guids = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
                foreach ($object in $State.Objects) { [void]$guids.Add([string]$object.object_guid) }
                $people = @($members | Where-Object { -not (Test-PcmGuiOrdinal $_.member_class 'Group') })
                $distinct = Get-PcmGuiOrdinalUnique @($people | ForEach-Object { [string]$_.member_guid })
                $answer.Members = @($distinct).Count
                $answer.Mailboxes = @($distinct | Where-Object { $guids.Contains([string]$_) }).Count
                $answer.Skipped = $answer.Members - $answer.Mailboxes
            }
        } finally { Close-PcmStore $connection }
    } catch { $answer.Error = $_.Exception.Message }
    return $answer
}

function Get-PcmGuiSnapshotFact {
    <# What the snapshot holds about one object, as lines of text for the object card of the Run page. #>
    param([Parameter(Mandatory)][object]$State, [AllowEmptyString()][string]$Identity)
    if (-not $State.Config -or -not $State.Snapshot -or -not $Identity) { return 'Nothing in the snapshot for this object.' }
    $match = @(Find-PcmGuiMailbox -State $State -Text $Identity) | Select-Object -First 1
    if (-not $match) { return 'Nothing in the snapshot for this object.' }
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add(('Snapshot #{0} ({1})' -f $State.Snapshot.id, (Format-PcmGuiTime ([string](Get-PcmGuiValue $State.Snapshot 'finished_utc' '')))))
    try {
        $connection = Open-PcmStore -Path $State.Config.Store.Path -Root $State.Root -ReadOnly
        try {
            $row = @(Invoke-PcmSql $connection 'SELECT email_addresses_json, recipient_type_details, organizational_unit FROM mailbox WHERE snapshot_id = @id AND object_guid = @g COLLATE NOCASE' @{ id = [long]$State.Snapshot.id; g = [string]$match.object_guid } -As Rows) | Select-Object -First 1
            if ($row -and $row.email_addresses_json) { foreach ($address in @($row.email_addresses_json | ConvertFrom-Json)) { foreach ($one in @($address)) { $lines.Add(([string]$one -replace '^smtp:', 'SMTP ')) } } }
            else { $lines.Add('SMTP ' + [string]$match.primary_smtp_address) }
            $lines.Add('Mailbox type: ' + $(if ($row -and $row.recipient_type_details) { [string]$row.recipient_type_details } else { ([string]$match.kind).ToLowerInvariant() }))
            if ($row -and $row.organizational_unit) { $lines.Add('OU: ' + [string]$row.organizational_unit) }
        } finally { Close-PcmStore $connection }
    } catch { $lines.Add('SMTP ' + [string]$match.primary_smtp_address); $lines.Add('Mailbox type: ' + ([string]$match.kind).ToLowerInvariant()) }
    return ($lines -join "`r`n")
}
#endregion

# SIG # Begin signature block
# MIImKwYJKoZIhvcNAQcCoIImHDCCJhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcPZ3mZUXkoFFL
# 71+CWtw87dfnMYZN6rYCV6GyfCahxaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEjCCGA4CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBKTOWXKq/p
# tXDaRsUJhMuX8M2pRQ2IBeKMH1AwfuTSoTANBgkqhkiG9w0BAQEFAASCAgCCfE9j
# Ly5YNEoko02AIsIEfxUbAJKOuftL62HZXkOznB4bNACqGjNH0pC9uKRNySVDC84X
# bAPRkasMBrnF6oWGBugP5vxwTUgIxOHkop/eWdAGWkZr5Xdvi60Q74/Dy8u817yw
# bV/Yqi64yeoLfw9crsEWmzJdwtyjwRWdMSxmu5ZkDVFwOGUvSCPiWGVnM/Z24g4o
# 6MBJsV9FxUL38aInloD5+dlmzNqSMmx7U9Z6k/OZ+yLV8u4fobKf5NZJcQGgdoBf
# cQUtBKGzmDALtk0GmPdNFwlkILaIgkrBmuAxTuwFmTC/zcuvapBmtfpd4Nfu9ldT
# NzASkN8fLNonaow7OFOOZYHx+h4EDmKgV1ACfhB5JJHF2lZnLBRhhq4OuegdKRfP
# zuz6R6WRXerHoroFi+k4R7mhy8ZyY9qQU1OkYRHaTvg76oRmu7d9J/+OQRRxZOWN
# 6JEd/ePmOTyeLPhu87JRW0lRwfUODUmGDZRt8sP2Z1Yd1kmfGu58ZxisMYjdyWYR
# /vpBTvFuGYqM68/qJUOmNtF/X/WJnS6rXBViyz7aEhZWvQMTWGxdqZmzvlHNTsft
# Ws17+7ELStFNr2U16unZF93KnWr3GLbgmVoBsaQWzaLFAB1Uv/9Wbv0M051+KSNn
# rbpDpXjojFlSiRzSNfAxrFQt+tkbciy1NkHI/qGCFM8wghTLBgorBgEEAYI3AwMB
# MYIUuzCCFLcGCSqGSIb3DQEHAqCCFKgwghSkAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIDvT2LxdrV5HVNBgG23xULfjQiAeKeMbUj9sQulE0LxPAggcmoNm/anW
# bhgPMjAyNjExMTUyMDA2MjhaMAMCAQECEGcpTryD/RksDIfp8citrzygghGnMIIE
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
# wzSPOzGCAlcwggJTAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTIwMDYyOFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIGzB
# RlSM4bPft519lvavyP2SKiDuS/57HQkDX+YH2QCsMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIESDBGAiEAqxTVvaimhc/HhObmZTxYRcZ4qrnlYGhU3GVcrurIz4QCIQDI
# opuLWBHMo09ldrDUJxku53Pr1W9MD8/81YGMHyY+gQ==
# SIG # End signature block
