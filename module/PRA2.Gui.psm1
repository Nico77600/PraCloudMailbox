<#
.SYNOPSIS
    PRA Cloud Mailbox - the window (Invoke-PraCloudMailbox.ps1 -Gui).

.DESCRIPTION
    A WPF front-end for the four actions. It never changes anything itself: every action runs
    Invoke-PraCloudMailbox.ps1 in a child process - Windows PowerShell 5.1 for Collect, PowerShell 7 for Check,
    Convert and Recover - exactly as on the command line, with the same log, transcript, journal and report. The
    window follows the run through a file of events (PRA_EVENT_FILE, one JSON object per line, written by
    PRA2.Common), asks for a stop with PRA_STOP_FILE (the run starts no new object, then ends normally and can be
    resumed) and answers the questions of the run (Entra Connect Manual mode) in the answer file of the event.

    What the window adds: the state at a glance (configuration, last snapshot, batches, this computer, the next
    step), the readiness of every object in a filterable table, waves (the objects ticked are written to a list
    given to -IdentityPath), a preview that must have run on exactly the same selection before Apply, and a typed
    confirmation (CONVERT, RECOVER) before any change.

    Runs in Windows PowerShell 5.1 and PowerShell 7 (Fluent theme of Windows 11 with PowerShell 7.5 or later,
    classic WPF controls with the same colours otherwise). The layout is in PRA2.Gui.xaml.

    Regions:
        1. Theme            Initialize-PraGuiTheme, Set-PraGuiTheme (Fluent or classic, light or dark)
        2. Data             what the window shows, read from the configuration, the snapshot database, the
                            journal and the last Check report; the command line of each run (pure functions)
        3. Window           New-PraGuiWindow, the pages and their tables
        4. Runs             Start-PraGuiRun, the events of the child process, the result
        5. Entry            Show-PraGui

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.0
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:Gui = $null
$script:IconFont = 'Segoe Fluent Icons, Segoe MDL2 Assets'
$script:Pages = @(
    [pscustomobject]@{ Name = 'Overview'; Glyph = 0xE80F; Title = 'Overview'; Detail = 'state and next step' }
    [pscustomobject]@{ Name = 'Collect'; Glyph = 0xE896; Title = '1  Collect'; Detail = 'Exchange server, every day' }
    [pscustomobject]@{ Name = 'Check'; Glyph = 0xE9D5; Title = '2  Check'; Detail = 'read-only, every week' }
    [pscustomobject]@{ Name = 'Convert'; Glyph = 0xE753; Title = '3  Convert'; Detail = 'disaster' }
    [pscustomobject]@{ Name = 'Recover'; Glyph = 0xE7A7; Title = '4  Recover'; Detail = 'infrastructure rebuilt' }
)
# Lines of the activity panel: icon and colour of each status of the console.
$script:LineStyle = @{
    Step = @{ Glyph = 0xE76C; Tone = 'Brand' }; Ok = @{ Glyph = 0xE73E; Tone = 'Success' }; Warn = @{ Glyph = 0xE7BA; Tone = 'Caution' }
    Fail = @{ Glyph = 0xEA39; Tone = 'Critical' }; Info = @{ Glyph = 0xE946; Tone = 'Secondary' }; Skip = @{ Glyph = 0xE72A; Tone = 'Tertiary' }
    Sub = @{ Glyph = 0xE76C; Tone = 'Tertiary' }
}
# At most this many detail lines (Planned: x, one per object) in the panel; the log has them all.
$script:MaxSubLines = 300

#region 1. Theme ------------------------------------------------------------------------------------

function New-PraGuiBrush {
    param([Parameter(Mandatory)][string]$Color)
    $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

function Test-PraGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    if (-not $personalize -or -not $personalize.PSObject.Properties['AppsUseLightTheme']) { return $false }
    return [string]$personalize.AppsUseLightTheme -eq '0'
}

function Initialize-PraGuiTheme {
    <#
        Loads WPF and applies the theme to the application: Fluent (.NET 9 and later, PowerShell 7.5+), light or dark
        as Windows (System), or Light / Dark for the documentation images. Returns Fluent and Dark.
    #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml, System.Data
    # One application per process: created once, never shut down by a closed window.
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $fluent = $null -ne [Windows.Application].GetProperty('ThemeMode')
    $dark = $Theme -eq 'Dark'
    # The classic controls (text boxes, lists) stay light: a dark window would only darken the background around them.
    if ($Theme -eq 'System' -and $fluent) { $dark = Test-PraGuiDarkMode }
    if ($fluent) {
        # Light or Dark, never System: without the setting (Windows Server) WPF would pick dark while the window is light.
        $app.ThemeMode = [Windows.ThemeMode]::new($(if ($dark) { 'Dark' } else { 'Light' }))
    }
    return [pscustomobject]@{ Fluent = $fluent; Dark = $dark; Application = $app }
}

function Set-PraGuiTheme {
    <#
        Colours of a window: the accent of the reports (crimson) on the accent resources, the status colours, and with
        the classic theme the Fluent resources the window uses, with the palette of the reports.
    #>
    param([Parameter(Mandatory)][Windows.Window]$Window, [Parameter(Mandatory)][object]$Theme)
    $dark = [bool]$Theme.Dark
    $r = $Window.Resources
    $set = { param([string[]]$Keys, [string]$LightColor, [string]$DarkColor) $b = New-PraGuiBrush $(if ($dark) { $DarkColor } else { $LightColor }); foreach ($k in $Keys) { $r[$k] = $b } }
    if (-not $Theme.Fluent) {
        & $set 'ApplicationBackgroundBrush' '#F7F4EF' '#202020'
        & $set 'CardBackgroundFillColorDefaultBrush' '#FFFFFF' '#2B2B2B'
        & $set 'CardStrokeColorDefaultBrush', 'ControlStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'ControlFillColorDefaultBrush' '#FFFFFF' '#2D2D2D'
        & $set 'ControlFillColorSecondaryBrush' '#F3F0EB' '#323232'
        & $set 'DividerStrokeColorDefaultBrush' '#E6E2DC' '#3D3D3D'
        & $set 'LayerFillColorDefaultBrush' '#FCFBF8' '#262626'
        & $set 'TextFillColorPrimaryBrush' '#242424' '#FFFFFF'
        & $set 'TextFillColorSecondaryBrush' '#5C5C5C' '#C5C5C5'
        & $set 'TextFillColorTertiaryBrush' '#8A8A8A' '#9A9A9A'
        # Column headers of the tables, like the Fluent ones.
        $header = [Windows.Style]::new([Windows.Controls.Primitives.DataGridColumnHeader])
        $header.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::BackgroundProperty, $r['LayerFillColorDefaultBrush'])))
        $header.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::ForegroundProperty, $r['TextFillColorSecondaryBrush'])))
        $header.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::PaddingProperty, ([Windows.Thickness]::new(6, 6, 6, 6)))))
        $header.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::BorderBrushProperty, $r['DividerStrokeColorDefaultBrush'])))
        $header.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::BorderThicknessProperty, ([Windows.Thickness]::new(0, 0, 0, 1)))))
        $header.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::FontWeightProperty, [Windows.FontWeights]::SemiBold)))
        $r[[Windows.Controls.Primitives.DataGridColumnHeader]] = $header
    }
    & $set 'AccentFillColorDefaultBrush', 'AccentButtonBackground', 'AccentButtonBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'AccentFillColorSecondaryBrush', 'AccentButtonBackgroundPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'AccentFillColorTertiaryBrush', 'AccentButtonBackgroundPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentTextFillColorPrimaryBrush' '#9A1A41' '#FD8EA1'
    & $set 'PraBrand' '#B11F4B' '#B11F4B'
    & $set 'PraBrandText' '#B11F4B' '#FD8EA1'
    & $set 'PraAccentSoft' '#14B11F4B' '#33FD8EA1'
    & $set 'PraSuccess' '#16A34A' '#4ADE80'
    & $set 'PraCaution' '#B45309' '#FBBF24'
    & $set 'PraCritical' '#DC2626' '#F87171'
    & $set 'PraInfoBackground' '#F3F3F3' '#2E2E2E'
    & $set 'PraInfoBorder' '#E0E0E0' '#3D3D3D'
    & $set 'PraCautionBackground' '#FFF7E8' '#33FBBF24'
    & $set 'PraCautionBorder' '#F5D7A1' '#66FBBF24'
    & $set 'PraCriticalBackground' '#FDECEC' '#33F87171'
    & $set 'PraCriticalBorder' '#F4B4B4' '#66F87171'
    & $set 'PraSuccessBackground' '#EAF7EE' '#334ADE80'
    & $set 'PraSuccessBorder' '#B7E4C4' '#664ADE80'
}

function Get-PraGuiToneKey {
    <# The resource key of a tone: Brand, Success, Caution, Critical, Primary, Secondary, Tertiary. #>
    param([string]$Tone = 'Primary')
    switch ($Tone) {
        'Brand' { return 'PraBrandText' } 'Success' { return 'PraSuccess' } 'Caution' { return 'PraCaution' } 'Critical' { return 'PraCritical' }
        'Secondary' { return 'TextFillColorSecondaryBrush' } 'Tertiary' { return 'TextFillColorTertiaryBrush' } default { return 'TextFillColorPrimaryBrush' }
    }
}

function Get-PraGuiBrush {
    <# The brush of a tone in the window ($null before the window exists). #>
    param([string]$Tone = 'Primary')
    if ($script:Gui -and $script:Gui.Form) { return $script:Gui.Form.TryFindResource((Get-PraGuiToneKey $Tone)) }
    return $null
}

function Set-PraGuiAccentButton {
    <# A button in the accent colour, with the size of the other buttons. #>
    param([Parameter(Mandatory)][Windows.Controls.Button]$Button, [Parameter(Mandatory)][Windows.Window]$Window, [Parameter(Mandatory)][object]$Theme)
    if ($Theme.Fluent) {
        $accent = $Window.TryFindResource('AccentButtonStyle')
        if ($accent) {
            $style = [Windows.Style]::new([Windows.Controls.Button], $accent)
            $style.Setters.Add(([Windows.Setter]::new([Windows.Controls.Control]::PaddingProperty, ([Windows.Thickness]::new(14, 6, 14, 6)))))
            $style.Setters.Add(([Windows.Setter]::new([Windows.FrameworkElement]::MarginProperty, $Button.Margin)))
            $style.Setters.Add(([Windows.Setter]::new([Windows.FrameworkElement]::MinWidthProperty, [double]96)))
            $Button.Style = $style
            return
        }
    }
    $Button.SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'AccentFillColorDefaultBrush')
    $Button.SetResourceReference([Windows.Controls.Control]::BorderBrushProperty, 'AccentFillColorDefaultBrush')
    $Button.Foreground = [Windows.Media.Brushes]::White
}
#endregion

#region 2. Data -------------------------------------------------------------------------------------

function Get-PraGuiEngine {
    <#
        The PowerShell that runs an action: Windows PowerShell 5.1 for Collect (the Exchange cmdlets), PowerShell 7.4
        or later for the others. Returns Ok, Path, Version and the problem when there is none.
    #>
    param([Parameter(Mandatory)][ValidateSet('Collect', 'Check', 'Convert', 'Recover')][string]$Action)
    if ($Action -eq 'Collect') {
        $path = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
        $ok = Test-Path -LiteralPath $path
        return [pscustomobject]@{ Ok = $ok; Path = $path; Version = '5.1'; Problem = $(if ($ok) { '' } else { 'Windows PowerShell 5.1 (powershell.exe) not found.' }) }
    }
    $candidates = [Collections.Generic.List[string]]::new()
    if ($PSVersionTable.PSEdition -eq 'Core') { $candidates.Add((Get-Process -Id $PID).Path) }
    $command = Get-Command pwsh.exe -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($command) { $candidates.Add($command.Source) }
    $candidates.Add((Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'))
    $best = $null
    foreach ($candidate in ($candidates | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)) {
        $text = [regex]::Match([string](Get-Item -LiteralPath $candidate).VersionInfo.ProductVersion, '^\d+(\.\d+){1,3}').Value
        if (-not $text) { continue }
        $version = [version]$text
        if ($version -ge [version]'7.4') { return [pscustomobject]@{ Ok = $true; Path = $candidate; Version = $text; Problem = '' } }
        if (-not $best) { $best = "$candidate ($text)" }
    }
    $problem = 'PowerShell 7.4 or later (pwsh.exe) not found: Check, Convert and Recover run in PowerShell 7 (winget install Microsoft.PowerShell).'
    if ($best) { $problem = "PowerShell 7.4 or later needed, found ${best}: Check, Convert and Recover run in PowerShell 7.4+." }
    return [pscustomobject]@{ Ok = $false; Path = ''; Version = ''; Problem = $problem }
}

function Get-PraGuiCommand {
    <#
        The command line of one run: the same script and parameters as by hand. Apply gets -Force: the window asked
        its own confirmation. Returns FilePath, Arguments and Display (the command as an operator would type it).
    #>
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ConfigPath, [Parameter(Mandatory)][hashtable]$Request, [Parameter(Mandatory)][string]$Engine)
    $quote = { param([string]$Value) '"' + $Value + '"' }
    $tool = [Collections.Generic.List[string]]::new()
    $tool.Add('-Action'); $tool.Add([string]$Request.Action)
    $mode = if ($Request.ContainsKey('Mode') -and $Request.Mode) { [string]$Request.Mode } else { 'Preview' }
    if ($Request.Action -ne 'Check') { $tool.Add('-Mode'); $tool.Add($mode) }
    if ($Request.ContainsKey('Scope') -and $Request.Scope -and $Request.Scope -ne 'All') { $tool.Add('-Scope'); $tool.Add([string]$Request.Scope) }
    if ($Request.ContainsKey('Batch') -and $Request.Batch) { $tool.Add('-Batch'); $tool.Add([string]$Request.Batch) }
    $shown = [Collections.Generic.List[string]]::new()
    foreach ($part in $tool) { $shown.Add($part) }
    if ($Request.ContainsKey('IdentityPath') -and $Request.IdentityPath) {
        $tool.Add('-IdentityPath'); $tool.Add((& $quote $Request.IdentityPath))
        $shown.Add('-IdentityPath'); $shown.Add((& $quote $Request.IdentityPath))
    }
    $tool.Add('-ConfigPath'); $tool.Add((& $quote $ConfigPath))
    if ($mode -eq 'Apply') { $tool.Add('-Force'); $shown.Add('-Force') }
    # -NonInteractive: a prompt nobody can see fails the run instead of waiting forever (the window answers through events).
    $arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -File {0} {1}' -f (& $quote (Join-Path $Root 'Invoke-PraCloudMailbox.ps1')), ($tool -join ' ')
    $display = '{0} -File .\Invoke-PraCloudMailbox.ps1 {1}' -f $(if ($Request.Action -eq 'Collect') { 'powershell.exe' } else { 'pwsh' }), ($shown -join ' ')
    return [pscustomobject]@{ FilePath = $Engine; Arguments = $arguments; Display = $display }
}

function Read-PraGuiEvents {
    <#
        The events written since Position (complete lines only: the run may be writing the next one). Returns the
        events and the new position.
    #>
    param([Parameter(Mandatory)][string]$Path, [long]$Position = 0)
    $none = [pscustomobject]@{ Events = @(); Position = $Position }
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
    if ($last -lt 0) { return $none }
    $text = [Text.Encoding]::UTF8.GetString($buffer, 0, $last + 1)
    $events = [Collections.Generic.List[object]]::new()
    foreach ($line in $text.Split([char]10)) {
        if (-not $line.Trim()) { continue }
        try { $events.Add(($line | ConvertFrom-Json)) } catch { $events.Add([pscustomobject]@{ kind = 'item'; status = 'Info'; text = $line.Trim(); identity = '' }) }
    }
    return [pscustomobject]@{ Events = $events.ToArray(); Position = $Position + $last + 1 }
}

function Read-PraGuiLastCheck {
    <# The newest Check report of the report folder (its CSV, delimiter ;), or $null. #>
    param([string]$Folder)
    if (-not $Folder -or -not (Test-Path -LiteralPath $Folder)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $Folder -Filter 'PRA2_*.csv' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 300)
    foreach ($file in $files) {
        $head = @(Get-Content -LiteralPath $file.FullName -TotalCount 2 -Encoding UTF8)
        if ($head.Count -lt 2) { continue }
        $first = @($head | ConvertFrom-Csv -Delimiter ';') | Select-Object -First 1
        if (-not $first -or -not $first.PSObject.Properties['Action'] -or $first.Action -ne 'Check') { continue }
        $html = [IO.Path]::ChangeExtension($file.FullName, '.html')
        return [pscustomobject]@{ Path = $file.FullName; Html = $(if (Test-Path -LiteralPath $html) { $html } else { '' }); Time = $file.LastWriteTime
            Rows = @(Import-Csv -LiteralPath $file.FullName -Delimiter ';' -Encoding UTF8) }
    }
    return $null
}

function Read-PraGuiState {
    <#
        What the window shows, read without changing anything: the configuration, the snapshots and the objects of the
        last one, the batches of the journal with their objects, and the last Check report. A file that cannot be read
        is reported in the state (ConfigError, SnapshotError, JournalError), never thrown.
    #>
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ConfigPath)
    $state = [pscustomobject]@{ Root = $Root; ConfigPath = $ConfigPath; Config = $null; ConfigError = ''; Snapshot = $null; Snapshots = @(); Objects = @()
        Permissions = 0; SnapshotError = ''; Batches = @(); Items = @(); JournalError = ''; Check = $null; Read = (Get-Date) }
    try { $state.Config = Import-PraConfiguration -Path $ConfigPath -Root $Root }
    catch { $state.ConfigError = $_.Exception.Message; return $state }
    $config = $state.Config
    if (Test-Path -LiteralPath $config.Store.Path) {
        try {
            $connection = Open-Pra2Store -Path $config.Store.Path -Root $Root -ReadOnly
            try {
                $state.Snapshots = @(Get-Pra2SnapshotList $connection)
                $state.Snapshot = Get-Pra2Snapshot $connection
                if ($state.Snapshot) {
                    $id = [long]$state.Snapshot.id
                    $state.Objects = @(Invoke-Pra2Sql $connection 'SELECT object_guid, kind, user_principal_name, primary_smtp_address, sam_account_name FROM mailbox WHERE snapshot_id = @id ORDER BY rowid' @{ id = $id } -As Rows)
                    $state.Permissions = [int](Invoke-Pra2Sql $connection 'SELECT COUNT(*) FROM permission WHERE snapshot_id = @id' @{ id = $id } -As Scalar)
                }
            } finally { Close-Pra2Store $connection }
        } catch { $state.SnapshotError = $_.Exception.Message }
    }
    if (Test-Path -LiteralPath $config.Store.JournalPath) {
        try {
            $journal = Open-Pra2Store -Path $config.Store.JournalPath -Root $Root -Kind Journal -ReadOnly
            try {
                $state.Batches = @(Invoke-Pra2Sql $journal 'SELECT * FROM batch ORDER BY created_utc DESC' -As Rows)
                $state.Items = @(Invoke-Pra2Sql $journal 'SELECT batch_id, object_guid, kind, identity, step, status, message FROM batch_item' -As Rows)
            } finally { Close-Pra2Store $journal }
        } catch { $state.JournalError = $_.Exception.Message }
    }
    $state.Check = Read-PraGuiLastCheck -Folder $config.Report.Folder
    return $state
}

function Get-PraGuiBatchSummary {
    <#
        One line per Convert batch: its objects, its status, and how far its Recover went. An object is waiting for
        Recover when it was converted (at least started) and no Recover batch of this Convert batch finished it.
    #>
    param([Parameter(Mandatory)][object]$State)
    $itemsByBatch = @{}
    foreach ($item in $State.Items) {
        $key = [string]$item.batch_id
        if (-not $itemsByBatch.ContainsKey($key)) { $itemsByBatch[$key] = [Collections.Generic.List[object]]::new() }
        $itemsByBatch[$key].Add($item)
    }
    $summary = [Collections.Generic.List[object]]::new()
    foreach ($batch in @($State.Batches | Where-Object { $_.action -eq 'Convert' })) {
        $id = [string]$batch.id
        $items = @(if ($itemsByBatch.ContainsKey($id)) { $itemsByBatch[$id] })
        $recovers = @($State.Batches | Where-Object { $_.action -eq 'Recover' -and [string]$_.convert_batch -eq $id })
        $back = New-Object 'Collections.Generic.HashSet[string]'
        foreach ($recover in $recovers) {
            $rid = [string]$recover.id
            if ($itemsByBatch.ContainsKey($rid)) { foreach ($done in @($itemsByBatch[$rid] | Where-Object { $_.status -eq 'Done' })) { [void]$back.Add([string]$done.object_guid) } }
        }
        $converted = @($items | Where-Object { $_.step -and $_.status -ne 'Skipped' })
        $waiting = @($converted | Where-Object { -not $back.Contains([string]$_.object_guid) })
        $recover = if (-not $recovers.Count) { if ($converted.Count) { 'Not started' } else { '-' } }
            elseif (@($recovers | Where-Object { $_.status -eq 'Complete' }).Count) { 'Complete' }
            else { 'Partial: {0} of {1} back' -f ($converted.Count - $waiting.Count), $converted.Count }
        $tone = if ($recover -eq 'Complete') { 'Success' } elseif ($recover -like 'Partial*') { 'Caution' } else { 'Secondary' }
        $summary.Add([pscustomobject]@{ Id = $id; Created = (Format-PraGuiTime ([string]$batch.created_utc))
                Status = [string]$batch.status; Objects = $items.Count; Users = @($items | Where-Object { $_.kind -eq 'User' }).Count
                Shared = @($items | Where-Object { $_.kind -ne 'User' }).Count; Converted = $converted.Count; Waiting = $waiting.Count
                Recover = $recover; Tone = $tone; Items = $items; Back = $back; Account = [string]$batch.account })
    }
    return $summary.ToArray()
}

function Get-PraGuiNextStep {
    <# The next thing to do, from the state: Title, Text, Page (or '' = the configuration), Button, Glyph. #>
    param([Parameter(Mandatory)][object]$State, [object[]]$Batches = @())
    $step = { param($Title, $Text, $Page, $Button, $Glyph) [pscustomobject]@{ Title = $Title; Text = $Text; Page = $Page; Button = $Button; Glyph = $Glyph } }
    if ($State.ConfigError) { return (& $step 'Fix the configuration' $State.ConfigError '' 'Open the configuration' 0xE713) }
    $config = $State.Config
    if (-not $config.Cloud.TenantId -or -not $config.Cloud.AppId -or -not $config.Cloud.CertificateThumbprint) {
        return (& $step 'Fill in the configuration' 'The tenant, the app registration and its certificate are empty: Check, Convert and Recover cannot sign in. Fill in config\PraCloudMailbox.config.psd1 (developer guide, chapter 6), then copy the same file to the Exchange server.' '' 'Open the configuration' 0xE713)
    }
    if ($State.SnapshotError) { return (& $step 'The snapshot database cannot be read' $State.SnapshotError 'Collect' 'Collect' 0xE896) }
    if (-not $State.Snapshot) { return (& $step 'Collect the first snapshot' 'On an Exchange server: Collect, then copy data\PraCloudMailbox.db to this computer (a daily scheduled task does both). After the disaster nothing on-premises can be read.' 'Collect' 'Collect' 0xE896) }
    $waiting = @($Batches | Where-Object { $_.Waiting -gt 0 })
    $partial = @($Batches | Where-Object { $_.Status -in @('Partial', 'Running') })
    if ($partial.Count) { return (& $step ('Convert batch {0} is not finished' -f $partial[0].Id) 'Some objects of this batch were not converted yet (Pending, stopped or failed). Preview its resume on the Convert page: only what is not done is taken again.' 'Convert' 'Convert' 0xE753) }
    if ($waiting.Count) {
        $count = ($waiting | Measure-Object -Property Waiting -Sum).Sum
        return (& $step ('{0} converted object(s) waiting for Recover' -f $count) 'Their mailboxes are in Exchange Online. Once Active Directory (restored), Exchange and Entra Connect work again, roll their batch back on the Recover page.' 'Recover' 'Recover' 0xE7A7)
    }
    $taken = [datetime]::Parse([string]$State.Snapshot.finished_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
    $age = ((Get-Date).ToUniversalTime() - $taken.ToUniversalTime()).TotalDays
    if ($age -gt $config.Store.MaxSnapshotAgeDays) { return (& $step ('The last snapshot is {0:0} day(s) old' -f $age) 'Permissions and addresses may have changed since: check the daily Collect on the Exchange server and the copy of the database.' 'Collect' 'Collect' 0xE896) }
    if (-not $State.Check -or $State.Check.Time.ToUniversalTime() -lt $taken.ToUniversalTime()) { return (& $step 'Check the last snapshot' 'Read-only: whether every object of the snapshot is ready for Convert, and the licences and permissions of the app.' 'Check' 'Check' 0xE9D5) }
    $notReady = @($State.Check.Rows | Where-Object { $_.FinalStatus -eq 'Error' }).Count
    if ($notReady) { return (& $step ('{0} object(s) not ready' -f $notReady) 'Convert would refuse them. Fix what the Check says (filter "Not ready"), then run Check again.' 'Check' 'Check' 0xE9D5) }
    return (& $step 'Ready' 'Every object of the last snapshot is ready. Keep the daily Collect and run Check every week; at the disaster, Convert.' 'Check' 'Check' 0xE9D5)
}

function Format-PraGuiTime {
    <# A UTC time of the store or the journal (ISO 8601), in local time: yyyy-MM-dd HH:mm; '' when empty. #>
    param([AllowEmptyString()][string]$Value)
    if (-not $Value) { return '' }
    $time = [datetime]::MinValue
    if (-not [datetime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$time)) { return $Value }
    if ($time.Kind -eq [DateTimeKind]::Unspecified) { $time = [datetime]::SpecifyKind($time, [DateTimeKind]::Utc) }
    return $time.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
}

function Format-PraGuiPath {
    <# A path under the tool folder as .\relative (shorter in the cards); any other path unchanged. #>
    param([AllowEmptyString()][string]$Path = '', [Parameter(Mandatory)][string]$Root)
    if (-not $Path) { return '' }
    try { $full = [IO.Path]::GetFullPath($Path); $base = [IO.Path]::GetFullPath($Root).TrimEnd('\') + '\' } catch { return $Path }
    if ($full.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) { return '.\' + $full.Substring($base.Length) }
    return $Path
}

function New-PraGuiTable {
    <#
        The rows of a table of the window: a DataTable, because thousands of rows must filter as you type
        (DataView.RowFilter) and the tick boxes bind both ways. Selected is a boolean, StatusBrush a brush, the rest text.
    #>
    param([Parameter(Mandatory)][string[]]$Columns)
    $table = [Data.DataTable]::new()
    foreach ($column in $Columns) {
        $type = switch ($column) { 'Selected' { [bool] } 'StatusBrush' { [object] } default { [string] } }
        [void]$table.Columns.Add($column, $type)
    }
    return , $table
}

function ConvertTo-PraGuiCheckTable {
    <#
        The rows of a Check report: Ready, Warning (ready, with a warning) or Not ready (Convert would refuse it), and
        the main reason. Brushes: a hashtable Tone -> brush (empty in the tests).
    #>
    param([AllowEmptyCollection()][object[]]$Rows = @(), [hashtable]$Brushes = @{})
    $table = New-PraGuiTable -Columns 'Identity', 'Kind', 'ObjectGuid', 'StatusText', 'Tone', 'StatusBrush', 'Licence', 'ExchangeOnline', 'Holds', 'Issue', 'Search'
    if (-not $Rows.Count) { return , $table }
    $names = @($Rows[0].PSObject.Properties | ForEach-Object { $_.Name })
    $columns = @('Identity', 'Kind', 'ObjectGuid', 'PrimarySmtpAddress', 'FinalStatus', 'Warnings', 'Detail', 'Licence', 'ExchangeOnline', 'Holds')
    $table.BeginLoadData()
    foreach ($row in $Rows) {
        $v = @{}
        foreach ($name in $columns) { $v[$name] = if ($names -contains $name) { [string]$row.$name } else { '' } }
        $isError = $v.FinalStatus -eq 'Error'
        $status = if ($isError) { 'Not ready' } elseif ($v.Warnings) { 'Warning' } else { 'Ready' }
        $tone = if ($isError) { 'Critical' } elseif ($v.Warnings) { 'Caution' } else { 'Success' }
        $issue = if ($isError -or -not $v.Warnings) { $v.Detail } else { $v.Warnings }
        if ($issue.Length -gt 260) { $issue = $issue.Substring(0, 257) + '...' }
        # A finding of the tenant or of the run, not of an object.
        if (-not $v.ObjectGuid -and -not $v.Identity) { $v.Identity = '(tenant)'; $issue = $issue -replace '^\[[^\]]+\]\s*', '' }
        $search = ('{0} {1} {2} {3}' -f $v.Identity, $v.PrimarySmtpAddress, $v.ObjectGuid, $issue).ToLowerInvariant()
        [void]$table.Rows.Add([object[]]@($v.Identity, $v.Kind, $v.ObjectGuid, $status, $tone, $Brushes[$tone], $v.Licence, $v.ExchangeOnline, $v.Holds, $issue, $search))
    }
    $table.EndLoadData()
    return , $table
}

function ConvertTo-PraGuiConvertTable {
    <#
        The objects of the last snapshot for the Convert page: their state in the last Check (Not checked without one)
        and their last Convert batch. InCloud = 1 when a batch converted the object (at least started it) and no Recover
        brought it back.
    #>
    param([AllowEmptyCollection()][object[]]$Objects = @(), [AllowNull()][Data.DataTable]$Check, [AllowEmptyCollection()][object[]]$Batches = @(),
        [hashtable]$Brushes = @{}, [AllowEmptyCollection()][string[]]$Ticked = @())
    $table = New-PraGuiTable -Columns 'Selected', 'Identity', 'Kind', 'ObjectGuid', 'StatusText', 'Tone', 'StatusBrush', 'Batch', 'InCloud', 'Issue', 'Search'
    $checks = @{}
    if ($Check) { foreach ($row in $Check.Rows) { if ($row['ObjectGuid']) { $checks[[string]$row['ObjectGuid']] = $row } } }
    $last = @{}
    foreach ($batch in $Batches) {
        foreach ($item in $batch.Items) {
            $guid = [string]$item.object_guid
            if ($last.ContainsKey($guid)) { continue }
            $started = [bool]$item.step -and $item.status -ne 'Skipped'
            $back = $batch.Back.Contains($guid)
            $text = if ($back) { '{0} · back' -f $batch.Id } elseif ($item.status -eq 'Done') { $batch.Id } elseif ($item.status -eq 'Planned') { '{0} · not started' -f $batch.Id }
                else { '{0} · {1}' -f $batch.Id, $item.status }
            $last[$guid] = @{ Text = $text; InCloud = $(if ($started -and -not $back) { '1' } else { '0' }) }
        }
    }
    $tick = New-Object 'Collections.Generic.HashSet[string]' ([string[]]@($Ticked), [StringComparer]::OrdinalIgnoreCase)
    $table.BeginLoadData()
    foreach ($object in $Objects) {
        $guid = [string]$object.object_guid
        $identity = if ($object.user_principal_name) { [string]$object.user_principal_name } else { [string]$object.primary_smtp_address }
        $status = 'Not checked'; $tone = 'Tertiary'; $issue = ''
        if ($checks.ContainsKey($guid)) { $c = $checks[$guid]; $status = [string]$c['StatusText']; $tone = [string]$c['Tone']; $issue = [string]$c['Issue'] }
        $converted = if ($last.ContainsKey($guid)) { $last[$guid] } else { @{ Text = ''; InCloud = '0' } }
        $search = ('{0} {1} {2} {3}' -f $identity, $object.primary_smtp_address, $object.sam_account_name, $guid).ToLowerInvariant()
        [void]$table.Rows.Add([object[]]@($tick.Contains($guid), $identity, [string]$object.kind, $guid, $status, $tone, $Brushes[$tone], $converted.Text, $converted.InCloud, $issue, $search))
    }
    $table.EndLoadData()
    return , $table
}

function ConvertTo-PraGuiRecoverTable {
    <#
        The objects of one Convert batch for the Recover page: Waiting (converted, not back yet), Back on-premises, or
        Nothing to roll back (never started or skipped).
    #>
    param([Parameter(Mandatory)][object]$Batch, [hashtable]$Brushes = @{}, [AllowEmptyCollection()][string[]]$Ticked = @())
    $table = New-PraGuiTable -Columns 'Selected', 'Identity', 'Kind', 'ObjectGuid', 'Converted', 'StatusText', 'Tone', 'StatusBrush', 'Search'
    $tick = New-Object 'Collections.Generic.HashSet[string]' ([string[]]@($Ticked), [StringComparer]::OrdinalIgnoreCase)
    $table.BeginLoadData()
    foreach ($item in $Batch.Items) {
        $guid = [string]$item.object_guid
        $converted = switch ([string]$item.status) {
            'Done' { 'Done' } 'Planned' { 'Not started' } 'Skipped' { 'Skipped' }
            default { if ($item.step) { '{0} at {1}' -f $item.status, $item.step } else { [string]$item.status } }
        }
        $status = 'Nothing to roll back'; $tone = 'Tertiary'
        if ($Batch.Back.Contains($guid)) { $status = 'Back on-premises'; $tone = 'Success' }
        elseif ($item.step -and $item.status -ne 'Skipped') { $status = 'Waiting'; $tone = 'Caution' }
        $search = ('{0} {1}' -f $item.identity, $guid).ToLowerInvariant()
        [void]$table.Rows.Add([object[]]@($tick.Contains($guid), [string]$item.identity, [string]$item.kind, $guid, $converted, $status, $tone, $Brushes[$tone], $search))
    }
    $table.EndLoadData()
    return , $table
}

function Get-PraGuiRowFilter {
    <#
        The DataView filter of a table: the text typed (in Search, as typed: wildcards and quotes are literal) and the
        clauses of the drop-down lists, joined with AND.
    #>
    param([AllowEmptyString()][string]$Text = '', [AllowEmptyCollection()][string[]]$Clauses = @())
    $parts = [Collections.Generic.List[string]]::new()
    $typed = $Text.Trim().ToLowerInvariant()
    if ($typed) {
        $literal = ($typed -replace '([\[\]\*%])', '[$1]').Replace("'", "''")
        $parts.Add(("Search LIKE '%{0}%'" -f $literal))
    }
    foreach ($clause in $Clauses) { if ($clause) { $parts.Add($clause) } }
    return ($parts -join ' AND ')
}

function Get-PraGuiTicked {
    <# The object GUIDs ticked in a table. #>
    param([AllowNull()][Data.DataTable]$Table)
    if (-not $Table) { return @() }
    return @($Table.Select('Selected = true') | ForEach-Object { [string]$_['ObjectGuid'] })
}

function Write-PraGuiWave {
    <# The objects ticked for a wave, as the CSV read by -IdentityPath (Identity = object GUID, Name for the reader). #>
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][object[]]$Rows)
    $null = New-Item -ItemType Directory -Path $Folder -Force
    $path = Join-Path $Folder ('wave-{0}-{1}.csv' -f $Name, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add('Identity,Name')
    foreach ($row in $Rows) { $lines.Add(('{0},"{1}"' -f $row.ObjectGuid, ([string]$row.Identity).Replace('"', "'"))) }
    [IO.File]::WriteAllLines($path, $lines, (New-Object Text.UTF8Encoding($true)))
    return $path
}

function Get-PraGuiSelectionKey {
    <# What a preview was run on: Apply is possible only on exactly the same selection. #>
    param([Parameter(Mandatory)][string]$Prefix, [AllowEmptyCollection()][string[]]$Guids = @())
    if (-not $Guids.Count) { return $Prefix }
    $sorted = @($Guids | Sort-Object) -join ','
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($sorted))).Replace('-', '').Substring(0, 16) } finally { $sha.Dispose() }
    return '{0}|{1}|{2}' -f $Prefix, $Guids.Count, $hash
}

function Get-PraGuiModuleVersion {
    <# The highest installed version of a module PowerShell 7 can load (AllUsers or CurrentUser), or ''. #>
    param([Parameter(Mandatory)][string]$Name)
    $folders = @((Join-Path $env:ProgramFiles "PowerShell\Modules\$Name"), (Join-Path ([Environment]::GetFolderPath('MyDocuments')) "PowerShell\Modules\$Name"),
        (Join-Path $env:ProgramFiles "WindowsPowerShell\Modules\$Name"), (Join-Path ([Environment]::GetFolderPath('MyDocuments')) "WindowsPowerShell\Modules\$Name"))
    $versions = foreach ($folder in $folders) {
        if (Test-Path -LiteralPath $folder) { Get-ChildItem -LiteralPath $folder -Directory -ErrorAction SilentlyContinue | ForEach-Object { $v = $null; if ([version]::TryParse($_.Name, [ref]$v)) { $v } } }
    }
    $top = @($versions | Sort-Object -Descending) | Select-Object -First 1
    if ($top) { return [string]$top }
    return ''
}
#endregion

#region 3. Window -----------------------------------------------------------------------------------

# Named elements of PRA2.Gui.xaml used by the module.
$script:ControlNames = @(
    'Root', 'EnvText', 'TenantText', 'VersionText', 'Nav', 'PageOverview', 'PageCollect', 'PageCheck', 'PageConvert', 'PageRecover',
    'NextCard', 'NextIcon', 'NextTitle', 'NextText', 'NextButton', 'OvConfig', 'OvBatches', 'OvSnapshot', 'OvComputer',
    'CollectFacts', 'CollectTask', 'CollectCopyTask', 'CollectPreview', 'CollectApply', 'CollectHint', 'SnapshotGrid',
    'CheckScope', 'CheckRun', 'CheckInfo', 'CheckTotal', 'CheckReady', 'CheckWarn', 'CheckNotReady', 'CheckFilter', 'CheckStatusFilter',
    'CheckKindFilter', 'CheckOpenReport', 'CheckCount', 'CheckGrid',
    'ConvertFacts', 'ConvertResumePanel', 'ConvertResumeText', 'ConvertResume', 'ConvertAll', 'ConvertWave', 'ConvertFilter',
    'ConvertKindFilter', 'ConvertStateFilter', 'ConvertTickShown', 'ConvertUntick', 'ConvertTicked', 'ConvertGrid', 'ConvertPreview',
    'ConvertApply', 'ConvertHint',
    'RecoverFacts', 'BatchGrid', 'RecoverAll', 'RecoverWave', 'RecoverFilter', 'RecoverTickShown', 'RecoverUntick', 'RecoverTicked',
    'ItemGrid', 'RecoverPreview', 'RecoverApply', 'RecoverHint',
    'StatusPill', 'StatusText', 'RunAction', 'RunMeta', 'RunStep', 'RunProgress', 'RunCurrent', 'RunObjectsBar', 'RunObjects', 'RunOk', 'RunWarn', 'RunFail',
    'ActivityEmpty', 'ActivityLog', 'ResultCard', 'ResultTitle', 'ResultText', 'BatchPanel', 'BatchText', 'CopyBatch', 'NextStepText',
    'OpenReport', 'OpenLog', 'StopRun', 'Footer', 'OpenConfig', 'OpenData', 'OpenLogs', 'RefreshState', 'CloseWindow')
# The drop-down lists: text shown -> value (scope of Check) or DataView clause (filters).
$script:Choices = @{
    CheckScope         = [ordered]@{ 'Every object' = 'All'; 'Users only' = 'UsersOnly'; 'Shared mailboxes only' = 'SharedOnly' }
    CheckStatusFilter  = [ordered]@{ 'Every status' = ''; 'Ready' = "StatusText = 'Ready'"; 'Ready, with a warning' = "StatusText = 'Warning'"; 'Not ready' = "StatusText = 'Not ready'" }
    KindFilter         = [ordered]@{ 'Every type' = ''; 'Users' = "Kind = 'User'"; 'Shared mailboxes' = "Kind = 'Shared'"; 'Rooms and equipment' = "Kind IN ('Room','Equipment')" }
    ConvertStateFilter = [ordered]@{ 'Every object' = ''; 'Not converted' = "InCloud = '0'"; 'Converted' = "InCloud = '1'"; 'Ready (last Check)' = "StatusText IN ('Ready','Warning')"
        'Not ready (last Check)' = "StatusText = 'Not ready'"; 'Not checked' = "StatusText = 'Not checked'" }
}
# Buttons that start a run: disabled while one runs.
$script:RunButtons = @('NextButton', 'CollectPreview', 'CollectApply', 'CheckRun', 'ConvertResume', 'ConvertPreview', 'ConvertApply', 'RecoverPreview', 'RecoverApply', 'RefreshState')

function Invoke-PraGuiSafe {
    <# Runs a handler of the window: an error is written in the activity panel, it never closes the window. #>
    param([Parameter(Mandatory)][scriptblock]$Action)
    try { & $Action }
    catch {
        $message = $_.Exception.Message
        try { Add-PraGuiLine -Status Fail -Text $message; Show-PraGuiLastLine } catch { Write-Warning $message }
    }
}

function New-PraGuiFact {
    <# One line of a card: label, value, and the tone of the value. #>
    param([Parameter(Mandatory)][string]$Label, [AllowEmptyString()][string]$Value = '', [string]$Tone = 'Primary')
    return [pscustomobject]@{ Label = $Label; Value = $(if ($Value) { $Value } else { '-' }); Brush = (Get-PraGuiBrush $Tone); Tone = $Tone }
}

function New-PraGuiNavItem {
    <# One entry of the navigation: icon, title, and what the page is for. #>
    param([Parameter(Mandatory)][object]$Page)
    $panel = [Windows.Controls.StackPanel]::new()
    $panel.Orientation = [Windows.Controls.Orientation]::Horizontal
    $icon = [Windows.Controls.TextBlock]::new()
    $icon.Text = [string][char]$Page.Glyph
    $icon.FontFamily = [Windows.Media.FontFamily]::new($script:IconFont)
    $icon.FontSize = 16; $icon.Width = 30; $icon.VerticalAlignment = [Windows.VerticalAlignment]::Center
    $icon.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'PraBrandText')
    $text = [Windows.Controls.StackPanel]::new()
    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = $Page.Title; $title.FontSize = 14; $title.FontWeight = [Windows.FontWeights]::SemiBold
    $title.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorPrimaryBrush')
    $detail = [Windows.Controls.TextBlock]::new()
    $detail.Text = $Page.Detail; $detail.FontSize = 11.5
    $detail.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorSecondaryBrush')
    [void]$text.Children.Add($title); [void]$text.Children.Add($detail)
    [void]$panel.Children.Add($icon); [void]$panel.Children.Add($text)
    $item = [Windows.Controls.ListBoxItem]::new()
    $item.Content = $panel
    $item.Tag = $Page.Name
    return $item
}

function Set-PraGuiTone {
    <# The text colour of an element, from a tone (follows the theme). #>
    param([Parameter(Mandatory)][Windows.Controls.TextBlock]$Element, [string]$Tone = 'Primary')
    $Element.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, (Get-PraGuiToneKey $Tone))
}

function Set-PraGuiHint {
    <# The line next to the Preview and Apply buttons of a page. #>
    param([Parameter(Mandatory)][ValidateSet('Collect', 'Convert', 'Recover')][string]$Page, [AllowEmptyString()][string]$Text, [string]$Tone = 'Secondary')
    $hint = $script:Gui.Controls[$Page + 'Hint']
    $hint.Text = $Text
    Set-PraGuiTone -Element $hint -Tone $Tone
}

function Get-PraGuiWorkFolder {
    <# The files of the window (runs, waves): the folder gui next to the database. #>
    $g = $script:Gui
    $data = if ($g.State -and $g.State.Config) { Split-Path -Parent $g.State.Config.Store.Path } else { Join-Path $g.Root 'data' }
    $folder = Join-Path $data 'gui'
    $null = New-Item -ItemType Directory -Path $folder -Force
    return $folder
}

function Open-PraGuiPath {
    <# A report or a log with its application, a folder in the Explorer, the configuration in Notepad. #>
    param([Parameter(Mandatory)][string]$Path, [switch]$Edit)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Not found: $Path" }
    if ($Edit) { Start-Process -FilePath 'notepad.exe' -ArgumentList ('"{0}"' -f $Path); return }
    if (Test-Path -LiteralPath $Path -PathType Container) { Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $Path); return }
    Start-Process -FilePath $Path
}

function Get-PraGuiCertificate {
    <# Where the certificate of the app is (CurrentUser\My or LocalMachine\My) and until when it is valid. #>
    param([AllowEmptyString()][string]$Thumbprint)
    if (-not $Thumbprint) { return [pscustomobject]@{ Text = 'not set'; Tone = 'Critical' } }
    foreach ($store in 'CurrentUser', 'LocalMachine') {
        $certificate = Get-Item -LiteralPath ('Cert:\{0}\My\{1}' -f $store, $Thumbprint) -ErrorAction SilentlyContinue
        if (-not $certificate) { continue }
        $days = ($certificate.NotAfter - (Get-Date)).TotalDays
        $tone = if (-not $certificate.HasPrivateKey -or $days -lt 0) { 'Critical' } elseif ($days -lt 30) { 'Caution' } else { 'Primary' }
        $text = '{0}\My, valid until {1:yyyy-MM-dd}{2}' -f $store, $certificate.NotAfter, $(if (-not $certificate.HasPrivateKey) { ', without its private key' } else { '' })
        return [pscustomobject]@{ Text = $text; Tone = $tone }
    }
    return [pscustomobject]@{ Text = "$Thumbprint not found in CurrentUser\My or LocalMachine\My"; Tone = 'Critical' }
}

function Get-PraGuiScopeText {
    <# The scope of the configuration, in words. #>
    param([Parameter(Mandatory)][hashtable]$Config)
    $s = $Config.Scope
    $base = switch ($s.Mode) { 'OU' { "OU $($s.SearchBase)" } 'Group' { "members of $($s.GroupDN)" } 'Csv' { "list $($s.CsvPath)" } default { if ($s.SearchBase) { $s.SearchBase } else { 'whole organisation' } } }
    $kinds = @(); if ($s.IncludeUsers) { $kinds += 'users' }; if ($s.IncludeShared) { $kinds += 'shared' }; if ($s.IncludeRoom) { $kinds += 'rooms' }; if ($s.IncludeEquipment) { $kinds += 'equipment' }
    return '{0} · {1}' -f $base, ($kinds -join ' + ')
}

function Get-PraGuiLicenceText {
    <# How the users get their Exchange plan. #>
    param([Parameter(Mandatory)][hashtable]$Config)
    $users = $Config.Licensing.Users
    switch ($users.Mode) {
        'Group' { if ($users.GroupId) { return "licence group $($users.GroupId)" } return 'licence group (GroupId not set)' }
        'Kiosk' { return 'Exchange Kiosk plan of the licence of each user' }
        default { return "$($users.SkuPartNumber), Exchange plan only" }
    }
}

function Get-PraGuiEntraConnectText {
    <# How Recover runs the synchronisation. #>
    param([Parameter(Mandatory)][hashtable]$Config, [string]$Root = '')
    $ec = $Config.EntraConnect
    switch ($ec.Mode) {
        'Remoting' { return "Remoting to $($ec.Server)" }
        'Script' { return 'script ' + $(if ($Root) { Format-PraGuiPath -Path $ec.ScriptPath -Root $Root } else { $ec.ScriptPath }) }
        default { return 'Manual: the window asks you to run each synchronisation' }
    }
}

function New-PraGuiWindow {
    <#
    .SYNOPSIS
        Builds the window and reads the state, without showing it: used by Show-PraGui, the tests and the
        documentation images. Returns Form, Controls and Lines.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ConfigPath, [string]$Version = '',
        [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')
    $look = Initialize-PraGuiTheme -Theme $Theme
    $window = [Windows.Markup.XamlReader]::Parse([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'PRA2.Gui.xaml')))
    Set-PraGuiTheme -Window $window -Theme $look
    $controls = @{}
    foreach ($name in $script:ControlNames) {
        $control = $window.FindName($name)
        if ($null -eq $control) { throw "PRA2.Gui.xaml: element '$name' not found." }
        $controls[$name] = $control
    }
    $lines = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
    $script:Gui = @{ Form = $window; Controls = $controls; Theme = $look; Root = $Root; ConfigPath = $ConfigPath; Version = $Version
        State = $null; Batches = @(); Tables = @{ Check = $null; Convert = $null; Recover = $null }; Brushes = @{}; SelectedBatch = ''
        Page = 'Overview'; Run = $null; InTick = $false; Lines = $lines; Pending = @{ Convert = $null; Recover = $null }; Updating = $false
        FilterPending = (New-Object 'Collections.Generic.HashSet[string]'); LastReport = ''; LastLog = ''
        Hooks = @{ Confirm = $null; Question = $null; Notice = $null }; Timer = $null; FilterTimer = $null }
    $g = $script:Gui
    foreach ($tone in 'Brand', 'Success', 'Caution', 'Critical', 'Primary', 'Secondary', 'Tertiary') { $g.Brushes[$tone] = Get-PraGuiBrush $tone }
    $window.Title = 'PRA Cloud Mailbox' + $(if ($Version) { " $Version" } else { '' })
    $controls.ActivityLog.ItemsSource = $lines
    foreach ($page in $script:Pages) { [void]$controls.Nav.Items.Add((New-PraGuiNavItem -Page $page)) }
    foreach ($pair in @(@('CheckScope', 'CheckScope'), @('CheckStatusFilter', 'CheckStatusFilter'), @('CheckKindFilter', 'KindFilter'),
            @('ConvertKindFilter', 'KindFilter'), @('ConvertStateFilter', 'ConvertStateFilter'))) {
        foreach ($text in $script:Choices[$pair[1]].Keys) { [void]$controls[$pair[0]].Items.Add($text) }
        $controls[$pair[0]].SelectedIndex = 0
    }
    foreach ($name in 'NextButton', 'CheckRun', 'CollectPreview', 'ConvertPreview', 'RecoverPreview') { Set-PraGuiAccentButton -Button $controls[$name] -Window $window -Theme $look }

    # The run is read every 400 ms; a filter applies 300 ms after the last key.
    $g.Timer = [Windows.Threading.DispatcherTimer]::new()
    $g.Timer.Interval = [TimeSpan]::FromMilliseconds(400)
    $g.Timer.Add_Tick({ Invoke-PraGuiSafe { Invoke-PraGuiTick } })
    $g.FilterTimer = [Windows.Threading.DispatcherTimer]::new()
    $g.FilterTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $g.FilterTimer.Add_Tick({
            Invoke-PraGuiSafe {
                $script:Gui.FilterTimer.Stop()
                $pages = @($script:Gui.FilterPending)
                $script:Gui.FilterPending.Clear()
                foreach ($page in $pages) { Update-PraGuiFilter -Page $page }
            }
        })

    $controls.Nav.Add_SelectionChanged({ Invoke-PraGuiSafe { $item = $script:Gui.Controls.Nav.SelectedItem; if ($item) { Select-PraGuiPage -Name ([string]$item.Tag) } } })
    $controls.NextButton.Add_Click({
            Invoke-PraGuiSafe {
                $page = [string]$script:Gui.Controls.NextButton.Tag
                if ($page) { Select-PraGuiPage -Name $page } else { Open-PraGuiPath -Path $script:Gui.ConfigPath -Edit }
            }
        })
    $controls.CollectCopyTask.Add_Click({ Invoke-PraGuiSafe { [Windows.Clipboard]::SetText([string]$script:Gui.Controls.CollectTask.Text) } })
    $controls.CollectPreview.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiRun -Request @{ Action = 'Collect'; Mode = 'Preview' } -Page 'Collect' } })
    $controls.CollectApply.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiCollect } })
    $controls.CheckRun.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiCheck } })
    $controls.CheckOpenReport.Add_Click({ Invoke-PraGuiSafe { if ($script:Gui.State.Check -and $script:Gui.State.Check.Html) { Open-PraGuiPath -Path $script:Gui.State.Check.Html } } })
    $controls.CheckFilter.Add_TextChanged({ Invoke-PraGuiSafe { Request-PraGuiFilter -Page 'Check' } })
    $controls.ConvertFilter.Add_TextChanged({ Invoke-PraGuiSafe { Request-PraGuiFilter -Page 'Convert' } })
    $controls.RecoverFilter.Add_TextChanged({ Invoke-PraGuiSafe { Request-PraGuiFilter -Page 'Recover' } })
    $controls.CheckStatusFilter.Add_SelectionChanged({ Invoke-PraGuiSafe { Update-PraGuiFilter -Page 'Check' } })
    $controls.CheckKindFilter.Add_SelectionChanged({ Invoke-PraGuiSafe { Update-PraGuiFilter -Page 'Check' } })
    $controls.ConvertKindFilter.Add_SelectionChanged({ Invoke-PraGuiSafe { Update-PraGuiFilter -Page 'Convert' } })
    $controls.ConvertStateFilter.Add_SelectionChanged({ Invoke-PraGuiSafe { Update-PraGuiFilter -Page 'Convert' } })
    $controls.ConvertAll.Add_Checked({ Invoke-PraGuiSafe { Update-PraGuiApplyState } })
    $controls.ConvertWave.Add_Checked({ Invoke-PraGuiSafe { Update-PraGuiApplyState } })
    $controls.RecoverAll.Add_Checked({ Invoke-PraGuiSafe { Update-PraGuiApplyState } })
    $controls.RecoverWave.Add_Checked({ Invoke-PraGuiSafe { Update-PraGuiApplyState } })
    # A tick box of a row (the column headers are buttons too: only the boxes count).
    $controls.ConvertGrid.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] {
            if ($args[1].OriginalSource -is [Windows.Controls.CheckBox]) { Invoke-PraGuiSafe { Update-PraGuiTicks -Page 'Convert' -Changed } }
        })
    $controls.ItemGrid.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] {
            if ($args[1].OriginalSource -is [Windows.Controls.CheckBox]) { Invoke-PraGuiSafe { Update-PraGuiTicks -Page 'Recover' -Changed } }
        })
    $controls.ConvertTickShown.Add_Click({ Invoke-PraGuiSafe { Set-PraGuiTicks -Page 'Convert' -Shown } })
    $controls.ConvertUntick.Add_Click({ Invoke-PraGuiSafe { Set-PraGuiTicks -Page 'Convert' } })
    $controls.RecoverTickShown.Add_Click({ Invoke-PraGuiSafe { Set-PraGuiTicks -Page 'Recover' -Shown } })
    $controls.RecoverUntick.Add_Click({ Invoke-PraGuiSafe { Set-PraGuiTicks -Page 'Recover' } })
    $controls.ConvertResume.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiPreview -Page 'Convert' -Resume } })
    $controls.ConvertPreview.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiPreview -Page 'Convert' } })
    $controls.ConvertApply.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiApply -Page 'Convert' } })
    $controls.RecoverPreview.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiPreview -Page 'Recover' } })
    $controls.RecoverApply.Add_Click({ Invoke-PraGuiSafe { Start-PraGuiApply -Page 'Recover' } })
    $controls.BatchGrid.Add_SelectionChanged({
            Invoke-PraGuiSafe {
                if ($script:Gui.Updating) { return }
                $item = $script:Gui.Controls.BatchGrid.SelectedItem
                if ($item) { Select-PraGuiBatch -Id ([string]$item.Id) }
            }
        })
    $controls.StopRun.Add_Click({ Invoke-PraGuiSafe { Stop-PraGuiRun } })
    $controls.CopyBatch.Add_Click({ Invoke-PraGuiSafe { [Windows.Clipboard]::SetText([string]$script:Gui.Controls.BatchText.Text) } })
    $controls.OpenReport.Add_Click({ Invoke-PraGuiSafe { if ($script:Gui.LastReport) { Open-PraGuiPath -Path $script:Gui.LastReport } } })
    $controls.OpenLog.Add_Click({ Invoke-PraGuiSafe { if ($script:Gui.LastLog) { Open-PraGuiPath -Path $script:Gui.LastLog } } })
    $controls.OpenConfig.Add_Click({ Invoke-PraGuiSafe { Open-PraGuiPath -Path $script:Gui.ConfigPath -Edit } })
    $controls.OpenData.Add_Click({ Invoke-PraGuiSafe { Open-PraGuiPath -Path (Split-Path -Parent (Get-PraGuiWorkFolder)) } })
    $controls.OpenLogs.Add_Click({
            Invoke-PraGuiSafe {
                $folder = if ($script:Gui.State.Config -and (Test-Path -LiteralPath $script:Gui.State.Config.Logging.Folder)) { $script:Gui.State.Config.Logging.Folder } else { $script:Gui.Root }
                Open-PraGuiPath -Path $folder
            }
        })
    $controls.RefreshState.Add_Click({ Invoke-PraGuiSafe { Update-PraGuiState; Add-PraGuiLine -Status Info -Text 'State read again (configuration, snapshot, journal, last Check).'; Show-PraGuiLastLine } })
    $controls.CloseWindow.Add_Click({ $script:Gui.Form.Close() })
    $window.Add_Closing({
            if ($script:Gui -and $script:Gui.Run) {
                $args[1].Cancel = $true
                Show-PraGuiNotice -Text 'An action is running. Wait for its end, or click "Stop after the current object" and wait: closing the window now would leave the run without anyone to follow it.'
            }
        })

    $area = [Windows.SystemParameters]::WorkArea
    $window.Width = [Math]::Min($window.Width, $area.Width)
    $window.Height = [Math]::Min($window.Height, $area.Height)
    $window.MinWidth = [Math]::Min($window.MinWidth, $area.Width)
    $window.MinHeight = [Math]::Min($window.MinHeight, $area.Height)

    Clear-PraGuiActivity
    $controls.RunAction.Text = 'Nothing running'
    $controls.RunMeta.Text = 'Each action runs Invoke-PraCloudMailbox.ps1 in its own PowerShell, with its log, transcript and report.'
    Set-PraGuiStatus -Text 'Ready' -Tone 'Secondary'
    Set-PraGuiBusy -Busy $false
    Update-PraGuiState
    $controls.Nav.SelectedIndex = 0
    return [pscustomobject]@{ Form = $window; Controls = $controls; Lines = $lines }
}

function Select-PraGuiPage {
    <# Shows one page and selects it in the navigation. #>
    param([Parameter(Mandatory)][ValidateSet('Overview', 'Collect', 'Check', 'Convert', 'Recover')][string]$Name)
    $g = $script:Gui; $c = $g.Controls
    foreach ($page in $script:Pages) {
        $c['Page' + $page.Name].Visibility = if ($page.Name -eq $Name) { [Windows.Visibility]::Visible } else { [Windows.Visibility]::Collapsed }
    }
    $g.Page = $Name
    $index = [Array]::IndexOf([string[]]@($script:Pages | ForEach-Object { $_.Name }), $Name)
    if ($c.Nav.SelectedIndex -ne $index) { $c.Nav.SelectedIndex = $index }
}

function Update-PraGuiState {
    <#
        Reads the state again and fills every page. The ticks are kept (the objects stay ticked), except on the pages
        named in ClearTicks (after their Apply).
    #>
    param([AllowEmptyCollection()][string[]]$ClearTicks = @())
    $g = $script:Gui
    $ticked = @{}
    foreach ($page in 'Convert', 'Recover') { $ticked[$page] = [string[]]@(if ($ClearTicks -notcontains $page) { Get-PraGuiTicked $g.Tables[$page] }) }
    $g.State = Read-PraGuiState -Root $g.Root -ConfigPath $g.ConfigPath
    $g.Batches = @(if (-not $g.State.ConfigError) { Get-PraGuiBatchSummary -State $g.State })
    Update-PraGuiHeader
    Update-PraGuiOverview
    Update-PraGuiCollectPage
    Update-PraGuiCheckPage
    Update-PraGuiConvertPage -Ticked $ticked.Convert
    Update-PraGuiRecoverPage -Ticked $ticked.Recover
    Update-PraGuiApplyState
    $data = if ($g.State.Config) { Format-PraGuiPath -Path (Split-Path -Parent $g.State.Config.Store.Path) -Root $g.Root } else { '-' }
    $g.Controls.Footer.Text = '{0}  ·  configuration {1}  ·  data {2}  ·  read at {3:HH:mm:ss}' -f $g.Root, (Format-PraGuiPath -Path $g.ConfigPath -Root $g.Root), $data, $g.State.Read
}

function Update-PraGuiHeader {
    $g = $script:Gui; $c = $g.Controls; $config = $g.State.Config
    if ($g.State.ConfigError) {
        $c.EnvText.Text = 'Configuration error'
        $c.TenantText.Text = $g.ConfigPath
    } else {
        $c.EnvText.Text = 'Environment ' + $config.Environment
        $c.TenantText.Text = if ($config.Cloud.Organization) { $config.Cloud.Organization } elseif ($config.Cloud.TenantId) { $config.Cloud.TenantId } else { 'tenant not set' }
    }
    $edition = if ($PSVersionTable.PSEdition -eq 'Core') { 'PowerShell' } else { 'Windows PowerShell' }
    $c.VersionText.Text = '{0}Nicolas Fabert · {1} {2} · {3}' -f $(if ($g.Version) { "v$($g.Version) · " } else { '' }), $edition, $PSVersionTable.PSVersion, $(if ($g.Theme.Fluent) { 'Fluent' } else { 'classic theme' })
}

function Update-PraGuiOverview {
    $g = $script:Gui; $c = $g.Controls; $state = $g.State; $config = $state.Config
    $next = Get-PraGuiNextStep -State $state -Batches $g.Batches
    $c.NextTitle.Text = $next.Title
    $c.NextText.Text = $next.Text
    $c.NextIcon.Text = [string][char]$next.Glyph
    $c.NextButton.Content = if ($next.Page) { 'Go to ' + $next.Button } else { $next.Button }
    $c.NextButton.Tag = $next.Page

    $computer = [Collections.Generic.List[object]]::new()
    $computer.Add((New-PraGuiFact 'Computer' $env:COMPUTERNAME))
    $computer.Add((New-PraGuiFact 'Account' ([Security.Principal.WindowsIdentity]::GetCurrent().Name)))
    $pwsh = Get-PraGuiEngine -Action Check
    $computer.Add($(if ($pwsh.Ok) { New-PraGuiFact 'PowerShell 7' ('{0} · Check, Convert, Recover' -f $pwsh.Version) } else { New-PraGuiFact 'PowerShell 7' $pwsh.Problem 'Critical' }))
    $graph = Get-PraGuiModuleVersion -Name 'Microsoft.Graph.Authentication'
    $exo = Get-PraGuiModuleVersion -Name 'ExchangeOnlineManagement'
    $computer.Add($(if ($graph) { New-PraGuiFact 'Graph module' "Microsoft.Graph.Authentication $graph" } else { New-PraGuiFact 'Graph module' 'Microsoft.Graph.Authentication not installed (Install-Module Microsoft.Graph.Authentication)' 'Critical' }))
    $computer.Add($(if ($exo) { New-PraGuiFact 'Exchange Online' "ExchangeOnlineManagement $exo" } else { New-PraGuiFact 'Exchange Online' 'ExchangeOnlineManagement not installed (Install-Module ExchangeOnlineManagement)' 'Critical' }))
    $computer.Add($(if ($env:ExchangeInstallPath) { New-PraGuiFact 'Exchange tools' 'installed: Collect can run here' } else { New-PraGuiFact 'Exchange tools' 'not on this computer: Collect runs on an Exchange server' 'Secondary' }))
    $c.OvComputer.ItemsSource = $computer.ToArray()

    if ($state.ConfigError) {
        $c.OvConfig.ItemsSource = @((New-PraGuiFact 'File' $g.ConfigPath), (New-PraGuiFact 'Error' $state.ConfigError 'Critical'))
        $c.OvSnapshot.ItemsSource = @(New-PraGuiFact 'Snapshot' 'unknown until the configuration is fixed' 'Secondary')
        $c.OvBatches.ItemsSource = @(New-PraGuiFact 'Journal' 'unknown until the configuration is fixed' 'Secondary')
        return
    }
    $facts = [Collections.Generic.List[object]]::new()
    $facts.Add((New-PraGuiFact 'File' (Format-PraGuiPath -Path $config._Path -Root $g.Root)))
    $facts.Add((New-PraGuiFact 'Environment' $config.Environment))
    $tenant = (@($config.Cloud.Organization, $config.Cloud.TenantId) | Where-Object { $_ }) -join ' · '
    $facts.Add($(if ($tenant) { New-PraGuiFact 'Tenant' $tenant } else { New-PraGuiFact 'Tenant' 'not set' 'Critical' }))
    $facts.Add($(if ($config.Cloud.AppId) { New-PraGuiFact 'App' $config.Cloud.AppId } else { New-PraGuiFact 'App' 'not set' 'Critical' }))
    $certificate = Get-PraGuiCertificate -Thumbprint ([string]$config.Cloud.CertificateThumbprint)
    $facts.Add((New-PraGuiFact 'Certificate' $certificate.Text $certificate.Tone))
    $facts.Add((New-PraGuiFact 'Scope' (Get-PraGuiScopeText -Config $config)))
    $facts.Add((New-PraGuiFact 'User licence' (Get-PraGuiLicenceText -Config $config)))
    $facts.Add($(if ($config.Licensing.Shared.SkuPartNumber) { New-PraGuiFact 'Shared mailboxes' "temporary $($config.Licensing.Shared.SkuPartNumber)" } else { New-PraGuiFact 'Shared mailboxes' 'temporary licence not set (Licensing.Shared.SkuPartNumber)' 'Caution' }))
    $facts.Add($(if ($config.Retention.HoldPolicy) { New-PraGuiFact 'Case hold' $config.Retention.HoldPolicy } else { New-PraGuiFact 'Case hold' 'not set (Retention.HoldPolicy): Recover needs it' 'Caution' }))
    $facts.Add((New-PraGuiFact 'Entra Connect' (Get-PraGuiEntraConnectText -Config $config -Root $g.Root)))
    $c.OvConfig.ItemsSource = $facts.ToArray()

    $snapshot = [Collections.Generic.List[object]]::new()
    if ($state.SnapshotError) { $snapshot.Add((New-PraGuiFact 'Database' $state.SnapshotError 'Critical')) }
    elseif (-not $state.Snapshot) { $snapshot.Add((New-PraGuiFact 'Snapshot' 'none: Collect on an Exchange server, then copy the database here' 'Caution')) }
    else {
        $s = $state.Snapshot
        $taken = Format-PraGuiTime ([string]$s.finished_utc)
        $age = 0.0
        try { $age = ((Get-Date).ToUniversalTime() - [datetime]::Parse([string]$s.finished_utc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()).TotalDays } catch { $age = 0.0 }
        $snapshot.Add((New-PraGuiFact 'Snapshot' ('#{0} · {1} ({2:0} day(s) ago)' -f $s.id, $taken, [Math]::Floor($age)) $(if ($age -gt $config.Store.MaxSnapshotAgeDays) { 'Caution' } else { 'Primary' })))
        $snapshot.Add((New-PraGuiFact 'Exchange server' ([string]$s.exchange_server)))
        $kinds = @($state.Objects | Group-Object -Property kind)
        $count = { param($kind) $group = @($kinds | Where-Object { $_.Name -eq $kind }); if ($group.Count) { $group[0].Count } else { 0 } }
        $other = (& $count 'Room') + (& $count 'Equipment')
        $snapshot.Add((New-PraGuiFact 'Objects' ('{0} user(s) · {1} shared{2}' -f (& $count 'User'), (& $count 'Shared'), $(if ($other) { " · $other room/equipment (not converted)" } else { '' }))))
        $snapshot.Add((New-PraGuiFact 'Permissions' ('{0} (shared mailboxes)' -f $state.Permissions)))
        $snapshot.Add((New-PraGuiFact 'Database' (Format-PraGuiPath -Path $config.Store.Path -Root $g.Root)))
    }
    if ($state.Check) {
        $table = $g.Tables.Check
        $text = 'last on {0:yyyy-MM-dd HH:mm}' -f $state.Check.Time
        $tone = 'Primary'
        if ($table) {
            $notReady = $table.Select("StatusText = 'Not ready'").Count
            $text += ' · {0} ready · {1} with a warning · {2} not ready' -f $table.Select("StatusText = 'Ready'").Count, $table.Select("StatusText = 'Warning'").Count, $notReady
            if ($notReady) { $tone = 'Critical' }
        }
        $snapshot.Add((New-PraGuiFact 'Check' $text $tone))
    } else { $snapshot.Add((New-PraGuiFact 'Check' 'never run (or its report is not in the report folder)' 'Caution')) }
    $c.OvSnapshot.ItemsSource = $snapshot.ToArray()

    $batches = [Collections.Generic.List[object]]::new()
    if ($state.JournalError) { $batches.Add((New-PraGuiFact 'Journal' $state.JournalError 'Critical')) }
    elseif (-not $g.Batches.Count) { $batches.Add((New-PraGuiFact 'Journal' 'no batch: nothing converted' 'Secondary')) }
    else {
        foreach ($b in @($g.Batches | Select-Object -First 8)) {
            $batches.Add((New-PraGuiFact ('Batch ' + $b.Id) ('{0} · {1} user(s), {2} shared · Convert {3} · Recover: {4}' -f $b.Created, $b.Users, $b.Shared, $b.Status, $b.Recover) $(if ($b.Status -in @('Partial', 'Running', 'Failed')) { 'Caution' } else { $b.Tone })))
        }
        if ($g.Batches.Count -gt 8) { $batches.Add((New-PraGuiFact '...' ('{0} older batch(es) on the Recover page' -f ($g.Batches.Count - 8)) 'Secondary')) }
    }
    $c.OvBatches.ItemsSource = $batches.ToArray()
}

function Update-PraGuiCollectPage {
    $g = $script:Gui; $c = $g.Controls; $state = $g.State; $config = $state.Config
    $task = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Action Collect -Mode Apply -Force' -f (Join-Path $g.Root 'Invoke-PraCloudMailbox.ps1')
    $default = Join-Path $g.Root 'config\PraCloudMailbox.config.psd1'
    if (-not [string]::Equals([IO.Path]::GetFullPath($g.ConfigPath), [IO.Path]::GetFullPath($default), [StringComparison]::OrdinalIgnoreCase)) { $task += ' -ConfigPath "{0}"' -f $g.ConfigPath }
    $c.CollectTask.Text = $task
    $facts = [Collections.Generic.List[object]]::new()
    if ($config) {
        $facts.Add((New-PraGuiFact 'Exchange' $(if ($config.Exchange.ConnectionMode -eq 'Remote') { "PowerShell remoting to $($config.Exchange.Server)" } else { 'this server (Exchange Management Shell)' })))
        $facts.Add((New-PraGuiFact 'Scope' (Get-PraGuiScopeText -Config $config)))
        $facts.Add((New-PraGuiFact 'Database' ('{0} · {1} snapshot(s) kept' -f (Format-PraGuiPath -Path $config.Store.Path -Root $g.Root), $config.Store.KeepSnapshots)))
    }
    $facts.Add($(if ($state.Snapshot) { New-PraGuiFact 'Last snapshot' ('#{0} · {1}' -f $state.Snapshot.id, (Format-PraGuiTime ([string]$state.Snapshot.finished_utc))) } else { New-PraGuiFact 'Last snapshot' 'none' 'Caution' }))
    $c.CollectFacts.ItemsSource = $facts.ToArray()
    $rows = foreach ($s in $state.Snapshots) {
        $counts = $null
        try { if ($s.counts_json) { $counts = [string]$s.counts_json | ConvertFrom-Json } } catch { $counts = $null }
        [pscustomobject]@{ Id = [string]$s.id; Taken = (Format-PraGuiTime ([string]$(if ($s.finished_utc) { $s.finished_utc } else { $s.started_utc }))); Status = [string]$s.status
            Server = [string]$s.exchange_server; Computer = [string]$s.computer
            Mailboxes = $(if ($counts -and $counts.PSObject.Properties['mailbox']) { [string]$counts.mailbox } else { '' })
            Permissions = $(if ($counts -and $counts.PSObject.Properties['permission']) { [string]$counts.permission } else { '' }) }
    }
    $c.SnapshotGrid.ItemsSource = @($rows)
    if ($config -and $config.Exchange.ConnectionMode -eq 'Local' -and -not $env:ExchangeInstallPath) {
        Set-PraGuiHint -Page Collect -Text 'The Exchange tools are not on this computer: Collect runs on an Exchange server (the scheduled task above), then the database is copied here.' -Tone Caution
    } else { Set-PraGuiHint -Page Collect -Text 'Preview reads Exchange and writes nothing; Collect now writes a new snapshot. Exchange is never changed.' }
}

function Update-PraGuiCheckPage {
    $g = $script:Gui; $c = $g.Controls; $check = $g.State.Check
    $table = ConvertTo-PraGuiCheckTable -Rows @(if ($check) { $check.Rows }) -Brushes $g.Brushes
    $g.Tables.Check = $table
    $c.CheckGrid.ItemsSource = $table.DefaultView
    $objects = "ObjectGuid <> ''"
    if ($check) {
        $c.CheckTotal.Text = [string]$table.Select($objects).Count
        $c.CheckReady.Text = [string]$table.Select("$objects AND StatusText = 'Ready'").Count
        $c.CheckWarn.Text = [string]$table.Select("$objects AND StatusText = 'Warning'").Count
        $c.CheckNotReady.Text = [string]$table.Select("StatusText = 'Not ready'").Count
        $c.CheckInfo.Text = 'Last Check: {0:yyyy-MM-dd HH:mm} · {1} object(s) · {2}' -f $check.Time, $table.Select($objects).Count, (Split-Path -Leaf $check.Path)
    } else {
        foreach ($name in 'CheckTotal', 'CheckReady', 'CheckWarn', 'CheckNotReady') { $c[$name].Text = '-' }
        $c.CheckInfo.Text = 'No Check report yet' + $(if ($g.State.Config) { ' in ' + (Format-PraGuiPath -Path $g.State.Config.Report.Folder -Root $g.Root) } else { '' })
    }
    $c.CheckOpenReport.IsEnabled = [bool]($check -and $check.Html)
    Update-PraGuiFilter -Page Check
}

function Update-PraGuiConvertPage {
    param([AllowEmptyCollection()][string[]]$Ticked = @())
    $g = $script:Gui; $c = $g.Controls; $state = $g.State; $config = $state.Config
    $table = ConvertTo-PraGuiConvertTable -Objects $state.Objects -Check $g.Tables.Check -Batches $g.Batches -Brushes $g.Brushes -Ticked $Ticked
    $g.Tables.Convert = $table
    $c.ConvertGrid.ItemsSource = $table.DefaultView
    $facts = [Collections.Generic.List[object]]::new()
    if (-not $state.Snapshot) { $facts.Add((New-PraGuiFact 'Snapshot' 'none: Collect first' 'Critical')) }
    else {
        $facts.Add((New-PraGuiFact 'Snapshot' ('#{0} · {1}' -f $state.Snapshot.id, (Format-PraGuiTime ([string]$state.Snapshot.finished_utc)))))
        $other = $table.Select("Kind IN ('Room','Equipment')").Count
        $facts.Add((New-PraGuiFact 'Objects' ('{0} user(s) · {1} shared mailbox(es){2}' -f $table.Select("Kind = 'User'").Count, $table.Select("Kind = 'Shared'").Count, $(if ($other) { " · $other room/equipment (refused by Convert)" } else { '' }))))
    }
    if ($config) {
        $facts.Add((New-PraGuiFact 'Users' ('identity to the cloud, Exchange plan by ' + (Get-PraGuiLicenceText -Config $config))))
        $facts.Add((New-PraGuiFact 'Shared mailboxes' $(if ($config.Licensing.Shared.SkuPartNumber) { 'waves of {0}, temporary {1}' -f $config.Licensing.Shared.Parallel, $config.Licensing.Shared.SkuPartNumber } else { 'temporary licence not set' }) $(if ($config.Licensing.Shared.SkuPartNumber) { 'Primary' } else { 'Caution' })))
    }
    $inCloud = $table.Select("InCloud = '1'").Count
    $facts.Add((New-PraGuiFact 'In Exchange Online' $(if ($inCloud) { '{0} object(s) converted, not recovered yet' -f $inCloud } else { 'none' })))
    $facts.Add($(if ($state.Check) { New-PraGuiFact 'Last Check' ('{0:yyyy-MM-dd HH:mm}' -f $state.Check.Time) } else { New-PraGuiFact 'Last Check' 'none: run Check first to know which objects are ready' 'Caution' }))
    $c.ConvertFacts.ItemsSource = $facts.ToArray()
    $partial = @($g.Batches | Where-Object { $_.Status -in @('Partial', 'Running') }) | Select-Object -First 1
    if ($partial) {
        $open = @($partial.Items | Where-Object { $_.status -notin @('Done', 'Skipped') }).Count
        $c.ConvertResumeText.Text = 'Batch {0} ({1}) is not finished: {2} object(s) not done (stopped, pending or failed). Its resume takes again only what is not done, with the same batch ID.' -f $partial.Id, $partial.Created, $open
        $c.ConvertResumePanel.Visibility = [Windows.Visibility]::Visible
    } else { $c.ConvertResumePanel.Visibility = [Windows.Visibility]::Collapsed }
    Update-PraGuiFilter -Page Convert
}

function Update-PraGuiRecoverPage {
    param([AllowEmptyCollection()][string[]]$Ticked = @())
    $g = $script:Gui; $c = $g.Controls; $config = $g.State.Config
    $rows = @(foreach ($b in $g.Batches) {
            [pscustomobject]@{ Id = $b.Id; Created = $b.Created; Users = $b.Users; Shared = $b.Shared; ConvertStatus = $b.Status; StatusText = $b.Recover; StatusBrush = $g.Brushes[$b.Tone]; Waiting = $b.Waiting }
        })
    $pick = @($rows | Where-Object { $_.Id -eq $g.SelectedBatch }) + @($rows | Where-Object { $_.Waiting -gt 0 }) + @($rows) | Select-Object -First 1
    $g.Updating = $true
    try {
        $c.BatchGrid.ItemsSource = $rows
        $c.BatchGrid.SelectedItem = $pick
    } finally { $g.Updating = $false }
    $facts = [Collections.Generic.List[object]]::new()
    if ($config) {
        $facts.Add((New-PraGuiFact 'Entra Connect' (Get-PraGuiEntraConnectText -Config $config -Root $g.Root)))
        $facts.Add($(if ($config.Retention.HoldPolicy) { New-PraGuiFact 'Case hold' $config.Retention.HoldPolicy } else { New-PraGuiFact 'Case hold' 'not set (Retention.HoldPolicy)' 'Critical' }))
    }
    $waiting = 0
    foreach ($b in $g.Batches) { $waiting += [int]$b.Waiting }
    $facts.Add((New-PraGuiFact 'Waiting' $(if ($waiting) { '{0} converted object(s) not back yet' -f $waiting } else { 'nothing to roll back' }) $(if ($waiting) { 'Caution' } else { 'Secondary' })))
    $c.RecoverFacts.ItemsSource = $facts.ToArray()
    Select-PraGuiBatch -Id $(if ($pick) { [string]$pick.Id } else { '' }) -Ticked $Ticked
}

function Select-PraGuiBatch {
    <# The objects of one Convert batch in the Recover table. #>
    param([AllowEmptyString()][string]$Id = '', [AllowEmptyCollection()][string[]]$Ticked = @())
    $g = $script:Gui; $c = $g.Controls
    $batch = @($g.Batches | Where-Object { $_.Id -eq $Id }) | Select-Object -First 1
    if ($Id -ne $g.SelectedBatch) { $Ticked = @() }
    $g.SelectedBatch = if ($batch) { $Id } else { '' }
    if (-not $batch) {
        $g.Tables.Recover = $null
        $c.ItemGrid.ItemsSource = $null
    } else {
        $table = ConvertTo-PraGuiRecoverTable -Batch $batch -Brushes $g.Brushes -Ticked $Ticked
        $g.Tables.Recover = $table
        $c.ItemGrid.ItemsSource = $table.DefaultView
    }
    Update-PraGuiFilter -Page Recover
    Update-PraGuiApplyState
}

function Request-PraGuiFilter {
    <# A key typed in a filter: the filter applies when the typing pauses. #>
    param([Parameter(Mandatory)][ValidateSet('Check', 'Convert', 'Recover')][string]$Page)
    $g = $script:Gui
    [void]$g.FilterPending.Add($Page)
    $g.FilterTimer.Stop()
    $g.FilterTimer.Start()
}

function Update-PraGuiFilter {
    <# Applies the filter text and the drop-down lists of a table. #>
    param([Parameter(Mandatory)][ValidateSet('Check', 'Convert', 'Recover')][string]$Page)
    $g = $script:Gui; $c = $g.Controls; $table = $g.Tables[$Page]
    if ($table) {
        $clauses = switch ($Page) {
            'Check' { $script:Choices.CheckStatusFilter[[string]$c.CheckStatusFilter.SelectedItem]; $script:Choices.KindFilter[[string]$c.CheckKindFilter.SelectedItem] }
            'Convert' { $script:Choices.KindFilter[[string]$c.ConvertKindFilter.SelectedItem]; $script:Choices.ConvertStateFilter[[string]$c.ConvertStateFilter.SelectedItem] }
            default { '' }
        }
        $table.DefaultView.RowFilter = Get-PraGuiRowFilter -Text ([string]$c[$Page + 'Filter'].Text) -Clauses ([string[]]@($clauses | Where-Object { $_ }))
    }
    Update-PraGuiCount -Page $Page
}

function Update-PraGuiCount {
    <# Rows shown, and rows ticked. #>
    param([Parameter(Mandatory)][ValidateSet('Check', 'Convert', 'Recover')][string]$Page)
    $g = $script:Gui; $c = $g.Controls; $table = $g.Tables[$Page]
    $shown = if ($table) { $table.DefaultView.Count } else { 0 }
    $total = if ($table) { $table.Rows.Count } else { 0 }
    if ($Page -eq 'Check') { $c.CheckCount.Text = '{0} of {1} shown' -f $shown, $total; return }
    $ticked = @(Get-PraGuiTicked $table).Count
    $c[$Page + 'Ticked'].Text = '{0} ticked · {1} of {2} shown' -f $ticked, $shown, $total
}

function Update-PraGuiTicks {
    <# After a box was ticked or unticked: the counts; ticking an object chooses the wave. #>
    param([Parameter(Mandatory)][ValidateSet('Convert', 'Recover')][string]$Page, [switch]$Changed)
    $g = $script:Gui; $c = $g.Controls
    if ($Changed -and @(Get-PraGuiTicked $g.Tables[$Page]).Count -and -not $c[$Page + 'Wave'].IsChecked) { $c[$Page + 'Wave'].IsChecked = $true }
    Update-PraGuiCount -Page $Page
    Update-PraGuiApplyState
}

function Set-PraGuiTicks {
    <#
        Shown: ticks the rows shown that can go in a wave (Convert: ready in the last Check and not converted;
        Recover: waiting). Otherwise unticks every row.
    #>
    param([Parameter(Mandatory)][ValidateSet('Convert', 'Recover')][string]$Page, [switch]$Shown)
    $g = $script:Gui; $c = $g.Controls; $table = $g.Tables[$Page]
    if (-not $table) { return }
    $grid = $c[$(if ($Page -eq 'Convert') { 'ConvertGrid' } else { 'ItemGrid' })]
    $view = $table.DefaultView
    $rows = [Collections.Generic.List[Data.DataRow]]::new()
    if ($Shown) {
        foreach ($rowView in $view) {
            $row = $rowView.Row
            $fits = if ($Page -eq 'Convert') { $row['InCloud'] -eq '0' -and $row['StatusText'] -in @('Ready', 'Warning') } else { $row['StatusText'] -eq 'Waiting' }
            if ($fits -and -not $row['Selected']) { $rows.Add($row) }
        }
    } else { foreach ($row in $table.Rows) { if ($row['Selected']) { $rows.Add($row) } } }
    # Thousands of rows: the table is detached from the grid while it changes.
    $grid.ItemsSource = $null
    try { foreach ($row in $rows) { $row['Selected'] = [bool]$Shown } }
    finally { $grid.ItemsSource = $view }
    if ($Shown -and -not $rows.Count) {
        $text = if ($Page -eq 'Convert') { 'No row shown is ready and not converted' + $(if (-not $g.State.Check) { ': run Check first (Not checked rows are never ticked for you).' } else { '.' }) } else { 'No row shown is waiting for Recover.' }
        Add-PraGuiLine -Status Info -Text $text
        Show-PraGuiLastLine
    }
    Update-PraGuiTicks -Page $Page -Changed:$Shown
}

function Get-PraGuiSelection {
    <#
        What Preview and Apply run on a page now: the request (action, batch), the GUIDs of a wave, the key that must
        be the same at Apply as at Preview, a label, or the problem when nothing can run.
    #>
    param([Parameter(Mandatory)][ValidateSet('Convert', 'Recover')][string]$Page)
    $g = $script:Gui; $c = $g.Controls
    $selection = [pscustomobject]@{ Request = @{ Action = $Page }; Guids = @(); Key = ''; Label = ''; Problem = ''; Resume = $false }
    if ($g.State.ConfigError) { $selection.Problem = 'Fix the configuration first.'; return $selection }
    if ($Page -eq 'Convert') {
        if (-not $g.State.Snapshot) { $selection.Problem = 'No snapshot: Collect first (on an Exchange server), then copy the database here.'; return $selection }
        $prefix = 'Convert|{0}' -f $g.State.Snapshot.id
        $wave = [bool]$c.ConvertWave.IsChecked
        $all = 'every object of snapshot #{0} ({1})' -f $g.State.Snapshot.id, @($g.State.Objects).Count
    } else {
        if (-not $g.SelectedBatch) { $selection.Problem = 'No Convert batch: nothing to roll back.'; return $selection }
        $selection.Request.Batch = $g.SelectedBatch
        $prefix = 'Recover|{0}' -f $g.SelectedBatch
        $wave = [bool]$c.RecoverWave.IsChecked
        $all = 'the whole batch {0}' -f $g.SelectedBatch
    }
    if ($wave) {
        $guids = [string[]]@(Get-PraGuiTicked $g.Tables[$Page])
        if (-not $guids.Count) { $selection.Problem = 'A wave: tick its objects in the table (or choose {0}).' -f $(if ($Page -eq 'Convert') { 'every object of the snapshot' } else { 'the whole batch' }); return $selection }
        $selection.Guids = $guids
        $selection.Key = Get-PraGuiSelectionKey -Prefix $prefix -Guids $guids
        $selection.Label = 'a wave of {0} object(s)' -f $guids.Count
    } else {
        $selection.Key = $prefix + '|all'
        $selection.Label = $all
    }
    return $selection
}

function Update-PraGuiApplyState {
    <#
        Apply is possible only after a preview of exactly the selection of now (same snapshot or batch, same objects
        ticked); any change asks for a new preview.
    #>
    $g = $script:Gui; $c = $g.Controls
    if (-not $g.State) { return }
    $busy = [bool]$g.Run
    foreach ($page in 'Convert', 'Recover') {
        $pending = $g.Pending[$page]
        $now = Get-PraGuiSelection -Page $page
        if ($pending) {
            $valid = $now.Key -eq $pending.Guard
            if ($valid -and $pending.Resume) { $valid = [bool]@($g.Batches | Where-Object { $_.Id -eq $pending.Batch -and $_.Status -in @('Partial', 'Running') }).Count }
            if (-not $valid) { $g.Pending[$page] = $null; $pending = $null }
        }
        $c[$page + 'Apply'].IsEnabled = [bool]$pending -and -not $busy
        if ($pending) { Set-PraGuiHint -Page $page -Text $pending.Hint -Tone $pending.Tone }
        elseif ($now.Problem) { Set-PraGuiHint -Page $page -Text $now.Problem -Tone 'Caution' }
        else { Set-PraGuiHint -Page $page -Text ('Preview first: {0}. A preview changes nothing.' -f $now.Label) }
    }
}

function Start-PraGuiCheck {
    $scope = $script:Choices.CheckScope[[string]$script:Gui.Controls.CheckScope.SelectedItem]
    Start-PraGuiRun -Request @{ Action = 'Check'; Scope = $(if ($scope) { $scope } else { 'All' }) } -Page 'Check'
}

function Start-PraGuiCollect {
    $g = $script:Gui
    $store = if ($g.State.Config) { $g.State.Config.Store.Path } else { 'the database' }
    if (-not (Show-PraGuiQuestion -Title 'Collect now' -Text ('Read Exchange and write a new snapshot to {0}? Exchange is not changed; the new snapshot becomes the one Check and Convert use.' -f $store))) { return }
    Start-PraGuiRun -Request @{ Action = 'Collect'; Mode = 'Apply' } -Page 'Collect'
}

function Start-PraGuiPreview {
    <# Preview of a page (or of the resume of the unfinished Convert batch): nothing is changed. #>
    param([Parameter(Mandatory)][ValidateSet('Convert', 'Recover')][string]$Page, [switch]$Resume)
    $g = $script:Gui
    if ($Resume) {
        $batch = @($g.Batches | Where-Object { $_.Status -in @('Partial', 'Running') }) | Select-Object -First 1
        if (-not $batch) { throw 'No unfinished Convert batch to resume.' }
        $selection = [pscustomobject]@{ Request = @{ Action = 'Convert'; Batch = $batch.Id }; Guids = @(); Key = 'Convert|resume|' + $batch.Id
            Label = 'the resume of batch ' + $batch.Id; Problem = ''; Resume = $true }
    } else {
        $selection = Get-PraGuiSelection -Page $Page
        if ($selection.Problem) { Set-PraGuiHint -Page $Page -Text $selection.Problem -Tone 'Caution'; return }
    }
    $request = $selection.Request.Clone()
    $request.Mode = 'Preview'
    if ($selection.Guids.Count) { $request.IdentityPath = Write-PraGuiWave -Folder (Get-PraGuiWorkFolder) -Name $Page.ToLowerInvariant() -Rows @($g.Tables[$Page].Select('Selected = true')) }
    $g.Pending[$Page] = $null
    Start-PraGuiRun -Request $request -Page $Page -Selection $selection
}

function Start-PraGuiApply {
    <# Apply of the last preview of the page, after the typed confirmation. #>
    param([Parameter(Mandatory)][ValidateSet('Convert', 'Recover')][string]$Page)
    $g = $script:Gui
    Update-PraGuiApplyState
    $pending = $g.Pending[$Page]
    if (-not $pending) { Set-PraGuiHint -Page $Page -Text 'Preview first: Apply runs exactly what the last preview showed.' -Tone 'Caution'; return }
    $request = $pending.Request.Clone()
    $request.Mode = 'Apply'
    $engine = Get-PraGuiEngine -Action $Page
    $command = Get-PraGuiCommand -Root $g.Root -ConfigPath $g.ConfigPath -Request $request -Engine $(if ($engine.Path) { $engine.Path } else { 'pwsh.exe' })
    $title = if ($Page -eq 'Convert') { 'Convert to Exchange Online' } else { 'Recover: back on-premises' }
    if (-not (Show-PraGuiConfirm -Word $Page.ToUpperInvariant() -Title $title -Text $pending.Confirm -Command $command.Display)) {
        Add-PraGuiLine -Status Info -Text 'Cancelled: nothing was run.'
        Show-PraGuiLastLine
        return
    }
    Start-PraGuiRun -Request $request -Page $Page -Selection $pending.Selection
}
#endregion

#region 4. Runs -------------------------------------------------------------------------------------

function Start-PraGuiRun {
    <#
        Runs one action in a child process, exactly as on the command line, and follows it: the timer of the window
        reads its events (Invoke-PraGuiTick). Nothing waits: the window stays usable while it runs.
    #>
    param([Parameter(Mandatory)][hashtable]$Request, [string]$Page = '', [object]$Selection)
    $g = $script:Gui; $c = $g.Controls
    if ($g.Run) { throw 'An action is already running: wait for its end (or stop it).' }
    $engine = Get-PraGuiEngine -Action $Request.Action
    if (-not $engine.Ok) { throw $engine.Problem }
    $base = Join-Path (Get-PraGuiWorkFolder) ('run-{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'), $Request.Action.ToLowerInvariant())
    $files = @{ Events = "$base.events.jsonl"; Stop = "$base.stop"; Out = "$base.out.txt"; Err = "$base.err.txt" }
    $command = Get-PraGuiCommand -Root $g.Root -ConfigPath $g.ConfigPath -Request $Request -Engine $engine.Path
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $command.FilePath
    $info.Arguments = $command.Arguments
    $info.WorkingDirectory = $g.Root
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    $info.EnvironmentVariables['PRA_EVENT_FILE'] = $files.Events
    $info.EnvironmentVariables['PRA_STOP_FILE'] = $files.Stop
    $info.EnvironmentVariables['PRA_ICONS'] = 'Ascii'
    if ($Request.Action -eq 'Collect') {
        # Windows PowerShell must not load the modules of PowerShell 7 (inherited when the window runs in pwsh).
        $paths = @([Environment]::GetEnvironmentVariable('PSModulePath', 'User'), [Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')) | Where-Object { $_ }
        $info.EnvironmentVariables['PSModulePath'] = ($paths -join ';')
    }
    $process = [Diagnostics.Process]::Start($info)
    # Read in the background (no PowerShell code on other threads): a full pipe would block the run.
    $output = $process.StandardOutput.ReadToEndAsync()
    $errors = $process.StandardError.ReadToEndAsync()
    $g.Run = @{ Request = $Request; Page = $Page; Selection = $Selection; Process = $process; Output = $output; Errors = $errors; Files = $files
        Position = [long]0; Started = (Get-Date); Command = $command; Result = $null; Summary = $null; LogFile = ''
        Counts = @{ Ok = 0; Warn = 0; Fail = 0 }; SubLines = 0; Asked = (New-Object 'Collections.Generic.HashSet[int]'); StopAsked = $false }
    Clear-PraGuiActivity
    $mode = [string]$Request['Mode']
    $c.RunAction.Text = '{0} · {1}' -f $Request.Action, $(if ($Request.Action -eq 'Check') { 'read-only' } elseif ($mode -eq 'Apply') { 'Apply' } else { 'Preview (nothing is changed)' })
    $c.RunMeta.Text = $command.Display
    Add-PraGuiLine -Status Info -Text ('Started: ' + $command.Display)
    Set-PraGuiStatus -Text 'Running' -Tone 'Brand'
    Set-PraGuiBusy -Busy $true
    $g.Timer.Start()
}

function Invoke-PraGuiTick {
    <# The timer of the window: the new events of the run, then its end. Never re-entered (a question is modal). #>
    $g = $script:Gui
    if (-not $g -or -not $g.Run -or $g.InTick) { return }
    $g.InTick = $true
    try {
        $run = $g.Run
        # Exited first: every event written before the end is then in the file.
        $exited = $run.Process.HasExited
        $read = Read-PraGuiEvents -Path $run.Files.Events -Position $run.Position
        $run.Position = $read.Position
        foreach ($record in $read.Events) { Invoke-PraGuiEvent -Record $record }
        if ($read.Events.Count) { Show-PraGuiLastLine }
        if ($exited) { Complete-PraGuiRun; return }
        $elapsed = Format-PraDuration ((Get-Date) - $run.Started).TotalSeconds
        $g.Controls.StatusText.Text = '{0} · {1}' -f $(if ($run.StopAsked) { 'Stopping' } else { 'Running' }), $elapsed
    } finally { $g.InTick = $false }
}

function Invoke-PraGuiEvent {
    <# One event of the run: a line of the activity panel, the step, the counters, a question, the result. #>
    param([Parameter(Mandatory)][object]$Record)
    $g = $script:Gui; $c = $g.Controls; $run = $g.Run
    try {
        switch ([string](Get-PraValue $Record 'kind' '')) {
            'start' {
                $run.LogFile = [string](Get-PraValue $Record 'logFile' '')
                $details = Get-PraValue $Record 'details' $null
                if ($details) { foreach ($property in $details.PSObject.Properties) { Add-PraGuiLine -Status Info -Text ('{0}: {1}' -f $property.Name, $property.Value) } }
            }
            'step' {
                $index = [int](Get-PraValue $Record 'index' 0); $total = [int](Get-PraValue $Record 'total' 0); $title = [string](Get-PraValue $Record 'title' '')
                $c.RunStep.Text = if ($total) { 'Step {0}/{1} · {2}' -f $index, $total, $title } else { $title }
                # [Math]::Min/Max pick the (int,int) overload and truncate when a whole-number literal meets a
                # fraction: 1.0 / 0.0 as explicit doubles keep the fraction.
                if ($total) { $c.RunProgress.Value = [Math]::Max(0.0, [Math]::Min(1.0, ($index - 1) / [double]$total)) }
                # A new step: the object progress of the previous one (if any) no longer applies.
                $c.RunObjects.Text = ''; $c.RunObjects.Visibility = [Windows.Visibility]::Collapsed
                $c.RunObjectsBar.Value = 0; $c.RunObjectsBar.Visibility = [Windows.Visibility]::Collapsed
                Add-PraGuiLine -Status Step -Text $(if ($total) { '{0}/{1}  {2}' -f $index, $total, $title } else { $title })
            }
            'progress' {
                $done = [int](Get-PraValue $Record 'done' 0); $total = [int](Get-PraValue $Record 'total' 0); $phase = [string](Get-PraValue $Record 'phase' '')
                $percent = [int](Get-PraValue $Record 'percent' 0); $eta = [string](Get-PraValue $Record 'eta' '')
                if ($total -gt 0) {
                    $c.RunObjectsBar.Value = [Math]::Max(0.0, [Math]::Min(1.0, $done / [double]$total)); $c.RunObjectsBar.Visibility = [Windows.Visibility]::Visible
                    $c.RunObjects.Text = '{0}: {1} of {2} ({3}%){4}' -f $phase, $done, $total, $percent, $(if ($eta) { ' · ' + $eta } else { '' })
                    $c.RunObjects.Visibility = [Windows.Visibility]::Visible
                }
            }
            'item' {
                $status = [string](Get-PraValue $Record 'status' 'Info')
                if (-not $script:LineStyle.ContainsKey($status)) { $status = 'Info' }
                $identity = [string](Get-PraValue $Record 'identity' '')
                if ($identity) { $c.RunCurrent.Text = 'Object: ' + $identity }
                if ($run.Counts.ContainsKey($status)) {
                    $run.Counts[$status]++
                    $c['Run' + $status].Text = [string]$run.Counts[$status]
                }
                if ($status -eq 'Sub') {
                    $run.SubLines++
                    if ($run.SubLines -gt $script:MaxSubLines) {
                        if ($run.SubLines -eq $script:MaxSubLines + 1) { Add-PraGuiLine -Status Info -Text 'More detail lines in the log of the run.' }
                        return
                    }
                }
                Add-PraGuiLine -Status $status -Text ([string](Get-PraValue $Record 'text' ''))
            }
            'ask' {
                $id = [int](Get-PraValue $Record 'id' 0)
                if (-not $run.Asked.Add($id)) { return }
                $text = [string](Get-PraValue $Record 'text' '')
                $answer = [string](Get-PraValue $Record 'answer' '')
                Add-PraGuiLine -Status Warn -Text ('Question: ' + $text)
                Show-PraGuiLastLine
                Set-PraGuiStatus -Text 'Waiting for you' -Tone 'Caution'
                $yes = $false
                if (-not $run.StopAsked) { $yes = Show-PraGuiQuestion -Title ([string](Get-PraValue $Record 'title' 'PRA Cloud Mailbox')) -Text $text }
                # Only the answer file the run named next to its own event file.
                if ($answer -and $answer.StartsWith($run.Files.Events + '.answer-', [StringComparison]::OrdinalIgnoreCase)) {
                    [IO.File]::WriteAllText($answer, $(if ($yes) { 'yes' } else { 'no' }))
                }
                Add-PraGuiLine -Status Info -Text ('Answer: ' + $(if ($yes) { 'yes' } else { 'no' }))
                Set-PraGuiStatus -Text 'Running' -Tone 'Brand'
            }
            'summary' { $run.Summary = $Record }
            'result' { $run.Result = $Record }
        }
    } catch { Add-PraGuiLine -Status Info -Text ('(event not shown: {0})' -f $_.Exception.Message) }
}

function Complete-PraGuiRun {
    <# The run ended: the result card, then what its page does with it (Complete-PraGuiAction). #>
    $g = $script:Gui; $run = $g.Run
    $g.Timer.Stop()
    $result = $run.Result
    try {
        $process = $run.Process
        [void]$process.WaitForExit(5000)
        $exitCode = -1
        try { $exitCode = $process.ExitCode } catch { $exitCode = -1 }
        $out = ''; $err = ''
        try { if ($run.Output.Wait(3000)) { $out = [string]$run.Output.Result } } catch { $out = '' }
        try { if ($run.Errors.Wait(3000)) { $err = [string]$run.Errors.Result } } catch { $err = '' }
        try { [IO.File]::WriteAllText($run.Files.Out, $out); [IO.File]::WriteAllText($run.Files.Err, $err) } catch { $null = $_ }
        $process.Dispose()
        Show-PraGuiRunResult -Run $run -ExitCode $exitCode -OutputText $out -ErrorText $err
    } finally {
        $g.Run = $null
        Set-PraGuiBusy -Busy $false
    }
    Complete-PraGuiAction -Run $run -Result $result
    Show-PraGuiLastLine
}

function Show-PraGuiRunResult {
    <# The result card and the status of a run that ended, from its events (result, summary) or, without them, its error output. #>
    param([Parameter(Mandatory)][hashtable]$Run, [int]$ExitCode = -1, [AllowEmptyString()][string]$OutputText = '', [AllowEmptyString()][string]$ErrorText = '')
    $c = $script:Gui.Controls
    $result = $Run.Result
    $action = [string]$Run.Request.Action
    $preview = $action -ne 'Check' -and [string]$Run.Request['Mode'] -ne 'Apply'
    if (-not $result) {
        $tail = @(($ErrorText -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 8)
        if (-not $tail.Count) { $tail = @(($OutputText -split "`r?`n") | Where-Object { $_.Trim() } | Select-Object -Last 8) }
        $text = if ($tail.Count) { $tail -join [Environment]::NewLine } else { 'No message: see the log and the transcript in the log folder.' }
        Add-PraGuiLine -Status Fail -Text ('{0} ended without a result (exit code {1}).' -f $action, $ExitCode)
        Set-PraGuiResult -Tone 'Critical' -Title ('{0} did not finish (exit code {1})' -f $action, $ExitCode) -Text $text -Log $Run.LogFile
        Set-PraGuiStatus -Text 'Failed' -Tone 'Critical'
    } else {
        $code = [int](Get-PraValue $result 'exitCode' 1)
        $planned = [int](Get-PraValue $result 'planned' 0)
        $warnings = [int](Get-PraValue $result 'warnings' 0)
        $tone = if ($code -eq 1) { 'Critical' } elseif ($code -eq 2 -or $warnings) { 'Caution' } else { 'Success' }
        # A preview that plans objects reports the objects not ready as errors: it is still a usable preview.
        if ($preview -and $planned -and $code -eq 1) { $tone = 'Caution' }
        $title = if ($Run.Summary) { [string](Get-PraValue $Run.Summary 'title' '') } else { '' }
        if (-not $title) { $title = '{0} finished' -f $action }
        $parts = [Collections.Generic.List[string]]::new()
        $parts.Add(('{0} done · {1} pending · {2} error(s)' -f (Get-PraValue $result 'success' 0), (Get-PraValue $result 'pending' 0), (Get-PraValue $result 'error' 0)))
        if ($planned) { $parts.Add(('{0} planned' -f $planned)) }
        if ($warnings) { $parts.Add(('{0} warning(s)' -f $warnings)) }
        $parts.Add((Format-PraDuration ([double](Get-PraValue $result 'seconds' 0))))
        $issues = @(@(Get-PraValue $result 'issues' @()) | ForEach-Object { $_ } | Where-Object { $_ } | Select-Object -First 3)
        $text = ($parts -join ' · ') + $(if ($issues.Count -and $code -eq 1) { [Environment]::NewLine + ($issues -join [Environment]::NewLine) } else { '' })
        Set-PraGuiResult -Tone $tone -Title $title -Text $text -Batch ([string](Get-PraValue $result 'batchId' '')) -Next ([string[]]@(@(Get-PraValue $result 'nextSteps' @()) | ForEach-Object { $_ } | Where-Object { $_ })) `
            -Report ([string](Get-PraValue $result 'htmlReport' '')) -Log ([string](Get-PraValue $result 'logFile' $Run.LogFile))
        $status = if ($action -eq 'Check') { 'Check done' } elseif ($preview) { 'Preview done' } elseif ($code -eq 1) { 'Failed' } elseif ($code -eq 2) { 'Pending' } else { 'Done' }
        Set-PraGuiStatus -Text $status -Tone $tone
        $c.RunProgress.Value = 1
    }
}

function Get-PraGuiPlanCount {
    <# The objects of a preview, from its CSV report: Planned (Users, Shared) and Blocked (not ready, left as they are). #>
    param([AllowEmptyString()][string]$Path)
    $count = [pscustomobject]@{ Planned = 0; Users = 0; Shared = 0; Blocked = 0 }
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $count }
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter ';' -Encoding UTF8)
    if (-not $rows.Count -or -not $rows[0].PSObject.Properties['FinalStatus']) { return $count }
    foreach ($row in $rows) {
        if ($row.FinalStatus -eq 'Planned') {
            $count.Planned++
            if ($row.Kind -eq 'User') { $count.Users++ } else { $count.Shared++ }
        } elseif ($row.FinalStatus -eq 'Error' -and $row.ObjectGuid) { $count.Blocked++ }
    }
    return $count
}

function Complete-PraGuiAction {
    <#
        What a page does with the end of its run: a preview of Convert or Recover opens Apply on exactly the same
        selection; a Check or a change reads the state again.
    #>
    param([Parameter(Mandatory)][hashtable]$Run, [AllowNull()][object]$Result)
    $g = $script:Gui
    $action = [string]$Run.Request.Action
    $mode = [string]$Run.Request['Mode']
    if ($action -in @('Convert', 'Recover') -and $mode -eq 'Preview') {
        $selection = $Run.Selection
        $count = Get-PraGuiPlanCount -Path $(if ($Result) { [string](Get-PraValue $Result 'csvReport' '') } else { '' })
        if ($Result -and $count.Planned -eq 0) { $count.Planned = [int](Get-PraValue $Result 'planned' 0) }
        $blocked = if ($count.Blocked) { ' · {0} not ready (left as they are)' -f $count.Blocked } else { '' }
        if (-not $Result -or $count.Planned -eq 0) {
            $text = if ($Result) { 'Preview of {0}: nothing to {1}{2}. See the activity.' -f $selection.Label, $action.ToLowerInvariant(), $blocked } else { 'The preview did not finish: see the activity.' }
            Set-PraGuiHint -Page $action -Text $text -Tone 'Caution'
            return
        }
        $split = '{0} user(s), {1} shared' -f $count.Users, $count.Shared
        $confirm = if ($action -eq 'Convert') {
            if ($selection.Resume) { 'Resume of batch {0}: {1} object(s) not finished ({2}) are taken again.{3}' -f $Run.Request.Batch, $count.Planned, $split, $blocked }
            else { '{0} object(s) are converted: {1} user(s) - identity managed in the cloud, Exchange plan, their Teams storage becomes their mailbox - then {2} shared mailbox(es), in waves, with the permissions of the snapshot.{3}' -f $count.Planned, $count.Users, $count.Shared, $blocked }
        } else { '{0} object(s) of batch {1} go back on-premises ({2}): users to Active Directory and their on-premises mailbox, their cloud data under the case hold; shared mailboxes become inactive mailboxes under the hold. Nothing is deleted from a mailbox.{3}' -f $count.Planned, $Run.Request.Batch, $split, $blocked }
        $confirm += ' Every change is journaled; the run can be stopped after the current object and resumed.'
        $guard = if ($selection.Resume) { (Get-PraGuiSelection -Page $action).Key } else { $selection.Key }
        $g.Pending[$action] = @{ Key = $selection.Key; Guard = $guard; Resume = [bool]$selection.Resume; Batch = [string]$Run.Request['Batch']; Request = $Run.Request.Clone(); Selection = $selection
            Planned = $count.Planned; Tone = $(if ($count.Blocked) { 'Caution' } else { 'Success' }); Confirm = $confirm
            Hint = ('Preview of {0}: {1} planned ({2}){3}. "{4}…" runs exactly this.' -f $selection.Label, $count.Planned, $split, $blocked, $action) }
        Update-PraGuiApplyState
        return
    }
    if ($action -eq 'Collect' -and $mode -ne 'Apply') { return }
    $clear = @()
    if ($action -in @('Convert', 'Recover')) {
        $g.Pending[$action] = $null; $clear = @($action)
        $g.Controls[$action + 'All'].IsChecked = $true
    }
    Update-PraGuiState -ClearTicks $clear
}

function Stop-PraGuiRun {
    <# Asks the run to stop: it finishes the object in progress, starts no other one, and ends normally. #>
    $g = $script:Gui; $run = $g.Run
    if (-not $run -or $run.StopAsked) { return }
    [IO.File]::WriteAllText($run.Files.Stop, (Get-Date).ToString('o'))
    $run.StopAsked = $true
    $g.Controls.StopRun.IsEnabled = $false
    Add-PraGuiLine -Status Warn -Text 'Stop asked: the object in progress is finished, no other one is started; the run then ends normally (journal, report). Run the same action again to go on.'
    Show-PraGuiLastLine
    Set-PraGuiStatus -Text 'Stopping' -Tone 'Caution'
}

function Add-PraGuiLine {
    <# One line of the activity panel, with the icon and the colour of the console status. #>
    param([string]$Status = 'Info', [AllowEmptyString()][string]$Text = '')
    $g = $script:Gui
    $style = if ($script:LineStyle.ContainsKey($Status)) { $script:LineStyle[$Status] } else { $script:LineStyle.Info }
    $textTone = switch ($Status) { 'Warn' { 'Caution' } 'Fail' { 'Critical' } { $_ -in @('Sub', 'Skip') } { 'Secondary' } default { 'Primary' } }
    $step = $Status -eq 'Step'
    $g.Lines.Add([pscustomobject]@{ Glyph = [string][char]$style.Glyph; Brush = $g.Brushes[$style.Tone]; Text = $Text; TextBrush = $g.Brushes[$textTone]
            Size = $(if ($step) { 13.5 } else { 12.5 }); Weight = $(if ($step) { [Windows.FontWeights]::SemiBold } else { [Windows.FontWeights]::Normal })
            Margin = [Windows.Thickness]::new($(if ($Status -eq 'Sub') { 18 } else { 0 }), $(if ($step) { 10 } else { 2 }), 0, 2); Time = (Get-Date -Format 'HH:mm:ss') })
    while ($g.Lines.Count -gt 2000) { $g.Lines.RemoveAt(0) }
    $g.Controls.ActivityEmpty.Visibility = [Windows.Visibility]::Collapsed
}

function Show-PraGuiLastLine {
    $g = $script:Gui
    if ($g.Lines.Count) { $g.Controls.ActivityLog.ScrollIntoView($g.Lines[$g.Lines.Count - 1]) }
}

function Clear-PraGuiActivity {
    $g = $script:Gui; $c = $g.Controls
    $g.Lines.Clear()
    $c.ActivityEmpty.Visibility = [Windows.Visibility]::Visible
    foreach ($name in 'RunOk', 'RunWarn', 'RunFail') { $c[$name].Text = '0' }
    $c.RunStep.Text = ''; $c.RunCurrent.Text = ''; $c.RunProgress.Value = 0
    $c.RunObjects.Text = ''; $c.RunObjects.Visibility = [Windows.Visibility]::Collapsed
    $c.RunObjectsBar.Value = 0; $c.RunObjectsBar.Visibility = [Windows.Visibility]::Collapsed
    $c.ResultCard.Visibility = [Windows.Visibility]::Collapsed
}

function Set-PraGuiStatus {
    <# The status pill of the activity panel. #>
    param([Parameter(Mandatory)][string]$Text, [string]$Tone = 'Secondary')
    $c = $script:Gui.Controls
    $pair = switch ($Tone) {
        'Success' { 'PraSuccess', 'PraSuccessBackground' } 'Caution' { 'PraCaution', 'PraCautionBackground' } 'Critical' { 'PraCritical', 'PraCriticalBackground' }
        'Brand' { 'PraBrandText', 'PraAccentSoft' } default { 'TextFillColorSecondaryBrush', 'PraInfoBackground' }
    }
    $c.StatusText.Text = $Text
    $c.StatusText.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[0])
    $c.StatusPill.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[1])
}

function Set-PraGuiResult {
    <# The card at the end of a run: what happened, the batch ID, the next steps, the report and the log. #>
    param([string]$Tone = 'Secondary', [string]$Title = '', [string]$Text = '', [string]$Batch = '', [AllowEmptyCollection()][string[]]$Next = @(), [string]$Report = '', [string]$Log = '')
    $g = $script:Gui; $c = $g.Controls
    $pair = switch ($Tone) {
        'Success' { 'PraSuccessBackground', 'PraSuccessBorder' } 'Caution' { 'PraCautionBackground', 'PraCautionBorder' }
        'Critical' { 'PraCriticalBackground', 'PraCriticalBorder' } default { 'PraInfoBackground', 'PraInfoBorder' }
    }
    $c.ResultCard.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[0])
    $c.ResultCard.SetResourceReference([Windows.Controls.Border]::BorderBrushProperty, $pair[1])
    $c.ResultTitle.Text = $Title
    $c.ResultText.Text = $Text
    $c.BatchText.Text = $Batch
    $c.BatchPanel.Visibility = if ($Batch) { [Windows.Visibility]::Visible } else { [Windows.Visibility]::Collapsed }
    $steps = @($Next | Where-Object { $_ })
    $c.NextStepText.Text = if ($steps.Count) { 'Command line: ' + ($steps -join [Environment]::NewLine) } else { '' }
    $c.NextStepText.Visibility = if ($steps.Count) { [Windows.Visibility]::Visible } else { [Windows.Visibility]::Collapsed }
    $g.LastReport = if ($Report -and (Test-Path -LiteralPath $Report)) { $Report } else { '' }
    $g.LastLog = if ($Log -and (Test-Path -LiteralPath $Log)) { $Log } else { '' }
    $c.OpenReport.IsEnabled = [bool]$g.LastReport
    $c.OpenLog.IsEnabled = [bool]$g.LastLog
    $c.ResultCard.Visibility = [Windows.Visibility]::Visible
}

function Set-PraGuiBusy {
    <# While a run is going: no other run, the stop button (not for Collect, which has no stop). #>
    param([Parameter(Mandatory)][bool]$Busy)
    $g = $script:Gui; $c = $g.Controls
    foreach ($name in $script:RunButtons) { $c[$name].IsEnabled = -not $Busy }
    $stoppable = $Busy -and $g.Run -and $g.Run.Request.Action -ne 'Collect'
    $c.StopRun.Visibility = if ($stoppable) { [Windows.Visibility]::Visible } else { [Windows.Visibility]::Collapsed }
    $c.StopRun.IsEnabled = $true
    Update-PraGuiApplyState
}

function Show-PraGuiConfirm {
    <# The typed confirmation of an Apply: the word (CONVERT, RECOVER) must be typed exactly. $true to go on. #>
    param([Parameter(Mandatory)][string]$Word, [Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Text, [string]$Command = '')
    $g = $script:Gui
    if ($g.Hooks.Confirm) { return [bool](& $g.Hooks.Confirm $Word $Title $Text $Command) }
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PRA Cloud Mailbox" Width="620" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" ShowInTaskbar="False"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" Background="{DynamicResource ApplicationBackgroundBrush}">
  <StackPanel Margin="24,20,24,20">
    <TextBlock x:Name="Heading" FontSize="19" FontWeight="SemiBold" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
    <TextBlock x:Name="Body" FontSize="13.5" TextWrapping="Wrap" Margin="0,10,0,0" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
    <Border Margin="0,14,0,0" Padding="10,8" CornerRadius="6" BorderThickness="1" Background="{DynamicResource PraInfoBackground}" BorderBrush="{DynamicResource PraInfoBorder}">
      <TextBlock x:Name="Command" FontFamily="Cascadia Mono, Consolas" FontSize="12" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
    </Border>
    <TextBlock x:Name="Prompt" FontSize="13" Margin="0,16,0,6" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
    <TextBox x:Name="Word" FontFamily="Cascadia Mono, Consolas" FontSize="14"/>
    <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0">
      <Button x:Name="Ok" IsDefault="True" IsEnabled="False" MinWidth="110" Padding="14,6" Margin="0,0,8,0"/>
      <Button x:Name="Cancel" Content="Cancel" IsCancel="True" MinWidth="96" Padding="14,6"/>
    </StackPanel>
  </StackPanel>
</Window>
'@
    $dialog = [Windows.Markup.XamlReader]::Parse($xaml)
    Set-PraGuiTheme -Window $dialog -Theme $g.Theme
    if ($g.Form.IsVisible) { $dialog.Owner = $g.Form } else { $dialog.WindowStartupLocation = [Windows.WindowStartupLocation]::CenterScreen }
    $dialog.Tag = $Word
    $dialog.FindName('Heading').Text = $Title
    $dialog.FindName('Body').Text = $Text
    $dialog.FindName('Command').Text = $Command
    $dialog.FindName('Prompt').Text = "Type $Word to confirm:"
    $ok = $dialog.FindName('Ok')
    $ok.Content = $Word.Substring(0, 1) + $Word.Substring(1).ToLowerInvariant()
    Set-PraGuiAccentButton -Button $ok -Window $dialog -Theme $g.Theme
    $dialog.FindName('Word').Add_TextChanged({
            param($box)
            $owner = [Windows.Window]::GetWindow($box)
            $owner.FindName('Ok').IsEnabled = [string]$box.Text -ceq [string]$owner.Tag
        })
    $ok.Add_Click({ param($button) [Windows.Window]::GetWindow($button).DialogResult = $true })
    $dialog.Add_ContentRendered({ param($shown) [void]$shown.FindName('Word').Focus() })
    return [bool]$dialog.ShowDialog()
}

function Show-PraGuiQuestion {
    <# A yes/no question (No by default). #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Text)
    $g = $script:Gui
    if ($g.Hooks.Question) { return [bool](& $g.Hooks.Question $Title $Text) }
    $buttons = [Windows.MessageBoxButton]::YesNo; $image = [Windows.MessageBoxImage]::Question; $default = [Windows.MessageBoxResult]::No
    $answer = if ($g.Form.IsVisible) { [Windows.MessageBox]::Show($g.Form, $Text, $Title, $buttons, $image, $default) } else { [Windows.MessageBox]::Show($Text, $Title, $buttons, $image, $default) }
    return $answer -eq [Windows.MessageBoxResult]::Yes
}

function Show-PraGuiNotice {
    <# A message with OK. #>
    param([Parameter(Mandatory)][string]$Text)
    $g = $script:Gui
    if ($g.Hooks.Notice) { [void](& $g.Hooks.Notice $Text); return }
    $buttons = [Windows.MessageBoxButton]::OK; $image = [Windows.MessageBoxImage]::Information
    if ($g.Form.IsVisible) { [void][Windows.MessageBox]::Show($g.Form, $Text, 'PRA Cloud Mailbox', $buttons, $image) } else { [void][Windows.MessageBox]::Show($Text, 'PRA Cloud Mailbox', $buttons, $image) }
}
#endregion

#region 5. Entry ------------------------------------------------------------------------------------

function Show-PraGui {
    <#
    .SYNOPSIS
        Opens the window (Invoke-PraCloudMailbox.ps1 -Gui) and returns when it is closed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ConfigPath, [string]$Version = '')
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
        throw 'The window needs a single-threaded apartment: start PowerShell with -STA (the default of powershell.exe and pwsh.exe).'
    }
    $window = New-PraGuiWindow -Root $Root -ConfigPath $ConfigPath -Version $Version
    # The files of the runs older than 30 days (the waves stay: the logs name them).
    try {
        $limit = (Get-Date).AddDays(-30)
        Get-ChildItem -LiteralPath (Get-PraGuiWorkFolder) -Filter 'run-*' -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $limit } | Remove-Item -Force -ErrorAction SilentlyContinue
    } catch { $null = $_ }
    # Ctrl+C in the console would stop the command that owns the window, and with it every handler of the window.
    $previousCtrlC = $null
    try { if (-not [Console]::IsInputRedirected) { $previousCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } } catch { $previousCtrlC = $null }
    try { [void]$window.Form.ShowDialog() }
    finally {
        if ($null -ne $previousCtrlC) { try { [Console]::TreatControlCAsInput = $previousCtrlC } catch { $null = $_ } }
        if ($script:Gui) {
            foreach ($timer in @($script:Gui.Timer, $script:Gui.FilterTimer)) { if ($timer) { $timer.Stop() } }
        }
        $script:Gui = $null
    }
}
#endregion

Export-ModuleMember -Function Show-PraGui, New-PraGuiWindow, Get-PraGuiEngine, Get-PraGuiCommand, Read-PraGuiEvents, Read-PraGuiLastCheck, Read-PraGuiState,
    Get-PraGuiBatchSummary, Get-PraGuiNextStep, Format-PraGuiTime, ConvertTo-PraGuiCheckTable, ConvertTo-PraGuiConvertTable, ConvertTo-PraGuiRecoverTable,
    Get-PraGuiRowFilter, Get-PraGuiTicked, Write-PraGuiWave, Get-PraGuiSelectionKey, Get-PraGuiPlanCount
