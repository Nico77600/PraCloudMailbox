<#
.SYNOPSIS
    PRA Cloud Mailbox - the guided console (Invoke-PraCloudMailbox.ps1 -Gui).

.DESCRIPTION
    Five pages: Setup (what this computer can do, run settings), Task (what to do, on which objects), Verify (a preview that must run first, then a typed
    confirmation), Run (the steps of the engine, object by object) and Next steps (what to do now, the command line, the batch ID).

    The console never changes anything itself: every run starts Invoke-PraCloudMailbox.ps1 in a child PowerShell 7.5 or later with its own console, exactly as on the
    command line, with the same log, transcript, journal and report. The window follows the run through a file of
    events (PRA_EVENT_FILE, one JSON object per line), asks for a stop with PRA_STOP_FILE (no new object is started, the run ends normally and can be continued) and
    answers the questions of the run (Entra Connect Manual mode) in the answer file of the event. What the operator changes on the Setup page is for that run only: the
    configuration file is never written (a generated copy is passed with -ConfigPath, the original with -BaseConfigPath).

    Runs in PowerShell 7.5 or later with the native Fluent theme of Windows 11 (light or dark). The layout is in PRA.Gui.xaml; the helpers are in PRA.Gui.*.ps1 (data, detection, process, run configuration, views, run).

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
#>
#Requires -Version 7.5
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
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
} ([IO.Path]::GetDirectoryName($PSScriptRoot))
# <<< PRA package integrity bootstrap
$script:Gui = $null
$script:GuiModuleRoot = $PSScriptRoot
. (Join-Path $PSScriptRoot 'PRA.Gui.Data.ps1')
. (Join-Path $PSScriptRoot 'PRA.Gui.Process.ps1')
. (Join-Path $PSScriptRoot 'PRA.Gui.RunConfig.ps1')
. (Join-Path $PSScriptRoot 'PRA.Gui.Detect.ps1')

# The text boxes, choices and tick boxes of the console, by key (the controls are F_<key>, S_<key>_<choice> and C_<key>_<choice>).
$script:GuiTextFields = @('Env', 'Tenant', 'Org', 'AppId', 'Thumb', 'Upn', 'Store', 'SnapId', 'ExpConfig', 'Identity', 'Ou', 'Group', 'Csv', 'Parallel', 'KeepFree', 'WaveSize')
$script:GuiChoices = [ordered]@{
    Scope = @('All', 'UsersOnly', 'SharedOnly'); Snap = @('Latest', 'Specific'); Signin = @('Certificate', 'Interactive')
    Source = @('Single', 'All', 'Ou', 'Group', 'Csv'); Wave = @('Whole', 'Wave'); Expect = @('Provisioned', 'Deprovisioned', 'Retained')
}
# The settings that can differ from the configuration file for one run (what the task asks, Source, Identity, Ou, Group, Csv, Wave, Expect, belongs to the task, not to the settings).
$script:GuiSettingKeys = @('Env', 'Scope', 'Snap', 'SnapId', 'ExpConfig', 'Store', 'Signin', 'Tenant', 'Org', 'AppId', 'Thumb', 'Upn', 'Parallel', 'KeepFree', 'WaveSize')
# The whole-number settings, with their range (WaveSize may also be empty: the same as Parallel).
$script:GuiNumberRange = @{ Parallel = @(1, 1000); KeepFree = @(0, 100000); WaveSize = @(1, 1000) }
$script:GuiSettingLabel = [ordered]@{ Env = 'Environment'; Scope = 'Scope'; Snap = 'Snapshot'; SnapId = 'Snapshot'; ExpConfig = 'Configuration file'; Store = 'Snapshot file'; Signin = 'Sign-in'; Tenant = 'Tenant ID'; Org = 'Organization'; AppId = 'Application ID'; Thumb = 'Certificate'; Upn = 'Administrator'; Parallel = 'Temporary licences'; KeepFree = 'Licences kept free'; WaveSize = 'Wave size' }
$script:GuiActionTitle = @{ Collect = 'Collect a snapshot'; Convert = 'Move mailboxes to Exchange Online'; Recover = 'Roll back a conversion'; Check = 'Check'; Schedule = 'Schedule the daily Collect' }
$script:GuiActionWord = @{ Collect = 'COLLECT'; Convert = 'CONVERT'; Recover = 'RECOVER' }
$script:GuiScopeText = @{ All = 'All'; UsersOnly = 'Users only'; SharedOnly = 'Shared only' }

#region Rows, theme
# The rows shown by the window are ExpandoObject instances: they notify WPF whenever a value is set, so a step, an object or a light changes on screen at once.
function New-PcmGuiRow {
    param([hashtable]$Data)
    $row = [Dynamic.ExpandoObject]::new()
    foreach ($key in $Data.Keys) { $row.$key = $Data[$key] }
    return , $row
}

function Set-PcmGuiRow {
    param($Row, [hashtable]$Data)
    foreach ($key in $Data.Keys) { $Row.$key = $Data[$key] }
}

function Test-PcmGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    if (-not $personalize -or -not $personalize.PSObject.Properties['AppsUseLightTheme']) { return $false }
    return [string]$personalize.AppsUseLightTheme -eq '0'
}

function Initialize-PcmGuiTheme {
    <#
        Loads WPF and applies the native Fluent theme (.NET 9 and later, PowerShell 7.5+), light or dark as Windows (System) or as chosen (Light, Dark). Returns Dark and the application.
    #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $property = [Windows.Application].GetProperty('ThemeMode')
    if (-not $property) { throw 'The console needs PowerShell 7.5 or later (.NET 9, which brings the Fluent theme of Windows). Open it with pwsh.exe 7.5 or later.' }
    $dark = [string]::Equals($Theme, 'Dark', [StringComparison]::OrdinalIgnoreCase)
    if ([string]::Equals($Theme, 'System', [StringComparison]::OrdinalIgnoreCase)) { $dark = Test-PcmGuiDarkMode }
    # Light or Dark, never System: without the setting (Windows Server) WPF would pick dark while the window is light.
    $mode = [Activator]::CreateInstance($property.PropertyType, [object[]]@($(if ($dark) { 'Dark' } else { 'Light' })))
    $property.SetValue($app, $mode, $null)
    return [pscustomobject]@{ Dark = $dark; Application = $app }
}

function Set-PcmGuiTheme {
    <#
        The colours of the console: the brand and the status colours, and the surface, card and text colours that the Fluent theme does not define. Two palettes, light and dark.
    #>
    param([Parameter(Mandatory)]$Window, [Parameter(Mandatory)]$Theme)
    $colors = @{
        PraBrand = @('#B11F4B', '#FD8EA1'); PraBrandText = @('#B11F4B', '#FD8EA1'); PraOnBrand = @('#FFFFFF', '#1B1B1B'); PraAccentSoft = @('#14B11F4B', '#33FD8EA1')
        PraCritical = @('#DC2626', '#F87171'); PraSuccess = @('#16A34A', '#4ADE80'); PraCaution = @('#B45309', '#FBBF24')
        PraInfoBg = @('#F0EEEA', '#2E2E2E'); PraSuccessBg = @('#EAF7EE', '#334ADE80'); PraCautionBg = @('#FFF7E8', '#33FBBF24'); PraCriticalBg = @('#FDECEC', '#33F87171')
        PraTrack = @('#E6E3DE', '#4A4A4A')
        AccentFillColorDefaultBrush = @('#B11F4B', '#FD8EA1'); AccentButtonBackground = @('#B11F4B', '#FD8EA1'); AccentButtonBorderBrush = @('#B11F4B', '#FD8EA1')
        AccentFillColorSecondaryBrush = @('#E6B11F4B', '#E6FD8EA1'); AccentButtonBackgroundPointerOver = @('#E6B11F4B', '#E6FD8EA1')
        AccentFillColorTertiaryBrush = @('#CCB11F4B', '#CCFD8EA1'); AccentButtonBackgroundPressed = @('#CCB11F4B', '#CCFD8EA1')
    }
    $surface = @{
        ApplicationBackgroundBrush = @('#F7F4EF', '#202020'); CardBackgroundFillColorDefaultBrush = @('#FFFFFF', '#2B2B2B'); CardStrokeColorDefaultBrush = @('#DEDEDE', '#3D3D3D')
        ControlFillColorDefaultBrush = @('#FFFFFF', '#2D2D2D'); ControlFillColorSecondaryBrush = @('#F5F5F5', '#363636'); DividerStrokeColorDefaultBrush = @('#DEDEDE', '#3D3D3D')
        LayerFillColorDefaultBrush = @('#FCFBF8', '#262626'); TextFillColorPrimaryBrush = @('#242424', '#FFFFFF'); TextFillColorSecondaryBrush = @('#5C5C5C', '#C5C5C5')
        TextFillColorTertiaryBrush = @('#8A8A8A', '#9A9A9A'); AccentButtonForeground = @('#FFFFFF', '#1B1B1B')
    }
    $index = [int][bool]$Theme.Dark
    foreach ($key in $colors.Keys) {
        $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($colors[$key][$index]))
        $brush.Freeze(); $Window.Resources[$key] = $brush
    }
    foreach ($key in $surface.Keys) {
        if ($null -ne $Window.TryFindResource($key)) { continue }
        $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($surface[$key][$index]))
        $brush.Freeze(); $Window.Resources[$key] = $brush
    }
}

function Get-PcmGuiBrush {
    <# The brush of a resource key in the window ($null before the window exists). #>
    param([string]$Key)
    if ($script:Gui -and $script:Gui.Form) { return $script:Gui.Form.TryFindResource($Key) }
    return $null
}

function Set-PcmGuiPill {
    <# Colours a pill (border and text) for a level: Ok, Warn, Fail, Brand or neutral. #>
    param($Border, $Text, [string]$Level)
    $pair = switch ($Level) {
        'Ok' { 'PraSuccessBg', 'PraSuccess' } 'Warn' { 'PraCautionBg', 'PraCaution' } 'Fail' { 'PraCriticalBg', 'PraCritical' }
        'Brand' { 'PraAccentSoft', 'PraBrandText' } default { 'PraInfoBg', 'TextFillColorSecondaryBrush' }
    }
    $Border.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[0])
    $Text.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[1])
}

function Set-PcmGuiLevelColor {
    <# The colour of a text for a level (Ok, Warn, Fail or neutral). #>
    param($Text, [string]$Level)
    $key = switch ($Level) { 'Ok' { 'PraSuccess' } 'Warn' { 'PraCaution' } 'Fail' { 'PraCritical' } default { 'TextFillColorSecondaryBrush' } }
    $Text.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $key)
}

function Get-PcmGuiVisibility {
    param([bool]$Visible)
    if ($Visible) { return 'Visible' }
    return 'Collapsed'
}
#endregion

#region Run settings: fields, defaults, overrides
function Get-PcmGuiField {
    param([string]$Key)
    $c = $script:Gui.Controls
    if ($script:GuiChoices.Contains($Key)) {
        foreach ($value in $script:GuiChoices[$Key]) {
            $radio = $c[('S_{0}_{1}' -f $Key, $value)]
            if ($radio -and $radio.IsChecked) { return [string]$value }
        }
        return [string]$script:GuiChoices[$Key][0]
    }
    if ($Key -eq 'Once') { return [bool]$c.F_Once.IsChecked }
    return ([string]$c[('F_' + $Key)].Text).Trim()
}

function Set-PcmGuiField {
    param([string]$Key, $Value)
    $g = $script:Gui; $c = $g.Controls
    $g.Updating++
    try {
        if ($script:GuiChoices.Contains($Key)) {
            foreach ($choice in $script:GuiChoices[$Key]) {
                $radio = $c[('S_{0}_{1}' -f $Key, $choice)]
                if ($radio) { $radio.IsChecked = (Test-PcmGuiOrdinal $choice $Value) }
                $mirror = $c[('C_{0}_{1}' -f $Key, $choice)]
                if ($mirror) { $mirror.IsChecked = (Test-PcmGuiOrdinal $choice $Value) }
            }
        }
        elseif ($Key -eq 'Once') { $c.F_Once.IsChecked = [bool]$Value }
        else { $c[('F_' + $Key)].Text = [string]$Value }
    } finally { $g.Updating-- }
}

function Set-PcmGuiDefault {
    <# The default of every run setting is the value of the configuration file: what the file says is what the run does unless the operator changes it. #>
    $g = $script:Gui
    $config = $g.State.Config
    if (-not $config) { $g.Defaults = @{ Env = ''; Scope = 'All'; Snap = 'Latest'; SnapId = ''; ExpConfig = ''; Store = ''; Signin = 'Certificate'; Tenant = ''; Org = ''; AppId = ''; Thumb = ''; Upn = ''; Parallel = '100'; KeepFree = '0'; WaveSize = '' }; return }
    $shared = $config.Licensing.Shared
    $keepFree = [string](Get-PcmGuiValue $shared 'KeepFree' 0)
    $waveSize = [string](Get-PcmGuiValue (Get-PcmGuiValue $config 'Recover' @{}) 'WaveSize' 0)
    $g.Defaults = @{
        Env = [string]$config.Environment; Scope = 'All'; Snap = 'Latest'; SnapId = ''; ExpConfig = ''; Store = [string]$config.Store.Path
        Signin = [string]$config.Cloud.SignIn; Tenant = [string]$config.Cloud.TenantId; Org = [string]$config.Cloud.Organization
        AppId = [string]$config.Cloud.AppId; Thumb = [string]$config.Cloud.CertificateThumbprint; Upn = [string]$config.Cloud.UserPrincipalName
        Parallel = [string]$shared.Parallel; KeepFree = $keepFree; WaveSize = $(if ($waveSize -eq '0') { '' } else { $waveSize })
    }
}

function Reset-PcmGuiField {
    param([string]$Key)
    Set-PcmGuiField $Key $script:Gui.Defaults[$Key]
    if ($Key -eq 'Snap') { Set-PcmGuiField 'SnapId' '' }
}

function Test-PcmGuiSettingUsed {
    <# Is a run setting meaningful here? The certificate fields only with the certificate sign-in, the administrator only with the interactive sign-in, the snapshot number only when a snapshot is chosen, the wave size only for Recover. #>
    param([string]$Key)
    $g = $script:Gui
    $signin = Get-PcmGuiField 'Signin'
    switch ($Key) {
        { $_ -in @('AppId', 'Thumb') } { return (Test-PcmGuiOrdinal $signin 'Certificate') }
        'Upn' { return (Test-PcmGuiOrdinal $signin 'Interactive') }
        'SnapId' { return (Test-PcmGuiOrdinal (Get-PcmGuiField 'Snap') 'Specific') }
        'WaveSize' { return $g.Action -eq 'Recover' }
        default { return $true }
    }
}

function ConvertTo-PcmGuiWholeNumber {
    <# A whole number typed with spaces around it (' 40 ') as a number; $null for anything else (empty, a decimal, '4 0', a sign). #>
    param([AllowNull()][AllowEmptyString()][string]$Text)
    $value = ([string]$Text).Trim()
    $number = 0
    if ($value -match '^\d{1,9}$' -and [int]::TryParse($value, [ref]$number)) { return $number }
    return $null
}

function Test-PcmGuiModified {
    param([string]$Key)
    $g = $script:Gui
    if (-not $g.Defaults) { return $false }
    if ($Key -eq 'SnapId') { return (Test-PcmGuiSettingUsed 'SnapId') -and -not (Test-PcmGuiOrdinal (Get-PcmGuiField 'SnapId') $g.Defaults.SnapId) }
    if ($Key -eq 'Store') { return -not (Test-PcmGuiOrdinal (Get-PcmGuiField 'Store') $g.Defaults.Store) -and [bool](Get-PcmGuiField 'Store') }
    if ($script:GuiNumberRange.ContainsKey($Key)) {
        $now = ConvertTo-PcmGuiWholeNumber (Get-PcmGuiField $Key); $was = ConvertTo-PcmGuiWholeNumber $g.Defaults[$Key]
        if ($null -ne $now -and $null -ne $was) { return $now -ne $was }
        return -not (Test-PcmGuiOrdinal (Get-PcmGuiField $Key) $g.Defaults[$Key])
    }
    return -not (Test-PcmGuiOrdinal (Get-PcmGuiField $Key) $g.Defaults[$Key])
}

function Get-PcmGuiChangedSetting {
    <# The labels of the run settings that differ from the configuration file and matter here (one per setting). #>
    $names = [Collections.Generic.List[string]]::new()
    foreach ($key in $script:GuiSettingKeys) {
        if ($key -eq 'SnapId') { continue }
        if (-not (Test-PcmGuiSettingUsed $key) -and $key -ne 'Env' -and $key -ne 'Scope' -and $key -ne 'ExpConfig') { continue }
        if (Test-PcmGuiModified $key) { $names.Add([string]$script:GuiSettingLabel[$key]) }
    }
    if ((Test-PcmGuiModified 'SnapId') -and -not $names.Contains('Snapshot')) { $names.Add('Snapshot') }
    return $names.ToArray()
}

function Get-PcmGuiOverride {
    <# The configuration settings the operator changed for this run, nested like the configuration (empty when nothing differs from the file). #>
    $g = $script:Gui
    $override = @{}; $cloud = @{}
    if (Test-PcmGuiModified 'Env') { $override.Environment = Get-PcmGuiField 'Env' }
    if (Test-PcmGuiModified 'Store') { $override.Store = @{ Path = (Get-PcmGuiField 'Store') } }
    if (Test-PcmGuiSettingUsed 'Signin') {
        foreach ($pair in @(@('Signin', 'SignIn'), @('Tenant', 'TenantId'), @('Org', 'Organization'))) {
            if (Test-PcmGuiModified $pair[0]) { $cloud[$pair[1]] = Get-PcmGuiField $pair[0] }
        }
        foreach ($pair in @(@('AppId', 'AppId'), @('Thumb', 'CertificateThumbprint'), @('Upn', 'UserPrincipalName'))) {
            if ((Test-PcmGuiSettingUsed $pair[0]) -and (Test-PcmGuiModified $pair[0])) { $cloud[$pair[1]] = Get-PcmGuiField $pair[0] }
        }
    }
    if ($cloud.Count) { $override.Cloud = $cloud }
    $shared = @{}
    foreach ($key in @('Parallel', 'KeepFree')) { if ((Test-PcmGuiSettingUsed $key) -and (Test-PcmGuiModified $key)) { $shared[$key] = [int](ConvertTo-PcmGuiWholeNumber (Get-PcmGuiField $key)) } }
    if ($shared.Count) { $override.Licensing = @{ Shared = $shared } }
    if ((Test-PcmGuiSettingUsed 'WaveSize') -and (Test-PcmGuiModified 'WaveSize')) { $override.Recover = @{ WaveSize = [int](ConvertTo-PcmGuiWholeNumber (Get-PcmGuiField 'WaveSize')) } }
    return $override
}

function Get-PcmGuiSettingsProblem {
    <# Cheap checks of what the operator typed; the engine validates the whole configuration again. #>
    $g = $script:Gui
    if (-not $g.State.Config) { return '' }
    if ((Get-PcmGuiField 'Env') -notmatch '^[A-Za-z0-9_.-]{1,64}$') { return 'Environment: 1 to 64 letters, digits, dot, dash or underscore.' }
    if ((Test-PcmGuiSettingUsed 'SnapId')) {
        $id = [long]0
        if (-not [long]::TryParse((Get-PcmGuiField 'SnapId'), [ref]$id) -or $id -lt 1) { return 'Snapshot: type the number of the snapshot (1 or more).' }
    }
    $tenant = Get-PcmGuiField 'Tenant'
    $guid = [guid]::Empty
    if ((Test-PcmGuiSettingUsed 'Tenant') -and $tenant -and -not [guid]::TryParse($tenant, [ref]$guid)) { return 'Tenant ID: a GUID (00000000-0000-0000-0000-000000000000).' }
    foreach ($key in @('Parallel', 'KeepFree', 'WaveSize')) {
        if (-not (Test-PcmGuiSettingUsed $key)) { continue }
        $text = Get-PcmGuiField $key
        if ($key -eq 'WaveSize' -and -not $text) { continue }
        $range = $script:GuiNumberRange[$key]; $number = ConvertTo-PcmGuiWholeNumber $text
        if ($null -eq $number -or $number -lt $range[0] -or $number -gt $range[1]) { return ('{0}: a whole number from {1} to {2}.' -f $script:GuiSettingLabel[$key], $range[0], $range[1]) }
    }
    $organization = Get-PcmGuiField 'Org'
    if ((Test-PcmGuiSettingUsed 'Org') -and $organization -and $organization -notmatch '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') { return 'Organization: tenant.onmicrosoft.com.' }
    return ''
}

function Get-PcmGuiCloudProblem {
    <# What a cloud action still needs from the settings (empty when nothing). #>
    if ((Test-PcmGuiOrdinal (Get-PcmGuiField 'Signin') 'Certificate')) {
        if (-not (Get-PcmGuiField 'AppId') -or -not (Get-PcmGuiField 'Thumb')) { return 'The certificate sign-in needs the application ID and the certificate thumbprint (Setup, Change).' }
        if (-not (Get-PcmGuiField 'Tenant') -or -not (Get-PcmGuiField 'Org')) { return 'The certificate sign-in needs the Tenant ID and the Organization (Setup, Change).' }
    } elseif (-not (Get-PcmGuiField 'Tenant')) { return 'The interactive sign-in needs the Tenant ID (Setup, Change).' }
    return ''
}
#endregion

#region Detection: lights, Microsoft 365, role
function Invoke-PcmGuiDetection {
    <# Runs the detection (the real one by default) with the settings of the window, and shows the lights. -Refresh looks at the installed modules again. #>
    param([switch]$Refresh)
    $g = $script:Gui
    $config = $g.State.Config
    $request = @{ SignIn = (Get-PcmGuiField 'Signin'); Thumbprint = (Get-PcmGuiField 'Thumb'); AppId = (Get-PcmGuiField 'AppId'); Tenant = (Get-PcmGuiField 'Tenant') }
    if ($config) {
        $request.MinExoVersion = [string]$config.Exo.MinModuleVersion
        $request.ExchangeServer = [string]$config.Exchange.Server
        $request.Snapshot = Get-PcmGuiSnapshotInfo -State $g.State
        $request.Journal = [pscustomobject]@{ Path = [string]$config.Store.JournalPath; Exists = [bool]$g.State.JournalExists; Batches = @($g.State.Batches).Count; Error = [string]$g.State.JournalError }
    }
    if ($Refresh) { $request.Refresh = $true; Stop-PcmGuiProbe }
    $g.Detect = & $g.Probe $request
    $g.Controls.RecheckNote.Text = 'Checked at ' + [datetime]::Now.ToString('HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    Add-PcmGuiProbeLight
    $g.LightRows.Clear()
    foreach ($light in @($g.Detect.Lights)) { $g.LightRows.Add((New-PcmGuiRow @{ Name = [string]$light.Name; Status = [string]$light.Status; Detail = [string]$light.Detail })) }
}

function Get-PcmGuiNetworkState {
    <# The state of the Microsoft 365 test: '' (not tested), Checking, Ok, Partial or Fail. #>
    $g = $script:Gui
    if ($g.NetJob) { return 'Checking' }
    if ($g.NetResult) { return [string]$g.NetResult.State }
    return ''
}

function Get-PcmGuiExchangeState {
    <# The state of the Exchange test: '' (not tested), Checking, Ok or Fail. #>
    $g = $script:Gui
    if ($g.ExchJob) { return 'Checking' }
    if ($g.ExchResult) { return $(if ($g.ExchResult.Ok) { 'Ok' } else { 'Fail' }) }
    return ''
}

function Test-PcmGuiNetworkDown {
    return (Test-PcmGuiOrdinalIn (Get-PcmGuiNetworkState) @('Fail', 'Partial'))
}

function Stop-PcmGuiProbe {
    $g = $script:Gui
    $g.NetJob = $null; $g.NetResult = $null; $g.ExchJob = $null; $g.ExchResult = $null
}

function Start-PcmGuiNetworkCheck {
    <# Sends the requests and returns at once; Complete-PcmGuiNetworkCheck (window timer) reads the answers. #>
    $g = $script:Gui
    $g.NetResult = $null
    $targets = @(foreach ($name in $script:GuiM365Hosts) {
            $task = $null; $failure = ''
            try { $task = & $g.NetProbe $name $g.NetSeconds } catch { $failure = $_.Exception.Message }
            [pscustomobject]@{ Host = $name; Task = $task; Failure = $failure }
        })
    $g.NetJob = @{ Targets = $targets; Clock = [Diagnostics.Stopwatch]::StartNew() }
    $g.Timer.Start()
}

function Start-PcmGuiExchangeCheck {
    <# Opens (and closes) a remote session to Exchange.Server in the background and returns at once; Complete-PcmGuiExchangeCheck (window timer) reads the answer. #>
    $g = $script:Gui
    $g.ExchResult = $null
    $server = [string]$g.Detect.ExchangeServer
    $task = $null; $failure = ''
    try { $task = & $g.ExchProbe $server $g.ExchSeconds } catch { $failure = $_.Exception.Message }
    $g.ExchJob = @{ Target = [pscustomobject]@{ Host = $server; Task = $task; Failure = $failure }; Clock = [Diagnostics.Stopwatch]::StartNew() }
    $g.Timer.Start()
}

function Complete-PcmGuiNetworkCheck {
    <# Called by the window timer: does nothing until both answers are in or the time is up. #>
    $g = $script:Gui; $job = $g.NetJob
    if (-not $job) { return }
    $expired = $job.Clock.Elapsed.TotalSeconds -ge $g.NetSeconds
    $waiting = @($job.Targets | Where-Object { $_.Task -and -not $_.Task.IsCompleted })
    if ($waiting.Count -and -not $expired) { return }
    $answers = @(ConvertTo-PcmGuiNetworkAnswer -Targets $job.Targets -Seconds $g.NetSeconds)
    $g.NetJob = $null
    $g.NetResult = [pscustomobject]@{ State = (Get-PcmGuiNetworkVerdict $answers); Answers = $answers }
    Sync-PcmGuiTimer
    $text = @($answers | ForEach-Object { if ($_.Ok) { '{0}: reachable' -f $_.Host } else { '{0}: not reachable ({1})' -f $_.Host, $_.Why } }) -join '; '
    Add-PcmGuiLine ('Microsoft 365 check from this computer. ' + $text) $(if ($g.NetResult.State -eq 'Ok') { 'Info' } else { 'Warn' })
    if (-not $g.Run) { Reset-PcmGuiPreview }
    Invoke-PcmGuiDetection
    Update-PcmGuiView
}

function Complete-PcmGuiExchangeCheck {
    <# Called by the window timer: does nothing until the answer is in or the time is up. #>
    $g = $script:Gui; $job = $g.ExchJob
    if (-not $job) { return }
    $expired = $job.Clock.Elapsed.TotalSeconds -ge $g.ExchSeconds
    if ($job.Target.Task -and -not $job.Target.Task.IsCompleted -and -not $expired) { return }
    $answer = ConvertTo-PcmGuiExchangeAnswer -Target $job.Target -Seconds $g.ExchSeconds
    $g.ExchJob = $null
    $g.ExchResult = $answer
    Sync-PcmGuiTimer
    Add-PcmGuiLine ('Exchange check from this computer. ' + $(if ($answer.Ok) { '{0}: a remote session opens.' -f $answer.Host } else { '{0}: no remote session ({1}).' -f $answer.Host, $answer.Why })) $(if ($answer.Ok) { 'Info' } else { 'Warn' })
    if (-not $g.Run) { Reset-PcmGuiPreview }
    Invoke-PcmGuiDetection
    Update-PcmGuiView
}

function Add-PcmGuiProbeLight {
    <# Adds the Microsoft 365 light and the Exchange light, starts their tests, and settles the role. Each light is absent where its test cannot apply (a test passes its own probes). #>
    $g = $script:Gui; $d = $g.Detect
    if ($d.CanCloudBase -and $g.NetProbe -and -not $g.NetResult -and -not $g.NetJob) { Start-PcmGuiNetworkCheck }
    if ($d.CanCollect -and $g.ExchProbe -and -not $g.ExchResult -and -not $g.ExchJob) { Start-PcmGuiExchangeCheck }
    $state = if ($d.CanCloudBase -and $g.NetProbe) { Get-PcmGuiNetworkState } else { '' }
    $exchangeState = if ($d.CanCollect -and $g.ExchProbe) { Get-PcmGuiExchangeState } else { '' }
    $items = [Collections.Generic.List[object]]::new()
    foreach ($item in @($d.Lights)) { $items.Add($item) }
    if ($g.NetProbe) {
        $answers = @(if ($g.NetResult) { $g.NetResult.Answers })
        if ($d.CanCloudBase) { $shown = Get-PcmGuiNetworkLight -State $state -Answers $answers; $status = $shown.Status; $detail = $shown.Detail }
        else { $status = 'NA'; $detail = 'Not tested: the lights above must be green first.' }
        $items.Insert([Math]::Min(2, $items.Count), [ordered]@{ Name = $script:GuiNetLight; Status = $status; Detail = $detail })
    }
    if ($g.ExchProbe) {
        $shown = Get-PcmGuiExchangeLight -State $exchangeState -Answer $g.ExchResult -Server ([string]$d.ExchangeServer)
        $at = $items.Count
        for ($i = 0; $i -lt $items.Count; $i++) { if (Test-PcmGuiOrdinal $items[$i].Name 'Daily Collect') { $at = $i; break } }
        $items.Insert($at, [ordered]@{ Name = $script:GuiExchangeLight; Status = $shown.Status; Detail = $shown.Detail })
    }
    Add-Member -InputObject $d -NotePropertyName Lights -NotePropertyValue $items.ToArray() -Force
    Add-Member -InputObject $d -NotePropertyName Net -NotePropertyValue $state -Force
    Add-Member -InputObject $d -NotePropertyName Role -NotePropertyValue (Get-PcmGuiRole -Detect $d -Net $(if ($g.NetProbe) { $state } else { 'Ok' }) -Exchange $exchangeState) -Force
}

function Test-PcmGuiCanCollect {
    <# Can this computer run Collect? Not known to be unable: PowerShell 7.5, Exchange.Server set, and Exchange not known to be unreachable. #>
    $d = $script:Gui.Detect
    if (-not $d -or -not $d.CanCollect) { return $false }
    return (Get-PcmGuiExchangeState) -ne 'Fail'
}

function Test-PcmGuiCanCloud {
    <# Can this computer run the cloud actions? Not known to be unable: the lights are green and Microsoft 365 is not known to be unreachable. #>
    $d = $script:Gui.Detect
    if (-not $d -or -not $d.CanCloudBase) { return $false }
    return -not (Test-PcmGuiNetworkDown)
}
#endregion

#region Task: what the operator asked
function Get-PcmGuiNeedsCloud {
    $g = $script:Gui
    return $g.Action -in @('Check', 'Convert', 'Recover')
}

function Get-PcmGuiCollectProblem {
    <# Why this computer cannot run Collect (empty when it can). #>
    $d = $script:Gui.Detect
    if (-not $d) { return '' }
    if (-not $d.CanCloudEngine) { return 'This computer cannot run Collect: it needs PowerShell 7.5 or later (see Setup).' }
    if (-not $d.ExchangeServer) { return 'This computer cannot run Collect: Exchange.Server is not set in the configuration file. Collect reads Exchange through a remote session to that server.' }
    if ((Get-PcmGuiExchangeState) -eq 'Fail') {
        $light = @($d.Lights | Where-Object { Test-PcmGuiOrdinal $_.Name $script:GuiExchangeLight })
        return $(if ($light.Count) { 'This computer cannot run Collect: ' + [string]$light[0].Detail } else { 'This computer cannot run Collect: Exchange cannot be reached from here.' })
    }
    return ''
}

function Get-PcmGuiSnapshotBlocker {
    <# The snapshot file was changed after its last Collect: nothing runs from it until it is restored (empty when it is intact or cannot be checked). #>
    $info = Get-PcmGuiSnapshotInfo -State $script:Gui.State
    if (-not (Test-PcmGuiOrdinal $info.Integrity 'Changed')) { return '' }
    return ('The snapshot file was changed after its last Collect, so nothing runs from it. ' + [string]$info.IntegrityDetail)
}

function Get-PcmGuiBlocker {
    <# Why this computer cannot do the chosen task yet (empty when it can). #>
    $g = $script:Gui
    if (-not $g.Action -or -not $g.Detect) { return '' }
    if ($g.Action -eq 'Collect') {
        if (-not (Test-PcmGuiCanCollect)) { return Get-PcmGuiCollectProblem }
        return ''
    }
    if (-not (Test-PcmGuiCanCloud)) {
        $reason = ''
        if (Test-PcmGuiNetworkDown) { $net = @($g.Detect.Lights | Where-Object { Test-PcmGuiOrdinal $_.Name $script:GuiNetLight }); $reason = if ($net.Count) { [string]$net[0].Detail } else { 'Microsoft 365 is not reachable from this computer.' } }
        else { foreach ($light in $g.Detect.Lights) { if ((Test-PcmGuiOrdinal $light.Status 'Fail') -and -not (Test-PcmGuiOrdinal $light.Name 'Snapshot')) { $reason = [string]$light.Detail; break } } }
        if (-not $reason) { $reason = 'a prerequisite is missing.' }
        return ('This computer cannot run the cloud actions yet: {0} Fix it on the Setup page, or run this on the cloud admin server.' -f $reason)
    }
    $cloud = Get-PcmGuiCloudProblem
    if ($cloud) { return $cloud }
    if (-not $g.State.Snapshot) { return 'There is no complete snapshot to work from. Copy the database written by Collect to this computer (Setup shows where the tool looks for it), then check again.' }
    return Get-PcmGuiSnapshotBlocker
}

function Get-PcmGuiSourceKind {
    <# The way the objects are picked for the current task: Single, All, Ou, Group or Csv; '' when the task picks no objects (Collect, a batch). #>
    $g = $script:Gui
    switch ($g.Action) {
        'Convert' { if ($g.Mode -eq 'New') { return Get-PcmGuiField 'Source' } return '' }
        'Check' { if ($g.Mode -eq 'New') { return Get-PcmGuiField 'Source' } return '' }
        default { return '' }
    }
}

function Get-PcmGuiBatchKind {
    <# The kind of the batch list of the current task, '' when the task has none. #>
    $g = $script:Gui
    switch ($g.Action) {
        'Convert' { if ($g.Mode -eq 'Continue') { return 'ContinueConvert' } }
        'Recover' { if ($g.Mode -eq 'Continue') { return 'ContinueRollback' } else { return 'Rollback' } }
        'Check' { if ($g.Mode -eq 'Continue') { return 'Follow' } }
    }
    return ''
}

function Get-PcmGuiTicked {
    <# The rows ticked in the grid of the wave. #>
    $g = $script:Gui
    return @($g.WaveRows | Where-Object { $_.Checked })
}

function Get-PcmGuiSourceProblem {
    <# What is missing to pick the objects (empty when nothing). Looks in the snapshot, never in the cloud. #>
    $g = $script:Gui
    switch (Get-PcmGuiSourceKind) {
        'Single' {
            $identity = Get-PcmGuiField 'Identity'
            if (-not $identity) { return 'Type the e-mail address, the sign-in name or the account name of the mailbox.' }
            if ($g.State.Snapshot -and -not @(Find-PcmGuiMailbox -State $g.State -Text $identity).Count) { return ("'{0}' is not in snapshot #{1}. Check the spelling." -f $identity, $g.State.Snapshot.id) }
        }
        'Ou' {
            $ou = Get-PcmGuiField 'Ou'
            if (-not $ou) { return 'Pick an OU, or type its name.' }
            if ($g.State.Snapshot -and -not @(Get-PcmGuiOuMatch -State $g.State -Text $ou).Count) { return ("No mailbox of snapshot #{0} is in '{1}'." -f $g.State.Snapshot.id, $ou) }
        }
        'Group' { if (-not (Get-PcmGuiField 'Group')) { return 'Type the name of the group.' } }
        'Csv' {
            $csv = Get-PcmGuiField 'Csv'
            if (-not $csv) { return 'Choose a CSV file.' }
            if (-not (Test-Path -LiteralPath $csv -PathType Leaf)) { return ("The file '{0}' does not exist." -f $csv) }
        }
    }
    return ''
}

function Test-PcmGuiTaskReady {
    $g = $script:Gui
    if (-not $g.Action) { return $false }
    if ($g.State.ConfigError -or (Get-PcmGuiSettingsProblem) -or (Get-PcmGuiBlocker)) { return $false }
    if ($g.Action -eq 'Collect') { return $true }
    $kind = Get-PcmGuiBatchKind
    if ($kind) {
        if (-not $g.SelectedBatch) { return $false }
        if ($g.Action -eq 'Recover' -and $g.Mode -eq 'New' -and (Get-PcmGuiField 'Wave') -eq 'Wave') { return @(Get-PcmGuiTicked).Count -gt 0 }
        return $true
    }
    return -not (Get-PcmGuiSourceProblem)
}

function Get-PcmGuiSnapshotNumber {
    <# The snapshot given to -Snapshot: the one typed on the Setup page, 0 for the latest complete one. #>
    if (-not (Test-PcmGuiOrdinal (Get-PcmGuiField 'Snap') 'Specific')) { return [long]0 }
    $id = [long]0
    if ([long]::TryParse((Get-PcmGuiField 'SnapId'), [ref]$id)) { return $id }
    return [long]0
}

function Get-PcmGuiWindowRequest {
    <# The request of the current task for the engine. Mode Preview or Apply; a wave is written to its CSV file (the same file for the preview and for the run that follows it). #>
    param([string]$Mode = 'Preview')
    $g = $script:Gui
    if (-not $g.Action) { throw 'Choose a task first.' }
    $action = $g.Action
    $values = @{ Action = $action; Mode = $(if ($action -eq 'Check') { 'Preview' } else { $Mode }); Scope = (Get-PcmGuiField 'Scope'); Snapshot = (Get-PcmGuiSnapshotNumber) }
    if ($action -eq 'Collect') { $values.Snapshot = [long]0 }
    switch (Get-PcmGuiSourceKind) {
        'Single' { $values.Identity = Get-PcmGuiField 'Identity' }
        'Ou' { $values.SearchBase = Get-PcmGuiField 'Ou' }
        'Group' { $values.Group = Get-PcmGuiField 'Group' }
        'Csv' { $values.CsvPath = Get-PcmGuiField 'Csv' }
    }
    if (Get-PcmGuiBatchKind) {
        if (-not $g.SelectedBatch) { throw 'Pick the batch first.' }
        $values.Batch = [string]$g.SelectedBatch.BatchId
        if ($action -eq 'Recover' -and $g.Mode -eq 'New' -and (Get-PcmGuiField 'Wave') -eq 'Wave') {
            $ticked = @(Get-PcmGuiTicked)
            if (-not $ticked.Count) { throw 'Tick the objects of the wave.' }
            if (-not $g.WaveFile -or -not $g.WaveKey -or $g.WaveKey -ne (Get-PcmGuiSelectionKey -Prefix $values.Batch -Guids @($ticked | ForEach-Object { $_.Guid }))) {
                $folder = Join-Path (Join-Path $g.Root 'logs') 'gui'
                Initialize-PcmGuiFolder (Join-Path $g.Root 'logs'); Initialize-PcmGuiFolder $folder
                $g.WaveFile = Write-PcmGuiWave -Folder $folder -Name $values.Batch -Rows $ticked
                $g.WaveKey = Get-PcmGuiSelectionKey -Prefix $values.Batch -Guids @($ticked | ForEach-Object { $_.Guid })
            }
            $values.CsvPath = $g.WaveFile
        }
        elseif ($action -eq 'Recover' -and @($g.SelectedBatch.Wave).Count) {
            $remaining = @($g.SelectedBatch.Wave)
            $selectionKey = Get-PcmGuiSelectionKey -Prefix ('continue-' + $values.Batch) -Guids @($remaining | ForEach-Object { $_.Guid })
            if (-not $g.WaveFile -or -not $g.WaveKey -or $g.WaveKey -ne $selectionKey) {
                $folder = Join-Path (Join-Path $g.Root 'logs') 'gui'
                Initialize-PcmGuiFolder (Join-Path $g.Root 'logs'); Initialize-PcmGuiFolder $folder
                $g.WaveFile = Write-PcmGuiWave -Folder $folder -Name $values.Batch -Rows $remaining
                $g.WaveKey = $selectionKey
            }
            $values.CsvPath = $g.WaveFile
        }
        if ($action -eq 'Check') {
            $values.Expect = Get-PcmGuiField 'Expect'
            $values.Once = [bool](Get-PcmGuiField 'Once')
        }
    }
    return New-PcmGuiRequest @values
}

function Get-PcmGuiSelectionKey {
    <# What a preview was run on: a run is possible only on exactly the same selection. #>
    param([Parameter(Mandatory)][string]$Prefix, [AllowEmptyCollection()][string[]]$Guids = @())
    if (-not $Guids.Count) { return $Prefix }
    $sorted = @($Guids | ForEach-Object { $_.ToLowerInvariant() } | Sort-Object) -join ','
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($sorted))).Replace('-', '').Substring(0, 16) } finally { $sha.Dispose() }
    return '{0}|{1}|{2}' -f $Prefix, $Guids.Count, $hash
}

function Get-PcmGuiCommandPreview {
    <# The command line equivalent to what the console would run now (settings changed in the window are not on the command line). #>
    $g = $script:Gui
    if (-not $g.Action) { return '' }
    try {
        $request = Get-PcmGuiWindowRequest -Mode 'Apply'
        return (Get-PcmGuiCommand -Root $g.Root -ConfigPath $g.ConfigPath -Request $request -Engine 'pwsh.exe').Display
    } catch { return '' }
}
#endregion

#region Batches and the wave grid
function Update-PcmGuiEnvLock {
    <# A batch only continues in its own environment: choosing one sets the environment and locks it. #>
    $g = $script:Gui
    $row = $g.SelectedBatch
    if ($row -and $row.Environment) {
        if ($null -eq $g.EnvSaved) { $g.EnvSaved = Get-PcmGuiField 'Env' }
        Set-PcmGuiField 'Env' ([string]$row.Environment); $g.EnvLock = [string]$row.Environment
    }
    elseif ($null -ne $g.EnvSaved) { Set-PcmGuiField 'Env' $g.EnvSaved; $g.EnvSaved = $null; $g.EnvLock = $null }
    else { $g.EnvLock = $null }
}

function Update-PcmGuiBatchList {
    <# Rebuilds the list of batches for the current task, keeping the selection when the batch is still there. #>
    $g = $script:Gui; $c = $g.Controls
    $keep = if ($g.SelectedBatch) { [string]$g.SelectedBatch.BatchId } else { '' }
    $g.Updating++
    try {
        $g.BatchRows.Clear()
        $selected = $null
        $kind = Get-PcmGuiBatchKind
        if ($kind -and $g.State.Config -and -not $g.State.JournalError) {
            foreach ($row in @(Get-PcmGuiBatchEntry -State $g.State -Kind $kind)) {
                $g.BatchRows.Add($row)
                if ($keep -and (Test-PcmGuiOrdinal $row.BatchId $keep) -and $row.Selectable) { $selected = $row }
            }
        }
        if ($selected) { $c.BatchList.SelectedItem = $selected } else { $c.BatchList.SelectedItem = $null }
        $g.SelectedBatch = $selected
        Update-PcmGuiEnvLock
    } finally { $g.Updating-- }
    $c.BatchEmpty.Text = if ($g.State.JournalError) { 'The journal cannot be read: ' + $g.State.JournalError } elseif ($kind -and $g.BatchRows.Count -eq 0) {
        switch ($kind) {
            'ContinueConvert' { 'No batch to continue. A new batch starts from the other choice above.' }
            'Rollback' { 'No batch with mailboxes in the cloud. Convert first, or copy the journal file (Setup shows where the tool looks for it) to this computer.' }
            'ContinueRollback' { 'No interrupted rollback.' }
            default { 'No batch yet: the journal is written by Convert and Recover.' }
        }
    } else { '' }
    Update-PcmGuiWaveRows
}

function Update-PcmGuiWaveRows {
    <# The grid of a wave: the mailboxes of the chosen batch that are in the cloud. The ticks are kept while the batch stays the same, and start empty for another batch. #>
    $g = $script:Gui
    $keep = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    $batchId = if ($g.SelectedBatch) { [string]$g.SelectedBatch.BatchId } else { '' }
    if ($batchId -and (Test-PcmGuiOrdinal $batchId $g.WaveBatch)) { foreach ($row in @(Get-PcmGuiTicked)) { [void]$keep.Add([string]$row.Guid) } }
    $g.WaveRows.Clear()
    $g.WaveBatch = $batchId
    if ($g.Action -ne 'Recover' -or -not $g.SelectedBatch -or $g.Mode -ne 'New') { return }
    foreach ($row in @(Get-PcmGuiWaveRow -State $g.State -BatchId $batchId)) {
        $g.WaveRows.Add((New-PcmGuiRow @{ Checked = $keep.Contains([string]$row.Guid); Guid = [string]$row.Guid; Name = [string]$row.Name; Upn = [string]$row.Upn; Type = [string]$row.Type; Result = [string]$row.Result }))
    }
}

function Update-PcmGuiWaveCount {
    $g = $script:Gui; $c = $g.Controls
    $c.WaveCount.Text = '{0} ticked of {1} in the cloud' -f @(Get-PcmGuiTicked).Count, $g.WaveRows.Count
}
#endregion

#region Pages, steps of the run (defined in the view and run files)
. (Join-Path $PSScriptRoot 'PRA.Gui.View.ps1')
. (Join-Path $PSScriptRoot 'PRA.Gui.Run.ps1')
#endregion

#region Window
function Set-PcmGuiWindowBounds {
    param($Window, $WorkArea = $null)
    if ($null -eq $WorkArea) { $WorkArea = [Windows.SystemParameters]::WorkArea }
    $Window.MinWidth = [Math]::Min($Window.MinWidth, $WorkArea.Width)
    $Window.MinHeight = [Math]::Min($Window.MinHeight, $WorkArea.Height)
    $Window.Width = [Math]::Min($Window.Width, $WorkArea.Width)
    $Window.Height = [Math]::Min($Window.Height, $WorkArea.Height)
    $Window.WindowStartupLocation = [Windows.WindowStartupLocation]::Manual
    $Window.Left = $WorkArea.Left + [Math]::Max(0, ($WorkArea.Width - $Window.Width) / 2)
    $Window.Top = $WorkArea.Top + [Math]::Max(0, ($WorkArea.Height - $Window.Height) / 2)
}

function Update-PcmGuiState {
    <# Reads the configuration, the snapshot and the journal again; the run settings go back to the values of the file. #>
    param([switch]$Refresh)
    $g = $script:Gui; $c = $g.Controls
    $c.ConfigPathText.Text = $g.ConfigPath
    $g.State = Read-PcmGuiState -Root $g.Root -ConfigPath $g.ConfigPath -SnapshotId $g.SnapshotId
    Set-PcmGuiDefault
    foreach ($key in $script:GuiSettingKeys) { Set-PcmGuiField $key $g.Defaults[$key] }
    Set-PcmGuiField 'ExpConfig' $(if (Test-PcmGuiSamePath $g.ConfigPath $g.InitialConfigPath) { '' } else { $g.ConfigPath })
    $g.SelectedBatch = $null; $g.EnvLock = $null; $g.EnvSaved = $null
    Reset-PcmGuiPreview
    Invoke-PcmGuiDetection -Refresh:$Refresh
    Update-PcmGuiBatchList
    Update-PcmGuiView
}

function Read-PcmGuiStateKeepingSettings {
    <# The state again after a run (the journal and the snapshot changed), without touching the run settings. #>
    $g = $script:Gui
    $g.State = Read-PcmGuiState -Root $g.Root -ConfigPath $g.ConfigPath -SnapshotId $g.SnapshotId
}

function Select-PcmGuiTask {
    param([string]$Action)
    $g = $script:Gui
    if ($g.Run) { return }
    $g.Action = $Action; $g.Mode = 'New'; $g.SelectedBatch = $null
    Update-PcmGuiEnvLock
    Reset-PcmGuiPreview
    $g.Controls.TaskError.Text = ''
    Update-PcmGuiBatchList
    Update-PcmGuiView
}

function Show-PcmGuiError {
    param([string]$Message)
    $c = $script:Gui.Controls
    Add-PcmGuiLine $Message 'Fail'
    $c.TaskError.Text = $Message; $c.ValidationText.Text = $Message; $c.Footer.Text = $Message
}

function Assert-PcmGuiUnder {
    <# A path reported by the engine is only opened when it is inside the folder of the configuration meant for it (reports, logs). #>
    param([string]$Path, [string]$Folder)
    if (-not $Path) { return '' }
    if (-not $Folder) { throw 'The configuration is not valid: nothing is opened.' }
    $full = [IO.Path]::GetFullPath($Path); $root = [IO.Path]::GetFullPath($Folder).TrimEnd('\') + '\'
    if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { throw ('Refused: {0} is not inside {1}.' -f $full, $root.TrimEnd('\')) }
    return $full
}

function Open-PcmGuiPath {
    param([string]$Path)
    if (-not $Path) { return }
    if (-not (Test-Path -LiteralPath $Path)) { throw "File or folder missing: $Path" }
    $extension = [IO.Path]::GetExtension($Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container) -and -not (Test-PcmGuiOrdinalIn $extension @('.csv', '.html', '.txt', '.log', '.json', '.psd1'))) { throw 'Only reports, logs, configuration files and folders may be opened.' }
    $info = New-Object Diagnostics.ProcessStartInfo
    $info.FileName = $Path; $info.UseShellExecute = $true
    try { $null = [Diagnostics.Process]::Start($info) }
    catch { if (Test-PcmGuiOrdinal $extension '.psd1') { $null = [Diagnostics.Process]::Start('notepad.exe', ('"{0}"' -f $Path)) } else { throw } }
}

function Get-PcmGuiHook {
    <# The hook of this name that the window was given (the tests and the pictures give some), or $null: a hashtable without the key is an error under strict mode. #>
    param([Parameter(Mandatory)][string]$Name)
    $hooks = $script:Gui.Hooks
    if ($hooks -and $hooks.ContainsKey($Name)) { return $hooks[$Name] }
    return $null
}

function Select-PcmGuiOu {
    <# The dialog that lists the OUs of the snapshot. Returns the name picked, or ''. #>
    param([object[]]$Units)
    $g = $script:Gui
    $pick = Get-PcmGuiHook 'PickOu'
    if ($pick) { return [string](& $pick $Units) }
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
 Title="Pick an OU" Width="720" Height="500" WindowStartupLocation="CenterOwner" FontFamily="Segoe UI" Background="{DynamicResource ApplicationBackgroundBrush}">
 <Grid Margin="18"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
 <TextBlock x:Name="Info" TextWrapping="Wrap" Margin="0,0,0,8" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
 <TextBox x:Name="Filter" Grid.Row="1" Margin="0,0,0,8" ToolTip="Filter the name"/>
 <ListBox x:Name="Units" Grid.Row="2" DisplayMemberPath="Text"/>
 <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
 <Button x:Name="Select" Content="Pick" IsDefault="True" Padding="14,6" Margin="0,0,8,0"/>
 <Button Content="Cancel" IsCancel="True" Padding="14,6"/></StackPanel></Grid>
</Window>
'@
    $dialog = [Windows.Markup.XamlReader]::Parse($xaml)
    Set-PcmGuiTheme -Window $dialog -Theme $g.Theme
    if ($g.Form.IsVisible) { $dialog.Owner = $g.Form }
    $items = @($Units | ForEach-Object { [pscustomobject]@{ Name = [string]$_.Name; Text = ('{0}  ({1})' -f $_.Name, (ConvertTo-PcmGuiPlural ([int]$_.Count) 'mailbox' 'mailboxes')) } })
    $dialog.FindName('Info').Text = 'The OUs of the snapshot. Pick one: its sub-OUs are included.'
    $dialog.Tag = $items
    $dialog.FindName('Units').ItemsSource = $items
    $dialog.FindName('Filter').Add_TextChanged({
            param($box)
            $owner = [Windows.Window]::GetWindow($box)
            $text = [string]$box.Text
            $owner.FindName('Units').ItemsSource = @($owner.Tag | Where-Object { $_.Name.IndexOf($text, [StringComparison]::OrdinalIgnoreCase) -ge 0 })
        })
    $dialog.FindName('Select').Add_Click({ param($button) $owner = [Windows.Window]::GetWindow($button); if ($owner.FindName('Units').SelectedItem) { $owner.DialogResult = $true } })
    $dialog.FindName('Units').Add_MouseDoubleClick({ param($list) if ($list.SelectedItem) { [Windows.Window]::GetWindow($list).DialogResult = $true } })
    if ($dialog.ShowDialog()) { return [string]$dialog.FindName('Units').SelectedItem.Name }
    return ''
}

function Show-PcmGuiQuestion {
    <# A yes/no question of the engine (No by default). #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Text)
    $g = $script:Gui
    $ask = Get-PcmGuiHook 'Question'
    if ($ask) { return [bool](& $ask $Title $Text) }
    $buttons = [Windows.MessageBoxButton]::YesNo; $image = [Windows.MessageBoxImage]::Question; $default = [Windows.MessageBoxResult]::No
    $answer = if ($g.Form.IsVisible) { [Windows.MessageBox]::Show($g.Form, $Text, $Title, $buttons, $image, $default) } else { [Windows.MessageBox]::Show($Text, $Title, $buttons, $image, $default) }
    return $answer -eq [Windows.MessageBoxResult]::Yes
}

function Show-PcmGuiNotice {
    param([Parameter(Mandatory)][string]$Text)
    $g = $script:Gui
    $notice = Get-PcmGuiHook 'Notice'
    if ($notice) { [void](& $notice $Text); return }
    $buttons = [Windows.MessageBoxButton]::OK; $image = [Windows.MessageBoxImage]::Information
    if ($g.Form.IsVisible) { [void][Windows.MessageBox]::Show($g.Form, $Text, 'PRA Cloud Mailbox', $buttons, $image) } else { [void][Windows.MessageBox]::Show($Text, 'PRA Cloud Mailbox', $buttons, $image) }
}

function Invoke-PcmGuiClick {
    param([string]$Name)
    $g = $script:Gui; $c = $g.Controls
    try {
        $c.TaskError.Text = ''; $c.ValidationText.Text = ''
        switch ($Name) {
            'BtnRecheck' { Invoke-PcmGuiDetection -Refresh; $c.RecheckNote.Text = 'Checked at ' + [datetime]::Now.ToString('HH:mm', [Globalization.CultureInfo]::InvariantCulture); Update-PcmGuiView }
            'BtnChange' { $c.SettingsCard.Visibility = Get-PcmGuiVisibility ($c.SettingsCard.Visibility -ne 'Visible') }
            'BtnDoneSettings' { $c.SettingsCard.Visibility = 'Collapsed' }
            'BtnResetAll' {
                foreach ($key in $script:GuiSettingKeys) { if ($key -ne 'ExpConfig' -and -not ($key -eq 'Env' -and $g.EnvLock)) { Reset-PcmGuiField $key } }
                Invoke-PcmGuiChange
            }
            'BtnOpenConfig' { Open-PcmGuiPath $g.ConfigPath }
            'BtnReload' { if (-not $g.Run) { Update-PcmGuiState } }
            'BtnPickOu' { $picked = Select-PcmGuiOu -Units @(Get-PcmGuiOuList -State $g.State); if ($picked) { Set-PcmGuiField 'Ou' $picked; Invoke-PcmGuiChange } }
            'BtnCheckGroup' { Start-PcmGuiGroupLookup }
            'BtnPickCsv' {
                $dialog = New-Object Microsoft.Win32.OpenFileDialog; $dialog.Filter = 'List (*.csv;*.txt)|*.csv;*.txt|All files (*.*)|*.*'
                if ($dialog.ShowDialog($g.Form)) { Set-PcmGuiField 'Csv' $dialog.FileName; Invoke-PcmGuiChange }
            }
            'BtnPickConfig' {
                $dialog = New-Object Microsoft.Win32.OpenFileDialog; $dialog.Filter = 'Configuration (*.psd1)|*.psd1'
                if ($dialog.ShowDialog($g.Form)) { $g.ConfigPath = $dialog.FileName; Update-PcmGuiState }
            }
            'BtnTickAll' { foreach ($row in $g.WaveRows) { $row.Checked = $true }; Invoke-PcmGuiChange }
            'BtnUntickAll' { foreach ($row in $g.WaveRows) { $row.Checked = $false }; Invoke-PcmGuiChange }
            'BtnFixHome' { Show-PcmGuiPage 1; $c.SettingsCard.Visibility = 'Visible' }
            'BtnEditSettings' { Show-PcmGuiPage 1; $c.SettingsCard.Visibility = 'Visible' }
            'BtnBack' { if (-not $g.Run -and $g.Page -gt 1 -and -not (Test-PcmGuiApplied)) { Show-PcmGuiPage ($g.Page - 1) } }
            'BtnPreview' { Start-PcmGuiRun -Mode Preview }
            'BtnSchedule' { Start-PcmGuiSchedule }
            'BtnStop' { Stop-PcmGuiRun }
            'BtnNext' {
                switch ($g.Page) {
                    1 { if (-not $g.NetJob) { Show-PcmGuiPage 2 } }
                    2 { Show-PcmGuiPage 3 }
                    3 { if ($g.Action -eq 'Check') { Start-PcmGuiRun -Mode Preview -Purpose Run } else { Start-PcmGuiRun -Mode Apply }; Show-PcmGuiPage 4 }
                    4 { if ($g.LastRun) { Show-PcmGuiPage 5 } }
                    5 {
                        $g.Action = ''; $g.Mode = 'New'; $g.SelectedBatch = $null; $g.Model = $null; $g.LastRun = $null
                        Update-PcmGuiEnvLock; Reset-PcmGuiPreview
                        Show-PcmGuiPage 2
                    }
                }
            }
            'BtnCopyNextCmd' { [Windows.Clipboard]::SetText($c.NextCmd.Text) }
            'BtnCopyContinueCmd' { [Windows.Clipboard]::SetText($c.ContinueCmd.Text) }
            'BtnCopyBatch' { [Windows.Clipboard]::SetText($c.BatchIdText.Text) }
            'BtnReport' { Open-PcmGuiPath (Assert-PcmGuiUnder ([string](Get-PcmGuiValue $g.LastResult 'htmlReport' '')) $(if ($g.State.Config) { [string]$g.State.Config.Report.Folder } else { '' })) }
            'BtnReportLink' { Open-PcmGuiPath (Assert-PcmGuiUnder ([string](Get-PcmGuiValue $g.LastResult 'htmlReport' '')) $(if ($g.State.Config) { [string]$g.State.Config.Report.Folder } else { '' })) }
            'BtnGlobalReport' { Open-PcmGuiPath (Assert-PcmGuiUnder ([string](Get-PcmGuiValue $g.LastResult 'globalReport' '')) $(if ($g.State.Config) { [string]$g.State.Config.Report.Folder } else { '' })) }
            'BtnLog' { Open-PcmGuiPath (Assert-PcmGuiUnder ([string](Get-PcmGuiValue $g.LastResult 'logFile' '')) $(if ($g.State.Config) { [string]$g.State.Config.Logging.Folder } else { '' })) }
            default {
                if ($Name.StartsWith('R_', [StringComparison]::Ordinal)) {
                    $key = $Name.Substring(2)
                    if (Test-PcmGuiOrdinal $key 'ExpConfig') { $g.ConfigPath = $g.InitialConfigPath; Update-PcmGuiState } else { Reset-PcmGuiField $key; Invoke-PcmGuiChange }
                }
            }
        }
    } catch { Show-PcmGuiError $_.Exception.Message }
}

function Invoke-PcmGuiChoice {
    <# A mode, a filter or the theme changed. #>
    param([string]$Name)
    $g = $script:Gui
    if (-not $g -or $g.Updating -gt 0) { return }
    if ($Name.StartsWith('Mode_', [StringComparison]::Ordinal)) {
        if ($g.Run) { return }
        $g.Mode = $Name.Substring(5); $g.SelectedBatch = $null
        Update-PcmGuiEnvLock; Reset-PcmGuiPreview; Update-PcmGuiBatchList; Update-PcmGuiView
        return
    }
    if ($Name.StartsWith('Fl_', [StringComparison]::Ordinal)) { $g.Filter = $Name.Substring(3); $g.VerView.Refresh(); return }
    if ($Name.StartsWith('Th_', [StringComparison]::Ordinal)) {
        $info = Initialize-PcmGuiTheme $Name.Substring(3)
        $g.Theme = $info; Set-PcmGuiTheme $g.Form $info
        return
    }
    if ($Name.StartsWith('C_Scope_', [StringComparison]::Ordinal)) {
        Set-PcmGuiField 'Scope' $Name.Substring(8)
        Invoke-PcmGuiChange
        return
    }
    if ($Name.StartsWith('S_Expect_', [StringComparison]::Ordinal) -or $Name.StartsWith('S_Wave_', [StringComparison]::Ordinal)) { Invoke-PcmGuiChange; return }
    Invoke-PcmGuiChange
}

function New-PcmGuiXmlNamespace {
    param([xml]$Xml)
    $manager = New-Object Xml.XmlNamespaceManager $Xml.NameTable
    $manager.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    return , $manager
}

function New-PcmGuiWindow {
    <#
    .SYNOPSIS
        Builds and binds the window without showing it and without starting anything. -PrerequisiteProbe replaces the real detection (tests, pictures); with it, Microsoft 365 and Exchange
        are not tested unless -NetworkProbe (a task per host) and -ExchangeProbe (a task for the Exchange server) are given too. -GroupProbe replaces the group lookup of the engine. -Hooks answer the questions and the dialogs.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [string]$ConfigPath = '', [string]$Version = '', [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System',
        [scriptblock]$PrerequisiteProbe, [scriptblock]$NetworkProbe, [double]$NetworkSeconds = 3, [scriptblock]$ExchangeProbe, [double]$ExchangeSeconds = 15, [scriptblock]$GroupProbe, [hashtable]$Hooks = @{})
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) { throw 'The console needs a single-threaded apartment: start PowerShell with -STA (the default of pwsh.exe).' }
    if ($script:Gui -and $script:Gui.Run) { throw 'Cannot replace a window while its engine is running.' }
    if ($script:Gui -and $script:Gui.Form.IsVisible) { throw 'A PRA window is already open in this host.' }
    if (-not $ConfigPath) { $ConfigPath = Join-Path $Root 'config\PraCloudMailbox.config.psd1' }
    $themeInfo = Initialize-PcmGuiTheme $Theme
    $text = [IO.File]::ReadAllText((Join-Path $script:GuiModuleRoot 'PRA.Gui.xaml'))
    $xml = [xml]$text
    $reader = New-Object Xml.XmlNodeReader $xml
    try { $form = [Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Dispose() }
    Set-PcmGuiWindowBounds -Window $form
    Set-PcmGuiTheme $form $themeInfo
    $controls = @{}
    foreach ($node in $xml.SelectNodes('//*[@x:Name]', (New-PcmGuiXmlNamespace $xml))) {
        $name = $node.GetAttribute('Name', 'http://schemas.microsoft.com/winfx/2006/xaml')
        $controls[$name] = $form.FindName($name)
    }
    $timer = New-Object Windows.Threading.DispatcherTimer
    $timer.Interval = [timespan]::FromMilliseconds(250)
    $probe = if ($PrerequisiteProbe) { $PrerequisiteProbe } else { { param($Request) Get-PcmGuiPrerequisite -Request $Request } }
    $netProbe = if ($NetworkProbe) { $NetworkProbe } elseif ($PrerequisiteProbe) { $null } else { { param($HostName, $Seconds) Start-PcmGuiHttpsProbe $HostName $Seconds } }
    $exchangeProbe = if ($ExchangeProbe) { $ExchangeProbe } elseif ($PrerequisiteProbe) { $null } else { { param($Server, $Seconds) Start-PcmGuiExchangeProbe $Server $Seconds } }
    $script:Gui = @{
        Form = $form; Controls = $controls; Timer = $timer; Root = [IO.Path]::GetFullPath($Root); ConfigPath = $ConfigPath; InitialConfigPath = $ConfigPath; SnapshotId = [long]0
        State = $null; Run = $null; Pending = $null; LastResult = $null; LastRun = $null; Model = $null; Updating = 1; InTick = $false; Page = 1
        Action = ''; Mode = 'New'; SelectedBatch = $null; EnvLock = $null; EnvSaved = $null; Defaults = $null; Detect = $null; Probe = $probe
        NetProbe = $netProbe; NetSeconds = $NetworkSeconds; NetJob = $null; NetResult = $null; GroupProbe = $GroupProbe; GroupLookup = $null
        ExchProbe = $exchangeProbe; ExchSeconds = $ExchangeSeconds; ExchJob = $null; ExchResult = $null
        Previewed = $false; PreviewError = ''; Filter = 'All'; Hooks = $Hooks; WaveFile = ''; WaveKey = ''; WaveBatch = ''; PreviewResult = $null
        Lines = (New-Object 'Collections.ObjectModel.ObservableCollection[object]')
        LightRows = (New-Object 'Collections.ObjectModel.ObservableCollection[object]'); BatchRows = (New-Object 'Collections.ObjectModel.ObservableCollection[object]')
        VerRows = (New-Object 'Collections.ObjectModel.ObservableCollection[object]'); NextRows = (New-Object 'Collections.ObjectModel.ObservableCollection[object]')
        TimelineRows = (New-Object 'Collections.ObjectModel.ObservableCollection[object]'); WaveRows = (New-Object 'Collections.ObjectModel.ObservableCollection[object]')
        Theme = $themeInfo; ThemeName = $Theme
    }
    $g = $script:Gui; $c = $controls
    $c.VersionText.Text = $(if ($Version) { 'v' + $Version } else { '' })
    $c.HeaderSub.ToolTip = 'Computer: ' + $env:COMPUTERNAME + ' · PowerShell ' + $PSVersionTable.PSVersion
    $c.ActivityLog.ItemsSource = $g.Lines; $c.Lights.ItemsSource = $g.LightRows; $c.BatchList.ItemsSource = $g.BatchRows; $c.NextSteps.ItemsSource = $g.NextRows
    $c.Timeline.ItemsSource = $g.TimelineRows; $c.WaveGrid.ItemsSource = $g.WaveRows; $c.VerList.ItemsSource = $g.VerRows
    $g.VerView = [Windows.Data.CollectionViewSource]::GetDefaultView($g.VerRows)
    $g.VerView.Filter = [Predicate[object]] {
        param($row)
        switch ($script:Gui.Filter) {
            'User' { return (Test-PcmGuiOrdinal $row.Type 'User') } 'Shared' { return (Test-PcmGuiOrdinal $row.Type 'Shared') }
            'Warn' { return (Test-PcmGuiOrdinalIn $row.Status @('Warn', 'Fail')) } default { return $true }
        }
    }
    foreach ($name in @($controls.Keys)) {
        $control = $controls[$name]
        if ($control -is [Windows.Controls.Button]) {
            $control.Tag = $name
            $control.Add_Click({ param($origin) Invoke-PcmGuiClick ([string]$origin.Tag) })
        }
        elseif ($control -is [Windows.Controls.RadioButton]) {
            if ($name.StartsWith('T_', [StringComparison]::Ordinal)) { $control.Tag = $name.Substring(2); $control.Add_Checked({ param($origin) if ($script:Gui.Updating -eq 0) { Select-PcmGuiTask ([string]$origin.Tag) } }) }
            elseif ($name.StartsWith('Nav', [StringComparison]::Ordinal) -and $name.Length -eq 4) {
                $control.Add_Checked({ param($origin)
                        $gui = $script:Gui
                        if ($gui.Updating -gt 0) { return }
                        $number = [int]([string]$origin.Name).Substring(3)
                        if (($number -ne $gui.Page) -and (Test-PcmGuiPageReachable $number)) { Show-PcmGuiPage $number }
                        else { $gui.Updating++; try { $gui.Controls[('Nav' + $gui.Page)].IsChecked = $true } finally { $gui.Updating-- } }
                    })
            }
            else { $control.Tag = $name; $control.Add_Checked({ param($origin) Invoke-PcmGuiChoice ([string]$origin.Tag) }) }
        }
        elseif ($control -is [Windows.Controls.TextBox] -and -not $control.IsReadOnly -and $name.StartsWith('F_', [StringComparison]::Ordinal)) {
            $control.Add_TextChanged({ Invoke-PcmGuiChange })
        }
    }
    $c.F_Once.Add_Checked({ Invoke-PcmGuiChange }); $c.F_Once.Add_Unchecked({ Invoke-PcmGuiChange })
    # The mailbox and the group are looked up when the field loses the focus or Enter is pressed; typing again clears the answer.
    $c.F_Group.Add_TextChanged({ if ($script:Gui.Updating -eq 0) { Set-PcmGuiGroupInfo '' } })
    $c.F_Group.Add_LostFocus({ if ($script:Gui.Updating -eq 0 -and $script:Gui.Page -eq 2) { try { Start-PcmGuiGroupLookup } catch { Set-PcmGuiGroupInfo $_.Exception.Message 'Warn' } } })
    $c.F_Group.Add_KeyDown({ if ([int]$args[1].Key -eq [int][Windows.Input.Key]::Return) { try { Start-PcmGuiGroupLookup } catch { Set-PcmGuiGroupInfo $_.Exception.Message 'Warn' } } })
    $c.ConfirmBox.Add_TextChanged({ if ($script:Gui.Updating -eq 0) { Update-PcmGuiView } })
    $c.WaveGrid.AddHandler([Windows.Controls.Primitives.ToggleButton]::ClickEvent, [Windows.RoutedEventHandler] {
            $box = $args[1].OriginalSource
            if ($box -is [Windows.Controls.CheckBox] -and $box.DataContext) { $box.DataContext.Checked = [bool]$box.IsChecked; Invoke-PcmGuiChange }
        })
    $c.BatchList.Add_SelectionChanged({
            $gui = $script:Gui
            if ($gui.Updating -gt 0 -or $gui.Run) { return }
            $gui.SelectedBatch = $gui.Controls.BatchList.SelectedItem
            # A check that follows a rollback expects the mailboxes back on-premises, one that follows a conversion expects them in the cloud.
            if ($gui.SelectedBatch -and $gui.Action -eq 'Check') { Set-PcmGuiField 'Expect' $(if ($gui.SelectedBatch.Action -eq 'Recover') { 'Deprovisioned' } else { 'Provisioned' }) }
            Update-PcmGuiEnvLock; Reset-PcmGuiPreview; Update-PcmGuiWaveRows; Update-PcmGuiView
        })
    $c.StepList.Add_SelectionChanged({
            $gui = $script:Gui
            if ($gui.Updating -gt 0 -or -not $gui.Model) { return }
            $gui.Model.Selected = $gui.Controls.StepList.SelectedItem; $gui.Model.Follow = $false
            Show-PcmGuiStepDetail $gui.Model.Selected
        })
    $c.DetObjList.Add_SelectionChanged({ $gui = $script:Gui; if ($gui.Model) { Show-PcmGuiObjectCard $gui.Controls.DetObjList.SelectedItem } })
    $timer.Add_Tick({ Invoke-PcmGuiTick })
    $form.Add_Closing({
            if ($script:Gui.Run) {
                $args[1].Cancel = $true
                Add-PcmGuiLine 'This window cannot be closed while a run is going: every change is journaled, wait for the end of the run, or stop it with "Stop after the current object".' 'Warn'
            }
        })
    $form.Add_Closed({ $script:Gui.Timer.Stop(); $script:Gui.Pending = $null; $script:Gui.NetJob = $null; $script:Gui.ExchJob = $null; if ($script:Gui.GroupLookup) { try { $script:Gui.GroupLookup.Process.Kill() } catch { Write-Verbose $_.Exception.Message }; $script:Gui.GroupLookup = $null }; if ($script:GuiNetHttp) { $script:GuiNetHttp.Dispose(); $script:GuiNetHttp = $null } })
    foreach ($name in @('System', 'Light', 'Dark')) { $c['Th_' + $name].IsChecked = (Test-PcmGuiOrdinal $name $Theme) }
    $c.Nav1.IsChecked = $true
    Set-PcmGuiField 'Source' 'Single'
    $g.Updating = 0
    Update-PcmGuiState -Refresh
    return $g
}

function Show-PcmGui {
    <#
    .SYNOPSIS
        Opens the console (Invoke-PraCloudMailbox.ps1 -Gui) and returns when it is closed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root, [string]$ConfigPath = '', [string]$Version = '')
    $window = New-PcmGuiWindow -Root $Root -ConfigPath $ConfigPath -Version $Version
    # The files of the runs older than 30 days (the logs and reports of the engine stay).
    try {
        $limit = (Get-Date).AddDays(-30)
        $folder = Join-Path (Join-Path $window.Root 'logs') 'gui'
        if (Test-Path -LiteralPath $folder) { Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $limit } | Remove-Item -Force -ErrorAction SilentlyContinue }
    } catch { Write-Verbose $_.Exception.Message }
    # Ctrl+C in the console would stop the command that owns the window, and with it every handler of the window.
    $previousCtrlC = $null
    try { if (-not [Console]::IsInputRedirected) { $previousCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } } catch { $previousCtrlC = $null }
    try { $null = $window.Form.ShowDialog() }
    finally {
        if ($null -ne $previousCtrlC) { try { [Console]::TreatControlCAsInput = $previousCtrlC } catch { Write-Verbose $_.Exception.Message } }
        if ($script:Gui -and $script:Gui.Timer) { $script:Gui.Timer.Stop() }
        $script:Gui = $null
    }
}
#endregion

Export-ModuleMember -Function Show-PcmGui, New-PcmGuiWindow, New-PcmGuiRequest, Get-PcmGuiCommand, Get-PcmGuiEngine, Read-PcmGuiEvents, Read-PcmGuiState, Get-PcmGuiBatchSummary,
    Get-PcmGuiBatchEntry, Get-PcmGuiWaveRow, Write-PcmGuiWave, Get-PcmGuiSelectionKey, Test-PcmGuiConfirmation, Format-PcmGuiTime, Get-PcmGuiPrerequisite, Get-PcmGuiRole,
    Get-PcmGuiNetworkWhy, Get-PcmGuiNetworkVerdict, ConvertTo-PcmGuiNativeArgument, New-PcmGuiRunConfig, ConvertTo-PcmGuiDataLiteral, Find-PcmGuiMailbox, Get-PcmGuiOuList, Get-PcmGuiOuMatch,
    Get-PcmGuiGroupInSnapshot, Start-PcmGuiChildProcess, Get-PcmGuiNetworkLight, Get-PcmGuiExchangeLight, Get-PcmGuiSnapshotInfo, Get-PcmGuiSignInLight

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDi8dWVSCrXQATM
# IYn8YlJ4yP/mfKMGrIw0715LEOmVDqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDjTDpe4tPr
# FmeUMN2EfBmhgBue/IIs6sjffMRVf22OCzANBgkqhkiG9w0BAQEFAASCAgCIoLkv
# zP2dkPeGUpduGtlaVf/UUDCIHUWQTB5XA82RIhrm3zN1YXJvxbkgAEbiN3ErL45s
# Cc6x2XLPVGolx/P8FluUZjjzq/ImXFoo1IyCZEstf+5nffl070VZhl5cL+dRdjrZ
# bTS/ZIwjkHgsP/ORLQWxWZLipO9o4yIHICqQe0PakSzesRisRbxz80CHOw2kmZBN
# cTPvB5VjUxhpJ80wCfLFJFxpOoeCirD/uy3KrANeIZOSFbUPgRPxJoE3LBfFDDGx
# BEuVN+yK4s0fc/ZG83QYHznJEbJvrOwss5fmWSVHfBjj6a0RhhuuiikfjkBhaHYl
# M36wHE/g/062isigKOs0mWARNN5S9K0blTIgrZuAvUf/TZEuIh463VENzwEvPw/n
# eZXfy9ycNYZHokm8W+QcAdX+mXdwAa9QpXPiPm56vTirz54gf8HMLUtBnD4krNdH
# PlBKuqmFna5aHNjaEqF1WZORdgGeBh6AfRsT25RgkJhNZ6MZtsMjxbMxm6pHoBev
# qE7ri4FxullMhI7XIc/YCX/nQAHz/Y+v2hbShmHG1ZEDNiZAk+uD0IlHpytWFvHA
# EXi9J9b0w6JDWmPjnG2r4MRRaNSqgXR4c9hVAmYiHTPqzFSn6HG+CBAX4a/+Eqs/
# vJIMAcjDkUE9effKjWUMi53xKIIbn7U7q2vwGaGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEILxlksRU5b2WMKbu/lD3s/z1xv5S8Z9pRqbH4yYehOywAggA+rhDaTIU
# eBgPMjAyNjExMTUyMDA2NDBaMAMCAQECEDFlmFLG5c+OcGXAUOzvrPSgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTIwMDY0MFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIGGT
# x0HhWV+nqM5C9eK1mbvsHIKOAu8DjLPSVB+8aPDgMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAoJKFhxZyCeTo8WN8Zsfy+kQ4vdWCgcmy7zP1iWpy8+wCIEZa
# PCiUkK/DcS5Ojb+f/rYg6Cb9T2LpE8sgrIx3/XUd
# SIG # End signature block
