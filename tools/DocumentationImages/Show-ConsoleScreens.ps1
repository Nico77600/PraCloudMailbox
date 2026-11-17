<#
.SYNOPSIS
    Photographs the real PRA Cloud Mailbox console against a synthetic runtime, for the documentation (repository tool, never shipped).
.DESCRIPTION
    Called by tools\New-DocumentationImages.ps1, once per scenario and theme, in its own PowerShell 7.5+ STA process. It builds the real window (module\PRA.Gui.psm1)
    in a copy of the tool that holds the synthetic engine StubEngine.ps1 instead of the real script, then drives the five pages like an operator and renders them to PNG.
    Synthetic values only, injected here and never read from the machine: computer SRV-PRA01, account CONTOSO\pra.operator, tool folder D:\PRA\PraCloudMailbox. Every text of the window is checked before each photo: a real machine name, user name or local path stops the run.
.PARAMETER Shots
    Output folder of the PNG files (named console-<shot>-<light|dark>.png; the shots of the collect scenario are named console-collect-<shot>-...).
.PARAMETER Scenario
    cloud     a computer that can run every action: setup, run settings, the tasks, preview, run, next steps
    collect   a computer that can only run Collect (no cloud modules): setup, the Collect task, its preview, run and next steps
    problems  a run that fails before it changes anything, then a Check that finds an object that is not ready
.PARAMETER Repository
    Root of the repository (default: two folders above this script).
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
#Requires -Version 7.5
param(
    [Parameter(Mandatory)][string]$Shots,
    [ValidateSet('Light', 'Dark')][string]$Theme = 'Light',
    [ValidateSet('cloud', 'collect', 'problems')][string]$Scenario = 'cloud',
    [string]$Repository
)
$ErrorActionPreference = 'Stop'
# Real values of this machine, kept only to make sure that none of them reaches a picture.
$real = @($env:COMPUTERNAME, $env:USERNAME, $env:USERDOMAIN, $env:USERPROFILE, $env:TEMP, $env:LOCALAPPDATA, $env:APPDATA) | Where-Object { $_ -and $_.Length -ge 3 }
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US'); [Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$env:COMPUTERNAME = 'SRV-PRA01'; $env:USERDOMAIN = 'CONTOSO'; $env:USERNAME = 'pra.operator'; $env:USERDNSDOMAIN = 'CONTOSO.COM'
if (-not $Repository) { $Repository = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent }
$repo = (Resolve-Path $Repository).Path
$work = Join-Path ([IO.Path]::GetTempPath()) ('pcm-doc-' + $Scenario + '-' + $Theme + '-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
try {
    . (Join-Path $PSScriptRoot 'New-ToolCopy.ps1')
    . (Join-Path $PSScriptRoot 'ConsoleFixture.ps1')
    $root = New-DocumentationToolCopy -Repository $repo -Destination $work
    Import-Module (Join-Path $root 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $root 'module\PRA.Store.psm1') -Force
    Import-Module (Join-Path $root 'module\PRA.Gui.psm1') -Force
    $fixture = New-ConsoleFixture -ToolRoot $root
    # A run that fails before it changes anything has no batch: the window of that scenario starts without a journal.
    if ($Scenario -eq 'problems') { [IO.File]::Delete((Join-Path $root 'data\PraCloudMailbox-journal.db')) }
    $syntheticRoot = 'D:\PRA\PraCloudMailbox'

    # The detection is the real one (Get-PcmGuiPrerequisite and the light builders of the window), fed with synthetic installed modules, a synthetic certificate and a synthetic scheduled task:
    # nothing leaves this process, Microsoft 365 and Exchange included.
    $guiModule = Get-Module PRA.Gui
    $fakeModules = @([pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = [version]'2.41.1' }, [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = [version]'3.10.1' })
    $fakeCertificate = [pscustomobject]@{ Certificate = [pscustomobject]@{ Subject = 'CN=PRA-Cloud-App'; NotAfter = [datetime]::Now.AddDays(230); HasPrivateKey = $true } }
    $isCloud = ($Scenario -ne 'collect')
    $probe = {
        param($Request)
        & $guiModule {
            param($Request, $IsCloud, $Modules, $Certificate)
            $r = @{}; foreach ($key in $Request.Keys) { $r[$key] = $Request[$key] }
            $r.Edition = 'Core'; $r.Version = [version]'7.5.1'
            $r.Task = [pscustomobject]@{ Found = $true; Text = ('Scheduled task "PRA Collect" found. Last run {0}: succeeded.' -f (Get-Date).AddHours(-14).ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)) }
            if ($IsCloud) { $r.Modules = $Modules; $r.Certificate = $Certificate } else { $r.Modules = @() }
            $d = Get-PcmGuiPrerequisite -Request $r
            $items = [Collections.Generic.List[object]]::new()
            foreach ($light in $d.Lights) { $items.Add($light) }
            $net = if ($d.CanCloudBase) { Get-PcmGuiNetworkLight -State Ok } else { @{ Status = 'NA'; Detail = 'Not tested: the lights above must be green first.' } }
            $items.Insert([Math]::Min(2, $items.Count), [ordered]@{ Name = $script:GuiNetLight; Status = $net.Status; Detail = $net.Detail })
            $exchange = Get-PcmGuiExchangeLight -State Ok -Server ([string]$d.ExchangeServer)
            $at = $items.Count
            for ($k = 0; $k -lt $items.Count; $k++) { if ($items[$k].Name -eq 'Daily Collect') { $at = $k; break } }
            $items.Insert($at, [ordered]@{ Name = $script:GuiExchangeLight; Status = $exchange.Status; Detail = $exchange.Detail })
            Add-Member -InputObject $d -NotePropertyName Lights -NotePropertyValue $items.ToArray() -Force
            return $d
        } $Request $isCloud $fakeModules $fakeCertificate
    }.GetNewClosure()
    $g = New-PcmGuiWindow -Root $root -ConfigPath $fixture.Config -Version '2.1.1' -Theme $Theme -PrerequisiteProbe $probe -Hooks @{ NoConsole = $true }
    function Pump($ms = 120) { $f = New-Object Windows.Threading.DispatcherFrame; $t = New-Object Windows.Threading.DispatcherTimer; $t.Interval = [timespan]::FromMilliseconds($ms); $t.Tag = $f; $t.Add_Tick({ param($s) $s.Tag.Continue = $false; $s.Stop() }); $t.Start(); [Windows.Threading.Dispatcher]::PushFrame($f) }
    function Walk($element, [scriptblock]$visit) { & $visit $element; $n = [Windows.Media.VisualTreeHelper]::GetChildrenCount($element); for ($i = 0; $i -lt $n; $i++) { Walk ([Windows.Media.VisualTreeHelper]::GetChild($element, $i)) $visit } }
    function Mask($text) { if (-not $text) { return $text }; ([string]$text).Replace($root, $syntheticRoot).Replace([string]$PSVersionTable.PSVersion, '7.5.1') }
    $leak = (@($real | ForEach-Object { [regex]::Escape($_) }) + @('REDMOND', 'DESKTOP-', 'C:\\Users', '\\AppData\\', 'scratchpad', 'Temp\\claude')) -join '|'
    $null = New-Item -ItemType Directory -Path $Shots -Force
    $theme = $Theme.ToLowerInvariant()
    function Shot([string]$name, [int]$Height = 940) {
        $g.Form.Height = $Height; Pump; $g.Form.UpdateLayout(); Pump; Pump
        $bad = New-Object 'Collections.Generic.List[string]'
        $restore = New-Object 'Collections.Generic.List[object]'
        # The texts are masked for the picture only: a text box gets its own text back (without the window noticing the change).
        $g.Updating++
        try {
            Walk $g.Controls.Root { param($e)
                if ($e -is [Windows.Controls.TextBlock]) { $m = Mask $e.Text; if ($m -ne $e.Text) { $e.Text = $m }; if ($e.Text -match $leak) { $bad.Add($e.Text) } }
                elseif ($e -is [Windows.Controls.TextBox]) { $m = Mask $e.Text; if ($m -ne $e.Text) { $restore.Add(@($e, $e.Text)); $e.Text = $m }; if ($e.Text -match $leak) { $bad.Add($e.Text) } }
                elseif ($e -is [Windows.Controls.ContentControl] -and $e.Content -is [string]) { $m = Mask $e.Content; if ($m -ne $e.Content) { $e.Content = $m }; if ($e.Content -match $leak) { $bad.Add($e.Content) } }
            }
            if ($bad.Count) { throw "Real value in the window ($name): $($bad[0])" }
            $g.Form.UpdateLayout()
            $el = $g.Controls.Root; $w = [int][Math]::Ceiling($el.ActualWidth); $h = [int][Math]::Ceiling($el.ActualHeight)
            $rtb = [Windows.Media.Imaging.RenderTargetBitmap]::new($w, $h, 96, 96, [Windows.Media.PixelFormats]::Pbgra32); $rtb.Render($el)
        } finally { foreach ($pair in $restore) { $pair[0].Text = $pair[1] }; $g.Updating-- }
        $png = [Windows.Media.Imaging.PngBitmapEncoder]::new(); $png.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $file = Join-Path $Shots ("console-$name-$theme.png"); $fs = [IO.File]::Create($file); try { $png.Save($fs) } finally { $fs.Close() }
        "shot console-$name-$theme"
    }
    $script:c = $g.Controls
    function Click($n) { $script:c[$n].RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Button]::ClickEvent))) }
    function WaitRun { $d = [datetime]::UtcNow.AddSeconds(120); while ($g.Run -and [datetime]::UtcNow -lt $d) { Pump 100 }; if ($g.Run) { throw 'timeout' }; Pump 200 }
    function Hold($flag) { [IO.File]::WriteAllText((Join-Path $root "hold.$flag"), '1') }
    function Release($flag) { Remove-Item (Join-Path $root "hold.$flag") -ErrorAction SilentlyContinue }
    function Scenario($holdAfter) { $s = [IO.File]::ReadAllText((Join-Path $root 'scenario.json')) | ConvertFrom-Json; $s.holdAfter = $holdAfter; [IO.File]::WriteAllText((Join-Path $root 'scenario.json'), ($s | ConvertTo-Json -Depth 6)) }
    $g.Form.WindowStartupLocation = 'Manual'; $g.Form.Left = 20; $g.Form.Top = 10; $g.Form.ShowActivated = $false
    $g.Form.Show(); Pump 400
    function Choose($task) { $c['T_' + $task].IsChecked = $true; Pump }
    function Preview { Click 'BtnPreview'; WaitRun }
    function SetBlocked([string]$name, [string]$message) { $s = [IO.File]::ReadAllText((Join-Path $root 'scenario.json')) | ConvertFrom-Json; foreach ($o in $s.objects) { if ($o.upn -eq $name) { $o.blocked = $message } }; [IO.File]::WriteAllText((Join-Path $root 'scenario.json'), ($s | ConvertTo-Json -Depth 6)) }
    if ($Scenario -eq 'collect') {
        Shot 'collect-setup'
        Click 'BtnNext'; Choose 'Collect'
        Shot 'collect-task'
        Click 'BtnNext'; Preview
        $c.ConfirmBox.Text = 'collect'; Pump
        Shot 'collect-verify-preview'
        Click 'BtnNext'; WaitRun
        Shot 'collect-run-done'
        Click 'BtnNext'; Pump
        Shot 'collect-next'
    } elseif ($Scenario -eq 'problems') {
        # A run that fails before it changes anything (the certificate is not on this computer), then a Check that finds an object that is not ready.
        Click 'BtnNext'; Choose 'Convert'; $c.S_Source_Single.IsChecked = $true; $c.F_Identity.Text = 'adele.vance@contoso.com'; Pump
        Click 'BtnNext'; Preview
        $c.ConfirmBox.Text = 'CONVERT'; Pump
        [IO.File]::WriteAllText((Join-Path $root 'fail.flag'), '1')
        Click 'BtnNext'; WaitRun
        Shot 'run-failed'
        Click 'BtnNext'; Pump
        Shot 'next-failed'
        [IO.File]::Delete((Join-Path $root 'fail.flag'))
        SetBlocked 'debra.berger@contoso.com' 'No Exchange Online recipient: the on-premises mailbox is not represented in Exchange Online (Exchange hybrid / sync).'
        Click 'BtnNext'; Pump
        Choose 'Check'; $c.S_Source_All.IsChecked = $true; Pump
        Click 'BtnNext'; Pump
        Click 'BtnNext'; WaitRun
        $c.DetObjList.SelectedIndex = 16; $c.DetObjList.ScrollIntoView($c.DetObjList.SelectedItem); Pump 300
        Shot 'run-check-blocked'
        Click 'BtnNext'; Pump
        Shot 'next-check-blocked'
    } else {
        Shot 'setup'
        Click 'BtnChange'; Pump; $c.S_Scope_UsersOnly.IsChecked = $true; Pump; $c.SettingsCard.BringIntoView(); Pump
        Shot 'setup-settings' 975
        $c.S_Scope_All.IsChecked = $true; $c.S_Signin_Interactive.IsChecked = $true; Pump 400; $c.SettingsCard.BringIntoView(); Pump
        Shot 'setup-signin' 975
        $c.S_Signin_Certificate.IsChecked = $true; Pump 400
        Click 'BtnDoneSettings'; $c.Page1.ScrollToTop(); Pump
        Click 'BtnNext'; Choose 'Convert'; $c.S_Source_Single.IsChecked = $true; $c.F_Identity.Text = 'adele.vance@contoso.com'; Pump 300
        Shot 'task-convert-one'
        $c.S_Source_Ou.IsChecked = $true; $c.F_Ou.Text = 'contoso.com/Paris/Users'; Pump 300
        Shot 'task-convert-ou'
        [IO.File]::WriteAllText((Join-Path $root 'data\wave1.csv'), "Identity`r`nadele.vance@contoso.com`r`nalex.wilber@contoso.com`r`nmegan.bowen@contoso.com`r`n")
        $c.S_Source_Csv.IsChecked = $true; $c.F_Csv.Text = (Join-Path $root 'data\wave1.csv'); Pump 300
        Shot 'task-convert-csv'
        $c.S_Source_Group.IsChecked = $true; $c.F_Group.Text = 'PRA-Wave1'; Click 'BtnCheckGroup'; Pump
        $lookupLimit = [datetime]::UtcNow.AddSeconds(60); while ($g.GroupLookup -and [datetime]::UtcNow -lt $lookupLimit) { Pump 100 }; Pump
        Shot 'task-convert-group'
        $c.Mode_Continue.IsChecked = $true; Pump
        Shot 'task-continue' 975
        Choose 'Recover'; $c.Mode_New.IsChecked = $true; Pump; $c.BatchList.SelectedIndex = 0; Pump
        Shot 'task-rollback' 975
        $c.S_Wave_Wave.IsChecked = $true; Pump
        foreach ($row in @($g.WaveRows | Select-Object -First 5)) { $row.Checked = $true }; & (Get-Module PRA.Gui) { Invoke-PcmGuiChange }; Pump
        Shot 'task-rollback-wave' 975
        Choose 'Check'; $c.Mode_Continue.IsChecked = $true; Pump; $c.BatchList.SelectedIndex = 0; Pump
        Shot 'task-check' 975
        Choose 'Convert'; $c.Mode_New.IsChecked = $true; $c.S_Source_Group.IsChecked = $true; $c.F_Group.Text = 'PRA-Wave1'; Pump
        Click 'BtnNext'
        Shot 'verify-before'
        Preview
        Shot 'verify-preview'
        $c.ConfirmBox.Text = 'convert'; Pump
        Shot 'verify-confirmed'
        Scenario 11; Hold 'users'
        Click 'BtnNext'
        $deadline = [datetime]::UtcNow.AddSeconds(60)
        while ([datetime]::UtcNow -lt $deadline) { Pump 200; $o = $g.Model.Objects['s6']; if ($o -and $o.Count -ge 11) { break } }
        $c.DetObjList.SelectedIndex = 1; Pump 300
        Shot 'run'
        Release 'users'; WaitRun
        Shot 'run-done'
        Click 'BtnNext'; Pump
        Shot 'next'
    }
    'SHOTS-OK'
} finally {
    if ($g -and $g.Form) { try { $g.Form.Close() } catch { Write-Verbose $_.Exception.Message } }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
