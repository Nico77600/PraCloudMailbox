#Requires -Version 7.4
<#
.SYNOPSIS
    Renders the screenshots of the guides (docs\images\pra-*.png) from the real tool, against the in-memory tenant.

.DESCRIPTION
    No tenant and no lab are needed, and the images always match the current code: the console and the reports
    are produced by Invoke-PraCloudMailbox.ps1 and its modules, with tests\FakeTenant.ps1 standing in for
    Microsoft Graph, Exchange Online, the eDiscovery case hold and Entra Connect (only the connections, the
    roles of the token, Entra Connect and the waits are replaced, as in the end-to-end tests). Names and paths
    are anonymised (contoso.com, C:\PRA\PraCloudMailbox).

        pra-console-check.png     console of a Check (read-only readiness of the last snapshot)
        pra-console-convert.png   console of a Convert -Mode Apply (users, then shared mailboxes)
        pra-console-recover.png   console of the Recover -Mode Apply of that batch
        pra-report-check.png      HTML report of the Check
        pra-report-convert.png    HTML report of the Convert
        pra-gui-overview.png      the window on the state left by the Convert: next step, configuration, snapshot, batches
        pra-gui-convert.png       the Convert page with the activity of that Convert (its events, read as the window does)

    The pages are captured with Microsoft Edge (headless), the window with RenderTargetBitmap (WPF). Run tools\Build-Documentation.ps1 afterwards: the
    guides embed the images.

.PARAMETER OutputFolder
    Default: docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (snapshot, journal, logs, reports, HTML pages) and shows its path.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0  (from Web Services Client for Exchange 1.0.0)
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'The in-memory tenant of the tests lives in global variables.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Replacements keep the signature of the real functions; Minutes is read in the run block.')]
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'docs\images' }
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('pra-doc-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$shown = 'C:\PRA\PraCloudMailbox'

#region Edge helpers (same as Web Services Client for Exchange) ---------------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,$Height", '--force-device-scale-factor=1', '--virtual-time-budget=3000', "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # The page writes its height in body[data-h]; Edge returns the DOM with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--virtual-time-budget=3000', '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function Save-Page([string]$Html, [string]$Name, [int]$Width, [int]$Height = 0) {
    if ($Height -le 0) { $Height = Get-PageHeight $Html $Width }
    $png = Join-Path $OutputFolder "$Name.png"
    Save-Screenshot $Html $png $Width $Height
    Write-Host ("  {0,-26} {1} x {2}" -f "$Name.png", $Width, $Height)
}
#endregion

#region The tool against the in-memory tenant -------------------------------------------------------------------------
# Console theme is chosen when PRA2.Common loads: emoji as in Windows Terminal, colours as ANSI sequences.
$env:PRA_ICONS = 'Emoji'
$env:PRA_ANSI = '1'
foreach ($name in 'PRA2.Common', 'PRA2.Store', 'PRA2.Cloud') {
    Remove-Module $name -ErrorAction SilentlyContinue
    Import-Module (Join-Path $root "module\$name.psm1") -Force
}
. (Join-Path $root 'tests\FakeTenant.ps1')
# The functions of the entry script, without its main block (which re-imports the modules and exits).
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })) {
    . ([scriptblock]::Create($function.Extent.Text))
}
$toolVersion = [regex]::Match($ast.Extent.Text, "(?m)^\`$toolVersion = '([^']+)'").Groups[1].Value

# What reaches Microsoft 365 or the Entra Connect server is replaced, as in tests\PraCloudMailbox.EndToEnd.Tests.ps1.
function Connect-Pra2Cloud { 'Microsoft Graph and Exchange Online (certificate of the app registration)' }
function Disconnect-Pra2Cloud { }
function Connect-Pra2Compliance { }
function Get-Pra2GraphRole { @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
function Invoke-Pra2EntraConnect {
    param([hashtable]$Context, [string]$Operation, [object]$Caller)
    $null = Invoke-PraFakeEntraConnect -Operation $Operation
    @{ Pause = 'SyncCycleEnabled=False InProgress=False'; Resume = 'SyncCycleEnabled=True InProgress=False'; Delta = 'SyncCycleEnabled=True InProgress=False' }[$Operation]
}
function Wait-Pra2Condition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
& (Get-Module PRA2.Cloud) {
    function script:Wait-Pra2Condition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
}

$configPath = Join-Path $work 'PraCloudMailbox.config.psd1'
Set-Content -LiteralPath $configPath -Encoding UTF8 -Value @"
@{
    Environment = 'PROD'
    Scope = @{ Mode = 'OU'; SearchBase = 'OU=Paris,DC=contoso,DC=com' }
    Store = @{ Path = '$work\data\PraCloudMailbox.db'; JournalPath = '$work\data\PraCloudMailbox-journal.db'; BackupFolder = '$work\data\backup' }
    Cloud = @{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; AppId = '44444444-4444-4444-4444-444444444444'; CertificateThumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'; DefaultUsageLocation = 'FR' }
    Polling = @{ IntervalSeconds = 20; MailboxTimeoutMinutes = 30; SyncTimeoutMinutes = 30; HoldTimeoutMinutes = 30 }
    Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }; Shared = @{ SkuPartNumber = 'E5' } }
    Retention = @{ HoldPolicy = 'PRA-Recover-Hold' }
    EntraConnect = @{ Mode = 'Remoting'; Server = 'aadc01.contoso.com' }
    Logging = @{ Folder = '$work\logs' }
    Report = @{ Folder = '$work\reports' }
}
"@

New-PraFakeTenant
$global:PraFake.CasePolicy.Name = 'PRA-Recover-Hold'
$global:PraFake.Group.Name = 'LIC-Exchange-Online'
$people = [ordered]@{}
foreach ($u in @(
        @{ Upn = 'jdupont@contoso.com'; Teams = $true }, @{ Upn = 'mmartin@contoso.com'; Teams = $true }, @{ Upn = 'lbernard@contoso.com'; Teams = $false }
        @{ Upn = 'compta@contoso.com'; Kind = 'Shared' }, @{ Upn = 'accueil@contoso.com'; Kind = 'Shared' })) {
    $kind = if ($u.ContainsKey('Kind')) { $u.Kind } else { 'User' }
    $people[$u.Upn] = Add-PraFakeUser -Upn $u.Upn -Kind $kind -Component $(if ($u.ContainsKey('Teams') -and $u.Teams) { [guid]::NewGuid().ToString() } else { '' }) -UsageLocation $(if ($kind -eq 'Shared') { '' } else { 'FR' })
}
$people['lbernard@contoso.com'].UsageLocation = ''

function Write-DocSnapshot {
    <# A complete snapshot of the five objects, as Collect writes it on the Exchange server. #>
    $byUpn = @{}; foreach ($ad in $global:PraFake.Ad) { $byUpn[$ad.Upn] = $ad }
    $cn = Open-Pra2Store -Path "$work\data\PraCloudMailbox.db" -Root $root
    try {
        $id = New-Pra2Snapshot $cn @{ RunId = [guid]::NewGuid().ToString('N'); Environment = 'PROD'; ToolVersion = $toolVersion; ExchangeServer = 'EX01.contoso.com'; ScopeJson = '{"Mode":"OU"}' }
        $rows = @(foreach ($upn in $people.Keys) {
                $ad = $byUpn[$upn]
                [ordered]@{ object_guid = $ad.ObjectGuid; kind = $ad.Kind; user_principal_name = $upn; primary_smtp_address = $upn; sam_account_name = ($upn -split '@')[0]
                    exchange_guid = $ad.OnPremGuid; immutable_id = $ad.ImmutableId; recipient_type_details = "$($ad.Kind)Mailbox" }
            })
        $null = Add-Pra2Row $cn mailbox $id $rows
        $grant = { param($mailbox, $upn, $right, $automap) [ordered]@{ mailbox_guid = $byUpn[$mailbox].ObjectGuid; access_right = $right; trustee = "CONTOSO\$(($upn -split '@')[0])"; trustee_kind = 'User'
                trustee_upn = $upn; trustee_guid = $byUpn[$upn].ObjectGuid; trustee_name = ($upn -split '@')[0]; resolved = 1; auto_mapping = $automap } }
        $permissions = @(
            (& $grant 'compta@contoso.com' 'jdupont@contoso.com' 'FullAccess' 1)
            [ordered]@{ mailbox_guid = $byUpn['compta@contoso.com'].ObjectGuid; access_right = 'FullAccess'; trustee = 'CONTOSO\GRP-Compta'; trustee_kind = 'Group'
                trustee_upn = $null; trustee_guid = 'grp-compta'; trustee_name = 'GRP-Compta'; resolved = 1; auto_mapping = $null }
            (& $grant 'compta@contoso.com' 'mmartin@contoso.com' 'SendAs' $null)
            (& $grant 'compta@contoso.com' 'jdupont@contoso.com' 'SendOnBehalf' $null)
            (& $grant 'accueil@contoso.com' 'lbernard@contoso.com' 'FullAccess' 1)
            (& $grant 'accueil@contoso.com' 'lbernard@contoso.com' 'SendAs' $null))
        $null = Add-Pra2Row $cn permission $id $permissions
        $memberUpns = 'mmartin@contoso.com', 'lbernard@contoso.com'
        $members = $memberUpns | ForEach-Object { [ordered]@{ group_guid = 'grp-compta'; member_kind = 'User'; member_upn = $_; member_name = ($_ -split '@')[0]; depth = 1 } }
        $null = Add-Pra2Row $cn group_member $id @($members)
        Set-Pra2SnapshotStatus $cn $id Complete @{ mailbox = $rows.Count }
    } finally { Close-Pra2Store $cn }
}
New-Item -ItemType Directory -Path "$work\data" -Force | Out-Null
Write-DocSnapshot

function Invoke-DocRun {
    <# One run as Invoke-PraCloudMailbox.ps1 does it (main block): banner, steps, summary card, reports. #>
    param([ValidateSet('Check', 'Convert', 'Recover')][string]$RunAction, [string]$RunMode = 'Preview', [string]$RunBatch = '', [double]$Minutes)
    $values = @{ Action = $RunAction; Mode = $(if ($RunAction -eq 'Check') { 'Preview' } else { $RunMode }); Batch = $RunBatch; Identity = ''; IdentityPath = ''; Scope = 'All'; Snapshot = [long]0; Force = $true; PassThru = $false; caller = $null }
    foreach ($name in $values.Keys) { Set-Variable -Name $name -Value $values[$name] -Scope Script }
    $script:effectiveMode = $values.Mode
    $script:context = @{
        Root = $root; Version = $toolVersion; RunId = ((Get-Date -Format 'yyyyMMdd_HHmmss') + '-' + [guid]::NewGuid().ToString('N'))
        StartTime = Get-Date; Action = $RunAction; Mode = $values.Mode; Phase = ''
        CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; StepIndex = 0; StepTotal = 0; Warnings = 0
        Issues = [System.Collections.Generic.List[object]]::new(); Rows = [System.Collections.Generic.List[object]]::new()
        BackupFiles = [System.Collections.Generic.List[string]]::new(); StateFiles = [System.Collections.Generic.List[string]]::new()
        Excluded = [System.Collections.Generic.List[string]]::new()
        LogFolder = ''; ReportFolder = ''; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $false
        Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''; Config = @{}; VerboseEnabled = $false
        BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null
    }
    $records = & {
        try {
            $context.Config = Import-PraConfiguration -Path $configPath -Root $root
            $script:config = $context.Config
            $context.LogFolder = $context.Config.Logging.Folder; $context.ReportFolder = $context.Config.Report.Folder
            Initialize-PraAudit $context
            $dot = [char]0x00B7
            $modeText = if ($RunAction -eq 'Check') { 'Read-only' } elseif ($effectiveMode -eq 'Apply') { 'Apply (changes are made)' } else { 'Preview (nothing is changed)' }
            $banner = [ordered]@{ 'Action' = @('Mode', $RunAction); 'Mode' = @($(if ($effectiveMode -eq 'Apply') { 'Write' } else { 'Plan' }), $modeText) }
            $banner['Scope'] = @('Target', $(if ($RunBatch) { "batch $RunBatch" } else { 'snapshot last complete' }))
            $banner['Tenant'] = @('Cloud', $context.Config.Cloud.Organization)
            $banner['Database'] = @('Folder', $context.Config.Store.Path)
            $banner['Config'] = @('Config', ('{0} {1} Environment {2}' -f (Split-Path $context.Config._Path -Leaf), $dot, $context.Config.Environment))
            $banner['Log'] = @('Log', $context.LogFile)
            Write-PraBanner -Context $context -Title 'PRA Cloud Mailbox' -Subtitle "Exchange DR $dot scenario 2 $dot on-premises lost $([char]0x2192) Exchange Online" -Details $banner
            $context.StepTotal = @{ Check = 4; Convert = 7; Recover = 7 }[$RunAction]
            switch ($RunAction) { 'Check' { Invoke-PraCheck } 'Convert' { Invoke-PraConvert } 'Recover' { Invoke-PraRecover } }
        } catch {
            $context.Issues.Add([pscustomobject]@{ Message = $_.Exception.Message; Source = 'Execution' }); $context.ExitCode = 1
        } finally {
            if ($context.Journal) { Close-Pra2Store $context.Journal; $context.Journal = $null }
            # The simulated tenant answers at once: the duration shown is the one measured in the lab.
            if ($Minutes) { $context.StartTime = (Get-Date).AddMinutes(-$Minutes) }
            $script:DocResult = Complete-PraRun -Context $context
        }
    } 6>&1
    Write-Host ("  {0,-8} {1}, exit code {2}, {3} object(s)" -f $RunAction, $script:DocResult.Status, $script:DocResult.ExitCode, $script:DocResult.TotalCount)
    if ($script:DocResult.ExitCode -eq 1) { throw "$RunAction failed: $(@($script:DocResult.Issues | ForEach-Object { $_.Message }) -join '; ')" }
    [pscustomobject]@{ Records = $records; Result = $script:DocResult; Html = $script:DocResult.HtmlReport }
}

function Get-ShownText([string]$Text) {
    # Work folder -> the folder of a real installation; this computer and this account -> neutral names.
    $account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Text = $Text.Replace($work, $shown).Replace($account, 'CONTOSO\pra-admin')
    # Also as escaped in the JSON data of the HTML report.
    $Text = $Text.Replace($work.Replace('\', '\\'), $shown.Replace('\', '\\')).Replace($account.Replace('\', '\\'), 'CONTOSO\\pra-admin')
    return $Text.Replace($env:COMPUTERNAME, 'PRA-ADMIN01')
}
#region Window images (WPF) ------------------------------------------------------------------------------------------
function Invoke-DocPump {
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background,
        [Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $frame)
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Save-WindowImage([Windows.Window]$Form, [string]$Name) {
    Invoke-DocPump; $Form.UpdateLayout(); Invoke-DocPump
    $content = $Form.Content
    $bitmap = [Windows.Media.Imaging.RenderTargetBitmap]::new([int][Math]::Ceiling($content.ActualWidth), [int][Math]::Ceiling($content.ActualHeight), 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($content)
    $encoder = [Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream = [IO.File]::Create((Join-Path $OutputFolder "$Name.png"))
    try { $encoder.Save($stream) } finally { $stream.Dispose() }
    Write-Host ("  {0,-26} {1} x {2}" -f "$Name.png", $bitmap.PixelWidth, $bitmap.PixelHeight)
}

function Save-DocWindow {
    <# The window on the state the Convert left (its objects wait for Recover), with the activity of that Convert. #>
    param([string]$EventsPath)
    $gui = Import-Module (Join-Path $root 'module\PRA2.Gui.psm1') -Force -PassThru
    # No certificate store or module folder of this computer in the images.
    & $gui {
        function script:Get-PraGuiCertificate { [pscustomobject]@{ Text = 'LocalMachine\My, valid until 2028-06-30'; Tone = 'Primary' } }
        function script:Get-PraGuiModuleVersion { param([string]$Name) @{ 'Microsoft.Graph.Authentication' = '2.40.0'; 'ExchangeOnlineManagement' = '3.10.1' }[$Name] }
    }
    $window = New-PraGuiWindow -Root $root -ConfigPath $configPath -Version $toolVersion -Theme Light
    $form = $window.Form
    $form.WindowStartupLocation = 'Manual'; $form.Left = -6000; $form.Top = -6000; $form.Width = 1460; $form.Height = 920
    $form.ShowActivated = $false; $form.ShowInTaskbar = $false
    $form.Show()
    try {
        # The events of the Convert, read and shown as the window does while a run goes.
        & $gui {
            param($events)
            $g = $script:Gui
            $command = 'pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Force'
            Clear-PraGuiActivity
            $g.Controls.RunAction.Text = 'Convert · Apply'
            $g.Controls.RunMeta.Text = $command
            Add-PraGuiLine -Status Info -Text ('Started: ' + $command)
            $g.Run = @{ Request = @{ Action = 'Convert'; Mode = 'Apply' }; Files = @{ Events = $events }; Counts = @{ Ok = 0; Warn = 0; Fail = 0 }; SubLines = 0
                Asked = [Collections.Generic.HashSet[int]]::new(); StopAsked = $false; LogFile = ''; Result = $null; Summary = $null }
            foreach ($record in (Read-PraGuiEvents -Path $events).Events) { Invoke-PraGuiEvent -Record $record }
            Show-PraGuiRunResult -Run $g.Run -ExitCode 0
            $g.Run = $null
            Set-PraGuiBusy -Busy $false
        } $EventsPath
        $hide = { param([string]$Text) (Get-ShownText $Text).Replace($root, $shown) }
        for ($i = 0; $i -lt $window.Lines.Count; $i++) {
            $line = $window.Lines[$i]
            $text = & $hide $line.Text
            if ($text -ne $line.Text) { $copy = $line.PSObject.Copy(); $copy.Text = $text; $window.Lines[$i] = $copy }
        }
        foreach ($name in 'OvConfig', 'OvSnapshot', 'OvBatches', 'OvComputer', 'CollectFacts', 'ConvertFacts', 'RecoverFacts') {
            $facts = @($window.Controls[$name].ItemsSource)
            $window.Controls[$name].ItemsSource = @(foreach ($fact in $facts) { $copy = $fact.PSObject.Copy(); $copy.Value = & $hide $fact.Value; $copy })
        }
        foreach ($name in 'Footer', 'RunMeta', 'ResultText', 'NextStepText', 'CheckInfo', 'CollectTask') { $window.Controls[$name].Text = & $hide $window.Controls[$name].Text }
        & $gui { Show-PraGuiLastLine }
        & $gui { Select-PraGuiPage -Name Overview }
        Save-WindowImage $form 'pra-gui-overview'
        & $gui { Select-PraGuiPage -Name Convert }
        Save-WindowImage $form 'pra-gui-convert'
    } finally { $form.Close() }
    Remove-Module PRA2.Gui -Force -ErrorAction SilentlyContinue
}
#endregion
Write-Host 'Runs (in-memory tenant)...'
# Every poll of the tool advances the simulated tenant by one tick, as in the end-to-end tests.
function global:Start-Sleep { Step-PraFake }
try {
    $check = Invoke-DocRun -RunAction Check -Minutes 0.7
    # The events of the Convert, as the window reads them (PRA_EVENT_FILE).
    $events = Join-Path $work 'convert.events.jsonl'
    & (Get-Module PRA2.Common) { $script:EventFile = $args[0] } $events
    try { $convert = Invoke-DocRun -RunAction Convert -RunMode Apply -Minutes 9.6 }
    finally { & (Get-Module PRA2.Common) { $script:EventFile = '' } }
    Write-Host 'Window (before the Recover):'
    Save-DocWindow -EventsPath $events
    $recover = Invoke-DocRun -RunAction Recover -RunMode Apply -RunBatch $convert.Result.BatchId -Minutes 14.4
} finally {
    Remove-Item function:global:Start-Sleep -ErrorAction SilentlyContinue
}
#endregion

#region Console images ------------------------------------------------------------------------------------------------
# Campbell palette of Windows Terminal, for the 16 colours the console theme uses (PRA_ANSI=1).
$palette = @{ 30 = '#0c0c0c'; 31 = '#c50f1f'; 32 = '#13a10e'; 33 = '#c19c00'; 34 = '#0037da'; 35 = '#881798'; 36 = '#3a96dd'; 37 = '#cccccc'
    90 = '#767676'; 91 = '#e74856'; 92 = '#16c60c'; 93 = '#f9f1a5'; 94 = '#3b78ff'; 95 = '#b4009e'; 96 = '#61d6d6'; 97 = '#f2f2f2' }
function ConvertFrom-Ansi([string]$Line) {
    $out = [Text.StringBuilder]::new(); $fg = $null
    foreach ($part in [regex]::Split($Line, '(\x1b\[[0-9;]*m)')) {
        if ($part -match '^\x1b\[([0-9;]*)m$') {
            $code = if ($Matches[1]) { [int]$Matches[1] } else { 0 }
            $fg = if ($palette.ContainsKey($code)) { $palette[$code] } else { $null }
            continue
        }
        if ($part -eq '') { continue }
        $text = [Net.WebUtility]::HtmlEncode($part)
        [void]$out.Append($(if ($fg) { "<span style=""color:$fg"">$text</span>" } else { $text }))
    }
    return $out.ToString()
}


function Save-Console([object[]]$Records, [string]$Command, [string]$Name) {
    $lines = foreach ($r in $Records) {
        $data = if ($r -is [Management.Automation.InformationRecord]) { $r.MessageData } else { $r }
        $text = if ($data -is [Management.Automation.HostInformationMessage]) { [string]$data.Message } else { [string]$data }
        Get-ShownText $text
    }
    $body = ($lines | ForEach-Object { ConvertFrom-Ansi $_ }) -join "`n"
    $console = @"
<!doctype html><html><head><meta charset="utf-8"><style>
body { margin:0; background:#ffffff; font-family:"Segoe UI", sans-serif; }
.win { width:1180px; margin:0; border-radius:10px; overflow:hidden; background:#0c0c0c; border:1px solid #2b2b2b; }
.bar { display:flex; align-items:center; gap:10px; height:38px; padding:0 14px; background:#202020; color:#d0d0d0; font-size:12.5px; }
.tab { padding:6px 14px; background:#0c0c0c; border-radius:8px 8px 0 0; margin-top:8px; }
pre { margin:0; padding:14px 18px 18px; color:#cccccc; font:13.5px/1.42 "Cascadia Mono", Consolas, monospace; white-space:pre-wrap; word-break:break-all; }
</style></head><body><div class="win"><div class="bar"><span class="tab">PowerShell 7.5</span></div>
<pre><span style="color:#cccccc">PS $shown&gt; $([Net.WebUtility]::HtmlEncode($Command))</span>
$body</pre></div>
<script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.win').getBoundingClientRect().height) + 2);</script></body></html>
"@
    $page = Join-Path $work "$Name.html"
    [IO.File]::WriteAllText($page, $console, [Text.UTF8Encoding]::new($false))
    Save-Page $page $Name 1182
}

Write-Host 'Images:'
Save-Console $check.Records '.\Invoke-PraCloudMailbox.ps1 -Action Check' 'pra-console-check'
Save-Console $convert.Records '.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply' 'pra-console-convert'
Save-Console $recover.Records ".\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch $($convert.Result.BatchId)" 'pra-console-recover'
#endregion

#region Report images -------------------------------------------------------------------------------------------------
function Save-Report([string]$Path, [string]$Name) {
    $html = Get-ShownText ([IO.File]::ReadAllText($Path))
    $measure = "<script>window.addEventListener('load', () => setTimeout(() => document.body.setAttribute('data-h', Math.ceil(document.body.getBoundingClientRect().height) + 24), 50));</script></body>"
    $page = Join-Path $work "$Name.html"
    [IO.File]::WriteAllText($page, $html.Replace('</body>', $measure), [Text.UTF8Encoding]::new($false))
    Save-Page $page $Name 1280
}
Save-Report $check.Html 'pra-report-check'
Save-Report $convert.Html 'pra-report-convert'
#endregion

foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
Remove-Variable -Name PraFake, PraFakeCommands -Scope Global -ErrorAction SilentlyContinue
foreach ($name in 'PRA2.Cloud', 'PRA2.Store', 'PRA2.Common') { Remove-Module $name -Force -ErrorAction SilentlyContinue }
Remove-Item Env:\PRA_ICONS, Env:\PRA_ANSI -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
