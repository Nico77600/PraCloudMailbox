#Requires -Version 7.5
<#
.SYNOPSIS
    Renders the screenshots of the guides (docs\images\pra-*.png) from the real tool, against the in-memory tenant.

.DESCRIPTION
    No tenant and no test environment are needed, and the images always match the current code: the consoles and the reports are produced by Invoke-PraCloudMailbox.ps1 and its modules, with
    tests\FakeTenant.ps1 standing in for Microsoft Graph, Exchange Online, the eDiscovery case hold and Entra Connect (only the connections, the roles of the token, Entra Connect and the waits are
    replaced, as in the end-to-end tests). Names and paths are synthetic (contoso.com, C:\PRA\PraCloudMailbox). Every picture exists in a light and a dark version (-light.png, -dark.png):

        cli-check-<theme>.png          console of a Check (read-only readiness of the last snapshot)
        cli-convert-group-<theme>.png  console of the Preview of a Convert on a group (header line, skipped members)
        cli-convert-<theme>.png        console of a Convert -Mode Apply (users, then shared mailboxes)
        cli-follow-<theme>.png         console of Check -Batch -Expect Provisioned -Once (the follow-up of the batch)
        cli-recover-<theme>.png        console of the Recover -Mode Apply of that batch
        report-check-<theme>.png       HTML report of the Check
        report-convert-<theme>.png     HTML report of the Convert
        report-global-<theme>.png      the global report of the batch: Convert, follow-up and Recover, every object end to end
        console-<screen>-<theme>.png   the guided console, page by page (the real window against the synthetic engine of tools\DocumentationImages:
                                       tools\DocumentationImages\Show-ConsoleScreens.ps1, one process per scenario and theme; synthetic names only, a real
                                       machine, user or path name stops the run)

    The pages are captured with Microsoft Edge (headless), the window with RenderTargetBitmap (WPF). Run tools\Build-Documentation.ps1 afterwards: the
    guides embed the images.
.PARAMETER OutputFolder
    Default: package\docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (snapshot, journal, logs, reports, HTML pages) and shows its path.

.PARAMETER SkipConsole
    Does not photograph the guided console (the pictures of the command line and of the reports only).

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0  (from Web Services Client for Exchange 1.0.0)
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars', '', Justification = 'The in-memory tenant of the tests lives in global variables.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'Replacements keep the signature of the real functions; Minutes is read in the run block.')]
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork,
    [switch]$SkipConsole
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'package\docs\images' }
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
    # The page writes its height in body[data-h]; Edge returns the DOM with --dump-dom. A headless Edge that exits without an answer is started again (up to three times).
    $url = 'file:///' + ($Html -replace '\\', '/')
    for ($attempt = 1; $attempt -le 3; $attempt++) {
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
        if ($m.Success) { return [int]$m.Groups[1].Value }
    }
    throw "Height not measured: $Html"
}
function Save-Page([string]$Html, [string]$Name, [int]$Width, [int]$Height = 0) {
    if ($Height -le 0) { $Height = Get-PageHeight $Html $Width }
    $png = Join-Path $OutputFolder "$Name.png"
    Save-Screenshot $Html $png $Width $Height
    Write-Host ("  {0,-26} {1} x {2}" -f "$Name.png", $Width, $Height)
}
#endregion

#region The tool against the in-memory tenant -------------------------------------------------------------------------
# Console theme is chosen when PRA.Common loads: emoji as in Windows Terminal, colours as ANSI sequences.
$env:PRA_ICONS = 'Emoji'
$env:PRA_ANSI = '1'
. (Join-Path $PSScriptRoot 'DocumentationImages\New-ToolCopy.ps1')
$package = New-DocumentationToolCopy -Repository $root -Destination (Join-Path $work 'package') -RealEngine
foreach ($name in 'PRA.Common', 'PRA.Store', 'PRA.Cloud') {
    Remove-Module $name -ErrorAction SilentlyContinue
    Import-Module (Join-Path $package "module\$name.psm1") -Force
}
. (Join-Path $root 'tests\FakeTenant.ps1')
# The functions of the entry script, without its main block (which re-imports the modules and exits).
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $package 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })) {
    . ([scriptblock]::Create($function.Extent.Text))
}
$toolVersion = Get-PcmToolVersion

# What reaches Microsoft 365 or the Entra Connect server is replaced, as in tests\PraCloudMailbox.EndToEnd.Tests.ps1.
function Connect-PcmCloud { 'Microsoft Graph and Exchange Online (certificate of the app registration)' }
function Disconnect-PcmCloud { }
function Connect-PcmCompliance { }
function Get-PcmGraphRole { @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
function Invoke-PcmEntraConnect {
    param([hashtable]$Context, [string]$Operation, [object]$Caller)
    $null = Invoke-PraFakeEntraConnect -Operation $Operation
    @{ Pause = 'SyncCycleEnabled=False InProgress=False'; Resume = 'SyncCycleEnabled=True InProgress=False'; Delta = 'SyncCycleEnabled=True InProgress=False' }[$Operation]
}
function Wait-PcmCondition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
& (Get-Module PRA.Cloud) {
    function script:Wait-PcmCondition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
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
        @{ Upn = 'adele.vance@contoso.com'; Teams = $true }, @{ Upn = 'megan.bowen@contoso.com'; Teams = $true }, @{ Upn = 'alex.wilber@contoso.com'; Teams = $false }
        @{ Upn = 'accounting@contoso.com'; Kind = 'Shared' }, @{ Upn = 'helpdesk@contoso.com'; Kind = 'Shared' })) {
    $kind = if ($u.ContainsKey('Kind')) { $u.Kind } else { 'User' }
    $people[$u.Upn] = Add-PraFakeUser -Upn $u.Upn -Kind $kind -Component $(if ($u.ContainsKey('Teams') -and $u.Teams) { [guid]::NewGuid().ToString() } else { '' }) -UsageLocation $(if ($kind -eq 'Shared') { '' } else { 'FR' })
}
$people['alex.wilber@contoso.com'].UsageLocation = ''

function Write-DocSnapshot {
    <# A complete snapshot of the five objects, as Collect writes it on the Exchange server. #>
    $byUpn = @{}; foreach ($ad in $global:PraFake.Ad) { $byUpn[$ad.Upn] = $ad }
    $cn = Open-PcmStore -Path "$work\data\PraCloudMailbox.db" -Root $package
    try {
        # The pictures show snapshot 42: the counter of the database starts there.
        $null = Invoke-PcmSql $cn "INSERT INTO sqlite_sequence (name, seq) VALUES ('snapshot', 41)"
        $id = New-PcmSnapshot $cn @{ RunId = [guid]::NewGuid().ToString('N'); Environment = 'PROD'; ToolVersion = $toolVersion; ExchangeServer = 'EX01.contoso.com'; ScopeJson = '{"Mode":"OU"}' }
        $rows = @(foreach ($upn in $people.Keys) {
                $ad = $byUpn[$upn]
                [ordered]@{ object_guid = $ad.ObjectGuid; kind = $ad.Kind; user_principal_name = $upn; primary_smtp_address = $upn; sam_account_name = ($upn -split '@')[0]
                    exchange_guid = $ad.OnPremGuid; immutable_id = $ad.ImmutableId; recipient_type_details = "$($ad.Kind)Mailbox" }
            })
        $null = Add-PcmRow $cn mailbox $id $rows
        $grant = { param($mailbox, $upn, $right, $automap) [ordered]@{ mailbox_guid = $byUpn[$mailbox].ObjectGuid; access_right = $right; trustee = "CONTOSO\$(($upn -split '@')[0])"; trustee_kind = 'User'
                trustee_upn = $upn; trustee_guid = $byUpn[$upn].ObjectGuid; trustee_name = ($upn -split '@')[0]; resolved = 1; auto_mapping = $automap } }
        $permissions = @(
            (& $grant 'accounting@contoso.com' 'adele.vance@contoso.com' 'FullAccess' 1)
            [ordered]@{ mailbox_guid = $byUpn['accounting@contoso.com'].ObjectGuid; access_right = 'FullAccess'; trustee = 'CONTOSO\GRP-Accounting'; trustee_kind = 'Group'
                trustee_upn = $null; trustee_guid = 'grp-accounting'; trustee_name = 'GRP-Accounting'; resolved = 1; auto_mapping = $null }
            (& $grant 'accounting@contoso.com' 'megan.bowen@contoso.com' 'SendAs' $null)
            (& $grant 'accounting@contoso.com' 'adele.vance@contoso.com' 'SendOnBehalf' $null)
            (& $grant 'helpdesk@contoso.com' 'alex.wilber@contoso.com' 'FullAccess' 1)
            (& $grant 'helpdesk@contoso.com' 'alex.wilber@contoso.com' 'SendAs' $null))
        $null = Add-PcmRow $cn permission $id $permissions
        $memberUpns = 'megan.bowen@contoso.com', 'alex.wilber@contoso.com'
        $members = $memberUpns | ForEach-Object { [ordered]@{ group_guid = 'grp-accounting'; member_kind = 'User'; member_upn = $_; member_name = ($_ -split '@')[0]; depth = 1 } }
        $null = Add-PcmRow $cn group_member $id @($members)
        $wave = 'b0000000-0000-0000-0000-000000000001'
        $null = Add-PcmRow $cn target_group $id @([ordered]@{ group_guid = $wave; name = 'PRA-Wave1'; sam_account_name = 'PRA-Wave1'; distinguished_name = 'CN=PRA-Wave1,OU=Groups,DC=contoso,DC=com'; primary_smtp_address = ''; display_name = 'PRA-Wave1'; member_count = 5 })
        $waveMembers = @(foreach ($upn in 'adele.vance@contoso.com', 'megan.bowen@contoso.com', 'alex.wilber@contoso.com') {
                $ad = $byUpn[$upn]
                [ordered]@{ group_guid = $wave; member_guid = $ad.ObjectGuid; member_class = 'User'; member_name = ($upn -split '@')[0]; member_upn = $upn; member_smtp = $upn; member_dn = ''; via_group_guid = ''; depth = 1 }
            })
        $waveMembers += [ordered]@{ group_guid = $wave; member_guid = 'c0000000-0000-0000-0000-000000000001'; member_class = 'Contact'; member_name = 'Print Room'; member_upn = ''; member_smtp = 'print.room@contoso.com'; member_dn = ''; via_group_guid = ''; depth = 1 }
        $waveMembers += [ordered]@{ group_guid = $wave; member_guid = 'c0000000-0000-0000-0000-000000000002'; member_class = 'Computer'; member_name = 'KIOSK-01'; member_upn = ''; member_smtp = ''; member_dn = ''; via_group_guid = ''; depth = 2 }
        $null = Add-PcmRow $cn target_member $id $waveMembers
        Set-PcmSnapshotStatus $cn $id Complete @{ mailbox = $rows.Count }
    } finally { Close-PcmStore $cn }
    $null = Write-PcmHashFile -Path "$work\data\PraCloudMailbox.db"
}
New-Item -ItemType Directory -Path "$work\data" -Force | Out-Null
Write-DocSnapshot

function Invoke-DocRun {
    <# One run as Invoke-PraCloudMailbox.ps1 does it (main block): banner, steps, summary card, reports. #>
    param([ValidateSet('Check', 'Convert', 'Recover')][string]$RunAction, [string]$RunMode = 'Preview', [string]$RunBatch = '', [double]$Minutes, [string]$RunGroup = '', [string]$RunExpect = '', [switch]$RunOnce)
    $values = @{ Action = $RunAction; Mode = $(if ($RunAction -eq 'Check') { 'Preview' } else { $RunMode }); Batch = $RunBatch; Identity = ''; IdentityPath = ''; Scope = 'All'; Snapshot = [long]0; Force = $true; PassThru = $false; ResolveOnly = $false; caller = $null
        CsvPath = ''; SearchBase = ''; Group = $RunGroup; MaxObjects = 0; Expect = $RunExpect; Once = [bool]$RunOnce; BaseConfigPath = ''; AcceptUnverifiedSnapshot = $false }
    foreach ($name in $values.Keys) { Set-Variable -Name $name -Value $values[$name] -Scope Script }
    $script:effectiveMode = $values.Mode
    $script:BoundOptions = [string[]]@($values.Keys | Where-Object { $values[$_] -and $_ -in 'Batch', 'Group', 'Expect', 'Once' })
    $script:context = @{
        Root = $package; Version = $toolVersion; RunId = ((Get-Date -Format 'yyyyMMdd_HHmmss') + '-' + [guid]::NewGuid().ToString('N'))
        StartTime = Get-Date; Action = $RunAction; Mode = $values.Mode; Phase = ''
        CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; StepIndex = 0; StepTotal = 0; Warnings = 0
        Issues = [System.Collections.Generic.List[object]]::new(); Rows = [System.Collections.Generic.List[object]]::new()
        BackupFiles = [System.Collections.Generic.List[string]]::new(); StateFiles = [System.Collections.Generic.List[string]]::new()
        Excluded = [System.Collections.Generic.List[string]]::new()
        LogFolder = ''; ReportFolder = ''; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $false
        Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''; Config = @{}; VerboseEnabled = $false
        BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null
        Approval = { param($Target, $Operation) $true }; ConfirmPreference = 'None'; FallbackNext = @(); NextOnFailure = $false; NoChangeMade = $false; Cancelled = $false
        SkippedMembers = [System.Collections.Generic.List[string]]::new(); ObjectsText = ''; ConfigOverrides = @(); GroupTarget = $null; LeftForLater = 0
        CloudDescription = ''; GlobalReportPath = ''; SnapshotCheck = ''; AcceptUnverifiedSnapshot = $false
    }
    Set-PcmWriteApproval -Approval $context.Approval
    $script:ConfigPathBound = $false; $script:ConfirmBound = $true
    $records = & {
        try {
            $context.Config = Import-PcmConfiguration -Path $configPath -Root $package
            $script:config = $context.Config
            $context.LogFolder = $context.Config.Logging.Folder; $context.ReportFolder = $context.Config.Report.Folder
            Initialize-PcmAudit $context
            $dot = [char]0x00B7
            $modeText = if ($RunAction -eq 'Check') { 'Read-only' } elseif ($effectiveMode -eq 'Apply') { 'Apply (changes are made)' } else { 'Preview (nothing is changed)' }
            $banner = [ordered]@{ 'Action' = @('Mode', $RunAction); 'Mode' = @($(if ($effectiveMode -eq 'Apply') { 'Write' } else { 'Plan' }), $modeText) }
            $banner['Scope'] = @('Target', $(if ($RunBatch) { "batch $RunBatch" } elseif ($RunGroup) { "group $RunGroup" } else { 'snapshot last complete' }))
            $banner['Tenant'] = @('Cloud', $context.Config.Cloud.Organization)
            $banner['Database'] = @('Folder', $context.Config.Store.Path)
            $banner['Config'] = @('Config', ('{0} {1} Environment {2}' -f (Split-Path $context.Config._Path -Leaf), $dot, $context.Config.Environment))
            $banner['Log'] = @('Log', $context.LogFile)
            Write-PcmBanner -Context $context -Title 'PRA Cloud Mailbox' -Subtitle "Exchange DR $dot scenario 2 $dot on-premises lost $([char]0x2192) Exchange Online" -Details $banner
            $context.StepTotal = @{ Check = 4; Convert = 7; Recover = 7 }[$RunAction]
            switch ($RunAction) { 'Check' { if ($RunBatch) { Invoke-PcmCheckBatch } else { Invoke-PcmCheck } } 'Convert' { Invoke-PcmConvert } 'Recover' { Invoke-PcmRecover } }
        } catch {
            $context.Issues.Add([pscustomobject]@{ Message = $_.Exception.Message; Source = 'Execution' }); $context.ExitCode = 1
        } finally {
            if ($context.Journal) {
                try { Write-PcmBatchRecord } catch { Write-Host ("batch record: $($_.Exception.Message)") }
                Close-PcmStore $context.Journal; $context.Journal = $null
            }
            # The simulated tenant answers at once: the duration shown is a typical one.
            if ($Minutes) { $context.StartTime = (Get-Date).AddMinutes(-$Minutes) }
            $script:DocResult = Complete-PcmRun -Context $context
            $script:DocGlobal = [string]$context.GlobalReportPath
        }
    } 6>&1
    Write-Host ("  {0,-8} {1}, exit code {2}, {3} object(s)" -f $RunAction, $script:DocResult.Status, $script:DocResult.ExitCode, $script:DocResult.TotalCount)
    if ($script:DocResult.ExitCode -eq 1) { throw "$RunAction failed: $(@($script:DocResult.Issues | ForEach-Object { $_.Message }) -join '; ')" }
    [pscustomobject]@{ Records = $records; Result = $script:DocResult; Html = $script:DocResult.HtmlReport; GlobalHtml = $script:DocGlobal }
}

function Get-ShownText([string]$Text) {
    # Work folder -> the folder of a real installation; this computer and this account -> neutral names.
    $account = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $Text = $Text.Replace((Join-Path $work 'package'), $shown).Replace($work, $shown).Replace($account, 'CONTOSO\pra-admin')
    # Also as escaped in the JSON data of the HTML report.
    $Text = $Text.Replace((Join-Path $work 'package').Replace('\', '\\'), $shown.Replace('\', '\\')).Replace($work.Replace('\', '\\'), $shown.Replace('\', '\\')).Replace($account.Replace('\', '\\'), 'CONTOSO\\pra-admin')
    return $Text.Replace($env:COMPUTERNAME, 'PRA-ADMIN01')
}
#endregion
Write-Host 'Runs (in-memory tenant)...'
# Every poll of the tool advances the simulated tenant by one tick, as in the end-to-end tests.
function global:Start-Sleep { Step-PraFake }
try {
    $check = Invoke-DocRun -RunAction Check -Minutes 0.7
    $groupRun = Invoke-DocRun -RunAction Convert -RunGroup 'PRA-Wave1' -Minutes 0.4
    $convert = Invoke-DocRun -RunAction Convert -RunMode Apply -Minutes 9.6
    $follow = Invoke-DocRun -RunAction Check -RunBatch $convert.Result.BatchId -RunExpect Provisioned -RunOnce -Minutes 0.5
    $recover = Invoke-DocRun -RunAction Recover -RunMode Apply -RunBatch $convert.Result.BatchId -Minutes 14.4
} finally {
    Remove-Item function:global:Start-Sleep -ErrorAction SilentlyContinue
}
#endregion

#region Console images ------------------------------------------------------------------------------------------------
# Campbell palette of Windows Terminal for the dark picture; a darker one of the same 16 colours for the light picture (PRA_ANSI=1).
$palettes = @{
    dark  = @{ Background = '#0c0c0c'; Border = '#2b2b2b'; Bar = '#202020'; BarText = '#d0d0d0'; Text = '#cccccc'
        Colours = @{ 30 = '#0c0c0c'; 31 = '#c50f1f'; 32 = '#13a10e'; 33 = '#c19c00'; 34 = '#0037da'; 35 = '#881798'; 36 = '#3a96dd'; 37 = '#cccccc'
            90 = '#767676'; 91 = '#e74856'; 92 = '#16c60c'; 93 = '#f9f1a5'; 94 = '#3b78ff'; 95 = '#b4009e'; 96 = '#61d6d6'; 97 = '#f2f2f2' } }
    light = @{ Background = '#fbfbfb'; Border = '#cfcfcf'; Bar = '#e9e9e9'; BarText = '#3b3b3b'; Text = '#2b2b2b'
        Colours = @{ 30 = '#2b2b2b'; 31 = '#c50f1f'; 32 = '#0e7a0d'; 33 = '#8f6b00'; 34 = '#0037da'; 35 = '#881798'; 36 = '#0e7490'; 37 = '#4a4a4a'
            90 = '#6a6a6a'; 91 = '#c42b3b'; 92 = '#107c10'; 93 = '#8a5a00'; 94 = '#1f5fd8'; 95 = '#a0148c'; 96 = '#0f7b88'; 97 = '#1b1b1b' } }
}
function ConvertFrom-Ansi([string]$Line, [hashtable]$Colours) {
    $out = [Text.StringBuilder]::new(); $fg = $null
    foreach ($part in [regex]::Split($Line, '(\x1b\[[0-9;]*m)')) {
        if ($part -match '^\x1b\[([0-9;]*)m$') {
            $code = if ($Matches[1]) { [int]$Matches[1] } else { 0 }
            $fg = if ($Colours.ContainsKey($code)) { $Colours[$code] } else { $null }
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
    foreach ($theme in 'light', 'dark') {
        $p = $palettes[$theme]
        $body = ($lines | ForEach-Object { ConvertFrom-Ansi $_ $p.Colours }) -join "`n"
        $console = @"
<!doctype html><html><head><meta charset="utf-8"><style>
body { margin:0; background:$($p.Background); font-family:"Segoe UI", sans-serif; }
.win { width:1180px; margin:0; border-radius:10px; overflow:hidden; background:$($p.Background); border:1px solid $($p.Border); }
.bar { display:flex; align-items:center; gap:10px; height:38px; padding:0 14px; background:$($p.Bar); color:$($p.BarText); font-size:12.5px; }
.tab { padding:6px 14px; background:$($p.Background); border-radius:8px 8px 0 0; margin-top:8px; }
pre { margin:0; padding:14px 18px 18px; color:$($p.Text); font:13.5px/1.42 "Cascadia Mono", Consolas, monospace; white-space:pre-wrap; word-break:break-all; }
</style></head><body><div class="win"><div class="bar"><span class="tab">PowerShell 7.5</span></div>
<pre><span style="color:$($p.Text)">PS $shown&gt; $([Net.WebUtility]::HtmlEncode($Command))</span>
$body</pre></div>
<script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.win').getBoundingClientRect().height) + 2);</script></body></html>
"@
        $page = Join-Path $work "$Name-$theme.html"
        [IO.File]::WriteAllText($page, $console, [Text.UTF8Encoding]::new($true))
        Save-Page $page "$Name-$theme" 1182
    }
}

Write-Host 'Images:'
Save-Console $check.Records '.\Invoke-PraCloudMailbox.ps1 -Action Check' 'cli-check'
Save-Console $groupRun.Records ".\Invoke-PraCloudMailbox.ps1 -Action Convert -Group PRA-Wave1" 'cli-convert-group'
Save-Console $convert.Records '.\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply' 'cli-convert'
Save-Console $follow.Records ".\Invoke-PraCloudMailbox.ps1 -Action Check -Batch $($convert.Result.BatchId) -Expect Provisioned -Once" 'cli-follow'
Save-Console $recover.Records ".\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Apply -Batch $($convert.Result.BatchId)" 'cli-recover'
#endregion
#region Guided console pictures ---------------------------------------------------------------------------------------
if (-not $SkipConsole) {
Write-Host 'Photographing the console (synthetic engine)...'
$shots = Join-Path $work 'console'
New-Item -ItemType Directory -Path $shots -Force | Out-Null
$harness = Join-Path $PSScriptRoot 'DocumentationImages\Show-ConsoleScreens.ps1'
$pwsh = (Get-Process -Id $PID).Path
$running = @()
foreach ($job in @(foreach ($scenario in 'cloud', 'collect', 'problems') { foreach ($theme in 'Light', 'Dark') { @{ Scenario = $scenario; Theme = $theme } } })) {
    $log = Join-Path $work ("console-$($job.Scenario)-$($job.Theme).log")
    $argList = @('-NoProfile', '-STA', '-File', "`"$harness`"", '-Shots', "`"$shots`"", '-Theme', $job.Theme, '-Scenario', $job.Scenario, '-Repository', "`"$root`"")
    $running += [pscustomobject]@{ Job = $job; Log = $log; Process = (Start-Process -FilePath $pwsh -ArgumentList $argList -PassThru -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError ($log + '.err')) }
}
foreach ($r in $running) {
    $null = $r.Process.WaitForExit()
    $text = [IO.File]::ReadAllText($r.Log)
    if ($text -notmatch 'SHOTS-OK') { throw "Console harness failed ($($r.Job.Scenario), $($r.Job.Theme)):`n$text`n$([IO.File]::ReadAllText($r.Log + '.err'))" }
}
Get-ChildItem -LiteralPath $shots -Filter 'console-*.png' | ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $OutputFolder $_.Name) -Force }
Write-Host ('  {0} console pictures' -f @(Get-ChildItem -LiteralPath $shots -Filter 'console-*.png').Count)
}
#endregion

#region Report images -------------------------------------------------------------------------------------------------
function Save-Report([string]$Path, [string]$Name) {
    $original = Get-ShownText ([IO.File]::ReadAllText($Path))
    $measure = "<script>window.addEventListener('load', () => setTimeout(() => document.body.setAttribute('data-h', Math.ceil(document.body.getBoundingClientRect().height) + 24), 50));</script></body>"
    foreach ($theme in 'light', 'dark') {
        # The report follows the colour scheme of the browser: the picture of the other theme asks for it.
        $html = $original.Replace('window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light"', ('"{0}"' -f $theme)).Replace('@media (prefers-color-scheme:dark){', $(if ($theme -eq 'dark') { '@media all{' } else { '@media not all{' }))
        $page = Join-Path $work "$Name-$theme.html"
        [IO.File]::WriteAllText($page, $html.Replace('</body>', $measure), [Text.UTF8Encoding]::new($true))
        Save-Page $page "$Name-$theme" 1280
    }
}
Save-Report $check.Html 'report-check'
Save-Report $convert.Html 'report-convert'
if (-not $recover.GlobalHtml -or -not (Test-Path -LiteralPath $recover.GlobalHtml)) { throw 'The Recover run wrote no global report.' }
Save-Report $recover.GlobalHtml 'report-global'
#endregion
foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
Remove-Variable -Name PraFake, PraFakeCommands -Scope Global -ErrorAction SilentlyContinue
foreach ($name in 'PRA.Cloud', 'PRA.Store', 'PRA.Common') { Remove-Module $name -Force -ErrorAction SilentlyContinue }
Remove-Item Env:\PRA_ICONS, Env:\PRA_ANSI -ErrorAction SilentlyContinue
Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
