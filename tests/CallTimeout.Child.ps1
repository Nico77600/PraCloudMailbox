<#
.SYNOPSIS
    Child side of PraCloudMailbox.CallTimeout.Tests.ps1 (PowerShell 7).
.DESCRIPTION
    Dot-sourced by a copy of the entry script that runs in its own pwsh process, as a scheduled task does: it loads the modules of the test package, the in-memory tenant of
    FakeTenant.ps1, a snapshot of two users and a configuration file, and replaces the sign-in by the time limit of the calls (the real watchdog, the real guard of the
    Exchange Online commands). The scenario (environment variable PCM_TEST_SCENARIO) says which call of the fake tenant does not answer:
      ReadHang     Get-OrganizationConfig never answers, and reacts to a stop (an HTTP call that is waiting)
      ReadStuck    Get-OrganizationConfig never answers and does not react to a stop (a call blocked in a native wait)
      WriteHang    Set-MailUser never answers, and reacts to a stop
      WriteDone    the write goes through (the run that continues the batch)
    Whatever the scenario, the file guard-after.txt of the work folder says how many guard aliases and proxy commands the session holds at its end.
.NOTES
    Author : Nicolas Fabert
#>
$ErrorActionPreference = 'Stop'
$root = $env:PCM_TEST_ROOT; $work = $env:PCM_TEST_WORK; $scenario = $env:PCM_TEST_SCENARIO; $testDir = $env:PCM_TEST_DIR
foreach ($name in 'PRA.Common', 'PRA.Store', 'PRA.Cloud') { Import-Module (Join-Path $root "module\$name.psm1") -Force -Global }
. (Join-Path $testDir 'FakeTenant.ps1')

New-PraFakeTenant
$global:PraFake.CasePolicy.Name = 'PRA-Recover-Hold'
$global:PraFake.Group.Name = 'LIC-Exchange-Online'
$people = [ordered]@{}
$number = 0
foreach ($upn in 'adele.vance@contoso.com', 'megan.bowen@contoso.com') {
    $number++
    $suffix = '{0:D12}' -f $number
    $user = Add-PraFakeUser -Upn $upn -Id "a0000000-0000-0000-0000-$suffix" -OnPremGuid "b0000000-0000-0000-0000-$suffix" -Component "c0000000-0000-0000-0000-$suffix" -UsageLocation 'FR'
    $global:PraFake.Ad[$global:PraFake.Ad.Count - 1].ObjectGuid = "d0000000-0000-0000-0000-$suffix"
    $people[$upn] = $user
}

$config = Join-Path $work 'child.config.psd1'
if (-not (Test-Path -LiteralPath $config)) {
    $null = New-Item -ItemType Directory -Path (Join-Path $work 'data') -Force
    Set-Content -LiteralPath $config -Encoding UTF8 -Value @"
@{
    Environment = 'PROD'
    Scope = @{ Mode = 'OU'; SearchBase = 'OU=Paris,DC=contoso,DC=com' }
    Store = @{ Path = '$work\data\PraCloudMailbox.db'; JournalPath = '$work\data\PraCloudMailbox-journal.db'; BackupFolder = '$work\data\backup' }
    Cloud = @{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; AppId = 'app'; CertificateThumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'; DefaultUsageLocation = 'FR' }
    Polling = @{ IntervalSeconds = 20; MailboxTimeoutMinutes = 30; SyncTimeoutMinutes = 30; HoldTimeoutMinutes = 30; CallTimeoutMinutes = 5 }
    Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }; Shared = @{ SkuPartNumber = 'E5' } }
    Retention = @{ HoldPolicy = 'PRA-Recover-Hold' }
    EntraConnect = @{ Mode = 'Remoting'; Server = 'aadc01.contoso.com' }
    Logging = @{ Folder = '$work\logs' }
    Report = @{ Folder = '$work\reports' }
}
"@
    $byUpn = @{}; foreach ($ad in $global:PraFake.Ad) { $byUpn[$ad.Upn] = $ad }
    $db = Join-Path $work 'data\PraCloudMailbox.db'
    $connection = Open-PcmStore -Path $db -Root $root
    try {
        $id = New-PcmSnapshot $connection @{ RunId = [guid]::NewGuid().ToString('N'); Environment = 'PROD'; ToolVersion = (Get-PcmToolVersion); ExchangeServer = 'EX01.contoso.com'; ScopeJson = '{"Mode":"OU"}' }
        $rows = @(foreach ($upn in $people.Keys) {
                $ad = $byUpn[$upn]
                [ordered]@{ object_guid = $ad.ObjectGuid; kind = $ad.Kind; user_principal_name = $upn; primary_smtp_address = $upn; sam_account_name = ($upn -split '@')[0]
                    exchange_guid = $ad.OnPremGuid; immutable_id = $ad.ImmutableId; recipient_type_details = "$($ad.Kind)Mailbox" }
            })
        $null = Add-PcmRow $connection mailbox $id $rows
        Set-PcmSnapshotStatus $connection $id Complete @{ mailbox = $rows.Count }
    } finally { Close-PcmStore $connection }
    $null = Write-PcmHashFile -Path $db
}

# The waits of the tenant happen on the next tick, as in the other suites; the calls that must not answer sleep with the real command.
function global:Start-Sleep { Step-PraFake }
function global:Get-PcmGraphRole { @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
function global:Get-PcmCloudModuleReport { param($MinExo) [pscustomobject]@{ Ok = $true; Graph = [pscustomobject]@{ Version = '2.41.1' }; Exo = [pscustomobject]@{ Version = '3.10.1' }; Messages = @() } }
function global:Invoke-PcmEntraConnect { param([hashtable]$Context, [string]$Operation, [object]$Caller) $null = Invoke-PraFakeEntraConnect -Operation $Operation; 'SyncCycleEnabled=True InProgress=False' }
function global:Wait-PcmCondition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
& (Get-Module PRA.Cloud) { function script:Wait-PcmCondition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds } }

$global:PraChildRealSetMailUser = (Get-Command Set-MailUser -CommandType Function).ScriptBlock
function global:Get-OrganizationConfig {
    [CmdletBinding()] param()
    if ($env:PCM_TEST_SCENARIO -eq 'ReadHang') { Microsoft.PowerShell.Utility\Start-Sleep -Seconds 3600 }
    if ($env:PCM_TEST_SCENARIO -eq 'ReadStuck') { [Threading.Thread]::Sleep(3600000) }
    $global:PraFake.Reads.Add('Get-OrganizationConfig')
    [pscustomobject]@{ InPlaceHolds = @($global:PraFake.OrgHolds) }
}
function global:Set-MailUser {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [object]$ExchangeGuid)
    if ($env:PCM_TEST_SCENARIO -eq 'WriteHang') { Microsoft.PowerShell.Utility\Start-Sleep -Seconds 3600 }
    & $global:PraChildRealSetMailUser @PSBoundParameters
}

# The sign-in is replaced; the watchdog and the guard of the commands are the real ones, with a deadline of a few seconds.
function global:Connect-PcmCloud {
    param([hashtable]$Context)
    Start-PcmCallGuard -Context $Context
    & (Get-Module PRA.Cloud) { $script:CallGuard.Seconds = 3; $script:CallGuard.Grace = 3 }
    Install-PcmCallGuard
    'Microsoft Graph and Exchange Online (fake tenant)'
}
function global:Disconnect-PcmCloud { param([hashtable]$Context) Stop-PcmCallGuard }
# What the session still holds of the guard when the process ends (the test reads the file): the aliases and the proxy commands must be gone.
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    $left = @(Get-Alias -ErrorAction SilentlyContinue | Where-Object { [string]$_.Definition -like 'PcmGuard_*' })
    $proxies = @(Get-Command -CommandType Function -Name 'PcmGuard_*' -ErrorAction SilentlyContinue)
    [IO.File]::WriteAllText((Join-Path $env:PCM_TEST_WORK 'guard-after.txt'), ('aliases={0} functions={1}' -f $left.Count, $proxies.Count))
}
function global:Connect-PcmCompliance { }
