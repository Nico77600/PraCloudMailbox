<#
.SYNOPSIS
    PRA Cloud Mailbox - cloud module: Microsoft Graph and Exchange Online (PowerShell 7).

.DESCRIPTION
    Certificate sign-in of the app registration (Cloud.AppId + Cloud.CertificateThumbprint) to Microsoft
    Graph (Microsoft.Graph.Authentication) and Exchange Online (ExchangeOnlineManagement 3.10+).

    It holds every cloud operation of Check, Convert and Recover: the state of the Entra ID object and of the
    Exchange Online recipient of each object of the Collect snapshot, the rules that tell whether Convert can
    run (Get-PcmReadiness, pure functions of the snapshot row and of the cloud state, tested without any
    cloud access), source of authority, licences, shared mailbox permissions, the eDiscovery case hold
    (Security & Compliance) and the Entra Connect operations of Recover. The states of many objects are read in bulk
    (Get-PcmCloudStateSet, Get-PcmExoStateSet): a few calls per 50 objects instead of five calls per object.

    Facts behind the rules:
      - Convert of a user = ExchangeGuid cleared, SOA transferred (isCloudManaged), then an Exchange plan:
        Exchange Online promotes the Teams storage (ComponentShared) to the primary mailbox.
      - An Exchange plan that is already enabled before the GUID is cleared does nothing: it must be
        removed and added again.
      - Any hold on the cloud mailbox blocks the Recover of a user.
      - usageLocation synced from AD is refused for licences right after the SOA transfer: it is set
        again from the cloud.

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
# Application permissions (Graph roles) needed by Convert and Recover; Check reads them from the token.
$script:RequiredGraphRoles = @(
    @{ Role = 'User.ReadWrite.All'; Why = 'usageLocation, delete the identity of a shared mailbox (Recover)' }
    @{ Role = 'User-OnPremisesSyncBehavior.ReadWrite.All'; Why = 'transfer of the source of authority (isCloudManaged)' }
    @{ Role = 'LicenseAssignment.ReadWrite.All'; Why = 'temporary licence of shared mailboxes, direct or Kiosk licences' }
    @{ Role = 'Organization.Read.All'; Why = 'licence counters (subscribedSkus)' }
)
$script:GroupModeRoles = @(
    @{ Role = 'GroupMember.ReadWrite.All'; Why = 'membership of the licence group' }
    @{ Role = 'Group-OnPremisesSyncBehavior.ReadWrite.All'; Why = 'licence group synchronised from AD: transfer of its source of authority' }
)
# Exchange plans that give a mailbox. EXCHANGE_S_FOUNDATION gives none (Teams, Office): never counted as a mailbox plan.
$script:ExchangePlanPattern = '^EXCHANGE_S_(STANDARD|ENTERPRISE|DESKLESS|ESSENTIALS)'
# Minimum module versions (the same numbers as Get-PcmCloudModuleReport in PRA.Common).
$script:MinGraphVersion = '2.41.1'
$script:MinExoVersion = '3.10.1'

#region Connections ---------------------------------------------------------------------------------

# The Microsoft modules (Graph, Exchange Online, Security & Compliance) write their own warnings to the console, in the language of the
# computer and outside the format of the tool. Their warnings are collected instead (Read-PcmModuleWarning) and written to the log; their errors
# are untouched (every call keeps -ErrorAction and its try/catch).
$script:PcmModuleWarnings = [System.Collections.ArrayList]::new()
$script:QuietModuleCommands = @('Connect-MgGraph', 'Disconnect-MgGraph', 'Invoke-MgGraphRequest', 'Connect-ExchangeOnline', 'Disconnect-ExchangeOnline', 'Connect-IPPSSession',
    'Get-Recipient', 'Get-MailUser', 'Get-Mailbox', 'Get-MailboxLocation', 'Get-User', 'Get-OrganizationConfig', 'Set-MailUser', 'Set-Mailbox',
    'Add-MailboxPermission', 'Get-MailboxPermission', 'Add-RecipientPermission', 'Get-RecipientPermission',
    'New-Mailbox', 'Set-CalendarProcessing', 'Get-CalendarProcessing', 'Set-MailboxCalendarConfiguration', 'Get-MailboxCalendarConfiguration',
    'New-MailContact', 'Set-MailContact', 'Get-MailContact', 'Get-ResourceConfig', 'Set-User', 'Remove-MailContact', 'Remove-Mailbox',
    'Get-DistributionGroup', 'Get-DistributionGroupMember', 'Add-DistributionGroupMember', 'New-DistributionGroup', 'Set-DistributionGroup',
    'Get-CaseHoldPolicy', 'Set-CaseHoldPolicy', 'New-CaseHoldPolicy', 'New-CaseHoldRule')

function Get-PcmModuleQuietDefault {
    <# Default parameters that keep the warnings of the Microsoft module commands off the console and collect them for the log. #>
    [CmdletBinding()]
    param([string]$VariableName = 'PcmModuleWarnings')
    $defaults = @{}
    foreach ($name in $script:QuietModuleCommands) {
        $defaults["${name}:WarningAction"] = 'SilentlyContinue'
        $defaults["${name}:WarningVariable"] = '+' + $VariableName
    }
    return $defaults
}
$PSDefaultParameterValues = Get-PcmModuleQuietDefault

function Read-PcmModuleWarning {
    <#
    .SYNOPSIS
        The warnings that the Microsoft modules gave since the last call (each text once, in order), and forgets them.
    #>
    [CmdletBinding()]
    param()
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $texts = @(foreach ($item in $script:PcmModuleWarnings) {
            $text = ([string]$(if ($item -is [System.Management.Automation.WarningRecord]) { $item.Message } else { $item })).Trim()
            if ($text -and $seen.Add($text)) { $text }
        })
    $script:PcmModuleWarnings.Clear()
    return $texts
}

function Add-PcmModuleWarning {
    <# Brings in the warnings that a caller collected from the Microsoft modules (the entry script), to be logged with the others. #>
    [CmdletBinding()]
    param([AllowNull()][object[]]$Warning)
    foreach ($item in @($Warning)) { if ($null -ne $item) { [void]$script:PcmModuleWarnings.Add($item) } }
}
$script:GraphRequestOwner = $null
$script:GraphTransportReady = $false
$script:GraphPreviousMaxRetry = $null
$script:WriteApproval = $null

function Set-PcmWriteApproval {
    <# Registers (or clears) the approval of the run: the script block that asks ShouldProcess before every change to the tenant or to Exchange Online. #>
    [CmdletBinding()]
    param([AllowNull()][scriptblock]$Approval)
    $script:WriteApproval = $Approval
}

function Confirm-PcmWrite {
    <#
    .SYNOPSIS
        Asks the approval of the run before ONE change to Microsoft Entra ID, Exchange Online, the compliance case hold or Entra Connect.
        Stops the run when there is no approval or when it is refused: nothing is written.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Target, [Parameter(Mandatory)][string]$Operation)
    if ($script:WriteApproval -isnot [scriptblock]) { throw ("No approval is registered for '{0}' on {1}: nothing was written." -f $Operation, $Target) }
    $answer = @(& $script:WriteApproval $Target $Operation)
    if ($answer.Count -ne 1 -or $answer[0] -isnot [bool] -or -not $answer[0]) { throw ("Operation refused: {0} on {1}. The run stops here and nothing more is written." -f $Operation, $Target) }
}

function Test-PcmGraphWrite {
    <# Does this Microsoft Graph request change something? Reads sent as POST (membership check) and the batch envelope (each request in it is asked on its own) do not. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Method, [Parameter(Mandatory)][string]$Uri)
    if ([string]::Equals($Method, 'GET', [StringComparison]::OrdinalIgnoreCase)) { return $false }
    $path = ($Uri -replace '\?.*$', '').TrimEnd('/')
    if ($path -match '/checkMemberGroups$' -or $path -match '/\$batch$') { return $false }
    return $true
}

#region Time limit of the calls to Microsoft 365 ---------------------------------------------------------

# Every call to Microsoft 365 (Microsoft Graph, Exchange Online, Security & Compliance, the sign-ins) ends within Polling.CallTimeoutMinutes. The Exchange Online
# and Graph sessions belong to the runspace of the run, so no call is moved: a watchdog, alone in its own runspace, only watches a deadline and never runs a
# Microsoft command. A call still pending at its deadline is stopped the way Ctrl+C stops it: the pipeline of the run is stopped, its finally blocks run,
# and the entry script ends the run (the message, the next steps, exit code 1). A call that does not react to the stop is ended with the process after a
# short grace time. After the first deadline every new call to Microsoft 365 is refused at once, so a cleanup never waits again.
$script:CallGuard = $null
$script:CallGuarded = @{}
$script:CallGuardAliasBefore = @{}
$script:CallGuardLastTimeout = ''
$script:CallGuardGraceSeconds = 30
$script:CallGuardCleanupSeconds = 30
$script:CallGuardInteractiveSeconds = 900
$script:CallGuardClientMarginSeconds = 30
$script:GraphCallSeconds = 300
$script:GraphPreviousClientTimeout = $null

$script:CallGuardWatchdog = @'
param($s)
$finish = {
    $lines = @([string]$s.Timeout, 'The call could not be interrupted, so this run was ended at once: its summary and its report were not written.')
    if ($s.Extra) { $lines += @($s.Extra) }
    $lines += 'Next steps:'
    foreach ($step in @($s.NextSteps)) { $lines += ('  ' + $step) }
    $lines += 'RESULT: FAIL'
    $text = ($lines -join [Environment]::NewLine)
    try { [Console]::Error.WriteLine($text) } catch { $null = $_ }
    try { $log = [string]$s.Context['LogFile']; if ($log) { [IO.File]::AppendAllText($log, $text + [Environment]::NewLine) } } catch { $null = $_ }
    [Environment]::Exit(1)
}
while (-not $s.Quit) {
    [Threading.Thread]::Sleep(250)
    if (-not $s.Armed) { continue }
    if ([DateTime]::UtcNow.Ticks -lt [long]$s.DeadlineTicks) { continue }
    $s.Expirations = [int]$s.Expirations + 1
    if ($s.Expirations -eq 1) { $s.Timeout = [string]$s.Message }
    $s.Expired = $true
    $s.Armed = $false
    if ($s.Expirations -gt 1 -or $null -eq $s.Pipeline) { & $finish }
    $s.Released = $false
    try { $s.Pipeline.StopAsync() } catch { & $finish }
    $until = [DateTime]::UtcNow.AddSeconds([int]$s.Grace)
    while (-not $s.Released -and -not $s.Quit -and [DateTime]::UtcNow -lt $until) { [Threading.Thread]::Sleep(250) }
    if (-not $s.Released -and -not $s.Quit) { & $finish }
}
'@

function Get-PcmCallTimeoutSecond {
    <# The seconds the tool waits for ONE answer of Microsoft 365: Polling.CallTimeoutMinutes, 5 minutes when the setting is not there. #>
    [CmdletBinding()]
    param([AllowNull()][object]$Config)
    $minutes = 5
    $polling = Get-PcmValue $Config 'Polling' $null
    $value = Get-PcmValue $polling 'CallTimeoutMinutes' $null
    if ($value -is [int] -and $value -ge 1) { $minutes = $value }
    return ($minutes * 60)
}

function Format-PcmCallDuration {
    <# A time limit in words: 5 minutes, 1 minute, 90 seconds. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$Seconds)
    if ($Seconds -ge 60 -and $Seconds % 60 -eq 0) { $minutes = [int]($Seconds / 60); return ('{0} minute{1}' -f $minutes, $(if ($minutes -eq 1) { '' } else { 's' })) }
    return ('{0} second{1}' -f $Seconds, $(if ($Seconds -eq 1) { '' } else { 's' }))
}

function Get-PcmCallTimeoutMessage {
    <# The sentence of a stop: what did not answer, during which step, and what to do. A change that did not answer may or may not have been applied. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$Seconds, [string]$Step = '', [string]$Call = '', [switch]$Write)
    $where = if ($Step) { $Step } else { 'this run' }
    $head = 'Microsoft 365 did not answer within {0} during {1}{2}' -f (Format-PcmCallDuration $Seconds), $where, $(if ($Call) { " ($Call)" } else { '' })
    if ($Write) {
        return ('{0}. That last change may or may not have been applied, and nothing else was changed: the next run reads it back before going on. Check https://status.cloud.microsoft and the service health of your tenant, then run the command given in the next steps.' -f $head)
    }
    return ('{0}; nothing more was changed. Check https://status.cloud.microsoft and the service health of your tenant, then run the same command again.' -f $head)
}

function Get-PcmCallTimeout {
    <# The message of the stop when a call to Microsoft 365 did not answer in time in this run, otherwise an empty text. #>
    [CmdletBinding()]
    param()
    $state = $script:CallGuard
    if ($null -ne $state -and [string]$state.Timeout) { return [string]$state.Timeout }
    return [string]$script:CallGuardLastTimeout
}

function Get-PcmRunningPipeline {
    <# The pipeline that runs the current call, or nothing when this PowerShell does not give it (the member is internal: it can move between versions). Never throws. #>
    [CmdletBinding()]
    param([string]$MemberName = 'GetCurrentlyRunningPipeline')
    try {
        $runspace = [System.Management.Automation.Runspaces.Runspace]::DefaultRunspace
        $method = $runspace.GetType().GetMethod($MemberName, [Reflection.BindingFlags]'Instance,Public,NonPublic', $null, [Type]::EmptyTypes, $null)
        if ($null -eq $method) { return $null }
        return $method.Invoke($runspace, @())
    } catch { return $null }
}

function Start-PcmCallGuard {
    <# Starts the watchdog of the run (once, before the first sign-in). The pipeline it can stop is the one that runs this call. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($null -ne $script:CallGuard) { Stop-PcmCallGuard }
    $script:CallGuardLastTimeout = ''
    $pipeline = Get-PcmRunningPipeline
    if ($null -eq $pipeline) {
        try { Write-PcmLog -Context $Context -Message 'The pipeline of this run cannot be reached on this PowerShell: a call to Microsoft 365 that does not answer in time ends the process (exit code 1) instead of stopping the run.' -Level Detail } catch { $null = $_ }
    }
    $extra = @()
    if (Test-PcmSameText ([string](Get-PcmValue $Context 'Action' '')) 'Recover') { $extra = @('If this run had paused the Entra Connect scheduler, resume it by hand: Set-ADSyncScheduler -SyncCycleEnabled $true.') }
    $state = [hashtable]::Synchronized(@{
            Armed = $false; Depth = 0; Epoch = 0; DeadlineTicks = [long]0; Message = ''; Timeout = ''; Expired = $false; Released = $false; Expirations = 0; Quit = $false
            Pipeline = $pipeline; Context = $Context; Seconds = (Get-PcmCallTimeoutSecond -Config (Get-PcmValue $Context 'Config' $null)); Grace = $script:CallGuardGraceSeconds
            NextSteps = @(); NextFor = $null; Extra = $extra; Runspace = $null; Watcher = $null; Handle = $null })
    try {
        $state.Runspace = [runspacefactory]::CreateRunspace([initialsessionstate]::CreateDefault2())
        $state.Runspace.Open()
        $state.Watcher = [powershell]::Create()
        $state.Watcher.Runspace = $state.Runspace
        $null = $state.Watcher.AddScript($script:CallGuardWatchdog).AddArgument($state)
        $state.Handle = $state.Watcher.BeginInvoke()
    } catch {
        $message = $_.Exception.Message
        foreach ($name in @('Watcher', 'Runspace')) { try { if ($null -ne $state[$name]) { $state[$name].Dispose() } } catch { $null = $_ } }
        throw ('The time limit of the calls to Microsoft 365 could not be started ({0}): nothing was done.' -f $message)
    }
    $script:CallGuard = $state
}

function Stop-PcmCallGuard {
    <# Stops the watchdog and removes the commands that go through it. What it saw (the message of a stop) is kept for the end of the run. #>
    [CmdletBinding()]
    param()
    Uninstall-PcmCallGuard
    $state = $script:CallGuard
    if ($null -eq $state) { return }
    $script:CallGuard = $null
    $script:CallGuardLastTimeout = [string]$state.Timeout
    $state.Quit = $true; $state.Armed = $false
    try { if ($null -ne $state.Handle) { $null = $state.Handle.AsyncWaitHandle.WaitOne(2000) } } catch { $null = $_ }
    foreach ($name in @('Watcher', 'Runspace')) { try { if ($null -ne $state[$name]) { $state[$name].Dispose() } } catch { $null = $_ } }
}

function Enter-PcmCall {
    <#
    .SYNOPSIS
        Arms the deadline of one call to Microsoft 365. Calls inside a call share the deadline of the outer one. -Cleanup is for the calls that end a run:
        they are still allowed after a stop, with a short deadline of their own.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Call, [switch]$Write, [int]$Seconds = 0, [switch]$Cleanup)
    $state = $script:CallGuard
    if ($null -eq $state) { return }
    if ($state.Expired -and -not $Cleanup) { throw [System.InvalidOperationException]::new([string]$state.Timeout) }
    $state.Depth = [int]$state.Depth + 1
    if ($state.Depth -gt 1) { return }
    $limit = if ($Seconds -gt 0) { $Seconds } else { [int]$state.Seconds }
    $context = $state.Context
    $step = [string](Get-PcmValue $context 'CurrentPhase' '')
    if (-not $step) { $step = [string](Get-PcmValue $context 'CurrentOperation' '') }
    $state.Message = Get-PcmCallTimeoutMessage -Seconds $limit -Step $step -Call $Call -Write:$Write
    $batch = [string](Get-PcmValue $context 'BatchId' '')
    if ($state.NextFor -ne $batch) {
        $state.NextFor = $batch
        $provider = Get-PcmValue $context 'FallbackProvider' $null
        if ($provider -is [scriptblock]) { try { $state.NextSteps = [string[]]@(& $provider | ForEach-Object { [string]$_ }) } catch { $null = $_ } }
    }
    $state.Epoch = [int]$state.Epoch + 1
    $state.DeadlineTicks = [DateTime]::UtcNow.AddSeconds($limit).Ticks
    $state.Armed = $true
}

function Exit-PcmCall {
    <# Disarms the deadline of the call that ends. When that call was the one that did not answer, the stop reached it: the watchdog is told. #>
    [CmdletBinding()]
    param()
    $state = $script:CallGuard
    if ($null -eq $state) { return }
    if ($state.Depth -gt 0) { $state.Depth = [int]$state.Depth - 1 }
    if ($state.Depth -eq 0) {
        $state.Armed = $false
        if ($state.Expired) { $state.Released = $true }
    }
}

function Set-PcmCallExpired {
    <# A call ended by the time limit of its own transport (the Graph client): the run stops like for a deadline of the watchdog. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message)
    $state = $script:CallGuard
    if ($null -eq $state) { return }
    if (-not [string]$state.Timeout) { $state.Timeout = $Message }
    $state.Expired = $true
}

function Get-PcmGuardedCommandName {
    <# The Exchange Online and Security & Compliance commands that go through the time limit (the sign-ins and Graph have their own place). #>
    [CmdletBinding()]
    param()
    return @(@($script:QuietModuleCommands | Where-Object { $_ -notmatch '^(Connect|Disconnect|Invoke)-' }) + @('Get-EXORecipient', 'Get-EXOMailbox') | Get-PcmDistinctText)
}

function Install-PcmCallGuard {
    <#
    .SYNOPSIS
        Puts the time limit in front of the Exchange Online and Security & Compliance commands that exist now. Each one gets a proxy with the same parameters
        and the same pipeline behaviour (an alias of the command name points to it); the proxy arms the deadline when the call starts and disarms it
        when it ends. Run it again after a connection that adds commands: the commands already guarded are left alone.
    #>
    [CmdletBinding()]
    param([string[]]$Name = @(Get-PcmGuardedCommandName))
    $state = $script:CallGuard
    if ($null -eq $state) { return }
    foreach ($commandName in $Name) {
        if ($script:CallGuarded.ContainsKey($commandName)) { continue }
        $original = Get-Command -Name $commandName -CommandType Function, Cmdlet -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -eq $original) { continue }
        try {
            $proxyName = 'PcmGuard_' + $commandName
            $text = [System.Management.Automation.ProxyCommand]::Create([System.Management.Automation.CommandMetadata]::new($original))
            $text = [regex]::Replace($text, '\$wrappedCmd = [^\r\n]+', ('$wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand(''{0}'', [System.Management.Automation.CommandTypes]''Function,Cmdlet'')' -f $commandName))
            $enter = "Enter-PcmCall -Call '{0}'{1}" -f $commandName, $(if ($commandName -match '^Get-') { '' } else { ' -Write' })
            $text = [regex]::Replace($text, '(?s)(begin\s*\{\s*try\s*\{)', { param($m) $m.Groups[1].Value + "`n        " + $enter + "`n" })
            if ($text -match 'clean\s*\{') { $text = [regex]::Replace($text, 'clean\s*\{', "clean {`n        Exit-PcmCall`n") } else { $text += "`nclean { Exit-PcmCall }`n" }
            # An alias of the same name that someone else made (another tool in this session) is put back when the guard goes.
            $script:CallGuardAliasBefore[$commandName] = Get-Alias -Name $commandName -Scope Global -ErrorAction SilentlyContinue | Select-Object -First 1
            Set-Item -Path ('Function:\global:' + $proxyName) -Value $ExecutionContext.InvokeCommand.NewScriptBlock($text)
            Set-Alias -Name $commandName -Value $proxyName -Scope Global -Force
            $script:CallGuarded[$commandName] = $proxyName
        } catch {
            Write-PcmLog -Context $state.Context -Message ('The time limit could not be put on {0}: {1}' -f $commandName, $_.Exception.Message) -Level Detail
        }
    }
}

function Uninstall-PcmCallGuard {
    <# Removes the proxies: the commands are called directly again. #>
    [CmdletBinding()]
    param()
    foreach ($commandName in @($script:CallGuarded.Keys)) {
        $proxyName = [string]$script:CallGuarded[$commandName]
        $ours = Get-Alias -Name $commandName -Scope Global -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($null -ne $ours -and [string]$ours.Definition -eq $proxyName) { Remove-Item -Path ('Alias:\' + $commandName) -Force -ErrorAction SilentlyContinue }
        foreach ($path in @(('Function:\global:' + $proxyName), ('Function:\' + $proxyName))) { Remove-Item -Path $path -Force -ErrorAction SilentlyContinue }
        $before = $script:CallGuardAliasBefore[$commandName]
        if ($null -ne $before -and -not (Get-Alias -Name $commandName -Scope Global -ErrorAction SilentlyContinue)) { Set-Alias -Name $commandName -Value ([string]$before.Definition) -Scope Global -Force -ErrorAction SilentlyContinue }
        $script:CallGuardAliasBefore.Remove($commandName)
        $script:CallGuarded.Remove($commandName)
    }
}

function Test-PcmClientTimeout {
    <# Did the call end because the HTTP client of the Graph SDK gave up waiting (its own time limit)? #>
    [CmdletBinding()]
    param([AllowNull()][Exception]$Exception)
    for ($e = $Exception; $e; $e = $e.InnerException) {
        if ([string]$e.Message -match 'HttpClient\.Timeout' -or ($e -is [Threading.Tasks.TaskCanceledException] -and $e.InnerException -is [TimeoutException])) { return $true }
    }
    return $false
}
#endregion

function Initialize-PcmGraphTransport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    foreach ($name in @('Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext')) {
        $null = Get-Command $name -ErrorAction Stop
    }
    if ($null -ne $script:GraphRequestOwner -or (Get-MgContext -ErrorAction Stop)) {
        throw 'Graph transport requires a fresh, unconnected PowerShell process; an existing SDK session cannot be reused safely.'
    }
    $request = Get-MgRequestContext -ErrorAction Stop
    if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry) {
        throw 'Cannot verify the Graph SDK request context; no cloud request is allowed.'
    }
    $script:GraphPreviousMaxRetry = [int]$request.MaxRetry
    $script:GraphPreviousClientTimeout = $(if ($request.PSObject.Properties['ClientTimeout'] -and $null -ne $request.ClientTimeout) { [timespan]$request.ClientTimeout } else { $null })
    $script:GraphCallSeconds = Get-PcmCallTimeoutSecond -Config (Get-PcmValue $Context 'Config' $null)
    $script:GraphRequestOwner = $Context
    # Set before authentication: this cmdlet discards the SDK's cached HTTP client. The client gives up a little after the watchdog would have stopped the call.
    $clientTimeout = [timespan]::FromSeconds($script:GraphCallSeconds + $script:CallGuardClientMarginSeconds)
    $null = Set-MgRequestContext -MaxRetry 0 -ClientTimeout ([int]$clientTimeout.TotalSeconds) -ErrorAction Stop
    $request = Get-MgRequestContext -ErrorAction Stop
    if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry -or $request.MaxRetry -ne 0) {
        throw 'Graph SDK MaxRetry=0 was not verified; no cloud request is allowed.'
    }
    if ($request.PSObject.Properties['ClientTimeout'] -and [timespan]$request.ClientTimeout -ne $clientTimeout) {
        throw 'Graph SDK ClientTimeout was not verified; no cloud request is allowed.'
    }
    $script:GraphTransportReady = $true
}

function Assert-PcmGraphTransport {
    [CmdletBinding()]
    param()
    if (-not $script:GraphTransportReady) { throw 'Graph transport is not initialized by Connect-PcmCloud; no cloud request is allowed.' }
    $request = Get-MgRequestContext -ErrorAction Stop
    if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry -or $request.MaxRetry -ne 0) {
        $script:GraphTransportReady = $false
        throw 'Graph SDK retry settings changed or cannot be verified; start a fresh PowerShell process. No cloud request is allowed.'
    }
}

function Connect-PcmCloud {
    <#
    .SYNOPSIS
        Certificate sign-in to Microsoft Graph and Exchange Online. Returns a short description.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $cloud = $Context.Config.Cloud
    if (Test-PcmSameText (Get-PcmValue $cloud 'SignIn' 'Certificate') 'Interactive') { return Connect-PcmCloudInteractive -Context $Context }
    foreach ($name in @('TenantId','Organization','AppId','CertificateThumbprint')) {
        if (-not $cloud[$name]) { throw "Configuration: Cloud.$name is empty (certificate sign-in of the app registration, see the administrator guide, chapter 4.3)." }
    }
    $certificate = Get-ChildItem -Path "Cert:\CurrentUser\My\$($cloud.CertificateThumbprint)", "Cert:\LocalMachine\My\$($cloud.CertificateThumbprint)" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $certificate) { throw "Certificate $($cloud.CertificateThumbprint) not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My." }
    if (-not $certificate.HasPrivateKey) { throw "Certificate $($cloud.CertificateThumbprint) has no private key on this computer." }
    if ($certificate.NotAfter -lt (Get-Date).AddDays(30)) { $Context['CertificateWarning'] = "The certificate expires on $($certificate.NotAfter.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture))." }
    if ($Context['Approval'] -isnot [scriptblock]) { throw 'The run gives no approval for its changes: it cannot sign in to the cloud. Nothing was done.' }
    Import-Module Microsoft.Graph.Authentication -MinimumVersion $script:MinGraphVersion -ErrorAction Stop
    try {
        Set-PcmWriteApproval -Approval $Context['Approval']
        Start-PcmCallGuard -Context $Context
        Initialize-PcmGraphTransport -Context $Context
        $Context['GraphConnected'] = $true # Own even a partially failed authentication, for cleanup.
        Enter-PcmCall -Call 'Connect-MgGraph'
        try { Connect-MgGraph -ClientId $cloud.AppId -TenantId $cloud.TenantId -CertificateThumbprint $cloud.CertificateThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop } finally { Exit-PcmCall }
        Assert-PcmGraphTransport
        Import-Module ExchangeOnlineManagement -MinimumVersion ([string](Get-PcmValue (Get-PcmValue $Context.Config 'Exo' @{}) 'MinModuleVersion' $script:MinExoVersion)) -ErrorAction Stop
        Enter-PcmCall -Call 'Connect-ExchangeOnline'
        try {
            Connect-ExchangeOnline -AppId $cloud.AppId -CertificateThumbprint $cloud.CertificateThumbprint -Organization $cloud.Organization -ShowBanner:$false `
                -CommandName @(Get-PcmExoCommandName -Context $Context) -ErrorAction Stop
        } finally { Exit-PcmCall }
        $Context['ExoConnected'] = $true
        Install-PcmCallGuard
        return ('Microsoft Graph and Exchange Online, app {0}, certificate {1} (valid until {2})' -f $cloud.AppId, $cloud.CertificateThumbprint.Substring(0, 8), $certificate.NotAfter.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture))
    } catch {
        Disconnect-PcmCloud -Context $Context
        throw
    }
}

function Get-PcmExoCommandName {
    <# The Exchange Online cmdlets the tool loads (only these are imported): the native resource cmdlets only when native resources are in the run. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $commands = @('Get-Recipient','Get-MailUser','Get-Mailbox','Get-MailboxLocation','Get-User','Get-OrganizationConfig','Set-MailUser','Set-Mailbox',
        'Add-MailboxPermission','Get-MailboxPermission','Add-RecipientPermission','Get-RecipientPermission')
    if ([bool](Get-PcmValue $Context 'NativeResourceScope' (Test-PcmSameText ([string](Get-PcmValue $Context.Config.Scope 'ResourceMode' 'SynchronizedShared')) 'CloudResource'))) {
        $commands += @('New-Mailbox','Set-CalendarProcessing','Get-CalendarProcessing','Set-MailboxCalendarConfiguration','Get-MailboxCalendarConfiguration',
            'New-MailContact','Set-MailContact','Get-MailContact','Get-ResourceConfig','Set-User',
            'Get-DistributionGroup','Get-DistributionGroupMember','Add-DistributionGroupMember','New-DistributionGroup','Set-DistributionGroup')
    }
    return $commands
}

function Connect-PcmCloudInteractive {
    <#
    .SYNOPSIS
        Interactive sign-in of an administrator (Cloud.SignIn = Interactive) to Microsoft Graph (delegated permissions) and Exchange Online. The
        sign-in window needs a console where someone can answer: it cannot run with -NonInteractive.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $cloud = $Context.Config.Cloud
    if (-not $cloud['TenantId']) { throw 'Configuration: Cloud.TenantId is empty (the interactive sign-in must go to the right tenant).' }
    if (Test-PcmNonInteractiveProcess) {
        throw 'Cloud.SignIn = Interactive needs a window where you can sign in, and this PowerShell was started with -NonInteractive. Open PowerShell 7 normally (without -NonInteractive) and run the same command again.'
    }
    if ($Context['Approval'] -isnot [scriptblock]) { throw 'The run gives no approval for its changes: it cannot sign in to the cloud. Nothing was done.' }
    Import-Module Microsoft.Graph.Authentication -MinimumVersion $script:MinGraphVersion -ErrorAction Stop
    try {
        Set-PcmWriteApproval -Approval $Context['Approval']
        Start-PcmCallGuard -Context $Context
        Initialize-PcmGraphTransport -Context $Context
        $Context['GraphConnected'] = $true
        if (Get-Command Set-MgGraphOption -ErrorAction SilentlyContinue) { try { $null = Set-MgGraphOption -EnableLoginByWAM $true -ErrorAction Stop } catch { $null = $_ } }
        $scopes = @(Get-PcmDelegatedScope -Config $Context.Config)
        # A person answers the sign-in window: its limit is longer than the one of a call.
        $signInSeconds = [Math]::Max((Get-PcmCallTimeoutSecond -Config $Context.Config), $script:CallGuardInteractiveSeconds)
        Enter-PcmCall -Call 'Connect-MgGraph' -Seconds $signInSeconds
        try { Connect-MgGraph -Scopes $scopes -TenantId $cloud.TenantId -ContextScope Process -NoWelcome -ErrorAction Stop } finally { Exit-PcmCall }
        Assert-PcmGraphTransport
        $account = [string](Get-PcmValue (Get-MgContext) 'Account' '')
        Import-Module ExchangeOnlineManagement -MinimumVersion ([string](Get-PcmValue (Get-PcmValue $Context.Config 'Exo' @{}) 'MinModuleVersion' $script:MinExoVersion)) -ErrorAction Stop
        $connect = @{ ShowBanner = $false; CommandName = @(Get-PcmExoCommandName -Context $Context); ErrorAction = 'Stop' }
        $hint = if ([string]$cloud['UserPrincipalName']) { [string]$cloud['UserPrincipalName'] } else { $account }
        if ($hint) { $connect.UserPrincipalName = $hint }
        Enter-PcmCall -Call 'Connect-ExchangeOnline' -Seconds $signInSeconds
        try { Connect-ExchangeOnline @connect } finally { Exit-PcmCall }
        $Context['ExoConnected'] = $true
        Install-PcmCallGuard
        return ('Microsoft Graph and Exchange Online, interactive sign-in{0} (delegated permissions: {1})' -f $(if ($account) { " as $account" } else { '' }), ($scopes -join ', '))
    } catch {
        Disconnect-PcmCloud -Context $Context
        throw
    }
}

function Disconnect-PcmCloud {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    Uninstall-PcmCallGuard
    if ($Context.ContainsKey('ExoConnected') -and $Context['ExoConnected']) {
        Enter-PcmCall -Call 'Disconnect-ExchangeOnline' -Seconds $script:CallGuardCleanupSeconds -Cleanup
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop 6>$null | Out-Null } catch { $null = $_ } finally { Exit-PcmCall }
        $Context['ExoConnected'] = $false
    }
    foreach ($text in @(Read-PcmModuleWarning)) { try { Write-PcmLog $Context ('Microsoft 365 module warning: ' + $text) -Level Detail } catch { $null = $_ } }
    if ([object]::ReferenceEquals($script:GraphRequestOwner, $Context)) {
        $script:GraphTransportReady = $false
        Set-PcmWriteApproval -Approval $null
        if ($Context.ContainsKey('GraphConnected') -and $Context['GraphConnected']) {
            Enter-PcmCall -Call 'Disconnect-MgGraph' -Seconds $script:CallGuardCleanupSeconds -Cleanup
            try { $null = Disconnect-MgGraph -ErrorAction Stop } catch { $null = $_ } finally { Exit-PcmCall }
            $Context['GraphConnected'] = $false
        }
        try {
            # Never reset a cached authenticated client just to restore request settings.
            if (Get-MgContext -ErrorAction Stop) { throw 'Graph is still connected; SDK request settings were not restored.' }
            $restore = @{ MaxRetry = $script:GraphPreviousMaxRetry }
            if ($null -ne $script:GraphPreviousClientTimeout) { $restore.ClientTimeout = [int][Math]::Ceiling($script:GraphPreviousClientTimeout.TotalSeconds) }
            $null = Set-MgRequestContext @restore -ErrorAction Stop
            $request = Get-MgRequestContext -ErrorAction Stop
            if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry -or $request.MaxRetry -ne $script:GraphPreviousMaxRetry) {
                throw 'Graph SDK request settings restoration could not be verified; use a fresh PowerShell process.'
            }
            if ($restore.ContainsKey('ClientTimeout') -and $request.PSObject.Properties['ClientTimeout'] -and [timespan]$request.ClientTimeout -ne $script:GraphPreviousClientTimeout) {
                throw 'Graph SDK request settings restoration could not be verified; use a fresh PowerShell process.'
            }
            $script:GraphRequestOwner = $null
            $script:GraphPreviousMaxRetry = $null
            $script:GraphPreviousClientTimeout = $null
        } catch { Write-Warning $_.Exception.Message }
    }
    if ($null -ne $script:CallGuard -and [object]::ReferenceEquals($script:CallGuard.Context, $Context)) { Stop-PcmCallGuard }
}

function Test-PcmNetworkFailure {
    <#
    .SYNOPSIS
        A failure of the connection itself, without an HTTP answer: timeout, reset, name resolution. It can happen
        after a long delta cycle. Only calls without -NoRetry can be sent again: reads, PATCH of a value, licence
        assignment, group member added when missing, protected legacy DELETE allowing 404.
    #>
    [CmdletBinding()]
    param([AllowNull()][Exception]$Exception)
    for ($e = $Exception; $e; $e = $e.InnerException) {
        if ($e -is [Net.Sockets.SocketException] -or $e -is [IO.IOException] -or $e -is [TimeoutException] -or $e -is [Threading.Tasks.TaskCanceledException]) { return $true }
        # HttpRequestException without a status code (.NET 5+) or without the property (.NET Framework): no answer at all.
        if ($e.GetType().FullName -eq 'System.Net.Http.HttpRequestException' -and (-not $e.PSObject.Properties['StatusCode'] -or $null -eq $e.StatusCode)) { return $true }
    }
    return $false
}

function Invoke-PcmGraph {
    <#
    .SYNOPSIS
        Microsoft Graph request with retries on throttling (429), transient errors (502, 503, 504) and connection
        failures without an answer (Test-PcmNetworkFailure). Returns $null for 404 when -AllowNotFound.
        The SDK transient retry handler is disabled before sign-in; this wrapper owns the retry budget.
        -NoRetry never resends an ambiguous response, for unprotected deletion. SDK authentication challenges
        and redirects are not controlled by MaxRetry.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, [ValidateSet('GET','POST','PATCH','DELETE')][string]$Method = 'GET', [object]$Body, [switch]$AllowNotFound, [switch]$NoRetry,
        [System.Collections.IDictionary]$Headers, [switch]$Approved)
    if (-not $Approved -and (Test-PcmGraphWrite -Method $Method -Uri $Uri)) { Confirm-PcmWrite -Target (($Uri -replace '\?.*$', '') -replace '^v1\.0/', '') -Operation "Microsoft Graph $Method" }
    $parameters = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $parameters.Body = ($Body | ConvertTo-Json -Depth 10 -Compress); $parameters.ContentType = 'application/json' }
    if ($Headers -and $Headers.Count) { $parameters.Headers = $Headers }
    $changes = Test-PcmGraphWrite -Method $Method -Uri $Uri
    if (-not $changes -and $Body -is [System.Collections.IDictionary] -and $Body.Contains('requests') -and ($Uri -replace '\?.*$', '').TrimEnd('/') -match '/\$batch$') {
        $changes = @(@($Body['requests']) | Where-Object { -not (Test-PcmSameText ([string](Get-PcmValue $_ 'method' 'GET')) 'GET') }).Count -gt 0
    }
    $callName = 'Microsoft Graph {0} {1}' -f $Method, (($Uri -replace '\?.*$', '') -replace '^v1\.0/', '')
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        Assert-PcmGraphTransport
        Enter-PcmCall -Call $callName -Write:$changes
        $armed = $true
        try { return Invoke-MgGraphRequest @parameters }
        catch {
            Exit-PcmCall; $armed = $false
            if (Test-PcmClientTimeout $_.Exception) {
                $message = Get-PcmCallTimeoutMessage -Seconds $script:GraphCallSeconds -Step ([string](Get-PcmValue $script:GraphRequestOwner 'CurrentPhase' '')) -Call $callName -Write:$changes
                Set-PcmCallExpired -Message $message
                throw [System.TimeoutException]::new($message, $_.Exception)
            }
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 }
            if ($code -eq 404 -and $AllowNotFound) { return $null }
            if (-not $NoRetry -and $code -in @(429, 502, 503, 504) -and $attempt -lt 5) { Start-Sleep -Seconds (5 * $attempt); continue }
            if (-not $NoRetry -and $code -eq 0 -and $attempt -lt 5 -and (Test-PcmNetworkFailure $_.Exception)) { Start-Sleep -Seconds (10 * $attempt); continue }
            # The Graph message (error.code: error.message) says what is wrong; the exception alone only says "BadRequest".
            $detail = ''
            try { $graphError = ([string]$_.ErrorDetails.Message | ConvertFrom-PcmJson).error; $detail = '{0}: {1}' -f $graphError.code, $graphError.message } catch { $detail = '' }
            if ($detail) { throw [System.InvalidOperationException]::new(('{0} {1} failed ({2}): {3}' -f $Method, ($Uri -replace '\?.*$', ''), $code, $detail), $_.Exception) }
            throw
        }
        finally { if ($armed) { Exit-PcmCall } }
    }
}

function Get-PcmGraphRole {
    <# Permissions of the current Microsoft Graph token: the application permissions (roles) of the app, or the delegated permissions (scopes) of the administrator who signed in. #>
    [CmdletBinding()]
    param()
    $context = Get-MgContext
    if (-not $context) { throw 'Not signed in to Microsoft Graph.' }
    return @($context.Scopes | ForEach-Object { [string]$_ })
}

function Get-PcmDelegatedScope {
    <#
    .SYNOPSIS
        The Microsoft Graph delegated permissions asked at an interactive sign-in: the delegated equivalents of the application permissions that
        Convert and Recover need (the same names), according to the licensing mode and the resources of the configuration.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config)
    $scopes = [System.Collections.Generic.List[string]]::new()
    foreach ($role in $script:RequiredGraphRoles) { $scopes.Add([string]$role.Role) }
    $groupMode = [string](Get-PcmValue (Get-PcmValue (Get-PcmValue $Config 'Licensing' @{}) 'Users' @{}) 'Mode' 'Group')
    if (Test-PcmSameText $groupMode 'Group') { foreach ($role in $script:GroupModeRoles) { $scopes.Add([string]$role.Role) } }
    elseif (Test-PcmSameText ([string](Get-PcmValue (Get-PcmValue $Config 'Scope' @{}) 'ResourceMode' 'SynchronizedShared')) 'CloudResource') { $scopes.Add([string]$script:GroupModeRoles[0].Role) }
    if (-not @($scopes | Where-Object { Test-PcmSameText $_ 'GroupMember.ReadWrite.All' }).Count) { $scopes.Add('GroupMember.Read.All') }
    return @($scopes | Get-PcmDistinctText)
}

function Test-PcmNonInteractiveProcess {
    <# Was this PowerShell started with -NonInteractive (a scheduled task, the window's engine)? An interactive sign-in cannot be shown there. #>
    [CmdletBinding()]
    param([string[]]$CommandLine = [Environment]::GetCommandLineArgs())
    foreach ($argument in @($CommandLine)) {
        $name = ([string]$argument).TrimStart('-', '/')
        if (([string]$argument).Length -gt 1 -and ([string]$argument)[0] -in @('-', '/') -and $name.Length -ge 4 -and 'noninteractive'.StartsWith($name, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

$script:GroupReadRoles = @('GroupMember.Read.All', 'Group.Read.All', 'GroupMember.ReadWrite.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All')

function Test-PcmGroupReadRole {
    <# Can this token read groups and their members? One of the Microsoft Graph permissions that read groups is enough. #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Roles = @())
    foreach ($role in $script:GroupReadRoles) { if (@($Roles | Where-Object { Test-PcmSameText $_ $role }).Count) { return $true } }
    return $false
}

function Get-PcmAdminRole {
    <# The Microsoft Entra administrator roles of the account that signed in (interactive sign-in). $null when they cannot be read. #>
    [CmdletBinding()]
    param()
    try {
        $roles = [System.Collections.Generic.List[string]]::new()
        $page = Invoke-PcmGraph -Uri 'v1.0/me/transitiveMemberOf/microsoft.graph.directoryRole?$select=displayName,roleTemplateId'
        foreach ($role in @($page.value)) { if ([string]$role.displayName) { $roles.Add([string]$role.displayName) } }
        return $roles.ToArray()
    } catch { return $null }
}

function Get-PcmRoleNeed {
    <#
    .SYNOPSIS
        The Microsoft Entra administrator roles an interactive sign-in needs for an action: each need lists the roles that give it (any one of them; the Global
        Administrator gives them all). Microsoft Purview and Exchange permissions are listed in the guide: they are not readable from Microsoft Graph.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Check', 'Convert', 'Recover')][string]$Action)
    $needs = @(
        [pscustomobject]@{ Need = 'read users, licences and Exchange Online recipients'; AnyOf = @('Global Reader', 'Global Administrator', 'Exchange Administrator', 'User Administrator') }
    )
    if ($Action -ne 'Check') {
        $needs = @(
            [pscustomobject]@{ Need = 'move the source of authority of users and groups to the cloud and back (Hybrid Identity Administrator)'; AnyOf = @('Hybrid Identity Administrator', 'Global Administrator') }
            [pscustomobject]@{ Need = 'set the usage location, assign and remove licences, delete identities (User Administrator)'; AnyOf = @('User Administrator', 'Global Administrator') }
            [pscustomobject]@{ Need = 'change mailboxes and mailbox permissions in Exchange Online (Exchange Administrator)'; AnyOf = @('Exchange Administrator', 'Global Administrator') }
            [pscustomobject]@{ Need = 'add users to the licence group (Groups Administrator or User Administrator)'; AnyOf = @('Groups Administrator', 'User Administrator', 'Global Administrator') }
        )
    }
    return $needs
}

function Test-PcmAdminRole {
    <# Compares the roles of the signed-in account with what an action needs. Level Info when everything is there or cannot be read, Warn for what seems to be missing. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Check', 'Convert', 'Recover')][string]$Action, [AllowNull()][string[]]$Roles)
    if ($null -eq $Roles) {
        return @([pscustomobject]@{ Level = 'Info'; Code = 'ADMIN_ROLE'; Message = 'The administrator roles of the account that signed in could not be read: the actions will tell what is refused.' })
    }
    $findings = [System.Collections.Generic.List[object]]::new()
    $held = @($Roles)
    foreach ($need in @(Get-PcmRoleNeed -Action $Action)) {
        $has = @($need.AnyOf | Where-Object { $role = $_; @($held | Where-Object { Test-PcmSameText $_ $role }).Count }).Count -gt 0
        if (-not $has) {
            [void]$findings.Add([pscustomobject]@{ Level = 'Warn'; Code = 'ADMIN_ROLE'; Message = ('The account that signed in holds none of these Microsoft Entra roles: {0}. They are needed to {1}.' -f ($need.AnyOf -join ', '), $need.Need) })
        }
    }
    if (-not $findings.Count) { [void]$findings.Add([pscustomobject]@{ Level = 'Ok'; Code = 'ADMIN_ROLE'; Message = ('Microsoft Entra roles of the account that signed in: {0}' -f $(if ($held.Count) { $held -join ', ' } else { 'none' })) }) }
    return $findings.ToArray()
}
#endregion

#region Follow-up of a batch ---------------------------------------------------------------------------

function Test-PcmSharedGrant {
    <# Is one permission of the snapshot there on a cloud shared mailbox? Read-only (the read-back of Grant-PcmSharedPermission). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][object]$Grant)
    switch ($Grant.Right) {
        'FullAccess' { return [bool](@(Get-MailboxPermission -Identity $Mailbox -User $Grant.Trustee -ErrorAction SilentlyContinue | Where-Object { -not $_.Deny -and @($_.AccessRights | ForEach-Object { [string]$_ }) -contains 'FullAccess' }).Count) }
        'SendAs' { return [bool](@(Get-RecipientPermission -Identity $Mailbox -Trustee $Grant.Trustee -ErrorAction SilentlyContinue | Where-Object { @($_.AccessRights | ForEach-Object { [string]$_ }) -contains 'SendAs' }).Count) }
        'SendOnBehalf' {
            $delegate = Get-Recipient -Identity $Grant.Trustee -ErrorAction SilentlyContinue
            $current = @((Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo | ForEach-Object { [string]$_ })
            return [bool]($delegate -and ($current -contains [string]$delegate.Name -or $current -contains [string]$delegate.Identity -or $current -contains [string]$delegate.ExternalDirectoryObjectId -or $current -contains [string]$delegate.Guid))
        }
    }
    return $false
}

function Test-PcmBatchExpectation {
    <#
    .SYNOPSIS
        Is one object of a batch in the state that was expected (-Expect of Check -Batch)? Pure function of the journal row, of the snapshot row, of the cloud
        state read by Get-PcmCloudStateSet and of the facts that the caller read (read-only).
    .DESCRIPTION
        Provisioned   user: cloud mailbox (UserMailbox), source of authority in the cloud, Exchange plan enabled, no longer synchronised from AD (what Convert waits for).
                      shared: SharedMailbox, tagged, source of authority in the cloud, permissions of the snapshot present. Native resource: the mailbox that the journal owns.
        Deprovisioned MailUser with the on-premises ExchangeGuid, synchronised from AD again, source of authority back in AD (what Recover waits for). Native resource: identity gone.
        Retained      what Recover keeps: a case hold of the tool on the cloud mailbox of a user, an inactive mailbox for a shared mailbox, nothing for a native resource without protection.
    .OUTPUTS
        Met, Detail (plain words), Stage (the Exchange Online state shown in the report).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Provisioned', 'Deprovisioned', 'Retained')][string]$Expect, [Parameter(Mandatory)][object]$Item, [AllowNull()][object]$Record,
        [AllowNull()][object]$State, [Parameter(Mandatory)][hashtable]$Config, [hashtable]$Facts = @{})
    $problems = [System.Collections.Generic.List[string]]::new()
    $ok = [System.Collections.Generic.List[string]]::new()
    $native = [bool](Get-PcmValue $Facts 'Native' $false)
    $isUser = Test-PcmSameText ([string]$Item.kind) 'User'
    $stage = ''
    if ($State -is [System.Management.Automation.ErrorRecord]) { return [pscustomobject]@{ Met = $false; Detail = ('the state could not be read: {0}' -f $State.Exception.Message); Stage = '' } }
    if ($native) {
        $exists = Get-PcmValue $Facts 'NativeExists' $null
        $stage = [string](Get-PcmValue $Facts 'NativeType' '')
        $mode = [string](Get-PcmValue $Facts 'NativeMode' 'None')
        switch ($Expect) {
            'Provisioned' {
                if ($exists -ne $true) { $problems.Add('the native resource is not in Exchange Online') }
                elseif ($stage -ne [string](Get-PcmValue $Facts 'CloudType' $stage)) { $problems.Add(('it is a {0}, the journal says {1}' -f $stage, (Get-PcmValue $Facts 'CloudType' ''))) }
                else { $ok.Add(('native {0} present, owned by the journal' -f $stage)) }
            }
            'Deprovisioned' { if ($exists -eq $false) { $ok.Add('the cloud identity of the native resource is gone') } else { $problems.Add('the cloud identity of the native resource is still there') } }
            'Retained' {
                if ($mode -eq 'None') { $ok.Add('native resource without protection: no hold and no inactive copy by design') }
                elseif ((Get-PcmValue $Facts 'InactiveFound' $null) -eq $true) { $ok.Add('inactive mailbox kept') } else { $problems.Add('no inactive mailbox found for this protected native resource') }
            }
        }
        return [pscustomobject]@{ Met = ($problems.Count -eq 0); Detail = (@($problems) + @($ok) -join '; '); Stage = $stage }
    }
    if (-not $State) { return [pscustomobject]@{ Met = $false; Detail = 'no Entra ID object found for this object'; Stage = '' } }
    $type = [string](Get-PcmValue $State 'RecipientType' '')
    $stage = $type
    $onPremGuid = [string](Get-PcmValue $Record 'exchange_guid' '')
    $behavior = Get-PcmValue $State 'IsCloudManaged' $null
    $user = Get-PcmValue $State 'User' $null
    $synced = [bool](Get-PcmValue $user 'onPremisesSyncEnabled' $false)
    switch ($Expect) {
        'Provisioned' {
            if ($isUser) { if ($type -eq 'UserMailbox') { $ok.Add('cloud mailbox present') } else { $problems.Add(('no cloud mailbox yet (Exchange Online shows {0})' -f $(if ($type) { $type } else { 'nothing' }))) } }
            else { if ($type -eq 'SharedMailbox') { $ok.Add('shared mailbox present') } else { $problems.Add(('not a shared mailbox yet (Exchange Online shows {0})' -f $(if ($type) { $type } else { 'nothing' }))) } }
            if ($behavior -eq $true) { $ok.Add('source of authority in the cloud') } else { $problems.Add('the source of authority is not in the cloud') }
            if ($type -match 'Mailbox$' -and [string](Get-PcmValue $State 'ExchangeGuid' '') -eq [guid]::Empty.ToString()) { $problems.Add('the mailbox has no ExchangeGuid') }
            if ($isUser) {
                $plan = Get-PcmValue $Facts 'HasPlan' $null
                if ($plan -eq $true) { $ok.Add('Exchange plan enabled') } elseif ($plan -eq $false) { $problems.Add('no Exchange plan enabled') }
            } else {
                $tag = [string]$Config.Retention.TagValue
                if ($tag -and [string](Get-PcmValue $State 'TagValue' '') -ne $tag) { $problems.Add(('{0} is not {1}' -f $Config.Retention.TagAttribute, $tag)) } elseif ($tag) { $ok.Add('retention tag set') }
                $failed = @(Get-PcmValue $Facts 'PermissionsFailed' @() | ForEach-Object { $_ } | Where-Object { $_ })
                $missing = Get-PcmValue $Facts 'PermissionsMissing' $null
                if ($failed.Count) { $problems.Add(('permissions that Convert could not grant: {0}' -f ($failed -join '; '))) }
                if ($null -ne $missing -and @($missing).Count) { $problems.Add(('permissions of the snapshot not found: {0}' -f (@($missing) -join '; '))) }
                elseif ($null -ne $missing -and -not $failed.Count) { $ok.Add('permissions of the snapshot present') }
            }
        }
        'Deprovisioned' {
            if ($type -eq 'MailUser' -and $onPremGuid -and [string](Get-PcmValue $State 'ExchangeGuid' '') -eq $onPremGuid) { $ok.Add('MailUser with the on-premises ExchangeGuid') }
            else { $problems.Add(('not a MailUser with the on-premises ExchangeGuid (Exchange Online shows {0})' -f $(if ($type) { $type } else { 'nothing' }))) }
            if ($synced) { $ok.Add('synchronised from AD') } else { $problems.Add('not synchronised from AD yet') }
            if ($behavior -eq $true) { $problems.Add('the source of authority is still in the cloud') }
        }
        'Retained' {
            if ($isUser) {
                $stamped = Get-PcmValue $Facts 'HoldStamped' $null
                if ($stamped -eq $true) { $ok.Add('case hold stamped on the cloud mailbox') } else { $problems.Add('no case hold of the tool on the cloud mailbox (Recover has not finished for this user)') }
            } else {
                $found = Get-PcmValue $Facts 'InactiveFound' $null
                if ($found -eq $true) { $ok.Add('inactive mailbox kept') }
                elseif ($found -eq $false) { $problems.Add('the inactive mailbox is not there') }
                else { $problems.Add('no inactive mailbox recorded for it (Recover has not deleted its identity yet)') }
            }
        }
    }
    return [pscustomobject]@{ Met = ($problems.Count -eq 0); Detail = ((@($problems) + @($ok)) -join '; '); Stage = $stage }
}
#endregion

#region Entra groups ----------------------------------------------------------------------------------

function Find-PcmEntraGroup {
    <#
    .SYNOPSIS
        The Microsoft Entra groups that a value names: an object ID (GUID), or a display name, an e-mail address or a mail nickname. The letter case does
        not matter: what Graph returns is checked again with the case ignored. Read-only.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value)
    $select = 'id,displayName,mail,mailNickname,onPremisesSyncEnabled,groupTypes'
    $key = $Value.Trim()
    $guid = [guid]::Empty
    if ([guid]::TryParse($key, [ref]$guid)) {
        $group = Invoke-PcmGraph -Uri ('v1.0/groups/{0}?$select={1}' -f $guid.ToString(), $select) -AllowNotFound
        if ($group) { return @($group) }
        return @()
    }
    $quoted = $key.Replace("'", "''")
    $filter = [uri]::EscapeDataString(("displayName eq '{0}' or mail eq '{0}' or mailNickname eq '{0}'" -f $quoted))
    $uri = 'v1.0/groups?$filter={0}&$select={1}&$top=999' -f $filter, $select
    $page = $null
    try { $page = Invoke-PcmGraph -Uri $uri }
    catch { $page = Invoke-PcmGraph -Uri ($uri + '&$count=true') -Headers @{ ConsistencyLevel = 'eventual' } }
    $found = [System.Collections.Generic.List[object]]::new()
    $first = [uri]('https://graph.microsoft.com/' + $uri)
    $guard = 0
    while ($page -and $guard -lt 50) {
        $guard++
        foreach ($group in @($page.value)) {
            if ((Test-PcmSameText ([string]$group.displayName) $key) -or (Test-PcmSameText ([string]$group.mail) $key) -or (Test-PcmSameText ([string]$group.mailNickname) $key)) { $found.Add($group) }
        }
        $next = [string](Get-PcmValue $page '@odata.nextLink' '')
        if (-not $next) { break }
        $parsed = [uri]$next
        if (-not (Test-PcmSameText $parsed.Host $first.Host) -or -not (Test-PcmSameText $parsed.AbsolutePath $first.AbsolutePath)) { throw 'Microsoft Graph returned a next page for another address than the one asked: the group search stops.' }
        $page = Invoke-PcmGraph -Uri $next
    }
    return $found.ToArray()
}

function Get-PcmEntraGroupMember {
    <#
    .SYNOPSIS
        Every member of a Microsoft Entra group, nested groups included (transitiveMembers, all the pages): users, devices, contacts and the nested groups
        themselves, each with its class. Read-only.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GroupId)
    $select = 'id,userPrincipalName,mail,onPremisesImmutableId,displayName'
    $uri = 'v1.0/groups/{0}/transitiveMembers?$select={1}&$top=999' -f ([guid]$GroupId).ToString(), $select
    $first = [uri]('https://graph.microsoft.com/' + $uri)
    $members = [System.Collections.Generic.List[object]]::new()
    $page = Invoke-PcmGraph -Uri $uri
    $guard = 0
    while ($page) {
        $guard++
        if ($guard -gt 5000) { throw 'The member list of the group has more pages than expected: the read stops.' }
        foreach ($item in @($page.value)) {
            $type = [string](Get-PcmValue $item '@odata.type' '')
            $class = switch -Regex ($type) { 'user$' { 'User' } 'group$' { 'Group' } 'device$' { 'Device' } 'orgContact$' { 'Contact' } 'servicePrincipal$' { 'ServicePrincipal' } default { $(if ($type) { ($type -split '\.')[-1] } else { 'User' }) } }
            $members.Add([pscustomobject]@{ Id = [string]$item.id; Class = $class; Name = [string](Get-PcmValue $item 'displayName' ''); Upn = [string](Get-PcmValue $item 'userPrincipalName' '')
                    Smtp = [string](Get-PcmValue $item 'mail' ''); ImmutableId = [string](Get-PcmValue $item 'onPremisesImmutableId' '') })
        }
        $next = [string](Get-PcmValue $page '@odata.nextLink' '')
        if (-not $next) { break }
        $parsed = [uri]$next
        if (-not (Test-PcmSameText $parsed.Host $first.Host) -or -not (Test-PcmSameText $parsed.AbsolutePath $first.AbsolutePath)) { throw 'Microsoft Graph returned a next page for another address than the one asked: the member read stops.' }
        $page = Invoke-PcmGraph -Uri $next
    }
    return $members.ToArray()
}
#endregion

#region Tenant facts ---------------------------------------------------------------------------------

function Get-PcmTenantFact {
    <#
    .SYNOPSIS
        Licences (subscribedSkus), directory synchronisation status, organisation-wide Exchange holds and, in Group
        mode, the licence group.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [switch]$SkipUserGroup, [switch]$NativeOnly)
    if ($NativeOnly) { return [pscustomobject]@{ Skus = @(); Organization = $null; LicenceGroup = $null; LicenceGroupCloudManaged = $null; OrgHolds = @() } }
    $skus = @((Invoke-PcmGraph -Uri 'v1.0/subscribedSkus').value)
    $organization = @((Invoke-PcmGraph -Uri 'v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime').value) | Select-Object -First 1
    $group = $null; $groupBehavior = $null
    $users = $Context.Config.Licensing.Users
    if (-not $SkipUserGroup -and $users.Mode -eq 'Group' -and $users.GroupId) {
        $group = Invoke-PcmGraph -Uri ("v1.0/groups/{0}?`$select=id,displayName,assignedLicenses,onPremisesSyncEnabled,securityEnabled,mailEnabled" -f $users.GroupId) -AllowNotFound
        if ($group) { $groupBehavior = Invoke-PcmGraph -Uri ("v1.0/groups/{0}/onPremisesSyncBehavior?`$select=isCloudManaged" -f $users.GroupId) -AllowNotFound }
    }
    # Retention policies applied to every mailbox (InPlaceHolds of the organisation, prefix mbx).
    $orgHolds = @()
    try { $orgHolds = @((Get-OrganizationConfig -ErrorAction Stop).InPlaceHolds | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^mbx' }) } catch { $orgHolds = @() }
    return [pscustomobject]@{ Skus = $skus; Organization = $organization; LicenceGroup = $group; LicenceGroupCloudManaged = $(if ($groupBehavior) { [bool]$groupBehavior.isCloudManaged } else { $null })
        OrgHolds = $orgHolds }
}

function Get-PcmMailboxPlanId {
    <# servicePlanId of every Exchange plan that gives a mailbox, in the given SKUs. #>
    param([AllowEmptyCollection()][object[]]$Skus = @())
    return @($Skus | ForEach-Object { $_.servicePlans } | Where-Object { $_ -and [string]$_.servicePlanName -match $script:ExchangePlanPattern } | ForEach-Object { [string]$_.servicePlanId } | Get-PcmDistinctText)
}

function Get-PcmSkuExchangePlan {
    <# Exchange service plans of a SKU (servicePlanName EXCHANGE_S_*), minus the disabled ones. #>
    param([Parameter(Mandatory)][object]$Sku, [string[]]$DisabledPlanIds = @())
    return @($Sku.servicePlans | Where-Object { [string]$_.servicePlanName -match $script:ExchangePlanPattern -and [string]$_.servicePlanId -notin $DisabledPlanIds } |
        ForEach-Object { [string]$_.servicePlanName })
}

function Get-PcmFreeUnit {
    param([Parameter(Mandatory)][object]$Sku)
    return [int]$Sku.prepaidUnits.enabled - [int]$Sku.consumedUnits
}

function Get-PcmSkuSpelling {
    <# A licence name as the tenant spells it (SkuPartNumber), whatever the letter case typed in the configuration; the typed name when the tenant has no such licence. #>
    [CmdletBinding()]
    param([AllowEmptyString()][string]$Name, [AllowNull()][object]$Tenant)
    $found = @(@($Tenant.Skus) | Where-Object { $_ -and [string]$_.skuPartNumber -eq $Name }) | Select-Object -First 1
    if ($found) { return [string]$found.skuPartNumber }
    return $Name
}

function Get-PcmSharedUnitCount {
    <#
    .SYNOPSIS
        The units of the shared mailbox licence (Licensing.Shared.SkuPartNumber) at the start of Convert: Free, UserDraw (the units the users of the same run take from that same
        licence, through the licence group or directly), KeepFree and Wave (how many shared mailboxes can hold a temporary unit at the same time: Licensing.Shared.Parallel,
        and never more than the units left after the users and KeepFree). $null when the licence is not one of the tenant. Pure function (no cloud call).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][object]$Tenant, [int]$UserCount = 0)
    $licensing = $Config.Licensing
    $name = [string](Get-PcmValue (Get-PcmValue $licensing 'Shared' @{}) 'SkuPartNumber' '')
    $sharedSku = @($Tenant.Skus | Where-Object { [string]$_.skuPartNumber -eq $name }) | Select-Object -First 1
    if (-not $name -or -not $sharedSku) { return $null }
    $sharedId = [string]$sharedSku.skuId
    $draw = 0
    if ($UserCount -gt 0) {
        $users = Get-PcmValue $licensing 'Users' @{}
        if (Test-PcmSameText ([string](Get-PcmValue $users 'Mode' 'Group')) 'Direct') {
            if ([string](Get-PcmValue $users 'SkuPartNumber' '') -eq $name) { $draw = $UserCount }
        }
        elseif ((Test-PcmSameText ([string](Get-PcmValue $users 'Mode' 'Group')) 'Group') -and $Tenant.PSObject.Properties['LicenceGroup'] -and $Tenant.LicenceGroup) {
            foreach ($assigned in @($Tenant.LicenceGroup.assignedLicenses)) {
                if ([string]$assigned.skuId -ne $sharedId) { continue }
                if (@(Get-PcmSkuExchangePlan -Sku $sharedSku -DisabledPlanIds @($assigned.disabledPlans | ForEach-Object { [string]$_ })).Count) { $draw = $UserCount }
            }
        }
    }
    $free = Get-PcmFreeUnit $sharedSku
    $keepFree = [Math]::Max(0, [int](Get-PcmValue (Get-PcmValue $licensing 'Shared' @{}) 'KeepFree' 0))
    $parallel = [Math]::Max(1, [int](Get-PcmValue (Get-PcmValue $licensing 'Shared' @{}) 'Parallel' 100))
    return [pscustomobject]@{ Name = [string]$sharedSku.skuPartNumber; Free = $free; UserDraw = $draw; KeepFree = $keepFree; Parallel = $parallel; Wave = [Math]::Max(0, [Math]::Min($parallel, ($free - $draw - $keepFree))) }
}

function Test-PcmTenantReadiness {
    <#
    .SYNOPSIS
        Tenant-level rules: app permissions, licences for the users (Group, Kiosk or Direct) and for the
        temporary licence of the shared mailboxes. Pure function (no cloud call).
    .PARAMETER Roles
        Application permissions of the token.
    .PARAMETER Tenant
        Output of Get-PcmTenantFact.
    .PARAMETER UserCount
        Users of the snapshot that will need an Exchange plan.
    .PARAMETER SharedCount
        Shared, room and equipment mailboxes of the snapshot.
    .OUTPUTS
        Findings: Level (Ok, Info, Warn, Error), Code, Message.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [AllowEmptyCollection()][string[]]$Roles = @(), [Parameter(Mandatory)][object]$Tenant,
        [int]$UserCount = 0, [int]$SharedCount = 0, [int]$ResourceCount = 0, [switch]$SkipUserLicensing, [switch]$NativeOnly, [ValidateSet('Certificate', 'Interactive')][string]$SignIn = 'Certificate')
    $findings = [System.Collections.Generic.List[object]]::new()
    $add = { param($Level, $Code, $Message) [void]$findings.Add([pscustomobject]@{ Level = $Level; Code = $Code; Message = $Message }) }
    $licensing = $Config.Licensing
    $needed = @($script:RequiredGraphRoles)
    if ($NativeOnly) { $needed = @($script:RequiredGraphRoles | Where-Object Role -eq 'User.ReadWrite.All') }
    if ((-not $SkipUserLicensing -and $licensing.Users.Mode -eq 'Group') -or $ResourceCount -gt 0) { $needed += $script:GroupModeRoles[0] }
    $missing = @($needed | Where-Object { $_.Role -notin $Roles -and -not ($_.Role -eq 'LicenseAssignment.ReadWrite.All' -and 'Directory.ReadWrite.All' -in $Roles) })
    $kindText = if ($SignIn -eq 'Interactive') { 'delegated permission' } else { 'application permission' }
    $fixText = if ($SignIn -eq 'Interactive') { ' An administrator of the tenant must consent to them at the sign-in.' } else { '' }
    if ($missing.Count) { & $add 'Error' 'APP_PERMISSION' ("Missing Graph $kindText(s): " + (($missing | ForEach-Object { '{0} ({1})' -f $_.Role, $_.Why }) -join '; ') + '.' + $fixText) }
    else { & $add 'Ok' 'APP_PERMISSION' ("Graph ${kindText}s present: " + (($needed | ForEach-Object { $_.Role }) -join ', ')) }

    $skuById = @{}; $skuByName = @{}
    foreach ($sku in @($Tenant.Skus)) { $skuById[[string]$sku.skuId] = $sku; $skuByName[[string]$sku.skuPartNumber] = $sku }
    $keepFree = [Math]::Max(0, [int](Get-PcmValue $licensing.Shared 'KeepFree' 0))
    $sharedName = [string]$licensing.Shared.SkuPartNumber
    $sharedSkuId = if ($sharedName -and $skuByName.ContainsKey($sharedName)) { [string]$skuByName[$sharedName].skuId } else { '' }
    if (-not $SkipUserLicensing) { switch ($licensing.Users.Mode) {
        'Group' {
            if (-not $licensing.Users.GroupId) { & $add 'Error' 'LICENCE_GROUP' 'Licensing.Users.GroupId is empty (object ID of the group that carries the Exchange plan).' }
            elseif (-not $Tenant.LicenceGroup) { & $add 'Error' 'LICENCE_GROUP' "Licence group $($licensing.Users.GroupId) not found in Entra ID." }
            else {
                $group = $Tenant.LicenceGroup
                $plans = @(); $free = $null; $names = @(); $drawsShared = $false
                foreach ($assigned in @($group.assignedLicenses)) {
                    $sku = $skuById[[string]$assigned.skuId]
                    if (-not $sku) { continue }
                    $exchange = @(Get-PcmSkuExchangePlan -Sku $sku -DisabledPlanIds @($assigned.disabledPlans | ForEach-Object { [string]$_ }))
                    if ($exchange.Count) { $plans += $exchange; $free = if ($null -eq $free) { Get-PcmFreeUnit $sku } else { [Math]::Min($free, (Get-PcmFreeUnit $sku)) }; $names += [string]$sku.skuPartNumber; if ($sharedSkuId -and [string]$sku.skuId -eq $sharedSkuId) { $drawsShared = $true } }
                }
                if (-not $plans.Count) { & $add 'Error' 'LICENCE_GROUP' "Group $($group.displayName) assigns no licence with an Exchange plan enabled." }
                else {
                    $reserved = if ($drawsShared -and $UserCount -gt 0) { $keepFree } else { 0 }
                    $level = if ($free -ge ($UserCount + $reserved)) { 'Ok' } else { 'Error' }
                    $reserveText = if ($reserved) { ', {0} of them kept free (Licensing.Shared.KeepFree, same licence as the shared mailboxes)' -f $reserved } else { '' }
                    & $add $level 'LICENCE_CAPACITY' ('Group {0}: {1} ({2}), {3} free unit(s) for {4} user(s){5}.' -f $group.displayName, ($names -join ', '), (($plans | Get-PcmDistinctText) -join ', '), $free, $UserCount, $reserveText)
                }
                if ($group.onPremisesSyncEnabled -and -not $Tenant.LicenceGroupCloudManaged) {
                    & $add 'Info' 'LICENCE_GROUP_SOA' ("Group $($group.displayName) is synchronised from AD: Convert transfers its source of authority to the cloud first (Group-OnPremisesSyncBehavior.ReadWrite.All), Recover gives it back to AD last.")
                    if ('Group-OnPremisesSyncBehavior.ReadWrite.All' -notin $Roles) { & $add 'Error' 'APP_PERMISSION' "Missing Graph $kindText Group-OnPremisesSyncBehavior.ReadWrite.All (licence group synchronised from AD)." }
                }
            }
        }
        'Direct' {
            $sku = $skuByName[[string]$licensing.Users.SkuPartNumber]
            if (-not $sku) { & $add 'Error' 'LICENCE_SKU' "Licensing.Users.SkuPartNumber $($licensing.Users.SkuPartNumber) is not a licence of the tenant." }
            else {
                $free = Get-PcmFreeUnit $sku
                $reserved = if ($sharedSkuId -and [string]$sku.skuId -eq $sharedSkuId -and $UserCount -gt 0) { $keepFree } else { 0 }
                $reserveText = if ($reserved) { ', {0} of them kept free (Licensing.Shared.KeepFree, same licence as the shared mailboxes)' -f $reserved } else { '' }
                & $add $(if ($free -ge ($UserCount + $reserved)) { 'Ok' } else { 'Error' }) 'LICENCE_CAPACITY' ('{0}: {1} free unit(s) for {2} user(s){3}.' -f $sku.skuPartNumber, $free, $UserCount, $reserveText)
            }
        }
        'Kiosk' { & $add 'Info' 'LICENCE_KIOSK' 'Kiosk mode: each user gets the Exchange Kiosk plan of the licence he already has (checked per user). No extra licence unit; 2 GB mailbox, no Outlook desktop.' }
    } }
    if ($SharedCount -gt 0) {
        $name = [string]$licensing.Shared.SkuPartNumber
        if (-not $name) { & $add 'Error' 'SHARED_LICENCE' 'Licensing.Shared.SkuPartNumber is empty: Convert needs a temporary licence with an Exchange plan for each shared mailbox (given back once it is converted).' }
        else {
            $sku = $skuByName[$name]
            if (-not $sku) { & $add 'Error' 'SHARED_LICENCE' "Licensing.Shared.SkuPartNumber $name is not a licence of the tenant." }
            elseif (-not @(Get-PcmSkuExchangePlan -Sku $sku).Count) { & $add 'Error' 'SHARED_LICENCE' "$name has no Exchange plan." }
            else {
                $free = Get-PcmFreeUnit $sku
                $usable = $free - $keepFree
                $parallel = [Math]::Max(1, [int](Get-PcmValue $licensing.Shared 'Parallel' 100))
                # The users of the same run take their units first, from this same licence when the licence group (or the direct licence) is this one.
                $units = Get-PcmSharedUnitCount -Config $Config -Tenant $Tenant -UserCount $(if ($SkipUserLicensing) { 0 } else { $UserCount })
                $draw = if ($units) { [int]$units.UserDraw } else { 0 }
                $name = [string]$sku.skuPartNumber
                $freeText = '{0} free unit(s)' -f $free
                if ($draw -gt 0) { $freeText += ', {0} taken by the users of this run' -f $draw }
                if ($keepFree -gt 0) { $freeText += ', {0} kept free (Licensing.Shared.KeepFree)' -f $keepFree }
                if ($usable -ge 1 -and ($usable - $draw) -lt 1) {
                    & $add 'Warn' 'SHARED_LICENCE' ('{0}: {1}; the users of this run take every unit that can be used, so no shared mailbox can hold a temporary unit: they are not started until a unit is free ({2} in scope). Free a unit.' -f $name, $freeText, $SharedCount)
                }
                elseif ($usable -ge 1) {
                    $wave = [Math]::Min($parallel, $usable - $draw)
                    & $add 'Ok' 'SHARED_LICENCE' ('{0}: {1}; the shared mailboxes are converted in waves of {2} (Licensing.Shared.Parallel = {3}, one temporary unit each, given back after each wave; {4} in scope).' -f $name, $freeText, $wave, $parallel, $SharedCount)
                } else {
                    $fix = if ($keepFree -gt 0) { 'Free a unit, or lower Licensing.Shared.KeepFree' } else { 'Free a unit' }
                    & $add 'Error' 'SHARED_LICENCE' ('{0}: {1}; no unit is left for the temporary licence of the {2} shared mailbox(es) in scope. {3}.' -f $name, $freeText, $SharedCount, $fix)
                }
            }
        }
    }
    $organization = $Tenant.Organization
    if ($organization -and $organization.onPremisesSyncEnabled -and $organization.onPremisesLastSyncDateTime) {
        $last = [datetime]$organization.onPremisesLastSyncDateTime
        & $add 'Info' 'DIRECTORY_SYNC' ('Last Entra Connect synchronisation: {0} UTC.' -f $last.ToUniversalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture))
    }
    $orgHolds = @(if ($Tenant.PSObject.Properties['OrgHolds']) { $Tenant.OrgHolds | Where-Object { $_ } })
    if ($orgHolds.Count) {
        & $add 'Warn' 'ORG_HOLD' ('Retention policies applied to every mailbox ({0}): they will also hold the cloud mailboxes of the users, and a hold blocks their Recover. Exclude the users in scope from these policies, or plan to lift them before Recover.' -f ($orgHolds -join ', '))
    }
    return $findings.ToArray()
}
#endregion

#region Objects --------------------------------------------------------------------------------------

function Get-PcmCloudState {
    <#
    .SYNOPSIS
        Entra ID object and Exchange Online recipient of one snapshot row (read-only). Get-PcmCloudStateSet for one row.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Record)
    $state = (Get-PcmCloudStateSet -Context $Context -Records @($Record))[[string]$Record.object_guid]
    if ($state -is [System.Management.Automation.ErrorRecord]) { throw $state }
    return $state
}
#endregion

#region Many objects at once ------------------------------------------------------------------------
# A state read one object at a time costs five calls (about 3 s): 15 hours for 21,000 objects. These functions read
# the same states in bulk (typical times for 21,700 users and 16,000 mailboxes): every user page by page
# (20 s), the source of authority by Graph batches of 20 (0.4 s each), Exchange Online by filters of 50 object IDs
# (0.3 to 0.7 s each).

$script:ExoFilterSize = 50
$script:GraphBatchSize = 20

function ConvertFrom-PcmMailboxLocation {
    <# MailboxLocations of Get-MailUser / Get-EXOMailbox ('1;<guid>;<type>;<database>;<id>') as Type and Guid, like Get-MailboxLocation. #>
    param([AllowNull()][AllowEmptyCollection()][object[]]$Value = @())
    return @(foreach ($text in @($Value | Where-Object { $_ })) {
            $parts = ([string]$text).Split(';')
            if ($parts.Count -ge 3 -and $parts[2]) { [pscustomobject]@{ Type = $parts[2]; Guid = $parts[1] } }
        })
}

function Get-PcmExoFilter {
    <# An OPATH filter that matches any of the object IDs. #>
    param([Parameter(Mandatory)][string[]]$Id)
    return (@($Id | ForEach-Object { "ExternalDirectoryObjectId -eq '{0}'" -f $_.Replace("'", "''") }) -join ' -or ')
}

function Get-PcmExoStateSet {
    <#
    .SYNOPSIS
        Get-PcmExoState of many Entra object IDs, with their mailbox locations (Teams storage): filters of 50 IDs on
        Get-EXORecipient, Get-User, Get-MailUser and Get-EXOMailbox. A chunk that fails is read again one object at a
        time. Returns a hashtable ID -> state (Type '' when there is no recipient).
    #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Id = @(), [string]$TagAttribute = 'CustomAttribute1')
    $result = @{}
    $ids = @($Id | Where-Object { $_ } | Get-PcmDistinctText)
    for ($start = 0; $start -lt $ids.Count; $start += $script:ExoFilterSize) {
        $chunk = [string[]]@($ids[$start..([Math]::Min($ids.Count, $start + $script:ExoFilterSize) - 1)])
        try {
            $filter = Get-PcmExoFilter -Id $chunk
            $recipients = @{}; $users = @{}; $mailUsers = @{}; $mailboxes = @{}
            foreach ($r in @(Get-EXORecipient -Filter $filter -ResultSize Unlimited -Properties ExternalDirectoryObjectId, ExchangeGuid, $TagAttribute -ErrorAction Stop)) { $recipients[[string]$r.ExternalDirectoryObjectId] = $r }
            foreach ($u in @(Get-User -Filter $filter -ResultSize Unlimited -ErrorAction Stop)) { $users[[string]$u.ExternalDirectoryObjectId] = $u }
            $mailUserIds = [string[]]@($recipients.Keys | Where-Object { [string]$recipients[$_].RecipientTypeDetails -eq 'MailUser' })
            $mailboxIds = [string[]]@($recipients.Keys | Where-Object { [string]$recipients[$_].RecipientTypeDetails -match 'Mailbox$' })
            if ($mailUserIds.Count) { foreach ($m in @(Get-MailUser -Filter (Get-PcmExoFilter -Id $mailUserIds) -ResultSize Unlimited -ErrorAction Stop)) { $mailUsers[[string]$m.ExternalDirectoryObjectId] = $m } }
            if ($mailboxIds.Count) {
                foreach ($m in @(Get-EXOMailbox -Filter (Get-PcmExoFilter -Id $mailboxIds) -ResultSize Unlimited -Properties ExternalDirectoryObjectId, InPlaceHolds, LitigationHoldEnabled, MailboxLocations -ErrorAction Stop)) { $mailboxes[[string]$m.ExternalDirectoryObjectId] = $m }
            }
            foreach ($one in $chunk) {
                $recipient = $recipients[$one]; $user = $users[$one]; $holds = @(); $locations = @(); $guid = ''
                if ($recipient) {
                    $guid = [string]$recipient.ExchangeGuid
                    if ($mailUsers.ContainsKey($one)) {
                        $holds = @($mailUsers[$one].InPlaceHolds | ForEach-Object { [string]$_ }); $guid = [string]$mailUsers[$one].ExchangeGuid
                        $locations = @(ConvertFrom-PcmMailboxLocation @($mailUsers[$one].MailboxLocations))
                    } elseif ($mailboxes.ContainsKey($one)) {
                        $mailbox = $mailboxes[$one]
                        $holds = @($mailbox.InPlaceHolds | ForEach-Object { [string]$_ }); if ($mailbox.LitigationHoldEnabled) { $holds += 'LitigationHold' }
                        $locations = @(ConvertFrom-PcmMailboxLocation @($mailbox.MailboxLocations))
                    }
                }
                $result[$one] = [pscustomobject]@{
                    Type = $(if ($recipient) { [string]$recipient.RecipientTypeDetails } else { '' }); ExchangeGuid = $guid
                    IsDirSynced = $(if ($user) { [bool]$user.IsDirSynced } else { $null }); Holds = $holds
                    Tag = $(if ($recipient -and $recipient.PSObject.Properties[$TagAttribute]) { [string]$recipient.$TagAttribute } else { '' })
                    UserType = $(if ($user) { [string]$user.RecipientTypeDetails } else { '' }); Locations = $locations
                }
            }
        } catch {
            Write-Verbose ("Exchange Online filter of {0} object(s) failed ({1}): one object at a time." -f $chunk.Count, $_.Exception.Message)
            foreach ($one in $chunk) {
                $state = Get-PcmExoState -Identity $one -TagAttribute $TagAttribute
                $locations = @(Get-MailboxLocation -User $one -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ForEach-Object { [pscustomobject]@{ Type = [string]$_.MailboxLocationType; Guid = [string]$_.MailboxGuid } })
                $state | Add-Member -NotePropertyName Locations -NotePropertyValue $locations
                $result[$one] = $state
            }
        }
    }
    return $result
}

function Get-PcmRetryAfterSecond {
    <# The Retry-After (seconds) that Microsoft Graph gave in the answer of one request of a batch: 0 when there is none, never more than two minutes. #>
    [CmdletBinding()]
    [OutputType([int])]
    param([AllowNull()][object]$Response)
    $headers = Get-PcmValue $Response 'headers' $null
    if ($null -eq $headers) { return 0 }
    $names = if ($headers -is [System.Collections.IDictionary]) { @($headers.Keys) } else { @($headers.PSObject.Properties.Name) }
    foreach ($name in $names) {
        if (-not [string]::Equals([string]$name, 'Retry-After', [StringComparison]::OrdinalIgnoreCase)) { continue }
        $seconds = 0
        if ([int]::TryParse([string](Get-PcmValue $headers ([string]$name) ''), [ref]$seconds) -and $seconds -gt 0) { return [Math]::Min($seconds, 120) }
    }
    return 0
}

function Invoke-PcmGraphBatch {
    <#
    .SYNOPSIS
        Many Microsoft Graph requests in batches of 20 (JSON batching). Throttled (429) or unavailable (5xx) requests
        are sent again (three rounds), then one by one. Every other answer is returned as it is.
    .PARAMETER Requests
        Hashtables: Method (GET, PATCH, POST, DELETE), Url (relative to v1.0, e.g. /users/<id>), Body (optional).
    .OUTPUTS
        One result per request, in the same order: Ok (2xx), Status, Body, Error ('code: message').
    #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Requests = @())
    $results = New-Object object[] $Requests.Count
    $pending = [System.Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $Requests.Count; $i++) { $pending.Add($i) }
    # Every change is approved before the first request is sent: a refusal sends nothing.
    foreach ($request in $Requests) {
        if (Test-PcmGraphWrite -Method ([string]$request.Method) -Uri ([string]$request.Url)) {
            Confirm-PcmWrite -Target (([string]$request.Url -replace '\?.*$', '').TrimStart('/')) -Operation ('Microsoft Graph ' + [string]$request.Method)
        }
    }
    $toResult = {
        param([int]$Status, [object]$Body)
        $text = ''
        if ($Status -ge 300 -and $Body) { $graphError = Get-PcmValue $Body 'error' $null; if ($graphError) { $text = '{0}: {1}' -f (Get-PcmValue $graphError 'code' ''), (Get-PcmValue $graphError 'message' '') } }
        [pscustomobject]@{ Ok = ($Status -ge 200 -and $Status -lt 300); Status = $Status; Body = $Body; Error = $text }
    }
    for ($round = 1; $round -le 3 -and $pending.Count; $round++) {
        $again = [System.Collections.Generic.List[int]]::new()
        $retryAfter = 0
        for ($start = 0; $start -lt $pending.Count; $start += $script:GraphBatchSize) {
            $chunk = @($pending.GetRange($start, [Math]::Min($script:GraphBatchSize, $pending.Count - $start)))
            $batch = @(foreach ($index in $chunk) {
                    $request = $Requests[$index]
                    $item = @{ id = [string]$index; method = [string]$request.Method; url = [string]$request.Url }
                    if ($request.ContainsKey('Body') -and $null -ne $request.Body) { $item.body = $request.Body; $item.headers = @{ 'Content-Type' = 'application/json' } }
                    $item
                })
            $answer = Invoke-PcmGraph -Method POST -Uri 'v1.0/$batch' -Body @{ requests = $batch } -Approved
            $seen = @{}
            foreach ($response in @($answer.responses)) {
                $index = [int]$response.id; $seen[$index] = $true
                $status = [int]$response.status
                if ($status -eq 429 -or $status -ge 500) { $again.Add($index); $retryAfter = [Math]::Max($retryAfter, (Get-PcmRetryAfterSecond -Response $response)) } else { $results[$index] = & $toResult $status (Get-PcmValue $response 'body' $null) }
            }
            foreach ($index in $chunk) { if (-not $seen.ContainsKey($index)) { $again.Add($index) } }
        }
        $pending = $again
        if ($pending.Count -and $round -lt 3) { Start-Sleep -Seconds ([Math]::Max((5 * $round), $retryAfter)) }
    }
    foreach ($index in $pending) {
        $request = $Requests[$index]
        $parameters = @{ Method = [string]$request.Method; Uri = ('v1.0' + [string]$request.Url); Approved = $true }
        if ($request.ContainsKey('Body') -and $null -ne $request.Body) { $parameters.Body = $request.Body }
        try { $results[$index] = & $toResult 200 (Invoke-PcmGraph @parameters) }
        catch {
            $code = 0; try { $code = [int]$_.Exception.InnerException.Response.StatusCode } catch { $code = 0 }
            if (-not $code) { try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 } }
            $results[$index] = [pscustomobject]@{ Ok = $false; Status = $code; Body = $null; Error = ($_.Exception.Message -replace '\s+', ' ') }
        }
    }
    return , $results
}

function Get-PcmSyncBehaviorSet {
    <#
    .SYNOPSIS
        isCloudManaged of many users (source of authority), by Graph batches. Returns a hashtable user ID -> $true,
        $false or $null (none). A request that keeps failing stops the read.
    #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$UserId = @())
    $result = @{}
    $ids = @($UserId | Where-Object { $_ } | Get-PcmDistinctText)
    $answers = Invoke-PcmGraphBatch -Requests @($ids | ForEach-Object { @{ Method = 'GET'; Url = ('/users/{0}/onPremisesSyncBehavior?$select=isCloudManaged' -f $_) } })
    for ($i = 0; $i -lt $ids.Count; $i++) {
        $answer = $answers[$i]
        if ($answer.Ok) { $result[$ids[$i]] = [bool](Get-PcmValue $answer.Body 'isCloudManaged' $false) }
        elseif ($answer.Status -eq 404) { $result[$ids[$i]] = $null }
        else { throw ('Source of authority of {0} not readable ({1}): {2}' -f $ids[$i], $answer.Status, $answer.Error) }
    }
    return $result
}

function Get-PcmGraphUserSet {
    <#
    .SYNOPSIS
        The Entra ID user of many snapshot rows: by onPremisesImmutableId (one match only), otherwise by UPN. Up to
        PageThreshold rows: one query per row; more: every user of the tenant, page by page. Returns a hashtable
        object_guid -> User and MatchedBy (immutableId, UPN or '' when none).
    #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Records = @(), [int]$PageThreshold = 100)
    $select = 'id,userPrincipalName,accountEnabled,onPremisesSyncEnabled,onPremisesImmutableId,usageLocation,assignedLicenses,assignedPlans,licenseAssignmentStates,serviceProvisioningErrors'
    $result = @{}
    if ($Records.Count -le $PageThreshold) {
        foreach ($record in $Records) {
            $user = $null; $matchedBy = ''
            if ($record.immutable_id) {
                $filter = [uri]::EscapeDataString(("onPremisesImmutableId eq '{0}'" -f ([string]$record.immutable_id).Replace("'", "''")))
                $found = @((Invoke-PcmGraph -Uri ("v1.0/users?`$filter={0}&`$select={1}" -f $filter, $select)).value)
                if ($found.Count -eq 1) { $user = $found[0]; $matchedBy = 'immutableId' }
            }
            if (-not $user -and $record.user_principal_name) {
                $user = Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select={1}" -f [uri]::EscapeDataString([string]$record.user_principal_name), $select) -AllowNotFound
                if ($user) { $matchedBy = 'UPN' }
            }
            $result[[string]$record.object_guid] = [pscustomobject]@{ User = $user; MatchedBy = $matchedBy }
        }
        return $result
    }
    # onPremisesImmutableId is base64: case-sensitive.
    $byImmutable = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    $byUpn = @{}
    $uri = 'v1.0/users?$select={0}&$top=999' -f $select
    while ($uri) {
        $page = Invoke-PcmGraph -Uri $uri
        foreach ($user in @($page.value)) {
            if ($user.onPremisesImmutableId) {
                $key = [string]$user.onPremisesImmutableId
                if ($byImmutable.ContainsKey($key)) { $byImmutable[$key] = $null } else { $byImmutable[$key] = $user }
            }
            if ($user.userPrincipalName) { $byUpn[[string]$user.userPrincipalName] = $user }
        }
        $uri = [string](Get-PcmValue $page '@odata.nextLink' '')
    }
    foreach ($record in $Records) {
        $user = $null; $matchedBy = ''
        # Several users with the same immutableId count as none, as the query by row does.
        if ($record.immutable_id -and $byImmutable.ContainsKey([string]$record.immutable_id) -and $byImmutable[[string]$record.immutable_id]) { $user = $byImmutable[[string]$record.immutable_id]; $matchedBy = 'immutableId' }
        elseif ($record.user_principal_name -and $byUpn.ContainsKey([string]$record.user_principal_name)) { $user = $byUpn[[string]$record.user_principal_name]; $matchedBy = 'UPN' }
        $result[[string]$record.object_guid] = [pscustomobject]@{ User = $user; MatchedBy = $matchedBy }
    }
    return $result
}

function Get-PcmCloudStateSet {
    <#
    .SYNOPSIS
        Get-PcmCloudState of many snapshot rows: Entra ID users, source of authority and Exchange Online in bulk.
        Returns a hashtable object_guid -> state, or the ErrorRecord of a row that could not be read.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [AllowEmptyCollection()][object[]]$Records = @(), [int]$PageThreshold = 100)
    $result = @{}
    if (-not $Records.Count) { return $result }
    $native = @($Records | Where-Object {
            (Test-PcmNativeResource -Config $Context.Config -Record $_) -or
            ($Context.ContainsKey('ResourceJournal') -and (Get-PcmResourceIntent -Item $Context.ResourceJournal[[string]$_.object_guid]))
        })
    if ($native.Count) {
        try {
            $inventory = Get-PcmResourceInventory -Records $native
            foreach ($record in $native) {
                $intent = if ($Context.ContainsKey('ResourceJournal')) { Get-PcmResourceIntent -Item $Context.ResourceJournal[[string]$record.object_guid] } else { $null }
                try { $result[[string]$record.object_guid] = Get-PcmResourceState -Record $record -Records $Records -Inventory $inventory -Intent $intent }
                catch { $result[[string]$record.object_guid] = $_ }
            }
        } catch { foreach ($record in $native) { $result[[string]$record.object_guid] = $_ } }
    }
    $legacy = @($Records | Where-Object { $_ -notin $native })
    if (-not $legacy.Count) { return $result }
    $users = Get-PcmGraphUserSet -Records $legacy -PageThreshold $PageThreshold
    $ids = [string[]]@($users.Values | Where-Object { $_.User } | ForEach-Object { [string]$_.User.id })
    $behaviors = Get-PcmSyncBehaviorSet -UserId $ids
    $exo = Get-PcmExoStateSet -Id $ids -TagAttribute ([string]$Context.Config.Retention.TagAttribute)
    foreach ($record in $legacy) {
        $guid = [string]$record.object_guid
        try {
            $match = $users[$guid]
            $user = $match.User
            $state = if ($user) { $exo[[string]$user.id] } else { $null }
            $behavior = if ($user -and $behaviors.ContainsKey([string]$user.id)) { $behaviors[[string]$user.id] } else { $null }
            $result[$guid] = [pscustomobject]@{
                MatchedBy = $match.MatchedBy; User = $user; IsCloudManaged = $behavior
                RecipientType = $(if ($state) { $state.Type } else { '' })
                ExchangeGuid = $(if ($state) { $state.ExchangeGuid } else { '' })
                Holds = @(if ($state) { $state.Holds }); Locations = @(if ($state) { $state.Locations }); TagValue = $(if ($state) { $state.Tag } else { '' })
            }
        } catch { $result[$guid] = $_ }
    }
    return $result
}
#endregion

#region Native, unsynchronised resource mailboxes --------------------------------------------------

function Test-PcmNativeResource {
    param([hashtable]$Config, [object]$Record)
    return ([string]$Record.kind -in @('Room', 'Equipment') -and [string](Get-PcmValue $Config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource')
}

function Get-PcmResourceIntent {
    param([AllowNull()][object]$Item)
    if (-not $Item -or -not (Get-PcmValue $Item 'original_json' '')) { return $null }
    $value = [string]$Item.original_json | ConvertFrom-PcmJson
    if ([string](Get-PcmValue $value 'lifecycle' '') -eq 'CloudResource') {
        $token = [string](Get-PcmValue $value 'token' '')
        if ($token -notmatch '^[a-f0-9]{32}$' -or [string](Get-PcmValue $value 'creationName' '') -cnotmatch ("^PRA-R-[a-f0-9]{{8}}-[a-f0-9]{{8}}-{0}$" -f $token) -or
            [string](Get-PcmValue $value 'markerAttribute' '') -notin @('CustomAttribute14','CustomAttribute15')) { throw 'Native resource journal ownership token/name/marker is invalid; no identity is adopted.' }
        $null = Get-PcmResourceProtectionMode -Intent $value
        return $value
    }
    return $null
}

function Get-PcmResourceProtectionMode {
    param([AllowNull()][object]$Intent)
    if (-not $Intent) { return 'None' }
    # Missing mode means the original licensed/case-hold contract, never an implicit downgrade.
    $mode = Get-PcmValue $Intent 'protectionMode' 'CaseHold'
    if ($mode -isnot [string] -or $mode -cnotin @('None','CaseHold')) { throw "Unknown native resource protectionMode '$mode'; no mutation." }
    if ($mode -eq 'None' -and ((Get-PcmValue $Intent 'licensingGroupId' '') -or (Get-PcmValue $Intent 'holdTag' ''))) {
        throw 'Unprotected journal conflicts with recorded licence group/hold; no protection is removed.'
    }
    $null = Get-PcmUnprotectedDeletionProgress -Intent $Intent
    if ($mode -eq 'CaseHold' -and ((Get-PcmValue $Intent 'unprotectedDeletionAttempted' $false) -or (Get-PcmValue $Intent 'unprotectedDeletionProgress' ''))) {
        throw 'Protected journal conflicts with an unprotected deletion record; no deletion.'
    }
    return $mode
}

function Assert-PcmUnprotectedResource {
    param([Parameter(Mandatory)][object]$Owned, [switch]$CheckOrganization)
    foreach ($property in @('assignedLicenses','licenseAssignmentStates')) {
        if (-not $Owned.User.PSObject.Properties[$property]) { throw "Resource $property not exposed; cannot verify unchanged unlicensed identity." }
        if (@($Owned.User.$property | Where-Object { $null -ne $_ }).Count) { throw "Unexpected resource ${property}: unprotected lifecycle requires no licence assignment; nothing is revoked or deleted." }
    }
    $mailbox = $Owned.Mailbox
    foreach ($property in @('InPlaceHolds','LitigationHoldEnabled','RetentionHoldEnabled','DelayHoldApplied','DelayReleaseHoldApplied','ComplianceTagHoldApplied')) {
        if (-not $mailbox.PSObject.Properties[$property]) { throw "Resource $property not exposed; cannot verify absence of holds. No hold is removed." }
        if ($property -eq 'InPlaceHolds') {
            if (@($mailbox.$property | Where-Object { $_ -and [string]$_ -notmatch '^-' }).Count) { throw 'Unexpected active resource InPlaceHolds; no deletion or hold removal.' }
        } elseif ($mailbox.$property -ne $false) { throw "Unexpected active resource $property; no deletion or hold removal." }
    }
    if ($CheckOrganization) { Assert-PcmResourceOrganizationUnheld }
}

function Assert-PcmResourceOrganizationUnheld {
    $org = Get-OrganizationConfig -ErrorAction Stop
    if (-not $org -or -not $org.PSObject.Properties['InPlaceHolds'] -or $null -eq $org.InPlaceHolds) { throw 'Organization InPlaceHolds not exposed; no unprotected resource mutation.' }
    if (@($org.InPlaceHolds | Where-Object { $_ -and [string]$_ -notmatch '^-' }).Count) { throw 'Unexpected organization-wide holds; unprotected resource mutation refused. No hold is removed.' }
}

function Get-PcmUnprotectedDeletionProgress {
    <# Issuing is persisted before sending DELETE; an interrupted/ambiguous response never authorises a second deletion of a restored identity. #>
    param([Parameter(Mandatory)][object]$Intent)
    $progress = Get-PcmValue $Intent 'unprotectedDeletionProgress' ''
    $attempted = Get-PcmValue $Intent 'unprotectedDeletionAttempted' $false
    if ($attempted -isnot [bool]) { throw 'Unprotected deletion intent must be a Boolean; no deletion.' }
    if ($progress -isnot [string] -or $progress -cnotin @('', 'Intent', 'Issuing', 'Issued', 'Verified')) { throw "Unknown unprotected deletion progress '$progress'; no deletion." }
    if (-not $progress) { return $(if ($attempted) { 'Issuing' } else { '' }) }
    if (-not $attempted) { throw 'Unprotected deletion progress is journaled without a deletion intent; no deletion.' }
    return $progress
}

function Assert-PcmUnprotectedDeletionRepeatable {
    <# Only a pre-request intent or a definitively rejected request can authorise another DELETE. #>
    param([Parameter(Mandatory)][object]$Intent)
    $progress = Get-PcmUnprotectedDeletionProgress -Intent $Intent
    if ($progress -eq 'Verified') { throw "Deleted resource identity $($Intent.cloudId) is active again after a verified deletion; no second deletion and no routing contact. Investigate the restore." }
    if ($progress -eq 'Issuing') { throw "Resource identity $($Intent.cloudId) is visible after an interrupted or ambiguous deletion request; no second deletion. Wait for deletion evidence, or investigate a restore." }
    if ($progress -eq 'Issued') { throw "Resource identity $($Intent.cloudId) is still or again visible after its deletion returned; no second deletion until the directory shows it deleted. Wait, or investigate a restore." }
}

function Test-PcmUnprotectedResourceDeleted {
    param([Parameter(Mandatory)][object]$Intent)
    if ((Get-PcmResourceProtectionMode $Intent) -ne 'None') { throw 'Deletion proof applies only to the unprotected lifecycle; protected resources keep their inactive-copy verification.' }
    $progress = Get-PcmUnprotectedDeletionProgress -Intent $Intent
    if (-not $progress) { throw 'No journaled unprotected deletion intent; missing active identity is not proof of safe recovery.' }
    $cloudId = [string](Get-PcmValue $Intent 'cloudId' '')
    $cloudGuid = [string](Get-PcmValue $Intent 'cloudGuid' '')
    if (-not $cloudId -or -not $cloudGuid) { throw 'Journaled resource immutable ID/ExchangeGuid is missing; deletion cannot be verified.' }
    if (Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id" -f $cloudId) -AllowNotFound) {
        if ($progress -eq 'Verified') { Assert-PcmUnprotectedDeletionRepeatable -Intent $Intent }
        return $false
    }
    $deleted = Invoke-PcmGraph -Uri ("v1.0/directory/deletedItems/{0}?`$select=id" -f $cloudId) -AllowNotFound
    if (-not $deleted) { return $false }
    if ([string](Get-PcmValue $deleted 'id' '') -ne $cloudId) { throw 'Deleted resource identity differs from journal; no routing contact.' }
    $soft = @(Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited -ErrorAction Stop |
        Where-Object { [string]$_.ExchangeGuid -eq $cloudGuid })
    if (-not $soft.Count) { return $false }
    if ($soft.Count -ne 1 -or [string](Get-PcmValue $soft[0] 'ExternalDirectoryObjectId' '') -ne $cloudId -or
        -not $soft[0].PSObject.Properties['IsInactiveMailbox'] -or $soft[0].IsInactiveMailbox -ne $false) {
        throw 'Exact resource soft-deleted identity/GUID or non-inactive state cannot be verified; no contact or purge.'
    }
    $deletedUser = Invoke-PcmGraph -Uri ("v1.0/directory/deletedItems/{0}?`$select=id,assignedLicenses,licenseAssignmentStates" -f $cloudId)
    if ([string](Get-PcmValue $deletedUser 'id' '') -ne $cloudId) { throw 'Deleted resource licence identity differs from journal; no routing contact.' }
    Assert-PcmUnprotectedResource -Owned ([pscustomobject]@{ User = $deletedUser; Mailbox = $soft[0] }) -CheckOrganization
    return $true
}

function Get-PcmUnprotectedRecoverCollisions {
    param([Parameter(Mandatory)][object]$Intent, [Parameter(Mandatory)][object]$Record,
        [Parameter(Mandatory)][object]$Inventory, [string]$OwnedContactId = '')
    $cloudId = [string](Get-PcmValue $Intent 'cloudId' '')
    $cloudGuid = [string](Get-PcmValue $Intent 'cloudGuid' '')
    $parsedId = [guid]::Empty
    $parsedGuid = [guid]::Empty
    if ([string](Get-PcmValue $Intent 'lifecycle' '') -cne 'CloudResource' -or
        (Get-PcmResourceProtectionMode $Intent) -cne 'None' -or
        (Get-PcmUnprotectedDeletionProgress $Intent) -cne 'Verified' -or
        (Get-PcmValue $Intent 'creationAttempted' $false) -isnot [bool] -or $Intent.creationAttempted -ne $true -or
        [string]$Record.kind -cnotin @('Room','Equipment') -or
        [string](Get-PcmValue $Intent 'cloudType' '') -cne "$($Record.kind)Mailbox" -or
        -not [guid]::TryParse($cloudId, [ref]$parsedId) -or $parsedId -eq [guid]::Empty -or
        -not [guid]::TryParse($cloudGuid, [ref]$parsedGuid) -or $parsedGuid -eq [guid]::Empty -or
        $cloudGuid -eq [string]$Record.exchange_guid) {
        throw 'Expected resource tombstone requires verified unprotected journal ownership, independent immutable IDs and source type; no contact.'
    }
    # Fresh service proof, not a caller-supplied flag. Retained proxies alone do not establish an Exchange write conflict.
    if (-not (Test-PcmUnprotectedResourceDeleted -Intent $Intent)) { throw 'Exact resource deletion proof is no longer available; no routing contact or purge.' }
    $expected = @($Inventory.Recipients | Where-Object {
        $when = Get-PcmValue $_ 'WhenSoftDeleted' $null
        $validWhen = $when -is [datetime] -or $when -is [datetimeoffset]
        if ($when -is [string] -and -not [string]::IsNullOrWhiteSpace($when)) {
            $date = [datetimeoffset]::MinValue
            $validWhen = [datetimeoffset]::TryParse($when, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)
        }
        [string](Get-PcmValue $_ 'ExternalDirectoryObjectId' '') -eq $cloudId -and
        [string](Get-PcmValue $_ 'ExchangeGuid' '') -eq $cloudGuid -and
        [string](Get-PcmValue $_ 'RecipientTypeDetails' '') -ceq [string]$Intent.cloudType -and $validWhen
    })
    if ($expected.Count -gt 1) { throw 'Expected resource tombstone inventory is ambiguous; no routing contact.' }
    $filtered = [pscustomobject]@{
        Users = @($Inventory.Users | Where-Object { [string]$_.id -ne $cloudId })
        Recipients = @($Inventory.Recipients | Where-Object { $_ -notin $expected })
    }
    foreach ($property in @('RecipientKeys','GraphImmutableIds')) {
        if ($Inventory.PSObject.Properties[$property]) { $filtered | Add-Member -NotePropertyName $property -NotePropertyValue $Inventory.$property }
    }
    Get-PcmResourceCollisions -Record $Record -Inventory $filtered -OwnedContactId $OwnedContactId
}

function Get-PcmResourceAddresses {
    param([Parameter(Mandatory)][object]$Record)
    $primary = [string]$Record.primary_smtp_address
    if ($primary -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw 'Resource primary SMTP address is invalid.' }
    $values = @("SMTP:$primary")
    $json = [string](Get-PcmValue $Record 'email_addresses_json' '')
    if ($json) { $values += @(($json | ConvertFrom-PcmJson) | ForEach-Object { [string]$_ }) }
    $legacy = [string](Get-PcmValue $Record 'legacy_exchange_dn' '')
    if ($legacy) { $values += "X500:$legacy" }
    # Keep the original primary, not an accidental primary flag on another collected proxy.
    return @($values | Where-Object { $_ -and $_ -ine "smtp:$primary" } | ForEach-Object { if ($_ -cmatch '^SMTP:') { 'smtp:' + $_.Substring(5) } else { $_ } } |
        Get-PcmDistinctText) + @("SMTP:$primary")
}

function Get-PcmResourceKeys {
    param([object]$Record)
    $keys = @(Get-PcmResourceAddresses $Record | ForEach-Object { $_.ToLowerInvariant() })
    foreach ($field in @('user_principal_name', 'primary_smtp_address')) {
        $value = [string](Get-PcmValue $Record $field '')
        if ($value) { $keys += ('smtp:' + $value.ToLowerInvariant()) }
    }
    $alias = [string](Get-PcmValue $Record 'alias' '')
    if (-not $alias) { $alias = ([string]$Record.primary_smtp_address -split '@')[0] }
    $keys += ('alias:' + $alias.ToLowerInvariant())
    return @($keys | Get-PcmDistinctText)
}

function Test-PcmRecipientNotFound {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    # EXO can serialize the server exception as a generic exception; inspect metadata, never its localized message.
    if ($ErrorRecord.CategoryInfo.Category -notin @('ObjectNotFound', 'NotSpecified')) { return $false }
    $exception = $ErrorRecord.Exception
    while ($exception) {
        if ($exception.GetType().FullName -ceq 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException') { return $true }
        $remoteTypeName = Get-PcmValue -Object $exception.Data['RemoteException'] -Name 'TypeName'
        if ($remoteTypeName -is [string] -and $remoteTypeName -ceq 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException') { return $true }
        $exception = $exception.InnerException
    }
    if ($ErrorRecord.CategoryInfo.Reason -ceq 'ManagementObjectNotFoundException') { return $true }
    return ([string]$ErrorRecord.FullyQualifiedErrorId -cmatch '(?:^|,)(?:Microsoft\.Exchange\.Configuration\.Tasks\.)?ManagementObjectNotFoundException(?:,|$)|\[FailureCategory=Cmdlet-ManagementObjectNotFoundException\]')
}

function Get-PcmLegacyRecipient {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LegacyDN)
    if ([string]::IsNullOrWhiteSpace($LegacyDN)) { throw 'Resource legacy identity is empty; absence is not verified.' }
    $lookupResults = [System.Collections.Generic.List[object]]::new()
    try {
        # RecipientIdParameter resolves LegacyExchangeDN and X500 proxies by equality; LegacyExchangeDN is not Get-Recipient OPATH.
        Get-Recipient -Identity $LegacyDN -IncludeSoftDeletedRecipients -ErrorAction Stop | ForEach-Object { $lookupResults.Add($_) }
    } catch {
        if ($lookupResults.Count -or -not (Test-PcmRecipientNotFound -ErrorRecord $_)) { throw }
    }
    foreach ($recipient in $lookupResults) {
        if ([string](Get-PcmValue $recipient 'LegacyExchangeDN' '') -ine $LegacyDN -and
            @((Get-PcmValue $recipient 'EmailAddresses' @()) | ForEach-Object { [string]$_ }) -inotcontains "x500:$LegacyDN") {
            throw 'Resource legacy identity lookup returned a non-exact recipient; absence is not verified.'
        }
        $recipient
    }
}

function Get-PcmResourceGraphFilter {
    param([object[]]$Records, [string[]]$Keys, [string[]]$ImmutableIds)
    # Non-SMTP proxy equality needs the advanced index; keep SMTP/scalar predicates on the default query path.
    # https://learn.microsoft.com/graph/aad-advanced-queries#user-properties
    $predicates = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $advancedPredicates = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $proxies = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($key in $Keys) {
        if ($key.StartsWith('alias:')) {
            $null = $predicates.Add(("mailNickname eq '{0}'" -f $key.Substring(6).Replace("'", "''")))
        } else {
            $null = $proxies.Add($key)
            if ($key.StartsWith('smtp:')) {
                foreach ($property in 'userPrincipalName', 'mail') {
                    $null = $predicates.Add(("{0} eq '{1}'" -f $property, $key.Substring(5).Replace("'", "''")))
                }
            }
        }
    }
    foreach ($record in $Records) {
        foreach ($address in @(Get-PcmResourceAddresses $record)) { $null = $proxies.Add($address) }
    }
    foreach ($proxy in @($proxies)) {
        # Preserve collected values and query both primary/secondary SMTP and X500 protocol spellings.
        if ($proxy -imatch '^(smtp|x500):') {
            $prefix, $address = $proxy -split ':', 2
            $null = $proxies.Add($prefix.ToUpperInvariant() + ':' + $address)
            $null = $proxies.Add($prefix.ToLowerInvariant() + ':' + $address)
        }
    }
    foreach ($proxy in $proxies) {
        $target = $advancedPredicates
        if ($proxy -imatch '^smtp:') { $target = $predicates }
        $null = $target.Add(("proxyAddresses/any(p:p eq '{0}')" -f $proxy.Replace("'", "''")))
    }
    foreach ($immutableId in $ImmutableIds) { $null = $predicates.Add(("onPremisesImmutableId eq '{0}'" -f $immutableId.Replace("'", "''"))) }
    foreach ($mode in 'Default', 'Advanced') {
        $modePredicates = $advancedPredicates
        if ($mode -eq 'Default') { $modePredicates = $predicates }
        $clauses = [System.Collections.Generic.List[string]]::new()
        foreach ($predicate in @($modePredicates | Sort-Object -CaseSensitive)) {
            $clause = '(' + $predicate + ')'
            if ($clause.Length -gt 2000 -or [uri]::EscapeDataString($clause).Length -gt 3000) {
                throw 'Resource inventory key exceeds the bounded Graph filter size; absence is not verified.'
            }
            $combined = (@($clauses.ToArray()) + @($clause)) -join ' or '
            if ($clauses.Count -and ($clauses.Count -ge 15 -or $combined.Length -gt 2000 -or [uri]::EscapeDataString($combined).Length -gt 3000)) {
                [pscustomobject]@{ Filter = $clauses -join ' or '; QueryMode = $mode }
                $clauses.Clear()
            }
            $clauses.Add($clause)
        }
        if ($clauses.Count) { [pscustomobject]@{ Filter = $clauses -join ' or '; QueryMode = $mode } }
    }
}

function Get-PcmResourceInventory {
    # Both inventories cover the exact selected-key union; Graph also covers the original immutable anchors.
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Records = @())
    if (-not $Records.Count) { throw 'Resource inventory requires selected records; absence is not verified.' }
    $keys = @(@(foreach ($record in $Records) {
            if (-not $record) { throw 'Resource inventory contains an empty selected record; absence is not verified.' }
            Get-PcmResourceKeys $record
        }) | Sort-Object -Unique)
    $filters = [System.Collections.Generic.List[string]]::new()
    $clauses = [System.Collections.Generic.List[string]]::new()
    foreach ($key in $keys) {
        $predicates = if ($key.StartsWith('alias:')) {
            "(Alias -eq '{0}')" -f $key.Substring(6).Replace("'", "''")
        } else {
            "(EmailAddresses -eq '{0}')" -f $key.Replace("'", "''")
            if ($key.StartsWith('smtp:')) { "(PrimarySmtpAddress -eq '{0}')" -f $key.Substring(5).Replace("'", "''") }
            if ($key.StartsWith('x500:') -and [string]::IsNullOrWhiteSpace($key.Substring(5))) { throw 'Resource legacy identity is empty; absence is not verified.' }
        }
        foreach ($predicate in $predicates) {
            if ($predicate.Length -gt 4000) { throw 'Resource inventory key exceeds the bounded recipient filter size; absence is not verified.' }
            if ($clauses.Count -and ($clauses.Count -ge $script:ExoFilterSize -or (($clauses -join ' -or ').Length + $predicate.Length + 5) -gt 4000)) {
                $filters.Add(($clauses -join ' -or '))
                $clauses.Clear()
            }
            $clauses.Add($predicate)
        }
    }
    if ($clauses.Count) { $filters.Add(($clauses -join ' -or ')) }
    if (-not $filters.Count) { throw 'Resource inventory has no selected keys; absence is not verified.' }
    $immutableIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    foreach ($record in $Records) {
        $immutableId = [string](Get-PcmValue $record 'immutable_id' '')
        if ($immutableId) { $null = $immutableIds.Add($immutableId) }
    }
    $graphFilters = @(Get-PcmResourceGraphFilter -Records $Records -Keys $keys -ImmutableIds @($immutableIds))
    if (-not $graphFilters.Count) { throw 'Resource inventory has no Graph predicates; absence is not verified.' }
    # No SilentlyContinue here: an incomplete recipient inventory is not evidence that addresses are free.
    $recipients = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($filter in $filters) {
        foreach ($recipient in @(Get-Recipient -Filter $filter -ResultSize Unlimited -IncludeSoftDeletedRecipients -ErrorAction Stop)) {
            $identity = [string](Get-PcmValue $recipient 'Guid' '')
            if (-not $identity) { $identity = [string](Get-PcmValue $recipient 'ExternalDirectoryObjectId' '') }
            if ($identity -and $seen.ContainsKey($identity)) { continue }
            if ($identity) { $seen[$identity] = $true }
            $recipients.Add($recipient)
        }
    }
    foreach ($key in @($keys | Where-Object { $_.StartsWith('x500:') })) {
        foreach ($recipient in @(Get-PcmLegacyRecipient -LegacyDN $key.Substring(5))) {
            $identity = [string](Get-PcmValue $recipient 'Guid' '')
            if (-not $identity) { $identity = [string](Get-PcmValue $recipient 'ExternalDirectoryObjectId' '') }
            if ($identity -and $seen.ContainsKey($identity)) { continue }
            if ($identity) { $seen[$identity] = $true }
            $recipients.Add($recipient)
        }
    }
    $users = [System.Collections.Generic.List[object]]::new()
    $userFacts = @{}
    $select = 'id,userPrincipalName,mail,mailNickname,proxyAddresses,onPremisesImmutableId,onPremisesSyncEnabled'
    $queries = @(foreach ($filter in $graphFilters) {
            $uri = 'https://graph.microsoft.com/v1.0/users?$select=' + $select + '&$top=999&$filter=' + [uri]::EscapeDataString($filter.Filter)
            $headers = $null
            if ($filter.QueryMode -eq 'Advanced') {
                # This index is eventually consistent, not an atomic directory snapshot; EXO checks remain independent.
                $uri += '&$count=true'
                $headers = @{ ConsistencyLevel = 'eventual' }
            }
            [pscustomobject]@{ Path = '/v1.0/users'; Uri = $uri; Headers = $headers }
        })
    # Keep the deleted-user scan: the equality union is not qualified on this endpoint.
    $queries += [pscustomobject]@{ Path = '/v1.0/directory/deletedItems/microsoft.graph.user'
        Uri = 'https://graph.microsoft.com/v1.0/directory/deletedItems/microsoft.graph.user?$select=' + $select + '&$top=999'; Headers = $null }
    foreach ($query in $queries) {
        $uri = $query.Uri
        $visited = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
        do {
            $parsedUri = $null
            if (-not [uri]::TryCreate($uri, [UriKind]::Absolute, [ref]$parsedUri) -or $parsedUri.Scheme -cne 'https' -or
                $parsedUri.Host -ine 'graph.microsoft.com' -or -not $parsedUri.IsDefaultPort -or $parsedUri.UserInfo -or
                $parsedUri.AbsolutePath -cne $query.Path -or $parsedUri.Fragment -or -not $visited.Add($uri)) {
                throw 'Resource Graph inventory returned an invalid or repeated page link; absence is not verified.'
            }
            $parameters = @{ Uri = $uri }
            # Graph does not carry the consistency header in nextLink; retain the mode on every page.
            if ($query.Headers) { $parameters.Headers = $query.Headers }
            $page = Invoke-PcmGraph @parameters
            if (-not $page -or -not $page.PSObject.Properties['value'] -or $null -eq $page.value -or
                $page.value -isnot [System.Collections.IEnumerable] -or $page.value -is [string] -or $page.value -is [System.Collections.IDictionary]) {
                throw 'Resource Graph inventory returned no value collection; absence is not verified.'
            }
            foreach ($user in @($page.value)) {
                foreach ($property in $select.Split(',')) {
                    if (-not $user -or -not $user.PSObject.Properties[$property]) { throw 'Resource Graph inventory returned limited identity information; absence is not verified.' }
                }
                if ($user.id -isnot [string] -or [string]::IsNullOrWhiteSpace($user.id) -or
                    $user.userPrincipalName -isnot [string] -or [string]::IsNullOrWhiteSpace($user.userPrincipalName)) {
                    throw 'Resource Graph inventory returned limited identity information; absence is not verified.'
                }
                foreach ($property in 'mail', 'mailNickname', 'onPremisesImmutableId') {
                    if ($null -ne $user.$property -and $user.$property -isnot [string]) { throw 'Resource Graph inventory returned invalid identity information; absence is not verified.' }
                }
                if ($null -eq $user.proxyAddresses -or $user.proxyAddresses -isnot [System.Collections.IEnumerable] -or
                    $user.proxyAddresses -is [string] -or $user.proxyAddresses -is [System.Collections.IDictionary] -or
                    @($user.proxyAddresses | Where-Object { $_ -isnot [string] -or [string]::IsNullOrWhiteSpace($_) }).Count -or
                    ($null -ne $user.onPremisesSyncEnabled -and $user.onPremisesSyncEnabled -isnot [bool])) {
                    throw 'Resource Graph inventory returned invalid identity information; absence is not verified.'
                }
                $candidate = @(('smtp:' + $user.userPrincipalName), ('smtp:' + [string]$user.mail), ('alias:' + [string]$user.mailNickname)) + @($user.proxyAddresses)
                if ($query.Path -ceq '/v1.0/users' -and -not @($candidate | Where-Object { $_.ToLowerInvariant() -in $keys }).Count -and
                    -not $immutableIds.Contains([string]$user.onPremisesImmutableId)) {
                    throw 'Resource Graph inventory returned a user outside the exact selected scope; absence is not verified.'
                }
                $facts = $user | Select-Object -Property $select.Split(',') | ConvertTo-Json -Depth 4 -Compress
                if ($userFacts.ContainsKey($user.id)) {
                    if ($userFacts[$user.id] -cne $facts) { throw 'Resource Graph inventory returned inconsistent duplicate identity information; absence is not verified.' }
                    continue
                }
                $userFacts[$user.id] = $facts
                $users.Add($user)
            }
            $uri = ''
            if ($page.PSObject.Properties['@odata.nextLink']) {
                $nextLink = $page.'@odata.nextLink'
                if ($nextLink -isnot [string] -or [string]::IsNullOrWhiteSpace($nextLink)) {
                    throw 'Resource Graph inventory returned an invalid page link; absence is not verified.'
                }
                $uri = $nextLink
            }
        } while ($uri)
    }
    return [pscustomobject]@{ Recipients = $recipients.ToArray(); Users = $users.ToArray(); RecipientKeys = $keys; GraphImmutableIds = @($immutableIds) }
}

function Get-PcmResourceCollisions {
    param([object]$Record, [object]$Inventory, [string]$OwnedId = '', [string]$OwnedGuid = '', [string]$OwnedContactId = '')
    $keys = @(Get-PcmResourceKeys $Record)
    if ($Inventory.PSObject.Properties['RecipientKeys'] -and @($keys | Where-Object { $_ -notin $Inventory.RecipientKeys }).Count) {
        throw 'Resource collision inventory does not cover all selected record keys; absence is not verified.'
    }
    $immutableId = [string](Get-PcmValue $Record 'immutable_id' '')
    if ($Inventory.PSObject.Properties['GraphImmutableIds'] -and $immutableId -and $immutableId -cnotin $Inventory.GraphImmutableIds) {
        throw 'Resource collision inventory does not cover the selected immutable anchor; absence is not verified.'
    }
    foreach ($user in $Inventory.Users) {
        if ($OwnedId -and [string]$user.id -eq $OwnedId) { continue }
        $candidate = @(('smtp:' + [string]$user.userPrincipalName), ('smtp:' + [string](Get-PcmValue $user 'mail' ''))) +
            @((Get-PcmValue $user 'proxyAddresses' @()) | ForEach-Object { [string]$_ })
        $candidate += ('alias:' + [string](Get-PcmValue $user 'mailNickname' ''))
        if (@($candidate | Where-Object { $_.ToLowerInvariant() -in $keys }).Count -or
            ($immutableId -and [string](Get-PcmValue $user 'onPremisesImmutableId' '') -ceq $immutableId)) {
            "Existing Entra identity $($user.id) ($($user.userPrincipalName)) collides with this native resource; never adopted."
        }
    }
    foreach ($recipient in $Inventory.Recipients) {
        $id = [string](Get-PcmValue $recipient 'ExternalDirectoryObjectId' '')
        $guid = [string](Get-PcmValue $recipient 'ExchangeGuid' '')
        $contactId = [string](Get-PcmValue $recipient 'Guid' '')
        if (($OwnedId -and $id -eq $OwnedId) -or ($OwnedGuid -and $guid -eq $OwnedGuid) -or
            ($OwnedContactId -and $contactId -eq $OwnedContactId)) { continue }
        $candidate = @(('smtp:' + [string](Get-PcmValue $recipient 'PrimarySmtpAddress' '')), ('alias:' + [string](Get-PcmValue $recipient 'Alias' ''))) +
            @((Get-PcmValue $recipient 'EmailAddresses' @()) | ForEach-Object { [string]$_ })
        $legacy = [string](Get-PcmValue $recipient 'LegacyExchangeDN' '')
        if ($legacy) { $candidate += "x500:$legacy" }
        if (@($candidate | Where-Object { $_.ToLowerInvariant() -in $keys }).Count) {
            "Existing Exchange recipient $($recipient.Identity) ($($recipient.RecipientTypeDetails), ID $id, ExchangeGuid $guid) collides with a resource address or alias; left untouched, including soft-deleted recipients."
        }
    }
}

function Get-PcmOwnedResource {
    param([Parameter(Mandatory)][object]$Intent, [Parameter(Mandatory)][object]$Record)
    $recipients = @(Get-Recipient -Filter ("Name -eq '{0}'" -f ([string]$Intent.creationName).Replace("'", "''")) -ResultSize Unlimited -ErrorAction Stop)
    if (-not $recipients.Count) { return $null }
    if ($recipients.Count -ne 1) { throw 'Ambiguous resource creation identity; no mailbox is adopted.' }
    $recipient = $recipients[0]
    if (-not [bool](Get-PcmValue $Intent 'creationAttempted' $false)) { throw 'Resource creation was never journaled as attempted; a pre-existing name is not adopted.' }
    $expectedType = "$($Record.kind)Mailbox"
    if ([string]$recipient.Name -cne [string]$Intent.creationName -or [string]$recipient.RecipientTypeDetails -ne $expectedType) { throw 'Resource creation name/type does not prove ownership.' }
    $id = [string]$recipient.ExternalDirectoryObjectId
    $guid = [string]$recipient.ExchangeGuid
    if (-not $id -or -not $guid -or $guid -eq [guid]::Empty.ToString() -or $guid -eq [string]$Record.exchange_guid) { throw 'Resource has no independent cloud identity/ExchangeGuid.' }
    $persistedId = [string](Get-PcmValue $Intent 'cloudId' '')
    $persistedGuid = [string](Get-PcmValue $Intent 'cloudGuid' '')
    if (($persistedId -and $id -ne $persistedId) -or ($persistedGuid -and $guid -ne $persistedGuid)) { throw 'Resource immutable ID or ExchangeGuid differs from the journal; nothing is changed.' }
    $mailbox = Get-Mailbox -Identity $id -ErrorAction Stop
    $marker = [string](Get-PcmValue $mailbox ([string]$Intent.markerAttribute) '')
    # New-Mailbox has no custom-attribute parameter. The unpredictable token is embedded in Name atomically;
    # only this journaled creation name can bridge the crash before the marker and IDs were persisted.
    if ($marker -cne [string]$Intent.token -and ($marker -or ($persistedId -and -not (Get-PcmValue $Intent 'markerPending' $false)))) { throw 'Resource ownership marker differs from the journal.' }
    $user = Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id,userPrincipalName,onPremisesSyncEnabled,onPremisesImmutableId,accountEnabled,assignedPlans,licenseAssignmentStates,assignedLicenses,usageLocation" -f $id) -AllowNotFound
    if (-not $user) { throw 'Owned resource Graph identity is not yet visible; retry provisioning, never adopt another user.' }
    $exoUser = Get-User -Identity $id -ErrorAction Stop
    if ([string](Get-PcmValue $user 'id' '') -ne $id -or
        -not $user.PSObject.Properties['accountEnabled'] -or $user.accountEnabled -isnot [bool] -or
        -not $user.PSObject.Properties['onPremisesSyncEnabled'] -or
        ($null -ne $user.onPremisesSyncEnabled -and $user.onPremisesSyncEnabled -isnot [bool]) -or
        -not $user.PSObject.Properties['onPremisesImmutableId'] -or
        ($null -ne $user.onPremisesImmutableId -and $user.onPremisesImmutableId -isnot [string])) {
        throw 'Resource Graph identity/account/source-anchor read is missing, malformed or inconsistent; nothing is changed.'
    }
    if ($user.onPremisesSyncEnabled -or $user.onPremisesImmutableId -or
        -not $exoUser.PSObject.Properties['IsDirSynced'] -or $exoUser.IsDirSynced -isnot [bool] -or $exoUser.IsDirSynced) { throw 'Resource identity is not cloud-only; nothing is changed.' }
    if ([string]$mailbox.ExchangeGuid -ne $guid -or [string]$mailbox.ExternalDirectoryObjectId -ne $id) { throw 'Resource mailbox immutable identity read is inconsistent.' }
    return [pscustomobject]@{ Recipient = $recipient; Mailbox = $mailbox; User = $user; Id = $id; Guid = $guid; Marker = $marker }
}

function Get-PcmResourceState {
    param([object]$Record, [object[]]$Records, [object]$Inventory, [AllowNull()][object]$Intent)
    if ((Get-PcmResourceProtectionMode $Intent) -eq 'None') { Assert-PcmResourceOrganizationUnheld }
    $owned = if ($Intent) { Get-PcmOwnedResource -Intent $Intent -Record $Record } else { $null }
    if ($Intent -and (Get-PcmValue $Intent 'cloudId' '') -and -not $owned) { throw 'Journal-owned active resource is missing; use Recover to verify deleted/inactive state.' }
    $collisions = @(Get-PcmResourceCollisions -Record $Record -Inventory $Inventory -OwnedId $(if ($owned) { $owned.Id } else { '' }))
    $keys = @(Get-PcmResourceKeys $Record)
    foreach ($other in @($Records | Where-Object { [string]$_.object_guid -ne [string]$Record.object_guid })) {
        if (@(Get-PcmResourceKeys $other | Where-Object { $_ -in $keys }).Count) { $collisions += "Competing selected snapshot object $($other.object_guid) claims a resource address/alias." }
    }
    if ($owned -and (Get-PcmResourceProtectionMode $Intent) -eq 'None') { Assert-PcmUnprotectedResource -Owned $owned }
    return [pscustomobject]@{ NativeResource = $true; JournalNative = [bool]$Intent; ProtectionMode = (Get-PcmResourceProtectionMode $Intent); Owned = [bool]$owned; AbsenceVerified = $true; Collisions = $collisions
        User = $(if ($owned) { $owned.User } else { $null }); MatchedBy = ''; IsCloudManaged = $null
        RecipientType = $(if ($owned) { [string]$owned.Recipient.RecipientTypeDetails } else { '' })
        ExchangeGuid = $(if ($owned) { $owned.Guid } else { '' }); Holds = @(); Locations = @(); TagValue = '' }
}

$script:ResourceCalendarFields = @('AutomateProcessing','AllowConflicts','BookingWindowInDays','MaximumDurationInMinutes','AllowRecurringMeetings',
    'EnforceSchedulingHorizon','ScheduleOnlyDuringWorkHours','ConflictPercentageAllowed','MaximumConflictInstances','ForwardRequestsToDelegates',
    'DeleteAttachments','DeleteComments','RemovePrivateProperty','DeleteSubject','AddOrganizerToSubject','DeleteNonCalendarItems','TentativePendingApproval',
    'EnableResponseDetails','OrganizerInfo','AllBookInPolicy','AllRequestInPolicy','AllRequestOutOfPolicy','AddAdditionalResponse','AdditionalResponse',
    'ProcessExternalMeetingMessages','RemoveOldMeetingMessages','ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy')
$script:ResourceRecipientFields = @('ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy')
$script:ResourceCalendarRequired = @('AutomateProcessing','AllBookInPolicy','AllRequestInPolicy','AllRequestOutOfPolicy',
    'ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy','ScheduleOnlyDuringWorkHours')
$script:ResourceWorkFields = @('WorkingHoursStartTime','WorkingHoursEndTime','WorkingHoursDays','WorkingHoursTimeZone')
$script:ResourceDirectoryFields = @('Office','StreetAddress','City','StateOrProvince','PostalCode','CountryOrRegion','Phone','Notes')
$script:ResourceDeliveryRecipients = @('AcceptMessagesOnlyFromSendersOrMembers','RejectMessagesFromSendersOrMembers','ModeratedBy','BypassModerationFromSendersOrMembers')
$script:ResourceDeliveryFields = @('RequireSenderAuthenticationEnabled','ModerationEnabled','SendModerationNotifications') + $script:ResourceDeliveryRecipients

function ConvertTo-PcmResourceWorkDays {
    param([AllowEmptyCollection()][object[]]$Days)
    $names = @(foreach ($value in $Days) {
            foreach ($name in ([string]$value -split ',\s*')) {
                switch ($name.Trim().ToLowerInvariant()) {
                    weekdays { 'monday'; 'tuesday'; 'wednesday'; 'thursday'; 'friday' }
                    weekenddays { 'saturday'; 'sunday' }
                    alldays { 'monday'; 'tuesday'; 'wednesday'; 'thursday'; 'friday'; 'saturday'; 'sunday' }
                    default {
                        if ($name.Trim().ToLowerInvariant() -notin @('monday','tuesday','wednesday','thursday','friday','saturday','sunday')) { throw 'Resource working-hours days are invalid: Collect is required.' }
                        $name.Trim().ToLowerInvariant()
                    }
                }
            }
        })
    if (-not $names.Count) { throw 'Resource working-hours days must explicitly name the weekdays: Collect is required.' }
    return @($names | Sort-Object -Unique)
}

function Read-PcmResourceSettings {
    param([object]$Record)
    $json = [string](Get-PcmValue $Record 'resource_settings_json' '')
    if (-not $json) { throw 'Native resource settings are missing (v1 snapshot): Collect is required before CloudResource Convert.' }
    $settings = $json | ConvertFrom-PcmJson
    if ([int](Get-PcmValue $settings 'SchemaVersion' 0) -ne 2 -or -not $settings.PSObject.Properties['Capacity'] -or
        -not $settings.PSObject.Properties['ResourceCustom'] -or -not $settings.PSObject.Properties['Calendar'] -or
        -not $settings.PSObject.Properties['CalendarConfiguration'] -or -not $settings.PSObject.Properties['Directory'] -or
        -not $settings.PSObject.Properties['RoomLists'] -or -not $settings.PSObject.Properties['Delivery']) { throw 'Incomplete resource metadata/settings schema: a new Collect is required.' }
    Test-PcmDeliverySettings -Delivery $settings.Delivery
    foreach ($field in $script:ResourceDirectoryFields) {
        if (-not $settings.Directory -or -not $settings.Directory.PSObject.Properties[$field] -or $settings.Directory.$field -isnot [string]) {
            throw "Resource descriptive property Directory.$field missing/invalid: a new Collect is required."
        }
    }
    if ($settings.RoomLists -isnot [array]) { throw 'Resource RoomLists must explicitly contain an array: a new Collect is required.' }
    if ([string]$Record.kind -ne 'Room' -and $settings.RoomLists.Count) { throw 'Only room mailboxes can belong to RoomLists.' }
    $seenLists = @{}
    foreach ($list in $settings.RoomLists) {
        Test-PcmRoomListMetadata -List $list
        if ($seenLists.ContainsKey($list.SourceGuid)) { throw 'Duplicate resource room list: a new Collect is required.' }
        $seenLists[$list.SourceGuid] = $true
    }
    if (($settings.Capacity -isnot [int] -and $settings.Capacity -isnot [long]) -or [long]$settings.Capacity -lt 0 -or [long]$settings.Capacity -gt [int]::MaxValue) { throw 'Resource capacity must be a non-negative integer: Collect is required.' }
    foreach ($property in @($settings.ResourceCustom)) { if ($property -isnot [string] -or -not $property) { throw 'Invalid resource custom property: Collect is required.' } }
    foreach ($field in $script:ResourceCalendarRequired) {
        if (-not $settings.Calendar.PSObject.Properties[$field]) { throw "Resource Calendar.$field missing: Collect is required (no booking-policy broadening)." }
    }
    foreach ($field in $script:ResourceWorkFields) {
        if (-not $settings.CalendarConfiguration.PSObject.Properties[$field]) { throw "Resource CalendarConfiguration.$field missing: Collect is required." }
    }
    foreach ($field in $script:ResourceRecipientFields) {
        foreach ($smtp in @($settings.Calendar.$field)) { if ([string]$smtp -notmatch '^[^@\s]+@[^@\s]+$') { throw "Resource Calendar.$field has a non-SMTP recipient: Collect is required." } }
    }
    if ($settings.Calendar.AutomateProcessing -notin @('None','AutoUpdate','AutoAccept')) { throw 'Invalid resource AutomateProcessing: Collect is required.' }
    $integerFields = @('BookingWindowInDays','MaximumDurationInMinutes','ConflictPercentageAllowed','MaximumConflictInstances')
    foreach ($field in $integerFields) {
        if (-not $settings.Calendar.PSObject.Properties[$field]) { continue }
        if (($settings.Calendar.$field -isnot [int] -and $settings.Calendar.$field -isnot [long]) -or [long]$settings.Calendar.$field -lt 0) { throw "Invalid integer Calendar.${field}: Collect is required." }
    }
    foreach ($field in @($script:ResourceCalendarFields | Where-Object { $_ -notin $script:ResourceRecipientFields -and $_ -notin $integerFields -and $_ -notin @('AutomateProcessing','AdditionalResponse') })) {
        if (-not $settings.Calendar.PSObject.Properties[$field]) { continue }
        if ($settings.Calendar.$field -isnot [bool]) { throw "Invalid boolean Calendar.${field}: Collect is required (no implicit booking-policy broadening)." }
    }
    foreach ($field in @('WorkingHoursStartTime','WorkingHoursEndTime')) {
        if ([string]$settings.CalendarConfiguration.$field -notmatch '^(?:[01][0-9]|2[0-3]):[0-5][0-9]:[0-5][0-9]$') { throw "Invalid CalendarConfiguration.${field}: Collect is required." }
    }
    if (-not [string]$settings.CalendarConfiguration.WorkingHoursTimeZone) { throw 'Resource working-hours time zone is missing: Collect is required.' }
    $null = ConvertTo-PcmResourceWorkDays -Days @($settings.CalendarConfiguration.WorkingHoursDays)
    return $settings
}

function Test-PcmDeliverySettings {
    param([object]$Delivery)
    foreach ($field in @('RequireSenderAuthenticationEnabled','ModerationEnabled')) {
        if (-not $Delivery -or -not $Delivery.PSObject.Properties[$field] -or $Delivery.$field -isnot [bool]) {
            throw "Resource Delivery.$field missing/invalid: a new Collect is required (no delivery-policy broadening)."
        }
    }
    if ([string](Get-PcmValue $Delivery 'SendModerationNotifications' '') -notin @('Never','Internal','Always')) { throw 'Invalid resource moderation notification setting: a new Collect is required.' }
    foreach ($field in $script:ResourceDeliveryRecipients) {
        if (-not $Delivery.PSObject.Properties[$field] -or $Delivery.$field -isnot [array]) { throw "Resource Delivery.$field must contain an explicit recipient array: a new Collect is required." }
        foreach ($smtp in $Delivery.$field) {
            if ($smtp -isnot [string] -or $smtp -notmatch '^[^@\s]+@[^@\s]+$') { throw "Resource Delivery.$field has an invalid SMTP recipient: a new Collect is required." }
        }
    }
    if ($Delivery.ModerationEnabled -and -not $Delivery.ModeratedBy.Count) { throw 'Resource moderation requires explicit moderators: a new Collect is required.' }
}

function Test-PcmRoomListMetadata {
    param([object]$List)
    foreach ($field in @('SourceGuid','Name','DisplayName','Alias','PrimarySmtpAddress','LegacyExchangeDN')) {
        if (-not $List -or -not $List.PSObject.Properties[$field] -or $List.$field -isnot [string] -or ($field -ne 'LegacyExchangeDN' -and -not $List.$field)) {
            throw "Resource RoomLists.$field missing/invalid: a new Collect is required."
        }
    }
    if ([int](Get-PcmValue $List 'SchemaVersion' 0) -ne 1 -or $List.PrimarySmtpAddress -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or
        -not $List.PSObject.Properties['HiddenFromAddressListsEnabled'] -or $List.HiddenFromAddressListsEnabled -isnot [bool] -or
        -not $List.PSObject.Properties['EmailAddresses'] -or $List.EmailAddresses -isnot [array]) { throw 'Incomplete room-list metadata: a new Collect is required.' }
    foreach ($address in $List.EmailAddresses) { if ($address -isnot [string] -or -not $address) { throw 'Invalid room-list address: a new Collect is required.' } }
    if ([string](Get-PcmValue $List 'MemberJoinRestriction' '') -notin @('Closed','Open','ApprovalRequired') -or
        [string](Get-PcmValue $List 'MemberDepartRestriction' '') -notin @('Closed','Open')) { throw 'Invalid room-list membership controls: a new Collect is required.' }
    Test-PcmDeliverySettings -Delivery (Get-PcmValue $List 'Delivery' $null)
}

function ConvertTo-PcmRoomListRecord {
    param([object]$List)
    return [pscustomobject]@{ object_guid = $List.SourceGuid; kind = 'RoomList'; alias = $List.Alias; primary_smtp_address = $List.PrimarySmtpAddress
        email_addresses_json = (ConvertTo-Json -InputObject @($List.EmailAddresses) -Compress); legacy_exchange_dn = $List.LegacyExchangeDN
        immutable_id = ''; user_principal_name = ''; exchange_guid = '' }
}

function Test-PcmResourceConfigurationReadiness {
    param([object]$Record, [object[]]$FutureResources = @(), [hashtable]$ResourceIntents = @{},
        [hashtable]$RoomListIntents = @{}, [object[]]$RoomListPlan = @(), [switch]$AllowRoomListCreation)
    $settings = Read-PcmResourceSettings $Record
    $command = Get-Command -Name Set-Mailbox -ErrorAction Stop
    foreach ($parameter in (@('ResourceCapacity','ResourceCustom') + $script:ResourceDeliveryFields)) {
        if (-not $command.Parameters.ContainsKey($parameter)) {
            throw "Set-Mailbox -$parameter is not exposed by this Exchange Online session/app. Check Exchange Online cmdlet support and RBAC before Convert; the resource is not created and tenant-wide resource schema is not changed."
        }
    }
    $userCommand = Get-Command -Name Set-User -ErrorAction Stop
    foreach ($field in $script:ResourceDirectoryFields) {
        if (-not $userCommand.Parameters.ContainsKey($field)) { throw "Set-User -$field is not exposed by this Exchange Online session/app; resource descriptive properties cannot be restored. Check RBAC before Convert." }
    }
    foreach ($list in $settings.RoomLists) {
        $listIntent = $RoomListIntents[[string]$list.SourceGuid]
        if (-not $AllowRoomListCreation -and (-not $listIntent -or $listIntent.status -ne 'Done')) { throw "Room list $($list.PrimarySmtpAddress) is not provisioned/verified; dependent rooms are not created." }
        $group = Get-PcmResourceCloudRoomList -List $list -Intent $listIntent -AllowMissing:$AllowRoomListCreation
        if ($group) { $null = Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop }
        else { Test-PcmRoomListCreationReadiness -List $list }
        $memberCommand = Get-Command -Name Add-DistributionGroupMember -ErrorAction Stop
        if (-not $memberCommand.Parameters.ContainsKey('BypassSecurityGroupManagerCheck')) { throw 'Add-DistributionGroupMember -BypassSecurityGroupManagerCheck is not exposed; check room-list membership RBAC before Convert.' }
    }
    $references = @(foreach ($field in $script:ResourceRecipientFields) { [pscustomobject]@{ Control = 'booking'; Values = @($settings.Calendar.$field); Field = $field } }
        foreach ($field in $script:ResourceDeliveryRecipients) { [pscustomobject]@{ Control = 'delivery'; Values = @($settings.Delivery.$field); Field = $field } })
    foreach ($reference in $references) {
        $field = $reference.Field
        foreach ($smtp in $reference.Values) {
            if ($AllowRoomListCreation) {
                $futureLists = @($RoomListPlan | Where-Object { "smtp:$smtp" -in @(Get-PcmResourceAddresses (ConvertTo-PcmRoomListRecord $_)) })
                if ($futureLists.Count -eq 1) { continue }
                if ($futureLists.Count -gt 1) { throw "Ambiguous planned room-list recipient $smtp; no policy is broadened." }
            }
            $future = @($FutureResources | Where-Object { "smtp:$smtp" -in @(Get-PcmResourceAddresses $_) })
            if ($future.Count) {
                if ($future.Count -ne 1) { throw "Resource $($reference.Control) recipient $smtp ($field) matches competing native snapshot objects; no policy is omitted or broadened." }
                $intent = $ResourceIntents[[string]$future[0].object_guid]
                if (-not $intent -or -not [string](Get-PcmValue $intent 'cloudId' '')) {
                    throw "Resource $($reference.Control) recipient $smtp ($field) references a native resource not yet journal-owned/provisioned. Provision that referenced resource first, then resume this Convert batch. This resource is not created; no booking recipient is omitted or policy broadened."
                }
                $owned = Get-PcmOwnedResource -Intent $intent -Record $future[0]
                if (-not $owned) { throw "Resource $($reference.Control) recipient $smtp ($field) cannot be verified as the journal-owned native resource; no policy is omitted or broadened." }
            }
            try { $resolved = @(Get-Recipient -Identity ([string]$smtp) -ErrorAction Stop) }
            catch { throw "Resource $($reference.Control) recipient $smtp ($field) could not be read/resolved before creation; no policy is omitted or broadened. $($_.Exception.Message)" }
            if ($resolved.Count -ne 1 -or ("smtp:$smtp" -notin @($resolved[0].EmailAddresses | ForEach-Object { [string]$_ }) -and [string]$resolved[0].PrimarySmtpAddress -ine [string]$smtp)) {
                throw "Resource $($reference.Control) recipient $smtp ($field) is unresolved/ambiguous; no policy is broadened."
            }
        }
    }
    if (@($settings.ResourceCustom).Count) {
        $resourceConfig = Get-ResourceConfig -ErrorAction Stop
        if (-not $resourceConfig -or -not $resourceConfig.PSObject.Properties['ResourcePropertySchema']) {
            throw 'Cloud ResourcePropertySchema is not exposed by Get-ResourceConfig; custom resource features cannot be verified. Check cmdlet access/support before Convert. The resource is not created and tenant-wide schema is not changed.'
        }
        $schema = @($resourceConfig.ResourcePropertySchema | ForEach-Object { [string]$_ })
        foreach ($property in @($settings.ResourceCustom)) {
            if ("$($Record.kind)/$property" -notin $schema) { throw "Resource custom property $($Record.kind)/$property is absent from the cloud resource schema; configure it explicitly before Convert. The tool does not alter tenant-wide resource configuration." }
        }
    }
}

function Get-PcmResourceCloudRoomList {
    param([object]$List, [AllowNull()][object]$Intent, [switch]$AllowMissing)
    $record = ConvertTo-PcmRoomListRecord $List
    $inventory = Get-PcmResourceInventory -Records @($record)
    if ($Intent) {
        if ([string]$Intent.sourceGuid -ne [string]$List.SourceGuid -or $Intent.operation -notin @('Create','Reuse')) { throw 'Invalid room-list journal identity/operation.' }
        if ($Intent.operation -eq 'Create') {
            if ([string]$Intent.token -notmatch '^[a-f0-9]{32}$' -or [string]$Intent.creationName -cnotmatch ("^PRA-L-[a-f0-9]{{8}}-[a-f0-9]{{8}}-{0}$" -f $Intent.token)) { throw 'Room-list journal ownership token/name is invalid; no group is adopted.' }
            $groups = @(Get-DistributionGroup -Filter ("Name -eq '{0}'" -f $Intent.creationName) -ResultSize Unlimited -ErrorAction Stop)
            if ($groups.Count -gt 1) { throw 'Ambiguous room-list ownership; no group is adopted.' }
            if ($groups.Count -and -not $Intent.creationAttempted) { throw 'Room-list creation was not journaled as attempted; no pre-existing group is adopted.' }
            if ($groups.Count) {
                $candidate = $groups[0]
                $marker = [string](Get-PcmValue $candidate 'CustomAttribute15' '')
                if ([string]$candidate.Name -cne [string]$Intent.creationName -or
                    ($marker -cne [string]$Intent.token -and ($marker -or ($Intent.cloudGuid -and -not (Get-PcmValue $Intent 'markerPending' $false)))) -or [string]$candidate.Guid -eq [string]$List.SourceGuid) { throw 'Room-list ownership differs from the journal; nothing is changed.' }
            }
        } else { $groups = @(Get-DistributionGroup -Identity ([string]$Intent.cloudGuid) -ErrorAction Stop) }
    } else {
        $groups = @(Get-DistributionGroup -Filter ("PrimarySmtpAddress -eq '{0}'" -f ([string]$List.PrimarySmtpAddress).Replace("'", "''")) -ResultSize Unlimited -ErrorAction Stop)
    }
    if ($groups.Count -eq 0) {
        if ($Intent -and $Intent.cloudGuid) { throw 'Journal-owned room list is missing; never replace its identity.' }
        $collisions = @(Get-PcmResourceCollisions -Record $record -Inventory $inventory)
        if ($collisions.Count) { throw ("Room-list collision: " + ($collisions -join '; ')) }
        if ($AllowMissing) { return $null }
        throw "Room list $($List.PrimarySmtpAddress) is missing; resume Convert to provision it before its rooms."
    }
    if ($groups.Count -ne 1) { throw "Room list $($List.PrimarySmtpAddress) is ambiguous; no group is adopted." }
    $group = $groups[0]
    $pendingCreation = $Intent -and $Intent.operation -eq 'Create' -and $Intent.status -ne 'Done'
    if ([string](Get-PcmValue $group 'RecipientTypeDetails' '') -ne 'RoomList' -or
        -not $group.PSObject.Properties['IsDirSynced'] -or $group.IsDirSynced -ne $false -or
        [string](Get-PcmValue $group 'PrimarySmtpAddress' '') -ine [string]$List.PrimarySmtpAddress -or
        [string](Get-PcmValue $group 'Alias' '') -ine [string]$List.Alias -or
        -not [string](Get-PcmValue $group 'Guid' '') -or [string]$group.Guid -eq [guid]::Empty.ToString() -or
        -not $group.PSObject.Properties['HiddenFromAddressListsEnabled'] -or
        $group.HiddenFromAddressListsEnabled -isnot [bool] -or
        (-not $pendingCreation -and [bool]$group.HiddenFromAddressListsEnabled -ne [bool]$List.HiddenFromAddressListsEnabled)) {
        throw "Room list $($List.PrimarySmtpAddress) must be a cloud-only RoomList with the original primary address, alias and visibility; no unrelated group is converted or taken over."
    }
    if ($Intent -and (($Intent.cloudGuid -and [string]$Intent.cloudGuid -ne [string]$group.Guid) -or
            ($Intent.cloudId -and [string]$Intent.cloudId -ne [string](Get-PcmValue $group 'ExternalDirectoryObjectId' '')))) { throw 'Room-list immutable identity differs from the journal; nothing is changed.' }
    $collisions = @(Get-PcmResourceCollisions -Record $record -Inventory $inventory -OwnedContactId ([string]$group.Guid))
    if ($collisions.Count) { throw ("Room-list collision: " + ($collisions -join '; ')) }
    return $group
}

function Test-PcmRoomListCreationReadiness {
    param([object]$List)
    Test-PcmRoomListMetadata $List
    foreach ($spec in @(
            @{ Name = 'New-DistributionGroup'; Fields = @('Name','Alias','DisplayName','PrimarySmtpAddress','RoomList','IgnoreNamingPolicy','MemberJoinRestriction','MemberDepartRestriction','RequireSenderAuthenticationEnabled') },
            @{ Name = 'Set-DistributionGroup'; Fields = @('EmailAddresses','HiddenFromAddressListsEnabled','CustomAttribute15','BypassSecurityGroupManagerCheck','MemberJoinRestriction','MemberDepartRestriction') + $script:ResourceDeliveryFields })) {
        $command = Get-Command -Name $spec.Name -ErrorAction Stop
        foreach ($field in $spec.Fields) {
            if (-not $command.Parameters.ContainsKey($field)) { throw "$($spec.Name) -$field is not exposed; check room-list creation RBAC before Convert." }
        }
    }
    foreach ($field in $script:ResourceDeliveryRecipients) { $null = Resolve-PcmResourceRecipients -Addresses $List.Delivery.$field -Control 'Room-list delivery' }
}

function Set-PcmRoomListConfiguration {
    param([object]$List, [object]$Intent, [object]$Group)
    if ($Intent.operation -ne 'Create') { throw 'Existing reused room lists are not reconfigured.' }
    $parameters = @{ Identity = [string]$Group.Guid; EmailAddresses = @(Get-PcmResourceAddresses (ConvertTo-PcmRoomListRecord $List))
        HiddenFromAddressListsEnabled = [bool]$List.HiddenFromAddressListsEnabled; CustomAttribute15 = [string]$Intent.token
        MemberJoinRestriction = $List.MemberJoinRestriction; MemberDepartRestriction = $List.MemberDepartRestriction
        BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' }
    foreach ($field in $script:ResourceDeliveryFields) {
        if ($field -in $script:ResourceDeliveryRecipients) { $parameters[$field] = @(Resolve-PcmResourceRecipients -Addresses $List.Delivery.$field -Control 'Room-list delivery') }
        else { $parameters[$field] = $List.Delivery.$field }
    }
    Confirm-PcmWrite -Target ([string]$Group.Guid) -Operation 'Exchange Online: configure the room list'
    Set-DistributionGroup @parameters
    $actual = Get-DistributionGroup -Identity ([string]$Group.Guid) -ErrorAction Stop
    foreach ($field in @('HiddenFromAddressListsEnabled','CustomAttribute15','MemberJoinRestriction','MemberDepartRestriction') + $script:ResourceDeliveryFields) {
        if (-not $actual.PSObject.Properties[$field]) { throw "Room-list configuration verification failed: $field is not exposed." }
        if ($field -in $script:ResourceDeliveryRecipients) {
            $wanted = @($parameters[$field] | ForEach-Object { [string](Get-Recipient -Identity $_ -ErrorAction Stop).Guid } | Sort-Object)
            $got = @($actual.$field | Where-Object { $null -ne $_ } | ForEach-Object { [string](Get-Recipient -Identity ([string]$_) -ErrorAction Stop).Guid } | Sort-Object)
        } else { $wanted = @([string]$parameters[$field]); $got = @([string]$actual.$field) }
        if (($wanted -join '|') -ine ($got -join '|')) { throw "Room-list configuration verification failed: $field." }
    }
    foreach ($address in $parameters.EmailAddresses) { if ($address -notin @($actual.EmailAddresses | ForEach-Object { [string]$_ })) { throw "Room-list address verification failed: $address." } }
    if ([string]$actual.DisplayName -ine [string]$List.DisplayName) { throw 'Room-list display name verification failed.' }
}

function Resolve-PcmResourceRecipients {
    param([AllowEmptyCollection()][object[]]$Addresses, [string]$Control)
    foreach ($smtp in $Addresses) {
        $resolved = @(Get-Recipient -Identity ([string]$smtp) -ErrorAction Stop)
        if ($resolved.Count -ne 1) { throw "$Control recipient $smtp is unresolved/ambiguous; policy is not broadened." }
        [string]$resolved[0].Identity
    }
}

function Get-PcmResourceGroup {
    param([Parameter(Mandatory)][string]$Id, [object[]]$Skus = @(), [int]$NewUnits = 0, [int]$OtherUnits = 0)
    $group = Invoke-PcmGraph -Uri ("v1.0/groups/{0}?`$select=id,displayName,assignedLicenses,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSecurityIdentifier,securityEnabled,mailEnabled,groupTypes" -f $Id) -AllowNotFound
    if (-not $group) { throw 'Resources.LicensingGroupId does not exist.' }
    if ($null -ne $group.onPremisesSyncEnabled -or (Get-PcmValue $group 'onPremisesLastSyncDateTime' $null) -or
        (Get-PcmValue $group 'onPremisesSecurityIdentifier' '') -or -not $group.securityEnabled -or $group.mailEnabled -or
        @((Get-PcmValue $group 'groupTypes' @()) | Where-Object { $_ -in @('Unified','DynamicMembership') }).Count) {
        throw 'Resources.LicensingGroupId must be a never-synchronised, assigned-membership cloud-only security licensing group (not a mail/Microsoft 365/dynamic group or a formerly AD-synchronised group). A user-group SOA rollback must not remove the active resource licence.'
    }
    $enterprise = @(); $enterpriseAssignments = @()
    foreach ($assignment in @($group.assignedLicenses)) {
        $sku = $Skus | Where-Object { [string]$_.skuId -eq [string]$assignment.skuId } | Select-Object -First 1
        if (-not $sku) { throw 'Resource group assigns an unknown SKU; capacity cannot be verified.' }
        if (($NewUnits + $OtherUnits) -gt 0 -and (Get-PcmFreeUnit $sku) -lt ($NewUnits + $OtherUnits)) { throw "Resource licensing capacity insufficient for $($NewUnits + $OtherUnits) new unit(s) of $($sku.skuPartNumber); resources keep their licence throughout active lifetime." }
        $enabled = @($sku.servicePlans | Where-Object { [string]$_.servicePlanName -eq 'EXCHANGE_S_ENTERPRISE' -and [string]$_.servicePlanId -notin @($assignment.disabledPlans) } |
            ForEach-Object { [string]$_.servicePlanId })
        $enterprise += $enabled
        if ($enabled.Count) { $enterpriseAssignments += [pscustomobject]@{ SkuId = [string]$sku.skuId; PlanIds = $enabled } }
    }
    if (-not $enterprise.Count) { throw 'Resource licensing group has no enabled EXCHANGE_S_ENTERPRISE plan (E3/E5 required for the eDiscovery hold).' }
    return [pscustomobject]@{ Id = $Id; Group = $group; EnterprisePlans = $enterprise; EnterpriseAssignments = $enterpriseAssignments }
}

function Test-PcmResourceLicence {
    param([object]$User, [object]$Group)
    foreach ($assignment in $Group.EnterpriseAssignments) {
        $states = @($User.licenseAssignmentStates | Where-Object { [string]$_.assignedByGroup -eq $Group.Id -and [string]$_.skuId -eq $assignment.SkuId -and
                [string]$_.state -eq 'Active' -and [string]$_.error -in @('', 'None') })
        foreach ($state in $states) {
            $plans = @($User.assignedPlans | Where-Object { [string]$_.capabilityStatus -eq 'Enabled' -and [string]$_.servicePlanId -in $assignment.PlanIds -and [string]$_.servicePlanId -notin @($state.disabledPlans) })
            if ($plans.Count) { return $true }
        }
    }
    return $false
}

function Get-PcmResourceDisabledCalendar {
    param([Parameter(Mandatory)][string]$Identity)
    return @{ Identity = $Identity; ErrorAction = 'Stop'; AutomateProcessing = 'None'
        AllBookInPolicy = $false; AllRequestInPolicy = $false; AllRequestOutOfPolicy = $false
        ForwardRequestsToDelegates = $false; ProcessExternalMeetingMessages = $false
        ResourceDelegates = @(); BookInPolicy = @(); RequestInPolicy = @(); RequestOutOfPolicy = @() }
}

function Assert-PcmResourceCalendar {
    param([hashtable]$Parameters, [string[]]$Fields)
    $actual = Get-CalendarProcessing -Identity $Parameters.Identity -ErrorAction Stop
    foreach ($field in $Fields) {
        if (-not $actual -or -not $actual.PSObject.Properties[$field]) { throw "Resource calendar verification failed: $field is not exposed." }
        if ($field -in $script:ResourceRecipientFields) {
            $wanted = @($Parameters[$field] | ForEach-Object { [string](Get-Recipient -Identity $_ -ErrorAction Stop).Guid } | Sort-Object)
            $got = @($actual.$field | Where-Object { $null -ne $_ } | ForEach-Object { [string](Get-Recipient -Identity ([string]$_) -ErrorAction Stop).Guid } | Sort-Object)
        } else {
            if ($null -eq $actual.$field) { throw "Resource calendar verification failed: $field is unreadable." }
            $wanted = @($Parameters[$field] | ForEach-Object { [string]$_ }); $got = @($actual.$field | ForEach-Object { [string]$_ })
        }
        if (($wanted -join '|') -ine ($got -join '|')) { throw "Resource calendar verification failed: $field." }
    }
}

function Set-PcmResourceProvisioningDisabled {
    param([Parameter(Mandatory)][string]$Identity)
    # Only the just-created, pinned identity or a verified unfinished journal-owned resource reaches this helper.
    $errors = [System.Collections.Generic.List[string]]::new()
    try {
        $calendar = Get-PcmResourceDisabledCalendar -Identity $Identity
        Confirm-PcmWrite -Target $Identity -Operation 'Exchange Online: switch off the booking of the resource'
        Set-CalendarProcessing @calendar
        Assert-PcmResourceCalendar -Parameters $calendar -Fields @($calendar.Keys | Where-Object { $_ -notin @('Identity','ErrorAction') })
    } catch { $errors.Add("Booking disable failed: $($_.Exception.Message)") }
    try {
        Confirm-PcmWrite -Target $Identity -Operation 'Exchange Online: hide the resource from the address lists'
        Set-Mailbox -Identity $Identity -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        if (-not $mailbox.PSObject.Properties['HiddenFromAddressListsEnabled'] -or $mailbox.HiddenFromAddressListsEnabled -ne $true) {
            throw 'Hidden provisioning verification failed.'
        }
    } catch { $errors.Add("Hidden provisioning failed: $($_.Exception.Message)") }
    if ($errors.Count) { throw ($errors -join ' | ') }
}

function Enable-PcmResourceBooking {
    param([object]$Record, [object]$Owned, [hashtable]$Calendar)
    $hidden = [bool](Get-PcmValue $Record 'hidden_from_address_lists' $false)
    Confirm-PcmWrite -Target ([string]$Owned.Id) -Operation 'Exchange Online: set the address list visibility of the resource'
    Set-Mailbox -Identity $Owned.Id -HiddenFromAddressListsEnabled $hidden -ErrorAction Stop
    $mailbox = Get-Mailbox -Identity $Owned.Id -ErrorAction Stop
    if (-not $mailbox.PSObject.Properties['HiddenFromAddressListsEnabled'] -or $mailbox.HiddenFromAddressListsEnabled -ne $hidden) {
        throw 'Resource final hidden verification failed.'
    }
    # Booking activation is the last cloud write/proof: no later mandatory metadata read can leave it open.
    Confirm-PcmWrite -Target ([string]$Owned.Id) -Operation 'Exchange Online: switch on the booking of the resource'
    Set-CalendarProcessing @Calendar
    Assert-PcmResourceCalendar -Parameters $Calendar -Fields @($Calendar.Keys | Where-Object { $_ -notin @('Identity','ErrorAction') })
}

function Set-PcmResourceConfiguration {
    param([object]$Record, [object]$Owned, [object]$Intent)
    $settings = Read-PcmResourceSettings $Record
    $parameters = @{ Identity = $Owned.Id; EmailAddresses = @(Get-PcmResourceAddresses $Record); ResourceCapacity = [int]$settings.Capacity
        ResourceCustom = @($settings.ResourceCustom); HiddenFromAddressListsEnabled = $true; ErrorAction = 'Stop' }
    $parameters[[string]$Intent.markerAttribute] = [string]$Intent.token
    $attributes = [string](Get-PcmValue $Record 'custom_attributes_json' '')
    if ($attributes) {
        $custom = $attributes | ConvertFrom-PcmJson
        foreach ($property in $custom.PSObject.Properties) {
            if ($property.Name -match '^CustomAttribute([1-9]|1[0-5])$' -and $property.Name -ne [string]$Intent.markerAttribute) { $parameters[$property.Name] = [string]$property.Value }
        }
    }
    foreach ($field in $script:ResourceDeliveryFields) {
        if ($field -in $script:ResourceDeliveryRecipients) {
            $parameters[$field] = @(Resolve-PcmResourceRecipients -Addresses $settings.Delivery.$field -Control 'Delivery')
        } else { $parameters[$field] = $settings.Delivery.$field }
    }
    Confirm-PcmWrite -Target ([string]$Owned.Id) -Operation 'Exchange Online: restore the settings of the resource mailbox'
    Set-Mailbox @parameters
    $directory = @{ Identity = $Owned.Id; Confirm = $false; ErrorAction = 'Stop' }
    foreach ($field in $script:ResourceDirectoryFields) {
        $directory[$field] = if ($settings.Directory.$field) { [string]$settings.Directory.$field } else { $null }
    }
    Confirm-PcmWrite -Target ([string]$Owned.Id) -Operation 'Exchange Online: restore the directory properties of the resource'
    Set-User @directory
    $actualUser = Get-User -Identity $Owned.Id -ErrorAction Stop
    foreach ($field in $script:ResourceDirectoryFields) {
        if (-not $actualUser.PSObject.Properties[$field] -or [string]$actualUser.$field -ine [string]$settings.Directory.$field) {
            throw "Resource descriptive property verification failed: $field."
        }
    }
    $calendar = @{ Identity = $Owned.Id; ErrorAction = 'Stop' }
    $fields = @($script:ResourceCalendarFields | Where-Object { $settings.Calendar.PSObject.Properties[$_] })
    foreach ($field in $fields) {
        $value = $settings.Calendar.$field
        if ($field -in $script:ResourceRecipientFields) {
            $value = @(Resolve-PcmResourceRecipients -Addresses @($value) -Control 'Booking')
        }
        $calendar[$field] = $value
    }
    # Stage non-booking settings without enabling the source policy before mandatory proofs and membership.
    $stagedCalendar = $calendar.Clone()
    $disabledCalendar = Get-PcmResourceDisabledCalendar -Identity $Owned.Id
    foreach ($field in $disabledCalendar.Keys) { $stagedCalendar[$field] = $disabledCalendar[$field] }
    Confirm-PcmWrite -Target ([string]$Owned.Id) -Operation 'Exchange Online: restore the booking settings of the resource (booking stays off)'
    Set-CalendarProcessing @stagedCalendar
    $work = @{ Identity = $Owned.Id; ErrorAction = 'Stop' }
    foreach ($field in $script:ResourceWorkFields) {
        if ($field -ne 'WorkingHoursDays') { $work[$field] = $settings.CalendarConfiguration.$field }
    }
    # The portable snapshot key is WorkingHoursDays; Exchange reads and writes WorkDays.
    $work.WorkDays = @(ConvertTo-PcmResourceWorkDays -Days @($settings.CalendarConfiguration.WorkingHoursDays))
    Confirm-PcmWrite -Target ([string]$Owned.Id) -Operation 'Exchange Online: restore the working hours of the resource'
    Set-MailboxCalendarConfiguration @work
    Assert-PcmResourceCalendar -Parameters $stagedCalendar -Fields @($stagedCalendar.Keys | Where-Object { $_ -notin @('Identity','ErrorAction') })
    $actualWork = Get-MailboxCalendarConfiguration -Identity $Owned.Id -ErrorAction Stop
    foreach ($field in $script:ResourceWorkFields) {
        $readName = if ($field -eq 'WorkingHoursDays') { 'WorkDays' } else { $field }
        if (-not $actualWork -or -not $actualWork.PSObject.Properties[$readName]) {
            throw "Resource work-hours verification failed: $readName is not exposed."
        }
        $readValue = $actualWork.$readName
        if ($null -eq $readValue) { throw "Resource work-hours verification failed: $readName is unreadable." }
        $value = if ($field -eq 'WorkingHoursDays') { $work.WorkDays } else { $work[$field] }
        $wanted = @($value | ForEach-Object { [string]$_ } | Sort-Object)
        $got = @(if ($field -in @('WorkingHoursStartTime', 'WorkingHoursEndTime')) { ([timespan]$readValue).ToString('c') }
            elseif ($field -eq 'WorkingHoursDays') { ConvertTo-PcmResourceWorkDays -Days @($readValue) }
            elseif ($field -eq 'WorkingHoursTimeZone' -and $readValue.PSObject.Properties['Id']) { [string]$readValue.Id }
            else { [string]$readValue })
        if (($wanted -join '|') -ine ($got -join '|')) { throw "Resource work-hours verification failed: $field." }
    }
    $mailbox = Get-Mailbox -Identity $Owned.Id -ErrorAction Stop
    foreach ($field in $script:ResourceDeliveryFields) {
        if (-not $mailbox.PSObject.Properties[$field]) { throw "Resource delivery restriction verification failed: $field is not exposed." }
        if ($field -in $script:ResourceDeliveryRecipients) {
            $wanted = @($parameters[$field] | ForEach-Object { [string](Get-Recipient -Identity $_ -ErrorAction Stop).Guid } | Sort-Object)
            $got = @($mailbox.$field | Where-Object { $null -ne $_ } | ForEach-Object { [string](Get-Recipient -Identity ([string]$_) -ErrorAction Stop).Guid } | Sort-Object)
        } else { $wanted = @([string]$parameters[$field]); $got = @([string]$mailbox.$field) }
        if (($wanted -join '|') -ine ($got -join '|')) { throw "Resource delivery restriction verification failed: $field." }
    }
    if ([int]$mailbox.ResourceCapacity -ne [int]$settings.Capacity -or [bool]$mailbox.HiddenFromAddressListsEnabled -ne $parameters.HiddenFromAddressListsEnabled -or
        (@($mailbox.ResourceCustom | Sort-Object) -join '|') -ine (@($settings.ResourceCustom | Sort-Object) -join '|')) { throw 'Resource capacity/custom/hidden verification failed.' }
    foreach ($address in @(Get-PcmResourceAddresses $Record)) {
        if ($address -notin @($mailbox.EmailAddresses | ForEach-Object { [string]$_ })) { throw "Resource address not applied: $address." }
    }
    if ([string]$mailbox.PrimarySmtpAddress -ine [string]$Record.primary_smtp_address) { throw 'Resource primary SMTP verification failed.' }
    return $calendar
}

function Set-PcmResourceRoomListMembership {
    param([object]$Record, [object]$Owned, [hashtable]$RoomListIntents = @{})
    $settings = Read-PcmResourceSettings $Record
    foreach ($list in $settings.RoomLists) {
        $group = Get-PcmResourceCloudRoomList -List $list -Intent $RoomListIntents[[string]$list.SourceGuid]
        $members = @(Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop)
        if (-not @($members | Where-Object { [string](Get-PcmValue $_ 'ExternalDirectoryObjectId' '') -eq [string]$Owned.Id }).Count) {
            Confirm-PcmWrite -Target ([string]$group.Guid) -Operation 'Exchange Online: add the resource to the room list'
            Add-DistributionGroupMember -Identity ([string]$group.Guid) -Member $Owned.Id -BypassSecurityGroupManagerCheck -ErrorAction Stop
        }
        $members = @(Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop)
        if (@($members | Where-Object { [string](Get-PcmValue $_ 'ExternalDirectoryObjectId' '') -eq [string]$Owned.Id }).Count -ne 1) {
            throw "Resource room-list membership verification failed: $($list.PrimarySmtpAddress)."
        }
    }
}

function Get-PcmResourceRoute {
    param([hashtable]$Config, [object]$Record, [switch]$RequireReady)
    if ($RequireReady -and -not [bool](Get-PcmValue $Config.Resources 'OnPremisesReady' $false)) { throw 'Resources.OnPremisesReady must be explicitly true before native resource deletion: rebuild the on-premises resource first.' }
    $domain = [string](Get-PcmValue $Config.Resources 'OnPremisesRoutingDomain' '')
    if ($domain -notmatch '^(?=.{1,253}$)(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,63}$') { throw 'Resources.OnPremisesRoutingDomain must be a valid hybrid routing DNS domain.' }
    $domains = @(Get-PcmResourceAddresses $Record | Where-Object { $_ -match '^smtp:' } | ForEach-Object { ($_ -split '@')[-1] })
    $cloudDomain = [string](Get-PcmValue (Get-PcmValue $Config 'Cloud' @{}) 'Organization' '')
    if ($domain -in $domains -or $domain -ieq $cloudDomain -or $domain -match '\.onmicrosoft\.com$') { throw 'Resource routing domain loops to an original SMTP/cloud domain; use the on-premises hybrid routing domain.' }
    $alias = [string](Get-PcmValue $Record 'alias' '')
    if (-not $alias) { $alias = ([string]$Record.primary_smtp_address -split '@')[0] }
    if ($alias -notmatch '^[a-zA-Z0-9._-]+$') { throw 'Resource alias cannot form a safe routing address.' }
    return "$alias@$domain"
}
#endregion

#region Readiness rules (pure functions) -----------------------------------------------------------

function Get-PcmReadiness {
    <#
    .SYNOPSIS
        Rules for one object of the snapshot: can Convert run, and what will happen. Pure function.
    .PARAMETER Record
        Row of the mailbox table.
    .PARAMETER State
        Output of Get-PcmCloudState (or the same shape in tests).
    .PARAMETER Skus
        subscribedSkus of the tenant (Kiosk mode: the user's licences must contain EXCHANGE_S_DESKLESS).
    .OUTPUTS
        Findings: Level (Ok, Info, Warn, Error), Field (Entra, ExchangeOnline, TeamsStorage, Licence, Holds), Code, Message.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [Parameter(Mandatory)][object]$Record, [Parameter(Mandatory)][object]$State,
        [AllowEmptyCollection()][object[]]$Skus = @())
    $findings = [System.Collections.Generic.List[object]]::new()
    $add = { param($Level, $Field, $Code, $Message) [void]$findings.Add([pscustomobject]@{ Level = $Level; Field = $Field; Code = $Code; Message = $Message }) }
    $isUser = [string]$Record.kind -eq 'User'
    if ((Test-PcmNativeResource -Config $Config -Record $Record) -or (Get-PcmValue $State 'NativeResource' $false)) {
        try { $null = Read-PcmResourceSettings -Record $Record }
        catch { & $add 'Error' 'ExchangeOnline' 'RESOURCE_COLLECT_REQUIRED' $_.Exception.Message }
        if (-not $Config.Scope.ConvertRooms -and -not (Get-PcmValue $State 'JournalNative' $false)) { & $add 'Error' 'ExchangeOnline' 'KIND_NOT_SUPPORTED' 'CloudResource also requires Scope.ConvertRooms = true.' }
        foreach ($collision in @((Get-PcmValue $State 'Collisions' @()) | ForEach-Object { $_ })) { & $add 'Error' 'Entra' 'RESOURCE_COLLISION' ([string]$collision) }
        if (-not (Get-PcmValue $State 'AbsenceVerified' $false)) { & $add 'Error' 'Entra' 'RESOURCE_ABSENCE_UNVERIFIED' 'Cloud identities and all Exchange recipients must be checked successfully; Graph absence alone is insufficient.' }
        # A missing or unknown mode never defaults to the unprotected lifecycle.
        $mode = Get-PcmValue $State 'ProtectionMode' ''
        if ($mode -isnot [string] -or $mode -cnotin @('None', 'CaseHold')) {
            & $add 'Error' 'Holds' 'RESOURCE_PROTECTION_UNKNOWN' "Native resource protection mode '$mode' is missing or unknown; no lifecycle is assumed and nothing is changed."
            return $findings.ToArray()
        }
        if ($mode -eq 'CaseHold' -and [string](Get-PcmValue (Get-PcmValue $Config 'Cloud' @{}) 'DefaultUsageLocation' '') -notmatch '^[A-Z]{2}$') {
            & $add 'Error' 'Licence' 'NO_USAGE_LOCATION' 'Cloud.DefaultUsageLocation must be a two-letter country for the new native resource licence.'
        }
        if ($mode -eq 'CaseHold') {
            & $add 'Info' 'ExchangeOnline' 'RESOURCE_NATIVE' 'Legacy protected native journal: enterprise licence/case hold and inactive-copy recovery retained.'
        } else {
            & $add 'Info' 'ExchangeOnline' 'RESOURCE_NATIVE' 'Create an independent RoomMailbox/EquipmentMailbox with disabled sign-in, no dedicated licence (standard resources up to 50 GB), no hold. No old on-premises content is restored.'
            & $add 'Warn' 'Holds' 'RESOURCE_UNPROTECTED' 'Recover has no hold and no inactive copy: cloud-created reservations/messages are NOT preserved or merged. Export/reconcile first; deletion requires explicit Resources.AllowUnprotectedDeletion = $true.'
        }
        return $findings.ToArray()
    }
    if ([string]$Record.kind -in @('Room', 'Equipment')) {
        if (-not $Config.Scope.ConvertRooms) {
            & $add 'Error' 'ExchangeOnline' 'KIND_NOT_SUPPORTED' ("{0} mailbox: not converted (Scope.ConvertRooms = `$false). It stays unreachable during the disaster; recreate it in Exchange Online by hand if needed, or set Scope.ConvertRooms to convert it like a shared mailbox." -f $Record.kind)
        } else {
            & $add 'Info' 'ExchangeOnline' 'ROOM_CONVERT_ENABLED' ("{0} mailbox: converted like a shared mailbox (Scope.ConvertRooms). Booking settings (capacity, auto-accept, booking policies) are not collected or recreated by this tool; the on-premises object is authoritative and comes back unchanged at Recover." -f $Record.kind)
        }
    }
    $user = $State.User
    if (-not $user) {
        & $add 'Error' 'Entra' 'ENTRA_NOT_FOUND' 'No Entra ID object (neither by onPremisesImmutableId nor by UPN): the object must be synchronised before the disaster.'
        return $findings.ToArray()
    }
    if ($State.MatchedBy -eq 'UPN') { & $add 'Warn' 'Entra' 'ENTRA_MATCHED_BY_UPN' "Found by UPN only: onPremisesImmutableId differs from the objectGUID ($($user.onPremisesImmutableId)); check the source anchor." }
    if ($State.IsCloudManaged -eq $true) { & $add 'Warn' 'Entra' 'ALREADY_CLOUD_MANAGED' 'Source of authority already in the cloud (isCloudManaged): Convert already started for this object.' }
    elseif (-not $user.onPremisesSyncEnabled) { & $add 'Error' 'Entra' 'ENTRA_NOT_SYNCED' 'Cloud-only object: not synchronised from AD, outside the scenario.' }
    else { & $add 'Ok' 'Entra' 'ENTRA_SYNCED' 'Synchronised from AD.' }
    $errors = @($user.serviceProvisioningErrors | Where-Object { $_ -and -not $_.isResolved })
    if ($errors.Count) { & $add 'Warn' 'Entra' 'PROVISIONING_ERROR' ('{0} unresolved provisioning error(s) in Entra ID (serviceProvisioningErrors).' -f $errors.Count) }

    $type = [string]$State.RecipientType
    $onPremGuid = [string]$Record.exchange_guid
    if (-not $type) { & $add 'Error' 'ExchangeOnline' 'EXO_NOT_FOUND' 'No Exchange Online recipient: the on-premises mailbox is not represented in Exchange Online (Exchange hybrid / sync).' }
    elseif ($type -eq 'MailUser') {
        if ([string]$State.ExchangeGuid -eq $onPremGuid) { & $add 'Ok' 'ExchangeOnline' 'EXO_MAILUSER' 'MailUser with the on-premises ExchangeGuid.' }
        elseif ([string]$State.ExchangeGuid -eq [guid]::Empty.ToString()) { & $add 'Warn' 'ExchangeOnline' 'GUID_CLEARED' 'MailUser with an empty ExchangeGuid: Convert already started for this object.' }
        else {
            & $add 'Warn' 'ExchangeOnline' 'GUID_MISMATCH' ("MailUser ExchangeGuid {0} differs from the on-premises mailbox {1}." -f $State.ExchangeGuid, $onPremGuid)
            if ($State.MatchedBy -eq 'UPN') { & $add 'Error' 'Entra' 'ENTRA_WRONG_OBJECT' ("The Entra ID account with this UPN is probably not this person: it was found by UPN only (its onPremisesImmutableId is '{0}', not the one of the snapshot) and its mailbox GUID {1} differs from the on-premises mailbox {2}. Nothing is changed on it. Fix the source anchor or the UPN in Entra ID, or leave this object out." -f $user.onPremisesImmutableId, $State.ExchangeGuid, $onPremGuid) }
        }
    } elseif ($type -match 'Mailbox$') {
        & $add 'Error' 'ExchangeOnline' 'ALREADY_CLOUD_MAILBOX' ("Already a cloud mailbox ($type): converted already, or migrated to Exchange Online (then it is outside the scenario).")
    } else { & $add 'Error' 'ExchangeOnline' 'EXO_UNEXPECTED_TYPE' "Unexpected Exchange Online recipient type $type." }

    $component = @($State.Locations | Where-Object { $_.Type -eq 'ComponentShared' })
    if ($isUser) {
        if ($component.Count) { & $add 'Info' 'TeamsStorage' 'TEAMS_STORAGE' 'ComponentShared storage present (Teams storage, or the cloud mailbox of an earlier conversion): Convert promotes it to the primary mailbox, its content (Teams chat copies included) is kept.' }
        else { & $add 'Info' 'TeamsStorage' 'NO_TEAMS_STORAGE' 'No Teams storage: Convert creates a new, empty cloud mailbox.' }
    } elseif ($component.Count) {
        & $add 'Info' 'TeamsStorage' 'COMPONENT_SHARED' 'ComponentShared storage present on this object (created by Microsoft 365, or the mailbox of an earlier conversion of the same object): the temporary licence promotes it to the shared mailbox, with its content. An inactive mailbox of a previous Recover is never linked to the recreated object.'
    }

    $plans = @($user.assignedPlans | Where-Object { $_ -and [string]$_.capabilityStatus -eq 'Enabled' })
    # With the SKUs: the mailbox plans by ID (Foundation and other Exchange-backed plans give no mailbox); without: service exchange.
    $mailboxPlans = @(Get-PcmMailboxPlanId -Skus $Skus)
    $plans = @(if ($mailboxPlans.Count) { $plans | Where-Object { [string]$_.servicePlanId -in $mailboxPlans } } else { $plans | Where-Object { [string]$_.service -eq 'exchange' } })
    if ($plans.Count) { & $add 'Error' 'Licence' 'EXCHANGE_PLAN_PRESENT' 'An Exchange Online mailbox plan is already enabled on this object: clearing the ExchangeGuid would not create the mailbox while the plan stays, so Convert refuses it. Remove that plan from the object before the disaster, or exclude it from the scope.' }
    if (-not $user.usageLocation) {
        if ($Config.Cloud.DefaultUsageLocation) { & $add 'Info' 'Licence' 'NO_USAGE_LOCATION' ('usageLocation empty: Convert sets it to {0} (Cloud.DefaultUsageLocation) before the licence.' -f $Config.Cloud.DefaultUsageLocation) }
        else { & $add 'Error' 'Licence' 'NO_USAGE_LOCATION' 'usageLocation empty and Cloud.DefaultUsageLocation not set: a licence needs a country. Set Cloud.DefaultUsageLocation (two letters) or the usageLocation of the object.' }
    }
    if ($isUser -and $Config.Licensing.Users.Mode -eq 'Kiosk') {
        $skuById = @{}; foreach ($sku in $Skus) { $skuById[[string]$sku.skuId] = $sku }
        $kiosk = @($user.assignedLicenses | Where-Object { $skuById.ContainsKey([string]$_.skuId) -and @($skuById[[string]$_.skuId].servicePlans | Where-Object { [string]$_.servicePlanName -eq 'EXCHANGE_S_DESKLESS' }).Count })
        if ($kiosk.Count) { & $add 'Ok' 'Licence' 'KIOSK_AVAILABLE' ('Exchange Kiosk plan available in ' + (($kiosk | ForEach-Object { $skuById[[string]$_.skuId].skuPartNumber }) -join ', ') + '.') }
        else { & $add 'Error' 'Licence' 'NO_KIOSK_PLAN' 'Kiosk mode: no licence of this user contains the Exchange Kiosk plan (EXCHANGE_S_DESKLESS).' }
    }

    $holds = @($State.Holds | Where-Object { $_ })
    if ($holds.Count) {
        $text = 'Hold(s) on the object: {0}. They follow the cloud mailbox; a user hold must be lifted during Recover.' -f ($holds -join ', ')
        & $add $(if ($isUser) { 'Warn' } else { 'Info' }) 'Holds' 'HOLD_PRESENT' $text
    }
    $tag = [string]$Config.Retention.TagValue
    if ($tag -and [string]$State.TagValue -eq $tag) { & $add 'Warn' 'Holds' 'TAG_ALREADY_SET' ("{0} is already {1} in Exchange Online." -f $Config.Retention.TagAttribute, $tag) }
    if ($isUser -and $tag -and [string]$State.TagValue -eq $tag) { & $add 'Warn' 'Holds' 'USER_TAGGED' 'A user must not carry the retention tag: a retention policy hold would block his Recover.' }
    return $findings.ToArray()
}

function Get-PcmTrusteeFinding {
    <#
    .SYNOPSIS
        Shared mailbox permissions of the snapshot: is every trustee (or expanded group member) known in
        Exchange Online? Uses $Lookup (scriptblock: identity -> $true/$false) so that tests need no cloud.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Record, [AllowEmptyCollection()][object[]]$Permissions = @(), [AllowEmptyCollection()][object[]]$Members = @(),
        [Parameter(Mandatory)][scriptblock]$Lookup)
    $findings = [System.Collections.Generic.List[object]]::new()
    $mine = @($Permissions | Where-Object { [string]$_.mailbox_guid -eq [string]$Record.object_guid })
    if (-not $mine.Count) { return $findings.ToArray() }
    $counts = @{ FullAccess = 0; SendAs = 0; SendOnBehalf = 0 }
    $unknown = [System.Collections.Generic.List[string]]::new()
    foreach ($permission in $mine) {
        $counts[[string]$permission.access_right]++
        $targets = @()
        if ([string]$permission.trustee_kind -eq 'Group') {
            $targets = @($Members | Where-Object { [string]$_.group_guid -eq [string]$permission.trustee_guid -and [string]$_.member_kind -eq 'User' } |
                ForEach-Object { if ($_.member_upn) { [string]$_.member_upn } else { [string]$_.member_smtp } })
            if (-not $targets.Count) { [void]$unknown.Add(('{0} {1} (group without user member)' -f $permission.access_right, $permission.trustee)) }
        } elseif (-not $permission.resolved) {
            [void]$unknown.Add(('{0} {1} (not resolved at Collect)' -f $permission.access_right, $permission.trustee))
        } else {
            $targets = @($(if ($permission.trustee_upn) { [string]$permission.trustee_upn } else { [string]$permission.trustee_smtp }))
        }
        foreach ($target in $targets) {
            if ($target -and -not (& $Lookup $target)) { [void]$unknown.Add(('{0} {1} (not found in Exchange Online)' -f $permission.access_right, $target)) }
        }
    }
    $summary = 'FullAccess {0} · SendAs {1} · SendOnBehalf {2}' -f $counts.FullAccess, $counts.SendAs, $counts.SendOnBehalf
    if ($unknown.Count) { [void]$findings.Add([pscustomobject]@{ Level = 'Warn'; Field = 'Permissions'; Code = 'TRUSTEE_UNKNOWN'; Message = ('{0}; not reproducible: {1}' -f $summary, ($unknown -join '; ')) }) }
    else { [void]$findings.Add([pscustomobject]@{ Level = 'Ok'; Field = 'Permissions'; Code = 'TRUSTEES_OK'; Message = $summary }) }
    return $findings.ToArray()
}
#endregion

#region Changes (Convert, Recover) -----------------------------------------------------------------
# Every function changes one thing and is safe to run again: the entry script reads the state first and the
# journal records each step, so an interrupted batch is resumed with the same -Batch.

function Wait-PcmCondition {
    <#
    .SYNOPSIS
        Calls -Test every IntervalSeconds until it returns a true value or TimeoutSeconds is reached. Each wait is announced to the window (event progress: Phase, attempt, maximum of attempts, seconds to the next check).
    .OUTPUTS
        @{ Ok = bool; Seconds = elapsed; Last = last value of -Test }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Test, [Parameter(Mandatory)][int]$TimeoutSeconds, [int]$IntervalSeconds = 20, [string]$Phase = '')
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $attempt = 0
    $maxAttempts = [int][Math]::Ceiling($TimeoutSeconds / [double][Math]::Max(1, $IntervalSeconds)) + 1
    do {
        $last = & $Test
        if ($last) { return [pscustomobject]@{ Ok = $true; Seconds = [int]$clock.Elapsed.TotalSeconds; Last = $last } }
        if ($clock.Elapsed.TotalSeconds -ge $TimeoutSeconds -or $attempt -ge $maxAttempts - 1) { break }
        $attempt++
        $pause = [Math]::Max(1, [Math]::Min($IntervalSeconds, $TimeoutSeconds - [int]$clock.Elapsed.TotalSeconds))
        Write-PcmEvent 'progress' @{ phase = $(if ($Phase) { $Phase } else { 'Waiting for Microsoft 365' }); attempt = $attempt; maxAttempts = [Math]::Max($attempt, $maxAttempts); nextInSeconds = $pause }
        Start-Sleep -Seconds $pause
    } while ($true)
    return [pscustomobject]@{ Ok = $false; Seconds = [int]$clock.Elapsed.TotalSeconds; Last = $last }
}

function Get-PcmExoState {
    <# Exchange Online view of one object: recipient type, ExchangeGuid, synchronised or not, holds, tag value. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Identity, [string]$TagAttribute = 'CustomAttribute1')
    $recipient = Get-Recipient -Identity $Identity -ErrorAction SilentlyContinue
    $user = Get-User -Identity $Identity -ErrorAction SilentlyContinue
    $holds = @(); $guid = ''
    if ($recipient) {
        $guid = [string]$recipient.ExchangeGuid
        if ([string]$recipient.RecipientTypeDetails -eq 'MailUser') {
            $mailUser = Get-MailUser -Identity $Identity -ErrorAction SilentlyContinue
            if ($mailUser) { $holds = @($mailUser.InPlaceHolds | ForEach-Object { [string]$_ }); $guid = [string]$mailUser.ExchangeGuid }
        } elseif ([string]$recipient.RecipientTypeDetails -match 'Mailbox$') {
            $mailbox = Get-Mailbox -Identity $Identity -ErrorAction SilentlyContinue
            if ($mailbox) { $holds = @($mailbox.InPlaceHolds | ForEach-Object { [string]$_ }); if ($mailbox.LitigationHoldEnabled) { $holds += 'LitigationHold' } }
        }
    }
    return [pscustomobject]@{
        Type = $(if ($recipient) { [string]$recipient.RecipientTypeDetails } else { '' }); ExchangeGuid = $guid
        IsDirSynced = $(if ($user) { [bool]$user.IsDirSynced } else { $null }); Holds = $holds
        Tag = $(if ($recipient -and $recipient.PSObject.Properties[$TagAttribute]) { [string]$recipient.$TagAttribute } else { '' })
        UserType = $(if ($user) { [string]$user.RecipientTypeDetails } else { '' })
    }
}

function Set-PcmSourceOfAuthority {
    <# Source of authority of a user or a group: cloud (isCloudManaged = true) or AD. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][bool]$Cloud, [ValidateSet('users','groups')][string]$Type = 'users')
    if (-not $PSCmdlet.ShouldProcess("$Type/$Id", "isCloudManaged = $Cloud")) { return }
    $null = Invoke-PcmGraph -Method PATCH -Uri ("v1.0/{0}/{1}/onPremisesSyncBehavior" -f $Type, $Id) -Body @{ isCloudManaged = $Cloud }
}

function Get-PcmSourceOfAuthority {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Id, [ValidateSet('users','groups')][string]$Type = 'users')
    $behavior = Invoke-PcmGraph -Uri ("v1.0/{0}/{1}/onPremisesSyncBehavior?`$select=isCloudManaged" -f $Type, $Id) -AllowNotFound
    if ($behavior) { return [bool]$behavior.isCloudManaged }
    return $null
}

function Set-PcmUsageLocation {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][ValidatePattern('^[A-Z]{2}$')][string]$Country)
    if (-not $PSCmdlet.ShouldProcess($UserId, "usageLocation = $Country")) { return }
    $null = Invoke-PcmGraph -Method PATCH -Uri ("v1.0/users/{0}" -f $UserId) -Body @{ usageLocation = $Country }
}

function Set-PcmGroupMember {
    <# Adds or removes a member of a group; already a member / already removed is not an error. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][ValidateSet('Add','Remove')][string]$Operation)
    if (-not $PSCmdlet.ShouldProcess("group $GroupId", "$Operation member $UserId")) { return }
    try {
        if ($Operation -eq 'Add') { $null = Invoke-PcmGraph -Method POST -Uri ("v1.0/groups/{0}/members/`$ref" -f $GroupId) -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" } }
        else { $null = Invoke-PcmGraph -Method DELETE -Uri ("v1.0/groups/{0}/members/{1}/`$ref" -f $GroupId, $UserId) -AllowNotFound }
    } catch {
        if ($Operation -eq 'Add' -and "$_" -match 'already exist') { return }
        throw
    }
}

function Test-PcmGroupMember {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][string]$UserId)
    $result = Invoke-PcmGraph -Method POST -Uri ("v1.0/users/{0}/checkMemberGroups" -f $UserId) -Body @{ groupIds = @($GroupId) }
    return @($result.value) -contains $GroupId
}

function Get-PcmDisabledPlanForExchangeOnly {
    <# disabledPlans that leave only the Exchange plan(s) of a SKU enabled (licence given for the mailbox only). #>
    param([Parameter(Mandatory)][object]$Sku)
    return @($Sku.servicePlans | Where-Object { [string]$_.servicePlanName -notmatch $script:ExchangePlanPattern -and [string]$_.appliesTo -eq 'User' } | ForEach-Object { [string]$_.servicePlanId })
}

function Set-PcmLicence {
    <# Assigns (with disabled plans) or removes one SKU directly on a user. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$SkuId, [Parameter(Mandatory)][ValidateSet('Add','Remove')][string]$Operation,
        [string[]]$DisabledPlans = @())
    if (-not $PSCmdlet.ShouldProcess($UserId, "$Operation licence $SkuId")) { return }
    $body = if ($Operation -eq 'Add') { @{ addLicenses = @(@{ skuId = $SkuId; disabledPlans = @($DisabledPlans) }); removeLicenses = @() } }
        else { @{ addLicenses = @(); removeLicenses = @($SkuId) } }
    # A first assignment right after the usageLocation change can be refused (400) for a few seconds.
    for ($attempt = 1; ; $attempt++) {
        try { $null = Invoke-PcmGraph -Method POST -Uri ("v1.0/users/{0}/assignLicense" -f $UserId) -Body $body; return }
        catch { if ($Operation -ne 'Add' -or $attempt -ge 4 -or "$($_.Exception.Message)" -notmatch '\(400\)|BadRequest') { throw }; Start-Sleep -Seconds 15 }
    }
}

function Get-PcmUserLicence {
    <# Licence assignment states of a user (direct and by group). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)
    $user = Invoke-PcmGraph -Uri ("v1.0/users/{0}?`$select=id,usageLocation,assignedLicenses,licenseAssignmentStates,onPremisesSyncEnabled" -f $UserId)
    return $user
}

function Get-PcmDirectLicence {
    <# Direct licence assignments of a user (not the ones inherited from a group): skuId and sorted disabledPlans. Pure function. #>
    param([AllowNull()][object]$Licence)
    if (-not $Licence -or -not $Licence.PSObject.Properties['licenseAssignmentStates']) { return @() }
    return @($Licence.licenseAssignmentStates | Where-Object { $_ -and -not $_.assignedByGroup } | ForEach-Object {
            [pscustomobject]@{ skuId = [string]$_.skuId; disabledPlans = @($_.disabledPlans | ForEach-Object { [string]$_ } | Sort-Object) } })
}

function Enable-PcmLicencePlan {
    <#
    .SYNOPSIS
        Enables plans of one SKU on a user without touching the other plans: in the existing direct assignment of the
        SKU, or by a new direct assignment with -DisabledPlans (an assignment inherited from a group cannot be changed
        per user; the direct one adds to it).
    .OUTPUTS
        'Present' (already enabled), 'Enabled' (existing direct assignment changed) or 'Assigned' (new direct assignment).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$SkuId, [Parameter(Mandatory)][string[]]$PlanIds, [string[]]$DisabledPlans = @())
    $current = Get-PcmDirectLicence (Get-PcmUserLicence -UserId $UserId) | Where-Object { $_.skuId -eq $SkuId } | Select-Object -First 1
    if ($current) {
        $keep = @($current.disabledPlans | Where-Object { $_ -notin $PlanIds })
        if ($keep.Count -eq @($current.disabledPlans).Count) { return 'Present' }
        Set-PcmLicence -UserId $UserId -SkuId $SkuId -Operation Add -DisabledPlans $keep
        return 'Enabled'
    }
    Set-PcmLicence -UserId $UserId -SkuId $SkuId -Operation Add -DisabledPlans $DisabledPlans
    return 'Assigned'
}

function Restore-PcmLicence {
    <#
    .SYNOPSIS
        Puts the direct assignment of one SKU back as it was before Convert: the same disabled plans when the user had
        it ($Original = direct assignments recorded by Convert), otherwise the assignment added by Convert is removed.
    .OUTPUTS
        'Unchanged', 'Restored', 'Removed' or 'Absent'.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$SkuId, [AllowEmptyCollection()][object[]]$Original = @())
    $before = @($Original | Where-Object { $_ -and [string]$_.skuId -eq $SkuId }) | Select-Object -First 1
    $current = Get-PcmDirectLicence (Get-PcmUserLicence -UserId $UserId) | Where-Object { $_.skuId -eq $SkuId } | Select-Object -First 1
    if ($before) {
        $wanted = @($before.disabledPlans | ForEach-Object { [string]$_ } | Sort-Object)
        if ($current -and (@($current.disabledPlans) -join ',') -eq ($wanted -join ',')) { return 'Unchanged' }
        Set-PcmLicence -UserId $UserId -SkuId $SkuId -Operation Add -DisabledPlans $wanted
        return 'Restored'
    }
    if ($current) { Set-PcmLicence -UserId $UserId -SkuId $SkuId -Operation Remove; return 'Removed' }
    return 'Absent'
}

function Remove-PcmIdentity {
    <# Deletes a user from Entra ID, then permanently from the recycle bin (shared mailbox Recover). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][ValidateSet('Delete','Purge')][string]$Operation)
    if (-not $PSCmdlet.ShouldProcess($UserId, "$Operation identity")) { return }
    if ($Operation -eq 'Delete') { $null = Invoke-PcmGraph -Method DELETE -Uri ("v1.0/users/{0}" -f $UserId) -AllowNotFound }
    else { $null = Invoke-PcmGraph -Method DELETE -Uri ("v1.0/directory/deletedItems/{0}" -f $UserId) -AllowNotFound }
}

function Test-PcmDeletedIdentity {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)
    return [bool](Invoke-PcmGraph -Uri ("v1.0/directory/deletedItems/{0}?`$select=id" -f $UserId) -AllowNotFound)
}

function Connect-PcmCompliance {
    <# Security & Compliance (search-only session) for the eDiscovery case hold, case hold cmdlets only. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($Context.ContainsKey('ComplianceConnected') -and $Context['ComplianceConnected']) { return }
    $cloud = $Context.Config.Cloud
    # -CommandName is required: the full session also loads Get-Recipient and Get-User, which hide the Exchange
    # Online cmdlets and return a stale view (a rolled back user still seen as UserMailbox).
    if (Test-PcmSameText (Get-PcmValue $cloud 'SignIn' 'Certificate') 'Interactive') {
        $hint = [string](Get-PcmValue $cloud 'UserPrincipalName' '')
        if (-not $hint -and (Get-Command Get-MgContext -ErrorAction SilentlyContinue)) { $hint = [string](Get-PcmValue (Get-MgContext) 'Account' '') }
        $interactive = @{ ShowBanner = $false; EnableSearchOnlySession = $true; ErrorAction = 'Stop'
            CommandName = @('Get-CaseHoldPolicy', 'Set-CaseHoldPolicy', 'New-CaseHoldPolicy', 'New-CaseHoldRule') }
        if ($hint) { $interactive.UserPrincipalName = $hint }
        Enter-PcmCall -Call 'Connect-IPPSSession' -Seconds ([Math]::Max((Get-PcmCallTimeoutSecond -Config $Context.Config), $script:CallGuardInteractiveSeconds))
        try { Connect-IPPSSession @interactive } finally { Exit-PcmCall }
    } else {
        Enter-PcmCall -Call 'Connect-IPPSSession'
        try {
            Connect-IPPSSession -AppId $cloud.AppId -CertificateThumbprint $cloud.CertificateThumbprint -Organization $cloud.Organization -ShowBanner:$false -EnableSearchOnlySession `
                -CommandName Get-CaseHoldPolicy, Set-CaseHoldPolicy, New-CaseHoldPolicy, New-CaseHoldRule -ErrorAction Stop
        } finally { Exit-PcmCall }
    }
    $Context['ComplianceConnected'] = $true
    Install-PcmCallGuard
}

# An eDiscovery case hold policy holds 1,000 mailboxes at most: the tool uses the configured policy, then <name>-02,
# <name>-03... in the same case (created with a rule that holds everything). Set-CaseHoldPolicy takes many locations
# in one call (12 to 15 s per call, whatever their number).
$script:HoldLocationChunk = 100

function Get-PcmHoldPolicySet {
    <#
    .SYNOPSIS
        The case hold policies of the tool: the configured one and its siblings <name>-NN in the same case, with their
        tag (UniH + GUID) and their locations (Entra object IDs). A sibling being deleted (Mode PendingDeletion, still
        listed for hours) is left out, but its number is not used again.
    .PARAMETER Limit
        Locations per policy (Retention.HoldPolicyLimit, 1,000 = the Microsoft Purview limit).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Policy, [ValidateRange(1, 1000)][int]$Limit = 1000)
    $base = Get-CaseHoldPolicy -Identity $Policy -ErrorAction Stop
    if ([string]$base.Mode -eq 'PendingDeletion') { throw "Case hold policy $Policy (Retention.HoldPolicy) is being deleted: use another policy." }
    $pattern = '^{0}-(\d{{2,3}})$' -f [regex]::Escape($Policy)
    $taken = [System.Collections.Generic.List[int]]::new()
    $names = @($Policy)
    foreach ($sibling in @(Get-CaseHoldPolicy -Case ([string]$base.CaseId) -ErrorAction Stop | Where-Object { [string]$_.Name -match $pattern } | Sort-Object { [int]([string]$_.Name -replace '^.*-', '') })) {
        $taken.Add([int]([string]$sibling.Name -replace '^.*-', ''))
        if ([string]$sibling.Mode -ne 'PendingDeletion') { $names += [string]$sibling.Name }
    }
    $policies = [System.Collections.Generic.List[object]]::new()
    $tags = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $names) {
        $item = if ($name -eq $Policy) { $base } else { Get-CaseHoldPolicy -Identity $name -ErrorAction Stop }
        $locations = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($location in @($item.ExchangeLocation)) { foreach ($value in @([string]$location.ImmutableIdentity, [string]$location.Name)) { if ($value) { [void]$locations.Add($value) } } }
        $tag = 'UniH{0}' -f $item.Guid
        [void]$tags.Add($tag)
        $policies.Add([pscustomobject]@{ Name = [string]$item.Name; Tag = $tag; Locations = $locations; Count = @($item.ExchangeLocation).Count })
    }
    return [pscustomobject]@{ Base = $Policy; CaseId = [string]$base.CaseId; Limit = $Limit; Policies = $policies; Tags = $tags; Taken = $taken }
}

function Test-PcmToolHold {
    <# Is one of the holds of an object a case hold of the tool? #>
    param([Parameter(Mandatory)][object]$Set, [AllowEmptyCollection()][string[]]$Holds = @())
    foreach ($hold in $Holds) { if ($hold -and $Set.Tags.Contains($hold)) { return $true } }
    return $false
}

function Invoke-PcmHoldChange {
    <#
    One Set-CaseHoldPolicy call for many locations. 'failed to be deployed' while the change is recorded
    counts as done for the locations the policy shows afterwards. Returns the IDs that were not changed.
    #>
    param([Parameter(Mandatory)][string]$Policy, [Parameter(Mandatory)][string[]]$Id, [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Operation)
    Confirm-PcmWrite -Target $Policy -Operation ('Case hold: {0} {1} mailbox location(s)' -f $Operation.ToLowerInvariant(), $Id.Count)
    try {
        if ($Operation -eq 'Add') { Set-CaseHoldPolicy -Identity $Policy -AddExchangeLocation $Id -ErrorAction Stop }
        else { Set-CaseHoldPolicy -Identity $Policy -RemoveExchangeLocation $Id -ErrorAction Stop }
        return @()
    } catch {
        $message = $_.Exception.Message -replace '\s+', ' '
        if ($message -notmatch 'failed to be deployed|PolicyDeploymentException|already|déjà|not found in|introuvable dans|does not exist in') { throw }
        $now = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($location in @((Get-CaseHoldPolicy -Identity $Policy -ErrorAction Stop).ExchangeLocation)) { foreach ($value in @([string]$location.ImmutableIdentity, [string]$location.Name)) { if ($value) { [void]$now.Add($value) } } }
        $missed = @($Id | Where-Object { ($Operation -eq 'Add') -ne $now.Contains($_) })
        if ($missed.Count -and $message -match 'failed to be deployed|PolicyDeploymentException') { Write-Warning "Case hold policy $Policy`: $Operation recorded, distribution pending ($message)" }
        return $missed
    }
}

function Add-PcmHoldLocationSet {
    <#
    .SYNOPSIS
        Puts many Entra object IDs on the case hold of the tool, in groups of 100 per call: the policies with room first
        (1,000 locations each), then new policies <name>-NN created in the same case. An ID already held is left as it
        is. Returns Policy (ID -> policy name) and Failed (ID -> message).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Set, [AllowEmptyCollection()][string[]]$Id = @())
    $result = [pscustomobject]@{ Policy = @{}; Failed = @{} }
    $todo = [System.Collections.Generic.List[string]]::new()
    foreach ($one in @($Id | Where-Object { $_ } | Get-PcmDistinctText)) {
        $holder = $Set.Policies | Where-Object { $_.Locations.Contains($one) } | Select-Object -First 1
        if ($holder) { $result.Policy[$one] = $holder.Name } else { $todo.Add($one) }
    }
    if (-not $todo.Count -or -not $PSCmdlet.ShouldProcess($Set.Base, "Add $($todo.Count) location(s)")) { return $result }
    $next = 0
    while ($next -lt $todo.Count) {
        $policy = $Set.Policies | Where-Object { $_.Count -lt $Set.Limit } | Select-Object -First 1
        if (-not $policy) {
            $number = ((@($Set.Taken) + 1) | Measure-Object -Maximum).Maximum + 1
            $name = '{0}-{1:00}' -f $Set.Base, $number
            Confirm-PcmWrite -Target $name -Operation 'Case hold: create a new policy (the limit of the previous one is reached)'
            $created = New-CaseHoldPolicy -Name $name -Case $Set.CaseId -Enabled $true -ErrorAction Stop
            $null = New-CaseHoldRule -Name "$name-Rule" -Policy $name -ErrorAction Stop
            $policy = [pscustomobject]@{ Name = $name; Tag = ('UniH{0}' -f $created.Guid); Locations = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); Count = 0 }
            $Set.Policies.Add($policy); [void]$Set.Tags.Add($policy.Tag); $Set.Taken.Add($number)
        }
        $size = [Math]::Min($script:HoldLocationChunk, [Math]::Min($Set.Limit - $policy.Count, $todo.Count - $next))
        $chunk = [string[]]@($todo.GetRange($next, $size))
        $next += $size
        $missed = @()
        try { $missed = @(Invoke-PcmHoldChange -Policy $policy.Name -Id $chunk -Operation Add) }
        catch { foreach ($one in $chunk) { $result.Failed[$one] = "case hold policy $($policy.Name): $($_.Exception.Message -replace '\s+', ' ')" }; continue }
        foreach ($one in $chunk) {
            if ($missed -contains $one) { $result.Failed[$one] = "not added to the case hold policy $($policy.Name)"; continue }
            [void]$policy.Locations.Add($one); $policy.Count++; $result.Policy[$one] = $policy.Name
        }
    }
    return $result
}

function Remove-PcmHoldLocationSet {
    <# Takes many Entra object IDs off the policy of the tool that holds each one (one call per policy and 100 IDs). Returns Failed (ID -> message). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Set, [AllowEmptyCollection()][string[]]$Id = @())
    $result = [pscustomobject]@{ Failed = @{} }
    foreach ($policy in @($Set.Policies)) {
        $mine = @($Id | Where-Object { $_ -and $policy.Locations.Contains($_) } | Get-PcmDistinctText)
        if (-not $mine.Count -or -not $PSCmdlet.ShouldProcess($policy.Name, "Remove $($mine.Count) location(s)")) { continue }
        for ($start = 0; $start -lt $mine.Count; $start += $script:HoldLocationChunk) {
            $chunk = [string[]]@($mine[$start..([Math]::Min($mine.Count, $start + $script:HoldLocationChunk) - 1)])
            $missed = @()
            try { $missed = @(Invoke-PcmHoldChange -Policy $policy.Name -Id $chunk -Operation Remove) }
            catch { foreach ($one in $chunk) { $result.Failed[$one] = "case hold policy $($policy.Name): $($_.Exception.Message -replace '\s+', ' ')" }; continue }
            foreach ($one in $chunk) {
                if ($missed -contains $one) { $result.Failed[$one] = "not removed from the case hold policy $($policy.Name)"; continue }
                [void]$policy.Locations.Remove($one); $policy.Count = [Math]::Max(0, $policy.Count - 1)
            }
        }
    }
    return $result
}

function Get-PcmKioskAssignment {
    <# Kiosk mode: the SKU of the user that contains EXCHANGE_S_DESKLESS and the disabled plans of a direct assignment that only adds Kiosk. #>
    param([Parameter(Mandatory)][object]$Licence, [Parameter(Mandatory)][object[]]$Skus)
    foreach ($state in @($Licence.licenseAssignmentStates)) {
        $sku = $Skus | Where-Object { [string]$_.skuId -eq [string]$state.skuId } | Select-Object -First 1
        if (-not $sku) { continue }
        $kiosk = $sku.servicePlans | Where-Object { [string]$_.servicePlanName -eq 'EXCHANGE_S_DESKLESS' } | Select-Object -First 1
        if (-not $kiosk) { continue }
        $disabled = @($state.disabledPlans | ForEach-Object { [string]$_ } | Where-Object { $_ -ne [string]$kiosk.servicePlanId })
        $direct = @($Licence.licenseAssignmentStates | Where-Object { [string]$_.skuId -eq [string]$sku.skuId -and -not $_.assignedByGroup })
        return [pscustomobject]@{ SkuId = [string]$sku.skuId; SkuName = [string]$sku.skuPartNumber; KioskPlanId = [string]$kiosk.servicePlanId; DisabledPlans = $disabled
            DirectWithKiosk = [bool](@($direct | Where-Object { @($_.disabledPlans) -notcontains [string]$kiosk.servicePlanId }).Count) }
    }
    return $null
}

function Get-PcmSharedGrant {
    <#
    .SYNOPSIS
        Permissions to reproduce for one shared mailbox, from the snapshot: one entry per right and per user
        (groups replaced by their user members). Pure function.
    .OUTPUTS
        Right (FullAccess, SendAs, SendOnBehalf), Trustee (UPN or SMTP), AutoMapping, Source.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Record, [AllowEmptyCollection()][object[]]$Permissions = @(), [AllowEmptyCollection()][object[]]$Members = @())
    $grants = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($permission in @($Permissions | Where-Object { [string]$_.mailbox_guid -eq [string]$Record.object_guid })) {
        $targets = @()
        if ([string]$permission.trustee_kind -eq 'Group') {
            $targets = @($Members | Where-Object { [string]$_.group_guid -eq [string]$permission.trustee_guid -and [string]$_.member_kind -eq 'User' } |
                ForEach-Object { [pscustomobject]@{ Id = $(if ($_.member_upn) { [string]$_.member_upn } else { [string]$_.member_smtp }); Source = "group $($permission.trustee_name)" } })
        } elseif ($permission.resolved -and [string]$permission.trustee_kind -eq 'User') {
            $targets = @([pscustomobject]@{ Id = $(if ($permission.trustee_upn) { [string]$permission.trustee_upn } else { [string]$permission.trustee_smtp }); Source = 'direct' })
        }
        foreach ($target in $targets) {
            if (-not $target.Id) { continue }
            $key = '{0}|{1}' -f $permission.access_right, $target.Id.ToLowerInvariant()
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $automap = if ([string]$permission.access_right -eq 'FullAccess') { $(if ($null -eq $permission.auto_mapping) { $true } else { [bool]$permission.auto_mapping }) } else { $null }
            [void]$grants.Add([pscustomobject]@{ Right = [string]$permission.access_right; Trustee = $target.Id; AutoMapping = $automap; Source = $target.Source })
        }
    }
    return $grants.ToArray()
}

function Grant-PcmSharedPermission {
    <#
    .SYNOPSIS
        Grants one permission on a cloud shared mailbox if it is not there yet, then reads it back.
    .OUTPUTS
        'Granted', 'Present' or throws.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][object]$Grant)
    $present = { Test-PcmSharedGrant -Mailbox $Mailbox -Grant $Grant }
    if (& $present) { return 'Present' }
    if (-not $PSCmdlet.ShouldProcess($Mailbox, "$($Grant.Right) to $($Grant.Trustee)")) { return 'Planned' }
    Confirm-PcmWrite -Target $Mailbox -Operation ('Exchange Online: grant {0} to {1}' -f $Grant.Right, $Grant.Trustee)
    switch ($Grant.Right) {
        'FullAccess' { $null = Add-MailboxPermission -Identity $Mailbox -User $Grant.Trustee -AccessRights FullAccess -InheritanceType All -AutoMapping ([bool]$Grant.AutoMapping) -ErrorAction Stop }
        'SendAs' { $null = Add-RecipientPermission -Identity $Mailbox -Trustee $Grant.Trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop }
        'SendOnBehalf' { Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{ Add = $Grant.Trustee } -ErrorAction Stop }
    }
    $check = Wait-PcmCondition -Test $present -TimeoutSeconds 120 -IntervalSeconds 10
    if (-not $check.Ok) { throw "$($Grant.Right) to $($Grant.Trustee) not visible after 2 min." }
    return 'Granted'
}

# A Delta cycle that Entra Connect refuses because the connector is busy is started again this many times, this many seconds apart (after the wait for the running cycle).
$script:EntraBusyAttempts = 3
$script:EntraBusySpacingSeconds = 60

function Invoke-PcmEntraConnect {
    <#
    .SYNOPSIS
        Entra Connect operation on the rebuilt server: Pause or Resume the scheduler, or run a Delta cycle and wait.
    .DESCRIPTION
        EntraConnect.Mode:
          Remoting  Invoke-Command to EntraConnect.Server (ADSync cmdlets), with the current account.
          Script    EntraConnect.ScriptPath -Operation <op>: any way to reach the server (exit 0 = done). EntraConnect.Server and EntraConnect.CredentialPath are given
                    as -Server and -CredentialPath to a script that declares those parameters (the sample does), so that a signed script needs no edit.
          Manual    The operator does it; the tool asks for a confirmation (in the console, or in the window when the
                    run belongs to it; impossible with -Force in the console).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][ValidateSet('Pause','Resume','Delta')][string]$Operation, [object]$Caller)
    $entra = $Context.Config.EntraConnect
    if (-not $PSCmdlet.ShouldProcess("Entra Connect ($($entra.Mode))", $Operation)) { return 'Planned' }
    if (-not (Test-PcmSameText $entra.Mode 'Manual')) { Confirm-PcmWrite -Target ("Entra Connect ({0})" -f $entra.Mode) -Operation $Operation }
    switch ($entra.Mode) {
        'Remoting' {
            $busyAttempts = $script:EntraBusyAttempts; $busySpacing = $script:EntraBusySpacingSeconds
            try {
                $output = Invoke-Command -ComputerName $entra.Server -ErrorAction Stop -ScriptBlock {
                    Import-Module ADSync -ErrorAction Stop
                    switch ($using:Operation) {
                        'Pause' { Set-ADSyncScheduler -SyncCycleEnabled $false }
                        'Resume' { Set-ADSyncScheduler -SyncCycleEnabled $true }
                        'Delta' {
                            # A cycle that is still running (the connector is busy) is waited for; a start that is refused for that reason is tried again a few times.
                            for ($attempt = 1; $attempt -le $using:busyAttempts; $attempt++) {
                                $t0 = Get-Date; while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 15) { Start-Sleep -Seconds 10 }
                                try { Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop | Out-Null; break }
                                catch {
                                    if ($_.Exception.Message -notmatch 'busy|in\s?progress|already (running|started)|retry operation') { throw }
                                    if ($attempt -ge $using:busyAttempts) { throw 'PCM-ENTRA-BUSY' }
                                    Start-Sleep -Seconds $using:busySpacing
                                }
                            }
                            Start-Sleep -Seconds 10
                            $t0 = Get-Date; while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 30) { Start-Sleep -Seconds 10 }
                        }
                    }
                    $s = Get-ADSyncScheduler
                    'SyncCycleEnabled={0} InProgress={1}' -f $s.SyncCycleEnabled, $s.SyncCycleInProgress
                }
            } catch {
                if ($_.Exception.Message -match 'PCM-ENTRA-BUSY') {
                    $batch = [string](Get-PcmValue $Context 'ConvertBatchId' '')
                    throw ('Entra Connect is busy with another synchronization cycle: wait for it to finish, then run Recover -Batch {0} again.' -f $(if ($batch) { $batch } else { '<id>' }))
                }
                throw
            }
            return (@($output) -join ' ')
        }
        'Script' {
            # A script that ends without 'exit' leaves $LASTEXITCODE unset (StrictMode) or at the value of an earlier call.
            $global:LASTEXITCODE = 0
            # The server and the credential file of the configuration go to a script that declares -Server and -CredentialPath (the signed sample does), so that it runs as it is signed;
            # a script of the administrator that does not declare them gets -Operation only.
            $scriptArguments = @{ Operation = $Operation }
            $declared = (Get-Command -Name $entra.ScriptPath -CommandType ExternalScript -ErrorAction Stop).Parameters
            if ($entra.Server -and $declared.ContainsKey('Server')) { $scriptArguments.Server = [string]$entra.Server }
            if ((Get-PcmValue $entra 'CredentialPath' '') -and $declared.ContainsKey('CredentialPath')) { $scriptArguments.CredentialPath = [string]$entra.CredentialPath }
            $output = & $entra.ScriptPath @scriptArguments
            $code = Get-Variable -Name LASTEXITCODE -Scope Global -ValueOnly -ErrorAction SilentlyContinue
            $lines = (@($output | ForEach-Object { [string]$_ -split '\r?\n' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' | ')
            if ($code) { throw "EntraConnect.ScriptPath -Operation $Operation ended with exit code $code$(if ($lines) { ': ' + $lines } else { '.' })" }
            return $lines
        }
        'Manual' {
            $text = @{ Pause = 'Pause the Entra Connect scheduler (Set-ADSyncScheduler -SyncCycleEnabled $false)'; Resume = 'Resume the Entra Connect scheduler (Set-ADSyncScheduler -SyncCycleEnabled $true)'; Delta = 'Run a delta synchronisation (Start-ADSyncSyncCycle -PolicyType Delta) and wait for its end' }[$Operation]
            # The window answers through its event file (PRA_EVENT_FILE); the console through the caller.
            if (-not $Caller -and -not $env:PRA_EVENT_FILE) { throw "EntraConnect.Mode = Manual needs the operator: $text." }
            $where = if ($entra.Server) { $entra.Server } else { 'the Entra Connect server' }
            if (-not (Request-PcmOperator -Question "$text on $where. Done?" -Title 'PRA Cloud Mailbox - Entra Connect' -Caller $Caller)) { throw "Entra Connect operation $Operation not confirmed by the operator." }
            return 'confirmed by the operator'
        }
    }
}

function Invoke-PcmEntraConnectObject {
    <#
    .SYNOPSIS
        Targeted synchronisation of single objects on the Entra Connect server, by their on-premises distinguished name.
    .DESCRIPTION
        Recover uses it for a deleted shared identity that a delta cycle did not recreate. It returns one answer per object (Dn, Ok, Detail) and does not throw for an object that failed.
        EntraConnect.Mode:
          Remoting  Invoke-Command to EntraConnect.Server. Invoke-ADSyncSingleObjectSync (Entra Connect 2.x) reads the object from Active Directory, synchronises it and exports it, one
                    object after the other. That cmdlet needs a session that can reach the management service of Entra Connect, which a remote session often cannot: the objects it does
                    not synchronise (and every object, once two were refused) are synchronised together by one cycle in which the Active Directory connector runs a full
                    synchronisation (Set-ADSyncSchedulerConnectorOverride -FullSyncRequired, then Start-ADSyncSyncCycle -PolicyType Delta, waited for).
          Script    EntraConnect.ScriptPath -Operation SyncObject -ObjectDn <dn>, one object after the other (the signed sample does it). An answer that says the full synchronisation of
                    the connector ran covers the objects that are left. A script that does not take -ObjectDn, or that refuses the operation, gives a failed answer.
          Manual    The operator is given the exact command for each object and confirms.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string[]]$ObjectDn, [object]$Caller)
    $entra = $Context.Config.EntraConnect
    $failed = { param($Detail) @($ObjectDn | ForEach-Object { [pscustomobject]@{ Dn = [string]$_; Ok = $false; Detail = [string]$Detail } }) }
    if (-not $PSCmdlet.ShouldProcess("Entra Connect ($($entra.Mode))", 'SyncObject')) { return @(& $failed 'planned') }
    if (-not (Test-PcmSameText $entra.Mode 'Manual')) { Confirm-PcmWrite -Target ("Entra Connect ({0})" -f $entra.Mode) -Operation 'SyncObject' }
    switch ($entra.Mode) {
        'Remoting' {
            $rows = @(Invoke-Command -ComputerName $entra.Server -ErrorAction Stop -ScriptBlock {
                    Import-Module ADSync -ErrorAction Stop
                    $single = Get-Command -Name Invoke-ADSyncSingleObjectSync -ErrorAction SilentlyContinue
                    # A cycle of the scheduler that is running is waited for; a call that Entra Connect refuses with an error (not with an answer) is tried again, as its services may be busy with that cycle.
                    $retry = {
                        param([scriptblock]$Call)
                        for ($attempt = 1; $attempt -le 3; $attempt++) {
                            $t0 = Get-Date; while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 15) { Start-Sleep -Seconds 10 }
                            try { return (& $Call) } catch { if ($attempt -ge 3) { throw }; Start-Sleep -Seconds 30 }
                        }
                    }
                    $answers = @{}
                    $left = [System.Collections.Generic.List[string]]::new()
                    $refused = 0
                    foreach ($dn in @($using:ObjectDn)) {
                        $detail = 'Invoke-ADSyncSingleObjectSync is not available on this server'
                        if ($single -and $refused -lt 2) {
                            try {
                                $report = (& $retry { Invoke-ADSyncSingleObjectSync -DistinguishedName $dn -NoHtmlReport -ErrorAction Stop | Out-String }) | ConvertFrom-Json
                                if ([string]$report.StatusInfo.Status -eq 'Success') { $answers[$dn] = @{ Ok = $true; Detail = ('{0} by Invoke-ADSyncSingleObjectSync' -f $report.Action) }; continue }
                                $detail = (@($report.StatusInfo.ErrorCode, $report.StatusInfo.Reason) | Where-Object { $_ }) -join ' '
                            } catch { $detail = $_.Exception.Message }
                            $refused++
                        } elseif ($single) { $detail = 'not tried: the objects before it were refused by Invoke-ADSyncSingleObjectSync' }
                        $left.Add($dn); $answers[$dn] = @{ Ok = $false; Detail = $detail }
                    }
                    if ($left.Count) {
                        try {
                            $connectors = @(Get-ADSyncConnector)
                            $names = @($left | ForEach-Object {
                                    $dn = $_
                                    ($connectors | Where-Object { $_.Type -eq 'AD' -and @($_.Partitions | Where-Object { $_.DN -and $dn.EndsWith($_.DN, [StringComparison]::OrdinalIgnoreCase) }).Count } | Select-Object -First 1).Name
                                } | Where-Object { $_ } | Select-Object -Unique)
                            if (-not $names.Count) { throw 'the Active Directory connector of the objects was not found' }
                            $null = & $retry { foreach ($name in $names) { Set-ADSyncSchedulerConnectorOverride -ConnectorName $name -FullSyncRequired $true -ErrorAction Stop } }
                            $null = & $retry { Start-ADSyncSyncCycle -PolicyType Delta -ErrorAction Stop }
                            Start-Sleep -Seconds 10
                            $t0 = Get-Date; while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 30) { Start-Sleep -Seconds 10 }
                            foreach ($dn in $left) { $answers[$dn] = @{ Ok = $true; Detail = ('full synchronisation of the Active Directory connector, one cycle for {0} object(s) ({1})' -f $left.Count, $answers[$dn].Detail) } }
                        } catch { foreach ($dn in $left) { $answers[$dn] = @{ Ok = $false; Detail = ('{0}; full synchronisation of the connector: {1}' -f $answers[$dn].Detail, $_.Exception.Message) } } }
                    }
                    foreach ($dn in @($using:ObjectDn)) { [pscustomobject]@{ Dn = $dn; Ok = [bool]$answers[$dn].Ok; Detail = [string]$answers[$dn].Detail } }
                })
            return @($rows | ForEach-Object { [pscustomobject]@{ Dn = [string]$_.Dn; Ok = [bool]$_.Ok; Detail = [string]$_.Detail } })
        }
        'Script' {
            $declared = (Get-Command -Name $entra.ScriptPath -CommandType ExternalScript -ErrorAction Stop).Parameters
            if (-not $declared.ContainsKey('ObjectDn')) { return @(& $failed 'EntraConnect.ScriptPath does not take -ObjectDn: it cannot synchronise one object') }
            $rows = [System.Collections.Generic.List[object]]::new()
            $covered = ''
            foreach ($dn in $ObjectDn) {
                if ($covered) { $rows.Add([pscustomobject]@{ Dn = [string]$dn; Ok = $true; Detail = $covered }); continue }
                $global:LASTEXITCODE = 0
                $scriptArguments = @{ Operation = 'SyncObject'; ObjectDn = [string]$dn }
                if ($entra.Server -and $declared.ContainsKey('Server')) { $scriptArguments.Server = [string]$entra.Server }
                if ((Get-PcmValue $entra 'CredentialPath' '') -and $declared.ContainsKey('CredentialPath')) { $scriptArguments.CredentialPath = [string]$entra.CredentialPath }
                try {
                    $output = & $entra.ScriptPath @scriptArguments
                    $code = Get-Variable -Name LASTEXITCODE -Scope Global -ValueOnly -ErrorAction SilentlyContinue
                    $lines = (@($output | ForEach-Object { [string]$_ -split '\r?\n' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' | ')
                    $rows.Add([pscustomobject]@{ Dn = [string]$dn; Ok = (-not $code); Detail = $(if ($code) { "exit code $code$(if ($lines) { ': ' + $lines })" } else { $lines }) })
                    if (-not $code -and $lines -match 'full synchronisation of the Active Directory connector') { $covered = $lines }
                } catch { $rows.Add([pscustomobject]@{ Dn = [string]$dn; Ok = $false; Detail = $_.Exception.Message }) }
            }
            return @($rows)
        }
        'Manual' {
            if (-not $Caller -and -not $env:PRA_EVENT_FILE) { return @(& $failed 'EntraConnect.Mode = Manual needs the operator') }
            $where = if ($entra.Server) { $entra.Server } else { 'the Entra Connect server' }
            $commands = @($ObjectDn | ForEach-Object { "Invoke-ADSyncSingleObjectSync -DistinguishedName '{0}' -NoHtmlReport" -f ([string]$_).Replace("'", "''") })
            $question = "On $where, run these commands one after the other (each answer must show Status Success):`n" + ($commands -join "`n") + "`nIf one of them fails, run Start-ADSyncSyncCycle -PolicyType Delta after Set-ADSyncSchedulerConnectorOverride -ConnectorName <the Active Directory connector> -FullSyncRequired `$true.`nDone?"
            if (-not (Request-PcmOperator -Question $question -Title 'PRA Cloud Mailbox - Entra Connect' -Caller $Caller)) { return @(& $failed 'not confirmed by the operator') }
            return @($ObjectDn | ForEach-Object { [pscustomobject]@{ Dn = [string]$_; Ok = $true; Detail = 'confirmed by the operator' } })
        }
    }
}
#endregion

Export-ModuleMember -Function Get-PcmRunningPipeline, Get-PcmCallTimeout, Get-PcmCallTimeoutSecond, Get-PcmCallTimeoutMessage, Start-PcmCallGuard, Stop-PcmCallGuard, Install-PcmCallGuard, Test-PcmClientTimeout,
    Connect-PcmCloud, Disconnect-PcmCloud, Get-PcmModuleQuietDefault, Read-PcmModuleWarning, Add-PcmModuleWarning, Set-PcmWriteApproval, Confirm-PcmWrite, Invoke-PcmGraph, Get-PcmGraphRole, Get-PcmTenantFact, Get-PcmSkuExchangePlan,
    Test-PcmSharedGrant, Test-PcmBatchExpectation, Get-PcmDelegatedScope, Test-PcmNonInteractiveProcess, Test-PcmGroupReadRole, Get-PcmAdminRole, Get-PcmRoleNeed, Test-PcmAdminRole, Find-PcmEntraGroup, Get-PcmEntraGroupMember, Get-PcmExoCommandName,
    Test-PcmNativeResource, Get-PcmResourceIntent, Get-PcmResourceProtectionMode, Assert-PcmUnprotectedResource, Assert-PcmResourceOrganizationUnheld, Test-PcmUnprotectedResourceDeleted, Get-PcmUnprotectedRecoverCollisions, Get-PcmUnprotectedDeletionProgress, Assert-PcmUnprotectedDeletionRepeatable, Get-PcmResourceAddresses, Get-PcmResourceInventory, Get-PcmResourceCollisions,
    Get-PcmOwnedResource, Read-PcmResourceSettings, Test-PcmResourceConfigurationReadiness, Get-PcmResourceGroup, Test-PcmResourceLicence, Set-PcmResourceProvisioningDisabled, Set-PcmResourceConfiguration, Enable-PcmResourceBooking, Set-PcmResourceRoomListMembership, Get-PcmResourceRoute,
    Get-PcmResourceKeys, ConvertTo-PcmRoomListRecord, Get-PcmResourceCloudRoomList, Test-PcmRoomListCreationReadiness, Set-PcmRoomListConfiguration,
    Get-PcmMailboxPlanId, Get-PcmFreeUnit, Get-PcmSkuSpelling, Get-PcmSharedUnitCount, Test-PcmTenantReadiness, Get-PcmCloudState, Get-PcmReadiness, Get-PcmTrusteeFinding,
    Get-PcmCloudStateSet, Get-PcmExoStateSet, Get-PcmSyncBehaviorSet, Invoke-PcmGraphBatch, Get-PcmGraphUserSet, ConvertFrom-PcmMailboxLocation, Test-PcmNetworkFailure,
    Wait-PcmCondition, Get-PcmExoState, Set-PcmSourceOfAuthority, Get-PcmSourceOfAuthority, Set-PcmUsageLocation, Set-PcmGroupMember,
    Test-PcmGroupMember, Get-PcmDisabledPlanForExchangeOnly, Set-PcmLicence, Get-PcmUserLicence, Get-PcmDirectLicence, Enable-PcmLicencePlan,
    Restore-PcmLicence, Remove-PcmIdentity, Test-PcmDeletedIdentity,
    Connect-PcmCompliance, Get-PcmHoldPolicySet, Test-PcmToolHold, Add-PcmHoldLocationSet, Remove-PcmHoldLocationSet, Get-PcmKioskAssignment, Get-PcmSharedGrant, Grant-PcmSharedPermission, Invoke-PcmEntraConnect, Invoke-PcmEntraConnectObject

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAAfFVmaISZ1iOT
# hE4LEVQtKAG/o3tZWGcZgOe03b0KUKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDah8Fhb89F
# j/dAWrRw+8ir8O37WcpAr+vtIrMSYlZtazANBgkqhkiG9w0BAQEFAASCAgA+6UHd
# Uvr0DsK6pBSBLjqAl2SKEcVBwDlk7msHIxeaO5gpd0V6fbFtFAeRLLI4QLV0hnji
# QLWZojwjwMLAYYx2Qsjq3U0MlVe+S1aZxmZ5J3sr89BIyxp3tSchmGEuge5w/i6r
# i6BaB+cJz489yd+stjov+AzcQr9YoJuV3YLia8rRsw4dgg8AgC2jbsrVrRPBkcyc
# kFbM6FXA4OHaShd3zhi7v95FTD/kEzO5vils3JBuiDZHg+AZ3SxEfNQLC13sKF77
# X44r8oXHcZBYiQsf7+Lf3QQ3wUk8WwIys5Wbg17qWseWKxTfdbq4OVJdzDBqnxWZ
# KTwlBzGLY7utAoTn+2b8yqSC6hNeAUiLoDqMzhvl9sTXkPVRm8gkNhAPgS6FgYJ0
# EgdCTAdinSfvPl07/r72gYMydLbTkOxgd9xAcZxyzAv/oxEw0vy7lUQPezeuVSnC
# NJ+uNFwbe/KsuW+rOzPHnlxd+viJcZ+MW+xhnGtZ65TPMiBWvlCqIerxwaPhGzJq
# mLHj6P3AayMdWwSXcR3OWRguG66ulIZvZi8gj8chqS4tGC12B3WxmtljzYwZJAGh
# 7XrGKMtvqS/xfIso2Tn79AcBT+NkgHeIak4eHr0izlmYQmz0HHmYDLZVxsdPpngN
# 0EoTIqyERKPjSk3U5FLK1AMt9MHPr+cetFCVhqGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIAaQOmNqe+/J+6nc8DWNFORwNfk9LhaNHsIBSRnVl5tpAggipVm4Curx
# GBgPMjAyNjExMTUyMDA2MjBaMAMCAQECEE3eph2uAR6JllzBtqHTZ6agghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTIwMDYyMFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIHxj
# rSBC0BKADGDDKMihyqzQPiAy7FRTIR1+3LXcAU+tMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAGEfWMwmZuclUMQbIvbdNJhc68Pynpdqw+XDOWX5IcgwIhAI9O
# 6i1frtMzL5VAQoeMh+hu2cMLQUzpUf/EYSIlDhkz
# SIG # End signature block
