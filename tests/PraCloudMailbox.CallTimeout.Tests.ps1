<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the time limit of the calls to Microsoft 365 (Pester 5+, offline).
.DESCRIPTION
    Polling.CallTimeoutMinutes (default, range, case), the sentence of a stop, the time limit of the Microsoft Graph transport, the guard that puts the limit in front of
    the Exchange Online commands (same parameters, same pipeline, same errors, armed during the call only), and the real entry script in its own pwsh process against the
    in-memory tenant of FakeTenant.ps1 (tests\CallTimeout.Child.ps1): a call that never answers ends the run with the plain message, the next steps, exit code 1 and
    a journal that the next run continues; a call that cannot be interrupted ends the process with the same message. PowerShell 7 only.
.NOTES
    Author : Nicolas Fabert
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    foreach ($name in 'PRA.Common', 'PRA.Store', 'PRA.Cloud') { Import-Module (Join-Path $script:Root "module\$name.psm1") -Force -Global }
    . (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'

    # The entry script as a command that runs in its own process: no bootstrap, and the child side (tests\CallTimeout.Child.ps1) in place of the module imports.
    $text = [IO.File]::ReadAllText((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'))
    $start = $text.IndexOf('# >>> PRA package integrity bootstrap'); $end = $text.IndexOf('# <<< PRA package integrity bootstrap')
    $text = $text.Remove($start, $end - $start)
    $text = [regex]::Replace($text, '(?m)^Import-Module \(Join-Path \$PSScriptRoot ''PraCloudMailbox\.psd1''\).*$', '')
    $text = [regex]::Replace($text, '(?m)^if \(\$Action -in @\(''Collect'', ''Schedule''\)\) \{ Import-Module.*\r?\n.*Import-Module.*$', '. (Join-Path $env:PCM_TEST_DIR ''CallTimeout.Child.ps1'')')
    # The folder of the package holds nothing but the signed files: the copy sits beside it and works on the package through PCM_TEST_ROOT.
    $text = $text.Replace('$PSScriptRoot', '$env:PCM_TEST_ROOT')
    $script:ChildEntry = Join-Path (Split-Path $script:Root -Parent) 'Invoke-Child.ps1'
    [IO.File]::WriteAllText($script:ChildEntry, $text, [Text.UTF8Encoding]::new($true))

    function script:Invoke-Child {
        <# One run of the entry script in its own process. Work: the folder of the run (snapshot, journal, logs); a run that continues a batch gives the same one. #>
        param([Parameter(Mandatory)][string]$Scenario, [Parameter(Mandatory)][string]$Work, [string[]]$Arguments = @(), [int]$TimeoutSeconds = 180)
        $info = [Diagnostics.ProcessStartInfo]::new($script:Pwsh)
        foreach ($argument in (@('-NoProfile', '-NonInteractive', '-File', $script:ChildEntry) + $Arguments + @('-ConfigPath', (Join-Path $Work 'child.config.psd1')))) { $info.ArgumentList.Add($argument) }
        $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true; $info.UseShellExecute = $false
        $info.Environment['PCM_TEST_ROOT'] = $script:Root; $info.Environment['PCM_TEST_WORK'] = $Work; $info.Environment['PCM_TEST_SCENARIO'] = $Scenario
        $info.Environment['PCM_TEST_DIR'] = $PSScriptRoot; $info.Environment['PRA_ICONS'] = 'Ascii'; $null = $info.Environment.Remove('PRA_ANSI')
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $process = [Diagnostics.Process]::Start($info)
        $out = $process.StandardOutput.ReadToEndAsync(); $err = $process.StandardError.ReadToEndAsync()
        $finished = $process.WaitForExit($TimeoutSeconds * 1000)
        if (-not $finished) { $process.Kill($true); $process.WaitForExit() }
        [pscustomobject]@{ Exit = $process.ExitCode; Out = $out.Result; Err = $err.Result; Seconds = $clock.Elapsed.TotalSeconds; Finished = $finished }
    }

    function script:New-ChildWork { $folder = Join-Path $TestDrive ('child-' + [guid]::NewGuid().ToString('N').Substring(0, 8)); $null = New-Item -ItemType Directory -Path $folder; $folder }

    function script:Get-ChildJournal {
        param([string]$Work, [string]$BatchId)
        $journal = Open-PcmStore -Path (Join-Path $Work 'data\PraCloudMailbox-journal.db') -Root $script:Root -Kind Journal -ReadOnly
        try {
            [pscustomobject]@{ Batch = (Get-PcmBatch $journal -Id $BatchId); Items = @(Get-PcmBatchItem $journal $BatchId)
                Events = @(Invoke-PcmSql $journal 'SELECT step, outcome, detail FROM batch_event WHERE batch_id = @b ORDER BY time_utc, rowid' @{ b = $BatchId } -As Rows) }
        } finally { Close-PcmStore $journal }
    }

    function script:Write-TestConfig {
        param([string]$Text)
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $folder
        $path = Join-Path $folder 'c.psd1'
        Set-Content -LiteralPath $path -Value ("@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'x' }; $Text }") -Encoding UTF8
        [pscustomobject]@{ Path = $path; Folder = $folder }
    }
}

AfterAll {
    foreach ($name in 'Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext', 'Reset-PraFakeGraphTransport') { Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Call time limit: the setting' {
    It 'is 5 minutes when the setting is not there' {
        $file = Write-TestConfig ''
        (Import-PcmConfiguration -Path $file.Path -Root $file.Folder).Polling.CallTimeoutMinutes | Should -Be 5
    }
    It 'accepts <Minutes> minute(s), whatever the letter case of the name' -TestCases @(@{ Minutes = 1 }, @{ Minutes = 10 }, @{ Minutes = 60 }) {
        param($Minutes)
        $file = Write-TestConfig "Polling = @{ calltimeoutMINUTES = $Minutes }"
        (Import-PcmConfiguration -Path $file.Path -Root $file.Folder).Polling.CallTimeoutMinutes | Should -Be $Minutes
    }
    It 'refuses <Value>, in plain words' -TestCases @(@{ Value = '0' }, @{ Value = '61' }, @{ Value = '-3' }, @{ Value = '2.5' }, @{ Value = "'five'" }) {
        param($Value)
        $file = Write-TestConfig "Polling = @{ CallTimeoutMinutes = $Value }"
        { Import-PcmConfiguration -Path $file.Path -Root $file.Folder } | Should -Throw '*Polling.CallTimeoutMinutes must be a whole number*'
    }
    It 'is documented in the configuration sample and in the administrator guide' {
        (Get-Content -LiteralPath (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1') -Raw) | Should -Match 'CallTimeoutMinutes\s*=\s*5'
        (Get-Content -LiteralPath (Join-Path $script:Root 'docs\PraCloudMailbox-Guide.md') -Raw) | Should -Match 'CallTimeoutMinutes'
    }
    It 'turns the minutes of a configuration into seconds, 5 minutes without one' {
        Get-PcmCallTimeoutSecond -Config @{ Polling = @{ CallTimeoutMinutes = 2 } } | Should -Be 120
        Get-PcmCallTimeoutSecond -Config @{} | Should -Be 300
        Get-PcmCallTimeoutSecond -Config $null | Should -Be 300
    }
}

Describe 'Call time limit: the sentence of a stop' {
    It 'says what did not answer, where, that nothing more was changed, where to look and what to do (a read)' {
        Get-PcmCallTimeoutMessage -Seconds 300 -Step 'Tenant prerequisites' -Call 'Get-Recipient' |
            Should -BeExactly 'Microsoft 365 did not answer within 5 minutes during Tenant prerequisites (Get-Recipient); nothing more was changed. Check https://status.cloud.microsoft and the service health of your tenant, then run the same command again.'
    }
    It 'says that a change that did not answer may or may not have been applied, and that the next run reads it back' {
        $message = Get-PcmCallTimeoutMessage -Seconds 600 -Step 'Users: source of authority and Exchange plan' -Call 'Set-MailUser' -Write
        $message | Should -BeLike 'Microsoft 365 did not answer within 10 minutes during Users: source of authority and Exchange plan (Set-MailUser). That last change may or may not have been applied*'
        $message | Should -BeLike '*the next run reads it back before going on. Check https://status.cloud.microsoft*then run the command given in the next steps.'
    }
    It 'writes one minute, seconds, and no step' {
        Get-PcmCallTimeoutMessage -Seconds 60 -Step 'S' | Should -BeLike 'Microsoft 365 did not answer within 1 minute during S;*'
        Get-PcmCallTimeoutMessage -Seconds 3 -Step 'S' | Should -BeLike 'Microsoft 365 did not answer within 3 seconds during S;*'
        Get-PcmCallTimeoutMessage -Seconds 300 | Should -BeLike 'Microsoft 365 did not answer within 5 minutes during this run;*'
    }
}

Describe 'Call time limit: Microsoft Graph' {
    BeforeAll {
        function global:Connect-MgGraph { [CmdletBinding()] param($ClientId, $TenantId, $CertificateThumbprint, $ContextScope, [switch]$NoWelcome, $Scopes) }
        function global:Disconnect-MgGraph { [CmdletBinding()] param() }
        function global:Connect-ExchangeOnline { [CmdletBinding()] param($AppId, $CertificateThumbprint, $Organization, $ShowBanner, $CommandName, $UserPrincipalName, [switch]$DisableWAM) }
        function global:Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param() }
        function global:Invoke-MgGraphRequest { [CmdletBinding()] param([string]$Method = 'GET', [string]$Uri, [object]$Body, [string]$ContentType, [string]$OutputType, [System.Collections.IDictionary]$Headers) }
    }
    AfterAll {
        foreach ($name in 'Connect-MgGraph', 'Disconnect-MgGraph', 'Invoke-MgGraphRequest', 'Connect-ExchangeOnline', 'Disconnect-ExchangeOnline') { Remove-Item -Path "Function:\$name" -ErrorAction SilentlyContinue }
    }
    BeforeEach {
        Reset-PraFakeGraphTransport
        & (Get-Module PRA.Cloud) { $script:GraphRequestOwner['CurrentPhase'] = 'Tenant prerequisites' }
        $script:graphCalls = 0
        Mock Start-Sleep -ModuleName PRA.Cloud { }
    }
    AfterEach { Stop-PcmCallGuard }
    It 'gives the HTTP client of the SDK a little more than the limit, and puts the previous value back after the disconnect' {
        $context = @{ Config = @{ Cloud = @{ TenantId = 'synthetic'; Organization = 'synthetic'; AppId = 'synthetic'; CertificateThumbprint = '12345678' }; Scope = @{}; Polling = @{ CallTimeoutMinutes = 2 } }
            Approval = { param($Target, $Operation) $true } }
        Mock Import-Module -ModuleName PRA.Cloud { }
        Mock Get-ChildItem -ModuleName PRA.Cloud { [pscustomobject]@{ HasPrivateKey = $true; NotAfter = (Get-Date).AddYears(1) } }
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { }
        Mock Connect-MgGraph -ModuleName PRA.Cloud { $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = $ClientId } }
        Mock Disconnect-MgGraph -ModuleName PRA.Cloud { $global:PraFakeGraphAuth = $null }
        & (Get-Module PRA.Cloud) { $script:GraphRequestOwner = $null; $script:GraphTransportReady = $false }
        $global:PraFakeGraphRequest.ClientTimeout = [timespan]::FromMinutes(5)
        $null = Connect-PcmCloud -Context $context
        $global:PraFakeGraphRequest.ClientTimeout | Should -Be ([timespan]::FromSeconds(150))
        Disconnect-PcmCloud -Context $context
        $global:PraFakeGraphRequest.ClientTimeout | Should -Be ([timespan]::FromMinutes(5))
        Get-PcmCallTimeout | Should -BeNullOrEmpty
    }
    It 'does not send a read again when the client gave up: the run stops with the plain message' {
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud {
            $script:graphCalls++
            throw [Threading.Tasks.TaskCanceledException]::new('The request was canceled due to the configured HttpClient.Timeout of 330 seconds elapsing.', [TimeoutException]::new('The operation was canceled.'))
        }
        $failure = try { Invoke-PcmGraph -Uri 'v1.0/subscribedSkus' } catch { $_ }
        $failure.Exception | Should -BeOfType ([TimeoutException])
        $failure.Exception.Message | Should -BeExactly 'Microsoft 365 did not answer within 5 minutes during Tenant prerequisites (Microsoft Graph GET subscribedSkus); nothing more was changed. Check https://status.cloud.microsoft and the service health of your tenant, then run the same command again.'
        $script:graphCalls | Should -Be 1
        Should -Invoke Start-Sleep -ModuleName PRA.Cloud -Times 0 -Exactly
    }
    It 'says that a change may or may not have been applied, also for a batch that holds one' -TestCases @(
        @{ Case = 'a PATCH'; Method = 'PATCH'; Uri = 'v1.0/users/u1'; Body = @{ usageLocation = 'FR' } }
        @{ Case = 'a batch with a DELETE'; Method = 'POST'; Uri = 'v1.0/$batch'; Body = @{ requests = @(@{ id = '0'; method = 'GET'; url = '/users/a' }, @{ id = '1'; method = 'DELETE'; url = '/users/b' }) } }) {
        param($Case, $Method, $Uri, $Body)
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { throw [Threading.Tasks.TaskCanceledException]::new('canceled', [TimeoutException]::new('timeout')) }
        $failure = try { Invoke-PcmGraph -Method $Method -Uri $Uri -Body $Body -Approved } catch { $_ }
        $failure.Exception.Message | Should -BeLike '*That last change may or may not have been applied*'
    }
    It 'says nothing more was changed for a batch that only reads' {
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { throw [Threading.Tasks.TaskCanceledException]::new('canceled', [TimeoutException]::new('timeout')) }
        $failure = try { Invoke-PcmGraph -Method POST -Uri 'v1.0/$batch' -Body @{ requests = @(@{ id = '0'; method = 'GET'; url = '/users/a' }) } -Approved } catch { $_ }
        $failure.Exception.Message | Should -BeLike '*; nothing more was changed.*'
    }
    It 'keeps retrying a refused or a broken connection as before: only the client time limit is final' {
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { $script:graphCalls++; throw [Exception]::new('down', [IO.IOException]::new('reset')) }
        { Invoke-PcmGraph -Uri 'v1.0/users/u1' } | Should -Throw '*down*'
        $script:graphCalls | Should -Be 5
    }
    It 'refuses every new call after the first one that did not answer, at once, and still allows the calls that end a run' {
        $context = @{ Config = @{ Polling = @{ CallTimeoutMinutes = 5 } }; CurrentPhase = 'Tenant prerequisites'; Action = 'Check' }
        Start-PcmCallGuard -Context $context
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { $script:graphCalls++; throw [Threading.Tasks.TaskCanceledException]::new('canceled', [TimeoutException]::new('timeout')) }
        $null = try { Invoke-PcmGraph -Uri 'v1.0/users/u1' } catch { $_ }
        $script:graphCalls | Should -Be 1
        Get-PcmCallTimeout | Should -BeLike 'Microsoft 365 did not answer within 5 minutes during Tenant prerequisites*'
        $second = try { Invoke-PcmGraph -Uri 'v1.0/users/u2' } catch { $_ }
        $second.Exception.Message | Should -BeLike 'Microsoft 365 did not answer within 5 minutes*'
        $script:graphCalls | Should -Be 1
        & (Get-Module PRA.Cloud) { Enter-PcmCall -Call 'Disconnect-MgGraph' -Seconds 30 -Cleanup; Exit-PcmCall; $true } | Should -BeTrue
        Stop-PcmCallGuard
        Get-PcmCallTimeout | Should -BeLike 'Microsoft 365 did not answer within 5 minutes*'
    }
}

Describe 'Call time limit: the guard of the Exchange Online commands' {
    BeforeAll {
        function global:Get-Mailbox {
            [CmdletBinding()] param([Parameter(Position = 0, ValueFromPipeline)][string]$Identity, [int]$ResultSize = 1)
            process {
                $global:PraGuardSeen = & (Get-Module PRA.Cloud) { if ($script:CallGuard) { [pscustomobject]@{ Armed = $script:CallGuard.Armed; Depth = $script:CallGuard.Depth; Message = $script:CallGuard.Message } } }
                if ($Identity -eq 'boom') { throw [System.InvalidOperationException]::new('boom text') }
                if ($Identity -eq 'warn') { Write-Warning 'careful'; return 'warned' }
                "mailbox:$Identity`:$ResultSize"
            }
        }
        function global:Set-Mailbox {
            [CmdletBinding()] param([string]$Identity)
            $global:PraGuardSeen = & (Get-Module PRA.Cloud) { if ($script:CallGuard) { [pscustomobject]@{ Armed = $script:CallGuard.Armed; Depth = $script:CallGuard.Depth; Message = $script:CallGuard.Message } } }
            "set:$Identity"
        }
    }
    AfterAll { foreach ($name in 'Get-Mailbox', 'Set-Mailbox') { Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue } }
    BeforeEach {
        $script:Context = @{ Config = @{ Polling = @{ CallTimeoutMinutes = 5 } }; CurrentPhase = 'Objects'; Action = 'Check' }
        Start-PcmCallGuard -Context $script:Context
        Install-PcmCallGuard -Name 'Get-Mailbox', 'Set-Mailbox', 'Get-DoesNotExist'
    }
    AfterEach { Stop-PcmCallGuard }

    It 'puts the guard in front of the commands that exist, and only of those' {
        (Get-Command Get-Mailbox).CommandType | Should -Be 'Alias'
        (Get-Command Set-Mailbox).CommandType | Should -Be 'Alias'
        Get-Command Get-DoesNotExist -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }
    It 'keeps the parameters, the positions, the pipeline and the output of the command' {
        Get-Mailbox 'alice' | Should -Be 'mailbox:alice:1'
        Get-Mailbox -Identity 'bob' -ResultSize 7 | Should -Be 'mailbox:bob:7'
        @('a', 'b') | Get-Mailbox | Should -Be @('mailbox:a:1', 'mailbox:b:1')
        Set-Mailbox -Identity 'carol' | Should -Be 'set:carol'
    }
    It 'keeps the errors of the command: the same exception type, and -ErrorAction decides as before' {
        $failure = try { Get-Mailbox 'boom' -ErrorAction Stop } catch { $_ }
        $failure.Exception | Should -BeOfType ([System.InvalidOperationException])
        $failure.Exception.Message | Should -Be 'boom text'
        { Get-Mailbox -Identity 'x' -NoSuchParameter 1 } | Should -Throw
        Get-Mailbox 'warn' -WarningAction SilentlyContinue | Should -Be 'warned'
    }
    It 'arms the deadline during the call and disarms it after, also when the command fails' {
        $null = Get-Mailbox 'alice'
        $global:PraGuardSeen.Armed | Should -BeTrue
        $global:PraGuardSeen.Depth | Should -Be 1
        & (Get-Module PRA.Cloud) { $script:CallGuard.Armed } | Should -BeFalse
        $null = try { Get-Mailbox 'boom' -ErrorAction Stop } catch { $null = $_ }
        & (Get-Module PRA.Cloud) { $script:CallGuard.Depth } | Should -Be 0
        & (Get-Module PRA.Cloud) { $script:CallGuard.Armed } | Should -BeFalse
    }
    It 'knows a read from a change: the sentence of the stop says so, with the step and the command' {
        $null = Get-Mailbox 'alice'
        $global:PraGuardSeen.Message | Should -BeExactly 'Microsoft 365 did not answer within 5 minutes during Objects (Get-Mailbox); nothing more was changed. Check https://status.cloud.microsoft and the service health of your tenant, then run the same command again.'
        $null = Set-Mailbox -Identity 'alice'
        $global:PraGuardSeen.Message | Should -BeLike 'Microsoft 365 did not answer within 5 minutes during Objects (Set-Mailbox). That last change may or may not have been applied*'
    }
    It 'keeps the deadline of the outer call when a call is made inside another one' {
        & (Get-Module PRA.Cloud) {
            Enter-PcmCall -Call 'outer'
            $first = $script:CallGuard.DeadlineTicks
            Enter-PcmCall -Call 'inner' -Seconds 5
            $same = ($script:CallGuard.DeadlineTicks -eq $first) -and $script:CallGuard.Depth -eq 2
            Exit-PcmCall; Exit-PcmCall
            $same -and -not $script:CallGuard.Armed
        } | Should -BeTrue
    }
    It 'refuses a call after a call did not answer, without sending it, and removes its commands with the guard' {
        & (Get-Module PRA.Cloud) { Set-PcmCallExpired -Message 'Microsoft 365 did not answer within 5 minutes during Objects (Get-Mailbox); nothing more was changed.' }
        $global:PraGuardSeen = $null
        $failure = try { Get-Mailbox 'alice' -ErrorAction Stop } catch { $_ }
        $failure.Exception.Message | Should -BeLike 'Microsoft 365 did not answer within 5 minutes*'
        $global:PraGuardSeen | Should -BeNullOrEmpty
        Stop-PcmCallGuard
        (Get-Command Get-Mailbox).CommandType | Should -Be 'Function'
        Get-Command 'PcmGuard_Get-Mailbox' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        Get-Mailbox 'alice' | Should -Be 'mailbox:alice:1'
    }
    It 'removes the aliases and the proxy commands when the run ends, at its normal end and after an error (a run that is stopped: see the real run below)' {
        $leftover = { @(Get-Command -Name 'PcmGuard_*' -ErrorAction SilentlyContinue).Count + @(Get-Alias -ErrorAction SilentlyContinue | Where-Object { [string]$_.Definition -like 'PcmGuard_*' }).Count }
        & $leftover | Should -BeGreaterThan 0
        Stop-PcmCallGuard
        & $leftover | Should -Be 0
        Start-PcmCallGuard -Context $script:Context; Install-PcmCallGuard -Name 'Get-Mailbox', 'Set-Mailbox'
        $null = try { Get-Mailbox 'boom' -ErrorAction Stop } catch { $_ }
        Disconnect-PcmCloud -Context $script:Context
        & $leftover | Should -Be 0
        (Get-Command Get-Mailbox).CommandType | Should -Be 'Function'
    }
    It 'covers the commands that the tool calls and none of the sign-ins' {
        $names = & (Get-Module PRA.Cloud) { Get-PcmGuardedCommandName }
        $names | Should -Contain 'Get-Recipient'
        $names | Should -Contain 'Set-MailUser'
        $names | Should -Contain 'New-CaseHoldRule'
        $names | Should -Contain 'Get-EXOMailbox'
        @($names | Where-Object { $_ -match '^(Connect|Disconnect|Invoke)-' }).Count | Should -Be 0
    }
}

Describe 'Call time limit: the pipeline that the watchdog stops' {
    AfterEach { Stop-PcmCallGuard }
    It 'is found on this PowerShell' {
        Get-PcmRunningPipeline | Should -Not -BeNullOrEmpty
    }
    It 'is nothing, and no error, when the member that gives it is not there (another version of PowerShell)' {
        { Get-PcmRunningPipeline -MemberName 'NoSuchMemberOfThisVersion' } | Should -Not -Throw
        Get-PcmRunningPipeline -MemberName 'NoSuchMemberOfThisVersion' | Should -BeNullOrEmpty
    }
    It 'falls back cleanly when no pipeline can be reached: the guard starts, says why in the log, and ends the process instead of the run' {
        $lines = [System.Collections.Generic.List[string]]::new()
        Mock Get-PcmRunningPipeline -ModuleName PRA.Cloud { $null }
        Mock Write-PcmLog -ModuleName PRA.Cloud { $lines.Add([string]$Message) }
        $context = @{ Config = @{ Polling = @{ CallTimeoutMinutes = 5 } }; CurrentPhase = 'Objects'; Action = 'Check' }
        { Start-PcmCallGuard -Context $context } | Should -Not -Throw
        $lines | Should -Contain 'The pipeline of this run cannot be reached on this PowerShell: a call to Microsoft 365 that does not answer in time ends the process (exit code 1) instead of stopping the run.'
        & (Get-Module PRA.Cloud) { $null -eq $script:CallGuard.Pipeline } | Should -BeTrue
        & (Get-Module PRA.Cloud) { $script:CallGuard.Watcher.InvocationStateInfo.State } | Should -Be 'Running'
    }
}

Describe 'Call time limit: a real run in its own process' {
    BeforeAll {
        $script:Pass = New-ChildWork
        $script:CheckPass = Invoke-Child -Scenario 'None' -Work $script:Pass -Arguments @('-Action', 'Check')
        $script:ConvertWork = New-ChildWork
        $script:ConvertPass = Invoke-Child -Scenario 'None' -Work $script:ConvertWork -Arguments @('-Action', 'Convert', '-Mode', 'Apply', '-Force')
        $script:ReadWork = New-ChildWork
        $script:Read = Invoke-Child -Scenario 'ReadHang' -Work $script:ReadWork -Arguments @('-Action', 'Check') -TimeoutSeconds 90
        $script:StuckWork = New-ChildWork
        $script:Stuck = Invoke-Child -Scenario 'ReadStuck' -Work $script:StuckWork -Arguments @('-Action', 'Check') -TimeoutSeconds 90
        $script:WriteWork = New-ChildWork
        $script:Write = Invoke-Child -Scenario 'WriteHang' -Work $script:WriteWork -Arguments @('-Action', 'Convert', '-Mode', 'Apply', '-Force') -TimeoutSeconds 90
        $script:BatchId = [regex]::Match($script:Write.Out, '(?m)^ *\S+ +Batch +(?<id>[0-9a-f]{8})\s*$').Groups['id'].Value
        $script:Journal = if ($script:BatchId) { Get-ChildJournal -Work $script:WriteWork -BatchId $script:BatchId }
        $script:Resume = if ($script:BatchId) { Invoke-Child -Scenario 'WriteDone' -Work $script:WriteWork -Arguments @('-Action', 'Convert', '-Mode', 'Apply', '-Batch', $script:BatchId, '-Force') }
        $script:ResumedJournal = if ($script:BatchId) { Get-ChildJournal -Work $script:WriteWork -BatchId $script:BatchId }
        $script:InterruptWork = New-ChildWork
        $script:Interrupted = Invoke-Child -Scenario 'Interrupt' -Work $script:InterruptWork -Arguments @('-Action', 'Convert', '-Mode', 'Apply', '-Force') -TimeoutSeconds 90
        $script:InterruptBatch = [regex]::Match($script:Interrupted.Out, '(?m)^ *\S+ +Batch +(?<id>[0-9a-f]{8})\s*$').Groups['id'].Value
        $script:InterruptJournal = if ($script:InterruptBatch) { Get-ChildJournal -Work $script:InterruptWork -BatchId $script:InterruptBatch }
    }

    It 'changes nothing for a run whose calls all answer: a Check and a Convert go through the guard and end as before' {
        $script:CheckPass.Exit | Should -Be 0 -Because $script:CheckPass.Out
        $script:CheckPass.Out | Should -Match '(?m)^\s*RESULT: PASS\s*$'
        $script:ConvertPass.Exit | Should -Be 0 -Because $script:ConvertPass.Out
        $script:ConvertPass.Out | Should -Match '2 done'
    }
    It 'removes the aliases and the proxy commands of the guard at the end of the run: a run that ends well, a run that fails, and a run that is stopped by the time limit' {
        foreach ($work in $script:Pass, $script:ConvertWork, $script:ReadWork, $script:WriteWork) {
            $marker = Join-Path $work 'guard-after.txt'
            Test-Path -LiteralPath $marker | Should -BeTrue -Because $work
            [IO.File]::ReadAllText($marker) | Should -BeExactly 'aliases=0 functions=0' -Because $work
        }
    }
    It 'ends a Check whose read never answers: exit code 1, FAIL, in a few seconds and not after the whole wait' {
        $script:Read.Finished | Should -BeTrue
        $script:Read.Exit | Should -Be 1 -Because $script:Read.Out
        $script:Read.Seconds | Should -BeLessThan 80
        $script:Read.Out | Should -Match '(?m)^\s*RESULT: FAIL\s*$'
        $script:Read.Out | Should -Match 'Check stopped - see the errors below'
    }
    It 'says in plain words what did not answer, in which step and what to do' {
        $script:Read.Out | Should -Match ([regex]::Escape('Microsoft 365 did not answer within 3 seconds during Tenant prerequisites (Get-OrganizationConfig); nothing more was changed. Check https://status.cloud.microsoft and the service health of your tenant, then run the same command again.'))
    }
    It 'ends with the next steps: nothing changed, the service health, and the same command' {
        $script:Read.Out | Should -Match 'Check only reads: nothing was changed\. Check the service health of Microsoft 365 \(https://status\.cloud\.microsoft\), then run the same command again:'
        $script:Read.Out | Should -Match ([regex]::Escape('.\Invoke-PraCloudMailbox.ps1 -Action Check -ConfigPath'))
    }
    It 'stops the Exchange Online commands the way Ctrl+C does and leaves the log with the message' {
        $log = @(Get-ChildItem -LiteralPath (Join-Path $script:ReadWork 'logs') -Filter '*.log' | Select-Object -First 1)
        $log.Count | Should -Be 1
        (Get-Content -LiteralPath $log[0].FullName -Raw) | Should -Match 'Microsoft 365 did not answer within 3 seconds during Tenant prerequisites'
    }
    It 'ends the process, with the same message, the next steps and exit code 1, when the call cannot be interrupted' {
        $script:Stuck.Finished | Should -BeTrue
        $script:Stuck.Exit | Should -Be 1
        $script:Stuck.Seconds | Should -BeLessThan 80
        $script:Stuck.Err | Should -Match ([regex]::Escape('Microsoft 365 did not answer within 3 seconds during Tenant prerequisites (Get-OrganizationConfig); nothing more was changed.'))
        $script:Stuck.Err | Should -Match 'The call could not be interrupted, so this run was ended at once'
        $script:Stuck.Err | Should -Match '(?s)Next steps:\s+Check only reads: nothing was changed\..*-Action Check'
        (Get-Content -LiteralPath @(Get-ChildItem -LiteralPath (Join-Path $script:StuckWork 'logs') -Filter '*.log')[0].FullName -Raw) | Should -Match 'The call could not be interrupted'
    }
    It 'ends the text of a process that is ended at once with a RESULT line, on the console and in the log' {
        $script:Stuck.Err | Should -Match '(?m)^RESULT: FAIL\s*$'
        (Get-Content -LiteralPath @(Get-ChildItem -LiteralPath (Join-Path $script:StuckWork 'logs') -Filter '*.log')[0].FullName -Raw) | Should -Match '(?m)^RESULT: FAIL\s*$'
    }
    It 'ends a Convert whose change never answers with the batch kept in the journal, and says that the change is unknown' {
        $script:Write.Finished | Should -BeTrue
        $script:Write.Exit | Should -Be 1 -Because $script:Write.Out
        $script:Write.Out | Should -Match 'Microsoft 365 did not answer within 3 seconds during Users: source of authority and Exchange plan \(Set-MailUser\)\. That last change may or may not have been applied'
        $script:Write.Out | Should -Match 'What this run already did is kept in the journal \(batch [0-9a-f]{8}\) and is never repeated\. Check the service health of Microsoft 365'
        $script:Write.Out | Should -Match ('-Action Convert -Mode Apply -Batch {0} -ConfigPath' -f $script:BatchId)
        $script:Write.Out | Should -Match '(?m)^\s*RESULT: FAIL\s*$'
    }
    It 'leaves the journal as a crash would, but not open for ever: the batch partial, no object done or failed, the stop written, no step recorded for the change that did not answer' {
        $script:BatchId | Should -Match '^[0-9a-f]{8}$'
        $script:Journal.Batch.status | Should -Be 'Partial'
        $script:Journal.Batch.note | Should -Match 'ended before it finished'
        @($script:Journal.Items).Count | Should -Be 2
        @($script:Journal.Items | Where-Object { $_.status -in @('Done', 'Failed') }).Count | Should -Be 0
        @($script:Journal.Events | Where-Object { $_.step -eq 'MicrosoftTimeout' -and $_.outcome -eq 'Fail' }).Count | Should -Be 1
        @($script:Journal.Events | Where-Object { $_.step -eq 'GuidCleared' }).Count | Should -Be 0
    }
    It 'ends a run that is stopped from outside (Ctrl+C) with exit code 1, FAIL and the command that continues the batch, never PASS' {
        $script:Interrupted.Finished | Should -BeTrue
        $script:Interrupted.Exit | Should -Be 1 -Because $script:Interrupted.Out
        $script:Interrupted.Out | Should -Match '(?m)^\s*RESULT: FAIL\s*$'
        $script:Interrupted.Out | Should -Not -Match '(?m)^\s*RESULT: PASS\s*$'
        $script:Interrupted.Out | Should -Not -Match 'Convert finished'
        $script:Interrupted.Out | Should -Match 'the run was interrupted before it finished'
        $script:Interrupted.Out | Should -Match ('-Action Convert -Mode Apply -Batch {0} -ConfigPath' -f $script:InterruptBatch)
    }
    It 'counts the objects of an interrupted Convert as pending and leaves the batch partial, not running and not complete' {
        $script:InterruptBatch | Should -Match '^[0-9a-f]{8}$'
        $script:InterruptJournal.Batch.status | Should -Be 'Partial'
        @($script:InterruptJournal.Items | Where-Object { $_.status -eq 'Done' }).Count | Should -Be 0
        $script:Interrupted.Out | Should -Match '0 done .{1,3} 2 pending .{1,3} 1 error'
    }
    It 'is continued by the command of the next steps: the same batch ends complete, every object done once' {
        $script:Resume.Exit | Should -Be 0 -Because $script:Resume.Out
        $script:Resume.Out | Should -Match "Resuming batch $($script:BatchId)"
        $script:ResumedJournal.Batch.status | Should -Be 'Complete'
        @($script:ResumedJournal.Items | Where-Object { $_.status -eq 'Done' }).Count | Should -Be 2
        @($script:ResumedJournal.Events | Where-Object { $_.step -eq 'GuidCleared' }).Count | Should -Be 2
    }
}
