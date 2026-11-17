#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
}

AfterAll {
    Restore-PcmGlobalState $script:GlobalBefore
}


Describe 'Graph transport lifecycle (offline context model)' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
        $script:Package = New-PcmTestPackage -Repository ([IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..')))
        Import-Module (Join-Path $script:Package 'module\PRA.Common.psm1') -Force
        Import-Module (Join-Path $script:Package 'module\PRA.Cloud.psm1') -Force
        . (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
        function global:Connect-MgGraph {
            [CmdletBinding()]
            param($ClientId, $TenantId, $CertificateThumbprint, $ContextScope, [switch]$NoWelcome)
        }
        function global:Disconnect-MgGraph { [CmdletBinding()]param() }
        . (Join-Path $PSScriptRoot 'ExoSession.Fake.ps1')
    }
    BeforeEach {
        Reset-PraFakeGraphTransport
        & (Get-Module PRA.Cloud) {
            $script:GraphRequestOwner = $null
            $script:GraphTransportReady = $false
        }
        $global:PraFakeGraphRequest.MaxRetry = 3
        $context = @{ Config = @{ Cloud = @{ TenantId = 'synthetic'; Organization = 'synthetic'; AppId = 'synthetic'; CertificateThumbprint = '12345678' }; Scope = @{} }
            Approval = { param($Target, $Operation) $true } }
        Mock Import-Module -ModuleName PRA.Cloud { }
        Mock Get-ChildItem -ModuleName PRA.Cloud { [pscustomobject]@{ HasPrivateKey = $true; NotAfter = (Get-Date).AddYears(1) } }
        Reset-PraFakeExoSession
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $Organization -AppId $AppId }
        Mock Connect-MgGraph -ModuleName PRA.Cloud {
            $global:PraFakeGraphRequest.MaxRetry | Should -Be 0
            $ContextScope | Should -BeExactly Process
            $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = $ClientId }
        }
        Mock Disconnect-MgGraph -ModuleName PRA.Cloud { $global:PraFakeGraphAuth = $null }
    }
    AfterAll {
        foreach ($name in @('Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext', 'Reset-PraFakeGraphTransport', 'Connect-MgGraph', 'Disconnect-MgGraph')) {
            Remove-Item -Path "function:$name" -ErrorAction SilentlyContinue
        }
        Remove-PraFakeExoSession
    }
    It 'initializes before authentication and restores only after disconnect, preserving other request settings' {
        $before = $global:PraFakeGraphRequest | ConvertTo-Json -Compress
        Connect-PcmCloud -Context $context | Should -Not -BeNullOrEmpty
        $global:PraFakeGraphClientGeneration | Should -Be 2
        Disconnect-PcmCloud -Context $context
        ($global:PraFakeGraphRequest | ConvertTo-Json -Compress) | Should -BeExactly $before
        $global:PraFakeGraphClientGeneration | Should -Be 3
        Should -Invoke Connect-MgGraph -ModuleName PRA.Cloud -Times 1 -Exactly
    }
    It 'rejects an existing SDK context even with MaxRetry=<Retry>, without replacing its cached client' -TestCases @(@{ Retry = 0 }, @{ Retry = 3 }) {
        param($Retry)
        $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = 'external' }
        $global:PraFakeGraphRequest.MaxRetry = $Retry
        { Connect-PcmCloud -Context $context } | Should -Throw '*existing SDK session*'
        $global:PraFakeGraphClientGeneration | Should -Be 1
        $global:PraFakeGraphAuth.ClientId | Should -BeExactly external
        Should -Invoke Connect-MgGraph -ModuleName PRA.Cloud -Times 0 -Exactly
        Should -Invoke Disconnect-MgGraph -ModuleName PRA.Cloud -Times 0 -Exactly
    }
    It 'fails before authentication when <Fault> prevents initialization' -TestCases @(
        @{ Fault = 'MissingCommand' }, @{ Fault = 'InitialReadFails' }, @{ Fault = 'InitialContextMissing' }
        @{ Fault = 'SetFails' }, @{ Fault = 'ReadbackNonzero' }, @{ Fault = 'ReadbackMissing' }) {
        param($Fault)
        switch ($Fault) {
            MissingCommand { Mock Get-Command -ModuleName PRA.Cloud { throw 'Missing SDK context command' } -ParameterFilter { $Name -eq 'Set-MgRequestContext' } }
            InitialReadFails { Mock Get-MgRequestContext -ModuleName PRA.Cloud { throw 'Initial SDK context read failed' } }
            InitialContextMissing { $global:PraFakeGraphRequest = $null }
            SetFails { Mock Set-MgRequestContext -ModuleName PRA.Cloud { throw 'SDK setting failed' } }
            ReadbackNonzero { Mock Set-MgRequestContext -ModuleName PRA.Cloud { $global:PraFakeGraphRequest } }
            ReadbackMissing { Mock Set-MgRequestContext -ModuleName PRA.Cloud { $global:PraFakeGraphRequest = $null } }
        }
        { Connect-PcmCloud -Context $context -WarningAction SilentlyContinue } | Should -Throw
        Should -Invoke Connect-MgGraph -ModuleName PRA.Cloud -Times 0 -Exactly
    }
    It 'cleans up failed authentication without resetting an authenticated cached client' {
        Mock Connect-MgGraph -ModuleName PRA.Cloud {
            $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = 'synthetic' }
            throw 'Synthetic authentication failure'
        }
        { Connect-PcmCloud -Context $context } | Should -Throw '*Synthetic authentication failure*'
        $global:PraFakeGraphAuth | Should -BeNullOrEmpty
        $global:PraFakeGraphRequest.MaxRetry | Should -Be 3
    }
    It 'does not restore request settings when SDK disconnect fails' {
        Connect-PcmCloud -Context $context | Out-Null
        Mock Disconnect-MgGraph -ModuleName PRA.Cloud { throw 'Synthetic disconnect failure' }
        { Disconnect-PcmCloud -Context $context -WarningAction SilentlyContinue } | Should -Throw '*Microsoft Graph sign-out: Synthetic disconnect failure*'
        $global:PraFakeGraphClientGeneration | Should -Be 2
        $global:PraFakeGraphRequest.MaxRetry | Should -Be 0
        & (Get-Module PRA.Cloud) { $script:GraphTransportReady } | Should -BeFalse
    }
    It 'refuses an Exchange Online session that already exists, and signs nothing out' {
        $null = Add-PraFakeExoConnection -TenantId 'contoso.onmicrosoft.com' -ConnectionId 'foreign'
        { Connect-PcmCloud -Context $context } | Should -Throw '*An Exchange Online session already exists in this PowerShell window*started with -NoProfile if your profile signs in to Exchange Online (nothing was signed out).'
        Should -Invoke Connect-ExchangeOnline -ModuleName PRA.Cloud -Times 0 -Exactly
        $global:PraFakeExo.Disconnects.Count | Should -Be 0
        @(Get-ConnectionInformation).Count | Should -Be 1
    }
    It 'refuses an Exchange Online module that cannot sign out one connection' {
        Mock Get-Command -ModuleName PRA.Cloud { [pscustomobject]@{ Parameters = @{ Confirm = $true } } } -ParameterFilter { $Name -eq 'Disconnect-ExchangeOnline' }
        Mock Get-Command -ModuleName PRA.Cloud { Microsoft.PowerShell.Core\Get-Command @PSBoundParameters }
        { Connect-PcmCloud -Context $context } | Should -Throw '*cannot sign out one connection*'
        Should -Invoke Connect-ExchangeOnline -ModuleName PRA.Cloud -Times 0 -Exactly
    }
    It 'refuses an Exchange Online session of another tenant, and signs out only its own connection' {
        $context.Config.Cloud.TenantId = '33333333-3333-3333-3333-333333333333'
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId 'other.onmicrosoft.com' -AppId $AppId -ConnectionId 'own' }
        { Connect-PcmCloud -Context $context } | Should -Throw "*The Exchange Online session is in tenant other.onmicrosoft.com, and Cloud.TenantId is 33333333-3333-3333-3333-333333333333. Nothing was changed. Check Cloud.Organization and Cloud.TenantId in the configuration file, then run the command again."
        @($global:PraFakeExo.Disconnects) | Should -Be @('own')
    }
    It 'says that no connection was found, and does not tell to close sessions' {
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { }
        $failure = try { Connect-PcmCloud -Context $context; '' } catch { $_.Exception.Message }
        $failure | Should -BeLike 'No Exchange Online connection was found in this PowerShell window after the sign-in.*'
        $failure | Should -Not -BeLike '*Close the other sessions*'
    }
    It 'refuses a Security and Compliance session of another tenant, and signs it out with the others' {
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $Organization -AppId $AppId -ConnectionId 'exchange' }
        Mock Connect-IPPSSession -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId 'other.onmicrosoft.com' -AppId $AppId -ConnectionId 'compliance' }
        Connect-PcmCloud -Context $context | Out-Null
        { Connect-PcmCompliance -Context $context } | Should -Throw '*The Security & Compliance session is in tenant other.onmicrosoft.com, and Cloud.TenantId is synthetic. Nothing was changed.*'
        $context.ContainsKey('ComplianceConnected') | Should -BeFalse
        Disconnect-PcmCloud -Context $context
        @($global:PraFakeExo.Disconnects) | Should -Be @('exchange', 'compliance')
    }
    It 'accepts the tenant GUID or the organisation domain, in any letter case' -TestCases @(
        @{ Tenant = 'AAAAAAAA-0000-0000-0000-000000000001' }, @{ Tenant = 'aaaaaaaa-0000-0000-0000-000000000001' }, @{ Tenant = 'CONTOSO.onmicrosoft.com' }) {
        param($Tenant)
        $context.Config.Cloud.TenantId = 'aaaaaaaa-0000-0000-0000-000000000001'
        $context.Config.Cloud.Organization = 'contoso.ONMICROSOFT.com'
        $global:PraFakeExo.InteractiveTenant = $Tenant
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $global:PraFakeExo.InteractiveTenant -AppId 'SYNTHETIC' }
        Connect-PcmCloud -Context $context | Should -Not -BeNullOrEmpty
        Disconnect-PcmCloud -Context $context
    }
    It 'refuses a session opened by another application' {
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $Organization -AppId 'another-app' -ConnectionId 'own' }
        { Connect-PcmCloud -Context $context } | Should -Throw '*opened by application another-app, and Cloud.AppId is synthetic*'
        @($global:PraFakeExo.Disconnects) | Should -Be @('own')
    }
    It 'signs out only the connection of the tool, and leaves the others open' {
        Connect-PcmCloud -Context $context | Out-Null
        $own = @($context['ExoConnectionIds'])
        $own.Count | Should -Be 1
        $null = Add-PraFakeExoConnection -TenantId 'contoso.onmicrosoft.com' -ConnectionId 'added-later'
        Disconnect-PcmCloud -Context $context
        @($global:PraFakeExo.Disconnects) | Should -Be $own
        @(Get-ConnectionInformation | ForEach-Object { $_.ConnectionId }) | Should -Be @('added-later')
        $context['ExoConnected'] | Should -BeFalse
    }
    It 'signs out the Security and Compliance connection that the run opened, and only that one' {
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $Organization -AppId $AppId -ConnectionId 'exchange' }
        Mock Connect-IPPSSession -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $Organization -AppId $AppId -ConnectionId 'compliance' }
        Connect-PcmCloud -Context $context | Out-Null
        Connect-PcmCompliance -Context $context
        @($context['ExoConnectionIds']) | Should -Be @('exchange', 'compliance')
        $null = Add-PraFakeExoConnection -TenantId 'contoso.onmicrosoft.com' -ConnectionId 'foreign'
        Disconnect-PcmCloud -Context $context
        @($global:PraFakeExo.Disconnects) | Should -Be @('exchange', 'compliance')
        @(Get-ConnectionInformation | ForEach-Object { $_.ConnectionId }) | Should -Be @('foreign')
    }
    It 'reports a sign-out that fails, and still signs out of Graph' {
        Connect-PcmCloud -Context $context | Out-Null
        $global:PraFakeExo.FailDisconnect = 'Synthetic sign-out failure'
        { Disconnect-PcmCloud -Context $context } | Should -Throw '*Exchange Online sign-out: Synthetic sign-out failure*'
        Should -Invoke Disconnect-MgGraph -ModuleName PRA.Cloud -Times 1 -Exactly
        $global:PraFakeGraphAuth | Should -BeNullOrEmpty
    }
    It 'signs out the connection that a failed sign-in left, and reports the failure of the sign-in' {
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud { $null = Add-PraFakeExoConnection -TenantId $Organization -AppId $AppId -ConnectionId 'partial'; throw 'Synthetic sign-in failure' }
        { Connect-PcmCloud -Context $context } | Should -Throw '*Synthetic sign-in failure*'
        @($global:PraFakeExo.Disconnects) | Should -Be @('partial')
        @(Get-ConnectionInformation).Count | Should -Be 0
    }
    It 'refuses to sign out when the connections of the run are not known' {
        $context['ExoConnected'] = $true
        { Disconnect-PcmCloud -Context $context } | Should -Throw '*Exchange Online sign-out: the Exchange Online connections of this run are not known*'
        $global:PraFakeExo.Disconnects.Count | Should -Be 0
    }
    It 'does not relinquish ownership when restoration cannot be verified' {
        Connect-PcmCloud -Context $context | Out-Null
        Mock Set-MgRequestContext -ModuleName PRA.Cloud { $global:PraFakeGraphRequest }
        Disconnect-PcmCloud -Context $context -WarningAction SilentlyContinue
        & (Get-Module PRA.Cloud) { $null -ne $script:GraphRequestOwner } | Should -BeTrue
        & (Get-Module PRA.Cloud) { $script:GraphTransportReady } | Should -BeFalse
    }
}

Describe 'Entra Connect check (authenticated, read-only)' {
    BeforeAll { Import-Module (Join-Path $script:Package 'module\PRA.Cloud.psm1') -Force }

    It 'answers Ok when the remote job answers, with a session that has a time limit' {
        Mock Invoke-Command -ModuleName PRA.Cloud { Start-Job -ScriptBlock { [pscustomobject]@{ Computer = 'AADC'; CycleEnabled = $true } } }
        $answer = Test-PcmSyncServer -Server 'aadc'
        $answer.Ok | Should -BeTrue
        $answer.Kind | Should -BeExactly 'Ok'
        Should -Invoke Invoke-Command -ModuleName PRA.Cloud -Times 1 -Exactly -ParameterFilter { $ComputerName -eq 'aadc' -and $AsJob -and $null -ne $SessionOption }
        @(Get-Job).Count | Should -Be 0
    }

    It 'says in plain words what is wrong (<Kind>)' -TestCases @(
        @{ Kind = 'Denied'; Message = 'Access is denied.'; Text = '*no right on aadc*ADSyncAdmins*' }
        @{ Kind = 'Denied'; Message = 'Accès refusé.'; Text = '*no right on aadc*' }
        @{ Kind = 'NoModule'; Message = 'The specified module ADSync was not loaded because no valid module file was found'; Text = '*not installed there*EntraConnect.Server*' }
        @{ Kind = 'Unreachable'; Message = 'Connecting to remote server aadc failed with the following error message: WinRM cannot process the request'; Text = '*does not answer to PowerShell remoting*port 5985*' }
        @{ Kind = 'Failed'; Message = 'Something else happened.'; Text = '*could not be checked on aadc: Something else happened*' }) {
        param($Kind, $Message, $Text)
        $global:PcmProbeMessage = $Message
        Mock Invoke-Command -ModuleName PRA.Cloud { throw $global:PcmProbeMessage }
        $answer = Test-PcmSyncServer -Server 'aadc'
        $answer.Ok | Should -BeFalse
        $answer.Kind | Should -BeExactly $Kind
        $answer.Reason | Should -BeLike $Text
        $answer.Technical | Should -Match ([regex]::Escape($Message.TrimEnd('.')))
        Remove-Variable -Name PcmProbeMessage -Scope Global
    }

    It 'says in plain words what is wrong when the remote job fails (<Kind>)' -TestCases @(
        @{ Kind = 'Denied'; Body = { throw 'Access is denied.' }; Text = '*no right on aadc*' }
        @{ Kind = 'Failed'; Body = { throw 'Something else happened.' }; Text = '*could not be checked on aadc: Something else happened*' }) {
        param($Kind, $Body, $Text)
        $global:PcmProbeJob = $Body
        Mock Invoke-Command -ModuleName PRA.Cloud { Start-Job -ScriptBlock $global:PcmProbeJob }
        $answer = Test-PcmSyncServer -Server 'aadc'
        Remove-Variable -Name PcmProbeJob -Scope Global
        $answer.Ok | Should -BeFalse
        $answer.Kind | Should -BeExactly $Kind
        $answer.Reason | Should -BeLike $Text
        @(Get-Job).Count | Should -Be 0
    }

    It 'treats an empty answer as a failure' {
        Mock Invoke-Command -ModuleName PRA.Cloud { Start-Job -ScriptBlock { } }
        $answer = Test-PcmSyncServer -Server 'aadc'
        $answer.Ok | Should -BeFalse
        $answer.Kind | Should -BeExactly 'Failed'
    }

    It 'stops a remote job that does not answer in time, and says so in plain words' {
        Mock Invoke-Command -ModuleName PRA.Cloud { Start-Job -ScriptBlock { Start-Sleep -Seconds 120 } }
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $answer = Test-PcmSyncServer -Server 'aadc' -TimeoutSeconds 5
        $watch.Stop()
        $answer.Ok | Should -BeFalse
        $answer.Kind | Should -BeExactly 'Unreachable'
        $answer.Reason | Should -BeLike 'aadc does not answer to PowerShell remoting (WinRM, port 5985) within 5 seconds*'
        $watch.Elapsed.TotalSeconds | Should -BeLessThan 60
        @(Get-Job).Count | Should -Be 0
    }

    It 'only reads: the probe holds no command that changes or starts anything' {
        $probe = & (Get-Module PRA.Cloud) { $script:PcmSyncProbe.ToString() }
        $probe | Should -Match 'Get-ADSyncScheduler'
        $probe | Should -Not -Match '(?i)\b(Set|Start|Invoke|Add|Remove|New|Stop|Suspend|Resume|Disable|Enable)-ADSync'
        $probe | Should -Not -Match '(?i)\b(Set|Start|Stop|Remove|New|Restart)-(Service|Item|Content)\b'
    }
}

Describe 'Retry-After of a direct Graph request' {
    BeforeAll {
        Import-Module (Join-Path $script:Package 'module\PRA.Cloud.psm1') -Force
        function global:Invoke-MgGraphRequest { [CmdletBinding()] param([string]$Method = 'GET', [string]$Uri, [object]$Body, [string]$ContentType, [string]$OutputType, [System.Collections.IDictionary]$Headers) }
        function global:New-PcmThrottleFailure {
            param([int]$Code, [string]$RetryAfter = '')
            $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Code)
            if ($RetryAfter) { $null = $response.Headers.TryAddWithoutValidation('retry-after', $RetryAfter) }
            $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Code.", $response)
            [System.Management.Automation.ErrorRecord]::new($exception, 'FakeGraph', 'InvalidOperation', $null)
        }
    }
    BeforeEach {
        $global:PcmSlept = [System.Collections.Generic.List[int]]::new()
        $global:PcmGraphFailures = [System.Collections.Generic.Queue[object]]::new()
        Mock Assert-PcmGraphTransport -ModuleName PRA.Cloud { }
        Mock Enter-PcmCall -ModuleName PRA.Cloud { }
        Mock Exit-PcmCall -ModuleName PRA.Cloud { }
        Mock Start-Sleep -ModuleName PRA.Cloud { $global:PcmSlept.Add([int]$Seconds) }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { if ($global:PcmGraphFailures.Count) { throw $global:PcmGraphFailures.Dequeue() }; [pscustomobject]@{ ok = $true } }
    }
    AfterAll {
        foreach ($name in 'Invoke-MgGraphRequest', 'New-PcmThrottleFailure') { Remove-Item -Path "Function:\$name" -ErrorAction SilentlyContinue }
        Remove-Variable -Name PcmSlept, PcmGraphFailures -Scope Global -ErrorAction SilentlyContinue
    }

    It 'waits for the Retry-After of the answer, whatever the status (<Code>), at least the usual pause' -TestCases @(
        @{ Code = 429 }, @{ Code = 502 }, @{ Code = 503 }, @{ Code = 504 }) {
        param($Code)
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code $Code -RetryAfter '17'))
        (Invoke-PcmGraph -Uri 'v1.0/users/x').ok | Should -BeTrue
        @($global:PcmSlept) | Should -Be @(17)
    }

    It 'reads a Retry-After that is an HTTP date, as the time left until it' {
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter ([datetimeoffset]::UtcNow.AddSeconds(46).ToString('r', [Globalization.CultureInfo]::InvariantCulture))))
        $null = Invoke-PcmGraph -Uri 'v1.0/users/x'
        $global:PcmSlept.Count | Should -Be 1
        $global:PcmSlept[0] | Should -BeGreaterThan 30
        $global:PcmSlept[0] | Should -BeLessOrEqual 47
    }

    It 'keeps the usual backoff when there is no Retry-After, a smaller one, a date in the past or a text that is not one' -TestCases @(
        @{ Header = ''; Expected = 5 }, @{ Header = '2'; Expected = 5 }, @{ Header = 'soon'; Expected = 5 }, @{ Header = '0'; Expected = 5 }
        @{ Header = ([datetimeoffset]::UtcNow.AddMinutes(-5).ToString('r', [Globalization.CultureInfo]::InvariantCulture)); Expected = 5 }) {
        param($Header, $Expected)
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 503 -RetryAfter $Header))
        $null = Invoke-PcmGraph -Uri 'v1.0/users/x'
        @($global:PcmSlept) | Should -Be @($Expected)
    }

    It 'grows the usual backoff with the attempt, and takes the larger of the two each time' {
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429))
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter '3'))
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter '40'))
        $null = Invoke-PcmGraph -Uri 'v1.0/users/x'
        @($global:PcmSlept) | Should -Be @(5, 10, 40)
    }

    It 'never waits more than two minutes' {
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter '900'))
        $null = Invoke-PcmGraph -Uri 'v1.0/users/x'
        @($global:PcmSlept) | Should -Be @(120)
    }

    It 'says in the log that Microsoft Graph asks to wait, when the answer makes the tool wait longer than the usual pause' {
        $global:PcmLogged = [System.Collections.Generic.List[string]]::new()
        Mock Write-PcmLog -ModuleName PRA.Cloud { $global:PcmLogged.Add([string]$Message) }
        & (Get-Module PRA.Cloud) { $script:GraphRequestOwner = @{} }
        try {
            $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter '17'))
            $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter '2'))
            $null = Invoke-PcmGraph -Uri 'v1.0/users/x'
        } finally { & (Get-Module PRA.Cloud) { $script:GraphRequestOwner = $null } }
        @($global:PcmLogged) | Should -Be @('Microsoft Graph asks to wait 17 s')
        Remove-Variable -Name PcmLogged -Scope Global
    }

    It 'does not send a request again that must not be resent' {
        $global:PcmGraphFailures.Enqueue((New-PcmThrottleFailure -Code 429 -RetryAfter '9'))
        { Invoke-PcmGraph -Uri 'v1.0/users/x' -NoRetry } | Should -Throw
        $global:PcmSlept.Count | Should -Be 0
    }
}

# Optional installed-SDK qualification: the normal offline gate does not install Graph or sign in.
Describe 'Installed Graph SDK transport (in-memory HTTP, no authentication)' -Skip:($PSVersionTable.PSEdition -ne 'Core' -or -not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    BeforeAll {
        $sdkSession = $null
        $savedRetry = $null
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module (Join-Path $script:Package 'module\PRA.Cloud.psm1') -Force
            # SDK 2.41 isolates Core in its own load context; resolve the session through
            # the imported cmdlet's metadata rather than PowerShell's default type lookup.
            $sdkAssembly = (Get-Command Get-MgRequestContext).ImplementingType.Assembly
            $initializer = $sdkAssembly.GetType('Microsoft.Graph.PowerShell.Authentication.Common.GraphSessionInitializer', $true)
            $sdkCore = $initializer.GetMethod('CreateInstance', [Reflection.BindingFlags]'Static,Public,NonPublic').ReturnType.Assembly
            $sdkSession = $sdkCore.GetType('Microsoft.Graph.PowerShell.Authentication.GraphSession', $true).GetProperty('Instance').GetValue($null)
            if ($sdkSession.AuthContext) { throw 'Offline SDK qualification refuses an authenticated process.' }
            $savedRetry = (Get-MgRequestContext).MaxRetry
        } catch {
            Write-Host ('Offline SDK fixture setup failed: {0}' -f $_.Exception.Message)
            throw
        }
        if (-not ('PraGraphOfflineHandler' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
public sealed class PraGraphOfflineHandler : HttpMessageHandler {
    public int Count;
    public int[] Statuses;
    public readonly List<string> Methods = new List<string>();
    public PraGraphOfflineHandler(int[] statuses) { Statuses = statuses; }
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
        if (request.RequestUri.Host != "graph.microsoft.com") throw new InvalidOperationException("Unexpected fixture host.");
        if (Count >= 20) throw new InvalidOperationException("Offline fixture request limit exceeded.");
        Methods.Add(request.Method.Method);
        int status = Statuses[Math.Min(Count++, Statuses.Length - 1)];
        var response = new HttpResponseMessage((HttpStatusCode)status) { RequestMessage = request };
        response.Content = new StringContent(status >= 400 ? "{\"error\":{\"code\":\"OfflineFixture\",\"message\":\"Synthetic HTTP failure\"}}" : "{}", Encoding.UTF8, "application/json");
        response.Headers.TryAddWithoutValidation("Retry-After", "1");
        return Task.FromResult(response);
    }
}
'@
        }
        function New-OfflineSdkClient {
            param([int[]]$Statuses)
            # Test-only reflection keeps the SDK's own complete builder/middleware; ONLY the final
            # network handler is replaced, before use. A null auth provider cannot obtain a token.
            $requestContext = (Get-MgRequestContext).PSObject.BaseObject
            $builder = $sdkAssembly.GetType('Microsoft.Graph.PowerShell.Authentication.Helpers.HttpHelpers', $true).GetMethod('GetGraphHttpClient', [Reflection.BindingFlags]'Static,NonPublic')
            if ($builder) {
                $client = $builder.Invoke($null, [object[]]@($null, $requestContext))
            } else {
                # 2.41 removed the null-provider overload. Use its native GraphClientFactory
                # and the same middleware order/options, without creating token credentials.
                $loadContext = [System.Runtime.Loader.AssemblyLoadContext]::GetLoadContext($sdkCore)
                $graphCore = $loadContext.LoadFromAssemblyName([Reflection.AssemblyName]::new('Microsoft.Graph.Core'))
                $kiota = $loadContext.LoadFromAssemblyName([Reflection.AssemblyName]::new('Microsoft.Kiota.Http.HttpClientLibrary'))
                $retryOption = [Activator]::CreateInstance($kiota.GetType('Microsoft.Kiota.Http.HttpClientLibrary.Middleware.Options.RetryHandlerOption', $true))
                $retryOption.Delay = $requestContext.RetryDelay
                $retryOption.MaxRetry = $requestContext.MaxRetry
                $retryOption.RetriesTimeLimit = $requestContext.RetriesTimeLimit
                $handlers = [Collections.Generic.List[Net.Http.DelegatingHandler]]::new()
                $authType = $sdkCore.GetType('Microsoft.Graph.PowerShell.Authentication.Core.Http.AuthenticationHandler', $true)
                $authConstructor = @($authType.GetConstructors() | Where-Object { $_.GetParameters().Count -eq 1 })[0]
                $handlers.Add($authConstructor.Invoke([object[]]@($null)))
                foreach ($name in 'NationalCloudHandler', 'ODataQueryOptionsHandler', 'HttpVersionHandler') {
                    $handlers.Add([Activator]::CreateInstance($sdkAssembly.GetType("Microsoft.Graph.PowerShell.Authentication.Handlers.$name", $true)))
                }
                $retryType = $kiota.GetType('Microsoft.Kiota.Http.HttpClientLibrary.Middleware.RetryHandler', $true)
                $handlers.Add($retryType.GetConstructors()[0].Invoke([object[]]@($retryOption)))
                $redirectType = $kiota.GetType('Microsoft.Kiota.Http.HttpClientLibrary.Middleware.RedirectHandler', $true)
                $handlers.Add($redirectType.GetConstructors()[0].Invoke([object[]]@([Type]::Missing)))
                $handlers.Add([Activator]::CreateInstance($sdkAssembly.GetType('Microsoft.Graph.PowerShell.Authentication.Handlers.RequestHeaderHandler', $true)))
                $factory = @($graphCore.GetType('Microsoft.Graph.GraphClientFactory', $true).GetMethods() | Where-Object {
                    $_.Name -eq 'Create' -and $_.GetParameters().Count -eq 6 -and
                    $_.GetParameters()[0].ParameterType -eq [Collections.Generic.IEnumerable[Net.Http.DelegatingHandler]]
                })[0]
                $client = $factory.Invoke($null, [object[]]@($handlers, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing))
                $client.Timeout = $requestContext.ClientTimeout
            }
            $field = [Net.Http.HttpMessageInvoker].GetField('_handler', [Reflection.BindingFlags]'Instance,NonPublic')
            $handler = $field.GetValue($client)
            while ($handler.InnerHandler -is [Net.Http.DelegatingHandler]) { $handler = $handler.InnerHandler }
            $handler.InnerHandler.Dispose()
            $terminal = [PraGraphOfflineHandler]::new($Statuses)
            $handler.InnerHandler = $terminal
            $sdkSession.GraphHttpClient = $client
            $sdkSession.Environment = (Get-MgEnvironment -Name Global).PSObject.BaseObject
            $sdkSession.AuthContext = [Activator]::CreateInstance($sdkAssembly.GetType('Microsoft.Graph.PowerShell.Authentication.AuthContext', $true))
            [pscustomobject]@{ Client = $client; Terminal = $terminal }
        }
    }
    BeforeEach {
        $sdkSession.AuthContext = $null
        $sdkSession.GraphHttpClient = $null
        $null = Set-MgRequestContext -MaxRetry 3
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $true }
        & (Get-Module PRA.Cloud) {
            $script:GraphRequestOwner = $null
            $script:GraphTransportReady = $false
        }
    }
    AfterEach {
        $sdkSession.AuthContext = $null
        if ($sdkSession.GraphHttpClient) { $sdkSession.GraphHttpClient.Dispose() }
        $sdkSession.GraphHttpClient = $null
    }
    AfterAll {
        if ($null -ne $sdkSession -and $null -ne $savedRetry) {
            $sdkSession.AuthContext = $null
            $null = Set-MgRequestContext -MaxRetry $savedRetry
        }
        Remove-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    }
    It 'demonstrates default SDK middleware really resends DELETE after HTTP <Status> (not a mocked cmdlet)' -TestCases @(@{ Status = 503 }, @{ Status = 504 }) {
        param($Status)
        $fixture = New-OfflineSdkClient @($Status, 204)
        $null = Microsoft.Graph.Authentication\Invoke-MgGraphRequest -Method DELETE -Uri 'v1.0/users/offline-id' -ErrorAction Stop
        $fixture.Terminal.Count | Should -Be 2
        $fixture.Terminal.Methods | Should -Be @('DELETE', 'DELETE')
    }
    It 'sends exactly one DELETE through the real SDK on HTTP <Status> after initialization (cached client=<Warm>)' -TestCases @(
        foreach ($status in 503, 504, 400, 403) { foreach ($warm in $false, $true) { @{ Status = $status; Warm = $warm } } }) {
        param($Status, $Warm)
        & (Get-Module PRA.Cloud) { Initialize-PcmGraphTransport -Context @{} }
        $fixture = New-OfflineSdkClient $(if ($Warm) { @(204, $Status, 204) } else { @($Status, 204) })
        if ($Warm) { $null = Invoke-PcmGraph -Method GET -Uri 'v1.0/users/offline-id' }
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
        $failure = try { Invoke-PcmGraph -Method DELETE -Uri 'v1.0/users/offline-id' -NoRetry } catch { $_ }
        $failure | Should -Not -BeNullOrEmpty
        $failure.FullyQualifiedErrorId | Should -Match 'InvokeGraphHttpResponseException'
        [int]$failure.Exception.Response.StatusCode | Should -Be $Status
        $fixture.Terminal.Count | Should -Be $(if ($Warm) { 2 } else { 1 })
        $fixture.Terminal.Methods | Should -Be $(if ($Warm) { @('GET', 'DELETE') } else { @('DELETE') })
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
        (Get-MgRequestContext).MaxRetry | Should -Be 0
    }
    It 'rejects an already cached external SDK client without resetting or sending it even with reported MaxRetry=<Reported>' -TestCases @(@{ Reported = 0 }, @{ Reported = 3 }) {
        param($Reported)
        $fixture = New-OfflineSdkClient @(204)
        # Model a stale context/client mismatch: changing a context property alone cannot rebuild middleware.
        $sdkSession.RequestContext.MaxRetry = $Reported
        { & (Get-Module PRA.Cloud) { Initialize-PcmGraphTransport -Context @{} } } | Should -Throw '*existing SDK session*'
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
        $fixture.Terminal.Count | Should -Be 0
        (Get-MgRequestContext).MaxRetry | Should -Be $Reported
    }
    It 'fails closed with no wire attempt when retry settings change after initialization' {
        & (Get-Module PRA.Cloud) { Initialize-PcmGraphTransport -Context @{} }
        $fixture = New-OfflineSdkClient @(503, 204)
        $sdkSession.RequestContext.MaxRetry = 3
        { Invoke-PcmGraph -Method DELETE -Uri 'v1.0/users/offline-id' -NoRetry } | Should -Throw '*retry settings changed*'
        $fixture.Terminal.Count | Should -Be 0
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
    }
    It 'keeps the ordinary <Method> retry budget at five wire attempts rather than multiplying SDK retries' -TestCases @(@{ Method = 'GET' }, @{ Method = 'PATCH' }, @{ Method = 'DELETE' }) {
        param($Method)
        & (Get-Module PRA.Cloud) { Initialize-PcmGraphTransport -Context @{} }
        $fixture = New-OfflineSdkClient @(503)
        Mock Start-Sleep -ModuleName PRA.Cloud { }
        { Invoke-PcmGraph -Method $Method -Uri 'v1.0/users/offline-id' } | Should -Throw
        $fixture.Terminal.Count | Should -Be 5
    }
}
