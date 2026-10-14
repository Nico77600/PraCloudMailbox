#Requires -Version 5.1

Describe 'Graph transport lifecycle (offline context model)' {
    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot '..\module\PRA2.Common.psm1') -Force
        Import-Module (Join-Path $PSScriptRoot '..\module\PRA2.Cloud.psm1') -Force
        . (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
        function global:Connect-MgGraph {
            [CmdletBinding()]
            param($ClientId, $TenantId, $CertificateThumbprint, $ContextScope, [switch]$NoWelcome)
        }
        function global:Disconnect-MgGraph { [CmdletBinding()]param() }
        function global:Connect-ExchangeOnline {
            [CmdletBinding()]
            param($AppId, $CertificateThumbprint, $Organization, $ShowBanner, $CommandName)
        }
    }
    BeforeEach {
        Reset-PraFakeGraphTransport
        & (Get-Module PRA2.Cloud) {
            $script:GraphRequestOwner = $null
            $script:GraphTransportReady = $false
        }
        $global:PraFakeGraphRequest.MaxRetry = 3
        $context = @{ Config = @{ Cloud = @{ TenantId = 'synthetic'; Organization = 'synthetic'; AppId = 'synthetic'; CertificateThumbprint = '12345678' }; Scope = @{} } }
        Mock Import-Module -ModuleName PRA2.Cloud { }
        Mock Get-ChildItem -ModuleName PRA2.Cloud { [pscustomobject]@{ HasPrivateKey = $true; NotAfter = (Get-Date).AddYears(1) } }
        Mock Connect-ExchangeOnline -ModuleName PRA2.Cloud { }
        Mock Connect-MgGraph -ModuleName PRA2.Cloud {
            $global:PraFakeGraphRequest.MaxRetry | Should -Be 0
            $ContextScope | Should -BeExactly Process
            $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = $ClientId }
        }
        Mock Disconnect-MgGraph -ModuleName PRA2.Cloud { $global:PraFakeGraphAuth = $null }
    }
    AfterAll {
        foreach ($name in @('Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext', 'Reset-PraFakeGraphTransport', 'Connect-MgGraph', 'Disconnect-MgGraph', 'Connect-ExchangeOnline')) {
            Remove-Item -Path "function:$name" -ErrorAction SilentlyContinue
        }
    }
    It 'initializes before authentication and restores only after disconnect, preserving other request settings' {
        $before = $global:PraFakeGraphRequest | ConvertTo-Json -Compress
        Connect-Pra2Cloud -Context $context | Should -Not -BeNullOrEmpty
        $global:PraFakeGraphClientGeneration | Should -Be 2
        Disconnect-Pra2Cloud -Context $context
        ($global:PraFakeGraphRequest | ConvertTo-Json -Compress) | Should -BeExactly $before
        $global:PraFakeGraphClientGeneration | Should -Be 3
        Should -Invoke Connect-MgGraph -ModuleName PRA2.Cloud -Times 1 -Exactly
    }
    It 'rejects an existing SDK context even with MaxRetry=<Retry>, without replacing its cached client' -TestCases @(@{ Retry = 0 }, @{ Retry = 3 }) {
        param($Retry)
        $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = 'external' }
        $global:PraFakeGraphRequest.MaxRetry = $Retry
        { Connect-Pra2Cloud -Context $context } | Should -Throw '*existing SDK session*'
        $global:PraFakeGraphClientGeneration | Should -Be 1
        $global:PraFakeGraphAuth.ClientId | Should -BeExactly external
        Should -Invoke Connect-MgGraph -ModuleName PRA2.Cloud -Times 0 -Exactly
        Should -Invoke Disconnect-MgGraph -ModuleName PRA2.Cloud -Times 0 -Exactly
    }
    It 'fails before authentication when <Fault> prevents initialization' -TestCases @(
        @{ Fault = 'MissingCommand' }, @{ Fault = 'InitialReadFails' }, @{ Fault = 'InitialContextMissing' }
        @{ Fault = 'SetFails' }, @{ Fault = 'ReadbackNonzero' }, @{ Fault = 'ReadbackMissing' }) {
        param($Fault)
        switch ($Fault) {
            MissingCommand { Mock Get-Command -ModuleName PRA2.Cloud { throw 'Missing SDK context command' } -ParameterFilter { $Name -eq 'Set-MgRequestContext' } }
            InitialReadFails { Mock Get-MgRequestContext -ModuleName PRA2.Cloud { throw 'Initial SDK context read failed' } }
            InitialContextMissing { $global:PraFakeGraphRequest = $null }
            SetFails { Mock Set-MgRequestContext -ModuleName PRA2.Cloud { throw 'SDK setting failed' } }
            ReadbackNonzero { Mock Set-MgRequestContext -ModuleName PRA2.Cloud { $global:PraFakeGraphRequest } }
            ReadbackMissing { Mock Set-MgRequestContext -ModuleName PRA2.Cloud { $global:PraFakeGraphRequest = $null } }
        }
        { Connect-Pra2Cloud -Context $context -WarningAction SilentlyContinue } | Should -Throw
        Should -Invoke Connect-MgGraph -ModuleName PRA2.Cloud -Times 0 -Exactly
    }
    It 'cleans up failed authentication without resetting an authenticated cached client' {
        Mock Connect-MgGraph -ModuleName PRA2.Cloud {
            $global:PraFakeGraphAuth = [pscustomobject]@{ ClientId = 'synthetic' }
            throw 'Synthetic authentication failure'
        }
        { Connect-Pra2Cloud -Context $context } | Should -Throw '*Synthetic authentication failure*'
        $global:PraFakeGraphAuth | Should -BeNullOrEmpty
        $global:PraFakeGraphRequest.MaxRetry | Should -Be 3
    }
    It 'does not restore request settings when SDK disconnect fails' {
        Connect-Pra2Cloud -Context $context | Out-Null
        Mock Disconnect-MgGraph -ModuleName PRA2.Cloud { throw 'Synthetic disconnect failure' }
        Disconnect-Pra2Cloud -Context $context -WarningAction SilentlyContinue
        $global:PraFakeGraphClientGeneration | Should -Be 2
        $global:PraFakeGraphRequest.MaxRetry | Should -Be 0
        & (Get-Module PRA2.Cloud) { $script:GraphTransportReady } | Should -BeFalse
    }
    It 'does not relinquish ownership when restoration cannot be verified' {
        Connect-Pra2Cloud -Context $context | Out-Null
        Mock Set-MgRequestContext -ModuleName PRA2.Cloud { $global:PraFakeGraphRequest }
        Disconnect-Pra2Cloud -Context $context -WarningAction SilentlyContinue
        & (Get-Module PRA2.Cloud) { $null -ne $script:GraphRequestOwner } | Should -BeTrue
        & (Get-Module PRA2.Cloud) { $script:GraphTransportReady } | Should -BeFalse
    }
}

# Optional installed-SDK qualification: the normal offline gate does not install Graph or sign in.
Describe 'Installed Graph SDK transport (in-memory HTTP, no authentication)' -Skip:($PSVersionTable.PSEdition -ne 'Core' -or -not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    BeforeAll {
        $sdkSession = $null
        $savedRetry = $null
        try {
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module (Join-Path $PSScriptRoot '..\module\PRA2.Cloud.psm1') -Force
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
        & (Get-Module PRA2.Cloud) {
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
        & (Get-Module PRA2.Cloud) { Initialize-Pra2GraphTransport -Context @{} }
        $fixture = New-OfflineSdkClient $(if ($Warm) { @(204, $Status, 204) } else { @($Status, 204) })
        if ($Warm) { $null = Invoke-Pra2Graph -Method GET -Uri 'v1.0/users/offline-id' }
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
        $failure = try { Invoke-Pra2Graph -Method DELETE -Uri 'v1.0/users/offline-id' -NoRetry } catch { $_ }
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
        { & (Get-Module PRA2.Cloud) { Initialize-Pra2GraphTransport -Context @{} } } | Should -Throw '*existing SDK session*'
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
        $fixture.Terminal.Count | Should -Be 0
        (Get-MgRequestContext).MaxRetry | Should -Be $Reported
    }
    It 'fails closed with no wire attempt when retry settings change after initialization' {
        & (Get-Module PRA2.Cloud) { Initialize-Pra2GraphTransport -Context @{} }
        $fixture = New-OfflineSdkClient @(503, 204)
        $sdkSession.RequestContext.MaxRetry = 3
        { Invoke-Pra2Graph -Method DELETE -Uri 'v1.0/users/offline-id' -NoRetry } | Should -Throw '*retry settings changed*'
        $fixture.Terminal.Count | Should -Be 0
        [object]::ReferenceEquals($sdkSession.GraphHttpClient, $fixture.Client) | Should -BeTrue
    }
    It 'keeps the ordinary <Method> retry budget at five wire attempts rather than multiplying SDK retries' -TestCases @(@{ Method = 'GET' }, @{ Method = 'PATCH' }, @{ Method = 'DELETE' }) {
        param($Method)
        & (Get-Module PRA2.Cloud) { Initialize-Pra2GraphTransport -Context @{} }
        $fixture = New-OfflineSdkClient @(503)
        Mock Start-Sleep -ModuleName PRA2.Cloud { }
        { Invoke-Pra2Graph -Method $Method -Uri 'v1.0/users/offline-id' } | Should -Throw
        $fixture.Terminal.Count | Should -Be 5
    }
}
