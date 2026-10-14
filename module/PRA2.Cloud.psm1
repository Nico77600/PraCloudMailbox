<#
.SYNOPSIS
    PRA Cloud Mailbox - cloud module: Microsoft Graph and Exchange Online (PowerShell 7).

.DESCRIPTION
    Certificate sign-in of the app registration (Cloud.AppId + Cloud.CertificateThumbprint) to Microsoft
    Graph (Microsoft.Graph.Authentication) and Exchange Online (ExchangeOnlineManagement 3.10+).

    It holds every cloud operation of Check, Convert and Recover: the state of the Entra ID object and of the
    Exchange Online recipient of each object of the Collect snapshot, the rules that tell whether Convert can
    run (Get-Pra2Readiness, pure functions of the snapshot row and of the cloud state, tested without any
    cloud access), source of authority, licences, shared mailbox permissions, the eDiscovery case hold
    (Security & Compliance) and the Entra Connect operations of Recover. The states of many objects are read in bulk
    (Get-Pra2CloudStateSet, Get-Pra2ExoStateSet): a few calls per 50 objects instead of five calls per object.

    Facts from the lab (2-7 October 2026) behind the rules:
      - Convert of a user = ExchangeGuid cleared, SOA transferred (isCloudManaged), then an Exchange plan:
        Exchange Online promotes the Teams storage (ComponentShared) to the primary mailbox.
      - An Exchange plan that is already enabled before the GUID is cleared does nothing: it must be
        removed and added again (test T2c).
      - Any hold on the cloud mailbox blocks the Recover of a user (tests T8, H).
      - usageLocation synced from AD is refused for licences right after the SOA transfer: it is set
        again from the cloud (test SH2).

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.0 (UNRELEASED)
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

#region Connections ---------------------------------------------------------------------------------

$script:GraphRequestOwner = $null
$script:GraphTransportReady = $false
$script:GraphPreviousMaxRetry = $null

function Initialize-Pra2GraphTransport {
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
    $script:GraphRequestOwner = $Context
    # Set before authentication: this cmdlet discards the SDK's cached HTTP client.
    $null = Set-MgRequestContext -MaxRetry 0 -ErrorAction Stop
    $request = Get-MgRequestContext -ErrorAction Stop
    if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry -or $request.MaxRetry -ne 0) {
        throw 'Graph SDK MaxRetry=0 was not verified; no cloud request is allowed.'
    }
    $script:GraphTransportReady = $true
}

function Assert-Pra2GraphTransport {
    [CmdletBinding()]
    param()
    if (-not $script:GraphTransportReady) { throw 'Graph transport is not initialized by Connect-Pra2Cloud; no cloud request is allowed.' }
    $request = Get-MgRequestContext -ErrorAction Stop
    if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry -or $request.MaxRetry -ne 0) {
        $script:GraphTransportReady = $false
        throw 'Graph SDK retry settings changed or cannot be verified; start a fresh PowerShell process. No cloud request is allowed.'
    }
}

function Connect-Pra2Cloud {
    <#
    .SYNOPSIS
        Certificate sign-in to Microsoft Graph and Exchange Online. Returns a short description.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $cloud = $Context.Config.Cloud
    foreach ($name in @('TenantId','Organization','AppId','CertificateThumbprint')) {
        if (-not $cloud[$name]) { throw "Configuration: Cloud.$name is empty (certificate sign-in of the app registration, see the developer guide, chapter 4.3)." }
    }
    $certificate = Get-ChildItem -Path "Cert:\CurrentUser\My\$($cloud.CertificateThumbprint)", "Cert:\LocalMachine\My\$($cloud.CertificateThumbprint)" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $certificate) { throw "Certificate $($cloud.CertificateThumbprint) not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My." }
    if (-not $certificate.HasPrivateKey) { throw "Certificate $($cloud.CertificateThumbprint) has no private key on this computer." }
    if ($certificate.NotAfter -lt (Get-Date).AddDays(30)) { $Context['CertificateWarning'] = "The certificate expires on $($certificate.NotAfter.ToString('yyyy-MM-dd'))." }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    try {
        Initialize-Pra2GraphTransport -Context $Context
        $Context['GraphConnected'] = $true # Own even a partially failed authentication, for cleanup.
        Connect-MgGraph -ClientId $cloud.AppId -TenantId $cloud.TenantId -CertificateThumbprint $cloud.CertificateThumbprint -ContextScope Process -NoWelcome -ErrorAction Stop
        Assert-Pra2GraphTransport
        Import-Module ExchangeOnlineManagement -MinimumVersion 3.10.0 -ErrorAction Stop
        $commands = @('Get-Recipient','Get-MailUser','Get-Mailbox','Get-MailboxLocation','Get-User','Get-OrganizationConfig','Set-MailUser','Set-Mailbox',
            'Add-MailboxPermission','Get-MailboxPermission','Add-RecipientPermission','Get-RecipientPermission')
        if ([bool](Get-PraValue $Context 'NativeResourceScope' ([string](Get-PraValue $Context.Config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource'))) {
            $commands += @('New-Mailbox','Set-CalendarProcessing','Get-CalendarProcessing','Set-MailboxCalendarConfiguration','Get-MailboxCalendarConfiguration',
                'New-MailContact','Set-MailContact','Get-MailContact','Get-ResourceConfig','Set-User',
                'Get-DistributionGroup','Get-DistributionGroupMember','Add-DistributionGroupMember','New-DistributionGroup','Set-DistributionGroup')
        }
        Connect-ExchangeOnline -AppId $cloud.AppId -CertificateThumbprint $cloud.CertificateThumbprint -Organization $cloud.Organization -ShowBanner:$false `
            -CommandName $commands -ErrorAction Stop
        $Context['ExoConnected'] = $true
        return ('Microsoft Graph and Exchange Online, app {0}, certificate {1} (valid until {2})' -f $cloud.AppId, $cloud.CertificateThumbprint.Substring(0, 8), $certificate.NotAfter.ToString('yyyy-MM-dd'))
    } catch {
        Disconnect-Pra2Cloud -Context $Context
        throw
    }
}

function Disconnect-Pra2Cloud {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($Context.ContainsKey('ExoConnected') -and $Context['ExoConnected']) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop } catch { $null = $_ }; $Context['ExoConnected'] = $false }
    if ([object]::ReferenceEquals($script:GraphRequestOwner, $Context)) {
        $script:GraphTransportReady = $false
        if ($Context.ContainsKey('GraphConnected') -and $Context['GraphConnected']) { try { $null = Disconnect-MgGraph -ErrorAction Stop } catch { $null = $_ }; $Context['GraphConnected'] = $false }
        try {
            # Never reset a cached authenticated client just to restore request settings.
            if (Get-MgContext -ErrorAction Stop) { throw 'Graph is still connected; SDK request settings were not restored.' }
            $null = Set-MgRequestContext -MaxRetry $script:GraphPreviousMaxRetry -ErrorAction Stop
            $request = Get-MgRequestContext -ErrorAction Stop
            if (-not $request -or -not $request.PSObject.Properties['MaxRetry'] -or $null -eq $request.MaxRetry -or $request.MaxRetry -ne $script:GraphPreviousMaxRetry) {
                throw 'Graph SDK request settings restoration could not be verified; use a fresh PowerShell process.'
            }
            $script:GraphRequestOwner = $null
            $script:GraphPreviousMaxRetry = $null
        } catch { Write-Warning $_.Exception.Message }
    }
}

function Test-Pra2NetworkFailure {
    <#
    .SYNOPSIS
        A failure of the connection itself, without an HTTP answer: timeout, reset, name resolution. Seen in the lab
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

function Invoke-Pra2Graph {
    <#
    .SYNOPSIS
        Microsoft Graph request with retries on throttling (429), transient errors (502, 503, 504) and connection
        failures without an answer (Test-Pra2NetworkFailure). Returns $null for 404 when -AllowNotFound.
        The SDK transient retry handler is disabled before sign-in; this wrapper owns the retry budget.
        -NoRetry never resends an ambiguous response, for unprotected deletion. SDK authentication challenges
        and redirects are not controlled by MaxRetry.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, [ValidateSet('GET','POST','PATCH','DELETE')][string]$Method = 'GET', [object]$Body, [switch]$AllowNotFound, [switch]$NoRetry,
        [System.Collections.IDictionary]$Headers)
    $parameters = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $parameters.Body = ($Body | ConvertTo-Json -Depth 10 -Compress); $parameters.ContentType = 'application/json' }
    if ($Headers -and $Headers.Count) { $parameters.Headers = $Headers }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        Assert-Pra2GraphTransport
        try { return Invoke-MgGraphRequest @parameters }
        catch {
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 }
            if ($code -eq 404 -and $AllowNotFound) { return $null }
            if (-not $NoRetry -and $code -in @(429, 502, 503, 504) -and $attempt -lt 5) { Start-Sleep -Seconds (5 * $attempt); continue }
            if (-not $NoRetry -and $code -eq 0 -and $attempt -lt 5 -and (Test-Pra2NetworkFailure $_.Exception)) { Start-Sleep -Seconds (10 * $attempt); continue }
            # The Graph message (error.code: error.message) says what is wrong; the exception alone only says "BadRequest".
            $detail = ''
            try { $graphError = ([string]$_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction Stop).error; $detail = '{0}: {1}' -f $graphError.code, $graphError.message } catch { $detail = '' }
            if ($detail) { throw [System.InvalidOperationException]::new(('{0} {1} failed ({2}): {3}' -f $Method, ($Uri -replace '\?.*$', ''), $code, $detail), $_.Exception) }
            throw
        }
    }
}

function Get-Pra2GraphRole {
    <# Application permissions (roles) of the current app-only token. #>
    [CmdletBinding()]
    param()
    $context = Get-MgContext
    if (-not $context) { throw 'Not signed in to Microsoft Graph.' }
    return @($context.Scopes | ForEach-Object { [string]$_ })
}
#endregion

#region Tenant facts ---------------------------------------------------------------------------------

function Get-Pra2TenantFact {
    <#
    .SYNOPSIS
        Licences (subscribedSkus), directory synchronisation status, organisation-wide Exchange holds and, in Group
        mode, the licence group.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [switch]$SkipUserGroup, [switch]$NativeOnly)
    if ($NativeOnly) { return [pscustomobject]@{ Skus = @(); Organization = $null; LicenceGroup = $null; LicenceGroupCloudManaged = $null; OrgHolds = @() } }
    $skus = @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value)
    $organization = @((Invoke-Pra2Graph -Uri 'v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime').value) | Select-Object -First 1
    $group = $null; $groupBehavior = $null
    $users = $Context.Config.Licensing.Users
    if (-not $SkipUserGroup -and $users.Mode -eq 'Group' -and $users.GroupId) {
        $group = Invoke-Pra2Graph -Uri ("v1.0/groups/{0}?`$select=id,displayName,assignedLicenses,onPremisesSyncEnabled,securityEnabled,mailEnabled" -f $users.GroupId) -AllowNotFound
        if ($group) { $groupBehavior = Invoke-Pra2Graph -Uri ("v1.0/groups/{0}/onPremisesSyncBehavior?`$select=isCloudManaged" -f $users.GroupId) -AllowNotFound }
    }
    # Retention policies applied to every mailbox (InPlaceHolds of the organisation, prefix mbx).
    $orgHolds = @()
    try { $orgHolds = @((Get-OrganizationConfig -ErrorAction Stop).InPlaceHolds | ForEach-Object { [string]$_ } | Where-Object { $_ -match '^mbx' }) } catch { $orgHolds = @() }
    return [pscustomobject]@{ Skus = $skus; Organization = $organization; LicenceGroup = $group; LicenceGroupCloudManaged = $(if ($groupBehavior) { [bool]$groupBehavior.isCloudManaged } else { $null })
        OrgHolds = $orgHolds }
}

function Get-Pra2MailboxPlanId {
    <# servicePlanId of every Exchange plan that gives a mailbox, in the given SKUs. #>
    param([AllowEmptyCollection()][object[]]$Skus = @())
    return @($Skus | ForEach-Object { $_.servicePlans } | Where-Object { $_ -and [string]$_.servicePlanName -match $script:ExchangePlanPattern } | ForEach-Object { [string]$_.servicePlanId } | Select-Object -Unique)
}

function Get-Pra2SkuExchangePlan {
    <# Exchange service plans of a SKU (servicePlanName EXCHANGE_S_*), minus the disabled ones. #>
    param([Parameter(Mandatory)][object]$Sku, [string[]]$DisabledPlanIds = @())
    return @($Sku.servicePlans | Where-Object { [string]$_.servicePlanName -match $script:ExchangePlanPattern -and [string]$_.servicePlanId -notin $DisabledPlanIds } |
        ForEach-Object { [string]$_.servicePlanName })
}

function Get-Pra2FreeUnit {
    param([Parameter(Mandatory)][object]$Sku)
    return [int]$Sku.prepaidUnits.enabled - [int]$Sku.consumedUnits
}

function Test-Pra2TenantReadiness {
    <#
    .SYNOPSIS
        Tenant-level rules: app permissions, licences for the users (Group, Kiosk or Direct) and for the
        temporary licence of the shared mailboxes. Pure function (no cloud call).
    .PARAMETER Roles
        Application permissions of the token.
    .PARAMETER Tenant
        Output of Get-Pra2TenantFact.
    .PARAMETER UserCount
        Users of the snapshot that will need an Exchange plan.
    .PARAMETER SharedCount
        Shared, room and equipment mailboxes of the snapshot.
    .OUTPUTS
        Findings: Level (Ok, Info, Warn, Error), Code, Message.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [AllowEmptyCollection()][string[]]$Roles = @(), [Parameter(Mandatory)][object]$Tenant,
        [int]$UserCount = 0, [int]$SharedCount = 0, [int]$ResourceCount = 0, [switch]$SkipUserLicensing, [switch]$NativeOnly)
    $findings = [System.Collections.Generic.List[object]]::new()
    $add = { param($Level, $Code, $Message) [void]$findings.Add([pscustomobject]@{ Level = $Level; Code = $Code; Message = $Message }) }
    $licensing = $Config.Licensing
    $needed = @($script:RequiredGraphRoles)
    if ($NativeOnly) { $needed = @($script:RequiredGraphRoles | Where-Object Role -eq 'User.ReadWrite.All') }
    if ((-not $SkipUserLicensing -and $licensing.Users.Mode -eq 'Group') -or $ResourceCount -gt 0) { $needed += $script:GroupModeRoles[0] }
    $missing = @($needed | Where-Object { $_.Role -notin $Roles -and -not ($_.Role -eq 'LicenseAssignment.ReadWrite.All' -and 'Directory.ReadWrite.All' -in $Roles) })
    if ($missing.Count) { & $add 'Error' 'APP_PERMISSION' ('Missing Graph application permission(s): ' + (($missing | ForEach-Object { '{0} ({1})' -f $_.Role, $_.Why }) -join '; ')) }
    else { & $add 'Ok' 'APP_PERMISSION' ('Graph application permissions present: ' + (($needed | ForEach-Object { $_.Role }) -join ', ')) }

    $skuById = @{}; $skuByName = @{}
    foreach ($sku in @($Tenant.Skus)) { $skuById[[string]$sku.skuId] = $sku; $skuByName[[string]$sku.skuPartNumber] = $sku }
    if (-not $SkipUserLicensing) { switch ($licensing.Users.Mode) {
        'Group' {
            if (-not $licensing.Users.GroupId) { & $add 'Error' 'LICENCE_GROUP' 'Licensing.Users.GroupId is empty (object ID of the group that carries the Exchange plan).' }
            elseif (-not $Tenant.LicenceGroup) { & $add 'Error' 'LICENCE_GROUP' "Licence group $($licensing.Users.GroupId) not found in Entra ID." }
            else {
                $group = $Tenant.LicenceGroup
                $plans = @(); $free = 0; $names = @()
                foreach ($assigned in @($group.assignedLicenses)) {
                    $sku = $skuById[[string]$assigned.skuId]
                    if (-not $sku) { continue }
                    $exchange = @(Get-Pra2SkuExchangePlan -Sku $sku -DisabledPlanIds @($assigned.disabledPlans | ForEach-Object { [string]$_ }))
                    if ($exchange.Count) { $plans += $exchange; $free += (Get-Pra2FreeUnit $sku); $names += [string]$sku.skuPartNumber }
                }
                if (-not $plans.Count) { & $add 'Error' 'LICENCE_GROUP' "Group $($group.displayName) assigns no licence with an Exchange plan enabled." }
                else {
                    $level = if ($free -ge $UserCount) { 'Ok' } else { 'Error' }
                    & $add $level 'LICENCE_CAPACITY' ('Group {0}: {1} ({2}), {3} free unit(s) for {4} user(s).' -f $group.displayName, ($names -join ', '), (($plans | Select-Object -Unique) -join ', '), $free, $UserCount)
                }
                if ($group.onPremisesSyncEnabled -and -not $Tenant.LicenceGroupCloudManaged) {
                    & $add 'Info' 'LICENCE_GROUP_SOA' ("Group $($group.displayName) is synchronised from AD: Convert transfers its source of authority to the cloud first (Group-OnPremisesSyncBehavior.ReadWrite.All), Recover gives it back to AD last (test T15).")
                    if ('Group-OnPremisesSyncBehavior.ReadWrite.All' -notin $Roles) { & $add 'Error' 'APP_PERMISSION' 'Missing Graph application permission Group-OnPremisesSyncBehavior.ReadWrite.All (licence group synchronised from AD).' }
                }
            }
        }
        'Direct' {
            $sku = $skuByName[[string]$licensing.Users.SkuPartNumber]
            if (-not $sku) { & $add 'Error' 'LICENCE_SKU' "Licensing.Users.SkuPartNumber $($licensing.Users.SkuPartNumber) is not a licence of the tenant." }
            else {
                $free = Get-Pra2FreeUnit $sku
                & $add $(if ($free -ge $UserCount) { 'Ok' } else { 'Error' }) 'LICENCE_CAPACITY' ('{0}: {1} free unit(s) for {2} user(s).' -f $sku.skuPartNumber, $free, $UserCount)
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
            elseif (-not @(Get-Pra2SkuExchangePlan -Sku $sku).Count) { & $add 'Error' 'SHARED_LICENCE' "$name has no Exchange plan." }
            else {
                $free = Get-Pra2FreeUnit $sku
                $level = if ($free -ge 1) { 'Ok' } else { 'Error' }
                $parallel = [Math]::Max(1, [int](Get-PraValue $licensing.Shared 'Parallel' 100))
                $wave = [Math]::Max(1, [Math]::Min($parallel, $free))
                & $add $level 'SHARED_LICENCE' ('{0}: {1} free unit(s); the shared mailboxes are converted in waves of {2} (Licensing.Shared.Parallel = {3}, one temporary unit each, given back after each wave; {4} in scope).' -f $name, $free, $wave, $parallel, $SharedCount)
            }
        }
    }
    $organization = $Tenant.Organization
    if ($organization -and $organization.onPremisesSyncEnabled -and $organization.onPremisesLastSyncDateTime) {
        $last = [datetime]$organization.onPremisesLastSyncDateTime
        & $add 'Info' 'DIRECTORY_SYNC' ('Last Entra Connect synchronisation: {0:yyyy-MM-dd HH:mm} UTC.' -f $last.ToUniversalTime())
    }
    $orgHolds = @(if ($Tenant.PSObject.Properties['OrgHolds']) { $Tenant.OrgHolds | Where-Object { $_ } })
    if ($orgHolds.Count) {
        & $add 'Warn' 'ORG_HOLD' ('Retention policies applied to every mailbox ({0}): they will also hold the cloud mailboxes of the users, and a hold blocks their Recover (tests T8, H). Exclude the users in scope from these policies, or plan to lift them before Recover.' -f ($orgHolds -join ', '))
    }
    return $findings.ToArray()
}
#endregion

#region Objects --------------------------------------------------------------------------------------

function Get-Pra2CloudState {
    <#
    .SYNOPSIS
        Entra ID object and Exchange Online recipient of one snapshot row (read-only). Get-Pra2CloudStateSet for one row.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Record)
    $state = (Get-Pra2CloudStateSet -Context $Context -Records @($Record))[[string]$Record.object_guid]
    if ($state -is [System.Management.Automation.ErrorRecord]) { throw $state }
    return $state
}
#endregion

#region Many objects at once ------------------------------------------------------------------------
# A state read one object at a time costs five calls (about 3 s): 15 hours for 21,000 objects. These functions read
# the same states in bulk (measured in a lab tenant of 21,700 users and 16,000 mailboxes): every user page by page
# (20 s), the source of authority by Graph batches of 20 (0.4 s each), Exchange Online by filters of 50 object IDs
# (0.3 to 0.7 s each).

$script:ExoFilterSize = 50
$script:GraphBatchSize = 20

function ConvertFrom-Pra2MailboxLocation {
    <# MailboxLocations of Get-MailUser / Get-EXOMailbox ('1;<guid>;<type>;<database>;<id>') as Type and Guid, like Get-MailboxLocation. #>
    param([AllowNull()][AllowEmptyCollection()][object[]]$Value = @())
    return @(foreach ($text in @($Value | Where-Object { $_ })) {
            $parts = ([string]$text).Split(';')
            if ($parts.Count -ge 3 -and $parts[2]) { [pscustomobject]@{ Type = $parts[2]; Guid = $parts[1] } }
        })
}

function Get-Pra2ExoFilter {
    <# An OPATH filter that matches any of the object IDs. #>
    param([Parameter(Mandatory)][string[]]$Id)
    return (@($Id | ForEach-Object { "ExternalDirectoryObjectId -eq '{0}'" -f $_.Replace("'", "''") }) -join ' -or ')
}

function Get-Pra2ExoStateSet {
    <#
    .SYNOPSIS
        Get-Pra2ExoState of many Entra object IDs, with their mailbox locations (Teams storage): filters of 50 IDs on
        Get-EXORecipient, Get-User, Get-MailUser and Get-EXOMailbox. A chunk that fails is read again one object at a
        time. Returns a hashtable ID -> state (Type '' when there is no recipient).
    #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Id = @(), [string]$TagAttribute = 'CustomAttribute1')
    $result = @{}
    $ids = @($Id | Where-Object { $_ } | Select-Object -Unique)
    for ($start = 0; $start -lt $ids.Count; $start += $script:ExoFilterSize) {
        $chunk = [string[]]@($ids[$start..([Math]::Min($ids.Count, $start + $script:ExoFilterSize) - 1)])
        try {
            $filter = Get-Pra2ExoFilter -Id $chunk
            $recipients = @{}; $users = @{}; $mailUsers = @{}; $mailboxes = @{}
            foreach ($r in @(Get-EXORecipient -Filter $filter -ResultSize Unlimited -Properties ExternalDirectoryObjectId, ExchangeGuid, $TagAttribute -ErrorAction Stop)) { $recipients[[string]$r.ExternalDirectoryObjectId] = $r }
            foreach ($u in @(Get-User -Filter $filter -ResultSize Unlimited -ErrorAction Stop)) { $users[[string]$u.ExternalDirectoryObjectId] = $u }
            $mailUserIds = [string[]]@($recipients.Keys | Where-Object { [string]$recipients[$_].RecipientTypeDetails -eq 'MailUser' })
            $mailboxIds = [string[]]@($recipients.Keys | Where-Object { [string]$recipients[$_].RecipientTypeDetails -match 'Mailbox$' })
            if ($mailUserIds.Count) { foreach ($m in @(Get-MailUser -Filter (Get-Pra2ExoFilter -Id $mailUserIds) -ResultSize Unlimited -ErrorAction Stop)) { $mailUsers[[string]$m.ExternalDirectoryObjectId] = $m } }
            if ($mailboxIds.Count) {
                foreach ($m in @(Get-EXOMailbox -Filter (Get-Pra2ExoFilter -Id $mailboxIds) -ResultSize Unlimited -Properties ExternalDirectoryObjectId, InPlaceHolds, LitigationHoldEnabled, MailboxLocations -ErrorAction Stop)) { $mailboxes[[string]$m.ExternalDirectoryObjectId] = $m }
            }
            foreach ($one in $chunk) {
                $recipient = $recipients[$one]; $user = $users[$one]; $holds = @(); $locations = @(); $guid = ''
                if ($recipient) {
                    $guid = [string]$recipient.ExchangeGuid
                    if ($mailUsers.ContainsKey($one)) {
                        $holds = @($mailUsers[$one].InPlaceHolds | ForEach-Object { [string]$_ }); $guid = [string]$mailUsers[$one].ExchangeGuid
                        $locations = @(ConvertFrom-Pra2MailboxLocation @($mailUsers[$one].MailboxLocations))
                    } elseif ($mailboxes.ContainsKey($one)) {
                        $mailbox = $mailboxes[$one]
                        $holds = @($mailbox.InPlaceHolds | ForEach-Object { [string]$_ }); if ($mailbox.LitigationHoldEnabled) { $holds += 'LitigationHold' }
                        $locations = @(ConvertFrom-Pra2MailboxLocation @($mailbox.MailboxLocations))
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
                $state = Get-Pra2ExoState -Identity $one -TagAttribute $TagAttribute
                $locations = @(Get-MailboxLocation -User $one -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ForEach-Object { [pscustomobject]@{ Type = [string]$_.MailboxLocationType; Guid = [string]$_.MailboxGuid } })
                $state | Add-Member -NotePropertyName Locations -NotePropertyValue $locations
                $result[$one] = $state
            }
        }
    }
    return $result
}

function Invoke-Pra2GraphBatch {
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
    $toResult = {
        param([int]$Status, [object]$Body)
        $text = ''
        if ($Status -ge 300 -and $Body) { $graphError = Get-PraValue $Body 'error' $null; if ($graphError) { $text = '{0}: {1}' -f (Get-PraValue $graphError 'code' ''), (Get-PraValue $graphError 'message' '') } }
        [pscustomobject]@{ Ok = ($Status -ge 200 -and $Status -lt 300); Status = $Status; Body = $Body; Error = $text }
    }
    for ($round = 1; $round -le 3 -and $pending.Count; $round++) {
        $again = [System.Collections.Generic.List[int]]::new()
        for ($start = 0; $start -lt $pending.Count; $start += $script:GraphBatchSize) {
            $chunk = @($pending.GetRange($start, [Math]::Min($script:GraphBatchSize, $pending.Count - $start)))
            $batch = @(foreach ($index in $chunk) {
                    $request = $Requests[$index]
                    $item = @{ id = [string]$index; method = [string]$request.Method; url = [string]$request.Url }
                    if ($request.ContainsKey('Body') -and $null -ne $request.Body) { $item.body = $request.Body; $item.headers = @{ 'Content-Type' = 'application/json' } }
                    $item
                })
            $answer = Invoke-Pra2Graph -Method POST -Uri 'v1.0/$batch' -Body @{ requests = $batch }
            $seen = @{}
            foreach ($response in @($answer.responses)) {
                $index = [int]$response.id; $seen[$index] = $true
                $status = [int]$response.status
                if ($status -eq 429 -or $status -ge 500) { $again.Add($index) } else { $results[$index] = & $toResult $status (Get-PraValue $response 'body' $null) }
            }
            foreach ($index in $chunk) { if (-not $seen.ContainsKey($index)) { $again.Add($index) } }
        }
        $pending = $again
        if ($pending.Count -and $round -lt 3) { Start-Sleep -Seconds (5 * $round) }
    }
    foreach ($index in $pending) {
        $request = $Requests[$index]
        $parameters = @{ Method = [string]$request.Method; Uri = ('v1.0' + [string]$request.Url) }
        if ($request.ContainsKey('Body') -and $null -ne $request.Body) { $parameters.Body = $request.Body }
        try { $results[$index] = & $toResult 200 (Invoke-Pra2Graph @parameters) }
        catch {
            $code = 0; try { $code = [int]$_.Exception.InnerException.Response.StatusCode } catch { $code = 0 }
            if (-not $code) { try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 } }
            $results[$index] = [pscustomobject]@{ Ok = $false; Status = $code; Body = $null; Error = ($_.Exception.Message -replace '\s+', ' ') }
        }
    }
    return , $results
}

function Get-Pra2SyncBehaviorSet {
    <#
    .SYNOPSIS
        isCloudManaged of many users (source of authority), by Graph batches. Returns a hashtable user ID -> $true,
        $false or $null (none). A request that keeps failing stops the read.
    #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$UserId = @())
    $result = @{}
    $ids = @($UserId | Where-Object { $_ } | Select-Object -Unique)
    $answers = Invoke-Pra2GraphBatch -Requests @($ids | ForEach-Object { @{ Method = 'GET'; Url = ('/users/{0}/onPremisesSyncBehavior?$select=isCloudManaged' -f $_) } })
    for ($i = 0; $i -lt $ids.Count; $i++) {
        $answer = $answers[$i]
        if ($answer.Ok) { $result[$ids[$i]] = [bool](Get-PraValue $answer.Body 'isCloudManaged' $false) }
        elseif ($answer.Status -eq 404) { $result[$ids[$i]] = $null }
        else { throw ('Source of authority of {0} not readable ({1}): {2}' -f $ids[$i], $answer.Status, $answer.Error) }
    }
    return $result
}

function Get-Pra2GraphUserSet {
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
                $found = @((Invoke-Pra2Graph -Uri ("v1.0/users?`$filter={0}&`$select={1}" -f $filter, $select)).value)
                if ($found.Count -eq 1) { $user = $found[0]; $matchedBy = 'immutableId' }
            }
            if (-not $user -and $record.user_principal_name) {
                $user = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select={1}" -f [uri]::EscapeDataString([string]$record.user_principal_name), $select) -AllowNotFound
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
        $page = Invoke-Pra2Graph -Uri $uri
        foreach ($user in @($page.value)) {
            if ($user.onPremisesImmutableId) {
                $key = [string]$user.onPremisesImmutableId
                if ($byImmutable.ContainsKey($key)) { $byImmutable[$key] = $null } else { $byImmutable[$key] = $user }
            }
            if ($user.userPrincipalName) { $byUpn[[string]$user.userPrincipalName] = $user }
        }
        $uri = [string](Get-PraValue $page '@odata.nextLink' '')
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

function Get-Pra2CloudStateSet {
    <#
    .SYNOPSIS
        Get-Pra2CloudState of many snapshot rows: Entra ID users, source of authority and Exchange Online in bulk.
        Returns a hashtable object_guid -> state, or the ErrorRecord of a row that could not be read.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [AllowEmptyCollection()][object[]]$Records = @(), [int]$PageThreshold = 100)
    $result = @{}
    if (-not $Records.Count) { return $result }
    $native = @($Records | Where-Object {
            (Test-Pra2NativeResource -Config $Context.Config -Record $_) -or
            ($Context.ContainsKey('ResourceJournal') -and (Get-Pra2ResourceIntent -Item $Context.ResourceJournal[[string]$_.object_guid]))
        })
    if ($native.Count) {
        try {
            $inventory = Get-Pra2ResourceInventory -Records $native
            foreach ($record in $native) {
                $intent = if ($Context.ContainsKey('ResourceJournal')) { Get-Pra2ResourceIntent -Item $Context.ResourceJournal[[string]$record.object_guid] } else { $null }
                try { $result[[string]$record.object_guid] = Get-Pra2ResourceState -Record $record -Records $Records -Inventory $inventory -Intent $intent }
                catch { $result[[string]$record.object_guid] = $_ }
            }
        } catch { foreach ($record in $native) { $result[[string]$record.object_guid] = $_ } }
    }
    $legacy = @($Records | Where-Object { $_ -notin $native })
    if (-not $legacy.Count) { return $result }
    $users = Get-Pra2GraphUserSet -Records $legacy -PageThreshold $PageThreshold
    $ids = [string[]]@($users.Values | Where-Object { $_.User } | ForEach-Object { [string]$_.User.id })
    $behaviors = Get-Pra2SyncBehaviorSet -UserId $ids
    $exo = Get-Pra2ExoStateSet -Id $ids -TagAttribute ([string]$Context.Config.Retention.TagAttribute)
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

function Test-Pra2NativeResource {
    param([hashtable]$Config, [object]$Record)
    return ([string]$Record.kind -in @('Room', 'Equipment') -and [string](Get-PraValue $Config.Scope 'ResourceMode' 'SynchronizedShared') -eq 'CloudResource')
}

function Get-Pra2ResourceIntent {
    param([AllowNull()][object]$Item)
    if (-not $Item -or -not (Get-PraValue $Item 'original_json' '')) { return $null }
    $value = [string]$Item.original_json | ConvertFrom-Json -ErrorAction Stop
    if ([string](Get-PraValue $value 'lifecycle' '') -eq 'CloudResource') {
        $token = [string](Get-PraValue $value 'token' '')
        if ($token -notmatch '^[a-f0-9]{32}$' -or [string](Get-PraValue $value 'creationName' '') -cnotmatch ("^PRA-R-[a-f0-9]{{8}}-[a-f0-9]{{8}}-{0}$" -f $token) -or
            [string](Get-PraValue $value 'markerAttribute' '') -notin @('CustomAttribute14','CustomAttribute15')) { throw 'Native resource journal ownership token/name/marker is invalid; no identity is adopted.' }
        $null = Get-Pra2ResourceProtectionMode -Intent $value
        return $value
    }
    return $null
}

function Get-Pra2ResourceProtectionMode {
    param([AllowNull()][object]$Intent)
    if (-not $Intent) { return 'None' }
    # Missing mode means the original licensed/case-hold contract, never an implicit downgrade.
    $mode = Get-PraValue $Intent 'protectionMode' 'CaseHold'
    if ($mode -isnot [string] -or $mode -cnotin @('None','CaseHold')) { throw "Unknown native resource protectionMode '$mode'; no mutation." }
    if ($mode -eq 'None' -and ((Get-PraValue $Intent 'licensingGroupId' '') -or (Get-PraValue $Intent 'holdTag' ''))) {
        throw 'Unprotected journal conflicts with recorded licence group/hold; no protection is removed.'
    }
    $null = Get-Pra2UnprotectedDeletionProgress -Intent $Intent
    if ($mode -eq 'CaseHold' -and ((Get-PraValue $Intent 'unprotectedDeletionAttempted' $false) -or (Get-PraValue $Intent 'unprotectedDeletionProgress' ''))) {
        throw 'Protected journal conflicts with an unprotected deletion record; no deletion.'
    }
    return $mode
}

function Assert-Pra2UnprotectedResource {
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
    if ($CheckOrganization) { Assert-Pra2ResourceOrganizationUnheld }
}

function Assert-Pra2ResourceOrganizationUnheld {
    $org = Get-OrganizationConfig -ErrorAction Stop
    if (-not $org -or -not $org.PSObject.Properties['InPlaceHolds'] -or $null -eq $org.InPlaceHolds) { throw 'Organization InPlaceHolds not exposed; no unprotected resource mutation.' }
    if (@($org.InPlaceHolds | Where-Object { $_ -and [string]$_ -notmatch '^-' }).Count) { throw 'Unexpected organization-wide holds; unprotected resource mutation refused. No hold is removed.' }
}

function Get-Pra2UnprotectedDeletionProgress {
    <# Issuing is persisted before sending DELETE; an interrupted/ambiguous response never authorises a second deletion of a restored identity. #>
    param([Parameter(Mandatory)][object]$Intent)
    $progress = Get-PraValue $Intent 'unprotectedDeletionProgress' ''
    $attempted = Get-PraValue $Intent 'unprotectedDeletionAttempted' $false
    if ($attempted -isnot [bool]) { throw 'Unprotected deletion intent must be a Boolean; no deletion.' }
    if ($progress -isnot [string] -or $progress -cnotin @('', 'Intent', 'Issuing', 'Issued', 'Verified')) { throw "Unknown unprotected deletion progress '$progress'; no deletion." }
    if (-not $progress) { return $(if ($attempted) { 'Issuing' } else { '' }) }
    if (-not $attempted) { throw 'Unprotected deletion progress is journaled without a deletion intent; no deletion.' }
    return $progress
}

function Assert-Pra2UnprotectedDeletionRepeatable {
    <# Only a pre-request intent or a definitively rejected request can authorise another DELETE. #>
    param([Parameter(Mandatory)][object]$Intent)
    $progress = Get-Pra2UnprotectedDeletionProgress -Intent $Intent
    if ($progress -eq 'Verified') { throw "Deleted resource identity $($Intent.cloudId) is active again after a verified deletion; no second deletion and no routing contact. Investigate the restore." }
    if ($progress -eq 'Issuing') { throw "Resource identity $($Intent.cloudId) is visible after an interrupted or ambiguous deletion request; no second deletion. Wait for deletion evidence, or investigate a restore." }
    if ($progress -eq 'Issued') { throw "Resource identity $($Intent.cloudId) is still or again visible after its deletion returned; no second deletion until the directory shows it deleted. Wait, or investigate a restore." }
}

function Test-Pra2UnprotectedResourceDeleted {
    param([Parameter(Mandatory)][object]$Intent)
    if ((Get-Pra2ResourceProtectionMode $Intent) -ne 'None') { throw 'Deletion proof applies only to the unprotected lifecycle; protected resources keep their inactive-copy verification.' }
    $progress = Get-Pra2UnprotectedDeletionProgress -Intent $Intent
    if (-not $progress) { throw 'No journaled unprotected deletion intent; missing active identity is not proof of safe recovery.' }
    $cloudId = [string](Get-PraValue $Intent 'cloudId' '')
    $cloudGuid = [string](Get-PraValue $Intent 'cloudGuid' '')
    if (-not $cloudId -or -not $cloudGuid) { throw 'Journaled resource immutable ID/ExchangeGuid is missing; deletion cannot be verified.' }
    if (Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id" -f $cloudId) -AllowNotFound) {
        if ($progress -eq 'Verified') { Assert-Pra2UnprotectedDeletionRepeatable -Intent $Intent }
        return $false
    }
    $deleted = Invoke-Pra2Graph -Uri ("v1.0/directory/deletedItems/{0}?`$select=id" -f $cloudId) -AllowNotFound
    if (-not $deleted) { return $false }
    if ([string](Get-PraValue $deleted 'id' '') -ne $cloudId) { throw 'Deleted resource identity differs from journal; no routing contact.' }
    $soft = @(Get-Mailbox -SoftDeletedMailbox -ResultSize Unlimited -ErrorAction Stop |
        Where-Object { [string]$_.ExchangeGuid -eq $cloudGuid })
    if (-not $soft.Count) { return $false }
    if ($soft.Count -ne 1 -or [string](Get-PraValue $soft[0] 'ExternalDirectoryObjectId' '') -ne $cloudId -or
        -not $soft[0].PSObject.Properties['IsInactiveMailbox'] -or $soft[0].IsInactiveMailbox -ne $false) {
        throw 'Exact resource soft-deleted identity/GUID or non-inactive state cannot be verified; no contact or purge.'
    }
    $deletedUser = Invoke-Pra2Graph -Uri ("v1.0/directory/deletedItems/{0}?`$select=id,assignedLicenses,licenseAssignmentStates" -f $cloudId)
    if ([string](Get-PraValue $deletedUser 'id' '') -ne $cloudId) { throw 'Deleted resource licence identity differs from journal; no routing contact.' }
    Assert-Pra2UnprotectedResource -Owned ([pscustomobject]@{ User = $deletedUser; Mailbox = $soft[0] }) -CheckOrganization
    return $true
}

function Get-Pra2UnprotectedRecoverCollisions {
    param([Parameter(Mandatory)][object]$Intent, [Parameter(Mandatory)][object]$Record,
        [Parameter(Mandatory)][object]$Inventory, [string]$OwnedContactId = '')
    $cloudId = [string](Get-PraValue $Intent 'cloudId' '')
    $cloudGuid = [string](Get-PraValue $Intent 'cloudGuid' '')
    $parsedId = [guid]::Empty
    $parsedGuid = [guid]::Empty
    if ([string](Get-PraValue $Intent 'lifecycle' '') -cne 'CloudResource' -or
        (Get-Pra2ResourceProtectionMode $Intent) -cne 'None' -or
        (Get-Pra2UnprotectedDeletionProgress $Intent) -cne 'Verified' -or
        (Get-PraValue $Intent 'creationAttempted' $false) -isnot [bool] -or $Intent.creationAttempted -ne $true -or
        [string]$Record.kind -cnotin @('Room','Equipment') -or
        [string](Get-PraValue $Intent 'cloudType' '') -cne "$($Record.kind)Mailbox" -or
        -not [guid]::TryParse($cloudId, [ref]$parsedId) -or $parsedId -eq [guid]::Empty -or
        -not [guid]::TryParse($cloudGuid, [ref]$parsedGuid) -or $parsedGuid -eq [guid]::Empty -or
        $cloudGuid -eq [string]$Record.exchange_guid) {
        throw 'Expected resource tombstone requires verified unprotected journal ownership, independent immutable IDs and source type; no contact.'
    }
    # Fresh service proof, not a caller-supplied flag. Retained proxies alone do not establish an Exchange write conflict.
    if (-not (Test-Pra2UnprotectedResourceDeleted -Intent $Intent)) { throw 'Exact resource deletion proof is no longer available; no routing contact or purge.' }
    $expected = @($Inventory.Recipients | Where-Object {
        $when = Get-PraValue $_ 'WhenSoftDeleted' $null
        $validWhen = $when -is [datetime] -or $when -is [datetimeoffset]
        if ($when -is [string] -and -not [string]::IsNullOrWhiteSpace($when)) {
            $date = [datetimeoffset]::MinValue
            $validWhen = [datetimeoffset]::TryParse($when, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$date)
        }
        [string](Get-PraValue $_ 'ExternalDirectoryObjectId' '') -eq $cloudId -and
        [string](Get-PraValue $_ 'ExchangeGuid' '') -eq $cloudGuid -and
        [string](Get-PraValue $_ 'RecipientTypeDetails' '') -ceq [string]$Intent.cloudType -and $validWhen
    })
    if ($expected.Count -gt 1) { throw 'Expected resource tombstone inventory is ambiguous; no routing contact.' }
    $filtered = [pscustomobject]@{
        Users = @($Inventory.Users | Where-Object { [string]$_.id -ne $cloudId })
        Recipients = @($Inventory.Recipients | Where-Object { $_ -notin $expected })
    }
    foreach ($property in @('RecipientKeys','GraphImmutableIds')) {
        if ($Inventory.PSObject.Properties[$property]) { $filtered | Add-Member -NotePropertyName $property -NotePropertyValue $Inventory.$property }
    }
    Get-Pra2ResourceCollisions -Record $Record -Inventory $filtered -OwnedContactId $OwnedContactId
}

function Get-Pra2ResourceAddresses {
    param([Parameter(Mandatory)][object]$Record)
    $primary = [string]$Record.primary_smtp_address
    if ($primary -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw 'Resource primary SMTP address is invalid.' }
    $values = @("SMTP:$primary")
    $json = [string](Get-PraValue $Record 'email_addresses_json' '')
    if ($json) { $values += @(($json | ConvertFrom-Json -ErrorAction Stop) | ForEach-Object { [string]$_ }) }
    $legacy = [string](Get-PraValue $Record 'legacy_exchange_dn' '')
    if ($legacy) { $values += "X500:$legacy" }
    # Keep the original primary, not an accidental primary flag on another collected proxy.
    return @($values | Where-Object { $_ -and $_ -ine "smtp:$primary" } | ForEach-Object { if ($_ -cmatch '^SMTP:') { 'smtp:' + $_.Substring(5) } else { $_ } } |
        Select-Object -Unique) + @("SMTP:$primary")
}

function Get-Pra2ResourceKeys {
    param([object]$Record)
    $keys = @(Get-Pra2ResourceAddresses $Record | ForEach-Object { $_.ToLowerInvariant() })
    foreach ($field in @('user_principal_name', 'primary_smtp_address')) {
        $value = [string](Get-PraValue $Record $field '')
        if ($value) { $keys += ('smtp:' + $value.ToLowerInvariant()) }
    }
    $alias = [string](Get-PraValue $Record 'alias' '')
    if (-not $alias) { $alias = ([string]$Record.primary_smtp_address -split '@')[0] }
    $keys += ('alias:' + $alias.ToLowerInvariant())
    return @($keys | Select-Object -Unique)
}

function Test-Pra2RecipientNotFound {
    param([Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    # EXO can serialize the server exception as a generic exception; inspect metadata, never its localized message.
    if ($ErrorRecord.CategoryInfo.Category -notin @('ObjectNotFound', 'NotSpecified')) { return $false }
    $exception = $ErrorRecord.Exception
    while ($exception) {
        if ($exception.GetType().FullName -ceq 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException') { return $true }
        $remoteTypeName = Get-PraValue -Object $exception.Data['RemoteException'] -Name 'TypeName'
        if ($remoteTypeName -is [string] -and $remoteTypeName -ceq 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException') { return $true }
        $exception = $exception.InnerException
    }
    if ($ErrorRecord.CategoryInfo.Reason -ceq 'ManagementObjectNotFoundException') { return $true }
    return ([string]$ErrorRecord.FullyQualifiedErrorId -cmatch '(?:^|,)(?:Microsoft\.Exchange\.Configuration\.Tasks\.)?ManagementObjectNotFoundException(?:,|$)|\[FailureCategory=Cmdlet-ManagementObjectNotFoundException\]')
}

function Get-Pra2LegacyRecipient {
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$LegacyDN)
    if ([string]::IsNullOrWhiteSpace($LegacyDN)) { throw 'Resource legacy identity is empty; absence is not verified.' }
    $lookupResults = [System.Collections.Generic.List[object]]::new()
    try {
        # RecipientIdParameter resolves LegacyExchangeDN and X500 proxies by equality; LegacyExchangeDN is not Get-Recipient OPATH.
        Get-Recipient -Identity $LegacyDN -IncludeSoftDeletedRecipients -ErrorAction Stop | ForEach-Object { $lookupResults.Add($_) }
    } catch {
        if ($lookupResults.Count -or -not (Test-Pra2RecipientNotFound -ErrorRecord $_)) { throw }
    }
    foreach ($recipient in $lookupResults) {
        if ([string](Get-PraValue $recipient 'LegacyExchangeDN' '') -ine $LegacyDN -and
            @((Get-PraValue $recipient 'EmailAddresses' @()) | ForEach-Object { [string]$_ }) -inotcontains "x500:$LegacyDN") {
            throw 'Resource legacy identity lookup returned a non-exact recipient; absence is not verified.'
        }
        $recipient
    }
}

function Get-Pra2ResourceGraphFilter {
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
        foreach ($address in @(Get-Pra2ResourceAddresses $record)) { $null = $proxies.Add($address) }
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

function Get-Pra2ResourceInventory {
    # Both inventories cover the exact selected-key union; Graph also covers the original immutable anchors.
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Records = @())
    if (-not $Records.Count) { throw 'Resource inventory requires selected records; absence is not verified.' }
    $keys = @(@(foreach ($record in $Records) {
            if (-not $record) { throw 'Resource inventory contains an empty selected record; absence is not verified.' }
            Get-Pra2ResourceKeys $record
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
        $immutableId = [string](Get-PraValue $record 'immutable_id' '')
        if ($immutableId) { $null = $immutableIds.Add($immutableId) }
    }
    $graphFilters = @(Get-Pra2ResourceGraphFilter -Records $Records -Keys $keys -ImmutableIds @($immutableIds))
    if (-not $graphFilters.Count) { throw 'Resource inventory has no Graph predicates; absence is not verified.' }
    # No SilentlyContinue here: an incomplete recipient inventory is not evidence that addresses are free.
    $recipients = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($filter in $filters) {
        foreach ($recipient in @(Get-Recipient -Filter $filter -ResultSize Unlimited -IncludeSoftDeletedRecipients -ErrorAction Stop)) {
            $identity = [string](Get-PraValue $recipient 'Guid' '')
            if (-not $identity) { $identity = [string](Get-PraValue $recipient 'ExternalDirectoryObjectId' '') }
            if ($identity -and $seen.ContainsKey($identity)) { continue }
            if ($identity) { $seen[$identity] = $true }
            $recipients.Add($recipient)
        }
    }
    foreach ($key in @($keys | Where-Object { $_.StartsWith('x500:') })) {
        foreach ($recipient in @(Get-Pra2LegacyRecipient -LegacyDN $key.Substring(5))) {
            $identity = [string](Get-PraValue $recipient 'Guid' '')
            if (-not $identity) { $identity = [string](Get-PraValue $recipient 'ExternalDirectoryObjectId' '') }
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
            $page = Invoke-Pra2Graph @parameters
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

function Get-Pra2ResourceCollisions {
    param([object]$Record, [object]$Inventory, [string]$OwnedId = '', [string]$OwnedGuid = '', [string]$OwnedContactId = '')
    $keys = @(Get-Pra2ResourceKeys $Record)
    if ($Inventory.PSObject.Properties['RecipientKeys'] -and @($keys | Where-Object { $_ -notin $Inventory.RecipientKeys }).Count) {
        throw 'Resource collision inventory does not cover all selected record keys; absence is not verified.'
    }
    $immutableId = [string](Get-PraValue $Record 'immutable_id' '')
    if ($Inventory.PSObject.Properties['GraphImmutableIds'] -and $immutableId -and $immutableId -cnotin $Inventory.GraphImmutableIds) {
        throw 'Resource collision inventory does not cover the selected immutable anchor; absence is not verified.'
    }
    foreach ($user in $Inventory.Users) {
        if ($OwnedId -and [string]$user.id -eq $OwnedId) { continue }
        $candidate = @(('smtp:' + [string]$user.userPrincipalName), ('smtp:' + [string](Get-PraValue $user 'mail' ''))) +
            @((Get-PraValue $user 'proxyAddresses' @()) | ForEach-Object { [string]$_ })
        $candidate += ('alias:' + [string](Get-PraValue $user 'mailNickname' ''))
        if (@($candidate | Where-Object { $_.ToLowerInvariant() -in $keys }).Count -or
            ($immutableId -and [string](Get-PraValue $user 'onPremisesImmutableId' '') -ceq $immutableId)) {
            "Existing Entra identity $($user.id) ($($user.userPrincipalName)) collides with this native resource; never adopted."
        }
    }
    foreach ($recipient in $Inventory.Recipients) {
        $id = [string](Get-PraValue $recipient 'ExternalDirectoryObjectId' '')
        $guid = [string](Get-PraValue $recipient 'ExchangeGuid' '')
        $contactId = [string](Get-PraValue $recipient 'Guid' '')
        if (($OwnedId -and $id -eq $OwnedId) -or ($OwnedGuid -and $guid -eq $OwnedGuid) -or
            ($OwnedContactId -and $contactId -eq $OwnedContactId)) { continue }
        $candidate = @(('smtp:' + [string](Get-PraValue $recipient 'PrimarySmtpAddress' '')), ('alias:' + [string](Get-PraValue $recipient 'Alias' ''))) +
            @((Get-PraValue $recipient 'EmailAddresses' @()) | ForEach-Object { [string]$_ })
        $legacy = [string](Get-PraValue $recipient 'LegacyExchangeDN' '')
        if ($legacy) { $candidate += "x500:$legacy" }
        if (@($candidate | Where-Object { $_.ToLowerInvariant() -in $keys }).Count) {
            "Existing Exchange recipient $($recipient.Identity) ($($recipient.RecipientTypeDetails), ID $id, ExchangeGuid $guid) collides with a resource address or alias; left untouched, including soft-deleted recipients."
        }
    }
}

function Get-Pra2OwnedResource {
    param([Parameter(Mandatory)][object]$Intent, [Parameter(Mandatory)][object]$Record)
    $recipients = @(Get-Recipient -Filter ("Name -eq '{0}'" -f ([string]$Intent.creationName).Replace("'", "''")) -ResultSize Unlimited -ErrorAction Stop)
    if (-not $recipients.Count) { return $null }
    if ($recipients.Count -ne 1) { throw 'Ambiguous resource creation identity; no mailbox is adopted.' }
    $recipient = $recipients[0]
    if (-not [bool](Get-PraValue $Intent 'creationAttempted' $false)) { throw 'Resource creation was never journaled as attempted; a pre-existing name is not adopted.' }
    $expectedType = "$($Record.kind)Mailbox"
    if ([string]$recipient.Name -cne [string]$Intent.creationName -or [string]$recipient.RecipientTypeDetails -ne $expectedType) { throw 'Resource creation name/type does not prove ownership.' }
    $id = [string]$recipient.ExternalDirectoryObjectId
    $guid = [string]$recipient.ExchangeGuid
    if (-not $id -or -not $guid -or $guid -eq [guid]::Empty.ToString() -or $guid -eq [string]$Record.exchange_guid) { throw 'Resource has no independent cloud identity/ExchangeGuid.' }
    $persistedId = [string](Get-PraValue $Intent 'cloudId' '')
    $persistedGuid = [string](Get-PraValue $Intent 'cloudGuid' '')
    if (($persistedId -and $id -ne $persistedId) -or ($persistedGuid -and $guid -ne $persistedGuid)) { throw 'Resource immutable ID or ExchangeGuid differs from the journal; nothing is changed.' }
    $mailbox = Get-Mailbox -Identity $id -ErrorAction Stop
    $marker = [string](Get-PraValue $mailbox ([string]$Intent.markerAttribute) '')
    # New-Mailbox has no custom-attribute parameter. The unpredictable token is embedded in Name atomically;
    # only this journaled creation name can bridge the crash before the marker and IDs were persisted.
    if ($marker -cne [string]$Intent.token -and ($marker -or ($persistedId -and -not (Get-PraValue $Intent 'markerPending' $false)))) { throw 'Resource ownership marker differs from the journal.' }
    $user = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id,userPrincipalName,onPremisesSyncEnabled,onPremisesImmutableId,accountEnabled,assignedPlans,licenseAssignmentStates,assignedLicenses,usageLocation" -f $id) -AllowNotFound
    if (-not $user) { throw 'Owned resource Graph identity is not yet visible; retry provisioning, never adopt another user.' }
    $exoUser = Get-User -Identity $id -ErrorAction Stop
    if ([string](Get-PraValue $user 'id' '') -ne $id -or
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

function Get-Pra2ResourceState {
    param([object]$Record, [object[]]$Records, [object]$Inventory, [AllowNull()][object]$Intent)
    if ((Get-Pra2ResourceProtectionMode $Intent) -eq 'None') { Assert-Pra2ResourceOrganizationUnheld }
    $owned = if ($Intent) { Get-Pra2OwnedResource -Intent $Intent -Record $Record } else { $null }
    if ($Intent -and (Get-PraValue $Intent 'cloudId' '') -and -not $owned) { throw 'Journal-owned active resource is missing; use Recover to verify deleted/inactive state.' }
    $collisions = @(Get-Pra2ResourceCollisions -Record $Record -Inventory $Inventory -OwnedId $(if ($owned) { $owned.Id } else { '' }))
    $keys = @(Get-Pra2ResourceKeys $Record)
    foreach ($other in @($Records | Where-Object { [string]$_.object_guid -ne [string]$Record.object_guid })) {
        if (@(Get-Pra2ResourceKeys $other | Where-Object { $_ -in $keys }).Count) { $collisions += "Competing selected snapshot object $($other.object_guid) claims a resource address/alias." }
    }
    if ($owned -and (Get-Pra2ResourceProtectionMode $Intent) -eq 'None') { Assert-Pra2UnprotectedResource -Owned $owned }
    return [pscustomobject]@{ NativeResource = $true; JournalNative = [bool]$Intent; ProtectionMode = (Get-Pra2ResourceProtectionMode $Intent); Owned = [bool]$owned; AbsenceVerified = $true; Collisions = $collisions
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

function ConvertTo-Pra2ResourceWorkDays {
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

function Read-Pra2ResourceSettings {
    param([object]$Record)
    $json = [string](Get-PraValue $Record 'resource_settings_json' '')
    if (-not $json) { throw 'Native resource settings are missing (v1 snapshot): Collect is required before CloudResource Convert.' }
    $settings = $json | ConvertFrom-Json -ErrorAction Stop
    if ([int](Get-PraValue $settings 'SchemaVersion' 0) -ne 2 -or -not $settings.PSObject.Properties['Capacity'] -or
        -not $settings.PSObject.Properties['ResourceCustom'] -or -not $settings.PSObject.Properties['Calendar'] -or
        -not $settings.PSObject.Properties['CalendarConfiguration'] -or -not $settings.PSObject.Properties['Directory'] -or
        -not $settings.PSObject.Properties['RoomLists'] -or -not $settings.PSObject.Properties['Delivery']) { throw 'Incomplete resource metadata/settings schema: a new Collect is required.' }
    Test-Pra2DeliverySettings -Delivery $settings.Delivery
    foreach ($field in $script:ResourceDirectoryFields) {
        if (-not $settings.Directory -or -not $settings.Directory.PSObject.Properties[$field] -or $settings.Directory.$field -isnot [string]) {
            throw "Resource descriptive property Directory.$field missing/invalid: a new Collect is required."
        }
    }
    if ($settings.RoomLists -isnot [array]) { throw 'Resource RoomLists must explicitly contain an array: a new Collect is required.' }
    if ([string]$Record.kind -ne 'Room' -and $settings.RoomLists.Count) { throw 'Only room mailboxes can belong to RoomLists.' }
    $seenLists = @{}
    foreach ($list in $settings.RoomLists) {
        Test-Pra2RoomListMetadata -List $list
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
    $null = ConvertTo-Pra2ResourceWorkDays -Days @($settings.CalendarConfiguration.WorkingHoursDays)
    return $settings
}

function Test-Pra2DeliverySettings {
    param([object]$Delivery)
    foreach ($field in @('RequireSenderAuthenticationEnabled','ModerationEnabled')) {
        if (-not $Delivery -or -not $Delivery.PSObject.Properties[$field] -or $Delivery.$field -isnot [bool]) {
            throw "Resource Delivery.$field missing/invalid: a new Collect is required (no delivery-policy broadening)."
        }
    }
    if ([string](Get-PraValue $Delivery 'SendModerationNotifications' '') -notin @('Never','Internal','Always')) { throw 'Invalid resource moderation notification setting: a new Collect is required.' }
    foreach ($field in $script:ResourceDeliveryRecipients) {
        if (-not $Delivery.PSObject.Properties[$field] -or $Delivery.$field -isnot [array]) { throw "Resource Delivery.$field must contain an explicit recipient array: a new Collect is required." }
        foreach ($smtp in $Delivery.$field) {
            if ($smtp -isnot [string] -or $smtp -notmatch '^[^@\s]+@[^@\s]+$') { throw "Resource Delivery.$field has an invalid SMTP recipient: a new Collect is required." }
        }
    }
    if ($Delivery.ModerationEnabled -and -not $Delivery.ModeratedBy.Count) { throw 'Resource moderation requires explicit moderators: a new Collect is required.' }
}

function Test-Pra2RoomListMetadata {
    param([object]$List)
    foreach ($field in @('SourceGuid','Name','DisplayName','Alias','PrimarySmtpAddress','LegacyExchangeDN')) {
        if (-not $List -or -not $List.PSObject.Properties[$field] -or $List.$field -isnot [string] -or ($field -ne 'LegacyExchangeDN' -and -not $List.$field)) {
            throw "Resource RoomLists.$field missing/invalid: a new Collect is required."
        }
    }
    if ([int](Get-PraValue $List 'SchemaVersion' 0) -ne 1 -or $List.PrimarySmtpAddress -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or
        -not $List.PSObject.Properties['HiddenFromAddressListsEnabled'] -or $List.HiddenFromAddressListsEnabled -isnot [bool] -or
        -not $List.PSObject.Properties['EmailAddresses'] -or $List.EmailAddresses -isnot [array]) { throw 'Incomplete room-list metadata: a new Collect is required.' }
    foreach ($address in $List.EmailAddresses) { if ($address -isnot [string] -or -not $address) { throw 'Invalid room-list address: a new Collect is required.' } }
    if ([string](Get-PraValue $List 'MemberJoinRestriction' '') -notin @('Closed','Open','ApprovalRequired') -or
        [string](Get-PraValue $List 'MemberDepartRestriction' '') -notin @('Closed','Open')) { throw 'Invalid room-list membership controls: a new Collect is required.' }
    Test-Pra2DeliverySettings -Delivery (Get-PraValue $List 'Delivery' $null)
}

function ConvertTo-Pra2RoomListRecord {
    param([object]$List)
    return [pscustomobject]@{ object_guid = $List.SourceGuid; kind = 'RoomList'; alias = $List.Alias; primary_smtp_address = $List.PrimarySmtpAddress
        email_addresses_json = (ConvertTo-Json -InputObject @($List.EmailAddresses) -Compress); legacy_exchange_dn = $List.LegacyExchangeDN
        immutable_id = ''; user_principal_name = ''; exchange_guid = '' }
}

function Test-Pra2ResourceConfigurationReadiness {
    param([object]$Record, [object[]]$FutureResources = @(), [hashtable]$ResourceIntents = @{},
        [hashtable]$RoomListIntents = @{}, [object[]]$RoomListPlan = @(), [switch]$AllowRoomListCreation)
    $settings = Read-Pra2ResourceSettings $Record
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
        $group = Get-Pra2ResourceCloudRoomList -List $list -Intent $listIntent -AllowMissing:$AllowRoomListCreation
        if ($group) { $null = Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop }
        else { Test-Pra2RoomListCreationReadiness -List $list }
        $memberCommand = Get-Command -Name Add-DistributionGroupMember -ErrorAction Stop
        if (-not $memberCommand.Parameters.ContainsKey('BypassSecurityGroupManagerCheck')) { throw 'Add-DistributionGroupMember -BypassSecurityGroupManagerCheck is not exposed; check room-list membership RBAC before Convert.' }
    }
    $references = @(foreach ($field in $script:ResourceRecipientFields) { [pscustomobject]@{ Control = 'booking'; Values = @($settings.Calendar.$field); Field = $field } }
        foreach ($field in $script:ResourceDeliveryRecipients) { [pscustomobject]@{ Control = 'delivery'; Values = @($settings.Delivery.$field); Field = $field } })
    foreach ($reference in $references) {
        $field = $reference.Field
        foreach ($smtp in $reference.Values) {
            if ($AllowRoomListCreation) {
                $futureLists = @($RoomListPlan | Where-Object { "smtp:$smtp" -in @(Get-Pra2ResourceAddresses (ConvertTo-Pra2RoomListRecord $_)) })
                if ($futureLists.Count -eq 1) { continue }
                if ($futureLists.Count -gt 1) { throw "Ambiguous planned room-list recipient $smtp; no policy is broadened." }
            }
            $future = @($FutureResources | Where-Object { "smtp:$smtp" -in @(Get-Pra2ResourceAddresses $_) })
            if ($future.Count) {
                if ($future.Count -ne 1) { throw "Resource $($reference.Control) recipient $smtp ($field) matches competing native snapshot objects; no policy is omitted or broadened." }
                $intent = $ResourceIntents[[string]$future[0].object_guid]
                if (-not $intent -or -not [string](Get-PraValue $intent 'cloudId' '')) {
                    throw "Resource $($reference.Control) recipient $smtp ($field) references a native resource not yet journal-owned/provisioned. Provision that referenced resource first, then resume this Convert batch. This resource is not created; no booking recipient is omitted or policy broadened."
                }
                $owned = Get-Pra2OwnedResource -Intent $intent -Record $future[0]
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

function Get-Pra2ResourceCloudRoomList {
    param([object]$List, [AllowNull()][object]$Intent, [switch]$AllowMissing)
    $record = ConvertTo-Pra2RoomListRecord $List
    $inventory = Get-Pra2ResourceInventory -Records @($record)
    if ($Intent) {
        if ([string]$Intent.sourceGuid -ne [string]$List.SourceGuid -or $Intent.operation -notin @('Create','Reuse')) { throw 'Invalid room-list journal identity/operation.' }
        if ($Intent.operation -eq 'Create') {
            if ([string]$Intent.token -notmatch '^[a-f0-9]{32}$' -or [string]$Intent.creationName -cnotmatch ("^PRA-L-[a-f0-9]{{8}}-[a-f0-9]{{8}}-{0}$" -f $Intent.token)) { throw 'Room-list journal ownership token/name is invalid; no group is adopted.' }
            $groups = @(Get-DistributionGroup -Filter ("Name -eq '{0}'" -f $Intent.creationName) -ResultSize Unlimited -ErrorAction Stop)
            if ($groups.Count -gt 1) { throw 'Ambiguous room-list ownership; no group is adopted.' }
            if ($groups.Count -and -not $Intent.creationAttempted) { throw 'Room-list creation was not journaled as attempted; no pre-existing group is adopted.' }
            if ($groups.Count) {
                $candidate = $groups[0]
                $marker = [string](Get-PraValue $candidate 'CustomAttribute15' '')
                if ([string]$candidate.Name -cne [string]$Intent.creationName -or
                    ($marker -cne [string]$Intent.token -and ($marker -or ($Intent.cloudGuid -and -not (Get-PraValue $Intent 'markerPending' $false)))) -or [string]$candidate.Guid -eq [string]$List.SourceGuid) { throw 'Room-list ownership differs from the journal; nothing is changed.' }
            }
        } else { $groups = @(Get-DistributionGroup -Identity ([string]$Intent.cloudGuid) -ErrorAction Stop) }
    } else {
        $groups = @(Get-DistributionGroup -Filter ("PrimarySmtpAddress -eq '{0}'" -f ([string]$List.PrimarySmtpAddress).Replace("'", "''")) -ResultSize Unlimited -ErrorAction Stop)
    }
    if ($groups.Count -eq 0) {
        if ($Intent -and $Intent.cloudGuid) { throw 'Journal-owned room list is missing; never replace its identity.' }
        $collisions = @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory)
        if ($collisions.Count) { throw ("Room-list collision: " + ($collisions -join '; ')) }
        if ($AllowMissing) { return $null }
        throw "Room list $($List.PrimarySmtpAddress) is missing; resume Convert to provision it before its rooms."
    }
    if ($groups.Count -ne 1) { throw "Room list $($List.PrimarySmtpAddress) is ambiguous; no group is adopted." }
    $group = $groups[0]
    $pendingCreation = $Intent -and $Intent.operation -eq 'Create' -and $Intent.status -ne 'Done'
    if ([string](Get-PraValue $group 'RecipientTypeDetails' '') -ne 'RoomList' -or
        -not $group.PSObject.Properties['IsDirSynced'] -or $group.IsDirSynced -ne $false -or
        [string](Get-PraValue $group 'PrimarySmtpAddress' '') -ine [string]$List.PrimarySmtpAddress -or
        [string](Get-PraValue $group 'Alias' '') -ine [string]$List.Alias -or
        -not [string](Get-PraValue $group 'Guid' '') -or [string]$group.Guid -eq [guid]::Empty.ToString() -or
        -not $group.PSObject.Properties['HiddenFromAddressListsEnabled'] -or
        $group.HiddenFromAddressListsEnabled -isnot [bool] -or
        (-not $pendingCreation -and [bool]$group.HiddenFromAddressListsEnabled -ne [bool]$List.HiddenFromAddressListsEnabled)) {
        throw "Room list $($List.PrimarySmtpAddress) must be a cloud-only RoomList with the original primary address, alias and visibility; no unrelated group is converted or taken over."
    }
    if ($Intent -and (($Intent.cloudGuid -and [string]$Intent.cloudGuid -ne [string]$group.Guid) -or
            ($Intent.cloudId -and [string]$Intent.cloudId -ne [string](Get-PraValue $group 'ExternalDirectoryObjectId' '')))) { throw 'Room-list immutable identity differs from the journal; nothing is changed.' }
    $collisions = @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory -OwnedContactId ([string]$group.Guid))
    if ($collisions.Count) { throw ("Room-list collision: " + ($collisions -join '; ')) }
    return $group
}

function Test-Pra2RoomListCreationReadiness {
    param([object]$List)
    Test-Pra2RoomListMetadata $List
    foreach ($spec in @(
            @{ Name = 'New-DistributionGroup'; Fields = @('Name','Alias','DisplayName','PrimarySmtpAddress','RoomList','IgnoreNamingPolicy','MemberJoinRestriction','MemberDepartRestriction','RequireSenderAuthenticationEnabled') },
            @{ Name = 'Set-DistributionGroup'; Fields = @('EmailAddresses','HiddenFromAddressListsEnabled','CustomAttribute15','BypassSecurityGroupManagerCheck','MemberJoinRestriction','MemberDepartRestriction') + $script:ResourceDeliveryFields })) {
        $command = Get-Command -Name $spec.Name -ErrorAction Stop
        foreach ($field in $spec.Fields) {
            if (-not $command.Parameters.ContainsKey($field)) { throw "$($spec.Name) -$field is not exposed; check room-list creation RBAC before Convert." }
        }
    }
    foreach ($field in $script:ResourceDeliveryRecipients) { $null = Resolve-Pra2ResourceRecipients -Addresses $List.Delivery.$field -Control 'Room-list delivery' }
}

function Set-Pra2RoomListConfiguration {
    param([object]$List, [object]$Intent, [object]$Group)
    if ($Intent.operation -ne 'Create') { throw 'Existing reused room lists are not reconfigured.' }
    $parameters = @{ Identity = [string]$Group.Guid; EmailAddresses = @(Get-Pra2ResourceAddresses (ConvertTo-Pra2RoomListRecord $List))
        HiddenFromAddressListsEnabled = [bool]$List.HiddenFromAddressListsEnabled; CustomAttribute15 = [string]$Intent.token
        MemberJoinRestriction = $List.MemberJoinRestriction; MemberDepartRestriction = $List.MemberDepartRestriction
        BypassSecurityGroupManagerCheck = $true; ErrorAction = 'Stop' }
    foreach ($field in $script:ResourceDeliveryFields) {
        if ($field -in $script:ResourceDeliveryRecipients) { $parameters[$field] = @(Resolve-Pra2ResourceRecipients -Addresses $List.Delivery.$field -Control 'Room-list delivery') }
        else { $parameters[$field] = $List.Delivery.$field }
    }
    Set-DistributionGroup @parameters
    $actual = Get-DistributionGroup -Identity ([string]$Group.Guid) -ErrorAction Stop
    foreach ($field in @('HiddenFromAddressListsEnabled','CustomAttribute15','MemberJoinRestriction','MemberDepartRestriction') + $script:ResourceDeliveryFields) {
        if (-not $actual.PSObject.Properties[$field]) { throw "Room-list configuration verification failed: $field is not exposed." }
        if ($field -in $script:ResourceDeliveryRecipients) {
            $wanted = @($parameters[$field] | ForEach-Object { [string](Get-Recipient -Identity $_ -ErrorAction Stop).Guid } | Sort-Object)
            $got = @($actual.$field | Where-Object { $null -ne $_ } | ForEach-Object { [string](Get-Recipient -Identity ([string]$_) -ErrorAction Stop).Guid } | Sort-Object)
        } else { $wanted = @([string]$parameters[$field]); $got = @([string]$actual.$field) }
        if (($wanted -join '|') -cne ($got -join '|')) { throw "Room-list configuration verification failed: $field." }
    }
    foreach ($address in $parameters.EmailAddresses) { if ($address -notin @($actual.EmailAddresses | ForEach-Object { [string]$_ })) { throw "Room-list address verification failed: $address." } }
    if ([string]$actual.DisplayName -cne [string]$List.DisplayName) { throw 'Room-list display name verification failed.' }
}

function Resolve-Pra2ResourceRecipients {
    param([AllowEmptyCollection()][object[]]$Addresses, [string]$Control)
    foreach ($smtp in $Addresses) {
        $resolved = @(Get-Recipient -Identity ([string]$smtp) -ErrorAction Stop)
        if ($resolved.Count -ne 1) { throw "$Control recipient $smtp is unresolved/ambiguous; policy is not broadened." }
        [string]$resolved[0].Identity
    }
}

function Get-Pra2ResourceGroup {
    param([Parameter(Mandatory)][string]$Id, [object[]]$Skus = @(), [int]$NewUnits = 0, [int]$OtherUnits = 0)
    $group = Invoke-Pra2Graph -Uri ("v1.0/groups/{0}?`$select=id,displayName,assignedLicenses,onPremisesSyncEnabled,onPremisesLastSyncDateTime,onPremisesSecurityIdentifier,securityEnabled,mailEnabled,groupTypes" -f $Id) -AllowNotFound
    if (-not $group) { throw 'Resources.LicensingGroupId does not exist.' }
    if ($null -ne $group.onPremisesSyncEnabled -or (Get-PraValue $group 'onPremisesLastSyncDateTime' $null) -or
        (Get-PraValue $group 'onPremisesSecurityIdentifier' '') -or -not $group.securityEnabled -or $group.mailEnabled -or
        @((Get-PraValue $group 'groupTypes' @()) | Where-Object { $_ -in @('Unified','DynamicMembership') }).Count) {
        throw 'Resources.LicensingGroupId must be a never-synchronised, assigned-membership cloud-only security licensing group (not a mail/Microsoft 365/dynamic group or a formerly AD-synchronised group). A user-group SOA rollback must not remove the active resource licence.'
    }
    $enterprise = @(); $enterpriseAssignments = @()
    foreach ($assignment in @($group.assignedLicenses)) {
        $sku = $Skus | Where-Object { [string]$_.skuId -eq [string]$assignment.skuId } | Select-Object -First 1
        if (-not $sku) { throw 'Resource group assigns an unknown SKU; capacity cannot be verified.' }
        if (($NewUnits + $OtherUnits) -gt 0 -and (Get-Pra2FreeUnit $sku) -lt ($NewUnits + $OtherUnits)) { throw "Resource licensing capacity insufficient for $($NewUnits + $OtherUnits) new unit(s) of $($sku.skuPartNumber); resources keep their licence throughout active lifetime." }
        $enabled = @($sku.servicePlans | Where-Object { [string]$_.servicePlanName -eq 'EXCHANGE_S_ENTERPRISE' -and [string]$_.servicePlanId -notin @($assignment.disabledPlans) } |
            ForEach-Object { [string]$_.servicePlanId })
        $enterprise += $enabled
        if ($enabled.Count) { $enterpriseAssignments += [pscustomobject]@{ SkuId = [string]$sku.skuId; PlanIds = $enabled } }
    }
    if (-not $enterprise.Count) { throw 'Resource licensing group has no enabled EXCHANGE_S_ENTERPRISE plan (E3/E5 required for the eDiscovery hold).' }
    return [pscustomobject]@{ Id = $Id; Group = $group; EnterprisePlans = $enterprise; EnterpriseAssignments = $enterpriseAssignments }
}

function Test-Pra2ResourceLicence {
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

function Get-Pra2ResourceDisabledCalendar {
    param([Parameter(Mandatory)][string]$Identity)
    return @{ Identity = $Identity; ErrorAction = 'Stop'; AutomateProcessing = 'None'
        AllBookInPolicy = $false; AllRequestInPolicy = $false; AllRequestOutOfPolicy = $false
        ForwardRequestsToDelegates = $false; ProcessExternalMeetingMessages = $false
        ResourceDelegates = @(); BookInPolicy = @(); RequestInPolicy = @(); RequestOutOfPolicy = @() }
}

function Assert-Pra2ResourceCalendar {
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
        if (($wanted -join '|') -cne ($got -join '|')) { throw "Resource calendar verification failed: $field." }
    }
}

function Set-Pra2ResourceProvisioningDisabled {
    param([Parameter(Mandatory)][string]$Identity)
    # Only the just-created, pinned identity or a verified unfinished journal-owned resource reaches this helper.
    $errors = [System.Collections.Generic.List[string]]::new()
    try {
        $calendar = Get-Pra2ResourceDisabledCalendar -Identity $Identity
        Set-CalendarProcessing @calendar
        Assert-Pra2ResourceCalendar -Parameters $calendar -Fields @($calendar.Keys | Where-Object { $_ -notin @('Identity','ErrorAction') })
    } catch { $errors.Add("Booking disable failed: $($_.Exception.Message)") }
    try {
        Set-Mailbox -Identity $Identity -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop
        if (-not $mailbox.PSObject.Properties['HiddenFromAddressListsEnabled'] -or $mailbox.HiddenFromAddressListsEnabled -ne $true) {
            throw 'Hidden provisioning verification failed.'
        }
    } catch { $errors.Add("Hidden provisioning failed: $($_.Exception.Message)") }
    if ($errors.Count) { throw ($errors -join ' | ') }
}

function Enable-Pra2ResourceBooking {
    param([object]$Record, [object]$Owned, [hashtable]$Calendar)
    $hidden = [bool](Get-PraValue $Record 'hidden_from_address_lists' $false)
    Set-Mailbox -Identity $Owned.Id -HiddenFromAddressListsEnabled $hidden -ErrorAction Stop
    $mailbox = Get-Mailbox -Identity $Owned.Id -ErrorAction Stop
    if (-not $mailbox.PSObject.Properties['HiddenFromAddressListsEnabled'] -or $mailbox.HiddenFromAddressListsEnabled -ne $hidden) {
        throw 'Resource final hidden verification failed.'
    }
    # Booking activation is the last cloud write/proof: no later mandatory metadata read can leave it open.
    Set-CalendarProcessing @Calendar
    Assert-Pra2ResourceCalendar -Parameters $Calendar -Fields @($Calendar.Keys | Where-Object { $_ -notin @('Identity','ErrorAction') })
}

function Set-Pra2ResourceConfiguration {
    param([object]$Record, [object]$Owned, [object]$Intent)
    $settings = Read-Pra2ResourceSettings $Record
    $parameters = @{ Identity = $Owned.Id; EmailAddresses = @(Get-Pra2ResourceAddresses $Record); ResourceCapacity = [int]$settings.Capacity
        ResourceCustom = @($settings.ResourceCustom); HiddenFromAddressListsEnabled = $true; ErrorAction = 'Stop' }
    $parameters[[string]$Intent.markerAttribute] = [string]$Intent.token
    $attributes = [string](Get-PraValue $Record 'custom_attributes_json' '')
    if ($attributes) {
        $custom = $attributes | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in $custom.PSObject.Properties) {
            if ($property.Name -match '^CustomAttribute([1-9]|1[0-5])$' -and $property.Name -ne [string]$Intent.markerAttribute) { $parameters[$property.Name] = [string]$property.Value }
        }
    }
    foreach ($field in $script:ResourceDeliveryFields) {
        if ($field -in $script:ResourceDeliveryRecipients) {
            $parameters[$field] = @(Resolve-Pra2ResourceRecipients -Addresses $settings.Delivery.$field -Control 'Delivery')
        } else { $parameters[$field] = $settings.Delivery.$field }
    }
    Set-Mailbox @parameters
    $directory = @{ Identity = $Owned.Id; Confirm = $false; ErrorAction = 'Stop' }
    foreach ($field in $script:ResourceDirectoryFields) {
        $directory[$field] = if ($settings.Directory.$field) { [string]$settings.Directory.$field } else { $null }
    }
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
            $value = @(Resolve-Pra2ResourceRecipients -Addresses @($value) -Control 'Booking')
        }
        $calendar[$field] = $value
    }
    # Stage non-booking settings without enabling the source policy before mandatory proofs and membership.
    $stagedCalendar = $calendar.Clone()
    $disabledCalendar = Get-Pra2ResourceDisabledCalendar -Identity $Owned.Id
    foreach ($field in $disabledCalendar.Keys) { $stagedCalendar[$field] = $disabledCalendar[$field] }
    Set-CalendarProcessing @stagedCalendar
    $work = @{ Identity = $Owned.Id; ErrorAction = 'Stop' }
    foreach ($field in $script:ResourceWorkFields) {
        if ($field -ne 'WorkingHoursDays') { $work[$field] = $settings.CalendarConfiguration.$field }
    }
    # The portable snapshot key is WorkingHoursDays; Exchange reads and writes WorkDays.
    $work.WorkDays = @(ConvertTo-Pra2ResourceWorkDays -Days @($settings.CalendarConfiguration.WorkingHoursDays))
    Set-MailboxCalendarConfiguration @work
    Assert-Pra2ResourceCalendar -Parameters $stagedCalendar -Fields @($stagedCalendar.Keys | Where-Object { $_ -notin @('Identity','ErrorAction') })
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
            elseif ($field -eq 'WorkingHoursDays') { ConvertTo-Pra2ResourceWorkDays -Days @($readValue) }
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
    foreach ($address in @(Get-Pra2ResourceAddresses $Record)) {
        if ($address -notin @($mailbox.EmailAddresses | ForEach-Object { [string]$_ })) { throw "Resource address not applied: $address." }
    }
    if ([string]$mailbox.PrimarySmtpAddress -ine [string]$Record.primary_smtp_address) { throw 'Resource primary SMTP verification failed.' }
    return $calendar
}

function Set-Pra2ResourceRoomListMembership {
    param([object]$Record, [object]$Owned, [hashtable]$RoomListIntents = @{})
    $settings = Read-Pra2ResourceSettings $Record
    foreach ($list in $settings.RoomLists) {
        $group = Get-Pra2ResourceCloudRoomList -List $list -Intent $RoomListIntents[[string]$list.SourceGuid]
        $members = @(Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop)
        if (-not @($members | Where-Object { [string](Get-PraValue $_ 'ExternalDirectoryObjectId' '') -eq [string]$Owned.Id }).Count) {
            Add-DistributionGroupMember -Identity ([string]$group.Guid) -Member $Owned.Id -BypassSecurityGroupManagerCheck -ErrorAction Stop
        }
        $members = @(Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited -ErrorAction Stop)
        if (@($members | Where-Object { [string](Get-PraValue $_ 'ExternalDirectoryObjectId' '') -eq [string]$Owned.Id }).Count -ne 1) {
            throw "Resource room-list membership verification failed: $($list.PrimarySmtpAddress)."
        }
    }
}

function Get-Pra2ResourceRoute {
    param([hashtable]$Config, [object]$Record, [switch]$RequireReady)
    if ($RequireReady -and -not [bool](Get-PraValue $Config.Resources 'OnPremisesReady' $false)) { throw 'Resources.OnPremisesReady must be explicitly true before native resource deletion: rebuild the on-premises resource first.' }
    $domain = [string](Get-PraValue $Config.Resources 'OnPremisesRoutingDomain' '')
    if ($domain -notmatch '^(?=.{1,253}$)(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,63}$') { throw 'Resources.OnPremisesRoutingDomain must be a valid hybrid routing DNS domain.' }
    $domains = @(Get-Pra2ResourceAddresses $Record | Where-Object { $_ -match '^smtp:' } | ForEach-Object { ($_ -split '@')[-1] })
    $cloudDomain = [string](Get-PraValue (Get-PraValue $Config 'Cloud' @{}) 'Organization' '')
    if ($domain -in $domains -or $domain -ieq $cloudDomain -or $domain -match '\.onmicrosoft\.com$') { throw 'Resource routing domain loops to an original SMTP/cloud domain; use the on-premises hybrid routing domain.' }
    $alias = [string](Get-PraValue $Record 'alias' '')
    if (-not $alias) { $alias = ([string]$Record.primary_smtp_address -split '@')[0] }
    if ($alias -notmatch '^[a-zA-Z0-9._-]+$') { throw 'Resource alias cannot form a safe routing address.' }
    return "$alias@$domain"
}
#endregion

#region Readiness rules (pure functions) -----------------------------------------------------------

function Get-Pra2Readiness {
    <#
    .SYNOPSIS
        Rules for one object of the snapshot: can Convert run, and what will happen. Pure function.
    .PARAMETER Record
        Row of the mailbox table.
    .PARAMETER State
        Output of Get-Pra2CloudState (or the same shape in tests).
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
    if ((Test-Pra2NativeResource -Config $Config -Record $Record) -or (Get-PraValue $State 'NativeResource' $false)) {
        try { $null = Read-Pra2ResourceSettings -Record $Record }
        catch { & $add 'Error' 'ExchangeOnline' 'RESOURCE_COLLECT_REQUIRED' $_.Exception.Message }
        if (-not $Config.Scope.ConvertRooms -and -not (Get-PraValue $State 'JournalNative' $false)) { & $add 'Error' 'ExchangeOnline' 'KIND_NOT_SUPPORTED' 'CloudResource also requires Scope.ConvertRooms = true.' }
        foreach ($collision in @((Get-PraValue $State 'Collisions' @()) | ForEach-Object { $_ })) { & $add 'Error' 'Entra' 'RESOURCE_COLLISION' ([string]$collision) }
        if (-not (Get-PraValue $State 'AbsenceVerified' $false)) { & $add 'Error' 'Entra' 'RESOURCE_ABSENCE_UNVERIFIED' 'Cloud identities and all Exchange recipients must be checked successfully; Graph absence alone is insufficient.' }
        # A missing or unknown mode never defaults to the unprotected lifecycle.
        $mode = Get-PraValue $State 'ProtectionMode' ''
        if ($mode -isnot [string] -or $mode -cnotin @('None', 'CaseHold')) {
            & $add 'Error' 'Holds' 'RESOURCE_PROTECTION_UNKNOWN' "Native resource protection mode '$mode' is missing or unknown; no lifecycle is assumed and nothing is changed."
            return $findings.ToArray()
        }
        if ($mode -eq 'CaseHold' -and [string](Get-PraValue (Get-PraValue $Config 'Cloud' @{}) 'DefaultUsageLocation' '') -notmatch '^[A-Z]{2}$') {
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
        else { & $add 'Warn' 'ExchangeOnline' 'GUID_MISMATCH' ("MailUser ExchangeGuid {0} differs from the on-premises mailbox {1}." -f $State.ExchangeGuid, $onPremGuid) }
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
    $mailboxPlans = @(Get-Pra2MailboxPlanId -Skus $Skus)
    $plans = @(if ($mailboxPlans.Count) { $plans | Where-Object { [string]$_.servicePlanId -in $mailboxPlans } } else { $plans | Where-Object { [string]$_.service -eq 'exchange' } })
    if ($plans.Count) { & $add 'Error' 'Licence' 'EXCHANGE_PLAN_PRESENT' 'An Exchange Online mailbox plan is already enabled on this object: clearing the ExchangeGuid would not create the mailbox while the plan stays (test T2c), so Convert refuses it. Remove that plan from the object before the disaster, or exclude it from the scope.' }
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
        $text = 'Hold(s) on the object: {0}. They follow the cloud mailbox; a user hold must be lifted during Recover (tests T8, H).' -f ($holds -join ', ')
        & $add $(if ($isUser) { 'Warn' } else { 'Info' }) 'Holds' 'HOLD_PRESENT' $text
    }
    $tag = [string]$Config.Retention.TagValue
    if ($tag -and [string]$State.TagValue -eq $tag) { & $add 'Warn' 'Holds' 'TAG_ALREADY_SET' ("{0} is already {1} in Exchange Online." -f $Config.Retention.TagAttribute, $tag) }
    if ($isUser -and $tag -and [string]$State.TagValue -eq $tag) { & $add 'Warn' 'Holds' 'USER_TAGGED' 'A user must not carry the retention tag: a retention policy hold would block his Recover.' }
    return $findings.ToArray()
}

function Get-Pra2TrusteeFinding {
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

function Wait-Pra2Condition {
    <#
    .SYNOPSIS
        Calls -Test every IntervalSeconds until it returns a true value or TimeoutSeconds is reached.
    .OUTPUTS
        @{ Ok = bool; Seconds = elapsed; Last = last value of -Test }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Test, [Parameter(Mandatory)][int]$TimeoutSeconds, [int]$IntervalSeconds = 20)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    do {
        $last = & $Test
        if ($last) { return [pscustomobject]@{ Ok = $true; Seconds = [int]$clock.Elapsed.TotalSeconds; Last = $last } }
        if ($clock.Elapsed.TotalSeconds -ge $TimeoutSeconds) { break }
        Start-Sleep -Seconds ([Math]::Max(1, [Math]::Min($IntervalSeconds, $TimeoutSeconds - [int]$clock.Elapsed.TotalSeconds)))
    } while ($true)
    return [pscustomobject]@{ Ok = $false; Seconds = [int]$clock.Elapsed.TotalSeconds; Last = $last }
}

function Get-Pra2ExoState {
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

function Set-Pra2SourceOfAuthority {
    <# Source of authority of a user or a group: cloud (isCloudManaged = true) or AD. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][bool]$Cloud, [ValidateSet('users','groups')][string]$Type = 'users')
    if (-not $PSCmdlet.ShouldProcess("$Type/$Id", "isCloudManaged = $Cloud")) { return }
    $null = Invoke-Pra2Graph -Method PATCH -Uri ("v1.0/{0}/{1}/onPremisesSyncBehavior" -f $Type, $Id) -Body @{ isCloudManaged = $Cloud }
}

function Get-Pra2SourceOfAuthority {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Id, [ValidateSet('users','groups')][string]$Type = 'users')
    $behavior = Invoke-Pra2Graph -Uri ("v1.0/{0}/{1}/onPremisesSyncBehavior?`$select=isCloudManaged" -f $Type, $Id) -AllowNotFound
    if ($behavior) { return [bool]$behavior.isCloudManaged }
    return $null
}

function Set-Pra2UsageLocation {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][ValidatePattern('^[A-Z]{2}$')][string]$Country)
    if (-not $PSCmdlet.ShouldProcess($UserId, "usageLocation = $Country")) { return }
    $null = Invoke-Pra2Graph -Method PATCH -Uri ("v1.0/users/{0}" -f $UserId) -Body @{ usageLocation = $Country }
}

function Set-Pra2GroupMember {
    <# Adds or removes a member of a group; already a member / already removed is not an error. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][ValidateSet('Add','Remove')][string]$Operation)
    if (-not $PSCmdlet.ShouldProcess("group $GroupId", "$Operation member $UserId")) { return }
    try {
        if ($Operation -eq 'Add') { $null = Invoke-Pra2Graph -Method POST -Uri ("v1.0/groups/{0}/members/`$ref" -f $GroupId) -Body @{ '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$UserId" } }
        else { $null = Invoke-Pra2Graph -Method DELETE -Uri ("v1.0/groups/{0}/members/{1}/`$ref" -f $GroupId, $UserId) -AllowNotFound }
    } catch {
        if ($Operation -eq 'Add' -and "$_" -match 'already exist') { return }
        throw
    }
}

function Test-Pra2GroupMember {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$GroupId, [Parameter(Mandatory)][string]$UserId)
    $result = Invoke-Pra2Graph -Method POST -Uri ("v1.0/users/{0}/checkMemberGroups" -f $UserId) -Body @{ groupIds = @($GroupId) }
    return @($result.value) -contains $GroupId
}

function Get-Pra2DisabledPlanForExchangeOnly {
    <# disabledPlans that leave only the Exchange plan(s) of a SKU enabled (licence given for the mailbox only). #>
    param([Parameter(Mandatory)][object]$Sku)
    return @($Sku.servicePlans | Where-Object { [string]$_.servicePlanName -notmatch $script:ExchangePlanPattern -and [string]$_.appliesTo -eq 'User' } | ForEach-Object { [string]$_.servicePlanId })
}

function Set-Pra2Licence {
    <# Assigns (with disabled plans) or removes one SKU directly on a user. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][string]$SkuId, [Parameter(Mandatory)][ValidateSet('Add','Remove')][string]$Operation,
        [string[]]$DisabledPlans = @())
    if (-not $PSCmdlet.ShouldProcess($UserId, "$Operation licence $SkuId")) { return }
    $body = if ($Operation -eq 'Add') { @{ addLicenses = @(@{ skuId = $SkuId; disabledPlans = @($DisabledPlans) }); removeLicenses = @() } }
        else { @{ addLicenses = @(); removeLicenses = @($SkuId) } }
    # A first assignment right after the usageLocation change can be refused (400) for a few seconds (lab 2026-10-07).
    for ($attempt = 1; ; $attempt++) {
        try { $null = Invoke-Pra2Graph -Method POST -Uri ("v1.0/users/{0}/assignLicense" -f $UserId) -Body $body; return }
        catch { if ($Operation -ne 'Add' -or $attempt -ge 4 -or "$($_.Exception.Message)" -notmatch '\(400\)|BadRequest') { throw }; Start-Sleep -Seconds 15 }
    }
}

function Get-Pra2UserLicence {
    <# Licence assignment states of a user (direct and by group). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)
    $user = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select=id,usageLocation,assignedLicenses,licenseAssignmentStates,onPremisesSyncEnabled" -f $UserId)
    return $user
}

function Get-Pra2DirectLicence {
    <# Direct licence assignments of a user (not the ones inherited from a group): skuId and sorted disabledPlans. Pure function. #>
    param([AllowNull()][object]$Licence)
    if (-not $Licence -or -not $Licence.PSObject.Properties['licenseAssignmentStates']) { return @() }
    return @($Licence.licenseAssignmentStates | Where-Object { $_ -and -not $_.assignedByGroup } | ForEach-Object {
            [pscustomobject]@{ skuId = [string]$_.skuId; disabledPlans = @($_.disabledPlans | ForEach-Object { [string]$_ } | Sort-Object) } })
}

function Enable-Pra2LicencePlan {
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
    $current = Get-Pra2DirectLicence (Get-Pra2UserLicence -UserId $UserId) | Where-Object { $_.skuId -eq $SkuId } | Select-Object -First 1
    if ($current) {
        $keep = @($current.disabledPlans | Where-Object { $_ -notin $PlanIds })
        if ($keep.Count -eq @($current.disabledPlans).Count) { return 'Present' }
        Set-Pra2Licence -UserId $UserId -SkuId $SkuId -Operation Add -DisabledPlans $keep
        return 'Enabled'
    }
    Set-Pra2Licence -UserId $UserId -SkuId $SkuId -Operation Add -DisabledPlans $DisabledPlans
    return 'Assigned'
}

function Restore-Pra2Licence {
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
    $current = Get-Pra2DirectLicence (Get-Pra2UserLicence -UserId $UserId) | Where-Object { $_.skuId -eq $SkuId } | Select-Object -First 1
    if ($before) {
        $wanted = @($before.disabledPlans | ForEach-Object { [string]$_ } | Sort-Object)
        if ($current -and (@($current.disabledPlans) -join ',') -eq ($wanted -join ',')) { return 'Unchanged' }
        Set-Pra2Licence -UserId $UserId -SkuId $SkuId -Operation Add -DisabledPlans $wanted
        return 'Restored'
    }
    if ($current) { Set-Pra2Licence -UserId $UserId -SkuId $SkuId -Operation Remove; return 'Removed' }
    return 'Absent'
}

function Remove-Pra2Identity {
    <# Deletes a user from Entra ID, then permanently from the recycle bin (shared mailbox Recover). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$UserId, [Parameter(Mandatory)][ValidateSet('Delete','Purge')][string]$Operation)
    if (-not $PSCmdlet.ShouldProcess($UserId, "$Operation identity")) { return }
    if ($Operation -eq 'Delete') { $null = Invoke-Pra2Graph -Method DELETE -Uri ("v1.0/users/{0}" -f $UserId) -AllowNotFound }
    else { $null = Invoke-Pra2Graph -Method DELETE -Uri ("v1.0/directory/deletedItems/{0}" -f $UserId) -AllowNotFound }
}

function Test-Pra2DeletedIdentity {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserId)
    return [bool](Invoke-Pra2Graph -Uri ("v1.0/directory/deletedItems/{0}?`$select=id" -f $UserId) -AllowNotFound)
}

function Connect-Pra2Compliance {
    <# Security & Compliance (search-only session) for the eDiscovery case hold, case hold cmdlets only. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($Context.ContainsKey('ComplianceConnected') -and $Context['ComplianceConnected']) { return }
    $cloud = $Context.Config.Cloud
    # -CommandName is required: the full session also loads Get-Recipient and Get-User, which hide the Exchange
    # Online cmdlets and return a stale view (a rolled back user still seen as UserMailbox - lab 2026-10-07).
    Connect-IPPSSession -AppId $cloud.AppId -CertificateThumbprint $cloud.CertificateThumbprint -Organization $cloud.Organization -ShowBanner:$false -EnableSearchOnlySession `
        -CommandName Get-CaseHoldPolicy, Set-CaseHoldPolicy, New-CaseHoldPolicy, New-CaseHoldRule -ErrorAction Stop
    $Context['ComplianceConnected'] = $true
}

# An eDiscovery case hold policy holds 1,000 mailboxes at most: the tool uses the configured policy, then <name>-02,
# <name>-03... in the same case (created with a rule that holds everything). Set-CaseHoldPolicy takes many locations
# in one call (12 to 15 s per call in the lab, whatever their number).
$script:HoldLocationChunk = 100

function Get-Pra2HoldPolicySet {
    <#
    .SYNOPSIS
        The case hold policies of the tool: the configured one and its siblings <name>-NN in the same case, with their
        tag (UniH + GUID) and their locations (Entra object IDs). A sibling being deleted (Mode PendingDeletion, still
        listed for hours - lab 9 Oct) is left out, but its number is not used again.
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

function Test-Pra2ToolHold {
    <# Is one of the holds of an object a case hold of the tool? #>
    param([Parameter(Mandatory)][object]$Set, [AllowEmptyCollection()][string[]]$Holds = @())
    foreach ($hold in $Holds) { if ($hold -and $Set.Tags.Contains($hold)) { return $true } }
    return $false
}

function Invoke-Pra2HoldChange {
    <#
    One Set-CaseHoldPolicy call for many locations. 'failed to be deployed' while the change is recorded (lab 7-8 Oct)
    counts as done for the locations the policy shows afterwards. Returns the IDs that were not changed.
    #>
    param([Parameter(Mandatory)][string]$Policy, [Parameter(Mandatory)][string[]]$Id, [Parameter(Mandatory)][ValidateSet('Add', 'Remove')][string]$Operation)
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

function Add-Pra2HoldLocationSet {
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
    foreach ($one in @($Id | Where-Object { $_ } | Select-Object -Unique)) {
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
            $created = New-CaseHoldPolicy -Name $name -Case $Set.CaseId -Enabled $true -ErrorAction Stop
            $null = New-CaseHoldRule -Name "$name-Rule" -Policy $name -ErrorAction Stop
            $policy = [pscustomobject]@{ Name = $name; Tag = ('UniH{0}' -f $created.Guid); Locations = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); Count = 0 }
            $Set.Policies.Add($policy); [void]$Set.Tags.Add($policy.Tag); $Set.Taken.Add($number)
        }
        $size = [Math]::Min($script:HoldLocationChunk, [Math]::Min($Set.Limit - $policy.Count, $todo.Count - $next))
        $chunk = [string[]]@($todo.GetRange($next, $size))
        $next += $size
        $missed = @()
        try { $missed = @(Invoke-Pra2HoldChange -Policy $policy.Name -Id $chunk -Operation Add) }
        catch { foreach ($one in $chunk) { $result.Failed[$one] = "case hold policy $($policy.Name): $($_.Exception.Message -replace '\s+', ' ')" }; continue }
        foreach ($one in $chunk) {
            if ($missed -contains $one) { $result.Failed[$one] = "not added to the case hold policy $($policy.Name)"; continue }
            [void]$policy.Locations.Add($one); $policy.Count++; $result.Policy[$one] = $policy.Name
        }
    }
    return $result
}

function Remove-Pra2HoldLocationSet {
    <# Takes many Entra object IDs off the policy of the tool that holds each one (one call per policy and 100 IDs). Returns Failed (ID -> message). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Set, [AllowEmptyCollection()][string[]]$Id = @())
    $result = [pscustomobject]@{ Failed = @{} }
    foreach ($policy in @($Set.Policies)) {
        $mine = @($Id | Where-Object { $_ -and $policy.Locations.Contains($_) } | Select-Object -Unique)
        if (-not $mine.Count -or -not $PSCmdlet.ShouldProcess($policy.Name, "Remove $($mine.Count) location(s)")) { continue }
        for ($start = 0; $start -lt $mine.Count; $start += $script:HoldLocationChunk) {
            $chunk = [string[]]@($mine[$start..([Math]::Min($mine.Count, $start + $script:HoldLocationChunk) - 1)])
            $missed = @()
            try { $missed = @(Invoke-Pra2HoldChange -Policy $policy.Name -Id $chunk -Operation Remove) }
            catch { foreach ($one in $chunk) { $result.Failed[$one] = "case hold policy $($policy.Name): $($_.Exception.Message -replace '\s+', ' ')" }; continue }
            foreach ($one in $chunk) {
                if ($missed -contains $one) { $result.Failed[$one] = "not removed from the case hold policy $($policy.Name)"; continue }
                [void]$policy.Locations.Remove($one); $policy.Count = [Math]::Max(0, $policy.Count - 1)
            }
        }
    }
    return $result
}

function Get-Pra2KioskAssignment {
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

function Get-Pra2SharedGrant {
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

function Grant-Pra2SharedPermission {
    <#
    .SYNOPSIS
        Grants one permission on a cloud shared mailbox if it is not there yet, then reads it back.
    .OUTPUTS
        'Granted', 'Present' or throws.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$Mailbox, [Parameter(Mandatory)][object]$Grant)
    $present = {
        switch ($Grant.Right) {
            'FullAccess' { [bool](@(Get-MailboxPermission -Identity $Mailbox -User $Grant.Trustee -ErrorAction SilentlyContinue | Where-Object { -not $_.Deny -and @($_.AccessRights | ForEach-Object { [string]$_ }) -contains 'FullAccess' }).Count) }
            'SendAs' { [bool](@(Get-RecipientPermission -Identity $Mailbox -Trustee $Grant.Trustee -ErrorAction SilentlyContinue | Where-Object { @($_.AccessRights | ForEach-Object { [string]$_ }) -contains 'SendAs' }).Count) }
            'SendOnBehalf' {
                $delegate = Get-Recipient -Identity $Grant.Trustee -ErrorAction SilentlyContinue
                $current = @((Get-Mailbox -Identity $Mailbox).GrantSendOnBehalfTo | ForEach-Object { [string]$_ })
                [bool]($delegate -and ($current -contains [string]$delegate.Name -or $current -contains [string]$delegate.Identity -or $current -contains [string]$delegate.ExternalDirectoryObjectId -or $current -contains [string]$delegate.Guid))
            }
        }
    }
    if (& $present) { return 'Present' }
    if (-not $PSCmdlet.ShouldProcess($Mailbox, "$($Grant.Right) to $($Grant.Trustee)")) { return 'Planned' }
    switch ($Grant.Right) {
        'FullAccess' { $null = Add-MailboxPermission -Identity $Mailbox -User $Grant.Trustee -AccessRights FullAccess -InheritanceType All -AutoMapping ([bool]$Grant.AutoMapping) -ErrorAction Stop }
        'SendAs' { $null = Add-RecipientPermission -Identity $Mailbox -Trustee $Grant.Trustee -AccessRights SendAs -Confirm:$false -ErrorAction Stop }
        'SendOnBehalf' { Set-Mailbox -Identity $Mailbox -GrantSendOnBehalfTo @{ Add = $Grant.Trustee } -ErrorAction Stop }
    }
    $check = Wait-Pra2Condition -Test $present -TimeoutSeconds 120 -IntervalSeconds 10
    if (-not $check.Ok) { throw "$($Grant.Right) to $($Grant.Trustee) not visible after 2 min." }
    return 'Granted'
}

function Invoke-Pra2EntraConnect {
    <#
    .SYNOPSIS
        Entra Connect operation on the rebuilt server: Pause or Resume the scheduler, or run a Delta cycle and wait.
    .DESCRIPTION
        EntraConnect.Mode:
          Remoting  Invoke-Command to EntraConnect.Server (ADSync cmdlets), with the current account.
          Script    EntraConnect.ScriptPath -Operation <op>: any way to reach the server (exit 0 = done).
          Manual    The operator does it; the tool asks for a confirmation (in the console, or in the window when the
                    run belongs to it; impossible with -Force in the console).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][ValidateSet('Pause','Resume','Delta')][string]$Operation, [object]$Caller)
    $entra = $Context.Config.EntraConnect
    if (-not $PSCmdlet.ShouldProcess("Entra Connect ($($entra.Mode))", $Operation)) { return 'Planned' }
    switch ($entra.Mode) {
        'Remoting' {
            $output = Invoke-Command -ComputerName $entra.Server -ErrorAction Stop -ScriptBlock {
                Import-Module ADSync -ErrorAction Stop
                switch ($using:Operation) {
                    'Pause' { Set-ADSyncScheduler -SyncCycleEnabled $false }
                    'Resume' { Set-ADSyncScheduler -SyncCycleEnabled $true }
                    'Delta' {
                        $t0 = Get-Date; while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 15) { Start-Sleep -Seconds 10 }
                        Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                        Start-Sleep -Seconds 10
                        $t0 = Get-Date; while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 30) { Start-Sleep -Seconds 10 }
                    }
                }
                $s = Get-ADSyncScheduler
                'SyncCycleEnabled={0} InProgress={1}' -f $s.SyncCycleEnabled, $s.SyncCycleInProgress
            }
            return (@($output) -join ' ')
        }
        'Script' {
            # A script that ends without 'exit' leaves $LASTEXITCODE unset (StrictMode) or at the value of an earlier call.
            $global:LASTEXITCODE = 0
            $output = & $entra.ScriptPath -Operation $Operation
            $code = Get-Variable -Name LASTEXITCODE -Scope Global -ValueOnly -ErrorAction SilentlyContinue
            if ($code) { throw "EntraConnect.ScriptPath -Operation $Operation ended with exit code $code." }
            return (@($output | ForEach-Object { [string]$_ -split '\r?\n' } | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ' | ')
        }
        'Manual' {
            $text = @{ Pause = 'Pause the Entra Connect scheduler (Set-ADSyncScheduler -SyncCycleEnabled $false)'; Resume = 'Resume the Entra Connect scheduler (Set-ADSyncScheduler -SyncCycleEnabled $true)'; Delta = 'Run a delta synchronisation (Start-ADSyncSyncCycle -PolicyType Delta) and wait for its end' }[$Operation]
            # The window answers through its event file (PRA_EVENT_FILE); the console through the caller.
            if (-not $Caller -and -not $env:PRA_EVENT_FILE) { throw "EntraConnect.Mode = Manual needs the operator: $text." }
            $where = if ($entra.Server) { $entra.Server } else { 'the Entra Connect server' }
            if (-not (Request-PraOperator -Question "$text on $where. Done?" -Title 'PRA Cloud Mailbox - Entra Connect' -Caller $Caller)) { throw "Entra Connect operation $Operation not confirmed by the operator." }
            return 'confirmed by the operator'
        }
    }
}
#endregion

Export-ModuleMember -Function Connect-Pra2Cloud, Disconnect-Pra2Cloud, Invoke-Pra2Graph, Get-Pra2GraphRole, Get-Pra2TenantFact, Get-Pra2SkuExchangePlan,
    Test-Pra2NativeResource, Get-Pra2ResourceIntent, Get-Pra2ResourceProtectionMode, Assert-Pra2UnprotectedResource, Assert-Pra2ResourceOrganizationUnheld, Test-Pra2UnprotectedResourceDeleted, Get-Pra2UnprotectedRecoverCollisions, Get-Pra2UnprotectedDeletionProgress, Assert-Pra2UnprotectedDeletionRepeatable, Get-Pra2ResourceAddresses, Get-Pra2ResourceInventory, Get-Pra2ResourceCollisions,
    Get-Pra2OwnedResource, Read-Pra2ResourceSettings, Test-Pra2ResourceConfigurationReadiness, Get-Pra2ResourceGroup, Test-Pra2ResourceLicence, Set-Pra2ResourceProvisioningDisabled, Set-Pra2ResourceConfiguration, Enable-Pra2ResourceBooking, Set-Pra2ResourceRoomListMembership, Get-Pra2ResourceRoute,
    Get-Pra2ResourceKeys, ConvertTo-Pra2RoomListRecord, Get-Pra2ResourceCloudRoomList, Test-Pra2RoomListCreationReadiness, Set-Pra2RoomListConfiguration,
    Get-Pra2MailboxPlanId, Get-Pra2FreeUnit, Test-Pra2TenantReadiness, Get-Pra2CloudState, Get-Pra2Readiness, Get-Pra2TrusteeFinding,
    Get-Pra2CloudStateSet, Get-Pra2ExoStateSet, Get-Pra2SyncBehaviorSet, Invoke-Pra2GraphBatch, Get-Pra2GraphUserSet, ConvertFrom-Pra2MailboxLocation, Test-Pra2NetworkFailure,
    Wait-Pra2Condition, Get-Pra2ExoState, Set-Pra2SourceOfAuthority, Get-Pra2SourceOfAuthority, Set-Pra2UsageLocation, Set-Pra2GroupMember,
    Test-Pra2GroupMember, Get-Pra2DisabledPlanForExchangeOnly, Set-Pra2Licence, Get-Pra2UserLicence, Get-Pra2DirectLicence, Enable-Pra2LicencePlan,
    Restore-Pra2Licence, Remove-Pra2Identity, Test-Pra2DeletedIdentity,
    Connect-Pra2Compliance, Get-Pra2HoldPolicySet, Test-Pra2ToolHold, Add-Pra2HoldLocationSet, Remove-Pra2HoldLocationSet, Get-Pra2KioskAssignment, Get-Pra2SharedGrant, Grant-Pra2SharedPermission, Invoke-Pra2EntraConnect

# SIG # Begin signature block
# MIIghQYJKoZIhvcNAQcCoIIgdjCCIHICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCfHmx+aOr2SrZg
# a/qtVDn7+BdQ7y2Nae4suVS5lgDsqaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYISbDCCEmgCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC3lNop+1pi
# 7+77XxZheLEjuq4lX58QZtM0Le+gymdN1DANBgkqhkiG9w0BAQEFAASCAgB6TD9I
# WWdLkdQsAIkxqZP5y8xLMNq3IVd2/l3AhEMonefoJOuyl9m4zqIrw0v/14Tq2saW
# 4gSFsqhPrk9+dVXDlZ7nLsTC59ui5KdZpbK/T24kJ3P9JZ4O5kbrXh1IMVtG23S+
# sxkCprLPFJBLpKju/QAMN75gwkBI/yBIiaZBOQY38CgbK0c5tkAV8xwcBhCUeAkm
# eOHMT19Nrxw2yUmteh0g25V0yfOmojA5SUt0JTkcCMZmPof4j57fYPGZSe0Ep5ss
# p41DT09Jmx6ZNB9UwZ8ecCJMuulzGs31Scfhef596A6EcC/JIVtOZQF2OOQD4kQp
# 7GVOyQe8m/hiyJ75ZtRLLa8oAGdv8hfd3LtT1v0Zcy497qHXpBgF0/3XJ2X106Dm
# lKXWDwOPVFGFfhV6MA9tea+T878bdtQDwPfkHyNL8yKidOakurZEnhKbBHnE1T/w
# VkO9K1j6ZpJ79XXATxC0tDyjIwefo8w47oC959olzDBD56R1sqkTM4xPH2Pl48N8
# ZefVcP0mw/VYmyCrUbKxoXb60VDGAgNhxkHNfUPSemeOUXvTi59xctARGiHkr6tx
# DZwGRyJfZ4cDk0lWALPOVtZUfRgBOYd+cU/VcpdnN8vbDSInl9LF9NrD3Gu23g7Z
# f1uIlb6ASfy/Ok3roqSV5OxcZkS2ovc66/btzaGCDykwgg8lBgorBgEEAYI3AwMB
# MYIPFTCCDxEGCSqGSIb3DQEHAqCCDwIwgg7+AgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIJ9GsqbkDUlFvRZOr7zQ12z8pxDu0g7nx+XIxy2J4i8lAggsIZVoQkLc
# NBgPMjAyNjEwMTQxNTIzNDlaMAMCAQECEF9fo06LFxXvrPcvFQwqrV2gggwAMIIE
# /DCCAuSgAwIBAgIQH2sWYtIuG2xd8cDBoGAOODANBgkqhkiG9w0BAQsFADBzMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS8wLQYDVQQDDCZTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSBSMTAeFw0yNTAyMTgxNjMyMDJaFw0zNDExMTIxODUwMDVa
# MG4xCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3Rv
# bjERMA8GA1UECgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1w
# aW5nIFVuaXQgMjAyNSBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABBv7UVHH
# b+ZVluxPXlfE3M6tg0Xnq8dic+O3vPOCRpalUM1vO9A+GRzSVVjyygHhYrBw62XL
# Fh1kv7e+yRd/aTajggFaMIIBVjAfBgNVHSMEGDAWgBQMnRAljpqnG5mHQ88IfuG9
# gZD0zzBRBggrBgEFBQcBAQRFMEMwQQYIKwYBBQUHMAKGNWh0dHA6Ly9jZXJ0LnNz
# bC5jb20vU1NMLmNvbS10aW1lU3RhbXBpbmctSS1SU0EtUjEuY2VyMFEGA1UdIARK
# MEgwPAYMKwYBBAGCqTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNz
# bC5jb20vcmVwb3NpdG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwRgYDVR0fBD8wPTA7oDmgN4Y1aHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29t
# LXRpbWVTdGFtcGluZy1JLVJTQS1SMS5jcmwwHQYDVR0OBBYEFM582cAEgMUkEGoJ
# 6hyrJT0R/ajSMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAgHN1
# JIW9ZlNip07FfY9r1LJwLKTl1h/y9r1BjDgHMEz7U7J7mAy32JxA1KLyjCCJqNku
# RfhDHGFEfZtjTOru7AQMDSDiUDvCCc3jaz5py9jGvwIe5npkLbXqXhA+SqCzG719
# USigwhHkYr7v8NiXoaY0Xu03p0BiZZreZlpngmGB9+86N94FqIFx2I5Z5iuu7d4j
# 7aAOZBNxfU9jFDIIGDtWfWeCpF0Q8rLE/3DHDlzJKs5Il3i/VoOl7rTU938oSrxM
# www6GDBlKX8rGHxvkjgIb9MXkVglJSDHsNAZDeCS/inaxKHA+ChuB8K4OxIxa6k1
# e+eHBypakAY3wuN7w7PlPmhN6k5IdqZn3HZHs2VjoR841Z9wmVPFdGDdkBZ7XOmb
# 5OZUaKufsvtER0VbY2DkzaqAaItt1kc+I+FUz7PiVz2PUzpegkftRRxvrPyIL08b
# lJTnbMQ0XrqDN0rEfuTLv4XhhnKyJaXhxCUdjO/cRumMzZ35QTUMItBWy1xGp0iy
# oFVnAya2pPsVMuAI+sy1zxlOS9l5iSxvKJ8gpnpOgjqYTa7u1eyZo+4JlcGoiiR1
# 7LthMTrF1q62tO5xOpI8txUZ6gtKGtkJV6wdj+vZLcwKZ2M93GqVpOo6UBlGjx1s
# KsVZ7AIUzJZli4NJ0M2KYcrfNnlebkYKpDnSqWwwggb8MIIE5KADAgECAhBtUhhw
# h+gjTYVgANCAj5NWMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTE5MTExMzE4NTAwNVoXDTM0MTExMjE4NTAwNVowczELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCu
# URAT0vk8IKAghd7JUBxkyeH9xek0/wp/MUjoclrFXqhh/fGH91Fc+7fm0MHCE7A+
# wmOiqBj9ODrJAYGq3rm33jCnHSsCBNWAQYyoauLq8IjqsS1JlXL29qDNMMdwZ8UN
# zQS7vWZMDJ40JSGNphMGTIA2qn2bohGtgRc4p1395ESypUOaGvJ3t0FNL3BuKmb6
# YctMcQUF2sqooMzd89h0E6ujdvBDo6ZwNnWoxj7YmfWjSXg33A5GuY9ym4QZM5OE
# Vgo8ebz/B+gyhyCLNNhh4Mb/4xvCTCMVmNYrBviGgdPZYrym8Zb84TQCmSuX0JlL
# La6WK1aO6qlwISbb9bVGh866ekKblC/XRP20gAu1CjvcYciUgNTrGFg8f8AJgQPO
# Cc1/CCdaJSYwhJpSdheKOnQgESgNmYZPhFOC6IKaMAUXk5U1tjTcFCgFvvArXtK4
# azAWUOO1Y3fdldIBL6LjkzLUCYJNkFXqhsBVcPMuB0nUDWvLJfPimstjJ8lF4S6E
# CxWnlWi7OElVwTnt1GtRqeY9ydvvGLntU+FecK7DbqHDUd366UreMkSBtzevAc9a
# qoZPnjVMjvFqV1pYOjzmTiVHZtAc80bAfFe5LLfJzPI6DntNyqobpwTevQpHqPDN
# 9qqNO83r3kaw8A9j+HZiSw2AX5cGdQP0kG0vhzfgBwIDAQABo4IBgTCCAX0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTCBgwYIKwYBBQUHAQEEdzB1MFEGCCsGAQUFBzAChkVodHRwOi8vd3d3LnNzbC5j
# b20vcmVwb3NpdG9yeS9TU0xjb21Sb290Q2VydGlmaWNhdGlvbkF1dGhvcml0eVJT
# QS5jcnQwIAYIKwYBBQUHMAGGFGh0dHA6Ly9vY3Nwcy5zc2wuY29tMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwOwYDVR0fBDQwMjAwoC6gLIYq
# aHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0GA1Ud
# DgQWBBQMnRAljpqnG5mHQ88IfuG9gZD0zzAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAJIZdQ2mWkLPGQfZ8vyU+sCb8BXpRJZaL3Ez3VDlE3uZk3cP
# xPtybVfLuqaci0W6SB22JTMttCiQMnIVOsXWnIuAbD/aFTcUkTLBI3xys+wEajzX
# aXJYWACDS47BRjDtYlDW14gLJxf8W6DQoH3jHDGGy8kGJFOlDKG7/YrK7UGfHtBA
# EDVe6lyZ+FtCsrk7dD/IiL/+Q3Q6SFASJLQ2XI89ihFugdYL77CiDNXrI2MFspQG
# swXEAGpHuaQDTHUp/LdR3TyrIsLlnzoLskUGswF/KF8+kpWUiKJNC4rPWtNrxlbX
# YRGgdEdx8SMjUTDClldcrknlFxbqHsVmr9xkT2QtFmG+dEq1v5fsIK0vHaHrWjMM
# maJ9i+4qGJSD0stYfQ6v0PddT7EpGxGd867Ada6FZyHwbuQSadMb0K0P0OC2r7rw
# qBUe0BaMqTa6LWzWItgBjGcObXeMxmbQqlEz2YtAcErkZvh0WABDDE4U8GyV/32F
# daAvJgTfe9MiL2nSBioYe/g5mHUSWAay/Ip1RQmQCvmF9sNfqlhJwkjy/1U1ibUk
# TIUBX3HgymyQvqQTZLLys6pL2tCdWcjI9YuLw30rgZm8+K387L7ycUvqrmQ3ZJlu
# jHl3r1hgV76s3WwMPgKk1bAEFMj+rRXimSC+Ev30hXZdqyMdl/il5Ksd0vhGMYIC
# WDCCAlQCAQEwgYcwczELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgUjECEB9rFmLSLhtsXfHAwaBg
# DjgwCwYJYIZIAWUDBAIBoIIBYTAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# HAYJKoZIhvcNAQkFMQ8XDTI2MTAxNDE1MjM0OVowKAYJKoZIhvcNAQk0MRswGTAL
# BglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIIrcedoXLvjh
# xlt+8jGAL7Qm9XcnKpO1TtTaozD+D63/MIHJBgsqhkiG9w0BCRACLzGBuTCBtjCB
# szCBsAQgVCr5oWqNci5mEUl4iumUwYqaruWmXLNEolSa+Wx5x4swgYswd6R1MHMx
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjER
# MA8GA1UECgwIU1NMIENvcnAxLzAtBgNVBAMMJlNTTC5jb20gVGltZXN0YW1waW5n
# IElzc3VpbmcgUlNBIENBIFIxAhAfaxZi0i4bbF3xwMGgYA44MAoGCCqGSM49BAMC
# BEcwRQIhAMyQy9QyT+Z15g5KPJZiSX7GkPDVP/7K6bX0ojX/ngvtAiBYdlXEZqCh
# 8nspPj+/Zc5nwLZ73ZZkqsCfC1uvG4sZxw==
# SIG # End signature block
