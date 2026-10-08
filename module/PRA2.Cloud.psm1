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
    (Security & Compliance) and the Entra Connect operations of Recover.

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
    Version : 1.0.0
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
    Connect-MgGraph -ClientId $cloud.AppId -TenantId $cloud.TenantId -CertificateThumbprint $cloud.CertificateThumbprint -NoWelcome -ErrorAction Stop
    $Context['GraphConnected'] = $true
    Import-Module ExchangeOnlineManagement -MinimumVersion 3.10.0 -ErrorAction Stop
    Connect-ExchangeOnline -AppId $cloud.AppId -CertificateThumbprint $cloud.CertificateThumbprint -Organization $cloud.Organization -ShowBanner:$false `
        -CommandName Get-Recipient, Get-MailUser, Get-Mailbox, Get-MailboxLocation, Get-User, Get-OrganizationConfig, Set-MailUser, Set-Mailbox,
        Add-MailboxPermission, Get-MailboxPermission, Add-RecipientPermission, Get-RecipientPermission -ErrorAction Stop
    $Context['ExoConnected'] = $true
    return ('Microsoft Graph and Exchange Online, app {0}, certificate {1} (valid until {2})' -f $cloud.AppId, $cloud.CertificateThumbprint.Substring(0, 8), $certificate.NotAfter.ToString('yyyy-MM-dd'))
}

function Disconnect-Pra2Cloud {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($Context.ContainsKey('ExoConnected') -and $Context['ExoConnected']) { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop } catch { $null = $_ }; $Context['ExoConnected'] = $false }
    if ($Context.ContainsKey('GraphConnected') -and $Context['GraphConnected']) { try { $null = Disconnect-MgGraph -ErrorAction Stop } catch { $null = $_ }; $Context['GraphConnected'] = $false }
}

function Invoke-Pra2Graph {
    <#
    .SYNOPSIS
        Microsoft Graph request with retries on throttling (429) and transient errors (502, 503, 504).
        Returns $null for 404 when -AllowNotFound.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri, [ValidateSet('GET','POST','PATCH','DELETE')][string]$Method = 'GET', [object]$Body, [switch]$AllowNotFound)
    $parameters = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
    if ($null -ne $Body) { $parameters.Body = ($Body | ConvertTo-Json -Depth 10 -Compress); $parameters.ContentType = 'application/json' }
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try { return Invoke-MgGraphRequest @parameters }
        catch {
            $code = 0
            try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 0 }
            if ($code -eq 404 -and $AllowNotFound) { return $null }
            if ($code -in @(429, 502, 503, 504) -and $attempt -lt 5) { Start-Sleep -Seconds (5 * $attempt); continue }
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
    param([Parameter(Mandatory)][hashtable]$Context)
    $skus = @((Invoke-Pra2Graph -Uri 'v1.0/subscribedSkus').value)
    $organization = @((Invoke-Pra2Graph -Uri 'v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime').value) | Select-Object -First 1
    $group = $null; $groupBehavior = $null
    $users = $Context.Config.Licensing.Users
    if ($users.Mode -eq 'Group' -and $users.GroupId) {
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
        [int]$UserCount = 0, [int]$SharedCount = 0)
    $findings = [System.Collections.Generic.List[object]]::new()
    $add = { param($Level, $Code, $Message) [void]$findings.Add([pscustomobject]@{ Level = $Level; Code = $Code; Message = $Message }) }
    $licensing = $Config.Licensing
    $needed = @($script:RequiredGraphRoles)
    if ($licensing.Users.Mode -eq 'Group') { $needed += $script:GroupModeRoles[0] }
    $missing = @($needed | Where-Object { $_.Role -notin $Roles -and -not ($_.Role -eq 'LicenseAssignment.ReadWrite.All' -and 'Directory.ReadWrite.All' -in $Roles) })
    if ($missing.Count) { & $add 'Error' 'APP_PERMISSION' ('Missing Graph application permission(s): ' + (($missing | ForEach-Object { '{0} ({1})' -f $_.Role, $_.Why }) -join '; ')) }
    else { & $add 'Ok' 'APP_PERMISSION' ('Graph application permissions present: ' + (($needed | ForEach-Object { $_.Role }) -join ', ')) }

    $skuById = @{}; $skuByName = @{}
    foreach ($sku in @($Tenant.Skus)) { $skuById[[string]$sku.skuId] = $sku; $skuByName[[string]$sku.skuPartNumber] = $sku }
    switch ($licensing.Users.Mode) {
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
    }
    if ($SharedCount -gt 0) {
        $name = [string]$licensing.Shared.SkuPartNumber
        if (-not $name) { & $add 'Error' 'SHARED_LICENCE' 'Licensing.Shared.SkuPartNumber is empty: Convert needs a temporary licence with an Exchange plan for each shared mailbox (one at a time).' }
        else {
            $sku = $skuByName[$name]
            if (-not $sku) { & $add 'Error' 'SHARED_LICENCE' "Licensing.Shared.SkuPartNumber $name is not a licence of the tenant." }
            elseif (-not @(Get-Pra2SkuExchangePlan -Sku $sku).Count) { & $add 'Error' 'SHARED_LICENCE' "$name has no Exchange plan." }
            else {
                $free = Get-Pra2FreeUnit $sku
                $level = if ($free -ge 1) { 'Ok' } else { 'Error' }
                & $add $level 'SHARED_LICENCE' ('{0}: {1} free unit(s); one is enough, the shared mailboxes are converted one after another ({2} in scope).' -f $name, $free, $SharedCount)
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
        Entra ID object and Exchange Online recipient of one snapshot row (read-only).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Record)
    $select = 'id,userPrincipalName,accountEnabled,onPremisesSyncEnabled,onPremisesImmutableId,usageLocation,assignedLicenses,assignedPlans,licenseAssignmentStates,serviceProvisioningErrors'
    $user = $null; $matchedBy = ''
    if ($Record.immutable_id) {
        $filter = [uri]::EscapeDataString(("onPremisesImmutableId eq '{0}'" -f ([string]$Record.immutable_id).Replace("'", "''")))
        $found = @((Invoke-Pra2Graph -Uri ("v1.0/users?`$filter={0}&`$select={1}" -f $filter, $select)).value)
        if ($found.Count -eq 1) { $user = $found[0]; $matchedBy = 'immutableId' }
    }
    if (-not $user -and $Record.user_principal_name) {
        $user = Invoke-Pra2Graph -Uri ("v1.0/users/{0}?`$select={1}" -f [uri]::EscapeDataString([string]$Record.user_principal_name), $select) -AllowNotFound
        if ($user) { $matchedBy = 'UPN' }
    }
    $behavior = $null; $recipient = $null; $mailUser = $null; $mailbox = $null; $locations = @()
    if ($user) {
        $behavior = Invoke-Pra2Graph -Uri ("v1.0/users/{0}/onPremisesSyncBehavior?`$select=isCloudManaged" -f $user.id) -AllowNotFound
        $recipient = Get-Recipient -Identity ([string]$user.id) -ErrorAction SilentlyContinue
        if ($recipient) {
            if ([string]$recipient.RecipientTypeDetails -eq 'MailUser') { $mailUser = Get-MailUser -Identity ([string]$user.id) -ErrorAction SilentlyContinue }
            elseif ([string]$recipient.RecipientTypeDetails -match 'Mailbox$') { $mailbox = Get-Mailbox -Identity ([string]$user.id) -ErrorAction SilentlyContinue }
        }
        $locations = @(Get-MailboxLocation -User ([string]$user.id) -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | ForEach-Object {
                [pscustomobject]@{ Type = [string]$_.MailboxLocationType; Guid = [string]$_.MailboxGuid } })
    }
    $holds = @()
    if ($mailUser) { $holds = @($mailUser.InPlaceHolds | ForEach-Object { [string]$_ }) }
    elseif ($mailbox) { $holds = @($mailbox.InPlaceHolds | ForEach-Object { [string]$_ }); if ($mailbox.LitigationHoldEnabled) { $holds += 'LitigationHold' } }
    $tagAttribute = [string]$Context.Config.Retention.TagAttribute
    $tagValue = if ($recipient -and $recipient.PSObject.Properties[$tagAttribute]) { [string]$recipient.$tagAttribute } else { '' }
    return [pscustomobject]@{
        MatchedBy = $matchedBy; User = $user; IsCloudManaged = $(if ($behavior) { [bool]$behavior.isCloudManaged } else { $null })
        RecipientType = $(if ($recipient) { [string]$recipient.RecipientTypeDetails } else { '' })
        ExchangeGuid = $(if ($mailUser) { [string]$mailUser.ExchangeGuid } elseif ($recipient) { [string]$recipient.ExchangeGuid } else { '' })
        Holds = $holds; Locations = $locations; TagValue = $tagValue
    }
}

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
    if ([string]$Record.kind -in @('Room', 'Equipment')) {
        & $add 'Error' 'ExchangeOnline' 'KIND_NOT_SUPPORTED' ("{0} mailbox: not converted by this version (booking settings not collected, not validated in the lab). It stays unreachable during the disaster; recreate it in Exchange Online by hand if needed." -f $Record.kind)
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
        -CommandName Get-CaseHoldPolicy, Set-CaseHoldPolicy -ErrorAction Stop
    $Context['ComplianceConnected'] = $true
}

function Get-Pra2HoldTag {
    <# InPlaceHolds value of the configured case hold policy ('UniH' + policy GUID). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Policy)
    $item = Get-CaseHoldPolicy -Identity $Policy -ErrorAction Stop
    return ('UniH{0}' -f $item.Guid)
}

function Set-Pra2HoldLocation {
    <#
    .SYNOPSIS
        Adds or removes an Exchange location of the case hold policy; already there / already gone is not an error.
    .DESCRIPTION
        Set-CaseHoldPolicy can answer "Policy ... failed to be deployed" while the change is recorded (the hold then
        follows when the distribution goes through, lab 7-8 Oct 2026): the policy is read again and, when the location
        list shows the change, the call counts as done (the caller waits for InPlaceHolds anyway).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][string]$Policy, [Parameter(Mandatory)][string]$Identity, [Parameter(Mandatory)][ValidateSet('Add','Remove')][string]$Operation)
    if (-not $PSCmdlet.ShouldProcess($Policy, "$Operation $Identity")) { return }
    try {
        if ($Operation -eq 'Add') { Set-CaseHoldPolicy -Identity $Policy -AddExchangeLocation $Identity -ErrorAction Stop }
        else { Set-CaseHoldPolicy -Identity $Policy -RemoveExchangeLocation $Identity -ErrorAction Stop }
    } catch {
        if ("$_" -match 'already|déjà|not found in|introuvable dans|does not exist in') { return }
        if ("$_" -match 'failed to be deployed|PolicyDeploymentException') {
            $locations = @((Get-CaseHoldPolicy -Identity $Policy -ErrorAction Stop).ExchangeLocation | ForEach-Object { [string]$_.Name; [string]$_.ImmutableIdentity; [string]$_ } | Where-Object { $_ })
            $present = $locations -contains $Identity
            if (($Operation -eq 'Add') -eq $present) {
                Write-Warning "Case hold policy $Policy`: $Operation $Identity recorded, distribution pending ($($_.Exception.Message -replace '\s+', ' '))"
                return
            }
        }
        throw
    }
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
          Manual    The operator does it; the tool asks for a confirmation (impossible with -Force).
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
            if (-not $Caller) { throw "EntraConnect.Mode = Manual needs the operator: $text." }
            $where = if ($entra.Server) { $entra.Server } else { 'the Entra Connect server' }
            if (-not $Caller.ShouldContinue("$text on $where. Done?", 'PRA Cloud Mailbox - Entra Connect')) { throw "Entra Connect operation $Operation not confirmed by the operator." }
            return 'confirmed by the operator'
        }
    }
}
#endregion

Export-ModuleMember -Function Connect-Pra2Cloud, Disconnect-Pra2Cloud, Invoke-Pra2Graph, Get-Pra2GraphRole, Get-Pra2TenantFact, Get-Pra2SkuExchangePlan,
    Get-Pra2MailboxPlanId, Get-Pra2FreeUnit, Test-Pra2TenantReadiness, Get-Pra2CloudState, Get-Pra2Readiness, Get-Pra2TrusteeFinding,
    Wait-Pra2Condition, Get-Pra2ExoState, Set-Pra2SourceOfAuthority, Get-Pra2SourceOfAuthority, Set-Pra2UsageLocation, Set-Pra2GroupMember,
    Test-Pra2GroupMember, Get-Pra2DisabledPlanForExchangeOnly, Set-Pra2Licence, Get-Pra2UserLicence, Get-Pra2DirectLicence, Enable-Pra2LicencePlan,
    Restore-Pra2Licence, Remove-Pra2Identity, Test-Pra2DeletedIdentity,
    Connect-Pra2Compliance, Get-Pra2HoldTag, Set-Pra2HoldLocation, Get-Pra2KioskAssignment, Get-Pra2SharedGrant, Grant-Pra2SharedPermission, Invoke-Pra2EntraConnect
