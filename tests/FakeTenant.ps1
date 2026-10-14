<#
.SYNOPSIS
    In-memory Microsoft 365 tenant for the end-to-end tests of Convert and Recover (PowerShell 7).
.DESCRIPTION
    Imitates what the lab showed (developer guide, chapter 3), not the whole service: Microsoft Graph (users, licence group,
    licences, source of authority, recycle bin), Exchange Online (recipients, mailbox provisioning, Teams storage,
    holds, permissions, inactive mailboxes), the eDiscovery case hold and Entra Connect.
    What takes time in Microsoft 365 happens on the next tick (Step-PraFake): the tests mock Start-Sleep and the
    wait function so that every poll advances the tenant by one tick.
    Faults: LicenceAssign400 (number of refused assignments), NeverMailbox (user IDs never provisioned),
    NoHoldStamp (the case hold is never stamped), DeleteFails (DELETE /users refused), NoRecreate (Entra Connect
    does not recreate deleted objects), SendAsFails (Add-RecipientPermission refused),
    CaseHoldDeployError (Set-CaseHoldPolicy records the change but reports 'failed to be deployed', lab 7-8 Oct),
    CaseHoldFails (Set-CaseHoldPolicy refuses every change),
    StopAfter (@{ Pattern; Path }: the stop file of the window is created after the first call that matches),
    BatchThrottle (number of requests of a Graph batch answered 429 before the next ones succeed), ExoFilterFails (the
    filtered Exchange Online reads fail: the tool reads one object at a time). PageSize: users per page of /users.
.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
if (-not ('Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' -as [type])) {
    Add-Type -TypeDefinition @'
namespace Microsoft.Exchange.Configuration.Tasks {
    public class ManagementObjectNotFoundException : System.Exception {
        public ManagementObjectNotFoundException(string message) : base(message) {}
    }
}
'@
}

function global:New-PraFakeTenant {
    Reset-PraFakeGraphTransport
    $global:PraFake = @{
        Tick = 0
        Skus = [ordered]@{
            'sku-e5' = @{ Part = 'E5'; Enabled = 10; Plans = @(
                    @{ Name = 'EXCHANGE_S_ENTERPRISE'; Id = 'plan-exo'; AppliesTo = 'User' }
                    @{ Name = 'TEAMS1'; Id = 'plan-teams'; AppliesTo = 'User' }
                    @{ Name = 'COMPANY_ADMIN'; Id = 'plan-company'; AppliesTo = 'Company' }) }
            'sku-teams' = @{ Part = 'TEAMS_ENTERPRISE'; Enabled = 10; Plans = @(
                    @{ Name = 'TEAMS1'; Id = 'plan-teams2'; AppliesTo = 'User' }
                    @{ Name = 'EXCHANGE_S_DESKLESS'; Id = 'plan-kiosk'; AppliesTo = 'User' }
                    @{ Name = 'EXCHANGE_S_FOUNDATION'; Id = 'plan-foundation'; AppliesTo = 'User' }) }
        }
        Group = @{ Id = '11111111-1111-1111-1111-111111111111'; Name = 'LIC-EXO'; Synced = $false; CloudManaged = $false
            Licences = @(@{ skuId = 'sku-e5'; disabledPlans = @('plan-teams') }); Members = [System.Collections.Generic.List[string]]::new() }
        Users = [ordered]@{}
        Contacts = [ordered]@{}
        RoomLists = [ordered]@{}
        Recycle = [ordered]@{}
        Ad = [System.Collections.Generic.List[hashtable]]::new()
        Inactive = [System.Collections.Generic.List[hashtable]]::new()
        # Case hold policies of the eDiscovery case; CasePolicy is the first one (the configured policy).
        CasePolicies = [System.Collections.Generic.List[hashtable]]::new()
        CasePolicy = $null
        CaseRules = [System.Collections.Generic.List[string]]::new()
        Scheduler = $true
        Faults = @{ LicenceAssign400 = 0; NeverMailbox = @(); NoHoldStamp = $false; DeleteFails = $false; DeleteRejectedStatus = 0; DeleteAmbiguousStatus = 0; NoRecreate = $false; SendAsFails = $false; CaseHoldDeployError = $false; CaseHoldFails = $false; StopAfter = $null; BatchThrottle = 0; ExoFilterFails = $false
            NewMailboxFails = $false; CrashAfterNewMailbox = $false; CrashAfterDelete = $false; CrashAfterContact = $false; NoEnterprise = $false
            RecipientReadFails = $false; RecipientQueryFailsAt = 0; RecipientIdentityError = $null; RecipientIdentityEmpty = $false; RejectUnfilteredRecipientInventory = $false; GraphReadFails = $false
            RejectUnfilteredGraphInventory = $false; GraphQueryFailsAt = 0; GraphQueryUnsupported = $false; GraphAdvancedStatus = 0; GraphAdvancedFailsAt = 0; SignInDisableFails = $false; ContactFails = $false; ContactSoftDeletedAliasBlocks = $false; ContactSoftDeletedProxyBlocks = $false; NoInactive = $false; WrongInactiveHold = $false; WorkDaysAliasRead = $false; WorkTimeZoneObjectRead = $false
            DirectoryWriteFails = $false; DirectoryReadMismatch = $false; DirectoryNotesReadMismatch = $false; DirectoryNotesReadMissing = $false
            RoomListReadFails = $false; RoomListAddFails = $false; RoomListNoMembership = $false; CrashAfterRoomListAdd = $false
            CalendarEnableFails = $false; CalendarEnabledReadMismatch = $false; BookingDisableFails = $false; HiddenFinalFails = $false
            NewRoomListFails = $false; CrashAfterNewRoomList = $false; RoomListConfigFails = $false; CrashAfterRoomListConfig = $false; RoomListConfigMismatch = $false
            RoomListExchangeOnly = $false; RoomListReadAfterCreateFails = $false; ResourceReadAfterCreateFails = $false; ContactReadAfterCreateFails = $false
            SignInPropagationReads = 0; SignInReadFault = ''
            NoLicenseReads = $false; NoSoftDeleted = $false; SoftDeletedAddressBlocks = $false; OrganizationReadFails = $false }
        PageSize = 999; FilterCalls = 0; CaseHoldCalls = 0
        OrgHolds = @()
        ResourceSchema = @('Room/Video','Equipment/Video')
        Calls = [System.Collections.Generic.List[string]]::new()
        Reads = [System.Collections.Generic.List[string]]::new()
        RecipientQueries = [System.Collections.Generic.List[object]]::new()
        GraphQueries = [System.Collections.Generic.List[object]]::new()
        UserWrites = [System.Collections.Generic.List[object]]::new()
    }
    $global:PraFake.CasePolicy = @{ Name = 'PRA-HOLD'; Guid = '22222222-2222-2222-2222-222222222222'; CaseId = 'case-pra'; Locations = [System.Collections.Generic.List[string]]::new() }
    $global:PraFake.CasePolicies.Add($global:PraFake.CasePolicy)
}

function global:Get-PraFakeHoldTag {
    <# InPlaceHolds tag of a case hold policy (the configured one by default). #>
    param([string]$Name)
    $policy = if ($Name) { $global:PraFake.CasePolicies | Where-Object { $_.Name -eq $Name } | Select-Object -First 1 } else { $global:PraFake.CasePolicy }
    'UniH' + $policy.Guid
}

function global:New-PraFakeResourceSettings {
    $calendar = [ordered]@{ AutomateProcessing = 'AutoAccept'; AllowConflicts = $false; BookingWindowInDays = 180; MaximumDurationInMinutes = 120
        AllowRecurringMeetings = $true; EnforceSchedulingHorizon = $true; ScheduleOnlyDuringWorkHours = $true; ConflictPercentageAllowed = 0; MaximumConflictInstances = 0
        ForwardRequestsToDelegates = $true; DeleteAttachments = $true; DeleteComments = $false; RemovePrivateProperty = $true; DeleteSubject = $false
        AddOrganizerToSubject = $true; DeleteNonCalendarItems = $true; TentativePendingApproval = $true; EnableResponseDetails = $true; OrganizerInfo = $true
        AllBookInPolicy = $false; AllRequestInPolicy = $false; AllRequestOutOfPolicy = $false; AddAdditionalResponse = $true; AdditionalResponse = 'Contact reception.'
        ProcessExternalMeetingMessages = $false; RemoveOldMeetingMessages = $true; ResourceDelegates = @('alice@contoso.com'); BookInPolicy = @('bob@contoso.com')
        RequestInPolicy = @('carol@contoso.com'); RequestOutOfPolicy = @() }
    return [ordered]@{ SchemaVersion = 2; Capacity = 12; ResourceCustom = @('Video'); Calendar = $calendar
        Directory = @{ Office = 'Building A / room 101'; StreetAddress = '1 Test Street'; City = 'Paris'; StateOrProvince = ''; PostalCode = '75001'; CountryOrRegion = 'France'; Phone = '+33100000000'; Notes = 'Reception resource' }
        RoomLists = @()
        Delivery = @{ RequireSenderAuthenticationEnabled = $true; ModerationEnabled = $false; SendModerationNotifications = 'Never'
            AcceptMessagesOnlyFromSendersOrMembers = @('bob@contoso.com'); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @() }
        CalendarConfiguration = @{ WorkingHoursStartTime = '08:00:00'; WorkingHoursEndTime = '18:00:00'; WorkingHoursDays = @('monday','tuesday','wednesday','thursday','friday'); WorkingHoursTimeZone = 'Romance Standard Time' } }
}

function global:Add-PraFakeUser {
    param([Parameter(Mandatory)][string]$Upn, [ValidateSet('User','Shared','Room','Equipment')][string]$Kind = 'User', [string]$Id = ([guid]::NewGuid().ToString()),
        [string]$OnPremGuid = ([guid]::NewGuid().ToString()), [string]$ImmutableId = ('imm-' + $Upn), [string]$Component, [string]$UsageLocation = 'FR')
    $user = @{ Id = $Id; Upn = $Upn; Kind = $Kind; ImmutableId = $ImmutableId; OnPremGuid = $OnPremGuid; Synced = $true; CloudManaged = $false
        UsageLocation = $UsageLocation; AccountEnabled = $true; Name = ($Upn -split '@')[0]; Alias = ($Upn -split '@')[0]; Mail = $Upn; Addresses = @("SMTP:$Upn"); Marker = ''; ResourceCapacity = 0
        ResourceCustom = @(); Hidden = $false; Calendar = @{}; WorkHours = @{}
        Directory = @{ Office = ''; StreetAddress = ''; City = ''; StateOrProvince = ''; PostalCode = ''; CountryOrRegion = ''; Phone = ''; Notes = '' }
        Delivery = @{ RequireSenderAuthenticationEnabled = $false; ModerationEnabled = $false; SendModerationNotifications = 'Never'
            AcceptMessagesOnlyFromSendersOrMembers = @(); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @() }
        Direct = [System.Collections.Generic.List[hashtable]]::new()
        Type = 'MailUser'; ExchangeGuid = $OnPremGuid; DirSynced = $true; Holds = [System.Collections.Generic.List[string]]::new(); Tag = ''
        Component = $(if ($Component) { $Component } else { $null }); PendingShared = $false
        FullAccess = [System.Collections.Generic.List[string]]::new(); AutoMap = @{}; SendAs = [System.Collections.Generic.List[string]]::new()
        SendOnBehalf = [System.Collections.Generic.List[string]]::new() }
    $global:PraFake.Users[$Id] = $user
    $global:PraFake.Ad.Add(@{ ImmutableId = $ImmutableId; Upn = $Upn; Kind = $Kind; OnPremGuid = $OnPremGuid; ObjectGuid = [guid]::NewGuid().ToString() })
    return $user
}

function global:Write-PraFakeCall {
    param([string]$Text)
    $global:PraFake.Calls.Add(('{0:000} {1}' -f $global:PraFake.Tick, $Text))
    # The window asks for a stop at a given moment: the stop file appears after the first call that matches.
    $stop = $global:PraFake.Faults.StopAfter
    if ($stop -and $Text -match $stop.Pattern -and -not (Test-Path -LiteralPath $stop.Path)) { Set-Content -LiteralPath $stop.Path -Value 'stop' }
}

function global:Resolve-PraFakeUser {
    param([string]$Identity)
    if (-not $Identity) { return $null }
    if ($global:PraFake.Users.Contains($Identity)) { return $global:PraFake.Users[$Identity] }
    foreach ($user in $global:PraFake.Users.Values) { if ($user.Upn -eq $Identity -or $user.Name -eq $Identity -or $user.Alias -eq $Identity -or $user.ExchangeGuid -eq $Identity -or @($user.Addresses) -contains $Identity -or @($user.Addresses) -contains "smtp:$Identity") { return $user } }
    return $null
}

function global:Get-PraFakeMailboxPlan {
    <# Mailbox plans (Exchange Online or Kiosk; never Foundation) enabled for a user, direct or through the licence group. #>
    param([hashtable]$User)
    $fake = $global:PraFake
    $licences = @($User.Direct) + @(if ($fake.Group.Members.Contains($User.Id)) { $fake.Group.Licences })
    foreach ($licence in $licences) {
        foreach ($plan in @($fake.Skus[$licence.skuId].Plans | Where-Object { $_.Id -in @('plan-exo', 'plan-kiosk') })) { if (@($licence.disabledPlans) -notcontains $plan.Id) { $plan.Id } }
    }
}

function global:Test-PraFakeExchangePlan { param([hashtable]$User) return [bool]@(Get-PraFakeMailboxPlan $User).Count }

function global:Step-PraFake {
    <# One tick: what Microsoft 365 does in the background between two polls. #>
    $fake = $global:PraFake
    $fake.Tick++
    foreach ($user in @($fake.Users.Values)) {
        if ($user.CloudManaged -and $user.DirSynced) { $user.DirSynced = $false }
        $plan = Test-PraFakeExchangePlan $user
        if ($user.Type -eq 'MailUser' -and $user.ExchangeGuid -eq [guid]::Empty.ToString() -and -not $user.DirSynced -and $plan -and $user.UsageLocation -and $fake.Faults.NeverMailbox -notcontains $user.Id) {
            # The Teams storage (ComponentShared) becomes the primary mailbox when there is one (lab test A).
            $user.Type = 'UserMailbox'; $user.ExchangeGuid = $(if ($user.Component) { $user.Component } else { [guid]::NewGuid().ToString() }); $user.Component = $null
            Write-PraFakeCall "Mailbox $($user.Upn) $($user.ExchangeGuid)"
        } elseif ($user.Type -eq 'UserMailbox' -and -not $plan -and $user.DirSynced -and $user.Holds.Count) {
            # Exchange cannot disable a mailbox on hold: the switch back is blocked (lab T8, H).
            if (-not $user.ContainsKey('Blocked')) { $user['Blocked'] = $true; Write-PraFakeCall "BlockedByHold $($user.Upn)" }
        } elseif ($user.Type -eq 'UserMailbox' -and -not $plan -and $user.DirSynced) {
            # Plan removed after the source of authority went back to AD: MailUser with the on-premises GUID,
            # the cloud mailbox stays as ComponentShared with its content (lab T8, T13).
            $user.Component = $user.ExchangeGuid; $user.ExchangeGuid = $user.OnPremGuid; $user.Type = 'MailUser'
            Write-PraFakeCall "BackOnPremises $($user.Upn)"
        } elseif ($user.Type -eq 'UserMailbox' -and -not $plan) {
            # Licence removed before the conversion to shared is effective: disabled at once (lab SH2).
            $user.Type = 'DisabledUser'; $user.PendingShared = $false
            Write-PraFakeCall "Disabled $($user.Upn)"
        } elseif ($user.Type -eq 'UserMailbox' -and $user.PendingShared) {
            $user.Type = 'SharedMailbox'; $user.PendingShared = $false
            Write-PraFakeCall "Shared $($user.Upn)"
        }
        foreach ($policy in $fake.CasePolicies) {
            $tag = 'UniH' + $policy.Guid
            $held = $policy.Locations.Contains($user.Id)
            if ($held -and -not $fake.Faults.NoHoldStamp -and -not $user.Holds.Contains($tag)) { $user.Holds.Add($tag); Write-PraFakeCall "HoldStamped $($user.Upn) $($policy.Name)" }
            elseif (-not $held -and $user.Holds.Contains($tag)) { [void]$user.Holds.Remove($tag); Write-PraFakeCall "HoldReleased $($user.Upn) $($policy.Name)" }
        }
    }
    foreach ($user in @($fake.Recycle.Values)) {
        $known = @($fake.Inactive | Where-Object { $_.ExchangeGuid -eq $user.ExchangeGuid }).Count
        if (-not $known -and -not $fake.Faults.NoInactive -and $user.Type -match 'Mailbox$' -and $user.Holds.Count) {
            $fake.Inactive.Add(@{ ExchangeGuid = $user.ExchangeGuid; Upn = $user.Upn; Holds = $(if ($fake.Faults.WrongInactiveHold) { @('UniHunrelated') } else { @($user.Holds) }); Tag = $user.Tag })
            Write-PraFakeCall "Inactive $($user.Upn) $($user.ExchangeGuid)"
        }
    }
}

function global:Invoke-PraFakeEntraConnect {
    param([Parameter(Mandatory)][ValidateSet('Pause','Resume','Delta')][string]$Operation)
    $fake = $global:PraFake
    Write-PraFakeCall "EntraConnect $Operation"
    switch ($Operation) {
        'Pause' { $fake.Scheduler = $false }
        'Resume' { $fake.Scheduler = $true }
        'Delta' {
            foreach ($user in $fake.Users.Values) {
                if (-not $user.Synced -or $user.CloudManaged) { continue }
                $user.DirSynced = $true
                # AD wins again: a MailUser whose ExchangeGuid was cleared in the cloud gets the on-premises one back.
                if ($user.Type -eq 'MailUser' -and $user.ExchangeGuid -eq [guid]::Empty.ToString()) { $user.ExchangeGuid = $user.OnPremGuid; Write-PraFakeCall "GuidFromAd $($user.Upn)" }
            }
            if ($fake.Faults.NoRecreate) { return 'fake Delta (objects not recreated)' }
            foreach ($ad in @($fake.Ad)) {
                if ([bool](Get-PraValue $ad 'SyncExcluded' $false)) { continue }
                if (@($fake.Users.Values | Where-Object { $_.ImmutableId -eq $ad.ImmutableId }).Count) { continue }
                $deleted = @($fake.Recycle.Values | Where-Object { $_.ImmutableId -eq $ad.ImmutableId }) | Select-Object -First 1
                if ($deleted) {
                    # Same immutableId still in the recycle bin: the sync restores the deleted object (lab SH).
                    $fake.Recycle.Remove($deleted.Id); $fake.Users[$deleted.Id] = $deleted
                    Write-PraFakeCall "Restored $($deleted.Upn) $($deleted.Id)"
                    continue
                }
                $new = Add-PraFakeUser -Upn $ad.Upn -Kind $ad.Kind -OnPremGuid $ad.OnPremGuid -ImmutableId $ad.ImmutableId -UsageLocation ''
                $global:PraFake.Ad.RemoveAt($global:PraFake.Ad.Count - 1)
                Write-PraFakeCall "Recreated $($ad.Upn) $($new.Id)"
            }
        }
    }
    return "fake $Operation"
}

function global:Invoke-PraFakeWait {
    <# Deterministic Wait-Pra2Condition: one tick per poll, at most 40 polls. #>
    param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds = 20)
    $interval = [Math]::Max(1, $IntervalSeconds)
    $polls = [Math]::Max(1, [Math]::Min(40, [int][Math]::Ceiling($TimeoutSeconds / $interval)))
    $last = $null
    for ($i = 0; $i -le $polls; $i++) {
        $last = & $Test
        if ($last) { return [pscustomobject]@{ Ok = $true; Seconds = $i * $interval; Last = $last } }
        Step-PraFake
    }
    return [pscustomobject]@{ Ok = $false; Seconds = $TimeoutSeconds; Last = $last }
}

#region Microsoft Graph -----------------------------------------------------------------------------------------------
function global:Stop-PraFakeGraph {
    param([int]$Code, [string]$GraphCode = '', [string]$Message = '')
    $response = [System.Net.Http.HttpResponseMessage]::new([System.Net.HttpStatusCode]$Code)
    $exception = [Microsoft.PowerShell.Commands.HttpResponseException]::new("Response status code does not indicate success: $Code.", $response)
    $record = [System.Management.Automation.ErrorRecord]::new($exception, 'FakeGraph', 'InvalidOperation', $null)
    if ($GraphCode) { $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new((@{ error = @{ code = $GraphCode; message = $Message } } | ConvertTo-Json -Compress)) }
    throw $record
}

function global:ConvertTo-PraFakeGraphUser {
    param([hashtable]$User)
    $group = $global:PraFake.Group
    $licences = @(); $states = @()
    foreach ($licence in $User.Direct) {
        $licences += [pscustomobject]@{ skuId = $licence.skuId; disabledPlans = @($licence.disabledPlans) }
        $states += [pscustomobject]@{ skuId = $licence.skuId; assignedByGroup = $null; disabledPlans = @($licence.disabledPlans); state = 'Active'; error = 'None' }
    }
    if ($group.Members.Contains($User.Id)) {
        foreach ($licence in $group.Licences) {
            $licences += [pscustomobject]@{ skuId = $licence.skuId; disabledPlans = @($licence.disabledPlans) }
            $states += [pscustomobject]@{ skuId = $licence.skuId; assignedByGroup = $group.Id; disabledPlans = @($licence.disabledPlans); state = 'Active'; error = 'None' }
        }
    }
    $plans = @(Get-PraFakeMailboxPlan $User | ForEach-Object { [pscustomobject]@{ service = 'exchange'; servicePlanId = $_; capabilityStatus = $(if ($global:PraFake.Faults.NoEnterprise -and $_ -eq 'plan-exo') { 'Suspended' } else { 'Enabled' }) } })
    if ($User.Direct.Count -or $group.Members.Contains($User.Id)) { $plans += [pscustomobject]@{ service = 'exchange'; servicePlanId = 'plan-foundation-other'; capabilityStatus = 'Enabled' } }
    return [pscustomobject]@{ id = $User.Id; userPrincipalName = $User.Upn; accountEnabled = $User.AccountEnabled; mail = (Get-PraValue $User 'Mail' $User.Upn); mailNickname = $User.Alias; proxyAddresses = @($User.Addresses)
        onPremisesSyncEnabled = $(if ($User.Synced -and -not $User.CloudManaged) { $true } else { $null }); onPremisesImmutableId = $User.ImmutableId
        usageLocation = $(if ($User.UsageLocation) { $User.UsageLocation } else { $null }); assignedLicenses = $licences; licenseAssignmentStates = $states
        assignedPlans = $plans; serviceProvisioningErrors = @() }
}

function global:ConvertFrom-PraFakeGraphFilter {
    param([string]$Filter)
    $state = @{ Position = 0; Tokens = [System.Collections.Generic.List[object]]::new() }
    $pattern = "\G\s*(?:(?<Proxy>proxyAddresses/any\(p:p\s+eq\s+'(?<ProxyValue>(?:[^']|'')*)'\))|(?<Property>userPrincipalName|mailNickname|mail|onPremisesImmutableId)\s+eq\s+'(?<Value>(?:[^']|'')*)'|(?<Operator>and\b|or\b|\(|\)))"
    while ($state.Position -lt $Filter.Length) {
        $match = [regex]::Match($Filter.Substring($state.Position), $pattern)
        if (-not $match.Success) { throw "Fake Graph: unsupported exact filter $Filter" }
        if ($match.Groups['Operator'].Success) { $state.Tokens.Add([pscustomobject]@{ Kind = $match.Groups['Operator'].Value }) }
        else {
            $proxy = $match.Groups['Proxy'].Success
            $state.Tokens.Add([pscustomobject]@{ Kind = 'Atom'; Property = $(if ($proxy) { 'proxyAddresses' } else { $match.Groups['Property'].Value })
                Value = $(if ($proxy) { $match.Groups['ProxyValue'].Value } else { $match.Groups['Value'].Value }).Replace("''", "'") })
        }
        $state.Position += $match.Length
        if ([string]::IsNullOrWhiteSpace($Filter.Substring($state.Position))) { break }
    }
    $state.Position = 0
    $primary = {
        if ($state.Position -ge $state.Tokens.Count) { throw 'Fake Graph: incomplete filter expression.' }
        $token = $state.Tokens[$state.Position]; $state.Position++
        if ($token.Kind -eq 'Atom') { return $token }
        if ($token.Kind -ne '(') { throw 'Fake Graph: expected equality predicate.' }
        $node = & $readOr
        if ($state.Position -ge $state.Tokens.Count -or $state.Tokens[$state.Position].Kind -ne ')') { throw 'Fake Graph: unclosed filter group.' }
        $state.Position++
        return $node
    }
    $readAnd = {
        $node = & $primary
        while ($state.Position -lt $state.Tokens.Count -and $state.Tokens[$state.Position].Kind -eq 'and') {
            $state.Position++
            $node = [pscustomobject]@{ Kind = 'and'; Left = $node; Right = (& $primary) }
        }
        return $node
    }
    $readOr = {
        $node = & $readAnd
        while ($state.Position -lt $state.Tokens.Count -and $state.Tokens[$state.Position].Kind -eq 'or') {
            $state.Position++
            $node = [pscustomobject]@{ Kind = 'or'; Left = $node; Right = (& $readAnd) }
        }
        return $node
    }
    $expression = & $readOr
    if ($state.Position -ne $state.Tokens.Count) { throw 'Fake Graph: trailing filter expression.' }
    return $expression
}

function global:Test-PraFakeGraphFilter {
    param([object]$Expression, [hashtable]$User)
    switch ($Expression.Kind) {
        and { return ((Test-PraFakeGraphFilter $Expression.Left $User) -and (Test-PraFakeGraphFilter $Expression.Right $User)) }
        or { return ((Test-PraFakeGraphFilter $Expression.Left $User) -or (Test-PraFakeGraphFilter $Expression.Right $User)) }
        Atom {
            switch ($Expression.Property) {
                userPrincipalName { return ($User.Upn -ieq $Expression.Value) }
                mail { return ((Get-PraValue $User 'Mail' $User.Upn) -ieq $Expression.Value) }
                mailNickname { return ($User.Alias -ieq $Expression.Value) }
                onPremisesImmutableId { return ($User.ImmutableId -ceq $Expression.Value) }
                proxyAddresses {
                    # Conservatively model protocol casing as significant; addresses themselves are case-insensitive.
                    foreach ($proxy in @($User.Addresses)) {
                        $prefix, $address = $proxy -split ':', 2
                        $wantedPrefix, $wantedAddress = $Expression.Value -split ':', 2
                        if ($prefix -ceq $wantedPrefix -and $address -ieq $wantedAddress) { return $true }
                    }
                    return $false
                }
            }
        }
    }
    throw 'Fake Graph: unknown filter expression.'
}

function global:Get-PraFakeGraphPage {
    param([string]$Path, [string]$Query, [System.Collections.IDictionary]$Headers)
    $fake = $global:PraFake
    $parameters = @{}
    foreach ($pair in @($Query -split '&' | Where-Object { $_ })) {
        $name, $value = $pair -split '=', 2
        $name = [uri]::UnescapeDataString($name)
        if ($parameters.ContainsKey($name)) { throw 'Fake Graph: duplicate query parameter.' }
        $parameters[$name] = [uri]::UnescapeDataString([string]$value)
    }
    $filter = [string]$parameters['$filter']
    $requestHeaders = @{}
    if ($Headers) { foreach ($name in $Headers.Keys) { $requestHeaders[$name] = $Headers[$name] } }
    $fake.GraphQueries.Add([pscustomobject]@{ Path = $Path; Filter = $filter; Query = $Query; Headers = $requestHeaders })
    if ($fake.Faults.GraphQueryFailsAt -eq $fake.GraphQueries.Count) { Stop-PraFakeGraph 403 'Authorization_RequestDenied' 'Graph query access denied.' }
    if ($Path -eq 'users' -and -not $filter -and $fake.Faults.RejectUnfilteredGraphInventory) { throw 'Unfiltered active Graph inventory is forbidden in this regression.' }
    if ($filter -and $fake.Faults.GraphQueryUnsupported) { Stop-PraFakeGraph 400 'Request_UnsupportedQuery' 'Unsupported equality property filter.' }
    $nonSmtp = @([regex]::Matches($filter, "proxyAddresses/any\(p:p\s+eq\s+'(?<Value>(?:[^']|'')*)'\)") |
        Where-Object { $_.Groups['Value'].Value -inotmatch '^smtp:' }).Count -gt 0
    $advanced = $parameters['$count'] -ceq 'true' -and $requestHeaders['ConsistencyLevel'] -ceq 'eventual'
    if (($nonSmtp -or $parameters.ContainsKey('$count') -or $requestHeaders.ContainsKey('ConsistencyLevel')) -and -not $advanced) {
        Stop-PraFakeGraph 400 'Request_UnsupportedQuery' 'Non-SMTP proxy equality requires $count=true and ConsistencyLevel: eventual.'
    }
    if ($advanced -and $fake.Faults.GraphAdvancedStatus -and
        ($fake.Faults.GraphAdvancedFailsAt -eq 0 -or $fake.Faults.GraphAdvancedFailsAt -le $fake.GraphQueries.Count)) {
        $status = [int]$fake.Faults.GraphAdvancedStatus
        $code = switch ($status) { 400 { 'Request_UnsupportedQuery' }; 403 { 'Authorization_RequestDenied' }; 429 { 'TooManyRequests' }; default { 'ServiceUnavailable' } }
        Stop-PraFakeGraph $status $code 'Advanced Graph query failed.'
    }
    $all = @(if ($Path -eq 'users') { $fake.Users.Values } else { $fake.Recycle.Values })
    if ($filter) {
        $expression = ConvertFrom-PraFakeGraphFilter $filter
        $all = @($all | Where-Object { Test-PraFakeGraphFilter -Expression $expression -User $_ })
    }
    $skip = if ($parameters.ContainsKey('$skiptoken')) { [int]$parameters['$skiptoken'] } else { 0 }
    $size = $fake.PageSize
    if ($parameters.ContainsKey('$top')) { $size = [Math]::Min($size, [int]$parameters['$top']) }
    if ($size -lt 1 -or $skip -lt 0) { throw 'Fake Graph: invalid paging parameters.' }
    $answer = [ordered]@{ value = @($all | Select-Object -Skip $skip -First $size | ForEach-Object { ConvertTo-PraFakeGraphUser $_ }) }
    if ($advanced -and $skip -eq 0) { $answer['@odata.count'] = $all.Count }
    if ($skip + $size -lt $all.Count) {
        $nextQuery = @($Query -split '&' | Where-Object { $_ -notmatch '^\$skiptoken=' }) -join '&'
        $answer['@odata.nextLink'] = 'https://graph.microsoft.com/v1.0/' + $Path + '?' + $nextQuery + '&$skiptoken=' + ($skip + $size)
    }
    return [pscustomobject]$answer
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method = 'GET', [string]$Uri, [object]$Body, [string]$ContentType, [string]$OutputType, [System.Collections.IDictionary]$Headers)
    $fake = $global:PraFake
    $relative = $Uri -replace '^https://graph\.microsoft\.com/', '' -replace '^v1\.0/', ''
    $path, $query = $relative -split '\?', 2
    $parts = @($path.Split('/') | ForEach-Object { [uri]::UnescapeDataString($_) })
    if ($Method -eq 'GET') {
        $fake.Reads.Add("Graph GET $path")
        if ($fake.Faults.NoLicenseReads -and $path -match '^(subscribedSkus|organization|groups)') { Stop-PraFakeGraph 403 'Authorization_RequestDenied' 'Licence/organization/group read denied.' }
    }
    if ($Method -eq 'GET' -and $fake.Faults.GraphReadFails -and $path -match '^users') { Stop-PraFakeGraph 403 'Authorization_RequestDenied' 'User read denied.' }
    $data = if ($Body) { $Body | ConvertFrom-Json -AsHashtable } else { @{} }
    if ($Method -ne 'GET' -and $path -ne '$batch') { Write-PraFakeCall "Graph $Method $path" }
    switch -Regex ($path) {
        '^subscribedSkus$' {
            $skus = foreach ($id in $fake.Skus.Keys) {
                $sku = $fake.Skus[$id]
                $used = @($fake.Users.Values | Where-Object { $user = $_; @($user.Direct | Where-Object { $_.skuId -eq $id }).Count -or ($fake.Group.Members.Contains($user.Id) -and @($fake.Group.Licences | Where-Object { $_.skuId -eq $id }).Count) }).Count
                [pscustomobject]@{ skuId = $id; skuPartNumber = $sku.Part; consumedUnits = $used; prepaidUnits = [pscustomobject]@{ enabled = $sku.Enabled }
                    servicePlans = @($sku.Plans | ForEach-Object { [pscustomobject]@{ servicePlanName = $_.Name; servicePlanId = $_.Id; appliesTo = $_.AppliesTo } }) }
            }
            return [pscustomobject]@{ value = @($skus) }
        }
        '^organization$' { return [pscustomobject]@{ value = @([pscustomobject]@{ id = 'org'; displayName = 'Contoso'; onPremisesSyncEnabled = $true; onPremisesLastSyncDateTime = '2026-10-07T08:00:00Z' }) } }
        '^groups/[^/]+$' {
            if ($parts[1] -ne $fake.Group.Id) { Stop-PraFakeGraph 404 'Request_ResourceNotFound' 'Group not found.' }
            $group = $fake.Group
            return [pscustomobject]@{ id = $group.Id; displayName = $group.Name; securityEnabled = (Get-PraValue $group 'SecurityEnabled' $true); mailEnabled = (Get-PraValue $group 'MailEnabled' $false); groupTypes = (Get-PraValue $group 'GroupTypes' @())
                onPremisesSyncEnabled = $(if ($group.Synced -and -not $group.CloudManaged) { $true } else { $null })
                onPremisesLastSyncDateTime = $(if ($group.Synced) { '2026-10-07T08:00:00Z' } else { $null })
                onPremisesSecurityIdentifier = $(if ($group.Synced) { 'S-1-5-21-100-200-300-400' } else { $null })
                assignedLicenses = @($group.Licences | ForEach-Object { [pscustomobject]@{ skuId = $_.skuId; disabledPlans = @($_.disabledPlans) } }) }
        }
        '^groups/[^/]+/onPremisesSyncBehavior$' {
            if ($Method -eq 'PATCH') { $fake.Group.CloudManaged = [bool]$data.isCloudManaged; return $null }
            return [pscustomobject]@{ isCloudManaged = $fake.Group.CloudManaged }
        }
        '^groups/[^/]+/members/\$ref$' {
            $id = ([string]$data['@odata.id']).Split('/')[-1]
            if ($fake.Group.Synced -and -not $fake.Group.CloudManaged) { Stop-PraFakeGraph 400 'Request_BadRequest' 'Cannot update a group synchronised from on-premises.' }
            if ($fake.Group.Members.Contains($id)) { Stop-PraFakeGraph 400 'Request_BadRequest' "One or more added object references already exist for the following modified properties: 'members'." }
            $fake.Group.Members.Add($id); return $null
        }
        '^groups/[^/]+/members/[^/]+/\$ref$' {
            if (-not $fake.Group.Members.Remove($parts[3])) { Stop-PraFakeGraph 404 'Request_ResourceNotFound' 'Member not found.' }
            return $null
        }
        '^users$' {
            return Get-PraFakeGraphPage -Path $path -Query ([string]$query) -Headers $Headers
        }
        '^\$batch$' {
            $responses = foreach ($request in @($data.requests)) {
                if ($fake.Faults.BatchThrottle -gt 0) { $fake.Faults.BatchThrottle--; [pscustomobject]@{ id = $request.id; status = 429; body = $null }; continue }
                $call = @{ Method = $request.method; Uri = ('v1.0' + $request.url) }
                if ($request.ContainsKey('body') -and $null -ne $request.body) { $call.Body = ($request.body | ConvertTo-Json -Depth 8 -Compress) }
                try { [pscustomobject]@{ id = $request.id; status = 200; body = (Invoke-MgGraphRequest @call) } }
                catch {
                    $code = 500; try { $code = [int]$_.Exception.Response.StatusCode } catch { $code = 500 }
                    $errorBody = $null; try { $errorBody = [string]$_.ErrorDetails.Message | ConvertFrom-Json } catch { $errorBody = $null }
                    [pscustomobject]@{ id = $request.id; status = $code; body = $errorBody }
                }
            }
            return [pscustomobject]@{ responses = @($responses) }
        }
        '^users/[^/]+$' {
            $user = Resolve-PraFakeUser $parts[1]
            if (-not $user) { Stop-PraFakeGraph 404 'Request_ResourceNotFound' "Resource '$($parts[1])' does not exist." }
            switch ($Method) {
                'GET' {
                    $result = ConvertTo-PraFakeGraphUser $user
                    if (Get-PraValue $user 'SignInVerificationActive' $false) {
                        $user.SignInPollReads++
                        if ($user.SignInPendingReads -gt 0) { $result.accountEnabled = $true; $user.SignInPendingReads-- }
                        switch ($fake.Faults.SignInReadFault) {
                            MissingAccount { $result.PSObject.Properties.Remove('accountEnabled') }
                            NullAccount { $result.accountEnabled = $null }
                            StringAccount { $result.accountEnabled = 'false' }
                            WrongId { $result.id = [guid]::NewGuid().ToString() }
                            Sync { $result.onPremisesSyncEnabled = $true }
                            Anchor { $result.onPremisesImmutableId = 'unexpected-anchor' }
                            MissingSync { $result.PSObject.Properties.Remove('onPremisesSyncEnabled') }
                            MissingAnchor { $result.PSObject.Properties.Remove('onPremisesImmutableId') }
                            NullAnchors { $result.onPremisesSyncEnabled = $null; $result.onPremisesImmutableId = $null }
                            ReadError { throw 'Sign-in read access denied.' }
                        }
                    }
                    return $result
                }
                'PATCH' {
                    if ($data.ContainsKey('usageLocation')) { $user.UsageLocation = [string]$data.usageLocation }
                    if ($data.ContainsKey('accountEnabled')) {
                        $user['SignInVerificationActive'] = $true; $user['SignInPollReads'] = 0
                        $user['SignInPendingReads'] = $fake.Faults.SignInPropagationReads
                        if (-not $fake.Faults.SignInDisableFails) { $user.AccountEnabled = [bool]$data.accountEnabled }
                    }
                    return $null
                }
                'DELETE' {
                    if ($fake.Faults.DeleteFails) { Stop-PraFakeGraph 403 'Authorization_RequestDenied' 'Insufficient privileges to complete the operation.' }
                    if ($fake.Faults.DeleteRejectedStatus) { Stop-PraFakeGraph $fake.Faults.DeleteRejectedStatus 'RequestRejected' 'Deletion refused.' }
                    $fake.Users.Remove($user.Id); $user.WhenSoftDeleted = [datetime]::UtcNow; $fake.Recycle[$user.Id] = $user
                    if ($fake.Faults.DeleteAmbiguousStatus) { Stop-PraFakeGraph $fake.Faults.DeleteAmbiguousStatus 'AmbiguousDelete' 'Deletion response lost.' }
                    if ($fake.Faults.CrashAfterDelete) { $fake.Faults.CrashAfterDelete = $false; throw 'Interrupted after identity deletion.' }
                    return $null
                }
            }
        }
        '^users/[^/]+/onPremisesSyncBehavior$' {
            $user = Resolve-PraFakeUser $parts[1]
            if (-not $user) { Stop-PraFakeGraph 404 'Request_ResourceNotFound' 'User not found.' }
            if ($Method -eq 'PATCH') { $user.CloudManaged = [bool]$data.isCloudManaged; return $null }
            return [pscustomobject]@{ isCloudManaged = $user.CloudManaged }
        }
        '^users/[^/]+/checkMemberGroups$' {
            return [pscustomobject]@{ value = @(@($data.groupIds) | Where-Object { $_ -eq $fake.Group.Id -and $fake.Group.Members.Contains($parts[1]) }) }
        }
        '^users/[^/]+/assignLicense$' {
            $user = Resolve-PraFakeUser $parts[1]
            if (@($data.addLicenses).Count) {
                if ($fake.Faults.LicenceAssign400 -gt 0) { $fake.Faults.LicenceAssign400--; Stop-PraFakeGraph 400 'Request_BadRequest' 'License assignment cannot be done for user with invalid usage location.' }
                if (-not $user.UsageLocation) { Stop-PraFakeGraph 400 'Request_BadRequest' 'License assignment cannot be done for user with invalid usage location.' }
                foreach ($licence in @($data.addLicenses)) {
                    @($user.Direct | Where-Object { $_.skuId -eq $licence.skuId }) | ForEach-Object { [void]$user.Direct.Remove($_) }
                    $user.Direct.Add(@{ skuId = [string]$licence.skuId; disabledPlans = @($licence.disabledPlans) })
                }
            }
            foreach ($skuId in @($data.removeLicenses)) { @($user.Direct | Where-Object { $_.skuId -eq $skuId }) | ForEach-Object { [void]$user.Direct.Remove($_) } }
            return $null
        }
        '^directory/deletedItems/[^/]+$' {
            if ($parts[2] -eq 'microsoft.graph.user') { return Get-PraFakeGraphPage -Path $path -Query ([string]$query) -Headers $Headers }
            if (-not $fake.Recycle.Contains($parts[2])) { Stop-PraFakeGraph 404 'Request_ResourceNotFound' 'Deleted item not found.' }
            if ($Method -eq 'DELETE') { $fake.Recycle.Remove($parts[2]); return $null }
            return ConvertTo-PraFakeGraphUser $fake.Recycle[$parts[2]]
        }
    }
    throw "Fake Graph: $Method $path is not implemented."
}
#endregion

#region Exchange Online and Security & Compliance ---------------------------------------------------------------------
function global:Get-PraFakeUserOrError {
    param([string]$Identity)
    $user = Resolve-PraFakeUser $Identity
    if (-not $user) { Write-Error "The operation couldn't be performed because object '$Identity' couldn't be found." }
    return $user
}

function global:Get-PraFakeFiltered {
    <# The users named by an OPATH filter of object IDs (ExternalDirectoryObjectId -eq '...' -or ...), as the bulk reads send it. #>
    param([string]$Filter)
    $global:PraFake.FilterCalls++
    foreach ($match in [regex]::Matches($Filter, "ExternalDirectoryObjectId -eq '([^']+)'")) {
        $id = $match.Groups[1].Value
        if ($global:PraFake.Users.Contains($id)) { $global:PraFake.Users[$id] }
    }
}

function global:ConvertTo-PraFakeRecipient {
    param([hashtable]$User)
    $result = [pscustomobject]@{ RecipientTypeDetails = $User.Type; ExchangeGuid = $User.ExchangeGuid; ExternalDirectoryObjectId = $User.Id; Guid = $User.Id
        Name = $User.Name; Alias = $User.Alias; EmailAddresses = @($User.Addresses); Identity = $User.Upn; PrimarySmtpAddress = $User.Upn
        LegacyExchangeDN = $(if ($User.ContainsKey('LegacyExchangeDN')) { $User.LegacyExchangeDN } else { '' }); CustomAttribute1 = $User.Tag; CustomAttribute15 = $User.Marker }
    if ($User.ContainsKey('WhenSoftDeleted')) { $result | Add-Member -NotePropertyName WhenSoftDeleted -NotePropertyValue $User.WhenSoftDeleted }
    return $result
}

function global:Get-PraFakeLocationText {
    <# MailboxLocations as Exchange Online writes them: '1;<guid>;<type>;<database>;<id>'. #>
    param([hashtable]$User)
    if ($User.Type -match 'Mailbox$') { '1;{0};Primary;NAMPRD01.PROD.OUTLOOK.COM;{1}' -f $User.ExchangeGuid, $User.Id }
    if ($User.Component) { '1;{0};ComponentShared;NAMPRD01.PROD.OUTLOOK.COM;{1}' -f $User.Component, $User.Id }
}

function global:Get-Recipient {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Filter, [object]$ResultSize, [switch]$IncludeSoftDeletedRecipients)
    if ($global:PraFake.Faults.ResourceReadAfterCreateFails -and $Filter -like "Name -eq 'PRA-R-*" -and
        @($global:PraFake.Users.Values | Where-Object { $_.Name -like 'PRA-R-*' }).Count) { throw 'Resource post-create read refused.' }
    if ($global:PraFake.Faults.ContactReadAfterCreateFails -and $Filter -like "Name -eq 'PRA-C-*" -and
        $global:PraFake.Contacts.Count) { throw 'Contact post-create read refused.' }
    if ($global:PraFake.Faults.RecipientReadFails) { throw 'Exchange recipient inventory access denied.' }
    if ($Identity -and $IncludeSoftDeletedRecipients) {
        $global:PraFake.RecipientQueries.Add([pscustomobject]@{ Identity = $Identity; Filter = ''; IncludeSoftDeletedRecipients = $true; ResultSize = $ResultSize; ErrorAction = $ErrorActionPreference })
        if ($global:PraFake.Faults.RecipientIdentityError) { $PSCmdlet.ThrowTerminatingError($global:PraFake.Faults.RecipientIdentityError) }
        if ($global:PraFake.Faults.RecipientIdentityEmpty) { return }
        $all = @($global:PraFake.Users.Values | ForEach-Object { ConvertTo-PraFakeRecipient $_ }) + @($global:PraFake.Contacts.Values | ForEach-Object { [pscustomobject]$_ }) +
            @($global:PraFake.RoomLists.Values | ForEach-Object { [pscustomobject]$_ }) + @($global:PraFake.Recycle.Values | ForEach-Object {
                    $recipient = ConvertTo-PraFakeRecipient $_
                    # Preserve the fake's existing proxy-release timing, but never hide a pure internal LegacyDN.
                    if (-not $_.Holds.Count -and -not $global:PraFake.Faults.SoftDeletedAddressBlocks) { $recipient.EmailAddresses = @() }
                    $recipient
                })
        $lookupResults = @(foreach ($recipient in $all) {
                $legacy = $recipient.PSObject.Properties['LegacyExchangeDN']
                $addresses = $recipient.PSObject.Properties['EmailAddresses']
                if (($legacy -and [string]$legacy.Value -ieq $Identity) -or ($addresses -and @($addresses.Value) -icontains "x500:$Identity")) { $recipient }
            })
        if (-not $lookupResults.Count) {
            $exception = [Exception]::new("Impossible d'effectuer l'opération, car l'objet est introuvable.")
            $exception.Data['RemoteException'] = [pscustomobject]@{ TypeName = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException'
                Message = $exception.Message; StackTrace = $null; InnerError = $null }
            $PSCmdlet.ThrowTerminatingError([System.Management.Automation.ErrorRecord]::new(
                $exception, '[],Write-ErrorMessage', 'NotSpecified', $Identity))
        }
        return $lookupResults
    }
    if ($Filter -or (-not $Identity -and $ResultSize)) {
        $global:PraFake.RecipientQueries.Add([pscustomobject]@{ Identity = ''; Filter = $Filter; IncludeSoftDeletedRecipients = [bool]$IncludeSoftDeletedRecipients; ResultSize = $ResultSize; ErrorAction = $ErrorActionPreference })
        if (-not $Filter -and $global:PraFake.Faults.RejectUnfilteredRecipientInventory) { throw 'Unfiltered recipient inventory is forbidden in this regression.' }
        if ($global:PraFake.Faults.RecipientQueryFailsAt -eq $global:PraFake.RecipientQueries.Count) { throw 'Exchange recipient filtered query access denied.' }
        $all = @($global:PraFake.Users.Values | ForEach-Object { ConvertTo-PraFakeRecipient $_ }) + @($global:PraFake.Contacts.Values | ForEach-Object { [pscustomobject]$_ }) +
            @($global:PraFake.RoomLists.Values | ForEach-Object { [pscustomobject]$_ })
        if ($IncludeSoftDeletedRecipients) { $all += @($global:PraFake.Recycle.Values | Where-Object { $_.Holds.Count -or $global:PraFake.Faults.SoftDeletedAddressBlocks } | ForEach-Object { ConvertTo-PraFakeRecipient $_ }) }
        if ($Filter) {
            $pattern = "(Alias|EmailAddresses|PrimarySmtpAddress|Name) -eq '((?:[^']|'')*)'"
            $predicates = [regex]::Matches($Filter, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $remainder = [regex]::Replace($Filter, $pattern, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            $remainder = $remainder -replace '(?i)-or', '' -replace '[\s()]', ''
            if (-not $predicates.Count -or $remainder) { throw "Fake Get-Recipient: unsupported OPATH filter $Filter" }
            return @(foreach ($recipient in $all) {
                    foreach ($predicate in $predicates) {
                        $field = $predicate.Groups[1].Value
                        $value = $predicate.Groups[2].Value.Replace("''", "'")
                        $property = $recipient.PSObject.Properties[$field]
                        if ($property -and @($property.Value | ForEach-Object { [string]$_ }) -contains $value) { $recipient; break }
                    }
                })
        }
        return $all
    }
    foreach ($contact in $global:PraFake.Contacts.Values) {
        if ($contact.Guid -eq $Identity -or $contact.Identity -eq $Identity -or $contact.PrimarySmtpAddress -eq $Identity -or $contact.EmailAddresses -contains "smtp:$Identity") { return [pscustomobject]$contact }
    }
    foreach ($group in $global:PraFake.RoomLists.Values) {
        if ($group.Guid -eq $Identity -or $group.PrimarySmtpAddress -eq $Identity) { return [pscustomobject]$group }
    }
    $user = Get-PraFakeUserOrError $Identity
    if (-not $user) { return }
    ConvertTo-PraFakeRecipient $user
}

function global:Get-EXORecipient {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Filter, [object]$ResultSize, [string[]]$Properties, [string[]]$RecipientTypeDetails)
    if ($global:PraFake.Faults.ExoFilterFails) { throw 'A server side error has occurred because of which the operation could not be completed.' }
    foreach ($user in @(Get-PraFakeFiltered $Filter)) { ConvertTo-PraFakeRecipient $user }
}

function global:Get-User {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Filter, [object]$ResultSize)
    $users = if ($Filter) { @(Get-PraFakeFiltered $Filter) } else { @(Get-PraFakeUserOrError $Identity) }
    foreach ($user in @($users | Where-Object { $_ })) {
        $result = $user.Directory.Clone()
        $result.IsDirSynced = $user.DirSynced; $result.ExternalDirectoryObjectId = $user.Id; $result.UserPrincipalName = $user.Upn
        $result.RecipientTypeDetails = if ($user.Type -match 'Mailbox$') { $user.Type } else { 'User' }
        if ($global:PraFake.Faults.DirectoryReadMismatch) { $result.City = 'Wrong city' }
        if ($global:PraFake.Faults.DirectoryNotesReadMismatch) { $result.Notes = 'Wrong notes' }
        if ($global:PraFake.Faults.DirectoryNotesReadMissing) { $result.Remove('Notes') }
        [pscustomobject]$result
    }
}

function global:Set-User {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param([string]$Identity, [string]$Office, [string]$StreetAddress, [string]$City, [string]$StateOrProvince,
        [string]$PostalCode, [string]$CountryOrRegion, [string]$Phone, [string]$Notes, [switch]$Force)
    $user = Get-PraFakeUserOrError $Identity
    Write-PraFakeCall "Set-User $($user.Upn)"
    $parameters = @{}
    foreach ($entry in $PSBoundParameters.GetEnumerator()) { $parameters[$entry.Key] = $entry.Value }
    $global:PraFake.UserWrites.Add($parameters)
    $fields = @('Office','StreetAddress','City','StateOrProvince','PostalCode','CountryOrRegion','Phone','Notes')
    # Model the observed noninteractive failure of Set-User's built-in pause without prompting the test host.
    if (@($fields | Where-Object { $PSBoundParameters.ContainsKey($_) }).Count -and -not $Force -and
        (-not $PSBoundParameters.ContainsKey('Confirm') -or [bool]$PSBoundParameters['Confirm'])) {
        throw 'A server side error has occurred because of which the operation could not be completed.'
    }
    if ($global:PraFake.Faults.DirectoryWriteFails) { throw 'Resource directory update refused.' }
    foreach ($field in $fields) {
        if ($PSBoundParameters.ContainsKey($field)) { $user.Directory[$field] = [string]$PSBoundParameters[$field] }
    }
}

function global:Get-DistributionGroup {
    [CmdletBinding()] param([string]$Identity, [string]$Filter, [object]$ResultSize)
    if ($global:PraFake.Faults.RoomListReadAfterCreateFails -and $global:PraFake.RoomLists.Count) { throw 'Room-list post-create read refused.' }
    $groups = @($global:PraFake.RoomLists.Values)
    if ($Identity) {
        $groups = @($groups | Where-Object { $_.Guid -eq $Identity -or $_.PrimarySmtpAddress -eq $Identity })
        if (-not $groups.Count) { throw "Room list $Identity not found." }
    }
    if ($Filter -match "^(Name|PrimarySmtpAddress) -eq '([^']+)'$") {
        $property = $Matches[1]; $value = $Matches[2]
        $groups = @($groups | Where-Object { $_[$property] -eq $value })
    }
    foreach ($group in $groups) {
        $result = $group.Clone()
        if ($global:PraFake.Faults.RoomListConfigMismatch) { $result.MemberJoinRestriction = 'Open' }
        [pscustomobject]$result
    }
}

function global:New-DistributionGroup {
    [CmdletBinding()] param([string]$Name, [string]$Alias, [string]$DisplayName, [string]$PrimarySmtpAddress, [switch]$RoomList,
        [switch]$IgnoreNamingPolicy, [string]$MemberJoinRestriction, [string]$MemberDepartRestriction, [bool]$RequireSenderAuthenticationEnabled)
    if ($global:PraFake.Faults.NewRoomListFails) { throw 'New-DistributionGroup refused.' }
    if (@($global:PraFake.RoomLists.Values | Where-Object { $_.Alias -eq $Alias -or $_.PrimarySmtpAddress -eq $PrimarySmtpAddress -or $_.Name -eq $Name }).Count) { throw 'Room-list identity already assigned.' }
    $guid = [guid]::NewGuid().ToString()
    $group = @{ Guid = $guid; ExternalDirectoryObjectId = $guid; Name = $Name; Alias = $Alias; DisplayName = $DisplayName; PrimarySmtpAddress = $PrimarySmtpAddress
        EmailAddresses = @("SMTP:$PrimarySmtpAddress"); Identity = $PrimarySmtpAddress; IsDirSynced = $false; RecipientTypeDetails = 'RoomList'; HiddenFromAddressListsEnabled = $false
        CustomAttribute15 = ''; MemberJoinRestriction = $MemberJoinRestriction; MemberDepartRestriction = $MemberDepartRestriction
        RequireSenderAuthenticationEnabled = $RequireSenderAuthenticationEnabled; ModerationEnabled = $false; SendModerationNotifications = 'Never'
        AcceptMessagesOnlyFromSendersOrMembers = @(); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @()
        Members = [System.Collections.Generic.List[string]]::new() }
    if ($global:PraFake.Faults.RoomListExchangeOnly) { $group.ExternalDirectoryObjectId = '' }
    $global:PraFake.RoomLists[$guid] = $group
    Write-PraFakeCall "New-RoomList $PrimarySmtpAddress Name=$Name"
    if ($global:PraFake.Faults.CrashAfterNewRoomList) { $global:PraFake.Faults.CrashAfterNewRoomList = $false; throw 'Interrupted after room-list creation.' }
    [pscustomobject]$group
}

function global:Set-DistributionGroup {
    [CmdletBinding()] param([string]$Identity, [string[]]$EmailAddresses, [bool]$HiddenFromAddressListsEnabled, [string]$CustomAttribute15,
        [switch]$BypassSecurityGroupManagerCheck, [string]$MemberJoinRestriction, [string]$MemberDepartRestriction,
        [bool]$RequireSenderAuthenticationEnabled, [bool]$ModerationEnabled, [string]$SendModerationNotifications,
        [string[]]$AcceptMessagesOnlyFromSendersOrMembers, [string[]]$RejectMessagesFromSendersOrMembers, [string[]]$ModeratedBy, [string[]]$BypassModerationFromSendersOrMembers)
    if ($global:PraFake.Faults.RoomListConfigFails) { throw 'Room-list configuration refused.' }
    $read = Get-DistributionGroup -Identity $Identity
    $group = $global:PraFake.RoomLists[$read.Guid]
    foreach ($key in $PSBoundParameters.Keys) {
        if ($key -notin @('Identity','BypassSecurityGroupManagerCheck','ErrorAction')) { $group[$key] = $PSBoundParameters[$key] }
    }
    Write-PraFakeCall "Set-RoomList $($group.PrimarySmtpAddress)"
    if ($global:PraFake.Faults.CrashAfterRoomListConfig) { $global:PraFake.Faults.CrashAfterRoomListConfig = $false; throw 'Interrupted after room-list configuration.' }
}

function global:Get-DistributionGroupMember {
    [CmdletBinding()] param([string]$Identity, [object]$ResultSize)
    if ($global:PraFake.Faults.RoomListReadFails) { throw 'Room list membership read refused.' }
    $group = Get-DistributionGroup -Identity $Identity
    foreach ($id in $group.Members) { ConvertTo-PraFakeRecipient (Get-PraFakeUserOrError $id) }
}

function global:Add-DistributionGroupMember {
    [CmdletBinding()] param([string]$Identity, [string]$Member, [switch]$BypassSecurityGroupManagerCheck)
    if ($global:PraFake.Faults.RoomListAddFails) { throw 'Room list membership update refused.' }
    $group = Get-DistributionGroup -Identity $Identity
    if (-not $global:PraFake.Faults.RoomListNoMembership -and -not $group.Members.Contains($Member)) { $group.Members.Add($Member) }
    Write-PraFakeCall "RoomList Add $($group.PrimarySmtpAddress) $Member"
    if ($global:PraFake.Faults.CrashAfterRoomListAdd) { $global:PraFake.Faults.CrashAfterRoomListAdd = $false; throw 'Interrupted after room-list membership update.' }
}

function global:Get-MailUser {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Filter, [object]$ResultSize)
    if ($Filter) { $users = @(Get-PraFakeFiltered $Filter | Where-Object { $_.Type -eq 'MailUser' }) }
    else {
        $user = Get-PraFakeUserOrError $Identity
        if (-not $user) { return }
        if ($user.Type -ne 'MailUser') { Write-Error "$Identity isn't a mail user."; return }
        $users = @($user)
    }
    foreach ($user in $users) {
        [pscustomobject]@{ ExchangeGuid = $user.ExchangeGuid; InPlaceHolds = @($user.Holds); CustomAttribute1 = $user.Tag; ExternalDirectoryObjectId = $user.Id
            MailboxLocations = @(Get-PraFakeLocationText $user) }
    }
}

function global:Get-EXOMailbox {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Filter, [object]$ResultSize, [string[]]$Properties)
    foreach ($user in @(Get-PraFakeFiltered $Filter | Where-Object { $_.Type -match 'Mailbox$' })) {
        [pscustomobject]@{ ExchangeGuid = $user.ExchangeGuid; InPlaceHolds = @($user.Holds); LitigationHoldEnabled = $false; ExternalDirectoryObjectId = $user.Id
            MailboxLocations = @(Get-PraFakeLocationText $user); PrimarySmtpAddress = $user.Upn }
    }
}

function global:Get-Mailbox {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [switch]$InactiveMailboxOnly, [switch]$SoftDeletedMailbox, [object]$ResultSize)
    if ($SoftDeletedMailbox) {
        $global:PraFake.Reads.Add('Get-Mailbox SoftDeletedMailbox')
        if ($global:PraFake.Faults.NoSoftDeleted) { return }
        foreach ($one in @($global:PraFake.Recycle.Values | Where-Object { $_.Type -match 'Mailbox$' -and (-not $Identity -or $_.ExchangeGuid -eq $Identity) })) {
            [pscustomobject]@{ ExchangeGuid = $one.ExchangeGuid; ExternalDirectoryObjectId = $one.Id; IsInactiveMailbox = [bool]$one.Holds.Count; InPlaceHolds = @($one.Holds)
                LitigationHoldEnabled = $false; RetentionHoldEnabled = $false; DelayHoldApplied = $false; DelayReleaseHoldApplied = $false; ComplianceTagHoldApplied = $false }
        }
        return
    }
    if ($InactiveMailboxOnly) {
        $global:PraFake.Reads.Add('Get-Mailbox InactiveMailboxOnly')
        $inactive = @($global:PraFake.Inactive | Where-Object { -not $Identity -or $_.ExchangeGuid -eq $Identity -or $_.Upn -eq $Identity })
        if (-not $inactive.Count) { if ($Identity) { Write-Error "The operation couldn't be performed because object '$Identity' couldn't be found." }; return }
        foreach ($one in $inactive) { [pscustomobject]@{ ExchangeGuid = $one.ExchangeGuid; InPlaceHolds = @($one.Holds); CustomAttribute1 = $one.Tag; PrimarySmtpAddress = $one.Upn } }
        return
    }
    $user = Get-PraFakeUserOrError $Identity
    if (-not $user) { return }
    if ($user.Type -notmatch 'Mailbox$') { Write-Error "$Identity isn't a mailbox."; return }
    $result = @{ ExchangeGuid = $user.ExchangeGuid; ExternalDirectoryObjectId = $user.Id; InPlaceHolds = @($user.Holds); LitigationHoldEnabled = $false; GrantSendOnBehalfTo = @($user.SendOnBehalf)
        CustomAttribute1 = $user.Tag; CustomAttribute15 = $user.Marker; CustomAttribute14 = (Get-PraValue $user 'Marker14' ''); DelayHoldApplied = $false; DelayReleaseHoldApplied = $false
        PrimarySmtpAddress = $user.Upn; EmailAddresses = @($user.Addresses); ResourceCapacity = $user.ResourceCapacity; ResourceCustom = @($user.ResourceCustom); HiddenFromAddressListsEnabled = $user.Hidden }
    foreach ($field in @('LitigationHoldEnabled','RetentionHoldEnabled','DelayHoldApplied','DelayReleaseHoldApplied','ComplianceTagHoldApplied')) {
        $result[$field] = Get-PraValue $user $field $false
    }
    foreach ($field in $user.Delivery.Keys) { $result[$field] = $user.Delivery[$field] }
    [pscustomobject]$result
}

function global:Get-MailboxLocation {
    [CmdletBinding()] param([string]$User)
    $item = Get-PraFakeUserOrError $User
    if (-not $item) { return }
    if ($item.Type -match 'Mailbox$') { [pscustomobject]@{ MailboxLocationType = 'Primary'; MailboxGuid = $item.ExchangeGuid } }
    if ($item.Component) { [pscustomobject]@{ MailboxLocationType = 'ComponentShared'; MailboxGuid = $item.Component } }
}

function global:Set-MailUser {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [object]$ExchangeGuid)
    $user = Get-PraFakeUserOrError $Identity
    if (-not $user) { return }
    if ($user.Type -ne 'MailUser') { Write-Error "$Identity isn't a mail user."; return }
    $user.ExchangeGuid = [string]$ExchangeGuid
    Write-PraFakeCall "Set-MailUser $($user.Upn) ExchangeGuid=$ExchangeGuid"
}

function global:Set-Mailbox {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Type, [string]$CustomAttribute1, [string]$CustomAttribute2, [string]$CustomAttribute3,
        [string]$CustomAttribute4, [string]$CustomAttribute5, [string]$CustomAttribute6, [string]$CustomAttribute7, [string]$CustomAttribute8,
        [string]$CustomAttribute9, [string]$CustomAttribute10, [string]$CustomAttribute11, [string]$CustomAttribute12, [string]$CustomAttribute13, [string]$CustomAttribute14, [string]$CustomAttribute15,
        [string[]]$EmailAddresses, [int]$ResourceCapacity, [string[]]$ResourceCustom, [bool]$HiddenFromAddressListsEnabled, [object]$GrantSendOnBehalfTo,
        [bool]$RequireSenderAuthenticationEnabled, [bool]$ModerationEnabled, [string]$SendModerationNotifications, [string[]]$AcceptMessagesOnlyFromSendersOrMembers,
        [string[]]$RejectMessagesFromSendersOrMembers, [string[]]$ModeratedBy, [string[]]$BypassModerationFromSendersOrMembers)
    $user = Get-PraFakeUserOrError $Identity
    if (-not $user) { return }
    if ($user.Type -notmatch 'Mailbox$') { Write-Error "$Identity isn't a mailbox."; return }
    if ($Type -eq 'Shared' -and $user.Type -ne 'SharedMailbox') { $user.PendingShared = $true; Write-PraFakeCall "Set-Mailbox $($user.Upn) Type=Shared" }
    if ($PSBoundParameters.ContainsKey('CustomAttribute1')) { $user.Tag = $CustomAttribute1; Write-PraFakeCall "Set-Mailbox $($user.Upn) CustomAttribute1=$CustomAttribute1" }
    if ($PSBoundParameters.ContainsKey('CustomAttribute15')) { $user.Marker = $CustomAttribute15 }
    if ($PSBoundParameters.ContainsKey('CustomAttribute14')) { $user['Marker14'] = $CustomAttribute14 }
    if ($PSBoundParameters.ContainsKey('EmailAddresses')) { $user.Addresses = @($EmailAddresses); $user.Upn = (@($EmailAddresses | Where-Object { $_ -cmatch '^SMTP:' })[0]).Substring(5) }
    if ($PSBoundParameters.ContainsKey('ResourceCapacity')) { $user.ResourceCapacity = $ResourceCapacity }
    if ($PSBoundParameters.ContainsKey('ResourceCustom')) { $user.ResourceCustom = @($ResourceCustom) }
    if ($PSBoundParameters.ContainsKey('HiddenFromAddressListsEnabled')) { $user.Hidden = $HiddenFromAddressListsEnabled }
    if ($PSBoundParameters.ContainsKey('HiddenFromAddressListsEnabled') -and -not $HiddenFromAddressListsEnabled -and $global:PraFake.Faults.HiddenFinalFails) { throw 'Final visibility write interrupted.' }
    foreach ($field in @($user.Delivery.Keys)) {
        if ($PSBoundParameters.ContainsKey($field)) { $user.Delivery[$field] = $PSBoundParameters[$field] }
    }
    if ($GrantSendOnBehalfTo -is [hashtable] -and $GrantSendOnBehalfTo.ContainsKey('Add')) {
        $delegate = Resolve-PraFakeUser ([string]$GrantSendOnBehalfTo.Add)
        if (-not $delegate) { Write-Error "Delegate $($GrantSendOnBehalfTo.Add) not found."; return }
        if (-not $user.SendOnBehalf.Contains($delegate.Id)) { $user.SendOnBehalf.Add($delegate.Id) }
        Write-PraFakeCall "SendOnBehalf $($user.Upn) $($delegate.Upn)"
    }
}

function global:New-Mailbox {
    [CmdletBinding()] param([string]$Name, [string]$Alias, [string]$DisplayName, [string]$PrimarySmtpAddress, [switch]$Room, [switch]$Equipment)
    if ($global:PraFake.Faults.NewMailboxFails) { throw 'New-Mailbox refused.' }
    if (Resolve-PraFakeUser $PrimarySmtpAddress) { throw 'Address already assigned.' }
    $kind = if ($Room) { 'Room' } else { 'Equipment' }
    $user = Add-PraFakeUser -Upn $PrimarySmtpAddress -Kind $kind -UsageLocation ''
    $global:PraFake.Ad.RemoveAt($global:PraFake.Ad.Count - 1)
    $user.Synced = $false; $user.DirSynced = $false; $user.ImmutableId = ''; $user.ExchangeGuid = [guid]::NewGuid().ToString()
    $user.Type = "${kind}Mailbox"; $user.Name = $Name; $user.Alias = $Alias; $user.AccountEnabled = [bool]$global:PraFake.Faults.SignInDisableFails
    $user.Calendar = @{ AutomateProcessing = 'AutoAccept'; AllBookInPolicy = $true; AllRequestInPolicy = $false; AllRequestOutOfPolicy = $false
        ForwardRequestsToDelegates = $true; ProcessExternalMeetingMessages = $false
        ResourceDelegates = @(); BookInPolicy = @(); RequestInPolicy = @(); RequestOutOfPolicy = @() }
    if ($global:PraFake.Faults.SignInPropagationReads -gt 0) { $user.AccountEnabled = $true }
    Write-PraFakeCall "New-Mailbox $PrimarySmtpAddress $kind Name=$Name"
    if ($global:PraFake.Faults.CrashAfterNewMailbox) { $global:PraFake.Faults.CrashAfterNewMailbox = $false; throw 'Interrupted after mailbox creation.' }
    ConvertTo-PraFakeRecipient $user
}

function global:Set-CalendarProcessing {
    [CmdletBinding()] param([string]$Identity, [string]$AutomateProcessing, [bool]$AllowConflicts, [int]$BookingWindowInDays, [int]$MaximumDurationInMinutes,
        [bool]$AllowRecurringMeetings, [bool]$EnforceSchedulingHorizon, [bool]$ScheduleOnlyDuringWorkHours, [int]$ConflictPercentageAllowed, [int]$MaximumConflictInstances,
        [bool]$ForwardRequestsToDelegates, [bool]$DeleteAttachments, [bool]$DeleteComments, [bool]$RemovePrivateProperty, [bool]$DeleteSubject, [bool]$AddOrganizerToSubject,
        [bool]$DeleteNonCalendarItems, [bool]$TentativePendingApproval, [bool]$EnableResponseDetails, [bool]$OrganizerInfo, [bool]$AllBookInPolicy, [bool]$AllRequestInPolicy,
        [bool]$AllRequestOutOfPolicy, [bool]$AddAdditionalResponse, [string]$AdditionalResponse, [bool]$ProcessExternalMeetingMessages, [bool]$RemoveOldMeetingMessages,
        [string[]]$ResourceDelegates, [string[]]$BookInPolicy, [string[]]$RequestInPolicy, [string[]]$RequestOutOfPolicy)
    $user = Get-PraFakeUserOrError $Identity
    if ($AutomateProcessing -eq 'None' -and $global:PraFake.Faults.BookingDisableFails) { throw 'Booking disable refused.' }
    if ($AutomateProcessing -ne 'None' -and $global:PraFake.Faults.CalendarEnableFails) { throw 'Booking activation refused.' }
    foreach ($key in $PSBoundParameters.Keys) { if ($key -notin @('Identity', 'ErrorAction')) { $user.Calendar[$key] = $PSBoundParameters[$key] } }
    Write-PraFakeCall "Set-CalendarProcessing $($user.Upn) AutomateProcessing=$($user.Calendar.AutomateProcessing) AllBookInPolicy=$($user.Calendar.AllBookInPolicy)"
}

function global:Get-CalendarProcessing {
    [CmdletBinding()] param([string]$Identity)
    $calendar = (Get-PraFakeUserOrError $Identity).Calendar.Clone()
    if ($calendar.AutomateProcessing -ne 'None' -and $global:PraFake.Faults.CalendarEnabledReadMismatch) { $calendar.AllBookInPolicy = -not $calendar.AllBookInPolicy }
    [pscustomobject]$calendar
}

function global:Set-MailboxCalendarConfiguration {
    [CmdletBinding()] param([string]$Identity, [string]$WorkingHoursStartTime, [string]$WorkingHoursEndTime, [string[]]$WorkDays, [string]$WorkingHoursTimeZone)
    $user = Get-PraFakeUserOrError $Identity
    $user.WorkHours = @{ WorkingHoursStartTime = $WorkingHoursStartTime; WorkingHoursEndTime = $WorkingHoursEndTime; WorkDays = ($WorkDays -join ', '); WorkingHoursTimeZone = $WorkingHoursTimeZone }
    Write-PraFakeCall "Set-MailboxCalendarConfiguration $($user.Upn) WorkDays=$($WorkDays -join ',')"
}

function global:Get-MailboxCalendarConfiguration {
    [CmdletBinding()] param([string]$Identity)
    $hours = (Get-PraFakeUserOrError $Identity).WorkHours.Clone()
    if ($global:PraFake.Faults.WorkDaysAliasRead) { $hours.WorkDays = $global:PraFake.Faults.WorkDaysAliasRead }
    if ($global:PraFake.Faults.WorkTimeZoneObjectRead) { $hours.WorkingHoursTimeZone = [pscustomobject]@{ Id = $hours.WorkingHoursTimeZone; DisplayName = 'Localized time-zone display name' } }
    [pscustomobject]$hours
}

function global:New-MailContact {
    [CmdletBinding()] param([string]$Name, [string]$Alias, [string]$ExternalEmailAddress)
    $global:PraFake.Reads.Add('New-MailContact attempt')
    if ($global:PraFake.Faults.ContactFails) { throw 'Contact address reuse blocked; deleted account retained.' }
    if ($global:PraFake.Faults.ContactSoftDeletedAliasBlocks -and @($global:PraFake.Recycle.Values | Where-Object { $_.Alias -eq $Alias }).Count) {
        throw 'Exchange New-MailContact uniqueness conflict: alias retained by soft-deleted recipient.'
    }
    $id = [guid]::NewGuid().ToString()
    $global:PraFake.Contacts[$id] = @{ Guid = $id; ExternalDirectoryObjectId = ''; ExchangeGuid = ''; Name = $Name; Alias = $Alias; Identity = $id; RecipientTypeDetails = 'MailContact'
        ExternalEmailAddress = $ExternalEmailAddress; PrimarySmtpAddress = $ExternalEmailAddress; EmailAddresses = @("SMTP:$ExternalEmailAddress"); CustomAttribute15 = ''; CustomAttribute14 = ''; HiddenFromAddressListsEnabled = $false }
    Write-PraFakeCall "New-MailContact $Name $ExternalEmailAddress"
    if ($global:PraFake.Faults.CrashAfterContact) { $global:PraFake.Faults.CrashAfterContact = $false; throw 'Interrupted after routing contact creation.' }
    [pscustomobject]$global:PraFake.Contacts[$id]
}

function global:Get-MailContact {
    [CmdletBinding()] param([string]$Identity)
    $contact = $global:PraFake.Contacts.Values | Where-Object { $_.Guid -eq $Identity -or $_.Name -eq $Identity } | Select-Object -First 1
    if (-not $contact) { throw 'Mail contact not found.' }
    [pscustomobject]$contact
}

function global:Set-MailContact {
    [CmdletBinding()] param([string]$Identity, [string]$ExternalEmailAddress, [string[]]$EmailAddresses, [bool]$HiddenFromAddressListsEnabled,
        [string]$CustomAttribute15, [string]$CustomAttribute14)
    $global:PraFake.Reads.Add('Set-MailContact attempt')
    if ($global:PraFake.Faults.ContactSoftDeletedProxyBlocks -and @($global:PraFake.Recycle.Values | Where-Object {
        @($_.Addresses | Where-Object { $_ -in $EmailAddresses }).Count
    }).Count) { throw 'Exchange Set-MailContact uniqueness conflict: proxy retained by soft-deleted recipient.' }
    $contact = $global:PraFake.Contacts[$Identity]
    $contact.ExternalEmailAddress = $ExternalEmailAddress; $contact.EmailAddresses = @($EmailAddresses)
    $contact.PrimarySmtpAddress = (@($EmailAddresses | Where-Object { $_ -cmatch '^SMTP:' })[0]).Substring(5)
    $contact.HiddenFromAddressListsEnabled = $HiddenFromAddressListsEnabled
    if ($PSBoundParameters.ContainsKey('CustomAttribute15')) { $contact.CustomAttribute15 = $CustomAttribute15 }
    if ($PSBoundParameters.ContainsKey('CustomAttribute14')) { $contact.CustomAttribute14 = $CustomAttribute14 }
    Write-PraFakeCall "Set-MailContact $Identity $ExternalEmailAddress"
}

function global:Get-MailboxPermission {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$User)
    $mailbox = Get-PraFakeUserOrError $Identity; $trustee = Resolve-PraFakeUser $User
    if ($mailbox -and $trustee -and $mailbox.FullAccess.Contains($trustee.Id)) { [pscustomobject]@{ User = $trustee.Upn; AccessRights = @('FullAccess'); Deny = $false } }
}

function global:Add-MailboxPermission {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$User, [string[]]$AccessRights, [string]$InheritanceType, [bool]$AutoMapping = $true)
    $mailbox = Get-PraFakeUserOrError $Identity; $trustee = Resolve-PraFakeUser $User
    if (-not $mailbox -or -not $trustee) { Write-Error "Add-MailboxPermission: $Identity / $User not found."; return }
    $mailbox.FullAccess.Add($trustee.Id); $mailbox.AutoMap[$trustee.Id] = $AutoMapping
    Write-PraFakeCall "FullAccess $($mailbox.Upn) $($trustee.Upn) AutoMapping=$AutoMapping"
}

function global:Get-RecipientPermission {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Trustee)
    $mailbox = Get-PraFakeUserOrError $Identity; $user = Resolve-PraFakeUser $Trustee
    if ($mailbox -and $user -and $mailbox.SendAs.Contains($user.Id)) { [pscustomobject]@{ Trustee = $user.Upn; AccessRights = @('SendAs') } }
}

function global:Add-RecipientPermission {
    [CmdletBinding(SupportsShouldProcess = $true)] param([Parameter(Position = 0)][string]$Identity, [string]$Trustee, [string[]]$AccessRights)
    $mailbox = Get-PraFakeUserOrError $Identity; $user = Resolve-PraFakeUser $Trustee
    if (-not $mailbox -or -not $user) { Write-Error "Add-RecipientPermission: $Identity / $Trustee not found."; return }
    if ($global:PraFake.Faults.SendAsFails) { Write-Error 'The operation could not be performed (transient).'; return }
    $mailbox.SendAs.Add($user.Id)
    Write-PraFakeCall "SendAs $($mailbox.Upn) $($user.Upn)"
}

function global:Get-OrganizationConfig {
    [CmdletBinding()] param()
    $global:PraFake.Reads.Add('Get-OrganizationConfig')
    if ($global:PraFake.Faults.OrganizationReadFails) { throw 'Organization holds access denied.' }
    [pscustomobject]@{ InPlaceHolds = @($global:PraFake.OrgHolds) }
}

function global:Get-ResourceConfig {
    [CmdletBinding()] param()
    [pscustomobject]@{ ResourcePropertySchema = @($global:PraFake.ResourceSchema) }
}

function global:Get-CaseHoldPolicy {
    <# -Identity: one policy with its locations (by object ID, as the lab shows). -Case: the policies of a case, WITHOUT their locations (lab 8 Oct). #>
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Case)
    $global:PraFake.Reads.Add('Get-CaseHoldPolicy')
    if ($Case) {
        return @($global:PraFake.CasePolicies | Where-Object { $_.CaseId -eq $Case } | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Guid = [guid]$_.Guid; CaseId = $_.CaseId; Mode = (Get-PraValue $_ 'Mode' 'Enforce'); ExchangeLocation = @() } })
    }
    $policy = $global:PraFake.CasePolicies | Where-Object { $_.Name -eq $Identity } | Select-Object -First 1
    if (-not $policy) { Write-Error "Policy $Identity not found."; return }
    [pscustomobject]@{ Name = $policy.Name; Guid = [guid]$policy.Guid; CaseId = $policy.CaseId; Mode = (Get-PraValue $policy 'Mode' 'Enforce')
        ExchangeLocation = @($policy.Locations | ForEach-Object { $u = Resolve-PraFakeUser $_; [pscustomobject]@{ Name = $(if ($u) { $u.Upn } else { $_ }); ImmutableIdentity = $_ } }) }
}

function global:Set-CaseHoldPolicy {
    <# Many locations in one call. One location already there (or not there) fails the call after the others are changed. #>
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string[]]$AddExchangeLocation, [string[]]$RemoveExchangeLocation)
    $policy = $global:PraFake.CasePolicies | Where-Object { $_.Name -eq $Identity } | Select-Object -First 1
    if (-not $policy) { Write-Error "Policy $Identity not found."; return }
    $global:PraFake.CaseHoldCalls++
    if ($global:PraFake.Faults.CaseHoldFails) { Write-Error "Policy $Identity cannot be changed now (fake fault)."; return }
    $problem = ''
    foreach ($one in @($AddExchangeLocation | Where-Object { $_ })) {
        $user = Resolve-PraFakeUser $one
        if (-not $user) { $problem = "The location $one couldn't be found."; continue }
        if ($policy.Locations.Contains($user.Id)) { $problem = "The location $one is already in the policy."; continue }
        $policy.Locations.Add($user.Id); Write-PraFakeCall "CaseHold Add $($user.Upn) $($policy.Name)"
    }
    foreach ($one in @($RemoveExchangeLocation | Where-Object { $_ })) {
        $user = Resolve-PraFakeUser $one
        if (-not $user -or -not $policy.Locations.Remove($user.Id)) { $problem = "The location $one was not found in the policy."; continue }
        Write-PraFakeCall "CaseHold Remove $($user.Upn) $($policy.Name)"
    }
    if ($problem) { Write-Error $problem; return }
    if ($global:PraFake.Faults.CaseHoldDeployError) { Write-Error "|Microsoft.Exchange.Management.UnifiedPolicy.PolicyDeploymentException|Policy '$($policy.Guid)' failed to be deployed."; return }
}

function global:New-CaseHoldPolicy {
    [CmdletBinding()] param([string]$Name, [string]$Case, [bool]$Enabled = $true)
    if ($global:PraFake.CasePolicies | Where-Object { $_.Name -eq $Name }) { Write-Error "A policy named $Name already exists."; return }
    $policy = @{ Name = $Name; Guid = [guid]::NewGuid().ToString(); CaseId = $Case; Locations = [System.Collections.Generic.List[string]]::new() }
    $global:PraFake.CasePolicies.Add($policy)
    Write-PraFakeCall "CaseHoldPolicy New $Name"
    [pscustomobject]@{ Name = $Name; Guid = [guid]$policy.Guid; CaseId = $Case; Enabled = $Enabled }
}

function global:New-CaseHoldRule {
    [CmdletBinding()] param([string]$Name, [string]$Policy)
    $global:PraFake.CaseRules.Add("$Policy/$Name")
    Write-PraFakeCall "CaseHoldRule New $Name"
    [pscustomobject]@{ Name = $Name; Policy = $Policy }
}
#endregion

$global:PraFakeCommands = @('Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext', 'Reset-PraFakeGraphTransport', 'Invoke-MgGraphRequest', 'Get-Recipient', 'Get-User', 'Get-MailUser', 'Get-Mailbox', 'Get-MailboxLocation', 'Set-MailUser', 'Set-Mailbox', 'Get-OrganizationConfig', 'Get-PraFakeMailboxPlan',
    'New-Mailbox', 'Set-CalendarProcessing', 'Get-CalendarProcessing', 'Set-MailboxCalendarConfiguration', 'Get-MailboxCalendarConfiguration', 'New-MailContact', 'Get-MailContact', 'Set-MailContact', 'Get-ResourceConfig',
    'Set-User', 'Get-DistributionGroup', 'Get-DistributionGroupMember', 'Add-DistributionGroupMember',
    'New-DistributionGroup', 'Set-DistributionGroup',
    'Get-MailboxPermission', 'Add-MailboxPermission', 'Get-RecipientPermission', 'Add-RecipientPermission', 'Get-CaseHoldPolicy', 'Set-CaseHoldPolicy', 'New-CaseHoldPolicy', 'New-CaseHoldRule',
    'New-PraFakeTenant', 'Get-PraFakeHoldTag', 'Add-PraFakeUser', 'Write-PraFakeCall', 'Resolve-PraFakeUser', 'Test-PraFakeExchangePlan', 'Step-PraFake',
    'New-PraFakeResourceSettings',
    'Invoke-PraFakeEntraConnect', 'Invoke-PraFakeWait', 'Stop-PraFakeGraph', 'ConvertTo-PraFakeGraphUser', 'Get-PraFakeUserOrError',
    'ConvertFrom-PraFakeGraphFilter', 'Test-PraFakeGraphFilter', 'Get-PraFakeGraphPage',
    'Get-PraFakeFiltered', 'ConvertTo-PraFakeRecipient', 'Get-PraFakeLocationText', 'Get-EXORecipient', 'Get-EXOMailbox')
