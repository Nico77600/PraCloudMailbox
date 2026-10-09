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

function global:New-PraFakeTenant {
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
        Recycle = [ordered]@{}
        Ad = [System.Collections.Generic.List[hashtable]]::new()
        Inactive = [System.Collections.Generic.List[hashtable]]::new()
        # Case hold policies of the eDiscovery case; CasePolicy is the first one (the configured policy).
        CasePolicies = [System.Collections.Generic.List[hashtable]]::new()
        CasePolicy = $null
        CaseRules = [System.Collections.Generic.List[string]]::new()
        Scheduler = $true
        Faults = @{ LicenceAssign400 = 0; NeverMailbox = @(); NoHoldStamp = $false; DeleteFails = $false; NoRecreate = $false; SendAsFails = $false; CaseHoldDeployError = $false; CaseHoldFails = $false; StopAfter = $null; BatchThrottle = 0; ExoFilterFails = $false }
        PageSize = 999; FilterCalls = 0; CaseHoldCalls = 0
        OrgHolds = @()
        Calls = [System.Collections.Generic.List[string]]::new()
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

function global:Add-PraFakeUser {
    param([Parameter(Mandatory)][string]$Upn, [ValidateSet('User','Shared','Room','Equipment')][string]$Kind = 'User', [string]$Id = ([guid]::NewGuid().ToString()),
        [string]$OnPremGuid = ([guid]::NewGuid().ToString()), [string]$ImmutableId = ('imm-' + $Upn), [string]$Component, [string]$UsageLocation = 'FR')
    $user = @{ Id = $Id; Upn = $Upn; Kind = $Kind; ImmutableId = $ImmutableId; OnPremGuid = $OnPremGuid; Synced = $true; CloudManaged = $false
        UsageLocation = $UsageLocation; Direct = [System.Collections.Generic.List[hashtable]]::new()
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
    foreach ($user in $global:PraFake.Users.Values) { if ($user.Upn -eq $Identity) { return $user } }
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
        if (-not $known -and $user.Type -match 'Mailbox$' -and $user.Holds.Count) {
            $fake.Inactive.Add(@{ ExchangeGuid = $user.ExchangeGuid; Upn = $user.Upn; Holds = @($user.Holds); Tag = $user.Tag })
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
    $plans = @(Get-PraFakeMailboxPlan $User | ForEach-Object { [pscustomobject]@{ service = 'exchange'; servicePlanId = $_; capabilityStatus = 'Enabled' } })
    if ($User.Direct.Count -or $group.Members.Contains($User.Id)) { $plans += [pscustomobject]@{ service = 'exchange'; servicePlanId = 'plan-foundation-other'; capabilityStatus = 'Enabled' } }
    return [pscustomobject]@{ id = $User.Id; userPrincipalName = $User.Upn; accountEnabled = $true
        onPremisesSyncEnabled = $(if ($User.Synced -and -not $User.CloudManaged) { $true } else { $null }); onPremisesImmutableId = $User.ImmutableId
        usageLocation = $(if ($User.UsageLocation) { $User.UsageLocation } else { $null }); assignedLicenses = $licences; licenseAssignmentStates = $states
        assignedPlans = $plans; serviceProvisioningErrors = @() }
}

function global:Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method = 'GET', [string]$Uri, [object]$Body, [string]$ContentType, [string]$OutputType)
    $fake = $global:PraFake
    $relative = $Uri -replace '^https://graph\.microsoft\.com/', '' -replace '^v1\.0/', ''
    $path, $query = $relative -split '\?', 2
    $parts = @($path.Split('/') | ForEach-Object { [uri]::UnescapeDataString($_) })
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
            return [pscustomobject]@{ id = $group.Id; displayName = $group.Name; securityEnabled = $true; mailEnabled = $false
                onPremisesSyncEnabled = $(if ($group.Synced -and -not $group.CloudManaged) { $true } else { $null })
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
            $filter = [uri]::UnescapeDataString([string]$query)
            if ($filter -match 'onPremisesImmutableId eq') {
                $immutable = [regex]::Match($filter, "onPremisesImmutableId eq '([^']*)'").Groups[1].Value
                return [pscustomobject]@{ value = @($fake.Users.Values | Where-Object { $_.ImmutableId -eq $immutable } | ForEach-Object { ConvertTo-PraFakeGraphUser $_ }) }
            }
            # Every user, page by page.
            $token = [regex]::Match($filter, '\$skiptoken=(\d+)')
            $skip = if ($token.Success) { [int]$token.Groups[1].Value } else { 0 }
            $all = @($fake.Users.Values)
            $page = @($all | Select-Object -Skip $skip -First $fake.PageSize | ForEach-Object { ConvertTo-PraFakeGraphUser $_ })
            $answer = [ordered]@{ value = $page }
            if ($skip + $fake.PageSize -lt $all.Count) { $answer['@odata.nextLink'] = 'https://graph.microsoft.com/v1.0/users?$select=id&$skiptoken={0}' -f ($skip + $fake.PageSize) }
            return [pscustomobject]$answer
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
                'GET' { return ConvertTo-PraFakeGraphUser $user }
                'PATCH' { if ($data.ContainsKey('usageLocation')) { $user.UsageLocation = [string]$data.usageLocation }; return $null }
                'DELETE' {
                    if ($fake.Faults.DeleteFails) { Stop-PraFakeGraph 403 'Authorization_RequestDenied' 'Insufficient privileges to complete the operation.' }
                    $fake.Users.Remove($user.Id); $fake.Recycle[$user.Id] = $user; return $null
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
            if (-not $fake.Recycle.Contains($parts[2])) { Stop-PraFakeGraph 404 'Request_ResourceNotFound' 'Deleted item not found.' }
            if ($Method -eq 'DELETE') { $fake.Recycle.Remove($parts[2]); return $null }
            return [pscustomobject]@{ id = $parts[2] }
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
    [pscustomobject]@{ RecipientTypeDetails = $User.Type; ExchangeGuid = $User.ExchangeGuid; ExternalDirectoryObjectId = $User.Id; Guid = $User.Id
        Name = ($User.Upn -split '@')[0]; Identity = $User.Upn; PrimarySmtpAddress = $User.Upn; CustomAttribute1 = $User.Tag }
}

function global:Get-PraFakeLocationText {
    <# MailboxLocations as Exchange Online writes them: '1;<guid>;<type>;<database>;<id>'. #>
    param([hashtable]$User)
    if ($User.Type -match 'Mailbox$') { '1;{0};Primary;NAMPRD01.PROD.OUTLOOK.COM;{1}' -f $User.ExchangeGuid, $User.Id }
    if ($User.Component) { '1;{0};ComponentShared;NAMPRD01.PROD.OUTLOOK.COM;{1}' -f $User.Component, $User.Id }
}

function global:Get-Recipient {
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity)
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
        [pscustomobject]@{ IsDirSynced = $user.DirSynced; ExternalDirectoryObjectId = $user.Id; UserPrincipalName = $user.Upn
            RecipientTypeDetails = $(if ($user.Type -match 'Mailbox$') { $user.Type } else { 'User' }) }
    }
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
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [switch]$InactiveMailboxOnly, [object]$ResultSize)
    if ($InactiveMailboxOnly) {
        $inactive = @($global:PraFake.Inactive | Where-Object { $_.ExchangeGuid -eq $Identity -or $_.Upn -eq $Identity })
        if (-not $inactive.Count) { Write-Error "The operation couldn't be performed because object '$Identity' couldn't be found."; return }
        return [pscustomobject]@{ ExchangeGuid = $inactive[0].ExchangeGuid; InPlaceHolds = @($inactive[0].Holds); CustomAttribute1 = $inactive[0].Tag; PrimarySmtpAddress = $inactive[0].Upn }
    }
    $user = Get-PraFakeUserOrError $Identity
    if (-not $user) { return }
    if ($user.Type -notmatch 'Mailbox$') { Write-Error "$Identity isn't a mailbox."; return }
    [pscustomobject]@{ ExchangeGuid = $user.ExchangeGuid; InPlaceHolds = @($user.Holds); LitigationHoldEnabled = $false; GrantSendOnBehalfTo = @($user.SendOnBehalf)
        CustomAttribute1 = $user.Tag; DelayHoldApplied = $false; DelayReleaseHoldApplied = $false }
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
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Type, [string]$CustomAttribute1, [object]$GrantSendOnBehalfTo)
    $user = Get-PraFakeUserOrError $Identity
    if (-not $user) { return }
    if ($user.Type -notmatch 'Mailbox$') { Write-Error "$Identity isn't a mailbox."; return }
    if ($Type -eq 'Shared' -and $user.Type -ne 'SharedMailbox') { $user.PendingShared = $true; Write-PraFakeCall "Set-Mailbox $($user.Upn) Type=Shared" }
    if ($PSBoundParameters.ContainsKey('CustomAttribute1')) { $user.Tag = $CustomAttribute1; Write-PraFakeCall "Set-Mailbox $($user.Upn) CustomAttribute1=$CustomAttribute1" }
    if ($GrantSendOnBehalfTo -is [hashtable] -and $GrantSendOnBehalfTo.ContainsKey('Add')) {
        $delegate = Resolve-PraFakeUser ([string]$GrantSendOnBehalfTo.Add)
        if (-not $delegate) { Write-Error "Delegate $($GrantSendOnBehalfTo.Add) not found."; return }
        if (-not $user.SendOnBehalf.Contains($delegate.Id)) { $user.SendOnBehalf.Add($delegate.Id) }
        Write-PraFakeCall "SendOnBehalf $($user.Upn) $($delegate.Upn)"
    }
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
    [pscustomobject]@{ InPlaceHolds = @($global:PraFake.OrgHolds) }
}

function global:Get-CaseHoldPolicy {
    <# -Identity: one policy with its locations (by object ID, as the lab shows). -Case: the policies of a case, WITHOUT their locations (lab 8 Oct). #>
    [CmdletBinding()] param([Parameter(Position = 0)][string]$Identity, [string]$Case)
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

$global:PraFakeCommands = @('Invoke-MgGraphRequest', 'Get-Recipient', 'Get-User', 'Get-MailUser', 'Get-Mailbox', 'Get-MailboxLocation', 'Set-MailUser', 'Set-Mailbox', 'Get-OrganizationConfig', 'Get-PraFakeMailboxPlan',
    'Get-MailboxPermission', 'Add-MailboxPermission', 'Get-RecipientPermission', 'Add-RecipientPermission', 'Get-CaseHoldPolicy', 'Set-CaseHoldPolicy', 'New-CaseHoldPolicy', 'New-CaseHoldRule',
    'New-PraFakeTenant', 'Get-PraFakeHoldTag', 'Add-PraFakeUser', 'Write-PraFakeCall', 'Resolve-PraFakeUser', 'Test-PraFakeExchangePlan', 'Step-PraFake',
    'Invoke-PraFakeEntraConnect', 'Invoke-PraFakeWait', 'Stop-PraFakeGraph', 'ConvertTo-PraFakeGraphUser', 'Get-PraFakeUserOrError',
    'Get-PraFakeFiltered', 'ConvertTo-PraFakeRecipient', 'Get-PraFakeLocationText', 'Get-EXORecipient', 'Get-EXOMailbox')
