#Requires -Version 5.1
Describe 'Native resource safety rules' {
    BeforeAll {
        $root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        Import-Module (Join-Path $root 'module\PRA2.Common.psm1') -Force
        Import-Module (Join-Path $root 'module\PRA2.Cloud.psm1') -Force
    }

    Describe 'CloudResource Collect wiring' {
        BeforeAll {
            $root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
            Import-Module (Join-Path $root 'module\PRA2.Common.psm1') -Force
            Import-Module (Join-Path $root 'module\PRA2.Collect.psm1') -Force
            Import-Module (Join-Path $root 'module\PRA2.Store.psm1') -Force
            $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
            foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -in @('Invoke-PraCollect','Skip-PraStep','New-PraRow') })) {
                . ([scriptblock]::Create($function.Extent.Text))
            }
            function Start-ResourceCollectTest {
                param([string]$ResourceMode, [string]$Kind, [string]$Mode = 'Preview')
                $config = @{ Collect = @{ MailboxStatistics = $false; SharedPermissions = $false; Contacts = $false; DistributionGroups = $false }
                    Scope = @{ ResourceMode = $ResourceMode }; Exchange = @{ DomainController = '' } }
                $run = @{ Config = $config; Mode = $Mode; Action = 'Collect'; CurrentOperation = ''; CurrentIdentity = ''
                    Excluded = [System.Collections.Generic.List[string]]::new(); Rows = [System.Collections.Generic.List[object]]::new(); NextSteps = @() }
                $mailbox = [pscustomobject]@{ PrimarySmtpAddress = 'resource@contoso.com'; RecipientTypeDetails = $Kind; Guid = 'source-guid' }
                foreach ($entry in @{ context = $run; Action = 'Collect'; Identity = ''; Scope = 'All'; Force = $true; CollectTestMailbox = $mailbox }.GetEnumerator()) {
                    Set-Variable -Name $entry.Key -Value $entry.Value -Scope 1
                }
            }
            Mock Connect-Pra2Exchange { 'offline' }
            Mock Get-Pra2ScopeMailbox { $CollectTestMailbox }
            Mock Get-Pra2ResourceSettings { @{ SchemaVersion = 1; Capacity = 12 } }
            Mock Get-Pra2SharedPermission { @{ Permissions = @(); Members = @(); Warnings = @() } }
            Mock ConvertTo-Pra2MailboxRecord {
                [ordered]@{ object_guid = [string]$Mailbox.Guid; kind = ([string]$Mailbox.RecipientTypeDetails -replace 'Mailbox$', '')
                    primary_smtp_address = [string]$Mailbox.PrimarySmtpAddress; user_principal_name = [string]$Mailbox.PrimarySmtpAddress; exchange_guid = 'source-exchange-guid'
                    legacy_exchange_dn = ''; email_addresses_json = '["SMTP:resource@contoso.com"]'
                    resource_settings_json = $(if ($ResourceSettings) { $ResourceSettings | ConvertTo-Json -Compress } else { $null }) }
            }
            Mock Write-PraStep { }
            Mock Write-PraItem { }
            Mock Write-PraLog { }
            Mock Skip-PraStep { }
            Mock Open-Pra2Store { throw 'Snapshot writing must not be reached.' }
        }

        It 'collects booking settings only for opted-in native resources (<ResourceMode>, <Kind>)' -TestCases @(
            @{ ResourceMode = 'CloudResource'; Kind = 'RoomMailbox'; Calls = 1 },
            @{ ResourceMode = 'CloudResource'; Kind = 'EquipmentMailbox'; Calls = 1 },
            @{ ResourceMode = 'CloudResource'; Kind = 'UserMailbox'; Calls = 0 },
            @{ ResourceMode = 'CloudResource'; Kind = 'SharedMailbox'; Calls = 0 },
            @{ ResourceMode = 'SynchronizedShared'; Kind = 'RoomMailbox'; Calls = 0 },
            @{ ResourceMode = 'SynchronizedShared'; Kind = 'EquipmentMailbox'; Calls = 0 }) {
            param($ResourceMode, $Kind, $Calls)
            Start-ResourceCollectTest -ResourceMode $ResourceMode -Kind $Kind
            Invoke-PraCollect
            Should -Invoke Get-Pra2ResourceSettings -Times $Calls -Exactly
            Should -Invoke ConvertTo-Pra2MailboxRecord -Times 1 -Exactly
            Should -Invoke ConvertTo-Pra2MailboxRecord -Times $Calls -Exactly -ParameterFilter { $null -ne $ResourceSettings }
            Should -Invoke Open-Pra2Store -Times 0 -Exactly
        }

        It 'aborts before snapshot writing when native settings fail or are missing (<Failure>)' -TestCases @(
            @{ Failure = 'ReadError' }, @{ Failure = 'Empty' }) {
            param($Failure)
            Start-ResourceCollectTest -ResourceMode CloudResource -Kind RoomMailbox -Mode Apply
            if ($Failure -eq 'ReadError') { Mock Get-Pra2ResourceSettings { throw 'Unresolved booking delegate.' } }
            else { Mock Get-Pra2ResourceSettings { $null } }
            { Invoke-PraCollect } | Should -Throw '*Collect stopped before writing the snapshot*resource@contoso.com*'
            Should -Invoke ConvertTo-Pra2MailboxRecord -Times 0 -Exactly
            Should -Invoke Open-Pra2Store -Times 0 -Exactly
        }
        It 'does not collect general mailbox rights for native rooms/equipment even in mixed scope (<Kind>)' -TestCases @(
            @{ Kind = 'RoomMailbox' }, @{ Kind = 'EquipmentMailbox' }) {
            param($Kind)
            Start-ResourceCollectTest -ResourceMode CloudResource -Kind $Kind
            $context.Config.Collect.SharedPermissions = $true
            $shared = [pscustomobject]@{ Guid = 'shared-guid'; PrimarySmtpAddress = 'shared@contoso.com'; RecipientTypeDetails = 'SharedMailbox' }
            Mock Get-Pra2ScopeMailbox { @($CollectTestMailbox, $shared) }
            Invoke-PraCollect
            Should -Invoke Get-Pra2SharedPermission -Times 1 -Exactly -ParameterFilter { $Mailbox.Guid -eq 'shared-guid' }
            Should -Invoke Get-Pra2SharedPermission -Times 0 -Exactly -ParameterFilter { $Mailbox.RecipientTypeDetails -eq $Kind }
        }
        It 'preserves legacy synced room/equipment general permission collection (<Kind>)' -TestCases @(
            @{ Kind = 'RoomMailbox' }, @{ Kind = 'EquipmentMailbox' }) {
            param($Kind)
            Start-ResourceCollectTest -ResourceMode SynchronizedShared -Kind $Kind
            $context.Config.Collect.SharedPermissions = $true
            Invoke-PraCollect
            Should -Invoke Get-Pra2SharedPermission -Times 1 -Exactly
        }
    }
    BeforeEach {
        $config = @{ Scope = @{ ResourceMode = 'CloudResource'; ConvertRooms = $true }
            Resources = @{ OnPremisesReady = $true; OnPremisesRoutingDomain = 'onprem.contoso.com' } }
        $record = [pscustomobject]@{ object_guid = 'source'; kind = 'Room'; alias = 'room'; primary_smtp_address = 'room@contoso.com'; user_principal_name = 'room@contoso.com'
            immutable_id = 'source-anchor'; exchange_guid = [guid]::NewGuid().ToString(); email_addresses_json = '["SMTP:room@contoso.com","smtp:alternate@contoso.com","X500:/o=legacy/cn=room"]'
            legacy_exchange_dn = '/o=old/cn=room'; resource_settings_json = '' }
    }

    It 'defaults to synchronized shared behavior and requires explicit CloudResource selection' {
        $config.Scope.Remove('ResourceMode')
        Test-Pra2NativeResource -Config $config -Record $record | Should -BeFalse
        $config.Scope.ResourceMode = 'CloudResource'
        Test-Pra2NativeResource -Config $config -Record $record | Should -BeTrue
        $record.kind = 'Shared'
        Test-Pra2NativeResource -Config $config -Record $record | Should -BeFalse
    }

    It 'does not infer address availability from absent Graph user' {
        $state = [pscustomobject]@{ NativeResource = $true; User = $null; Collisions = @(); AbsenceVerified = $false; Owned = $false }
        @(Get-Pra2Readiness -Config $config -Record $record -State $state | Where-Object Code -eq RESOURCE_ABSENCE_UNVERIFIED).Count | Should -Be 1
    }

    It 'requires settings recollection and explicit ConvertRooms even in CloudResource mode' {
        $config.Scope.ConvertRooms = $false
        $state = [pscustomobject]@{ NativeResource = $true; User = $null; Collisions = @(); AbsenceVerified = $true; Owned = $false }
        $findings = @(Get-Pra2Readiness -Config $config -Record $record -State $state)
        $findings.Code | Should -Contain RESOURCE_COLLECT_REQUIRED
        $findings.Code | Should -Contain KIND_NOT_SUPPORTED
        $findings.Code | Should -Not -Contain ENTRA_NOT_FOUND
    }

    It 'refuses schema-one rows lacking the resource settings property without a property-access error' {
        $record.PSObject.Properties.Remove('resource_settings_json')
        { Read-Pra2ResourceSettings -Record $record } | Should -Throw '*v1 snapshot*Collect is required*'
    }

    It 'preserves all SMTP/X500 including legacy DN with exactly the original primary' {
        $addresses = @(Get-Pra2ResourceAddresses $record)
        $addresses | Should -Contain 'SMTP:room@contoso.com'
        $addresses | Should -Contain 'smtp:alternate@contoso.com'
        $addresses | Should -Contain 'X500:/o=legacy/cn=room'
        $addresses | Should -Contain 'X500:/o=old/cn=room'
        @($addresses | Where-Object { $_ -cmatch '^SMTP:' }).Count | Should -Be 1
    }

    It 'detects a non-user Exchange recipient proxy collision independently of Graph' {
        $inventory = [pscustomobject]@{ Users = @(); Recipients = @([pscustomobject]@{ Identity = 'unrelated'; RecipientTypeDetails = 'MailContact'; Alias = 'other'
                    PrimarySmtpAddress = 'other@contoso.com'; EmailAddresses = @('smtp:alternate@contoso.com'); Guid = 'contact' }) }
        @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 1
    }

    It 'does not ignore a different resource GUID just because it has the same address' {
        $inventory = [pscustomobject]@{ Users = @(); Recipients = @([pscustomobject]@{ Identity = 'room'; RecipientTypeDetails = 'RoomMailbox'; Alias = 'room'
                    PrimarySmtpAddress = 'room@contoso.com'; EmailAddresses = @(); ExchangeGuid = 'different'; ExternalDirectoryObjectId = 'foreign' }) }
        @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory -OwnedId 'owned' -OwnedGuid 'owned-guid').Count | Should -Be 1
    }

    It 'accepts only an enabled enterprise plan and active assignment from the specific resource group' {
        $group = [pscustomobject]@{ Id = 'resource-group'; EnterprisePlans = @('enterprise'); EnterpriseAssignments = @([pscustomobject]@{ SkuId = 'sku-enterprise'; PlanIds = @('enterprise') }) }
        $user = [pscustomobject]@{ assignedPlans = @([pscustomobject]@{ servicePlanId = 'enterprise'; capabilityStatus = 'Enabled' })
            licenseAssignmentStates = @([pscustomobject]@{ skuId = 'sku-enterprise'; assignedByGroup = 'different-group'; state = 'Active'; error = 'None'; disabledPlans = @() }) }
        Test-Pra2ResourceLicence -User $user -Group $group | Should -BeFalse
        $user.licenseAssignmentStates[0].assignedByGroup = 'resource-group'
        Test-Pra2ResourceLicence -User $user -Group $group | Should -BeTrue
        $user.assignedPlans[0].capabilityStatus = 'Suspended'
        Test-Pra2ResourceLicence -User $user -Group $group | Should -BeFalse
    }

    It 'refuses missing, malformed, looping and unasserted resource routing (<Fault>)' -TestCases @(
        @{ Fault = 'Empty' }, @{ Fault = 'Malformed' }, @{ Fault = 'PrimaryLoop' }, @{ Fault = 'ProxyLoop' }, @{ Fault = 'NotReady' }) {
        param($Fault)
        switch ($Fault) {
            Empty { $config.Resources.OnPremisesRoutingDomain = '' }
            Malformed { $config.Resources.OnPremisesRoutingDomain = 'https://onprem.contoso.com' }
            PrimaryLoop { $config.Resources.OnPremisesRoutingDomain = 'contoso.com' }
            ProxyLoop { $record.email_addresses_json = '["SMTP:room@contoso.com","smtp:room@onprem.contoso.com"]' }
            NotReady { $config.Resources.OnPremisesReady = $false }
        }
        { Get-Pra2ResourceRoute -Config $config -Record $record -RequireReady } | Should -Throw
    }

    It 'returns the exact hybrid destination without mutating the source' {
        Get-Pra2ResourceRoute -Config $config -Record $record -RequireReady | Should -Be 'room@onprem.contoso.com'
        $record.primary_smtp_address | Should -Be 'room@contoso.com'
    }

    It 'selects persisted lifecycle independently of current configuration' {
        $item = [pscustomobject]@{ original_json = '{"lifecycle":"CloudResource","cloudId":"immutable-id","cloudGuid":"independent-guid","token":"0123456789abcdef0123456789abcdef","creationName":"PRA-R-01234567-89abcdef-0123456789abcdef0123456789abcdef","markerAttribute":"CustomAttribute15"}' }
        $config.Scope.ResourceMode = 'SynchronizedShared'
        (Get-Pra2ResourceIntent -Item $item).cloudId | Should -Be 'immutable-id'
        $item.original_json = '{"entraId":"legacy"}'
        Get-Pra2ResourceIntent -Item $item | Should -BeNullOrEmpty
    }

    It 'never downgrades missing/legacy protection mode or accepts an unknown mode' {
        Get-Pra2ResourceProtectionMode -Intent $null | Should -BeExactly None
        Get-Pra2ResourceProtectionMode -Intent ([pscustomobject]@{ lifecycle = 'CloudResource' }) | Should -BeExactly CaseHold
        Get-Pra2ResourceProtectionMode -Intent ([pscustomobject]@{ protectionMode = 'None' }) | Should -BeExactly None
        { Get-Pra2ResourceProtectionMode -Intent ([pscustomobject]@{ protectionMode = '' }) } | Should -Throw '*Unknown*'
        { Get-Pra2ResourceProtectionMode -Intent ([pscustomobject]@{ protectionMode = @('None') }) } | Should -Throw '*Unknown*'
        { Get-Pra2ResourceProtectionMode -Intent ([pscustomobject]@{ protectionMode = 'None'; holdTag = 'UniHlegacy' }) } | Should -Throw '*conflicts*'
    }

    It 'does not require usageLocation or licensing for the new native readiness path' {
        $config.Cloud = @{ DefaultUsageLocation = '' }
        $state = [pscustomobject]@{ NativeResource = $true; Collisions = @(); AbsenceVerified = $true; ProtectionMode = 'None' }
        $findings = @(Get-Pra2Readiness -Config $config -Record $record -State $state)
        $findings.Code | Should -Not -Contain NO_USAGE_LOCATION
        $findings.Code | Should -Contain RESOURCE_UNPROTECTED
        $state.ProtectionMode = 'CaseHold'
        @(Get-Pra2Readiness -Config $config -Record $record -State $state).Code | Should -Contain NO_USAGE_LOCATION
    }

    It 'fails closed on a missing, empty, misspelled or non-string readiness protection mode (<Name>)' -TestCases @(
        @{ Name = 'missing'; Mode = '<missing>' }, @{ Name = 'empty'; Mode = '' }, @{ Name = 'lowercase none'; Mode = 'none' }
        @{ Name = 'lowercase hold'; Mode = 'casehold' }, @{ Name = 'unknown'; Mode = 'Unprotected' }, @{ Name = 'space'; Mode = 'None ' }
        @{ Name = 'single-element array'; Mode = @('None') }, @{ Name = 'boolean'; Mode = $false }, @{ Name = 'null'; Mode = $null }) {
        param($Name, $Mode)
        $state = [pscustomobject]@{ NativeResource = $true; Collisions = @(); AbsenceVerified = $true }
        if ($Name -ne 'missing') { $state | Add-Member -NotePropertyName ProtectionMode -NotePropertyValue $Mode }
        $findings = @(Get-Pra2Readiness -Config $config -Record $record -State $state)
        @($findings | Where-Object { $_.Code -eq 'RESOURCE_PROTECTION_UNKNOWN' -and $_.Level -eq 'Error' }).Count | Should -Be 1
        $findings.Code | Should -Not -Contain RESOURCE_UNPROTECTED
        $findings.Code | Should -Not -Contain RESOURCE_NATIVE
        $findings.Code | Should -Not -Contain ENTRA_NOT_FOUND
    }

    It 'keeps the explicit known readiness modes unchanged (<Mode>)' -TestCases @(
        @{ Mode = 'None'; Present = 'RESOURCE_UNPROTECTED' }, @{ Mode = 'CaseHold'; Present = 'RESOURCE_NATIVE' }) {
        param($Mode, $Present)
        $config.Cloud = @{ DefaultUsageLocation = 'FR' }
        $state = [pscustomobject]@{ NativeResource = $true; Collisions = @(); AbsenceVerified = $true; ProtectionMode = $Mode }
        $findings = @(Get-Pra2Readiness -Config $config -Record $record -State $state)
        $findings.Code | Should -Contain $Present
        $findings.Code | Should -Not -Contain RESOURCE_PROTECTION_UNKNOWN
    }

    It 'treats a journal without protectionMode as the legacy case-hold contract and never as unprotected' {
        $json = '{"lifecycle":"CloudResource","cloudId":"immutable-id","cloudGuid":"independent-guid","token":"0123456789abcdef0123456789abcdef","creationName":"PRA-R-01234567-89abcdef-0123456789abcdef0123456789abcdef","markerAttribute":"CustomAttribute15","licensingGroupId":"11111111-1111-1111-1111-111111111111","holdTag":"UniHlegacy"}'
        $intent = Get-Pra2ResourceIntent -Item ([pscustomobject]@{ original_json = $json })
        Get-Pra2ResourceProtectionMode -Intent $intent | Should -BeExactly CaseHold
        $bad = $json.Replace('"licensingGroupId"', '"protectionMode":"Unprotected","licensingGroupId"')
        { Get-Pra2ResourceIntent -Item ([pscustomobject]@{ original_json = $bad }) } | Should -Throw '*Unknown*'
    }

    It 'rejects inconsistent durable deletion progress in a journal intent (<Fault>)' -TestCases @(
        @{ Fault = 'Unknown'; Json = '"protectionMode":"None","unprotectedDeletionAttempted":true,"unprotectedDeletionProgress":"Done"'; Message = '*Unknown unprotected deletion progress*' }
        @{ Fault = 'ProgressArray'; Json = '"protectionMode":"None","unprotectedDeletionAttempted":true,"unprotectedDeletionProgress":["Verified"]'; Message = '*Unknown unprotected deletion progress*' }
        @{ Fault = 'AttemptType'; Json = '"protectionMode":"None","unprotectedDeletionAttempted":"true"'; Message = '*intent must be a Boolean*' }
        @{ Fault = 'NoIntent'; Json = '"protectionMode":"None","unprotectedDeletionProgress":"Verified"'; Message = '*without a deletion intent*' }
        @{ Fault = 'CaseHoldMix'; Json = '"unprotectedDeletionAttempted":true,"unprotectedDeletionProgress":"Issued"'; Message = '*Protected journal conflicts*' }) {
        param($Fault, $Json, $Message)
        $text = '{"lifecycle":"CloudResource","cloudId":"immutable-id","cloudGuid":"independent-guid","token":"0123456789abcdef0123456789abcdef","creationName":"PRA-R-01234567-89abcdef-0123456789abcdef0123456789abcdef","markerAttribute":"CustomAttribute15",' + $Json + '}'
        { Get-Pra2ResourceIntent -Item ([pscustomobject]@{ original_json = $text }) } | Should -Throw $Message
    }

    It 'derives deletion progress from the legacy attempted flag and refuses an ambiguous or completed DELETE (<Expected>)' -TestCases @(
        @{ Intent = @{}; Expected = ''; Repeatable = $true }
        @{ Intent = @{ unprotectedDeletionAttempted = $true }; Expected = 'Issuing'; Repeatable = $false }
        @{ Intent = @{ unprotectedDeletionAttempted = $true; unprotectedDeletionProgress = 'Intent' }; Expected = 'Intent'; Repeatable = $true }
        @{ Intent = @{ unprotectedDeletionAttempted = $true; unprotectedDeletionProgress = 'Issuing' }; Expected = 'Issuing'; Repeatable = $false }
        @{ Intent = @{ unprotectedDeletionAttempted = $true; unprotectedDeletionProgress = 'Issued' }; Expected = 'Issued'; Repeatable = $false }
        @{ Intent = @{ unprotectedDeletionAttempted = $true; unprotectedDeletionProgress = 'Verified' }; Expected = 'Verified'; Repeatable = $false }) {
        param($Intent, $Expected, $Repeatable)
        $object = [pscustomobject]($Intent + @{ cloudId = 'immutable-id' })
        Get-Pra2UnprotectedDeletionProgress -Intent $object | Should -BeExactly $Expected
        if ($Repeatable) { { Assert-Pra2UnprotectedDeletionRepeatable -Intent $object } | Should -Not -Throw }
        else { { Assert-Pra2UnprotectedDeletionRepeatable -Intent $object } | Should -Throw '*no second deletion*' }
    }

    # This exercises the PowerShell 7 cloud runtime (Graph HTTP metadata and URI encoding), not the 5.1 Collect path.
    Describe 'Selected resource recipient inventory' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
        BeforeAll {
            . (Join-Path $PSScriptRoot 'FakeTenant.ps1')
        }
        BeforeEach {
            New-PraFakeTenant
            $global:PraFake.Faults.RejectUnfilteredRecipientInventory = $true
            $global:PraFake.Faults.RejectUnfilteredGraphInventory = $true
        }
        AfterAll {
            foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "function:$name" -ErrorAction SilentlyContinue }
            Remove-Variable -Name PraFake, PraFakeCommands -Scope Global -ErrorAction SilentlyContinue
        }

        It 'finds unrelated <State> recipients claiming only the selected <Key> key' -TestCases @(
            foreach ($state in 'Active', 'SoftDeleted') {
                foreach ($key in 'Alias', 'Primary', 'ProxySMTP', 'ProxyX500', 'LegacyDN', 'UPN') { @{ State = $state; Key = $key } }
            }) {
            param($State, $Key)
            $record.user_principal_name = 'old-logon@contoso.com'
            $recipient = @{ Identity = 'unrelated-outside-prefix'; Name = 'unrelated-outside-prefix'; Guid = 'foreign-guid'
                RecipientTypeDetails = 'MailContact'; ExternalDirectoryObjectId = ''; ExchangeGuid = ''; Alias = 'foreign-alias'
                PrimarySmtpAddress = 'foreign@contoso.com'; EmailAddresses = @(); LegacyExchangeDN = '' }
            switch ($Key) {
                Alias { $recipient.Alias = 'ROOM' }
                Primary { $recipient.PrimarySmtpAddress = 'ROOM@CONTOSO.COM' }
                ProxySMTP { $recipient.EmailAddresses = @('SMTP:ALTERNATE@CONTOSO.COM') }
                ProxyX500 { $recipient.EmailAddresses = @('X500:/o=legacy/cn=room') }
                LegacyDN { $recipient.LegacyExchangeDN = '/o=old/cn=room' }
                UPN { $recipient.EmailAddresses = @('smtp:old-logon@contoso.com') }
            }
            if ($State -eq 'Active') { $global:PraFake.Contacts['foreign-guid'] = $recipient }
            else {
                $deleted = Add-PraFakeUser -Upn $recipient.PrimarySmtpAddress -Id 'foreign-id'
                $deleted.Alias = $recipient.Alias; $deleted.Addresses = $recipient.EmailAddresses
                $deleted['LegacyExchangeDN'] = $recipient.LegacyExchangeDN
                $deleted.Holds.Add('foreign-hold')
                $global:PraFake.Users.Remove('foreign-id')
                $global:PraFake.Recycle['foreign-id'] = $deleted
            }
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Recipients.Count | Should -Be 1
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -BeGreaterThan 0
            $global:PraFake.RecipientQueries.Count | Should -BeGreaterThan 0
            foreach ($query in $global:PraFake.RecipientQueries) {
                if ($query.Identity) {
                    ('x500:' + $query.Identity) | Should -BeIn $inventory.RecipientKeys
                    $query.Filter | Should -BeNullOrEmpty
                    $query.ResultSize | Should -BeNullOrEmpty
                } else {
                    $query.Filter | Should -Not -BeNullOrEmpty
                    $query.Filter | Should -Not -Match '-like|PRA-|LegacyExchangeDN'
                    $query.ResultSize | Should -Be Unlimited
                }
                $query.IncludeSoftDeletedRecipients | Should -BeTrue
                [string]$query.ErrorAction | Should -Be Stop
            }
        }

        It 'rejects the real unsupported LegacyExchangeDN OPATH syntax in the fake backend' {
            { Get-Recipient -Filter "LegacyExchangeDN -eq '/o=old/cn=room'" -IncludeSoftDeletedRecipients -ErrorAction Stop } |
                Should -Throw '*unsupported OPATH filter*'
        }

        It 'queries each distinct legacy/X500 key by explicit identity, including an unheld pure soft-deleted LegacyDN' {
            $deleted = Add-PraFakeUser -Upn 'unrelated@contoso.com' -Id 'deleted-legacy'
            $deleted['LegacyExchangeDN'] = '/O=OLD/CN=ROOM'
            $global:PraFake.Users.Remove($deleted.Id); $global:PraFake.Recycle[$deleted.Id] = $deleted
            $inventory = Get-Pra2ResourceInventory -Records @($record, $record)
            $queries = @($global:PraFake.RecipientQueries | Where-Object Identity)
            $queries.Count | Should -Be 2
            $queries.Identity | Should -Contain '/o=legacy/cn=room'
            $queries.Identity | Should -Contain '/o=old/cn=room'
            $inventory.Recipients.Count | Should -Be 1
            @($inventory.Recipients[0].EmailAddresses) | Should -Not -Contain 'X500:/o=old/cn=room'
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 1
        }

        It 'accepts only explicit typed or serialized not-found metadata (<Shape>)' -TestCases @(
            @{ Shape = 'Typed' }, @{ Shape = 'Wrapped' }, @{ Shape = 'FQID' }, @{ Shape = 'QualifiedFQID' }, @{ Shape = 'RemoteFQID' },
            @{ Shape = 'RemoteData' }, @{ Shape = 'RemoteDictionary' }, @{ Shape = 'RemoteWrapped' }) {
            param($Shape)
            $exception = [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException]::new('Localized absent recipient, without any English matching.')
            $fqid = 'UnrelatedTransportWrapper'
            $category = 'ObjectNotFound'
            switch ($Shape) {
                Wrapped { $exception = [InvalidOperationException]::new('Wrapped server exception.', $exception) }
                FQID { $exception = [Exception]::new('Serialized remote failure.'); $fqid = 'ManagementObjectNotFoundException,Get-Recipient' }
                QualifiedFQID { $exception = [Exception]::new('Serialized remote failure.'); $fqid = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException,Get-Recipient' }
                RemoteFQID { $exception = [Exception]::new('Serialized remote failure.'); $fqid = '[Server=EXO,RequestId=request] [FailureCategory=Cmdlet-ManagementObjectNotFoundException],Write-ErrorMessage'; $category = 'NotSpecified' }
            }
            if ($Shape -in @('RemoteData', 'RemoteDictionary', 'RemoteWrapped')) {
                $exception = [Exception]::new("Impossible d'effectuer l'opération, car l'objet est introuvable.")
                $remote = @{ TypeName = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException'; Message = $exception.Message
                    StackTrace = $null; InnerError = $null; Properties = @{ type = @{ Value = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException'; TypeAnnotation = $null } } }
                $exception.Data['RemoteException'] = if ($Shape -eq 'RemoteDictionary') { $remote } else { [pscustomobject]$remote }
                $fqid = '[],Write-ErrorMessage'; $category = 'NotSpecified'
                if ($Shape -eq 'RemoteWrapped') { $exception = [InvalidOperationException]::new('Wrapped REST failure.', $exception) }
            }
            $global:PraFake.Faults.RecipientIdentityError = [System.Management.Automation.ErrorRecord]::new($exception, $fqid, $category, '/o=old/cn=room')
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Recipients.Count | Should -Be 0
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 0
            @($global:PraFake.RecipientQueries | Where-Object Identity).Count | Should -Be 2
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'propagates generic REST errors unless guarded RemoteException.TypeName proves absence (<Failure>)' -TestCases @(
            @{ Failure = 'MissingRemote'; Category = 'NotSpecified' }
            @{ Failure = 'NullRemote'; Category = 'NotSpecified' }
            @{ Failure = 'MissingType'; Category = 'NotSpecified' }
            @{ Failure = 'NullType'; Category = 'NotSpecified' }
            @{ Failure = 'EmptyType'; Category = 'NotSpecified'; RemoteType = '' }
            @{ Failure = 'StringRemote'; Category = 'NotSpecified' }
            @{ Failure = 'ArrayType'; Category = 'NotSpecified' }
            @{ Failure = 'ShortType'; Category = 'NotSpecified'; RemoteType = 'ManagementObjectNotFoundException' }
            @{ Failure = 'PrefixType'; Category = 'NotSpecified'; RemoteType = 'OtherMicrosoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' }
            @{ Failure = 'SuffixType'; Category = 'NotSpecified'; RemoteType = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundExceptionUnexpected' }
            @{ Failure = 'CaseType'; Category = 'NotSpecified'; RemoteType = 'microsoft.exchange.configuration.tasks.managementobjectnotfoundexception' }
            @{ Failure = 'WhitespaceType'; Category = 'NotSpecified'; RemoteType = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException ' }
            @{ Failure = 'WrongType'; Category = 'NotSpecified'; RemoteType = 'System.Exception' }
            @{ Failure = 'Throttle429'; Category = 'NotSpecified'; RemoteType = 'Microsoft.Exchange.Data.Throttling.ThrottlingException' }
            @{ Failure = 'AccessDenied'; Category = 'NotSpecified'; RemoteType = 'System.UnauthorizedAccessException' }
            @{ Failure = 'Ambiguous'; Category = 'NotSpecified'; RemoteType = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectAmbiguousException' }
            @{ Failure = 'PermissionCategory'; Category = 'PermissionDenied'; RemoteType = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' }
            @{ Failure = 'ThrottleCategory'; Category = 'ResourceUnavailable'; RemoteType = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' }
            @{ Failure = 'AmbiguousCategory'; Category = 'InvalidArgument'; RemoteType = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' }) {
            param($Failure, $Category, $RemoteType)
            $exception = [Exception]::new("Impossible d'effectuer l'opération, car l'objet est introuvable.")
            $remote = [pscustomobject]@{ TypeName = $RemoteType; Properties = @{ type = @{ Value = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' } } }
            switch ($Failure) {
                MissingType { $remote.PSObject.Properties.Remove('TypeName') }
                StringRemote { $remote = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' }
                ArrayType { $remote.TypeName = @('Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException') }
            }
            if ($Failure -ne 'MissingRemote') { $exception.Data['RemoteException'] = if ($Failure -eq 'NullRemote') { $null } else { $remote } }
            $global:PraFake.Faults.RecipientIdentityError = [System.Management.Automation.ErrorRecord]::new($exception, '[],Write-ErrorMessage', $Category, '/o=old/cn=room')
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
            @($global:PraFake.RecipientQueries | Where-Object Identity).Count | Should -Be 1
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'never interprets another legacy identity failure as absence (<Failure>)' -TestCases @(
            @{ Failure = 'Permission'; Category = 'PermissionDenied'; FQID = 'ManagementObjectNotFoundException'; Message = 'Access denied.' }
            @{ Failure = 'Transient429'; Category = 'ResourceUnavailable'; FQID = 'TooManyRequests'; Message = '429 throttled.' }
            @{ Failure = 'Transient503'; Category = 'NotSpecified'; FQID = 'ServiceUnavailable'; Message = '503 retry later.' }
            @{ Failure = 'Ambiguous'; Category = 'InvalidArgument'; FQID = 'ManagementObjectAmbiguousException'; Message = 'Ambiguous legacy identity.' }
            @{ Failure = 'MessageOnly'; Category = 'ObjectNotFound'; FQID = 'GenericFailure'; Message = 'ManagementObjectNotFoundException could not be found.' }
            @{ Failure = 'Generic404'; Category = 'ObjectNotFound'; FQID = 'HttpFailure'; Message = '404 endpoint not found.' }
            @{ Failure = 'FQIDSuffix'; Category = 'ObjectNotFound'; FQID = 'ManagementObjectNotFoundExceptionUnexpected'; Message = 'Other exception.' }
            @{ Failure = 'FQIDPrefix'; Category = 'ObjectNotFound'; FQID = 'OtherManagementObjectNotFoundException'; Message = 'Other exception.' }
            @{ Failure = 'TypedWrongCategory'; Category = 'PermissionDenied'; FQID = 'ManagementObjectNotFoundException'; Message = 'Access denied.' }) {
            param($Failure, $Category, $FQID, $Message)
            $exception = if ($Failure -eq 'TypedWrongCategory') { [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException]::new($Message) } else { [Exception]::new($Message) }
            $global:PraFake.Faults.RecipientIdentityError = [System.Management.Automation.ErrorRecord]::new($exception, $FQID, $Category, '/o=old/cn=room')
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
            @($global:PraFake.RecipientQueries | Where-Object Identity).Count | Should -Be 1
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'accepts a successful empty exact identity result without an unfiltered fallback' {
            $global:PraFake.Faults.RecipientIdentityEmpty = $true
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Recipients.Count | Should -Be 0
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 0
            @($global:PraFake.RecipientQueries | Where-Object { -not $_.Filter -and -not $_.Identity }).Count | Should -Be 0
        }

        It 'keeps three selected resources primary SMTP distinct from their display UPN and reports failed inventories per object (<Fail>)' -TestCases @(
            @{ Fail = $false }, @{ Fail = $true }) {
            param($Fail)
            $null = Add-PraFakeUser -Upn 'sender-a@contoso.com'
            $null = Add-PraFakeUser -Upn 'sender-b@contoso.com'
            $records = @(foreach ($name in 'equipment', 'room-a', 'room-b') {
                    $copy = $record | Select-Object *
                    $copy.object_guid = "source-$name"; $copy.kind = $(if ($name -eq 'equipment') { 'Equipment' } else { 'Room' })
                    $copy.alias = "fixture-$name"; $copy.primary_smtp_address = "fixture-$name@tenant.onmicrosoft.com"
                    $copy.user_principal_name = "fixture-$name@source.example"; $copy.legacy_exchange_dn = "/o=Source/cn=fixture-$name"
                    $copy.email_addresses_json = ConvertTo-Json -InputObject @("SMTP:$($copy.primary_smtp_address)") -Compress
                    $copy
                })
            if ($Fail) {
                $global:PraFake.Faults.RecipientIdentityError = [System.Management.Automation.ErrorRecord]::new(
                    [Exception]::new('Legacy service unavailable.'), 'ServiceUnavailable', 'ResourceUnavailable', '/o=Source/cn=fixture-equipment')
            }
            $states = Get-Pra2CloudStateSet -Context @{ Config = $config } -Records $records
            $states.Count | Should -Be 3
            foreach ($row in $records) {
                $row.primary_smtp_address | Should -Match '@tenant.onmicrosoft.com$'
                $row.user_principal_name | Should -Match '@source\.example$'
                if ($Fail) { $states[$row.object_guid] | Should -BeOfType ([System.Management.Automation.ErrorRecord]) }
                else {
                    $states[$row.object_guid].AbsenceVerified | Should -BeTrue
                    $states[$row.object_guid].Collisions.Count | Should -Be 0
                    @($global:PraFake.RecipientQueries | Where-Object { $_.Identity -ieq $row.legacy_exchange_dn }).Count | Should -Be 1
                }
            }
            $global:PraFake.Users.Count | Should -Be 2
            $global:PraFake.Contacts.Count | Should -Be 0
            $global:PraFake.RoomLists.Count | Should -Be 0
            $global:PraFake.Calls.Count | Should -Be 0
            if (-not $Fail) {
                $activeQueries = @($global:PraFake.GraphQueries | Where-Object Path -eq users)
                $activeQueries.Count | Should -BeGreaterThan 0
                $activeQueries.Count | Should -BeLessThan 10
                @($activeQueries | Where-Object { $_.Headers.Count -eq 0 }).Count | Should -Be 2
                @($activeQueries | Where-Object { $_.Headers['ConsistencyLevel'] -ceq 'eventual' }).Count | Should -Be 1
                foreach ($query in $activeQueries) {
                    $query.Filter | Should -Not -BeNullOrEmpty
                    $query.Filter.Length | Should -BeLessOrEqual 2000
                    [uri]::EscapeDataString($query.Filter).Length | Should -BeLessOrEqual 3000
                }
                @($global:PraFake.GraphQueries | Where-Object Path -eq 'directory/deletedItems/microsoft.graph.user').Count | Should -Be 1
            }
        }

        It 'retains distinct exact legacy claims for collision detection and journal ownership exclusions' {
            foreach ($id in 'first', 'second') {
                $global:PraFake.Contacts[$id] = @{ Guid = $id; Identity = $id; RecipientTypeDetails = 'MailContact'; Alias = $id; PrimarySmtpAddress = "$id@contoso.com"
                    EmailAddresses = @(); LegacyExchangeDN = '/o=old/cn=room' }
            }
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Recipients.Count | Should -Be 2
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 2
            $collisions = @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory -OwnedContactId first)
            $collisions.Count | Should -Be 1
            $collisions[0] | Should -Match second
        }

        It 'fails closed if identity returns non-exact or partial failure results (<Result>)' -TestCases @(
            @{ Result = 'Wrong' }, @{ Result = 'PartialNotFound' }, @{ Result = 'PartialRemoteNotFound' }) {
            param($Result)
            Mock Get-Recipient -ModuleName PRA2.Cloud {
                if (-not $Identity) { return }
                [pscustomobject]@{ Guid = 'foreign'; Identity = 'foreign'; EmailAddresses = @(); LegacyExchangeDN = $(if ($Result -eq 'Wrong') { '/o=foreign/cn=room' } else { $Identity }) }
                if ($Result -in @('PartialNotFound', 'PartialRemoteNotFound')) {
                    $exception = [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException]::new('Partial read.')
                    $fqid = 'ManagementObjectNotFoundException'; $category = 'ObjectNotFound'
                    if ($Result -eq 'PartialRemoteNotFound') {
                        $exception = [Exception]::new("Impossible d'effectuer l'opération, car l'objet est introuvable.")
                        $exception.Data['RemoteException'] = [pscustomobject]@{ TypeName = 'Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException' }
                        $fqid = '[],Write-ErrorMessage'; $category = 'NotSpecified'
                    }
                    Write-Error -ErrorRecord ([System.Management.Automation.ErrorRecord]::new($exception, $fqid, $category, $Identity)) -ErrorAction Stop
                }
            }
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
        }

        It 'keeps all active and deleted Graph pages for UPN-only and immutable-anchor collisions' {
            $global:PraFake.PageSize = 1
            $null = Add-PraFakeUser -Upn 'irrelevant@contoso.com'
            $record.user_principal_name = 'old-logon@contoso.com'
            $user = Add-PraFakeUser -Upn $record.user_principal_name -Id 'upn-only'
            $user.Alias = 'foreign'; $user.Addresses = @()
            $deleted = Add-PraFakeUser -Upn 'deleted@contoso.com' -Id 'immutable-only' -ImmutableId $record.immutable_id
            $global:PraFake.Users.Remove($deleted.Id)
            $global:PraFake.Recycle[$deleted.Id] = $deleted
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Users.Count | Should -Be 2
            $collisions = @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory)
            ($collisions -join ';') | Should -Match 'upn-only'
            ($collisions -join ';') | Should -Match 'immutable-only'
        }

        It 'safely quotes apostrophes and treats wildcard and OPATH text as exact values' {
            $record.alias = "O'Brien* -or Alias -eq 'foreign'"
            $record.primary_smtp_address = "o'brien@contoso.com"
            $record.user_principal_name = "old'o@contoso.com"
            $record.email_addresses_json = '["X500:/o=Legacy/cn=O''Brien*(test)"]'
            $record.legacy_exchange_dn = "/o=Old/cn=O'Brien*(test)"
            $global:PraFake.Contacts['exact'] = @{ Guid = 'exact'; Identity = 'exact'; Alias = $record.alias
                PrimarySmtpAddress = $record.primary_smtp_address; EmailAddresses = @("X500:/o=Legacy/cn=O'Brien*(test)"); LegacyExchangeDN = $record.legacy_exchange_dn }
            $global:PraFake.Contacts['foreign'] = @{ Guid = 'foreign'; Identity = 'foreign'; Alias = 'foreign'; PrimarySmtpAddress = 'foreign@contoso.com'; EmailAddresses = @() }
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Recipients.Count | Should -Be 1
            $inventory.Recipients[0].Guid | Should -Be exact
            $global:PraFake.RecipientQueries[0].Filter | Should -Match "o''brien"
            @($global:PraFake.RecipientQueries | Where-Object Identity).Identity | Should -Contain "/o=old/cn=o'brien*(test)"
        }

        It 'deduplicates shared target keys and chunk results but not distinct recipients with duplicate claims' {
            $records = @($record, $record) + @(1..35 | ForEach-Object {
                    [pscustomobject]@{ primary_smtp_address = "resource$_@contoso.com"; alias = "resource$_"; user_principal_name = ''
                        email_addresses_json = '[]'; legacy_exchange_dn = '' }
                })
            foreach ($id in 'first', 'second') {
                $global:PraFake.Contacts[$id] = @{ Guid = $id; Identity = $id; RecipientTypeDetails = 'MailContact'; Alias = 'room'; PrimarySmtpAddress = 'foreign@contoso.com'
                    EmailAddresses = @('smtp:alternate@contoso.com', 'smtp:resource35@contoso.com') }
            }
            $inventory = Get-Pra2ResourceInventory -Records $records
            $inventory.Recipients.Count | Should -Be 2
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 2
            $queries = @($global:PraFake.RecipientQueries | Where-Object Filter)
            $queries.Count | Should -BeGreaterThan 1
            $allClauses = @(foreach ($query in $queries) {
                    $query.Filter.Length | Should -BeLessOrEqual 4000
                    $query.Filter | Should -Not -Match 'LegacyExchangeDN'
                    $matches = [regex]::Matches($query.Filter, "(Alias|EmailAddresses|PrimarySmtpAddress) -eq '((?:[^']|'')*)'")
                    $matches.Count | Should -BeGreaterThan 0
                    $matches.Count | Should -BeLessOrEqual 50
                    foreach ($match in $matches) { $match.Value }
                })
            @($allClauses | Select-Object -Unique).Count | Should -Be $allClauses.Count
            foreach ($other in $records) { { Get-Pra2ResourceCollisions -Record $other -Inventory $inventory } | Should -Not -Throw }
        }

        It 'bounds long X500 filters by characters as well as clause count' {
            $addresses = @(1..20 | ForEach-Object { 'X500:/o=legacy/cn=' + ('a' * 350) + $_ })
            $record.email_addresses_json = ConvertTo-Json -InputObject $addresses -Compress
            $null = Get-Pra2ResourceInventory -Records @($record)
            $global:PraFake.RecipientQueries.Count | Should -BeGreaterThan 1
            foreach ($query in @($global:PraFake.RecipientQueries | Where-Object Filter)) { $query.Filter.Length | Should -BeLessOrEqual 4000 }
        }

        It 'never emits an empty first chunk when one exact Graph clause reaches the character limit' {
            $record.alias = 'a' * (2000 - "(mailNickname eq '')".Length)
            $null = Get-Pra2ResourceInventory -Records @($record)
            @($global:PraFake.GraphQueries | Where-Object { $_.Filter.Length -eq 2000 }).Count | Should -Be 1
            foreach ($query in @($global:PraFake.RecipientQueries | Where-Object Filter)) {
                $query.Filter | Should -Not -BeNullOrEmpty
                $query.Filter.Length | Should -BeLessOrEqual 4000
            }
        }

        It 'does not return partial inventory when a later chunk cannot be read' {
            $record.email_addresses_json = ConvertTo-Json -InputObject @(1..60 | ForEach-Object { "smtp:proxy$_@contoso.com" }) -Compress
            $global:PraFake.Faults.RecipientQueryFailsAt = 2
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw '*filtered query access denied*'
            $global:PraFake.RecipientQueries.Count | Should -Be 2
        }

        It 'fails closed on unreadable Graph after successful exact Exchange queries' {
            $global:PraFake.Faults.GraphReadFails = $true
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
            $global:PraFake.RecipientQueries.Count | Should -BeGreaterThan 0
        }

        It 'fails closed on missing, empty, null or invalid selections before any query (<Selection>)' -TestCases @(
            @{ Selection = 'Missing' }, @{ Selection = 'Empty' }, @{ Selection = 'Null' }, @{ Selection = 'Invalid' }, @{ Selection = 'Oversize' }, @{ Selection = 'EmptyLegacy' }) {
            param($Selection)
            switch ($Selection) {
                Missing { { Get-Pra2ResourceInventory } | Should -Throw '*requires selected records*' }
                Empty { { Get-Pra2ResourceInventory -Records @() } | Should -Throw '*requires selected records*' }
                Null { { Get-Pra2ResourceInventory -Records @($null) } | Should -Throw '*empty selected record*' }
                Invalid { $record.primary_smtp_address = 'invalid'; { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw '*invalid*' }
                Oversize { $record.alias = 'a' * 4001; { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw '*bounded recipient filter size*' }
                EmptyLegacy { $record.email_addresses_json = '["X500:"]'; { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw '*legacy identity is empty*' }
            }
            $global:PraFake.RecipientQueries.Count | Should -Be 0
        }

        It 'refuses using a targeted inventory to prove absence for an unqueried record' {
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $record.alias = 'another-alias'
            { Get-Pra2ResourceCollisions -Record $record -Inventory $inventory } | Should -Throw '*does not cover all selected record keys*'
        }

        It 'queries room-list record addresses and alias rather than the tenant or resource-member scope' {
            $list = [pscustomobject]@{ SourceGuid = 'list-source'; Alias = 'building'; PrimarySmtpAddress = 'building@contoso.com'
                EmailAddresses = @('SMTP:building@contoso.com', 'smtp:building-proxy@contoso.com'); LegacyExchangeDN = '/o=legacy/cn=building' }
            $global:PraFake.Contacts['list-collision'] = @{ Guid = 'list-collision'; Identity = 'foreign'; RecipientTypeDetails = 'MailContact'; Alias = 'foreign'
                PrimarySmtpAddress = 'foreign@contoso.com'; EmailAddresses = @('smtp:building-proxy@contoso.com') }
            { Get-Pra2ResourceCloudRoomList -List $list -AllowMissing } | Should -Throw '*Room-list collision*'
            $filter = $global:PraFake.RecipientQueries[0].Filter
            $filter | Should -Match "Alias -eq 'building'"
            $filter | Should -Match 'building-proxy@contoso.com'
            $filter | Should -Match '/o=legacy/cn=building'
            $filter | Should -Not -Match 'room@contoso.com'
        }

        It 'preserves journal-owned immutable ID, GUID and contact exclusions without excluding competing claims' {
            $global:PraFake.Contacts['owned-contact'] = @{ Guid = 'owned-contact'; Identity = 'owned-contact'; RecipientTypeDetails = 'MailContact'
                Alias = 'room'; PrimarySmtpAddress = 'room@contoso.com'; EmailAddresses = @() }
            $owned = Add-PraFakeUser -Upn 'room@contoso.com' -Id 'owned-id'
            $owned.ExchangeGuid = 'owned-guid'
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory -OwnedId 'owned-id' -OwnedGuid 'owned-guid' -OwnedContactId 'owned-contact').Count | Should -Be 0
            $global:PraFake.Contacts['competing'] = @{ Guid = 'competing'; Identity = 'competing'; RecipientTypeDetails = 'MailContact'
                Alias = 'room'; PrimarySmtpAddress = 'foreign@contoso.com'; EmailAddresses = @() }
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $collisions = @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory -OwnedId 'owned-id' -OwnedGuid 'owned-guid' -OwnedContactId 'owned-contact')
            $collisions.Count | Should -Be 1
            $collisions[0] | Should -Match 'competing'
        }

        It 'does not turn incomplete <Path> Graph pages into absence (<Fault>)' -TestCases @(
            foreach ($path in 'users', 'directory/deletedItems/microsoft.graph.user') {
                foreach ($fault in 'MissingValue', 'NullValue', 'ScalarValue', 'LimitedIdentity', 'NextPageFails') { @{ Path = $path; Fault = $fault } }
            }) {
            param($Path, $Fault)
            Mock Invoke-Pra2Graph -ModuleName PRA2.Cloud {
                if ($Uri -match '[?&]fail=true') { throw 'Graph next page access denied.' }
                if ($Uri -like "https://graph.microsoft.com/v1.0/$Path`?*") {
                    switch ($Fault) {
                        MissingValue { return [pscustomobject]@{ other = @() } }
                        NullValue { return [pscustomobject]@{ value = $null } }
                        ScalarValue { return [pscustomobject]@{ value = [pscustomobject]@{ id = 'not-a-collection' } } }
                        LimitedIdentity { return [pscustomobject]@{ value = @([pscustomobject]@{ id = ''; userPrincipalName = 'hidden@contoso.com' }) } }
                        NextPageFails { return [pscustomobject]@{ value = @(); '@odata.nextLink' = "https://graph.microsoft.com/v1.0/$Path`?fail=true" } }
                    }
                }
                return [pscustomobject]@{ value = @() }
            }
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
        }

        It 'finds an unknown active user outside naming prefixes by only <Key>' -TestCases @(
            foreach ($key in 'UPNPrimary', 'UPNProxy', 'UPNCollected', 'MailPrimary', 'MailProxy', 'MailCollected', 'Alias',
                'SMTPPrimary', 'smtpSecondary', 'SMTPSecondary', 'X500Proxy', 'x500Proxy', 'Legacy', 'OriginalProxy', 'Anchor') { @{ Key = $key } }) {
            param($Key)
            $record.user_principal_name = 'old-logon@contoso.com'
            $record.immutable_id = 'AbC+/Def=='
            $record.email_addresses_json = '["SMTP:room@contoso.com","smtp:alternate@contoso.com","X500:/O=Legacy/CN=Room"]'
            $user = Add-PraFakeUser -Upn 'unknown-outside-prefix@contoso.com' -Id 'unknown-active' -ImmutableId 'different-anchor'
            $user.Mail = 'other@contoso.com'; $user.Alias = 'foreign'; $user.Addresses = @()
            switch ($Key) {
                UPNPrimary { $user.Upn = 'ROOM@CONTOSO.COM' }
                UPNProxy { $user.Upn = 'ALTERNATE@CONTOSO.COM' }
                UPNCollected { $user.Upn = 'OLD-LOGON@CONTOSO.COM' }
                MailPrimary { $user.Mail = 'ROOM@CONTOSO.COM' }
                MailProxy { $user.Mail = 'ALTERNATE@CONTOSO.COM' }
                MailCollected { $user.Mail = 'OLD-LOGON@CONTOSO.COM' }
                Alias { $user.Alias = 'ROOM' }
                SMTPPrimary { $user.Addresses = @('SMTP:ROOM@CONTOSO.COM') }
                smtpSecondary { $user.Addresses = @('smtp:ALTERNATE@CONTOSO.COM') }
                SMTPSecondary { $user.Addresses = @('SMTP:ALTERNATE@CONTOSO.COM') }
                X500Proxy { $user.Addresses = @('X500:/O=LEGACY/CN=ROOM') }
                x500Proxy { $user.Addresses = @('x500:/O=LEGACY/CN=ROOM') }
                Legacy { $user.Addresses = @('X500:/O=OLD/CN=ROOM') }
                OriginalProxy { $user.Addresses = @('X500:/O=Legacy/CN=Room') }
                Anchor { $user.ImmutableId = $record.immutable_id }
            }
            1..30 | ForEach-Object { $null = Add-PraFakeUser -Upn "unrelated$_@contoso.com" }
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Users.Count | Should -Be 1
            $inventory.Users[0].id | Should -Be unknown-active
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory | Where-Object { $_ -like 'Existing Entra*' }).Count | Should -Be 1
            foreach ($query in @($global:PraFake.GraphQueries | Where-Object Path -eq users)) {
                $query.Filter | Should -Not -BeNullOrEmpty
                $query.Filter | Should -Not -Match 'startswith|endswith|displayName|PRA-|\$count|eventual'
                $query.Filter.Length | Should -BeLessOrEqual 2000
                [uri]::EscapeDataString($query.Filter).Length | Should -BeLessOrEqual 3000
                $query.Query | Should -Match '&\$filter=%28'
            }
            if ($Key -eq 'OriginalProxy') {
                ($global:PraFake.GraphQueries.Filter -join ' ') | Should -Match ([regex]::Escape("p eq 'X500:/O=Legacy/CN=Room'"))
            }
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'keeps immutable anchors case-sensitive and rejects reuse of an inventory for another anchor' {
            $record.immutable_id = 'AbC+/Def=='
            $wrong = Add-PraFakeUser -Upn 'wrong@contoso.com' -Id 'wrong-anchor' -ImmutableId 'abc+/def=='
            $right = Add-PraFakeUser -Upn 'right@contoso.com' -Id 'right-anchor' -ImmutableId $record.immutable_id
            $deleted = Add-PraFakeUser -Upn 'deleted@contoso.com' -Id 'deleted-wrong-anchor' -ImmutableId $wrong.ImmutableId
            $global:PraFake.Users.Remove($deleted.Id); $global:PraFake.Recycle[$deleted.Id] = $deleted
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Users.Count | Should -Be 2
            $inventory.Users.id | Should -Not -Contain wrong-anchor
            $collisions = @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory)
            $collisions.Count | Should -Be 1
            $collisions[0] | Should -Match right-anchor
            $record.immutable_id = $wrong.ImmutableId
            { Get-Pra2ResourceCollisions -Record $record -Inventory $inventory } | Should -Throw '*does not cover the selected immutable anchor*'
        }

        It 'quotes and encodes exact Graph values without interpreting apostrophes, ampersands, plus or query text' {
            $record.alias = "O'Brien*' or mail eq 'foreign"
            $record.primary_smtp_address = "o'brien+&@contoso.com"
            $record.user_principal_name = "old'o+&@contoso.com"
            $record.email_addresses_json = '["X500:/o=Legacy/cn=O''Brien*(test)&$filter=mail eq ''foreign''"]'
            $record.immutable_id = "AbC+/'&?=="
            $exact = Add-PraFakeUser -Upn 'exact@contoso.com' -Id 'exact' -ImmutableId $record.immutable_id
            $null = Add-PraFakeUser -Upn 'foreign@contoso.com' -Id 'foreign'
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Users.Count | Should -Be 1
            $inventory.Users[0].id | Should -Be exact
            $filters = $global:PraFake.GraphQueries.Filter -join ' '
            $filters | Should -Match ([regex]::Escape("mailNickname eq 'o''brien*'' or mail eq ''foreign'"))
            $filters | Should -Match ([regex]::Escape("onPremisesImmutableId eq 'AbC+/''&?=='"))
            $queries = $global:PraFake.GraphQueries.Query -join ' '
            $queries | Should -Match '%2B'; $queries | Should -Match '%26'; $queries | Should -Match '%3F'
        }

        It 'faithfully parses Graph AND/OR grouping and preserves the exact filter on every page' {
            $first = Add-PraFakeUser -Upn 'first@contoso.com' -Id first -ImmutableId 'AbC=='
            $first.Alias = "o'brien"
            $second = Add-PraFakeUser -Upn 'second@contoso.com' -Id second -ImmutableId 'AbC=='
            $second.Mail = 'claim@contoso.com'
            $wrong = Add-PraFakeUser -Upn 'wrong@contoso.com' -Id wrong -ImmutableId 'abc=='
            $wrong.Mail = 'claim@contoso.com'
            $global:PraFake.PageSize = 1
            $filter = "(mailNickname eq 'o''brien' or mail eq 'claim@contoso.com') and onPremisesImmutableId eq 'AbC=='"
            $uri = 'https://graph.microsoft.com/v1.0/users?$select=id&$top=999&$filter=' + [uri]::EscapeDataString($filter)
            $page = Invoke-MgGraphRequest -Uri $uri
            @($page.value).Count | Should -Be 1
            $page.value[0].id | Should -Be first
            $next = Invoke-MgGraphRequest -Uri $page.'@odata.nextLink'
            @($next.value).Count | Should -Be 1
            $next.value[0].id | Should -Be second
            $next.PSObject.Properties.Name | Should -Not -Contain '@odata.nextLink'
            $global:PraFake.GraphQueries[0].Filter | Should -BeExactly $filter
            $global:PraFake.GraphQueries[1].Filter | Should -BeExactly $filter
            { ConvertFrom-PraFakeGraphFilter "startswith(mail,'prefix')" } | Should -Throw '*unsupported exact filter*'
            { ConvertFrom-PraFakeGraphFilter "mail eq 'a' or" } | Should -Throw
        }

        It 'segregates exact non-SMTP OR predicates from default SMTP, UPN, alias and anchor queries' {
            $record.email_addresses_json = '["SMTP:room@contoso.com","smtp:alternate@contoso.com","X500:/O=Legacy/CN=Room","SIP:room@contoso.com"]'
            $record.user_principal_name = 'old-logon@contoso.com'
            $null = Get-Pra2ResourceInventory -Records @($record)
            $active = @($global:PraFake.GraphQueries | Where-Object Path -eq users)
            $default = @($active | Where-Object { $_.Headers.Count -eq 0 })
            $advanced = @($active | Where-Object { $_.Headers.Count -gt 0 })
            $default.Count | Should -BeGreaterThan 0
            $advanced.Count | Should -Be 1
            foreach ($query in $default) {
                $query.Query | Should -Not -Match '\$count|eventual'
                $query.Filter | Should -Not -Match "p eq '(?:x500|sip):"
            }
            ($default.Filter -join ' ') | Should -Match 'userPrincipalName eq'
            ($default.Filter -join ' ') | Should -Match 'mail eq'
            ($default.Filter -join ' ') | Should -Match 'mailNickname eq'
            ($default.Filter -join ' ') | Should -Match 'onPremisesImmutableId eq'
            ($default.Filter -join ' ') | Should -Match "p eq 'SMTP:"
            $advanced[0].Headers.ConsistencyLevel | Should -BeExactly eventual
            $advanced[0].Query | Should -Match '&\$count=true(?:&|$)'
            $advanced[0].Filter | Should -Match ' or '
            $advanced[0].Filter | Should -Match "p eq 'X500:/O=Legacy/CN=Room'"
            $advanced[0].Filter | Should -Match "p eq 'x500:/O=Legacy/CN=Room'"
            $advanced[0].Filter | Should -Match "p eq 'SIP:room@contoso.com'"
            $advanced[0].Filter | Should -Not -Match "mailNickname|userPrincipalName|onPremisesImmutableId|mail eq|p eq 'smtp:"
            $deleted = @($global:PraFake.GraphQueries | Where-Object Path -ne users)
            $deleted.Count | Should -Be 1
            $deleted[0].Headers.Count | Should -Be 0
            $deleted[0].Query | Should -Not -Match '\$count'
        }

        It 'leaves SMTP-only inventory and ordinary Graph callers in default mode' {
            $record.email_addresses_json = '["SMTP:room@contoso.com","smtp:alternate@contoso.com"]'
            $record.legacy_exchange_dn = ''
            $null = Get-Pra2ResourceInventory -Records @($record)
            $active = @($global:PraFake.GraphQueries | Where-Object Path -eq users)
            $active.Count | Should -Be 1
            $active[0].Headers.Count | Should -Be 0
            $active[0].Query | Should -Not -Match '\$count'
            $null = Invoke-Pra2Graph -Uri "v1.0/users?`$filter=mailNickname%20eq%20'room'"
            $global:PraFake.GraphQueries[-1].Headers.Count | Should -Be 0
            $global:PraFake.GraphQueries[-1].Query | Should -Not -Match '\$count'
        }

        It 'forwards optional dictionary headers and retains advanced count/header on every X500 page' {
            $global:PraFake.PageSize = 1
            foreach ($id in 'first-x500', 'second-x500') {
                $user = Add-PraFakeUser -Upn "$id@contoso.com" -Id $id
                $user.Alias = $id; $user.Addresses = @('X500:/O=LEGACY/CN=ROOM')
            }
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Users.Count | Should -Be 2
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory | Where-Object { $_ -like 'Existing Entra*' }).Count | Should -Be 2
            $advanced = @($global:PraFake.GraphQueries | Where-Object { $_.Headers.Count -gt 0 })
            $advanced.Count | Should -Be 2
            $advanced[1].Query | Should -Match '\$skiptoken=1'
            $advanced[1].Filter | Should -BeExactly $advanced[0].Filter
            foreach ($query in $advanced) {
                $query.Headers.ConsistencyLevel | Should -BeExactly eventual
                $query.Query | Should -Match '&\$count=true(?:&|$)'
            }
            $filter = "proxyAddresses/any(p:p eq 'X500:/O=Legacy/CN=Room')"
            $headers = [System.Collections.Generic.Dictionary[string,object]]::new()
            $headers.Add('ConsistencyLevel', 'eventual')
            $page = Invoke-Pra2Graph -Uri ('https://graph.microsoft.com/v1.0/users?$count=true&$filter=' + [uri]::EscapeDataString($filter)) -Headers $headers
            $page.'@odata.count' | Should -Be 2
            $global:PraFake.GraphQueries[-1].Headers.ConsistencyLevel | Should -BeExactly eventual
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'rejects non-SMTP equality without both advanced requirements (<Fault>)' -TestCases @(
            @{ Fault = 'Default' }, @{ Fault = 'CountOnly' }, @{ Fault = 'HeaderOnly' },
            @{ Fault = 'FalseCount' }, @{ Fault = 'WrongHeader' }) {
            param($Fault)
            $filter = "(mailNickname eq 'room') or (proxyAddresses/any(p:p eq 'X500:/o=legacy/cn=room'))"
            $uri = 'https://graph.microsoft.com/v1.0/users?$filter=' + [uri]::EscapeDataString($filter)
            $parameters = @{ Uri = $uri }
            if ($Fault -in @('CountOnly', 'WrongHeader')) { $parameters.Uri += '&$count=true' }
            if ($Fault -eq 'FalseCount') { $parameters.Uri += '&$count=false' }
            if ($Fault -in @('HeaderOnly', 'FalseCount')) { $parameters.Headers = @{ ConsistencyLevel = 'eventual' } }
            if ($Fault -eq 'WrongHeader') { $parameters.Headers = @{ ConsistencyLevel = 'strong' } }
            { Invoke-Pra2Graph @parameters } | Should -Throw '*Request_UnsupportedQuery*requires*'
            $global:PraFake.GraphQueries.Count | Should -Be 1
        }

        It 'fails closed on advanced <Status> errors on <Page> without fallback or mutation' -TestCases @(
            foreach ($status in 400, 403, 429, 503) {
                foreach ($page in 'First', 'Next') { @{ Status = $status; Page = $page } }
            }) {
            param($Status, $Page)
            Mock Start-Sleep -ModuleName PRA2.Cloud { }
            $global:PraFake.Faults.GraphAdvancedStatus = $Status
            if ($Page -eq 'Next') {
                $global:PraFake.PageSize = 1
                foreach ($id in 'one', 'two') {
                    $user = Add-PraFakeUser -Upn "$id@contoso.com" -Id $id
                    $user.Addresses = @('X500:/o=legacy/cn=room')
                }
                $global:PraFake.Faults.GraphAdvancedFailsAt = 3
            }
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
            $attempts = if ($Status -in @(429, 503)) { 5 } else { 1 }
            $global:PraFake.GraphQueries.Count | Should -Be ($attempts + $(if ($Page -eq 'Next') { 2 } else { 1 }))
            @($global:PraFake.GraphQueries | Where-Object { $_.Headers.Count -gt 0 }).Count | Should -Be ($attempts + $(if ($Page -eq 'Next') { 1 } else { 0 }))
            foreach ($query in $global:PraFake.GraphQueries) {
                $query.Filter | Should -Not -BeNullOrEmpty
                if ($query.Headers.Count) {
                    $query.Headers.ConsistencyLevel | Should -BeExactly eventual
                    $query.Query | Should -Match '\$count=true'
                }
            }
            Should -Invoke Start-Sleep -ModuleName PRA2.Cloud -Times $(if ($Status -in @(429, 503)) { 4 } else { 0 }) -Exactly
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'requires the complete advanced value collection, not count metadata (<Fault>)' -TestCases @(
            @{ Fault = 'CountOnly' }, @{ Fault = 'NullValue' }, @{ Fault = 'ScalarValue' },
            @{ Fault = 'LimitedIdentity' }, @{ Fault = 'NextPageFails' }) {
            param($Fault)
            Mock Invoke-Pra2Graph -ModuleName PRA2.Cloud {
                param($Uri, $Headers)
                if (-not $Headers) { return [pscustomobject]@{ value = @() } }
                if ($Uri -match '[?&]fail=true') { throw 'Advanced Graph next page access denied.' }
                switch ($Fault) {
                    CountOnly { return [pscustomobject]@{ '@odata.count' = 0 } }
                    NullValue { return [pscustomobject]@{ '@odata.count' = 0; value = $null } }
                    ScalarValue { return [pscustomobject]@{ '@odata.count' = 0; value = [pscustomobject]@{ id = 'claim' } } }
                    LimitedIdentity { return [pscustomobject]@{ '@odata.count' = 1; value = @([pscustomobject]@{ id = 'claim'; userPrincipalName = 'other@contoso.com' }) } }
                    NextPageFails { return [pscustomobject]@{ '@odata.count' = 0; value = @(); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users?fail=true' } }
                }
            }
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
            Should -Invoke Invoke-Pra2Graph -ModuleName PRA2.Cloud -Times $(if ($Fault -eq 'NextPageFails') { 2 } else { 1 }) -Exactly -ParameterFilter {
                $Headers -and $Headers['ConsistencyLevel'] -ceq 'eventual'
            }
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'does not let optional <Metadata> count metadata hide an exact X500 collision' -TestCases @(
            @{ Metadata = 'Missing' }, @{ Metadata = 'Zero' }) {
            param($Metadata)
            $user = Add-PraFakeUser -Upn 'outside@contoso.com' -Id 'x500-only'
            $user.Addresses = @('X500:/o=legacy/cn=room')
            $global:PraGraphX500User = ConvertTo-PraFakeGraphUser $user
            Mock Invoke-Pra2Graph -ModuleName PRA2.Cloud {
                param($Uri, $Headers)
                if (-not $Headers) { return [pscustomobject]@{ value = @() } }
                $page = [ordered]@{ value = @($global:PraGraphX500User) }
                if ($Metadata -eq 'Zero') { $page['@odata.count'] = 0 }
                return [pscustomobject]$page
            }
            try {
                $inventory = Get-Pra2ResourceInventory -Records @($record)
                $inventory.Users.Count | Should -Be 1
                @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory | Where-Object { $_ -like 'Existing Entra*' }).Count | Should -Be 1
            } finally { Remove-Variable PraGraphX500User -Scope Global }
        }

        It 'deduplicates Graph users across bounded OR chunks and pages without losing distinct duplicate claims' {
            $global:PraFake.PageSize = 1
            $records = @($record, $record) + @(1..35 | ForEach-Object {
                    [pscustomobject]@{ primary_smtp_address = "resource$_@contoso.com"; alias = "resource$_"; user_principal_name = ''
                        email_addresses_json = '[]'; legacy_exchange_dn = ''; immutable_id = "AbC$_==" }
                })
            foreach ($id in 'first', 'second') {
                $user = Add-PraFakeUser -Upn "$id@contoso.com" -Id $id -ImmutableId 'AbC35=='
                $user.Alias = 'room'; $user.Addresses = @('SMTP:alternate@contoso.com', 'smtp:resource35@contoso.com')
            }
            $deleted = Add-PraFakeUser -Upn 'deleted@contoso.com' -Id deleted -ImmutableId 'source-anchor'
            $global:PraFake.Users.Remove($deleted.Id); $global:PraFake.Recycle[$deleted.Id] = $deleted
            $inventory = Get-Pra2ResourceInventory -Records $records
            $inventory.Users.Count | Should -Be 3
            $inventory.Users.id | Should -Contain first; $inventory.Users.id | Should -Contain second; $inventory.Users.id | Should -Contain deleted
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory | Where-Object { $_ -like 'Existing Entra*' }).Count | Should -Be 3
            $queries = @($global:PraFake.GraphQueries | Where-Object Path -eq users)
            @($queries | Where-Object Query -Match '\$skiptoken=').Count | Should -BeGreaterThan 1
            $filters = @($queries.Filter | Sort-Object -Unique)
            $filters.Count | Should -BeGreaterThan 1
            $allClauses = @(foreach ($filter in $filters) {
                    $filter.Length | Should -BeLessOrEqual 2000
                    [uri]::EscapeDataString($filter).Length | Should -BeLessOrEqual 3000
                    $matches = [regex]::Matches($filter, "\((?:\w+ eq '(?:[^']|'')*'|proxyAddresses/any\(p:p eq '(?:[^']|'')*'\))\)")
                    $matches.Count | Should -BeLessOrEqual 15
                    foreach ($match in $matches) { $match.Value }
                })
            @($allClauses | Sort-Object -CaseSensitive -Unique).Count | Should -Be $allClauses.Count
            foreach ($row in $records) { { Get-Pra2ResourceCollisions -Record $row -Inventory $inventory } | Should -Not -Throw }
        }

        It 'preserves two immutable anchors differing only by Base64 case in the selected union' {
            $record.immutable_id = 'AbC=='
            $other = $record | Select-Object *
            $other.immutable_id = 'abc=='
            $null = Add-PraFakeUser -Upn 'upper@contoso.com' -Id upper -ImmutableId 'AbC=='
            $null = Add-PraFakeUser -Upn 'lower@contoso.com' -Id lower -ImmutableId 'abc=='
            $inventory = Get-Pra2ResourceInventory -Records @($record, $other)
            $inventory.GraphImmutableIds.Count | Should -Be 2
            $inventory.Users.Count | Should -Be 2
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 1
            @(Get-Pra2ResourceCollisions -Record $other -Inventory $inventory).Count | Should -Be 1
        }

        It 'fails closed without whole-tenant or query-mode fallbacks on <Failure>' -TestCases @(
            @{ Failure = 'Unsupported' }, @{ Failure = 'PermissionLaterChunk' }, @{ Failure = 'PermissionLaterPage' }) {
            param($Failure)
            switch ($Failure) {
                Unsupported { $global:PraFake.Faults.GraphQueryUnsupported = $true }
                PermissionLaterChunk { $record.user_principal_name = 'old-logon@contoso.com'; $global:PraFake.Faults.GraphQueryFailsAt = 2 }
                PermissionLaterPage {
                    $global:PraFake.PageSize = 1
                    foreach ($id in 'one', 'two') { $u = Add-PraFakeUser -Upn "$id@contoso.com"; $u.Alias = 'room' }
                    $global:PraFake.Faults.GraphQueryFailsAt = 2
                }
            }
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw
            $global:PraFake.GraphQueries.Count | Should -Be $(if ($Failure -eq 'Unsupported') { 1 } else { 2 })
            foreach ($query in $global:PraFake.GraphQueries) {
                $query.Filter | Should -Not -BeNullOrEmpty
                $query.Query | Should -Not -Match 'startswith'
                if ($query.Headers.Count) {
                    $query.Query | Should -Match '\$count=true'
                    $query.Headers.ConsistencyLevel | Should -BeExactly eventual
                } else { $query.Query | Should -Not -Match '\$count|eventual' }
            }
            $global:PraFake.Calls.Count | Should -Be 0
        }

        It 'rejects oversized Graph predicates before any remote query' {
            $record.alias = 'a' * 2000
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw '*bounded Graph filter size*'
            $global:PraFake.GraphQueries.Count | Should -Be 0
            $global:PraFake.RecipientQueries.Count | Should -Be 0
        }

        It 'fails closed on malformed, missing or out-of-scope Graph identity information (<Fault>)' -TestCases @(
            foreach ($property in 'id', 'userPrincipalName', 'mail', 'mailNickname', 'proxyAddresses', 'onPremisesImmutableId', 'onPremisesSyncEnabled') {
                @{ Fault = "Missing-$property"; Property = $property }
            }
            foreach ($fault in 'WrongScope', 'NullId', 'NullUPN', 'WhitespaceId', 'NullProxies', 'StringProxies', 'NullProxy', 'NonStringMail', 'NonStringAnchor', 'NonBooleanSync', 'ConflictingDuplicate') { @{ Fault = $fault } }) {
            param($Fault, $Property)
            $user = Add-PraFakeUser -Upn 'room@contoso.com' -Id 'claim'
            $global:PraGraphBadUser = ConvertTo-PraFakeGraphUser $user
            switch -Wildcard ($Fault) {
                'Missing-*' { $global:PraGraphBadUser.PSObject.Properties.Remove($Property) }
                WrongScope { $global:PraGraphBadUser.userPrincipalName = 'unknown@contoso.com'; $global:PraGraphBadUser.mail = 'unknown@contoso.com'; $global:PraGraphBadUser.mailNickname = 'foreign'; $global:PraGraphBadUser.proxyAddresses = @() }
                NullId { $global:PraGraphBadUser.id = $null }
                NullUPN { $global:PraGraphBadUser.userPrincipalName = $null }
                WhitespaceId { $global:PraGraphBadUser.id = ' ' }
                NullProxies { $global:PraGraphBadUser.proxyAddresses = $null }
                StringProxies { $global:PraGraphBadUser.proxyAddresses = 'SMTP:room@contoso.com' }
                NullProxy { $global:PraGraphBadUser.proxyAddresses = @($null) }
                NonStringMail { $global:PraGraphBadUser.mail = @('room@contoso.com') }
                NonStringAnchor { $global:PraGraphBadUser.onPremisesImmutableId = @('source-anchor') }
                NonBooleanSync { $global:PraGraphBadUser.onPremisesSyncEnabled = 'true' }
            }
            Mock Invoke-Pra2Graph -ModuleName PRA2.Cloud {
                $user = $global:PraGraphBadUser
                if ($Fault -eq 'ConflictingDuplicate') {
                    $other = $user | Select-Object *
                    $other.mailNickname = 'different-alias'
                    return [pscustomobject]@{ value = @($user, $other) }
                }
                return [pscustomobject]@{ value = @($user) }
            }
            try { { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw }
            finally { Remove-Variable PraGraphBadUser -Scope Global }
        }

        It 'rejects invalid or repeated Graph page links (<Link>)' -TestCases @(
            @{ Link = $null }, @{ Link = '' }, @{ Link = ' ' }, @{ Link = @('https://graph.microsoft.com/v1.0/users?x=1') }
            @{ Link = 'http://graph.microsoft.com/v1.0/users?x=1' }, @{ Link = 'https://foreign.example/v1.0/users?x=1' }
            @{ Link = 'https://graph.microsoft.com/v1.0/groups?x=1' }, @{ Link = 'v1.0/users?x=1' }
            @{ Link = 'https://graph.microsoft.com/v1.0/users?x=1'; Repeat = $true }) {
            param($Link, $Repeat)
            Mock Invoke-Pra2Graph -ModuleName PRA2.Cloud { return [pscustomobject]@{ value = @(); '@odata.nextLink' = $Link } }
            { Get-Pra2ResourceInventory -Records @($record) } | Should -Throw '*page link*'
            Should -Invoke Invoke-Pra2Graph -ModuleName PRA2.Cloud -Times $(if ($Repeat) { 2 } else { 1 }) -Exactly
        }

        It 'treats a complete empty scoped Graph collection as absence and re-reads fresh collisions on resume' {
            $inventory = Get-Pra2ResourceInventory -Records @($record)
            $inventory.Users.Count | Should -Be 0
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 0
            $user = Add-PraFakeUser -Upn 'new-foreign@contoso.com' -Id 'new-foreign' -ImmutableId $record.immutable_id
            $fresh = Get-Pra2ResourceInventory -Records @($record)
            $fresh.Users.Count | Should -Be 1
            @(Get-Pra2ResourceCollisions -Record $record -Inventory $fresh).Count | Should -Be 1
            @($global:PraFake.GraphQueries | Where-Object Path -eq users).Count | Should -Be 4
        }
    }

    Describe 'Unprotected deletion proof' {
        BeforeAll {
            $script:stubs = @()
            foreach ($name in 'Get-Mailbox', 'Get-OrganizationConfig') {
                if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
                    New-Item -Path "function:global:$name" -Value { param([switch]$SoftDeletedMailbox, [object]$ResultSize) } | Out-Null
                    $script:stubs += $name
                }
            }
            Mock Invoke-Pra2Graph -ModuleName PRA2.Cloud {
                $fake = $global:PraResFake
                if ($Uri -match '^v1\.0/users/') { return $fake.Active }
                if ($Uri -match 'deletedItems/[^?]+\?\$select=id,assignedLicenses') { return $fake.DeletedUser }
                if ($Uri -match 'deletedItems/') { return $fake.Deleted }
                throw "Unexpected Graph URI $Uri"
            }
            Mock Get-Mailbox -ModuleName PRA2.Cloud { return $global:PraResFake.Soft }
            Mock Get-OrganizationConfig -ModuleName PRA2.Cloud { return [pscustomobject]@{ InPlaceHolds = @($global:PraResFake.OrgHolds) } }
        }
        AfterAll {
            foreach ($name in $script:stubs) { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
            Remove-Variable -Name PraResFake -Scope Global -ErrorAction SilentlyContinue
        }
        BeforeEach {
            $global:PraResFake = @{
                Active = $null; Deleted = [pscustomobject]@{ id = 'immutable-id' }
                DeletedUser = [pscustomobject]@{ id = 'immutable-id'; assignedLicenses = @(); licenseAssignmentStates = @() }
                Soft = @([pscustomobject]@{ ExchangeGuid = 'independent-guid'; ExternalDirectoryObjectId = 'immutable-id'; IsInactiveMailbox = $false; InPlaceHolds = @()
                        LitigationHoldEnabled = $false; RetentionHoldEnabled = $false; DelayHoldApplied = $false; DelayReleaseHoldApplied = $false; ComplianceTagHoldApplied = $false })
                OrgHolds = @() }
            $intent = [pscustomobject]@{ lifecycle = 'CloudResource'; protectionMode = 'None'; cloudId = 'immutable-id'; cloudGuid = 'independent-guid'
                unprotectedDeletionAttempted = $true; unprotectedDeletionProgress = 'Issued' }
        }

        It 'proves deletion only from the exact deleted identity, one non-inactive soft-deleted GUID, and no hold or licence' {
            Test-Pra2UnprotectedResourceDeleted -Intent $intent | Should -BeTrue
        }
        It 'refuses to treat a missing deletion intent as proof' {
            $intent.PSObject.Properties.Remove('unprotectedDeletionAttempted'); $intent.PSObject.Properties.Remove('unprotectedDeletionProgress')
            { Test-Pra2UnprotectedResourceDeleted -Intent $intent } | Should -Throw '*No journaled unprotected deletion intent*'
        }
        It 'refuses protected journals and missing immutable identifiers' {
            $intent.protectionMode = 'CaseHold'
            { Test-Pra2UnprotectedResourceDeleted -Intent $intent } | Should -Throw
            $intent.protectionMode = 'None'; $intent.cloudGuid = ''
            { Test-Pra2UnprotectedResourceDeleted -Intent $intent } | Should -Throw '*immutable ID/ExchangeGuid is missing*'
        }
        It 'keeps waiting while the identity is still visible after an intent or issued deletion (<Progress>)' -TestCases @(
            @{ Progress = 'Intent' }, @{ Progress = 'Issued' }) {
            param($Progress)
            $intent.unprotectedDeletionProgress = $Progress
            $global:PraResFake.Active = [pscustomobject]@{ id = 'immutable-id' }
            Test-Pra2UnprotectedResourceDeleted -Intent $intent | Should -BeFalse
        }
        It 'refuses an identity that is active again after verified deletion' {
            $intent.unprotectedDeletionProgress = 'Verified'
            $global:PraResFake.Active = [pscustomobject]@{ id = 'immutable-id' }
            { Test-Pra2UnprotectedResourceDeleted -Intent $intent } | Should -Throw '*active again*'
        }
        It 'is not proven without the deleted identity or without a soft-deleted mailbox (<Missing>)' -TestCases @(
            @{ Missing = 'Deleted' }, @{ Missing = 'Soft' }) {
            param($Missing)
            $global:PraResFake[$Missing] = $(if ($Missing -eq 'Soft') { @() } else { $null })
            Test-Pra2UnprotectedResourceDeleted -Intent $intent | Should -BeFalse
        }
        It 'refuses wrong or ambiguous deletion evidence (<Fault>)' -TestCases @(
            @{ Fault = 'DeletedId' }, @{ Fault = 'SoftId' }, @{ Fault = 'Duplicate' }, @{ Fault = 'Inactive' }, @{ Fault = 'InactiveMissing' }) {
            param($Fault)
            switch ($Fault) {
                DeletedId { $global:PraResFake.Deleted = [pscustomobject]@{ id = 'another-id' } }
                SoftId { $global:PraResFake.Soft[0].ExternalDirectoryObjectId = 'another-id' }
                Duplicate { $global:PraResFake.Soft = @($global:PraResFake.Soft[0], $global:PraResFake.Soft[0]) }
                Inactive { $global:PraResFake.Soft[0].IsInactiveMailbox = $true }
                InactiveMissing { $global:PraResFake.Soft[0].PSObject.Properties.Remove('IsInactiveMailbox') }
            }
            { Test-Pra2UnprotectedResourceDeleted -Intent $intent } | Should -Throw
        }
        It 'refuses holds, licences or unreadable protection facts on the deleted resource (<Fault>)' -TestCases @(
            @{ Fault = 'InPlaceHold' }, @{ Fault = 'Litigation' }, @{ Fault = 'DelayHold' }, @{ Fault = 'Licence' }, @{ Fault = 'LicenceState' }, @{ Fault = 'NotExposed' }, @{ Fault = 'OrgHold' }) {
            param($Fault)
            switch ($Fault) {
                InPlaceHold { $global:PraResFake.Soft[0].InPlaceHolds = @('UniHunexpected') }
                Litigation { $global:PraResFake.Soft[0].LitigationHoldEnabled = $true }
                DelayHold { $global:PraResFake.Soft[0].DelayHoldApplied = $true }
                Licence { $global:PraResFake.DeletedUser.assignedLicenses = @([pscustomobject]@{ skuId = 'sku' }) }
                LicenceState { $global:PraResFake.DeletedUser.licenseAssignmentStates = @([pscustomobject]@{ skuId = 'sku' }) }
                NotExposed { $global:PraResFake.DeletedUser.PSObject.Properties.Remove('licenseAssignmentStates') }
                OrgHold { $global:PraResFake.OrgHolds = @('mbxunexpected:2') }
            }
            { Test-Pra2UnprotectedResourceDeleted -Intent $intent } | Should -Throw
        }

        Describe 'Expected tombstone collision qualification' {
            BeforeEach {
                $intent.cloudId = '11111111-1111-1111-1111-111111111111'
                $intent.cloudGuid = '22222222-2222-2222-2222-222222222222'
                $intent.unprotectedDeletionProgress = 'Verified'
                $intent | Add-Member -NotePropertyName creationAttempted -NotePropertyValue $true
                $intent | Add-Member -NotePropertyName cloudType -NotePropertyValue RoomMailbox
                $global:PraResFake.Deleted.id = $intent.cloudId
                $global:PraResFake.DeletedUser.id = $intent.cloudId
                $global:PraResFake.Soft[0].ExternalDirectoryObjectId = $intent.cloudId
                $global:PraResFake.Soft[0].ExchangeGuid = $intent.cloudGuid
                $record = [pscustomobject]@{ kind = 'Room'; exchange_guid = '33333333-3333-3333-3333-333333333333'
                    primary_smtp_address = 'room@contoso.com'; user_principal_name = 'room@contoso.com'; alias = 'room'; immutable_id = 'source-anchor' }
                $recipient = [pscustomobject]@{ ExternalDirectoryObjectId = $intent.cloudId; ExchangeGuid = $intent.cloudGuid
                    Guid = $intent.cloudId; Identity = 'retained-room'; Name = 'not-used-as-deletion-proof'; RecipientTypeDetails = 'RoomMailbox'
                    WhenSoftDeleted = [datetime]::UtcNow; Alias = 'room'; EmailAddresses = @('SMTP:room@contoso.com'); PrimarySmtpAddress = 'room@contoso.com' }
                $inventory = [pscustomobject]@{ Recipients = @($recipient); Users = @([pscustomobject]@{ id = $intent.cloudId; userPrincipalName = 'room@contoso.com' })
                    RecipientKeys = @('smtp:room@contoso.com','alias:room'); GraphImmutableIds = @('source-anchor') }
            }

            It 'excludes only the exact retained tombstone after fresh proof, without changing general collisions (<DateKind>)' -TestCases @(
                @{ DateKind = 'DateTime' }, @{ DateKind = 'DateTimeOffset' }, @{ DateKind = 'String' }) {
                param($DateKind)
                if ($DateKind -eq 'DateTimeOffset') { $recipient.WhenSoftDeleted = [datetimeoffset]::UtcNow }
                if ($DateKind -eq 'String') { $recipient.WhenSoftDeleted = '2026-10-14T12:51:45+02:00' }
                $before = $inventory | ConvertTo-Json -Depth 5 -Compress
                @(Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory).Count | Should -Be 0
                @(Get-Pra2ResourceCollisions -Record $record -Inventory $inventory).Count | Should -Be 2
                ($inventory | ConvertTo-Json -Depth 5 -Compress) | Should -BeExactly $before
                Should -Invoke Get-Mailbox -ModuleName PRA2.Cloud -Times 1 -Exactly
            }

            It 'blocks mismatched identities, unproven deletion dates and incorrect recipient types (<Fault>)' -TestCases @(
                @{ Fault = 'WrongId' }, @{ Fault = 'WrongGuid' }, @{ Fault = 'BothWrong' }, @{ Fault = 'MissingWhen' }
                @{ Fault = 'NullWhen' }, @{ Fault = 'EmptyWhen' }, @{ Fault = 'InvalidWhen' }, @{ Fault = 'ArrayWhen' }
                @{ Fault = 'WrongType' }, @{ Fault = 'ActiveRecipient' }) {
                param($Fault)
                switch ($Fault) {
                    WrongId { $recipient.ExternalDirectoryObjectId = [guid]::NewGuid().ToString() }
                    WrongGuid { $recipient.ExchangeGuid = [guid]::NewGuid().ToString() }
                    BothWrong { $recipient.ExternalDirectoryObjectId = [guid]::NewGuid().ToString(); $recipient.ExchangeGuid = [guid]::NewGuid().ToString() }
                    MissingWhen { $recipient.PSObject.Properties.Remove('WhenSoftDeleted') }
                    NullWhen { $recipient.WhenSoftDeleted = $null }
                    EmptyWhen { $recipient.WhenSoftDeleted = ' ' }
                    InvalidWhen { $recipient.WhenSoftDeleted = 'not-a-date' }
                    ArrayWhen { $recipient.WhenSoftDeleted = @([datetime]::UtcNow) }
                    WrongType { $recipient.RecipientTypeDetails = 'EquipmentMailbox' }
                    ActiveRecipient { $recipient.WhenSoftDeleted = $null; $recipient.Name = 'PRA-R-old-prefix-is-not-proof' }
                }
                @(Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory).Count | Should -Be 1
            }

            It 'blocks foreign active/soft-deleted and recycled competing recipients with the same GUID (<Foreign>)' -TestCases @(
                @{ Foreign = 'Active' }, @{ Foreign = 'SoftDeleted' }, @{ Foreign = 'SameGuid' }) {
                param($Foreign)
                $other = $recipient | Select-Object *
                $other.Guid = [guid]::NewGuid().ToString()
                $other.ExternalDirectoryObjectId = [guid]::NewGuid().ToString()
                if ($Foreign -ne 'SameGuid') { $other.ExchangeGuid = [guid]::NewGuid().ToString() }
                if ($Foreign -eq 'Active') { $other.WhenSoftDeleted = $null }
                $inventory.Recipients += $other
                $collisions = @(Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory)
                $collisions.Count | Should -Be 1
                $collisions[0] | Should -Match ([regex]::Escape($other.ExternalDirectoryObjectId))
            }

            It 'refuses ambiguous duplicates of the expected recipient' {
                $inventory.Recipients += ($recipient | Select-Object *)
                { Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory } | Should -Throw '*inventory is ambiguous*'
            }

            It 'refuses changed fresh deletion/protection facts after journal verification (<Fault>)' -TestCases @(
                @{ Fault = 'ActiveRestored' }, @{ Fault = 'DeletedId' }, @{ Fault = 'DeletedLicenceId' }, @{ Fault = 'SoftId' }
                @{ Fault = 'SoftGuid' }, @{ Fault = 'Held' }, @{ Fault = 'Licensed' }, @{ Fault = 'Inactive' }, @{ Fault = 'OrganizationHold' }) {
                param($Fault)
                switch ($Fault) {
                    ActiveRestored { $global:PraResFake.Active = [pscustomobject]@{ id = $intent.cloudId } }
                    DeletedId { $global:PraResFake.Deleted.id = 'changed-id' }
                    DeletedLicenceId { $global:PraResFake.DeletedUser.id = 'changed-id' }
                    SoftId { $global:PraResFake.Soft[0].ExternalDirectoryObjectId = 'changed-id' }
                    SoftGuid { $global:PraResFake.Soft[0].ExchangeGuid = [guid]::NewGuid().ToString() }
                    Held { $global:PraResFake.Soft[0].DelayHoldApplied = $true }
                    Licensed { $global:PraResFake.DeletedUser.assignedLicenses = @([pscustomobject]@{ skuId = 'sku' }) }
                    Inactive { $global:PraResFake.Soft[0].IsInactiveMailbox = $true }
                    OrganizationHold { $global:PraResFake.OrgHolds = @('mbxunexpected:2') }
                }
                { Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory } | Should -Throw
            }

            It 'refuses incomplete ownership, source type, protection and durable progress pins (<Fault>)' -TestCases @(
                @{ Fault = 'NoMode' }, @{ Fault = 'Protected' }, @{ Fault = 'Issued' }, @{ Fault = 'NoCreateIntent' }
                @{ Fault = 'NoId' }, @{ Fault = 'NoGuid' }, @{ Fault = 'SourceGuid' }, @{ Fault = 'SourceKind' }, @{ Fault = 'CloudType' }) {
                param($Fault)
                switch ($Fault) {
                    NoMode { $intent.PSObject.Properties.Remove('protectionMode') }
                    Protected { $intent.protectionMode = 'CaseHold' }
                    Issued { $intent.unprotectedDeletionProgress = 'Issued' }
                    NoCreateIntent { $intent.creationAttempted = $false }
                    NoId { $intent.cloudId = '' }
                    NoGuid { $intent.cloudGuid = '' }
                    SourceGuid { $intent.cloudGuid = $record.exchange_guid }
                    SourceKind { $record.kind = 'User' }
                    CloudType { $intent.cloudType = 'EquipmentMailbox' }
                }
                { Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory } | Should -Throw
            }

            It 'preserves selected-key and immutable-anchor coverage guards (<Guard>)' -TestCases @(
                @{ Guard = 'RecipientKeys'; Message = '*all selected record keys*' }
                @{ Guard = 'GraphImmutableIds'; Message = '*selected immutable anchor*' }) {
                param($Guard, $Message)
                $inventory.$Guard = @()
                { Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory } | Should -Throw $Message
            }

            It 'keeps a live Graph immutable-anchor collision even without address overlap' {
                $inventory.Users += [pscustomobject]@{ id = 'foreign-id'; userPrincipalName = 'renamed@contoso.com'; onPremisesImmutableId = 'source-anchor' }
                $collisions = @(Get-Pra2UnprotectedRecoverCollisions -Intent $intent -Record $record -Inventory $inventory)
                $collisions.Count | Should -Be 1
                $collisions[0] | Should -Match 'foreign-id'
            }
        }
    }
}
