<#
.SYNOPSIS
    PRA Cloud Mailbox - end-to-end tests of Convert and Recover against an in-memory tenant (tests\FakeTenant.ps1).
.DESCRIPTION
    Runs the real functions of Invoke-PraCloudMailbox.ps1 (loaded without its main block) and of the modules, under
    StrictMode Latest, against a fake Microsoft 365 tenant that reproduces the behaviours of the service that the tool depends on. Only
    the connections, the Graph roles of the token, Entra Connect and the waits are replaced. PowerShell 7 only
    (Convert and Recover run in PowerShell 7).
.NOTES
    Author  : Nicolas Fabert
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
}

AfterAll {
    Restore-PcmGlobalState $script:GlobalBefore
}


Describe 'Convert and Recover end to end (fake tenant)' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
    BeforeAll {
        $repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
        $root = New-PcmTestPackage -Repository $repo
        Import-Module (Join-Path $root 'module\PRA.Common.psm1') -Force
        Import-Module (Join-Path $root 'module\PRA.Store.psm1') -Force
        Import-Module (Join-Path $root 'module\PRA.Cloud.psm1') -Force
        . (Join-Path $PSScriptRoot 'FakeTenant.ps1')
        $global:PraFakeSetUser = (Get-Command Set-User).ScriptBlock
        # The functions of the entry script, without its main block (which re-imports the modules and would drop the mocks).
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
        foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })) {
            . ([scriptblock]::Create($function.Extent.Text))
        }
        # Short waits on demand ($global:PraE2EWait), otherwise the configured minutes.
        function Get-PcmWaitSeconds {
            param([string]$Name)
            if ($global:PraE2EWait -and $global:PraE2EWait.ContainsKey($Name)) { return $global:PraE2EWait[$Name] }
            return [int]$context.Config.Polling[$Name] * 60
        }
        $global:PraE2E = @{ Root = $root }
        $script:ConfigPathBound = $false
        $script:ConfirmBound = $false
        $script:E2ECaller = [pscustomobject]@{ Asked = [System.Collections.Generic.List[string]]::new() }
        $script:E2ECaller | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($Target, $Operation) $this.Asked.Add("$Operation|$Target"); $true }

        function Write-E2ESnapshot {
            <# A new complete snapshot of the objects in scope, as Collect would write it. #>
            $fake = $global:PraFake
            $byUpn = @{}; foreach ($ad in $fake.Ad) { $byUpn[$ad.Upn] = $ad }
            $cn = Open-PcmStore -Path $global:PraE2E.Db -Root $global:PraE2E.Root
            try {
                $id = New-PcmSnapshot $cn @{ RunId = [guid]::NewGuid().ToString('N'); Environment = 'TEST'; ToolVersion = 'test'; ExchangeServer = 'EX1'; ScopeJson = '{}' }
                $rows = @(foreach ($upn in $global:PraE2E.InScope) {
                        $ad = $byUpn[$upn]
                        [ordered]@{ object_guid = $ad.ObjectGuid; kind = $ad.Kind; user_principal_name = $upn; primary_smtp_address = $upn; sam_account_name = ($upn -split '@')[0]
                            exchange_guid = $ad.OnPremGuid; immutable_id = $ad.ImmutableId; recipient_type_details = "$($ad.Kind)Mailbox"
                            alias = ($upn -split '@')[0]; display_name = $upn; hidden_from_address_lists = $(if ($ad.ContainsKey('Hidden')) { [int]$ad.Hidden } else { 1 })
                            name = ($upn -split '@')[0]
                            distinguished_name = $(if ($ad.ContainsKey('Dn')) { $ad.Dn } else { 'CN={0},OU={1},OU=Paris,DC=contoso,DC=com' -f ($upn -split '@')[0], $(if ($ad.Kind -eq 'User') { 'Staff' } else { 'Shared' }) })
                            organizational_unit = $(if ($ad.ContainsKey('Ou')) { $ad.Ou } else { 'contoso.com/Paris/{0}' -f $(if ($ad.Kind -eq 'User') { 'Staff' } else { 'Shared' }) })
                            email_addresses_json = (ConvertTo-Json -InputObject @("SMTP:$upn", "smtp:proxy-$upn", "X500:/o=CONTOSO/ou=Legacy/cn=$upn") -Compress)
                            legacy_exchange_dn = "/o=CONTOSO/ou=Old/cn=$upn"
                            resource_settings_json = $(if ($ad.ContainsKey('ResourceSettings')) { $ad.ResourceSettings | ConvertTo-Json -Depth 10 -Compress } else { $null }) }
                    })
                $null = Add-PcmRow $cn mailbox $id $rows
                $shared = $byUpn['compta@contoso.com']
                $user = { param($upn, $right, $automap) [ordered]@{ mailbox_guid = $shared.ObjectGuid; access_right = $right; trustee = "CONTOSO\$(($upn -split '@')[0])"; trustee_kind = 'User'
                        trustee_upn = $upn; trustee_guid = $byUpn[$upn].ObjectGuid; trustee_name = ($upn -split '@')[0]; resolved = 1; auto_mapping = $automap } }
                $permissions = @(
                    (& $user 'alice@contoso.com' 'FullAccess' 1)
                    [ordered]@{ mailbox_guid = $shared.ObjectGuid; access_right = 'FullAccess'; trustee = 'CONTOSO\grp-compta'; trustee_kind = 'Group'
                        trustee_upn = $null; trustee_guid = 'grp-compta'; trustee_name = 'grp-compta'; resolved = 1; auto_mapping = $null }
                    (& $user 'bob@contoso.com' 'SendAs' $null)
                    (& $user 'alice@contoso.com' 'SendOnBehalf' $null))
                foreach ($ad in @($fake.Ad | Where-Object { $_.ContainsKey('ResourceSettings') -and $_.Upn -in $global:PraE2E.InScope })) {
                    foreach ($right in @('FullAccess','SendAs','SendOnBehalf')) {
                        $permissions += [ordered]@{ mailbox_guid = $ad.ObjectGuid; access_right = $right; trustee = 'CONTOSO\alice'; trustee_kind = 'User'
                            trustee_upn = 'alice@contoso.com'; trustee_guid = $byUpn['alice@contoso.com'].ObjectGuid; trustee_name = 'alice'; resolved = 1; auto_mapping = 0 }
                    }
                }
                $null = Add-PcmRow $cn permission $id $permissions
                $members = @('alice@contoso.com', 'carol@contoso.com') | ForEach-Object { [ordered]@{ group_guid = 'grp-compta'; member_kind = 'User'; member_upn = $_; member_name = ($_ -split '@')[0]; depth = 1 } }
                $null = Add-PcmRow $cn group_member $id @($members)
                $groupRows = @(foreach ($group in @($global:PraE2E.Groups)) {
                        [ordered]@{ group_guid = $group.Guid; name = $group.Name; sam_account_name = $group.Sam; distinguished_name = $group.Dn; primary_smtp_address = $group.Smtp; display_name = $group.Name
                            member_count = @($group.Members | Where-Object { $_.Class -ne 'Group' }).Count } })
                $memberRows = @(foreach ($group in @($global:PraE2E.Groups)) {
                        foreach ($member in @($group.Members)) {
                            [ordered]@{ group_guid = $group.Guid; member_guid = $(if ($member.ContainsKey('Upn') -and $byUpn.ContainsKey($member.Upn)) { $byUpn[$member.Upn].ObjectGuid } else { $member.Guid }); member_class = $member.Class
                                member_name = $member.Name; member_upn = $(if ($member.ContainsKey('Upn')) { $member.Upn } else { $null }); member_smtp = $(if ($member.ContainsKey('Smtp')) { $member.Smtp } else { $null })
                                member_dn = $member.Dn; via_group_guid = $(if ($member.ContainsKey('Via')) { $member.Via } else { $group.Guid }); depth = $(if ($member.ContainsKey('Depth')) { $member.Depth } else { 1 }) } } })
                if (-not $global:PraE2E.NoGroupTables) { $null = Add-PcmRow $cn target_group $id $groupRows; $null = Add-PcmRow $cn target_member $id $memberRows }
                Set-PcmSnapshotStatus $cn $id Complete @{ mailbox = $rows.Count }
                if ($global:PraE2E.NoGroupTables) { $null = Invoke-PcmSql $cn 'DROP TABLE target_member'; $null = Invoke-PcmSql $cn 'DROP TABLE target_group' }
            } finally { Close-PcmStore $cn }
        }

        function Start-E2E {
            <# Sets, in the test, the script-level variables that the functions of the entry script read. #>
            param([ValidateSet('Check', 'Convert', 'Recover')][string]$Action, [string]$Mode = 'Apply', [string]$Batch = '', [string]$Identity = '', [string]$Scope = 'All', [Alias('IdentityPath')][string]$CsvPath = '',
                [string]$SearchBase = '', [string]$Group = '', [int]$MaxObjects = 0, [string]$Expect = '', [switch]$Once, [switch]$ResolveOnly)
            Close-E2E
            # The snapshot file is the one that a Collect left, with its hash file (a Convert or a Recover in Apply mode stops on a snapshot that has none).
            if (Test-Path -LiteralPath $global:PraE2E.Db -PathType Leaf) { $null = Write-PcmHashFile -Path $global:PraE2E.Db }
            $config = Import-PcmConfiguration -Path $global:PraE2E.Config -Root $global:PraE2E.Root
            $run = @{ Root = $global:PraE2E.Root; Version = 'test'; RunId = [guid]::NewGuid().ToString('N'); StartTime = Get-Date; Action = $Action; Mode = $Mode; Phase = ''
                CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; StepIndex = 0; StepTotal = 7; Warnings = 0
                Issues = [System.Collections.Generic.List[object]]::new(); Rows = [System.Collections.Generic.List[object]]::new()
                BackupFiles = [System.Collections.Generic.List[string]]::new(); StateFiles = [System.Collections.Generic.List[string]]::new(); Excluded = [System.Collections.Generic.List[string]]::new()
                LogFolder = $config.Logging.Folder; ReportFolder = $config.Report.Folder; LogFile = (Join-Path $global:PraE2E.Dir 'run.log'); TranscriptPath = ''; TranscriptStarted = $false
                NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''; Config = $config; VerboseEnabled = $false
                BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null
                SkippedMembers = [System.Collections.Generic.List[string]]::new(); ObjectsText = ''; ConfigOverrides = @(); GroupTarget = $null; LeftForLater = 0; CloudDescription = ''; GlobalReportPath = '' }
            if ($CsvPath) { $run.Identities = @(Read-PcmIdentityFile -Path $CsvPath) }
            $run.Approval = New-PcmApprovalCallback $script:E2ECaller $run
            $run.ConfirmPreference = 'None'
            $run.FallbackNext = @(); $run.NextOnFailure = $false; $run.NoChangeMade = $false; $run.Cancelled = $false
            Set-PcmWriteApproval -Approval $run.Approval
            $values = @{ context = $run; Action = $Action; Mode = $Mode; Batch = $Batch; Identity = $Identity; CsvPath = $CsvPath; SearchBase = $SearchBase; Group = $Group; MaxObjects = $MaxObjects
                Expect = $Expect; Once = [bool]$Once; ResolveOnly = [bool]$ResolveOnly; BaseConfigPath = ''; Scope = $Scope; Snapshot = [long]0; Force = $true; PassThru = $false; toolVersion = 'test'; caller = $null
                effectiveMode = $Mode; ConfigPath = $global:PraE2E.Config }
            foreach ($name in $values.Keys) { Set-Variable -Name $name -Value $values[$name] -Scope 1 }
        }

        function Get-E2ERow { param([string]$Identity) @($context.Rows | Where-Object { $_.Identity -eq $Identity }) | Select-Object -First 1 }
        function Add-E2EGroup {
            <# A group captured by Collect in the snapshot (Collect.TargetGroups): the default one has users, a shared mailbox, a computer, a contact, a user that is not in the snapshot and a nested group. #>
            param([string]$Name = 'PRA-Wave1', [string]$Sam = '', [hashtable[]]$Members)
            $guid = [guid]::NewGuid().ToString()
            if (-not $PSBoundParameters.ContainsKey('Members')) {
                $Members = @(
                    @{ Class = 'User'; Name = 'alice'; Upn = 'alice@contoso.com'; Dn = 'CN=alice,OU=Staff,OU=Paris,DC=contoso,DC=com' }
                    @{ Class = 'User'; Name = 'bob'; Upn = 'bob@contoso.com'; Dn = 'CN=bob,OU=Staff,OU=Paris,DC=contoso,DC=com' }
                    @{ Class = 'User'; Name = 'compta'; Upn = 'compta@contoso.com'; Dn = 'CN=compta,OU=Shared,OU=Paris,DC=contoso,DC=com' }
                    @{ Class = 'Computer'; Name = 'PC01'; Guid = [guid]::NewGuid().ToString(); Dn = 'CN=PC01,OU=Computers,DC=contoso,DC=com' }
                    @{ Class = 'Contact'; Name = 'Jane Contact'; Guid = [guid]::NewGuid().ToString(); Dn = 'CN=Jane Contact,OU=Contacts,DC=contoso,DC=com' }
                    @{ Class = 'User'; Name = 'dave'; Upn = 'dave@contoso.com'; Guid = [guid]::NewGuid().ToString(); Dn = 'CN=dave,OU=Staff,OU=Paris,DC=contoso,DC=com' }
                    @{ Class = 'Group'; Name = 'PRA-Sub'; Guid = [guid]::NewGuid().ToString(); Dn = 'CN=PRA-Sub,OU=Groups,DC=contoso,DC=com' })
            }
            $global:PraE2E.Groups.Add(@{ Guid = $guid; Name = $Name; Sam = $(if ($Sam) { $Sam } else { $Name.ToLowerInvariant() + '-sam' }); Dn = "CN=$Name,OU=Groups,DC=contoso,DC=com"
                    Smtp = ($Name.ToLowerInvariant() -replace '[^a-z0-9]', '') + '@contoso.com'; Members = @($Members) })
            Write-E2ESnapshot
            return $guid
        }
        function Get-E2ECallIndex {
            <# Position of the first call that matches, after -After (-1 when absent). #>
            param([string]$Pattern, [int]$After = -1)
            $calls = $global:PraFake.Calls
            for ($i = $After + 1; $i -lt $calls.Count; $i++) { if ($calls[$i] -match $Pattern) { return $i } }
            return -1
        }
        function Assert-E2EBookingDisabled {
            param([hashtable]$Resource)
            $Resource.Calendar.AutomateProcessing | Should -BeExactly None
            foreach ($field in 'AllBookInPolicy','AllRequestInPolicy','AllRequestOutOfPolicy','ForwardRequestsToDelegates','ProcessExternalMeetingMessages') {
                $Resource.Calendar[$field] | Should -BeFalse
            }
            foreach ($field in 'ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy') { @($Resource.Calendar[$field]).Count | Should -Be 0 }
            $Resource.Hidden | Should -BeTrue
        }
        function Get-E2EJournal {
            param([string]$BatchId)
            $journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal -ReadOnly
            try { return [pscustomobject]@{ Batch = (Get-PcmBatch $journal -Id $BatchId); Items = @(Get-PcmBatchItem $journal $BatchId); Changes = @(Get-PcmTenantChange $journal $BatchId) } }
            finally { Close-PcmStore $journal }
        }
        function Close-E2E {
            $runContext = Get-Variable context -ValueOnly -ErrorAction SilentlyContinue
            if ($runContext -and $runContext.Journal) { Close-PcmStore $runContext.Journal; $runContext.Journal = $null }
        }
        function Set-E2EUsersLicensing {
            param([ValidateSet('Kiosk', 'Direct')][string]$Mode, [string]$Sku = '')
            $text = [IO.File]::ReadAllText($global:PraE2E.Config)
            $text = $text.Replace("Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }", "Users = @{ Mode = '$Mode'; SkuPartNumber = '$Sku' }")
            [IO.File]::WriteAllText($global:PraE2E.Config, $text, [Text.UTF8Encoding]::new($true))
        }

        function Add-E2ENativeResource {
            param([string]$Upn = 'room@contoso.com', [string]$Kind = 'Room', [switch]$Only)
            $ad = @{ Upn = $Upn; Kind = $Kind; OnPremGuid = [guid]::NewGuid().ToString(); ObjectGuid = [guid]::NewGuid().ToString(); ImmutableId = "imm-$Upn"; ResourceSettings = (New-PraFakeResourceSettings); SyncExcluded = $true }
            $global:PraFake.Ad.Add($ad)
            if ($Only) { $global:PraE2E.InScope.Clear() }
            $global:PraE2E.InScope.Add($Upn)
            $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("Scope = @{ Mode = 'Auto' }", "Scope = @{ Mode = 'Auto'; ConvertRooms = `$true; ResourceMode = 'CloudResource' }")
            # Only these disposable fake objects acknowledge data loss; the shipped default stays false.
            if ($text -notmatch 'Resources =') { $text = $text.Replace("Environment = 'TEST'", "Environment = 'TEST'`r`n    Resources = @{ LicensingGroupId = '11111111-1111-1111-1111-111111111111'; OnPremisesRoutingDomain = 'onprem.contoso.com'; OnPremisesReady = `$true; AllowUnprotectedDeletion = `$true; Parallel = 100 }") }
            [IO.File]::WriteAllText($global:PraE2E.Config, $text, [Text.UTF8Encoding]::new($true))
            Write-E2ESnapshot
            return $ad
        }

        function Add-E2ERoomList {
            param([hashtable]$Resource, [switch]$Missing)
            $guid = [guid]::NewGuid().ToString()
            $list = @{ Guid = $guid; Name = 'Building A'; DisplayName = 'Building A'; Alias = 'building-a'; Identity = 'building-a@contoso.com'
                ExternalDirectoryObjectId = $guid; PrimarySmtpAddress = 'building-a@contoso.com'; EmailAddresses = @('SMTP:building-a@contoso.com')
                RecipientTypeDetails = 'RoomList'; IsDirSynced = $false; HiddenFromAddressListsEnabled = $false
                CustomAttribute15 = ''; MemberJoinRestriction = 'Closed'; MemberDepartRestriction = 'Closed'
                Members = [System.Collections.Generic.List[string]]::new() }
            if (-not $Missing) { $global:PraFake.RoomLists[$guid] = $list }
            $Resource.ResourceSettings.RoomLists = @(@{ SchemaVersion = 1; SourceGuid = [guid]::NewGuid().ToString(); Name = $list.Name; DisplayName = $list.DisplayName
                    Alias = $list.Alias; PrimarySmtpAddress = $list.PrimarySmtpAddress; HiddenFromAddressListsEnabled = $false
                    EmailAddresses = @('SMTP:building-a@contoso.com','smtp:all-rooms@contoso.com','X500:/o=legacy/cn=building-a'); LegacyExchangeDN = '/o=native/cn=building-a'
                    MemberJoinRestriction = 'Closed'; MemberDepartRestriction = 'Closed'
                    Delivery = @{ RequireSenderAuthenticationEnabled = $true; ModerationEnabled = $false; SendModerationNotifications = 'Never'
                        AcceptMessagesOnlyFromSendersOrMembers = @(); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @() } })
            Write-E2ESnapshot
            return $list
        }

        function Set-E2ELegacyProtectedResource {
            param([string]$BatchId, [switch]$ExplicitMode)
            foreach ($item in @(Get-PcmBatchItem $context.Journal $BatchId | Where-Object kind -in @('Room','Equipment'))) {
                $intent = $item.original_json | ConvertFrom-Json
                $intent.PSObject.Properties.Remove('protectionMode')
                if ($ExplicitMode) { $intent | Add-Member -NotePropertyName protectionMode -NotePropertyValue CaseHold }
                $intent | Add-Member -NotePropertyName licensingGroupId -NotePropertyValue $global:PraFake.Group.Id -Force
                Set-PcmBatchItem $context.Journal $BatchId $item.object_guid @{ original_json = ($intent | ConvertTo-Json -Depth 15 -Compress) }
                $user = Resolve-PraFakeUser $item.identity
                if ($user) {
                    $user.UsageLocation = 'FR'
                    if (-not $global:PraFake.Group.Members.Contains($user.Id)) { $global:PraFake.Group.Members.Add($user.Id) }
                }
            }
        }

        Mock Connect-PcmCloud { 'Microsoft Graph and Exchange Online (fake tenant)' }
        Mock Connect-PcmCompliance { }
        Mock Get-PcmGraphRole { @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
        Mock Invoke-PcmEntraConnect { Invoke-PraFakeEntraConnect -Operation $Operation }
        Mock Wait-PcmCondition { Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
        Mock Wait-PcmCondition -ModuleName PRA.Cloud { Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
        Mock Start-Sleep { Step-PraFake }
        Mock Start-Sleep -ModuleName PRA.Cloud { Step-PraFake }
        Mock Write-Host -ModuleName PRA.Common { }
    }

    BeforeEach {
        Set-StrictMode -Version Latest
        $global:PraE2EWait = @{}
        $dir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $dir
        $global:PraE2E.Dir = $dir
        $global:PraE2E.Db = Join-Path $dir 'snapshot.db'
        $global:PraE2E.Config = Join-Path $dir 'config.psd1'
        Set-Content -LiteralPath $global:PraE2E.Config -Encoding UTF8 -Value @"
@{
    Environment = 'TEST'
    Scope = @{ Mode = 'Auto' }
    Store = @{ Path = '$dir\snapshot.db'; JournalPath = '$dir\journal.db'; BackupFolder = '$dir\backup' }
    Cloud = @{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; AppId = 'app'; CertificateThumbprint = 'ABCDEF0123456789'; DefaultUsageLocation = 'FR' }
    Polling = @{ IntervalSeconds = 5; MailboxTimeoutMinutes = 1; SyncTimeoutMinutes = 1; HoldTimeoutMinutes = 1 }
    Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }; Shared = @{ SkuPartNumber = 'E5' } }
    Retention = @{ HoldPolicy = 'PRA-HOLD' }
    EntraConnect = @{ Mode = 'Manual'; Server = 'aadc' }
    Logging = @{ Folder = '$dir\logs' }
    Report = @{ Folder = '$dir\reports' }
}
"@
        New-PraFakeTenant
        $global:PraFake.Faults.RejectUnfilteredRecipientInventory = $true
        $script:alice = Add-PraFakeUser -Upn 'alice@contoso.com' -Component ([guid]::NewGuid().ToString())
        $script:aliceTeams = $script:alice.Component
        $script:bob = Add-PraFakeUser -Upn 'bob@contoso.com'
        $script:carol = Add-PraFakeUser -Upn 'carol@contoso.com'
        $script:compta = Add-PraFakeUser -Upn 'compta@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope = [System.Collections.Generic.List[string]]@('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com')
        $global:PraE2E.Groups = [System.Collections.Generic.List[hashtable]]::new()
        $global:PraE2E.NoGroupTables = $false
        Write-E2ESnapshot
    }

    AfterEach {
        Close-E2E
        $global:PraE2EWait = @{}
    }

    AfterAll {
        foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "function:$name" -ErrorAction SilentlyContinue }
        Remove-Variable -Name PraFake, PraFakeCommands, PraFakeSetUser, PraE2E, PraE2EWait -Scope Global -ErrorAction SilentlyContinue
    }

    It 'creates and recovers an unlicensed unprotected native <Kind> with no sync/SOA/shared conversion' -TestCases @(@{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        param($Kind)
        $ad = Add-E2ENativeResource -Upn 'native@contoso.com' -Kind $Kind -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'native@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'native@contoso.com').Detail
        $native = Resolve-PraFakeUser 'native@contoso.com'
        $native.Type | Should -Be "${Kind}Mailbox"
        $native.ExchangeGuid | Should -Not -Be $ad.OnPremGuid
        $native.AccountEnabled | Should -BeFalse
        $native.Synced | Should -BeFalse
        $native.Marker | Should -Not -BeNullOrEmpty
        $native.ResourceCapacity | Should -Be 12
        $native.Hidden | Should -BeTrue
        $native.Calendar.AllBookInPolicy | Should -BeFalse
        $native.Calendar.ResourceDelegates | Should -Contain 'alice@contoso.com'
        $native.Calendar.BookInPolicy | Should -Contain 'bob@contoso.com'
        $native.Calendar.RequestInPolicy | Should -Contain 'carol@contoso.com'
        $native.Calendar.RequestOutOfPolicy.Count | Should -Be 0
        $native.WorkHours.WorkingHoursTimeZone | Should -Be 'Romance Standard Time'
        $native.FullAccess.Count | Should -Be 0
        $native.SendAs.Count | Should -Be 0
        $native.SendOnBehalf.Count | Should -Be 0
        $native.AutoMap.Count | Should -Be 0
        foreach ($field in $ad.ResourceSettings.Directory.Keys) {
            $native.Directory[$field] | Should -BeExactly $ad.ResourceSettings.Directory[$field]
        }
        $native.Delivery.RequireSenderAuthenticationEnabled | Should -BeTrue
        $native.Delivery.AcceptMessagesOnlyFromSendersOrMembers | Should -Be @('bob@contoso.com')
        $global:PraFake.Group.Members | Should -Not -Contain $native.Id
        $native.UsageLocation | Should -BeNullOrEmpty
        $native.Direct.Count | Should -Be 0
        $savedGuid = $native.ExchangeGuid
        $batch = $context.BatchId
        $journal = Get-E2EJournal $batch
        $intent = $journal.Items[0].original_json | ConvertFrom-Json
        $intent.lifecycle | Should -Be CloudResource
        $intent.protectionMode | Should -BeExactly None
        $intent.PSObject.Properties.Name | Should -Not -Contain licensingGroupId
        $intent.cloudId | Should -Be $native.Id
        $intent.cloudGuid | Should -Be $savedGuid
        $intent.markerAttribute | Should -Not -Be $context.Config.Retention.TagAttribute
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'native@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.Recycle.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Inactive.Count | Should -Be 0
        $global:PraFake.Recycle[$native.Id].ExchangeGuid | Should -Be $savedGuid
        (Get-E2ERow 'native@contoso.com').Detail | Should -Match 'NO hold/NO inactive copy'
        Should -Invoke Connect-PcmCompliance -Times 0 -Exactly
        $global:PraFake.Reads | Should -Not -Contain 'Get-Mailbox InactiveMailboxOnly'
        $global:PraFake.Contacts.Count | Should -Be 1
        $contact = @($global:PraFake.Contacts.Values)[0]
        $contact.ExternalEmailAddress | Should -Be 'native@onprem.contoso.com'
        $contact.EmailAddresses | Should -Contain 'SMTP:native@contoso.com'
        $contact.EmailAddresses | Should -Contain 'smtp:proxy-native@contoso.com'
        $contact.EmailAddresses | Should -Contain 'X500:/o=CONTOSO/ou=Old/cn=native@contoso.com'
        $contact.HiddenFromAddressListsEnabled | Should -BeTrue
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser|onPremisesSyncBehavior|EntraConnect|Type=Shared|assignLicense|Graph DELETE directory/deletedItems|CaseHold|Graph POST groups/' }) | Should -BeNullOrEmpty
        @($global:PraFake.Ad | Where-Object { $_.ObjectGuid -eq $ad.ObjectGuid })[0].OnPremGuid | Should -Be $ad.OnPremGuid
    }

    It 'keeps mixed users/shared/native resources on their separate lifecycle paths' {
        $null = Add-E2ENativeResource
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | Where-Object FinalStatus -ne Success) | Should -BeNullOrEmpty
        $script:compta.Type | Should -Be SharedMailbox
        (Resolve-PraFakeUser 'room@contoso.com').Type | Should -Be RoomMailbox
        $script:compta.Direct.Count | Should -Be 0
        $global:PraFake.Group.Members | Should -Not -Contain (Resolve-PraFakeUser 'room@contoso.com').Id
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        @($context.Rows | Where-Object FinalStatus -ne Success) | Should -BeNullOrEmpty
        $global:PraFake.Recycle.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match ([regex]::Escape("users/$($native.Id)/onPremisesSyncBehavior")) }) | Should -BeNullOrEmpty
    }

    It 'checks/previews/converts/recovers native-only with enterprise <Capacity>, no licensing reads or compliance' -TestCases @(
        @{ Capacity = 'Full25of25' }, @{ Capacity = 'Absent' }) {
        param($Capacity)
        $null = Add-E2ENativeResource -Only
        if ($Capacity -eq 'Absent') { $global:PraFake.Skus.Remove('sku-e5') }
        else {
            $global:PraFake.Skus['sku-e5'].Enabled = 25
            for ($i = 1; $i -le 25; $i++) {
                $consumer = Add-PraFakeUser -Upn "licensed-$i@contoso.com"
                $consumer.Direct.Add(@{ skuId = 'sku-e5'; disabledPlans = @('plan-exo') })
            }
        }
        $global:PraFake.Faults.NoLicenseReads = $true
        Mock Get-PcmGraphRole { @('User.ReadWrite.All') }
        Mock Connect-PcmCompliance { throw 'No compliance permission.' }
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("DefaultUsageLocation = 'FR'", "DefaultUsageLocation = ''").Replace("HoldPolicy = 'PRA-HOLD'", "HoldPolicy = ''").Replace("LicensingGroupId = '11111111-1111-1111-1111-111111111111'", "LicensingGroupId = ''")
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Check -Mode Preview
        Invoke-PcmCheck
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.Calls.Count | Should -Be 0
        Start-E2E -Action Convert -Mode Preview
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Planned
        $global:PraFake.Calls.Count | Should -Be 0
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $native.UsageLocation | Should -BeNullOrEmpty
        Close-E2E
        Start-E2E -Action Recover -Mode Preview -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').Warnings | Should -Match 'NO hold, NO inactive copy'
        Close-E2E
        Mock Confirm-PcmApply {}
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        Should -Invoke Confirm-PcmApply -Times 1 -Exactly -ParameterFilter { $Question -match 'NO hold, NO inactive copy.*NOT preserved' }
        Should -Invoke Connect-PcmCompliance -Times 0 -Exactly
        @($global:PraFake.Reads | Where-Object { $_ -match 'subscribedSkus|organization$|groups/|CaseHold|InactiveMailboxOnly' }).Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'assignLicense|CaseHold|onPremisesSyncBehavior|Graph POST groups|Graph DELETE directory/deletedItems' }).Count | Should -Be 0
    }

    It 'blocks unprotected deletion by default before any mutation (<Mode>)' -TestCases @(
        @{ Mode = 'Preview' }, @{ Mode = 'Apply' }) {
        param($Mode)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $before = $global:PraFake.Calls.Count
        Close-E2E
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace('; AllowUnprotectedDeletion = $true', '')
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Recover -Mode $Mode -Batch $batch
        $context.Config.Resources.AllowUnprotectedDeletion | Should -BeFalse
        if ($Mode -eq 'Apply') { { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*' } else { Invoke-PcmRecover }
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'AllowUnprotectedDeletion.*cloud reservations/messages loss'
        (Get-E2ERow 'room@contoso.com').Warnings | Should -Match 'NO hold, NO inactive copy'
        $global:PraFake.Calls.Count | Should -Be $before
        $global:PraFake.Users.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
        Should -Invoke Connect-PcmCompliance -Times 0 -Exactly
    }

    It 'refuses organization holds or unreadable holds before creating native lists or resources (<Fault>)' -TestCases @(
        @{ Fault = 'Hold' }, @{ Fault = 'Read' }) {
        param($Fault)
        $room = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $room -Missing
        if ($Fault -eq 'Hold') { $global:PraFake.OrgHolds = @('mbxunexpected:2') }
        else { $global:PraFake.Faults.OrganizationReadFails = $true }
        Start-E2E -Action Check -Mode Preview
        Invoke-PcmCheck
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw
        $global:PraFake.Calls.Count | Should -Be 0
        $global:PraFake.RoomLists.Count | Should -Be 0
        Resolve-PraFakeUser 'room@contoso.com' | Should -BeNullOrEmpty
    }

    It 'refuses unexpected unprotected resource protection/licensing before deletion (<Fault>)' -TestCases @(
        @{ Fault = 'InPlaceHolds' }, @{ Fault = 'LitigationHoldEnabled' }, @{ Fault = 'RetentionHoldEnabled' },
        @{ Fault = 'DelayHoldApplied' }, @{ Fault = 'DelayReleaseHoldApplied' }, @{ Fault = 'ComplianceTagHoldApplied' },
        @{ Fault = 'OrganizationHold' }, @{ Fault = 'OrganizationRead' }, @{ Fault = 'DirectLicence' }, @{ Fault = 'GroupLicence' }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        switch ($Fault) {
            InPlaceHolds { $native.Holds.Add('UniHunexpected') }
            OrganizationHold { $global:PraFake.OrgHolds = @('mbxunexpected:2') }
            OrganizationRead { $global:PraFake.Faults.OrganizationReadFails = $true }
            DirectLicence { $native.Direct.Add(@{ skuId = 'sku-e5'; disabledPlans = @() }) }
            GroupLicence { $global:PraFake.Group.Members.Add($native.Id) }
            default { $native[$Fault] = $true }
        }
        $before = $global:PraFake.Calls.Count
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*'
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Calls.Count | Should -Be $before
        $global:PraFake.Users.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
    }

    It 'keeps exact deletion pending without soft-deleted proof and safely resumes' {
        $Fault = 'NoSoftDeleted'
        $room = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $room -Missing
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults[$Fault] = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Pending
        $global:PraFake.Recycle.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
        Close-E2E
        $global:PraFake.Faults[$Fault] = $false
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold' }).Count | Should -Be 0
        $global:PraFake.RoomLists.Count | Should -Be 1
        @($global:PraFake.RoomLists.Values)[0].Members | Should -Not -Contain @($global:PraFake.Contacts.Keys)[0]
    }

    It 'attempts routing with only its exact verified retained resource tombstone (<Kind>)' -TestCases @(
        @{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        param($Kind)
        $null = Add-E2ENativeResource -Kind $Kind -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $addresses = @($native.Addresses)
        $global:PraFake.Faults.SoftDeletedAddressBlocks = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        (Get-E2EJournal $context.BatchId).Items[0].status | Should -Be Done
        $global:PraFake.Recycle.Contains($native.Id) | Should -BeTrue
        $native.Addresses | Should -Be $addresses
        $native.WhenSoftDeleted | Should -BeOfType ([datetime])
        $global:PraFake.Contacts.Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold|Remove-Mailbox' }) | Should -BeNullOrEmpty
    }

    It 'keeps Exchange contact uniqueness errors Partial and resumes without another DELETE (<Fault>)' -TestCases @(
        @{ Fault = 'ContactSoftDeletedAliasBlocks'; Contacts = 0; Cmdlet = 'New-MailContact' }
        @{ Fault = 'ContactSoftDeletedProxyBlocks'; Contacts = 1; Cmdlet = 'Set-MailContact' }) {
        param($Fault, $Contacts, $Cmdlet)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $addresses = @($native.Addresses)
        $global:PraFake.Faults.SoftDeletedAddressBlocks = $true
        $global:PraFake.Faults[$Fault] = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        $recover = $context.BatchId
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match "$Cmdlet uniqueness conflict"
        $journal = Get-E2EJournal $recover
        $journal.Batch.status | Should -Be Partial
        $journal.Items[0].status | Should -Not -Be Done
        $intent = $journal.Items[0].original_json | ConvertFrom-Json
        $intent.unprotectedDeletionProgress | Should -BeExactly Verified
        $global:PraFake.Reads | Should -Contain "$Cmdlet attempt"
        $global:PraFake.Contacts.Count | Should -Be $Contacts
        if ($Contacts) {
            $intent.contactId | Should -Be @($global:PraFake.Contacts.Keys)[0]
            $intent.contactMarkerPending | Should -BeTrue
        }
        $native.Addresses | Should -Be $addresses
        Close-E2E
        $global:PraFake.Faults[$Fault] = $false
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        $context.BatchId | Should -Be $recover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-MailContact ' }).Count | Should -Be 1
        $native.Addresses | Should -Be $addresses
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold|Remove-Mailbox' }) | Should -BeNullOrEmpty
    }

    It 'blocks a foreign soft-deleted resource alias even when the owned tombstone is verified' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults.ContactFails = $true
        $global:PraFake.Faults.SoftDeletedAddressBlocks = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        $foreign = $native.Clone()
        $foreign.Id = [guid]::NewGuid().ToString()
        $foreign.ExchangeGuid = [guid]::NewGuid().ToString()
        $global:PraFake.Recycle[$foreign.Id] = $foreign
        $global:PraFake.Faults.ContactFails = $false
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Pending
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match ([regex]::Escape($foreign.Id))
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        $global:PraFake.Contacts.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-MailContact|Set-MailContact|Graph DELETE directory/deletedItems|CaseHold' }) | Should -BeNullOrEmpty
    }

    It 'pins a returned contact GUID before a failed post-create read and resumes that contact' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $global:PraFake.Faults.SoftDeletedAddressBlocks = $true
        $global:PraFake.Faults.ContactReadAfterCreateFails = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Contact post-create read refused'
        $recover = $context.BatchId
        $journal = Get-E2EJournal $recover
        $journal.Batch.status | Should -Be Partial
        $intent = $journal.Items[0].original_json | ConvertFrom-Json
        $intent.contactId | Should -Be @($global:PraFake.Contacts.Keys)[0]
        $intent.contactMarkerPending | Should -BeTrue
        $global:PraFake.Reads | Should -Not -Contain 'Set-MailContact attempt'
        Close-E2E
        $global:PraFake.Faults.ContactReadAfterCreateFails = $false
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        $context.BatchId | Should -Be $recover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-MailContact ' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold|Remove-Mailbox' }) | Should -BeNullOrEmpty
    }

    It 'records durable deletion progress and resumes an interrupted DELETE exactly once (<Fault>)' -TestCases @(
        @{ Fault = 'DeleteFails'; First = 'Intent'; Deletes = 2 }
        @{ Fault = 'CrashAfterDelete'; First = 'Issuing'; Deletes = 1 }
        @{ Fault = 'DeleteRejectedStatus'; Value = 400; First = 'Intent'; Deletes = 2 }
        @{ Fault = 'DeleteRejectedStatus'; Value = 403; First = 'Intent'; Deletes = 2 }
        @{ Fault = 'DeleteAmbiguousStatus'; Value = 503; First = 'Issuing'; Deletes = 1 }
        @{ Fault = 'DeleteAmbiguousStatus'; Value = 504; First = 'Issuing'; Deletes = 1 }) {
        param($Fault, $First, $Deletes, $Value = $true)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults[$Fault] = $Value
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $recover = $context.BatchId
        $intent = (Get-E2EJournal $recover).Items[0].original_json | ConvertFrom-Json
        $intent.unprotectedDeletionAttempted | Should -BeTrue
        $intent.unprotectedDeletionProgress | Should -BeExactly $First
        $global:PraFake.Contacts.Count | Should -Be 0
        Close-E2E
        $global:PraFake.Faults[$Fault] = $false
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be $Deletes
        $global:PraFake.Recycle.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 1
        $intent = (Get-E2EJournal $context.BatchId).Items[0].original_json | ConvertFrom-Json
        $intent.unprotectedDeletionProgress | Should -BeExactly Verified
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold' }) | Should -BeNullOrEmpty
    }

    It 'never deletes an identity again when it is restored after a deletion was issued or verified (<Fault>)' -TestCases @(
        @{ Fault = 'NoSoftDeleted'; Progress = 'Issued'; Message = 'still or again visible' }
        @{ Fault = 'ContactFails'; Progress = 'Verified'; Message = 'active again after a verified deletion' }
        @{ Fault = 'CrashAfterDelete'; Progress = 'Issuing'; Message = 'interrupted or ambiguous deletion request' }
        @{ Fault = 'DeleteAmbiguousStatus'; Value = 503; Progress = 'Issuing'; Message = 'interrupted or ambiguous deletion request' }
        @{ Fault = 'DeleteAmbiguousStatus'; Value = 504; Progress = 'Issuing'; Message = 'interrupted or ambiguous deletion request' }) {
        param($Fault, $Progress, $Message, $Value = $true)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults[$Fault] = $Value
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be $(if ($Fault -eq 'NoSoftDeleted') { 'Pending' } else { 'Error' })
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        $intent = (Get-E2EJournal $context.BatchId).Items[0].original_json | ConvertFrom-Json
        $intent.unprotectedDeletionProgress | Should -BeExactly $Progress
        $global:PraFake.Contacts.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        # The same immutable identity and mailbox return to the active directory (recycle-bin restore).
        $global:PraFake.Recycle.Remove($native.Id)
        $global:PraFake.Users[$native.Id] = $native
        $global:PraFake.Faults[$Fault] = $false
        $before = $global:PraFake.Calls.Count
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*'
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match $Message
        $global:PraFake.Calls.Count | Should -Be $before
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        $global:PraFake.Users.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold' }) | Should -BeNullOrEmpty
    }

    It 'refuses an identity restored between deletion verification and routing contact creation' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults.ContactFails = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2EJournal $context.BatchId).Items[0].original_json | ConvertFrom-Json | ForEach-Object { $_.unprotectedDeletionProgress } | Should -BeExactly Verified
        $global:PraFake.Recycle.Remove($native.Id)
        $global:PraFake.Users[$native.Id] = $native
        $global:PraFake.Faults.ContactFails = $false
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-MailContact ' }).Count | Should -Be 0
        $global:PraFake.Contacts.Count | Should -Be 0
    }

    It 'rechecks the deleted immutable identity after address inventory and before routing a contact' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        Mock Get-PcmResourceInventory {
            if ($global:PraFake.Recycle.Contains($native.Id)) {
                $global:PraFake.Recycle.Remove($native.Id)
                $global:PraFake.Users[$native.Id] = $native
            }
            # The inventory completed before the restore became visible; the final ID read must catch it.
            [pscustomobject]@{ Users = @(); Recipients = @() }
        }
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'active again'
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-MailContact|Set-MailContact|Graph DELETE directory/deletedItems|CaseHold' }) | Should -BeNullOrEmpty
        $global:PraFake.Users.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
    }

    It 'fails closed on an unknown, empty or conflicting journal protection mode before any mutation (<Mode>)' -TestCases @(
        @{ Mode = 'Unprotected' }, @{ Mode = '' }, @{ Mode = 'none' }, @{ Mode = 'ConflictingNone' }) {
        param($Mode)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $item = @(Get-PcmBatchItem $context.Journal $batch | Where-Object kind -eq Room)[0]
        $intent = $item.original_json | ConvertFrom-Json
        if ($Mode -eq 'ConflictingNone') { $intent | Add-Member -NotePropertyName holdTag -NotePropertyValue 'UniHlegacy' -Force }
        else { $intent.protectionMode = $Mode }
        Set-PcmBatchItem $context.Journal $batch $item.object_guid @{ original_json = ($intent | ConvertTo-Json -Depth 15 -Compress) }
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $before = $global:PraFake.Calls.Count
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        { Invoke-PcmRecover } | Should -Throw
        $global:PraFake.Calls.Count | Should -Be $before
        $global:PraFake.Users.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        { Invoke-PcmConvert } | Should -Throw
        $global:PraFake.Calls.Count | Should -Be $before
    }

    It 'retains the original protected native journal contract (<ExplicitMode>)' -TestCases @(@{ ExplicitMode = $false }, @{ ExplicitMode = $true }) {
        param($ExplicitMode)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Set-E2ELegacyProtectedResource -BatchId $batch -ExplicitMode:$ExplicitMode
        $native = Resolve-PraFakeUser 'room@contoso.com'
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        $context.Config.Resources.AllowUnprotectedDeletion = $false
        $context.Config.Resources.LicensingGroupId = ''
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.Inactive[0].ExchangeGuid | Should -Be $native.ExchangeGuid
        $global:PraFake.Inactive[0].Holds | Should -Contain (Get-PraFakeHoldTag)
        Should -Invoke Connect-PcmCompliance -Times 1 -Exactly
    }

    It 'resumes original protected native Convert journals without downgrading (<ExplicitMode>)' -TestCases @(
        @{ ExplicitMode = $false }, @{ ExplicitMode = $true }) {
        param($ExplicitMode)
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.CrashAfterNewMailbox = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Set-E2ELegacyProtectedResource -BatchId $batch -ExplicitMode:$ExplicitMode
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $intent = (Get-E2EJournal $batch).Items[0].original_json | ConvertFrom-Json
        Get-PcmResourceProtectionMode $intent | Should -BeExactly CaseHold
        $global:PraFake.Group.Members | Should -Contain $intent.cloudId
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 1
    }

    It 'refuses a changed Recover ownership identity on resume' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $global:PraFake.Faults.NoSoftDeleted = $true
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        $recover = $context.BatchId
        $item = @(Get-PcmBatchItem $context.Journal $recover)[0]
        $intent = $item.original_json | ConvertFrom-Json
        $intent.cloudId = [guid]::NewGuid().ToString()
        Set-PcmBatchItem $context.Journal $recover $item.object_guid @{ original_json = ($intent | ConvertTo-Json -Depth 15 -Compress) }
        $before = $global:PraFake.Calls.Count
        $global:PraFake.Faults.NoSoftDeleted = $false
        Start-E2E -Action Recover -Batch $batch
        { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'ownership cloudId differs'
        $global:PraFake.Calls.Count | Should -Be $before
        $global:PraFake.Contacts.Count | Should -Be 0
    }

    It 'restores a room-list membership without replacing existing members and resumes an interrupted addition' {
        $ad = Add-E2ENativeResource -Only
        $list = Add-E2ERoomList -Resource $ad
        $existing = Add-PraFakeUser -Upn 'existing-room@contoso.com' -Kind Room
        $existing.Type = 'RoomMailbox'
        $list.Members.Add($existing.Id)
        $global:PraFake.Faults.CrashAfterRoomListAdd = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $list.Members | Should -Contain $native.Id
        $list.Members | Should -Contain $existing.Id
        $batch = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $list.Members.Count | Should -Be 2
        @($global:PraFake.Calls | Where-Object { $_ -match 'RoomList Add' }).Count | Should -Be 1
    }

    It 'adds membership only after final resource verification succeeds, and resumes that room without recreation' {
        $room = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $room -Missing
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud {
            $user = Resolve-PraFakeUser $Identity
            $user.AccountEnabled = $true
            [pscustomobject]$user.WorkHours
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'disabled sign-in verification failed'
        (Get-E2EJournal $batch).Batch.status | Should -Be Partial
        Assert-E2EBookingDisabled $native
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept' }).Count | Should -Be 0
        @($global:PraFake.RoomLists.Values)[0].Members.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'RoomList Add ' }) | Should -BeNullOrEmpty
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud { [pscustomobject](Resolve-PraFakeUser $Identity).WorkHours }
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.RoomLists.Values)[0].Members | Should -Contain $native.Id
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox room@contoso.com Room ' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'RoomList Add ' }).Count | Should -Be 1
    }

    It 'creates a missing true RoomList before its rooms and deduplicates their shared list' {
        $first = Add-E2ENativeResource -Upn 'first@contoso.com' -Only
        $null = Add-E2ERoomList -Resource $first -Missing
        $first.ResourceSettings.RoomLists[0].HiddenFromAddressListsEnabled = $true
        $second = Add-E2ENativeResource -Upn 'second@contoso.com'
        $second.ResourceSettings.RoomLists = $first.ResourceSettings.RoomLists
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | Where-Object FinalStatus -ne Success) | Should -BeNullOrEmpty
        $global:PraFake.RoomLists.Count | Should -Be 1
        $list = @($global:PraFake.RoomLists.Values)[0]
        $list.RecipientTypeDetails | Should -Be RoomList
        $list.IsDirSynced | Should -BeFalse
        $list.HiddenFromAddressListsEnabled | Should -BeTrue
        $list.DisplayName | Should -Be 'Building A'
        $list.Alias | Should -Be 'building-a'
        $list.Guid | Should -Not -Be $first.ResourceSettings.RoomLists[0].SourceGuid
        $list.CustomAttribute15 | Should -Match '^[a-f0-9]{32}$'
        $list.EmailAddresses | Should -Contain 'SMTP:building-a@contoso.com'
        $list.EmailAddresses | Should -Contain 'smtp:all-rooms@contoso.com'
        $list.EmailAddresses | Should -Contain 'X500:/o=legacy/cn=building-a'
        $list.EmailAddresses | Should -Contain 'X500:/o=native/cn=building-a'
        $list.RequireSenderAuthenticationEnabled | Should -BeTrue
        $list.MemberJoinRestriction | Should -Be Closed
        $list.Members.Count | Should -Be 2
        $list.Members | Should -Contain (Resolve-PraFakeUser 'first@contoso.com').Id
        $list.Members | Should -Contain (Resolve-PraFakeUser 'second@contoso.com').Id
        (Get-E2ECallIndex 'New-RoomList') | Should -BeLessThan (Get-E2ECallIndex 'New-Mailbox')
        (Get-E2ECallIndex 'Set-RoomList') | Should -BeLessThan (Get-E2ECallIndex 'New-Mailbox')
        (Get-E2ECallIndex 'New-Mailbox first@') | Should -BeLessThan (Get-E2ECallIndex 'RoomList Add')
        (Get-E2ECallIndex ('RoomList Add .* ' + (Resolve-PraFakeUser 'first@contoso.com').Id)) |
            Should -BeLessThan (Get-E2ECallIndex 'New-Mailbox second@')
        $journal = Get-E2EJournal $context.BatchId
        $change = @($journal.Changes | Where-Object kind -eq NativeRoomList)
        $change.Count | Should -Be 1
        $saved = $change[0].original_json | ConvertFrom-Json
        $saved.cloudGuid | Should -Be $list.Guid
        $saved.cloudId | Should -Be $list.ExternalDirectoryObjectId
        $saved.operation | Should -Be Create
        $saved.status | Should -Be Done
        $journal.Batch.status | Should -Be Complete
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-RoomList' }).Count | Should -Be 1
    }

    It 'creates only lists and memberships needed by selected native rooms' {
        $first = Add-E2ENativeResource -Upn 'first@contoso.com' -Only
        $null = Add-E2ERoomList -Resource $first -Missing
        $second = Add-E2ENativeResource -Upn 'second@contoso.com'
        $second.ResourceSettings.RoomLists = @($first.ResourceSettings.RoomLists[0].Clone())
        $second.ResourceSettings.RoomLists[0].SourceGuid = [guid]::NewGuid().ToString()
        $second.ResourceSettings.RoomLists[0].Alias = 'building-b'
        $second.ResourceSettings.RoomLists[0].PrimarySmtpAddress = 'building-b@contoso.com'
        $second.ResourceSettings.RoomLists[0].EmailAddresses = @('SMTP:building-b@contoso.com')
        $second.ResourceSettings.RoomLists[0].LegacyExchangeDN = ''
        Write-E2ESnapshot
        Start-E2E -Action Convert -Identity 'first@contoso.com'
        Invoke-PcmConvert
        $global:PraFake.RoomLists.Count | Should -Be 1
        @($global:PraFake.RoomLists.Values)[0].PrimarySmtpAddress | Should -Be 'building-a@contoso.com'
        @($global:PraFake.RoomLists.Values)[0].Members.Count | Should -Be 1
        Resolve-PraFakeUser 'second@contoso.com' | Should -BeNullOrEmpty
    }

    It 'creates every required list before the first room starts' {
        $first = Add-E2ENativeResource -Upn 'first@contoso.com' -Only
        $null = Add-E2ERoomList -Resource $first -Missing
        $second = Add-E2ENativeResource -Upn 'second@contoso.com'
        $second.ResourceSettings.RoomLists = @($first.ResourceSettings.RoomLists[0].Clone())
        $second.ResourceSettings.RoomLists[0].SourceGuid = [guid]::NewGuid().ToString()
        $second.ResourceSettings.RoomLists[0].Alias = 'building-b'
        $second.ResourceSettings.RoomLists[0].DisplayName = 'Building B'
        $second.ResourceSettings.RoomLists[0].PrimarySmtpAddress = 'building-b@contoso.com'
        $second.ResourceSettings.RoomLists[0].EmailAddresses = @('SMTP:building-b@contoso.com')
        $second.ResourceSettings.RoomLists[0].LegacyExchangeDN = ''
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | Where-Object FinalStatus -ne Success) | Should -BeNullOrEmpty
        $global:PraFake.RoomLists.Count | Should -Be 2
        (Get-E2ECallIndex 'Set-RoomList building-b@') | Should -BeLessThan (Get-E2ECallIndex 'New-Mailbox')
        foreach ($list in $global:PraFake.RoomLists.Values) { $list.Members.Count | Should -Be 1 }
    }

    It 'blocks unexposed room-list creation commands in preview without tenant changes' {
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $original = Microsoft.PowerShell.Core\Get-Command New-DistributionGroup
        Mock Get-Command -ModuleName PRA.Cloud { Microsoft.PowerShell.Core\Get-Command -Name $Name -ErrorAction Stop }
        Mock Get-Command -ModuleName PRA.Cloud -ParameterFilter { $Name -eq 'New-DistributionGroup' } {
            $parameters = @{}
            foreach ($field in $original.Parameters.Keys) { if ($field -ne 'RoomList') { $parameters[$field] = $original.Parameters[$field] } }
            [pscustomobject]@{ Parameters = $parameters }
        }
        Start-E2E -Action Convert -Mode Preview
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'New-DistributionGroup -RoomList.*RBAC'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'resumes interrupted list creation/configuration without duplicate lists or lost memberships (<Fault>)' -TestCases @(
        @{ Fault = 'CrashAfterNewRoomList' }, @{ Fault = 'CrashAfterRoomListConfig' }, @{ Fault = 'RoomListConfigFails' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $global:PraFake.Faults[$Fault] = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 0
        $batch = $context.BatchId
        $saved = @(Get-E2EJournal $batch | ForEach-Object Changes | Where-Object kind -eq NativeRoomList)[0].original_json | ConvertFrom-Json
        $saved.creationAttempted | Should -BeTrue
        $saved.status | Should -Be Failed
        $saved.creationName | Should -Match $saved.token
        $global:PraFake.Faults[$Fault] = $false
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        $global:PraFake.RoomLists.Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-RoomList' }).Count | Should -Be 1
        @($global:PraFake.RoomLists.Values)[0].Members.Count | Should -Be 1
        (Get-E2EJournal $batch).Batch.status | Should -Be Complete
    }

    It 'creates and resumes an Exchange-only list, pinning its returned GUID before reads (<Fault>)' -TestCases @(
        @{ Fault = 'RoomListReadAfterCreateFails' }, @{ Fault = 'RoomListConfigFails' }, @{ Fault = 'CrashAfterNewRoomList' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $global:PraFake.Faults.RoomListExchangeOnly = $true
        $global:PraFake.Faults[$Fault] = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $list = @($global:PraFake.RoomLists.Values)[0]
        $saved = @((Get-E2EJournal $batch).Changes | Where-Object kind -eq NativeRoomList)[0].original_json | ConvertFrom-Json
        $saved.creationAttempted | Should -BeTrue
        $saved.cloudId | Should -BeNullOrEmpty
        if ($Fault -ne 'CrashAfterNewRoomList') {
            $saved.cloudGuid | Should -Be $list.Guid
            $saved.markerPending | Should -BeTrue
        } else { $saved.cloudGuid | Should -BeNullOrEmpty }
        $list.CustomAttribute15 | Should -BeNullOrEmpty
        $global:PraFake.Faults[$Fault] = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        $saved = @((Get-E2EJournal $batch).Changes | Where-Object kind -eq NativeRoomList)[0].original_json | ConvertFrom-Json
        $saved.cloudGuid | Should -Be $list.Guid
        $saved.cloudId | Should -BeNullOrEmpty
        $saved.markerPending | Should -BeFalse
        $list.Members.Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-RoomList|New-Mailbox' }).Count | Should -Be 2
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.RoomLists.Count | Should -Be 1
        foreach ($member in $list.Members) { $global:PraFake.Contacts.Contains($member) | Should -BeFalse }
    }

    It 'reuses an Exchange-only list without changing settings or unrelated members' {
        $ad = Add-E2ENativeResource -Only
        $list = Add-E2ERoomList -Resource $ad
        $list.ExternalDirectoryObjectId = ''; $list.MemberJoinRestriction = 'Open'
        $other = Add-PraFakeUser -Upn 'existing-room@contoso.com' -Kind Room
        $list.Members.Add($other.Id)
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $saved = @((Get-E2EJournal $context.BatchId).Changes | Where-Object kind -eq NativeRoomList)[0].original_json | ConvertFrom-Json
        $saved.operation | Should -Be Reuse
        $saved.cloudGuid | Should -Be $list.Guid
        $saved.cloudId | Should -BeNullOrEmpty
        $list.MemberJoinRestriction | Should -Be Open
        $list.Members | Should -Contain $other.Id
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-RoomList|Set-RoomList' }).Count | Should -Be 0
    }

    It 'rejects changed early-pinned Exchange-only list ownership (<Fault>)' -TestCases @(
        @{ Fault = 'Guid' }, @{ Fault = 'Marker' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $global:PraFake.Faults.RoomListExchangeOnly = $true
        $global:PraFake.Faults.RoomListConfigFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $list = @($global:PraFake.RoomLists.Values)[0]
        if ($Fault -eq 'Guid') { $list.Guid = [guid]::NewGuid().ToString() }
        else { $list.CustomAttribute15 = 'foreign-marker' }
        $before = $global:PraFake.Calls.Count
        $global:PraFake.Faults.RoomListConfigFails = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Calls.Count | Should -Be $before
    }

    It 'pins a directory ID when later exposed and refuses loss or drift (<Fault>, <Operation>)' -TestCases @(
        foreach ($operation in 'Create','Reuse') {
            foreach ($fault in 'Missing','Changed') { @{ Fault = $fault; Operation = $operation } }
        }) {
        param($Fault, $Operation)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing:($Operation -eq 'Create')
        if ($Operation -eq 'Create') {
            $global:PraFake.Faults.RoomListExchangeOnly = $true
            $global:PraFake.Faults.RoomListConfigFails = $true
        } else { @($global:PraFake.RoomLists.Values)[0].ExternalDirectoryObjectId = '' }
        $global:PraFake.Faults.NewMailboxFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $list = @($global:PraFake.RoomLists.Values)[0]
        $id = [guid]::NewGuid().ToString()
        $list.ExternalDirectoryObjectId = $id
        $global:PraFake.Faults.RoomListConfigFails = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        $saved = @((Get-E2EJournal $batch).Changes | Where-Object kind -eq NativeRoomList)[0].original_json | ConvertFrom-Json
        $saved.cloudId | Should -Be $id
        $list.ExternalDirectoryObjectId = $(if ($Fault -eq 'Missing') { '' } else { [guid]::NewGuid().ToString() })
        $before = $global:PraFake.Calls.Count
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Calls.Count | Should -Be $before
    }

    It 'never repeats an attempted list create when no exact object is visible' {
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $global:PraFake.Faults.NewRoomListFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $global:PraFake.Faults.NewRoomListFails = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.RoomLists.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-RoomList|New-Mailbox' }).Count | Should -Be 0
    }

    It 'does not overwrite a reused list or its existing configuration' {
        $ad = Add-E2ENativeResource -Only
        $list = Add-E2ERoomList -Resource $ad
        $list.MemberJoinRestriction = 'Open'
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $list.MemberJoinRestriction | Should -Be Open
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-RoomList|Set-RoomList' }).Count | Should -Be 0
        $intent = @((Get-E2EJournal $context.BatchId).Changes | Where-Object kind -eq NativeRoomList)[0].original_json | ConvertFrom-Json
        $intent.operation | Should -Be Reuse
        $intent.cloudGuid | Should -Be $list.Guid
    }

    It 'refuses replaced or tampered journal-owned lists during resume (<Fault>)' -TestCases @(
        @{ Fault = 'Marker' }, @{ Fault = 'Identity' }, @{ Fault = 'Name' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $global:PraFake.Faults.NewMailboxFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $list = @($global:PraFake.RoomLists.Values)[0]
        switch ($Fault) {
            Marker { $list.CustomAttribute15 = 'foreign-marker' }
            Name { $list.Name = 'Unrelated' }
            Identity {
                $global:PraFake.RoomLists.Clear()
                $list.Guid = [guid]::NewGuid().ToString(); $list.ExternalDirectoryObjectId = $list.Guid
                $global:PraFake.RoomLists[$list.Guid] = $list
            }
        }
        $before = $global:PraFake.Calls.Count
        $global:PraFake.Faults.NewMailboxFails = $false
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Calls.Count | Should -Be $before
        $global:PraFake.RoomLists.Count | Should -Be 1
    }

    It 'blocks cloud recipient/proxy/alias collisions before list or room creation (<Collision>)' -TestCases @(
        @{ Collision = 'Primary' }, @{ Collision = 'Proxy' }, @{ Collision = 'X500' }, @{ Collision = 'Alias' }) {
        param($Collision)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad -Missing
        $foreign = Add-PraFakeUser -Upn 'foreign@contoso.com'
        switch ($Collision) {
            Primary { $foreign.Addresses += 'smtp:building-a@contoso.com' }
            Proxy { $foreign.Addresses += 'smtp:all-rooms@contoso.com' }
            X500 { $foreign.Addresses += 'X500:/o=native/cn=building-a' }
            Alias { $foreign.Alias = 'building-a' }
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.RoomLists.Count | Should -Be 0
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'rejects inconsistent/shared room-list snapshot identities before any changes (<Conflict>)' -TestCases @(
        @{ Conflict = 'Metadata' }, @{ Conflict = 'ListAlias' }, @{ Conflict = 'MailboxAlias' }) {
        param($Conflict)
        $first = Add-E2ENativeResource -Upn 'first@contoso.com' -Only
        $null = Add-E2ERoomList -Resource $first -Missing
        $second = Add-E2ENativeResource -Upn 'second@contoso.com'
        $second.ResourceSettings.RoomLists = @($first.ResourceSettings.RoomLists[0].Clone())
        switch ($Conflict) {
            Metadata { $second.ResourceSettings.RoomLists[0].DisplayName = 'Different source' }
            ListAlias { $second.ResourceSettings.RoomLists[0].SourceGuid = [guid]::NewGuid().ToString() }
            MailboxAlias { $first.ResourceSettings.RoomLists[0].Alias = 'first'; $second.ResourceSettings.RoomLists = @() }
        }
        Write-E2ESnapshot
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*nothing is changed*'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'keeps list failures scoped to dependent rooms and does not report a partial batch complete' {
        $room = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $room -Missing
        $null = Add-E2ENativeResource -Upn 'equipment@contoso.com' -Kind Equipment
        $global:PraFake.Faults.RoomListConfigMismatch = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'equipment@contoso.com').FinalStatus | Should -Be Success
        Resolve-PraFakeUser 'room@contoso.com' | Should -BeNullOrEmpty
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        $global:PraFake.RoomLists.Count | Should -Be 1
    }

    It 'retains created cloud lists on Recover without inserting routing contacts or deleting groups' {
        $room = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $room -Missing
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $list = @($global:PraFake.RoomLists.Values)[0]
        $groupGuid = $list.Guid
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.RoomLists.Count | Should -Be 1
        @($global:PraFake.RoomLists.Values)[0].Guid | Should -Be $groupGuid
        $global:PraFake.Contacts.Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE.*groups|Remove-DistributionGroup' }).Count | Should -Be 0
    }

    It 'refuses unsuitable room lists before creating a mailbox (<Fault>)' -TestCases @(
        @{ Fault = 'Synced' }, @{ Fault = 'WrongType' }, @{ Fault = 'Visibility' }, @{ Fault = 'Primary' }, @{ Fault = 'Read' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        $list = Add-E2ERoomList -Resource $ad
        switch ($Fault) {
            Synced { $list.IsDirSynced = $true }
            WrongType { $list.RecipientTypeDetails = 'MailUniversalDistributionGroup' }
            Visibility { $list.HiddenFromAddressListsEnabled = $true }
            Primary { Mock Get-DistributionGroup -ModuleName PRA.Cloud { [pscustomobject]@{ PrimarySmtpAddress = 'wrong@contoso.com'; RecipientTypeDetails = 'RoomList'; IsDirSynced = $false; Guid = 'wrong'; HiddenFromAddressListsEnabled = $false } } }
            Read { $global:PraFake.Faults.RoomListReadFails = $true }
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox|RoomList Add|Set-User' }).Count | Should -Be 0
    }

    It 'models the noninteractive built-in Set-User pause without changing descriptive data (<Confirmation>)' -TestCases @(
        @{ Confirmation = 'Omitted' }, @{ Confirmation = 'True' }) {
        param($Confirmation)
        $parameters = @{ Identity = $script:alice.Id; Office = 'Changed office'; ErrorAction = 'Stop' }
        if ($Confirmation -eq 'True') { $parameters.Confirm = $true }
        $before = $script:alice.Directory | ConvertTo-Json -Compress
        { Set-User @parameters } | Should -Throw '*A server side error has occurred*'
        ($script:alice.Directory | ConvertTo-Json -Compress) | Should -BeExactly $before
    }

    It 'restores all eight descriptive fields including full Notes with explicitly false confirmation and no Force (<Kind>)' -TestCases @(
        @{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        param($Kind)
        $ad = Add-E2ENativeResource -Kind $Kind -Only
        $ad.ResourceSettings.Directory.Notes = "Full captured notes: réception.`r`nSecond line; retain every word and trailing space. "
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.UserWrites.Count | Should -Be 1
        $write = $global:PraFake.UserWrites[0]
        $write.Identity | Should -BeExactly $native.Id
        $write.ContainsKey('Confirm') | Should -BeTrue
        $write.Confirm | Should -Not -BeNullOrEmpty
        [bool]$write.Confirm | Should -BeFalse
        $write.ContainsKey('Force') | Should -BeFalse
        $ad.ResourceSettings.Directory.Count | Should -Be 8
        $readback = Get-User -Identity $native.Id -ErrorAction Stop
        foreach ($field in $ad.ResourceSettings.Directory.Keys) {
            $write.ContainsKey($field) | Should -BeTrue
            [string]$write[$field] | Should -BeExactly $ad.ResourceSettings.Directory[$field]
            $native.Directory[$field] | Should -BeExactly $ad.ResourceSettings.Directory[$field]
            $readback.$field | Should -BeExactly $ad.ResourceSettings.Directory[$field]
        }
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Complete
    }

    It 'keeps a missing noninteractive confirmation Partial, hidden and booking-disabled' {
        $null = Add-E2ENativeResource -Only
        Mock Set-User -ModuleName PRA.Cloud {
            & $global:PraFakeSetUser -Identity $Identity -Office $Office -ErrorAction Stop
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'A server side error has occurred'
        $global:PraFake.UserWrites[0].ContainsKey('Confirm') | Should -BeFalse
        Assert-E2EBookingDisabled (Resolve-PraFakeUser 'room@contoso.com')
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept|RoomList Add|Graph DELETE' }).Count | Should -Be 0
    }

    It 'reports descriptive-property and room-list update/verification failures without deleting the resource (<Fault>)' -TestCases @(
        @{ Fault = 'DirectoryWriteFails'; Kind = 'Room' }, @{ Fault = 'DirectoryReadMismatch'; Kind = 'Room' },
        @{ Fault = 'DirectoryWriteFails'; Kind = 'Equipment' }, @{ Fault = 'DirectoryReadMismatch'; Kind = 'Equipment' },
        @{ Fault = 'DirectoryNotesReadMismatch'; Kind = 'Room' }, @{ Fault = 'DirectoryNotesReadMissing'; Kind = 'Room' },
        @{ Fault = 'DirectoryNotesReadMismatch'; Kind = 'Equipment' }, @{ Fault = 'DirectoryNotesReadMissing'; Kind = 'Equipment' },
        @{ Fault = 'RoomListAddFails'; Kind = 'Room' }, @{ Fault = 'RoomListNoMembership'; Kind = 'Room' }) {
        param($Fault, $Kind)
        $ad = Add-E2ENativeResource -Kind $Kind -Only
        if ($Kind -eq 'Room') { $null = Add-E2ERoomList -Resource $ad }
        $global:PraFake.Faults[$Fault] = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        Resolve-PraFakeUser 'room@contoso.com' | Should -Not -BeNullOrEmpty
        Assert-E2EBookingDisabled (Resolve-PraFakeUser 'room@contoso.com')
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept' }).Count | Should -Be 0
        if ($Fault -like 'Directory*') {
            @($global:PraFake.Calls | Where-Object { $_ -match 'RoomList Add' }).Count | Should -Be 0
            (Get-E2ECallIndex 'Set-CalendarProcessing .*AutomateProcessing=None') | Should -BeLessThan (Get-E2ECallIndex 'Set-User ')
        }
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE' }).Count | Should -Be 0
    }

    It 'activates exactly the intended source policy only after resource proof and membership (<Processing>)' -TestCases @(
        @{ Processing = 'AutoAccept' }, @{ Processing = 'AutoUpdate' }, @{ Processing = 'None' }) {
        param($Processing)
        $ad = Add-E2ENativeResource -Only
        $ad.Hidden = $false
        $ad.ResourceSettings.Calendar.AutomateProcessing = $Processing
        $ad.ResourceSettings.Calendar.AllBookInPolicy = $true
        $list = Add-E2ERoomList -Resource $ad
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud {
            $native = Resolve-PraFakeUser $Identity
            Assert-E2EBookingDisabled $native
            @($global:PraFake.RoomLists.Values)[0].Members.Count | Should -Be 0
            [pscustomobject]$native.WorkHours
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success -Because (Get-E2ERow 'room@contoso.com').Detail
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $native.Hidden | Should -BeFalse
        $native.Calendar.AutomateProcessing | Should -BeExactly $Processing
        foreach ($field in $ad.ResourceSettings.Calendar.Keys) {
            @($native.Calendar[$field]) | Should -Be @($ad.ResourceSettings.Calendar[$field])
        }
        $list.Members | Should -Contain $native.Id
        (Get-E2ECallIndex 'RoomList Add') | Should -BeLessThan (Get-E2ECallIndex 'Set-CalendarProcessing .*AllBookInPolicy=True')
        (Get-E2EJournal $context.BatchId).Items[0].status | Should -BeExactly Done
    }

    It 'keeps native booking closed before mandatory proof failures and after failed activation (<Fault>)' -TestCases @(
        @{ Fault = 'WorkHoursReadFails' }, @{ Fault = 'FinalIdentityReadFails' },
        @{ Fault = 'CalendarEnableFails' }, @{ Fault = 'CalendarEnabledReadMismatch' }, @{ Fault = 'HiddenFinalFails' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        $ad.Hidden = $false
        $list = Add-E2ERoomList -Resource $ad
        if ($Fault -eq 'WorkHoursReadFails') {
            Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud { throw 'Mandatory work-hours read failed.' }
        } elseif ($Fault -eq 'FinalIdentityReadFails') {
            Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud {
                $global:PraFake.Faults.GraphReadFails = $true
                [pscustomobject](Resolve-PraFakeUser $Identity).WorkHours
            }
        } else { $global:PraFake.Faults[$Fault] = $true }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $native = Resolve-PraFakeUser 'room@contoso.com'
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2EJournal $batch).Batch.status | Should -Be Partial
        Assert-E2EBookingDisabled $native
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE' }).Count | Should -Be 0
        if ($Fault -in @('WorkHoursReadFails','FinalIdentityReadFails')) {
            $list.Members.Count | Should -Be 0
            @($global:PraFake.Calls | Where-Object { $_ -match 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept' }).Count | Should -Be 0
        } else {
            $list.Members | Should -Contain $native.Id
            if ($Fault -notin @('CalendarEnableFails','HiddenFinalFails')) {
                (Get-E2ECallIndex 'RoomList Add') | Should -BeLessThan (Get-E2ECallIndex 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept')
            } elseif ($Fault -eq 'HiddenFinalFails') {
                @($global:PraFake.Calls | Where-Object { $_ -match 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept' }).Count | Should -Be 0
            }
            $global:PraFake.Faults[$Fault] = $false
            Start-E2E -Action Convert -Batch $batch
            Invoke-PcmConvert
            (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
            $native.Calendar.AutomateProcessing | Should -BeExactly AutoAccept
            $native.Calendar.BookInPolicy | Should -Contain 'bob@contoso.com'
            $native.Calendar.ResourceDelegates | Should -Contain 'alice@contoso.com'
            $native.Hidden | Should -BeFalse
            @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 1
            @($global:PraFake.Calls | Where-Object { $_ -match 'RoomList Add' }).Count | Should -Be 1
        }
    }

    It 'reports failed restrictive initialization and failed cleanup without descriptive writes or success' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.BookingDisableFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'FAIL-CLOSED CLEANUP FAILED.*Booking disable refused'
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        (Resolve-PraFakeUser 'room@contoso.com').Hidden | Should -BeTrue
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-User|RoomList Add|Graph DELETE' }).Count | Should -Be 0
    }

    It 'does not relock a completed resource when resuming its batch with changed metadata' {
        $ad = Add-E2ENativeResource -Only
        $ad.Hidden = $false
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $batch = $context.BatchId
        $calendar = $native.Calendar | ConvertTo-Json -Depth 5 -Compress
        $calls = $global:PraFake.Calls.Count
        $ad.ResourceSettings.Remove('Directory')
        Write-E2ESnapshot
        Start-E2E -Action Convert -Batch $batch
        try { Invoke-PcmConvert } catch { $_.Exception.Message | Should -Match 'No object can be converted|No complete snapshot' }
        $native.Calendar | ConvertTo-Json -Depth 5 -Compress | Should -BeExactly $calendar
        $native.Hidden | Should -BeFalse
        $global:PraFake.Calls.Count | Should -Be $calls
        (Get-E2EJournal $batch).Items[0].status | Should -BeExactly Done
    }

    It 'previews automatic creation of a missing room list without changing the tenant' {
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad
        $global:PraFake.RoomLists.Clear()
        Start-E2E -Action Convert -Mode Preview
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Planned
        $global:PraFake.RoomLists.Count | Should -Be 0
        $global:PraFake.Calls.Count | Should -Be 0
    }
    It 'checks room-list readiness without changing the tenant (<Missing>)' -TestCases @(
        @{ Missing = $false }, @{ Missing = $true }) {
        param($Missing)
        $ad = Add-E2ENativeResource -Only
        $null = Add-E2ERoomList -Resource $ad
        if ($Missing) { $global:PraFake.RoomLists.Clear() }
        Start-E2E -Action Check -Mode Preview
        Invoke-PcmCheck
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'requires a new Collect for earlier or incomplete resource metadata (<Fault>)' -TestCases @(
        @{ Fault = 'Version' }, @{ Fault = 'Directory' }, @{ Fault = 'Office' }, @{ Fault = 'RoomLists' }, @{ Fault = 'RoomListArray' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        switch ($Fault) {
            Version { $ad.ResourceSettings.SchemaVersion = 1 }
            Directory { $ad.ResourceSettings.Remove('Directory') }
            Office { $ad.ResourceSettings.Directory.Remove('Office') }
            RoomLists { $ad.ResourceSettings.Remove('RoomLists') }
            RoomListArray { $ad.ResourceSettings.RoomLists = 'not an array' }
        }
        Write-E2ESnapshot
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Collect is required'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'ignores general rights even when an older snapshot contains unresolved native mailbox trustees' {
        $ad = Add-E2ENativeResource -Only
        $cn = Open-PcmStore -Path $global:PraE2E.Db -Root $global:PraE2E.Root
        try {
            $snap = Get-PcmSnapshot $cn
            $null = Add-PcmRow $cn permission $snap.id @(@{ mailbox_guid = $ad.ObjectGuid; access_right = 'FullAccess'
                    trustee = 'missing'; trustee_kind = 'User'; trustee_upn = 'missing@contoso.com'; resolved = 0 })
        } finally { Close-PcmStore $cn }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        (Resolve-PraFakeUser 'room@contoso.com').FullAccess.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'FullAccess|SendAs|SendOnBehalf' }).Count | Should -Be 0
    }
    It 'restores restricted invitation delivery and moderation without general mailbox permissions' {
        $ad = Add-E2ENativeResource -Only
        $ad.ResourceSettings.Delivery.ModerationEnabled = $true
        $ad.ResourceSettings.Delivery.ModeratedBy = @('alice@contoso.com')
        $ad.ResourceSettings.Delivery.BypassModerationFromSendersOrMembers = @('bob@contoso.com')
        $ad.ResourceSettings.Delivery.RejectMessagesFromSendersOrMembers = @('carol@contoso.com')
        $ad.ResourceSettings.Delivery.SendModerationNotifications = 'Internal'
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $native.Delivery.ModerationEnabled | Should -BeTrue
        $native.Delivery.ModeratedBy | Should -Be @('alice@contoso.com')
        $native.Delivery.BypassModerationFromSendersOrMembers | Should -Be @('bob@contoso.com')
        $native.Delivery.RejectMessagesFromSendersOrMembers | Should -Be @('carol@contoso.com')
        $native.Delivery.SendModerationNotifications | Should -Be Internal
        $native.FullAccess.Count | Should -Be 0
    }
    It 'preserves group-based booking and invitation restrictions without expanding mailbox grants' {
        $ad = Add-E2ENativeResource -Only
        $roomGroup = Add-E2ERoomList -Resource $ad
        $ad.ResourceSettings.RoomLists = @()
        $roomGroup.RecipientTypeDetails = 'MailUniversalSecurityGroup'
        $roomGroup.Members.Add($script:bob.Id)
        $ad.ResourceSettings.Calendar.BookInPolicy = @($roomGroup.PrimarySmtpAddress)
        $ad.ResourceSettings.Delivery.AcceptMessagesOnlyFromSendersOrMembers = @($roomGroup.PrimarySmtpAddress)
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $native.Calendar.BookInPolicy | Should -Be @($roomGroup.PrimarySmtpAddress)
        $native.Delivery.AcceptMessagesOnlyFromSendersOrMembers | Should -Be @($roomGroup.PrimarySmtpAddress)
        $roomGroup.Members.Count | Should -Be 1
        $native.FullAccess.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'RoomList Add|FullAccess|SendAs|SendOnBehalf' }).Count | Should -Be 0
    }
    It 'blocks invalid or unresolved invitation restrictions without broadening delivery (<Fault>)' -TestCases @(
        @{ Fault = 'Missing' }, @{ Fault = 'Authentication' }, @{ Fault = 'SenderList' }, @{ Fault = 'NoModerator' }, @{ Fault = 'Unresolved' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        switch ($Fault) {
            Missing { $ad.ResourceSettings.Remove('Delivery') }
            Authentication { $ad.ResourceSettings.Delivery.RequireSenderAuthenticationEnabled = 'false' }
            SenderList { $ad.ResourceSettings.Delivery.Remove('RejectMessagesFromSendersOrMembers') }
            NoModerator { $ad.ResourceSettings.Delivery.ModerationEnabled = $true }
            Unresolved { $ad.ResourceSettings.Delivery.AcceptMessagesOnlyFromSendersOrMembers = @('missing@contoso.com') }
        }
        Write-E2ESnapshot
        Start-E2E -Action Convert
        if ($Fault -eq 'Unresolved') {
            Invoke-PcmConvert
            (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'delivery recipient.*before creation'
        } else { { Invoke-PcmConvert } | Should -Throw '*No object*' }
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'safely refuses a later disaster while the previous Recover routing contact/deleted identity still claims the resource addresses' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $contact = @($global:PraFake.Contacts.Values)[0]
        $previousGuid = @($global:PraFake.Recycle.Values)[0].ExchangeGuid
        $calls = $global:PraFake.Calls.Count
        Close-E2E
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'MailContact'
        $global:PraFake.Contacts.Contains($contact.Guid) | Should -BeTrue
        $contact.ExternalEmailAddress | Should -Be 'room@onprem.contoso.com'
        @($global:PraFake.Recycle.Values)[0].ExchangeGuid | Should -Be $previousGuid
        $global:PraFake.Inactive.Count | Should -Be 0
        $global:PraFake.Calls.Count | Should -Be $calls
    }

    It 'previews native resources without changing the tenant or writing an intent' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert -Mode Preview
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Planned
        $global:PraFake.Calls.Count | Should -Be 0
        $context.BatchId | Should -BeNullOrEmpty
    }

    It 'does not impose resource group prerequisites on unrelated user/shared requests' {
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("Scope = @{ Mode = 'Auto' }", "Scope = @{ Mode = 'Auto'; ResourceMode = 'CloudResource' }")
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | Where-Object FinalStatus -ne Success) | Should -BeNullOrEmpty
        $context.Config.Resources.LicensingGroupId | Should -BeNullOrEmpty
    }

    It 'does not impose unrelated user/direct or shared licensing on native-only requests' {
        $null = Add-E2ENativeResource -Only
        Set-E2EUsersLicensing -Mode Direct -Sku 'NOT-A-TENANT-SKU'
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
    }

    It 'refuses competing chosen snapshot resource addresses before creation' {
        $first = Add-E2ENativeResource -Only
        $second = Add-E2ENativeResource -Upn 'equipment@contoso.com' -Kind Equipment
        $cn = Open-PcmStore -Path $global:PraE2E.Db -Root $global:PraE2E.Root
        try { $null = Invoke-PcmSql $cn 'UPDATE mailbox SET email_addresses_json = @a WHERE object_guid = @g' @{ a = '["SMTP:equipment@contoso.com","smtp:proxy-room@contoso.com"]'; g = $second.ObjectGuid } }
        finally { Close-PcmStore $cn }
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Competing'
        (Get-E2ERow 'equipment@contoso.com').Detail | Should -Match 'Competing'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'does not adopt existing resource/user/proxy/X500/alias collisions (<Collision>)' -TestCases @(
        @{ Collision = 'Resource' }, @{ Collision = 'User' }, @{ Collision = 'Proxy' }, @{ Collision = 'X500' }, @{ Collision = 'LegacyDN' }, @{ Collision = 'Alias' }, @{ Collision = 'Immutable' }) {
        param($Collision)
        $ad = Add-E2ENativeResource -Only
        $existing = Add-PraFakeUser -Upn $(if ($Collision -in @('Resource','User')) { 'room@contoso.com' } else { 'unrelated@contoso.com' })
        if ($Collision -eq 'Resource') { $existing.Type = 'RoomMailbox'; $existing.Synced = $false; $existing.DirSynced = $false; $existing.ImmutableId = '' }
        if ($Collision -eq 'Proxy') { $existing.Addresses += 'smtp:proxy-room@contoso.com' }
        if ($Collision -eq 'X500') { $existing.Addresses += 'X500:/o=CONTOSO/ou=Old/cn=room@contoso.com' }
        if ($Collision -eq 'LegacyDN') { $existing['LegacyExchangeDN'] = '/o=CONTOSO/ou=Old/cn=room@contoso.com' }
        if ($Collision -eq 'Alias') { $existing.Alias = 'room' }
        if ($Collision -eq 'Immutable') { $existing.ImmutableId = $ad.ImmutableId }
        $guid = $existing.ExchangeGuid
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'collid'
        $existing.ExchangeGuid | Should -Be $guid
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'fails closed when the Exchange recipient inventory cannot be read even when Graph has no resource' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.RecipientReadFails = $true
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*denied*'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'reports native Check errors for legacy identity <Failure> without mutation or false readiness' -TestCases @(
        @{ Failure = 'Permission'; Category = 'PermissionDenied'; FQID = 'Unauthorized' }
        @{ Failure = 'Transient'; Category = 'ResourceUnavailable'; FQID = 'ServiceUnavailable' }
        @{ Failure = 'Ambiguous'; Category = 'InvalidArgument'; FQID = 'ManagementObjectAmbiguousException' }) {
        param($Failure, $Category, $FQID)
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.RecipientIdentityError = [System.Management.Automation.ErrorRecord]::new(
            [Exception]::new("Legacy lookup $Failure failure."), $FQID, $Category, '/o=CONTOSO/ou=Old/cn=room@contoso.com')
        Start-E2E -Action Check -Mode Preview
        Invoke-PcmCheck
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match "Legacy lookup $Failure failure"
        $global:PraFake.Calls.Count | Should -Be 0
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'fails closed on Graph access failure instead of treating it as resource absence' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.GraphReadFails = $true
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*403*'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'refuses a deleted cloud identity collision without purging it' {
        $null = Add-E2ENativeResource -Only
        $existing = Add-PraFakeUser -Upn 'room@contoso.com'
        $global:PraFake.Users.Remove($existing.Id); $global:PraFake.Recycle[$existing.Id] = $existing
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        $global:PraFake.Recycle.Contains($existing.Id) | Should -BeTrue
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'requires recollection of a resource in a v1 snapshot' {
        $ad = Add-E2ENativeResource -Only
        $ad.Remove('ResourceSettings')
        Write-E2ESnapshot
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Collect is required'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'rejects incomplete or malformed booking/work-hours settings before creation (<Fault>)' -TestCases @(
        @{ Fault = 'Boolean' }, @{ Fault = 'MissingList' }, @{ Fault = 'Days' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        switch ($Fault) {
            Boolean { $ad.ResourceSettings.Calendar.AllBookInPolicy = 'false' }
            MissingList { $ad.ResourceSettings.Calendar.Remove('BookInPolicy') }
            Days { $ad.ResourceSettings.CalendarConfiguration.WorkingHoursDays = @() }
        }
        Write-E2ESnapshot
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*No object*'
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Collect is required'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'rejects the portable snapshot WorkingHoursDays name as a calendar write parameter' {
        $command = Get-Command Set-MailboxCalendarConfiguration
        $command.Parameters.ContainsKey('WorkDays') | Should -BeTrue
        $command.Parameters.ContainsKey('WorkingHoursDays') | Should -BeFalse
        { Set-MailboxCalendarConfiguration -Identity 'room@contoso.com' -WorkingHoursDays @('Monday') } |
            Should -Throw '*WorkingHoursDays*'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'binds portable WorkingHoursDays to native WorkDays for writes and reads (<Kind>, <Days>)' -TestCases @(
        @{ Kind = 'Room'; Days = @('Monday','Wednesday','Friday'); Expected = 'friday,monday,wednesday' },
        @{ Kind = 'Equipment'; Days = @('WeekendDays'); Expected = 'saturday,sunday' },
        @{ Kind = 'Room'; Days = @('AllDays'); Expected = 'friday,monday,saturday,sunday,thursday,tuesday,wednesday' }) {
        param($Kind, $Days, $Expected)
        $ad = Add-E2ENativeResource -Kind $Kind -Only
        $ad.ResourceSettings.CalendarConfiguration.WorkingHoursDays = $Days
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -like "*Set-MailboxCalendarConfiguration room@contoso.com WorkDays=$Expected" }).Count | Should -Be 1
        $hours = Get-MailboxCalendarConfiguration -Identity 'room@contoso.com'
        $hours.WorkDays | Should -Be ($Expected -replace ',', ', ')
        $hours.PSObject.Properties.Name | Should -Not -Contain WorkingHoursDays
        $ad.ResourceSettings.SchemaVersion | Should -Be 2
        $ad.ResourceSettings.CalendarConfiguration.WorkingHoursDays | Should -Be $Days
        $ad.ResourceSettings.CalendarConfiguration.ContainsKey('WorkDays') | Should -BeFalse
    }

    It 'fails restoration when native WorkDays read-back differs from the WorkDays write' {
        $null = Add-E2ENativeResource -Only
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud {
            [pscustomobject]@{ WorkingHoursStartTime = '08:00:00'; WorkingHoursEndTime = '18:00:00'
                WorkDays = 'Saturday'; WorkingHoursTimeZone = 'Romance Standard Time' }
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Resource work-hours verification failed: WorkingHoursDays'
        @($global:PraFake.Calls | Where-Object { $_ -like '*Set-MailboxCalendarConfiguration room@contoso.com WorkDays=friday,monday,thursday,tuesday,wednesday' }).Count | Should -Be 1
    }

    It 'normalizes native WorkDays flags without broadening the portable working hours (<Kind>, <Days>)' -TestCases @(
        @{ Kind = 'Room'; Days = 'Weekdays'; Expected = 'friday,monday,thursday,tuesday,wednesday' },
        @{ Kind = 'Equipment'; Days = 'WeekendDays'; Expected = 'saturday,sunday' },
        @{ Kind = 'Room'; Days = 'AllDays'; Expected = 'friday,monday,saturday,sunday,thursday,tuesday,wednesday' }) {
        param($Kind, $Days, $Expected)
        $ad = Add-E2ENativeResource -Kind $Kind -Only
        $ad.ResourceSettings.CalendarConfiguration.WorkingHoursDays = @($Days)
        $global:PraFake.Faults.WorkDaysAliasRead = $Days
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -like "*Set-MailboxCalendarConfiguration room@contoso.com WorkDays=$Expected" }).Count | Should -Be 1
        (Resolve-PraFakeUser 'room@contoso.com').WorkHours.WorkDays | Should -Be ($Expected -replace ',', ', ')
        (Get-MailboxCalendarConfiguration -Identity 'room@contoso.com').WorkDays | Should -Be $Days
        $ad.ResourceSettings.CalendarConfiguration.WorkingHoursDays | Should -Be @($Days)
    }

    It 'fails restoration when native WorkDays is absent or unreadable without a portable-name fallback (<Fault>)' -TestCases @(
        @{ Fault = 'Missing' }, @{ Fault = 'Null' }, @{ Fault = 'Unreadable' },
        @{ Fault = 'Empty' }, @{ Fault = 'EmptyArray' }, @{ Fault = 'Unknown' }, @{ Fault = 'Read' }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        $script:workDaysReadFault = $Fault
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud {
            if ($script:workDaysReadFault -eq 'Read') { throw 'WorkDays calendar read unavailable' }
            $hours = [pscustomobject]@{ WorkingHoursStartTime = '08:00:00'; WorkingHoursEndTime = '18:00:00'
                WorkingHoursDays = 'Weekdays'; WorkingHoursTimeZone = 'Romance Standard Time' }
            switch ($script:workDaysReadFault) {
                Null { $hours | Add-Member -NotePropertyName WorkDays -NotePropertyValue $null }
                Unreadable { $hours | Add-Member -MemberType ScriptProperty -Name WorkDays -Value { throw 'WorkDays getter unavailable' } }
                Empty { $hours | Add-Member -NotePropertyName WorkDays -NotePropertyValue '' }
                EmptyArray { $hours | Add-Member -NotePropertyName WorkDays -NotePropertyValue @() }
                Unknown { $hours | Add-Member -NotePropertyName WorkDays -NotePropertyValue 'UnknownDays' }
            }
            $hours
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'WorkDays|working-hours days'
        @($global:PraFake.Calls | Where-Object { $_ -like '*Set-MailboxCalendarConfiguration room@contoso.com WorkDays=friday,monday,thursday,tuesday,wednesday' }).Count | Should -Be 1
    }

    It 'fails restoration when other native working-hours metadata is absent, unreadable or differs (<Field>, <Fault>)' -TestCases @(
        @{ Field = 'WorkingHoursStartTime'; Fault = 'Missing' }, @{ Field = 'WorkingHoursEndTime'; Fault = 'Null' },
        @{ Field = 'WorkingHoursTimeZone'; Fault = 'Missing' }, @{ Field = 'WorkingHoursTimeZone'; Fault = 'Null' },
        @{ Field = 'WorkingHoursTimeZone'; Fault = 'EmptyId' }, @{ Field = 'WorkingHoursTimeZone'; Fault = 'Mismatch' },
        @{ Field = 'WorkingHoursStartTime'; Fault = 'TimeSpanDays' }, @{ Field = 'WorkingHoursEndTime'; Fault = 'Mismatch' }) {
        param($Field, $Fault)
        $null = Add-E2ENativeResource -Only
        $script:workReadField = $Field; $script:workReadFault = $Fault
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Cloud {
            $hours = [pscustomobject]@{ WorkingHoursStartTime = '08:00:00'; WorkingHoursEndTime = '18:00:00'
                WorkDays = 'Weekdays'; WorkingHoursTimeZone = 'Romance Standard Time' }
            switch ($script:workReadFault) {
                Missing { $hours.PSObject.Properties.Remove($script:workReadField) }
                Null { $hours.($script:workReadField) = $null }
                EmptyId { $hours.($script:workReadField) = [pscustomobject]@{ Id = ''; DisplayName = 'Romance Standard Time' } }
                Mismatch { $hours.($script:workReadField) = $(if ($script:workReadField -eq 'WorkingHoursTimeZone') { 'UTC' } else { '19:00:00' }) }
                TimeSpanDays { $hours.($script:workReadField) = [timespan]'1.08:00:00' }
            }
            $hours
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match "Resource work-hours verification failed: $Field"
    }

    It 'restores a minimum portable resource snapshot without demanding unexposed optional scalars (<Kind>)' -TestCases @(
        @{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        param($Kind)
        $ad = Add-E2ENativeResource -Kind $Kind -Only
        $required = @('AutomateProcessing','AllBookInPolicy','AllRequestInPolicy','AllRequestOutOfPolicy',
            'ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy','ScheduleOnlyDuringWorkHours')
        foreach ($key in @($ad.ResourceSettings.Calendar.Keys)) { if ($key -notin $required) { $ad.ResourceSettings.Calendar.Remove($key) } }
        $global:PraFake.Faults.WorkTimeZoneObjectRead = $true
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $native.Calendar.Keys.Count | Should -Be ($required.Count + 2)
        $native.Calendar.ForwardRequestsToDelegates | Should -BeFalse
        $native.Calendar.ProcessExternalMeetingMessages | Should -BeFalse
        $native.Calendar.AllBookInPolicy | Should -BeFalse
        $native.Calendar.ContainsKey('BookingWindowInDays') | Should -BeFalse
        $native.WorkHours.WorkingHoursTimeZone | Should -Be 'Romance Standard Time'
    }

    It 'blocks a future native delegate before creation and restores it exactly after the referenced resource is provisioned' {
        $first = Add-E2ENativeResource -Upn 'first@contoso.com' -Only
        $null = Add-E2ENativeResource -Upn 'second@contoso.com'
        $first.ResourceSettings.Calendar.ResourceDelegates = @('second@contoso.com')
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'first@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'first@contoso.com').Detail | Should -Match 'references a native resource not yet journal-owned/provisioned'
        Resolve-PraFakeUser 'first@contoso.com' | Should -BeNullOrEmpty
        (Get-E2ERow 'second@contoso.com').FinalStatus | Should -Be Success
        $batch = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'first@contoso.com').FinalStatus | Should -Be Success
        (Resolve-PraFakeUser 'first@contoso.com').Calendar.ResourceDelegates | Should -Contain 'second@contoso.com'
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox ' }).Count | Should -Be 2
    }

    It 'blocks mutually dependent native delegates without creating mailboxes or omitting recipients' {
        $first = Add-E2ENativeResource -Upn 'first@contoso.com' -Only
        $second = Add-E2ENativeResource -Upn 'second@contoso.com'
        $first.ResourceSettings.Calendar.ResourceDelegates = @('second@contoso.com')
        $second.ResourceSettings.Calendar.ResourceDelegates = @('first@contoso.com')
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | Where-Object FinalStatus -eq Error).Count | Should -Be 2
        $global:PraFake.Calls.Count | Should -Be 0
        @($context.Rows | Where-Object { $_.Detail -notmatch 'no booking recipient is omitted' }).Count | Should -Be 0
    }

    It 'ignores resource licensing group plan/capacity/authority for new native resources (<Fault>)' -TestCases @(
        @{ Fault = 'Plan' }, @{ Fault = 'Capacity' }, @{ Fault = 'Synced' }, @{ Fault = 'FormerlySynced' }, @{ Fault = 'Dynamic' }, @{ Fault = 'Security' }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        if ($Fault -eq 'Plan') { $global:PraFake.Group.Licences[0].disabledPlans += 'plan-exo' }
        if ($Fault -eq 'Capacity') { $global:PraFake.Skus['sku-e5'].Enabled = 0 }
        if ($Fault -eq 'Synced') { $global:PraFake.Group.Synced = $true }
        if ($Fault -eq 'FormerlySynced') { $global:PraFake.Group.Synced = $true; $global:PraFake.Group.CloudManaged = $true }
        if ($Fault -eq 'Dynamic') { $global:PraFake.Group['GroupTypes'] = @('DynamicMembership') }
        if ($Fault -eq 'Security') { $global:PraFake.Group['SecurityEnabled'] = $false }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $global:PraFake.Group.Members.Count | Should -Be 0
        @($global:PraFake.Calls | Where-Object { $_ -match 'onPremisesSyncBehavior|assignLicense|Graph POST groups' }).Count | Should -Be 0
    }

    It 'creates multiple native resources without consuming the single available enterprise unit' {
        $null = Add-E2ENativeResource -Only
        $null = Add-E2ENativeResource -Upn 'equipment@contoso.com' -Kind Equipment
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | Where-Object FinalStatus -eq Success).Count | Should -Be 2
        $global:PraFake.Group.Members.Count | Should -Be 0
    }

    It 'still validates mixed-scope user licensing capacity before changing anything' {
        $null = Add-E2ENativeResource
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*prerequisites*'
        $global:PraFake.Calls.Count | Should -Be 0
    }

    It 'resumes an interrupted resource creation only by journal ownership' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.CrashAfterNewMailbox = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $batch = $context.BatchId
        $item = (Get-E2EJournal $batch).Items[0]
        $item.entra_id | Should -BeNullOrEmpty
        ($item.original_json | ConvertFrom-Json).token | Should -Not -BeNullOrEmpty
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox ' }).Count | Should -Be 1
    }

    It 'refuses ambiguous interrupted resource ownership without changing either resource' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.CrashAfterNewMailbox = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $owned = Resolve-PraFakeUser 'room@contoso.com'
        $other = Add-PraFakeUser -Upn 'other@contoso.com'
        $other.Name = $owned.Name
        $before = $global:PraFake.Calls.Count
        Close-E2E
        Start-E2E -Action Convert -Batch $batch
        { Invoke-PcmConvert } | Should -Throw '*Ambiguous*'
        $global:PraFake.Calls.Count | Should -Be $before
    }

    It 'retains intent on creation failure and never performs guessed-SMTP cleanup' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.NewMailboxFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2EJournal $context.BatchId).Items[0].step | Should -Be ResourceCreationIntent
        @($global:PraFake.Calls | Where-Object { $_ -match 'DELETE' }) | Should -BeNullOrEmpty
    }

    It 'keeps native creation intent when configuration changes after the creation response was lost' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.CrashAfterNewMailbox = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Close-E2E
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("ResourceMode = 'CloudResource'", "ResourceMode = 'SynchronizedShared'").Replace('ConvertRooms = $true', 'ConvertRooms = $false')
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        (Resolve-PraFakeUser 'room@contoso.com').Type | Should -Be RoomMailbox
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser|onPremisesSyncBehavior|Type=Shared' }) | Should -BeNullOrEmpty
    }

    It 'never marks a resource done when sign-in disabling is not verified' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.SignInDisableFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'sign-in'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE' }) | Should -BeNullOrEmpty
    }

    It 'pins the returned mailbox IDs before a post-create read failure and resumes without another create' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.ResourceReadAfterCreateFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $user = Resolve-PraFakeUser 'room@contoso.com'
        $saved = (Get-E2EJournal $batch).Items[0].original_json | ConvertFrom-Json
        $saved.cloudId | Should -Be $user.Id
        $saved.cloudGuid | Should -Be $user.ExchangeGuid
        $saved.markerPending | Should -BeTrue
        $user.Marker | Should -BeNullOrEmpty
        Assert-E2EBookingDisabled $user
        $global:PraFake.Faults.ResourceReadAfterCreateFails = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 1
    }

    It 'waits for disabled sign-in propagation with null cloud-only anchors' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.SignInPropagationReads = 3
        $global:PraFake.Faults.SignInReadFault = 'NullAnchors'
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        $user = Resolve-PraFakeUser 'room@contoso.com'
        $user.SignInPollReads | Should -BeGreaterOrEqual 4
        $user.AccountEnabled | Should -BeFalse
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 1
    }

    It 'fails immediately on invalid Graph sign-in/source identity instead of waiting (<Fault>)' -TestCases @(
        foreach ($fault in 'MissingAccount','NullAccount','StringAccount','WrongId','Sync','Anchor','MissingSync','MissingAnchor','ReadError') { @{ Fault = $fault } }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.SignInPropagationReads = 3
        $global:PraFake.Faults.SignInReadFault = $Fault
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $user = Resolve-PraFakeUser 'room@contoso.com'
        $user.SignInPollReads | Should -Be 1
        $saved = (Get-E2EJournal $context.BatchId).Items[0].original_json | ConvertFrom-Json
        $saved.cloudId | Should -Be $user.Id
        $saved.cloudGuid | Should -Be $user.ExchangeGuid
        $user.Marker | Should -BeNullOrEmpty
        Assert-E2EBookingDisabled $user
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-CalendarProcessing .*AutomateProcessing=AutoAccept|Graph DELETE' }).Count | Should -Be 0
    }

    It 'retains early pins on sign-in timeout and resumes the exact same equipment' {
        $null = Add-E2ENativeResource -Upn 'equipment@contoso.com' -Kind Equipment -Only
        $global:PraE2EWait = @{ MailboxTimeoutMinutes = 40 }
        $global:PraFake.Faults.SignInDisableFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $user = Resolve-PraFakeUser 'equipment@contoso.com'
        $saved = (Get-E2EJournal $batch).Items[0].original_json | ConvertFrom-Json
        $saved.cloudId | Should -Be $user.Id
        $saved.cloudGuid | Should -Be $user.ExchangeGuid
        $user.SignInPollReads | Should -BeLessOrEqual 41
        $global:PraFake.Faults.SignInDisableFails = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'equipment@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 1
    }

    It 'never repeats an attempted resource create when the exact object is absent' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.NewMailboxFails = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $global:PraFake.Faults.NewMailboxFails = $false
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'already attempted'
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox' }).Count | Should -Be 0
    }
    It 'refuses missing cloud custom schema or unresolved booking recipients before mailbox creation (<Fault>)' -TestCases @(
        @{ Fault = 'Schema' }, @{ Fault = 'Recipient' }) {
        param($Fault)
        $ad = Add-E2ENativeResource -Only
        if ($Fault -eq 'Schema') { $global:PraFake.ResourceSchema = @() }
        else { $ad.ResourceSettings.Calendar.ResourceDelegates = @('missing@contoso.com'); Write-E2ESnapshot }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox|Set-ResourceConfig' }) | Should -BeNullOrEmpty
    }

    It 'blocks unsupported resource custom parameters before creation with an actionable session/RBAC finding' {
        $null = Add-E2ENativeResource -Only
        Mock Get-Command -ModuleName PRA.Cloud -ParameterFilter { $Name -eq 'Set-Mailbox' } {
            [pscustomobject]@{ Parameters = @{ ResourceCapacity = $null } }
        }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'Set-Mailbox -ResourceCustom.*RBAC'
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox|Set-Mailbox|Set-ResourceConfig' }) | Should -BeNullOrEmpty
    }

    It 'blocks unexposed cloud resource schema rather than mutating it or guessing custom-property support' {
        $null = Add-E2ENativeResource -Only
        Mock Get-ResourceConfig -ModuleName PRA.Cloud { [pscustomobject]@{ Name = 'unsupported schema view' } }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2ERow 'room@contoso.com').Detail | Should -Match 'ResourcePropertySchema is not exposed'
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-Mailbox|Set-Mailbox|Set-ResourceConfig' }) | Should -BeNullOrEmpty
    }

    It 'waits for enabled enterprise plan and retains active resource/licence membership' {
        $null = Add-E2ENativeResource -Only
        $global:PraFake.Faults.CrashAfterNewMailbox = $true
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Set-E2ELegacyProtectedResource -BatchId $batch
        Close-E2E
        $global:PraFake.Faults.NoEnterprise = $true
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Pending
        $global:PraFake.Group.Members | Should -Contain (Resolve-PraFakeUser 'room@contoso.com').Id
    }

    It 'recovers native lifecycle from the journal after ResourceMode/ConvertRooms configuration changes' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Close-E2E
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("ResourceMode = 'CloudResource'", "ResourceMode = 'SynchronizedShared'").Replace('ConvertRooms = $true', 'ConvertRooms = $false')
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect|onPremisesSyncBehavior' }) | Should -BeNullOrEmpty
    }

    It 'does not delete a native resource without its specific hold, correct GUID or active licence (<Fault>)' -TestCases @(
        @{ Fault = 'Hold' }, @{ Fault = 'Guid' }, @{ Fault = 'Licence' }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $resource = Resolve-PraFakeUser 'room@contoso.com'
        Set-E2ELegacyProtectedResource -BatchId $batch
        if ($Fault -eq 'Hold') { $global:PraFake.Faults.NoHoldStamp = $true; $resource.Holds.Add('UniHunrelated') }
        if ($Fault -eq 'Guid') { $resource.ExchangeGuid = [guid]::NewGuid().ToString() }
        if ($Fault -eq 'Licence') { $global:PraFake.Faults.NoEnterprise = $true }
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        if ($Fault -eq 'Guid') { { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*' } else { Invoke-PcmRecover }
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }) | Should -BeNullOrEmpty
        $global:PraFake.Users.Contains($resource.Id) | Should -BeTrue
        $global:PraFake.Group.Members | Should -Contain $resource.Id
    }

    It 'resumes resource deletion/contact side effects before the journal update (<Fault>)' -TestCases @(
        @{ Fault = 'CrashAfterDelete' }, @{ Fault = 'CrashAfterContact' }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $global:PraFake.Faults[$Fault] = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'New-MailContact ' }).Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems' }) | Should -BeNullOrEmpty
    }

    It 'retains inactive data and deleted identity when contact address reuse is blocked' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Set-E2ELegacyProtectedResource -BatchId $batch
        $resource = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults.ContactFails = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Recycle.Contains($resource.Id) | Should -BeTrue
        $global:PraFake.Inactive[0].ExchangeGuid | Should -Be $resource.ExchangeGuid
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold Remove' }) | Should -BeNullOrEmpty
    }

    It 'retains a deletion pending inactive provisioning and resumes without repeating deletion' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Set-E2ELegacyProtectedResource -BatchId $batch
        $native = Resolve-PraFakeUser 'room@contoso.com'
        $global:PraFake.Faults.NoInactive = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Pending
        $global:PraFake.Recycle.Contains($native.Id) | Should -BeTrue
        $global:PraFake.Contacts.Count | Should -Be 0
        Close-E2E
        $global:PraFake.Faults.NoInactive = $false
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }).Count | Should -Be 1
    }

    It 'never creates routing contact when the exact inactive GUID lacks the expected specific tool hold' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Set-E2ELegacyProtectedResource -BatchId $batch
        $global:PraFake.Faults.WrongInactiveHold = $true
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $global:PraFake.Contacts.Count | Should -Be 0
        $global:PraFake.Inactive.Count | Should -Be 1
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE directory/deletedItems|CaseHold Remove' }) | Should -BeNullOrEmpty
    }

    It 'refuses a pre-existing contact that imitates the journal creation name before any contact attempt' {
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        $intent = (Get-E2EJournal $batch).Items[0].original_json | ConvertFrom-Json
        $null = New-MailContact -Name ($intent.creationName -replace '^PRA-R-', 'PRA-C-') -Alias 'foreign-contact' -ExternalEmailAddress 'foreign@onprem.contoso.com'
        $contact = @($global:PraFake.Contacts.Values)[0]
        Close-E2E
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        $contact.ExternalEmailAddress | Should -Be 'foreign@onprem.contoso.com'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/|Set-MailContact' }) | Should -BeNullOrEmpty
    }

    It 'requires on-premises readiness and a non-loop routing domain before deletion (<Fault>)' -TestCases @(
        @{ Fault = 'Ready' }, @{ Fault = 'Loop' }) {
        param($Fault)
        $null = Add-E2ENativeResource -Only
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Close-E2E
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        if ($Fault -eq 'Ready') { $text = $text.Replace('OnPremisesReady = $true', 'OnPremisesReady = $false') }
        else { $text = $text.Replace("'onprem.contoso.com'", "'contoso.com'") }
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Recover -Batch $batch
        { Invoke-PcmRecover } | Should -Throw '*Nothing to roll*'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE users/' }) | Should -BeNullOrEmpty
    }

    It 'keeps a mixed Recover partial when native prerequisites are blocked and resumes only the native resource' {
        $null = Add-E2ENativeResource
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        Close-E2E
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace('OnPremisesReady = $true', 'OnPremisesReady = $false')
        [IO.File]::WriteAllText($global:PraE2E.Config, $text)
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Error
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Partial
        $doneCalls = $global:PraFake.Calls.Count
        Close-E2E
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace('OnPremisesReady = $false', 'OnPremisesReady = $true'))
        Start-E2E -Action Recover -Batch $batch
        Invoke-PcmRecover
        $context.Rows.Count | Should -Be 1
        (Get-E2ERow 'room@contoso.com').FinalStatus | Should -Be Success
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be Complete
        @($global:PraFake.Calls | Select-Object -Skip $doneCalls | Where-Object { $_ -match 'EntraConnect|onPremisesSyncBehavior' }) | Should -BeNullOrEmpty
    }

    It 'Preview plans every object and changes nothing' {
        Start-E2E -Action Convert -Mode Preview
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Planned', 'Planned', 'Planned')
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph (PATCH|POST|DELETE)|Set-|CaseHold|EntraConnect' }) | Should -BeNullOrEmpty
        $context.BatchId | Should -BeNullOrEmpty
    }

    It 'converts the users then the shared mailbox in the right order' {
        Start-E2E -Action Convert
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
        # Users: the Teams storage becomes the mailbox; cloud source of authority; licence group.
        $alice.Type | Should -Be 'UserMailbox'
        $alice.ExchangeGuid | Should -Be $aliceTeams
        $alice.CloudManaged | Should -BeTrue
        $global:PraFake.Group.Members | Should -Contain $alice.Id
        $bob.Type | Should -Be 'UserMailbox'
        (Get-E2ECallIndex "Set-MailUser alice@contoso.com ExchangeGuid=0{8}-") | Should -BeLessThan (Get-E2ECallIndex "Graph PATCH users/$($alice.Id)/onPremisesSyncBehavior")
        # Shared: SharedMailbox, tag, the 4 permissions of the snapshot, temporary licence gone.
        $compta.Type | Should -Be 'SharedMailbox'
        $compta.Tag | Should -Be 'Converted'
        $compta.Direct.Count | Should -Be 0
        @($compta.FullAccess) | Should -Be @($alice.Id, $carol.Id)
        $compta.AutoMap[$carol.Id] | Should -BeTrue
        @($compta.SendAs) | Should -Be @($bob.Id)
        @($compta.SendOnBehalf) | Should -Be @($alice.Id)
        # The licence is removed only once the mailbox is a SharedMailbox (otherwise it is disabled at once).
        $added = Get-E2ECallIndex "Graph POST users/$($compta.Id)/assignLicense"
        $shared = Get-E2ECallIndex '^\d+ Shared compta@'
        $removed = Get-E2ECallIndex "Graph POST users/$($compta.Id)/assignLicense" -After $added
        $shared | Should -BeGreaterThan $added
        $removed | Should -BeGreaterThan $shared
        $journal = Get-E2EJournal $context.BatchId
        $journal.Batch.status | Should -Be 'Complete'
        @($journal.Items | ForEach-Object { $_.status }) | Should -Be @('Done', 'Done', 'Done')
        $context.NextSteps -join ' ' | Should -Match "Recover -Batch $($context.BatchId)"
    }

    It 'retries the temporary licence refused just after the usageLocation change' {
        $global:PraFake.Faults.LicenceAssign400 = 2
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($global:PraFake.Calls | Where-Object { $_ -match "Graph POST users/$($compta.Id)/assignLicense" }).Count | Should -Be 4
        $alice.CloudManaged | Should -BeFalse
    }

    It 'refuses rooms and existing cloud mailboxes, and converts the others' {
        $null = Add-PraFakeUser -Upn 'room1@contoso.com' -Kind Room
        $dave = Add-PraFakeUser -Upn 'dave@contoso.com'
        $dave.Type = 'UserMailbox'; $dave.ExchangeGuid = [guid]::NewGuid().ToString(); $dave.DirSynced = $true
        $dave.Direct.Add(@{ skuId = 'sku-e5'; disabledPlans = @() })
        $global:PraE2E.InScope.Add('room1@contoso.com'); $global:PraE2E.InScope.Add('dave@contoso.com')
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        (Get-E2ERow 'room1@contoso.com').Detail | Should -Match 'Room mailbox: not converted'
        (Get-E2ERow 'dave@contoso.com').Detail | Should -Match 'Already a cloud mailbox'
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($global:PraFake.Calls | Where-Object { $_ -match 'room1|dave' }) | Should -BeNullOrEmpty
    }

    It 'changes nothing when the licences are short' {
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert
        { Invoke-PcmConvert } | Should -Throw '*Tenant prerequisites not met*'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph (PATCH|POST|DELETE)|Set-' }) | Should -BeNullOrEmpty
    }

    It 'resumes a batch: a user still provisioning is Pending, the next run finishes him without repeating what is done' {
        $global:PraFake.Faults.NeverMailbox = @($bob.Id)
        $global:PraE2EWait = @{ MailboxTimeoutMinutes = 1 }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $batch = $context.BatchId
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Pending'
        (Get-E2EJournal $batch).Batch.status | Should -Be 'Partial'
        Close-E2E

        $global:PraFake.Faults.NeverMailbox = @()
        $global:PraE2EWait = @{}
        Start-E2E -Action Convert -Batch $batch
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com')
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        $bob.Type | Should -Be 'UserMailbox'
        (Get-E2EJournal $batch).Batch.status | Should -Be 'Complete'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph POST groups/.+/members' }).Count | Should -Be 2
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser bob@' }).Count | Should -Be 1
    }

    It 'resumes only the objects asked by -Identity or -Scope when -Batch is given, and refuses an identity that is not left to resume' {
        $global:PraFake.Faults.NeverMailbox = @($bob.Id, $compta.Id)
        $global:PraE2EWait = @{ MailboxTimeoutMinutes = 1 }
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convertBatch = $context.BatchId
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Pending'
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        Close-E2E

        $global:PraFake.Faults.NeverMailbox = @()
        $global:PraE2EWait = @{}
        Start-E2E -Action Convert -Batch $convertBatch.ToUpperInvariant() -Identity 'BOB@Contoso.com'
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com')
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        (Get-E2EJournal $convertBatch).Batch.status | Should -Be 'Partial'
        Close-E2E

        Start-E2E -Action Convert -Batch $convertBatch -Identity 'carol@contoso.com'
        { Invoke-PcmConvert } | Should -Throw "*carol@contoso.com is not in the objects of Convert batch $convertBatch still to resume*Nothing was changed*"
        Close-E2E

        Start-E2E -Action Convert -Batch $convertBatch -Scope UsersOnly
        { Invoke-PcmConvert } | Should -Throw '*No object can be converted*'
        @($context.Rows).Count | Should -Be 0
        Close-E2E

        Start-E2E -Action Convert -Batch $convertBatch -Scope SharedOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('compta@contoso.com')
        (Get-E2EJournal $convertBatch).Batch.status | Should -Be 'Complete'
    }

    It 'refuses before any change a list that names an object that is not in the snapshot, in any letter case' {
        $wave = Join-Path $global:PraE2E.Dir 'ghost.txt'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value @('ALICE@contoso.com', 'Ghost@Contoso.com', 'phantom@contoso.com')
        Start-E2E -Action Convert -CsvPath $wave
        { Invoke-PcmConvert } | Should -Throw '*2 of the 3 identities asked for are not in snapshot*Ghost@Contoso.com, phantom@contoso.com*Nothing was changed*'
        $context.NoChangeMade | Should -BeTrue
        @($global:PraFake.Calls).Count | Should -Be 0
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value @('alice@contoso.com', 'carol@contoso.com')
        Close-E2E
        Start-E2E -Action Check -CsvPath $wave
        { Invoke-PcmCheck } | Should -Throw '*1 of the 2 identities asked for is not in snapshot*carol@contoso.com*'
    }

    It 'names the batch that already converted an object asked by name, and changes nothing' {
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        Invoke-PcmConvert
        $convertBatch = $context.BatchId
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        Close-E2E
        $calls = @($global:PraFake.Calls).Count

        Start-E2E -Action Convert -Identity 'ALICE@Contoso.com'
        { Invoke-PcmConvert } | Should -Throw "*alice@contoso.com is already converted by Convert batch ${convertBatch}: nothing was changed*"
        $context.NoChangeMade | Should -BeTrue
        $context.NextOnFailure | Should -BeTrue
        ($context.NextSteps -join ' ') | Should -Match "-Action Convert -Mode Apply -Batch $convertBatch"
        ($context.NextSteps -join ' ') | Should -Match "-Action Recover -Batch $convertBatch"
        @($global:PraFake.Calls).Count | Should -Be $calls
        Close-E2E

        Start-E2E -Action Recover -Batch $convertBatch
        Invoke-PcmRecover
        Close-E2E
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        { Invoke-PcmConvert } | Should -Not -Throw
    }

    It 'asks nothing in a Preview: it never writes' {
        $script:E2ECaller.Asked.Clear()
        Start-E2E -Action Convert -Mode Preview
        Invoke-PcmConvert
        $script:E2ECaller.Asked.Count | Should -Be 0
        @($global:PraFake.Calls).Count | Should -Be 0
    }

    It 'asks the approval of the run before each change of an Apply' {
        $script:E2ECaller.Asked.Clear()
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $script:E2ECaller.Asked.Count | Should -BeGreaterThan 5
        @($script:E2ECaller.Asked | Where-Object { $_ -match '^Microsoft Graph PATCH\|users/.+/onPremisesSyncBehavior' }).Count | Should -BeGreaterThan 0
        @($script:E2ECaller.Asked | Where-Object { $_ -match '^Exchange Online: clear the ExchangeGuid' }).Count | Should -BeGreaterThan 0
    }

    It 'writes nothing when the approval is refused, and the object says why' {
        Start-E2E -Action Convert
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $false }
        try { Invoke-PcmConvert } catch { $null = $_ }
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph (PATCH|POST|DELETE)|Set-|New-|Add-' }) | Should -BeNullOrEmpty
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'Operation refused'
    }

    It 'switches a licence group synchronised from AD to the cloud, and gives it back after the last user' {
        $global:PraFake.Group.Synced = $true
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PcmConvert
        $global:PraFake.Group.CloudManaged | Should -BeTrue
        (Get-E2ECallIndex 'Graph PATCH groups/') | Should -BeLessThan (Get-E2ECallIndex 'Graph POST groups/.+/members')
        $convert = $context.BatchId
        @((Get-E2EJournal $convert).Changes | ForEach-Object { $_.kind }) | Should -Be @('GroupSoa')
        Close-E2E

        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        $global:PraFake.Group.CloudManaged | Should -BeFalse
        [bool]((Get-E2EJournal $convert).Changes[0].restored) | Should -BeTrue
    }

    It 'rolls the batch back: users on-premises under the case hold, shared mailbox inactive and recreated by Entra Connect' {
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        $aliceCloud = $alice.ExchangeGuid; $bobCloud = $bob.ExchangeGuid; $sharedCloud = $compta.ExchangeGuid; $oldShared = $compta.Id
        Close-E2E
        $global:PraFake.Calls.Clear()

        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
        $tag = Get-PraFakeHoldTag
        foreach ($pair in @(@($alice, $aliceCloud), @($bob, $bobCloud))) {
            $user = $pair[0]
            $user.Type | Should -Be 'MailUser'
            $user.ExchangeGuid | Should -Be $global:PraFake.Ad.Where({ $_.Upn -eq $user.Upn })[0].OnPremGuid
            $user.Component | Should -Be $pair[1]
            $user.CloudManaged | Should -BeFalse
            $user.Holds | Should -Contain $tag
            $global:PraFake.Group.Members | Should -Not -Contain $user.Id
            (Get-E2ECallIndex "CaseHold Add $($user.Upn)") | Should -BeGreaterThan (Get-E2ECallIndex "BackOnPremises $($user.Upn)")
        }
        # Shared: hold stamped, scheduler paused, identity deleted, inactive, purged, scheduler resumed, recreated.
        $order = @('HoldStamped compta@', 'EntraConnect Pause', "Graph DELETE users/$oldShared", 'Inactive compta@', "Graph DELETE directory/deletedItems/$oldShared", 'EntraConnect Resume', 'Recreated compta@')
        $last = -1
        foreach ($step in $order) { $index = Get-E2ECallIndex ([regex]::Escape($step)) -After $last; $index | Should -BeGreaterThan $last -Because $step; $last = $index }
        $global:PraFake.Scheduler | Should -BeTrue
        $inactive = @($global:PraFake.Inactive | Where-Object { $_.ExchangeGuid -eq $sharedCloud })
        $inactive.Count | Should -Be 1
        $inactive[0].Holds | Should -Contain $tag
        $inactive[0].Tag | Should -Be 'Converted'
        $global:PraFake.Recycle.Contains($oldShared) | Should -BeFalse
        $new = Resolve-PraFakeUser 'compta@contoso.com'
        $new.Id | Should -Not -Be $oldShared
        $new.Type | Should -Be 'MailUser'
        $new.ExchangeGuid | Should -Be $global:PraFake.Ad.Where({ $_.Upn -eq 'compta@contoso.com' })[0].OnPremGuid
        $new.Holds | Should -Not -Contain $tag
        $journal = Get-E2EJournal $context.BatchId
        $journal.Batch.status | Should -Be 'Complete'
        $journal.Batch.convert_batch | Should -Be $convert
        ($journal.Items | Where-Object { $_.kind -eq 'Shared' }).entra_id | Should -Be $new.Id
    }

    It 'lifts the case hold of a user for the switch back, then puts it back' {
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        # The user was put under the case hold during the disaster. The policy answers "failed to be
        # deployed" while recording every change: the rollback must go on.
        $global:PraFake.CasePolicy.Locations.Add($alice.Id); Step-PraFake
        $alice.Holds | Should -Contain (Get-PraFakeHoldTag)
        $global:PraFake.Faults.CaseHoldDeployError = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover 3>$null
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $alice.Type | Should -Be 'MailUser'
        $alice.Holds | Should -Contain (Get-PraFakeHoldTag)
        $order = @('CaseHold Remove alice@', 'HoldReleased alice@', 'BackOnPremises alice@', 'CaseHold Add alice@', 'HoldStamped alice@')
        $last = -1
        foreach ($step in $order) { $index = Get-E2ECallIndex ([regex]::Escape($step)) -After $last; $index | Should -BeGreaterThan $last -Because $step; $last = $index }
        @($global:PraFake.Calls | Where-Object { $_ -match 'BlockedByHold' }) | Should -BeNullOrEmpty
    }

    It 'leaves a user whose cloud mailbox carries another hold untouched' {
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $alice.Holds.Add('mbxdeadbeef00000000000000000000beef:2')
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'other hold\(s\).*mbxdeadbeef'
        $alice.CloudManaged | Should -BeTrue
        $alice.Type | Should -Be 'UserMailbox'
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Partial'
    }

    It 'never deletes a shared identity whose case hold is not stamped' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.NoHoldStamp = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'case hold is not stamped.*NOT deleted'
        (Resolve-PraFakeUser $compta.Id).Type | Should -Be 'SharedMailbox'
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect|Graph DELETE' }) | Should -BeNullOrEmpty
    }

    It 'resumes the Entra Connect scheduler even when the identity cannot be deleted' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.DeleteFails = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ECallIndex 'EntraConnect Resume') | Should -BeGreaterThan (Get-E2ECallIndex 'EntraConnect Pause')
        $global:PraFake.Scheduler | Should -BeTrue
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect Delta' }) | Should -BeNullOrEmpty
    }

    It 'never lifts the case hold of a user who is already back when Recover runs again (review 1)' {
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.NoHoldStamp = $true
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Pending', 'Pending')
        Close-E2E
        $global:PraFake.Faults.NoHoldStamp = $false
        Step-PraFake
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        @($global:PraFake.Calls | Where-Object { $_ -match 'CaseHold (Remove|Add)|HoldReleased|Graph PATCH' }) | Should -BeNullOrEmpty
        $alice.Holds | Should -Contain (Get-PraFakeHoldTag)
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        { Invoke-PcmRecover } | Should -Throw '*already rolled back*'
    }

    It 'keeps the licence group in the cloud until the users of every wave are back (review 2)' {
        $global:PraFake.Group.Synced = $true
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        Invoke-PcmConvert
        $first = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Identity 'bob@contoso.com'
        Invoke-PcmConvert
        $second = $context.BatchId
        @((Get-E2EJournal $second).Changes).Count | Should -Be 0
        Close-E2E
        Start-E2E -Action Recover -Batch $first
        Invoke-PcmRecover
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $global:PraFake.Group.CloudManaged | Should -BeTrue
        $bob.Type | Should -Be 'UserMailbox'
        Close-E2E
        Start-E2E -Action Recover -Batch $second
        Invoke-PcmRecover
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        $global:PraFake.Group.CloudManaged | Should -BeFalse
        [bool]((Get-E2EJournal $first).Changes[0].restored) | Should -BeTrue
    }

    It 'finishes a shared mailbox whose identity was deleted by an earlier run (review 4)' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.NoRecreate = $true
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Pending'
        Close-E2E
        $global:PraFake.Faults.NoRecreate = $false
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE|CaseHold' }) | Should -BeNullOrEmpty
        (Get-E2ECallIndex 'Recreated compta@') | Should -BeGreaterThan -1
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
    }

    It 'rolls back like a user a shared mailbox whose Convert stopped half-way, without deleting it (review 4)' {
        $global:PraFake.Faults.NeverMailbox = @($compta.Id)
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        $compta.CloudManaged | Should -BeTrue
        $compta.Direct.Count | Should -Be 1
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'not fully converted'
        $compta.CloudManaged | Should -BeFalse
        $compta.Type | Should -Be 'MailUser'
        $compta.ExchangeGuid | Should -Be $global:PraFake.Ad.Where({ $_.Upn -eq 'compta@contoso.com' })[0].OnPremGuid
        $compta.Direct.Count | Should -Be 0
        $compta.Holds | Should -Contain (Get-PraFakeHoldTag)
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE|EntraConnect Pause' }) | Should -BeNullOrEmpty
    }

    It 'Kiosk mode: enables Kiosk in the existing Teams licence, then puts the licence back as it was (review 5)' {
        Set-E2EUsersLicensing -Mode Kiosk
        $alice.Direct.Add(@{ skuId = 'sku-teams'; disabledPlans = @('plan-kiosk') })
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        Invoke-PcmConvert
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $alice.Type | Should -Be 'UserMailbox'
        @($alice.Direct | Where-Object { $_.skuId -eq 'sku-teams' }).Count | Should -Be 1
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $teams = @($alice.Direct | Where-Object { $_.skuId -eq 'sku-teams' })
        $teams.Count | Should -Be 1
        @($teams[0].disabledPlans) | Should -Be @('plan-kiosk')
    }

    It 'Direct mode: enables Exchange in the licence the user already had, then restores it (review 5)' {
        Set-E2EUsersLicensing -Mode Direct -Sku 'E5'
        $bob.Direct.Add(@{ skuId = 'sku-e5'; disabledPlans = @('plan-exo', 'plan-teams') })
        # Every unit is used (bob holds his): a user who already has the SKU takes no new unit.
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert -Identity 'bob@contoso.com'
        Invoke-PcmConvert
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        @(($bob.Direct | Where-Object { $_.skuId -eq 'sku-e5' }).disabledPlans) | Should -Be @('plan-teams')
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        @($bob.Direct | Where-Object { $_.skuId -eq 'sku-e5' }).Count | Should -Be 1
        @(($bob.Direct | Where-Object { $_.skuId -eq 'sku-e5' }).disabledPlans | Sort-Object) | Should -Be @('plan-exo', 'plan-teams')
    }

    It 'converts the shared mailboxes in waves of Licensing.Shared.Parallel, phase by phase for the whole wave' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $rh = Add-PraFakeUser -Upn 'rh@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com'); $global:PraE2E.InScope.Add('rh@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; Parallel = 2 }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
        foreach ($mailbox in @($compta, $accueil, $rh)) { $mailbox.Type | Should -Be 'SharedMailbox'; @($mailbox.Direct).Count | Should -Be 0 }
        # Wave 1 = compta and accueil, phase by phase: both GUIDs cleared before either becomes shared; rh starts after them.
        (Get-E2ECallIndex 'Set-MailUser accueil@') | Should -BeLessThan (Get-E2ECallIndex 'Set-Mailbox compta@contoso.com Type=Shared')
        (Get-E2ECallIndex 'Set-MailUser rh@') | Should -BeGreaterThan (Get-E2ECallIndex 'Set-Mailbox accueil@contoso.com CustomAttribute1')
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match '4/4 permission\(s\) granted'
        @($compta.SendAs) | Should -Be @($bob.Id)
    }
    It 'emits a progress event per wave of a multi-wave Convert, and a rough estimate in the plan' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $rh = Add-PraFakeUser -Upn 'rh@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com'); $global:PraE2E.InScope.Add('rh@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; Parallel = 2 }"), [Text.UTF8Encoding]::new($true))
        $events = Join-Path $global:PraE2E.Dir 'progress-convert.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        try {
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PcmConvert
        } finally { & (Get-Module PRA.Common) { $script:EventFile = '' } }
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
        $records = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json })
        $progress = @($records | Where-Object { $_.kind -eq 'progress' -and $_.phase -eq 'Shared mailboxes' })
        @($progress | ForEach-Object { $_.done }) | Should -Be @(0, 2)
        @($progress | ForEach-Object { $_.total }) | Should -Be @(3, 3)
        @($records | Where-Object { $_.kind -eq 'item' -and $_.text -match 'Rough estimate: about' }).Count | Should -Be 1
    }
    It 'converts and recovers a room mailbox like a shared one when Scope.ConvertRooms is set' {
        $room = Add-PraFakeUser -Upn 'salle-paris@contoso.com' -Kind Room -UsageLocation ''
        $global:PraE2E.InScope.Add('salle-paris@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Scope = @{ Mode = 'Auto' }", "Scope = @{ Mode = 'Auto'; IncludeRoom = `$true; ConvertRooms = `$true }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Identity 'salle-paris@contoso.com'
        Invoke-PcmConvert
        (Get-E2ERow 'salle-paris@contoso.com').FinalStatus | Should -Be 'Success'
        $room.Type | Should -Be 'SharedMailbox'
        $convert = $context.BatchId
        $roomCloud = $room.ExchangeGuid; $oldId = $room.Id
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'salle-paris@contoso.com').FinalStatus | Should -Be 'Success'
        $new = Resolve-PraFakeUser 'salle-paris@contoso.com'
        $new.Id | Should -Not -Be $oldId
        $new.Type | Should -Be 'MailUser'
        @($global:PraFake.Inactive | Where-Object { $_.ExchangeGuid -eq $roomCloud }).Count | Should -Be 1
    }
    It 'still refuses a room mailbox when Scope.ConvertRooms is not set' {
        $room = Add-PraFakeUser -Upn 'salle-paris@contoso.com' -Kind Room -UsageLocation ''
        $global:PraE2E.InScope.Add('salle-paris@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Scope = @{ Mode = 'Auto' }", "Scope = @{ Mode = 'Auto'; IncludeRoom = `$true }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Identity 'salle-paris@contoso.com'
        { Invoke-PcmConvert } | Should -Throw '*No object can be converted*'
        (Get-E2ERow 'salle-paris@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'salle-paris@contoso.com').Detail | Should -Match 'not converted'
        $room.Type | Should -Be 'MailUser'
    }
    It 'makes the waves no larger than the free units of the temporary licence' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        # Parallel = 100 by default, one free unit: one shared mailbox at a time, the unit given back between them.
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        (Get-E2ECallIndex 'Set-MailUser accueil@') | Should -BeGreaterThan (Get-E2ECallIndex 'Set-Mailbox compta@contoso.com CustomAttribute1')
        $accueil.Type | Should -Be 'SharedMailbox'
    }

    It 'keeps Licensing.Shared.KeepFree units out of the waves' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; KeepFree = 1 }"), [Text.UTF8Encoding]::new($true))
        # Two free units, one kept: one shared mailbox at a time.
        $global:PraFake.Skus['sku-e5'].Enabled = 2
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        (Get-E2ECallIndex 'Set-MailUser accueil@') | Should -BeGreaterThan (Get-E2ECallIndex 'Set-Mailbox compta@contoso.com CustomAttribute1')
        $accueil.Type | Should -Be 'SharedMailbox'
    }
    It 'stops before the first change when the kept units are all the free units' {
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; KeepFree = 1 }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Scope SharedOnly
        { Invoke-PcmConvert } | Should -Throw '*prerequisites not met*'
        $compta.Type | Should -Be 'MailUser'
    }
    It 'does not start a shared mailbox when the kept units are all that is left (not started, nothing changed)' {
        $rh = Add-PraFakeUser -Upn 'rh@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('rh@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; KeepFree = 1 }"), [Text.UTF8Encoding]::new($true))
        $global:PraFake.Skus['sku-e5'].Enabled = 2
        $global:PraFake.Faults.NeverMailbox = @($compta.Id)
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'rh@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'rh@contoso.com').Detail | Should -Match 'Licensing\.Shared\.KeepFree'
        (Get-E2ERow 'rh@contoso.com').Detail | Should -Match "Convert -Batch $($context.BatchId)"
        $rh.CloudManaged | Should -BeFalse
        @($global:PraFake.Calls | Where-Object { $_ -match "rh@|$($rh.Id)" }) | Should -BeNullOrEmpty
    }
    It 'rolls back with Recover.WaveSize, independent from Licensing.Shared.Parallel' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; Parallel = 1 }").Replace("Environment = 'TEST'", "Environment = 'TEST'`r`n    Recover = @{ WaveSize = 2 }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        (Get-E2ECallIndex 'Set-MailUser accueil@') | Should -BeGreaterThan (Get-E2ECallIndex 'Set-Mailbox compta@contoso.com CustomAttribute1')
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect (Pause|Resume)' }).Count | Should -Be 2
    }
    It 'rolls back in waves of Licensing.Shared.Parallel when Recover.WaveSize is not set' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; Parallel = 1 }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect (Pause|Resume)' }).Count | Should -Be 4
    }
    It 'rolls back in waves: one delta cycle for the users of a wave, one scheduler pause for the shared mailboxes of a wave' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        $oldIds = @($compta.Id, $accueil.Id)
        Close-E2E
        $global:PraFake.Calls.Clear(); $global:PraFake.CaseHoldCalls = 0
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success', 'Success')
        $calls = $global:PraFake.Calls
        # Users: both back to AD before the one delta cycle of their wave; one case hold call for both.
        $pause = Get-E2ECallIndex 'EntraConnect Pause'
        $resume = Get-E2ECallIndex 'EntraConnect Resume'
        @($calls[0..$pause] | Where-Object { $_ -match 'EntraConnect Delta' }).Count | Should -Be 1
        (Get-E2ECallIndex "Graph PATCH users/$($bob.Id)/onPremisesSyncBehavior") | Should -BeLessThan (Get-E2ECallIndex 'EntraConnect Delta')
        # Shared: the scheduler paused once, both identities deleted and purged meanwhile, then one delta cycle.
        @($calls | Where-Object { $_ -match 'EntraConnect (Pause|Resume)' }).Count | Should -Be 2
        foreach ($id in $oldIds) {
            (Get-E2ECallIndex "Graph DELETE users/$id") | Should -BeGreaterThan $pause
            (Get-E2ECallIndex "Graph DELETE directory/deletedItems/$id") | Should -BeLessThan $resume
        }
        @($calls | Where-Object { $_ -match 'EntraConnect Delta' }).Count | Should -Be 2
        $global:PraFake.CaseHoldCalls | Should -Be 2
        $global:PraFake.Scheduler | Should -BeTrue
        @($global:PraFake.Inactive).Count | Should -Be 2
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
    }
    It 'emits a progress event per wave of a multi-wave Recover (users and shared), and a rough estimate in the plan' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; Parallel = 1 }"), [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $events = Join-Path $global:PraE2E.Dir 'progress-recover.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        try {
            Start-E2E -Action Recover -Batch $convert
            $context.UserWaveSize = 1
            Invoke-PcmRecover
        } finally { & (Get-Module PRA.Common) { $script:EventFile = '' } }
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success', 'Success')
        $records = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json })
        $userProgress = @($records | Where-Object { $_.kind -eq 'progress' -and $_.phase -eq 'Users' })
        @($userProgress | ForEach-Object { $_.done }) | Should -Be @(0, 1)
        @($userProgress | ForEach-Object { $_.total }) | Should -Be @(2, 2)
        $sharedProgress = @($records | Where-Object { $_.kind -eq 'progress' -and $_.phase -eq 'Shared mailboxes' })
        @($sharedProgress | ForEach-Object { $_.done }) | Should -Be @(0, 1)
        @($sharedProgress | ForEach-Object { $_.total }) | Should -Be @(2, 2)
        @($records | Where-Object { $_.kind -eq 'item' -and $_.text -match 'Rough estimate: about' }).Count | Should -Be 1
    }

    It 'opens a new case hold policy in the same case when the ones of the tool are full (Retention.HoldPolicyLimit)' {
        $text = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace("Retention = @{ HoldPolicy = 'PRA-HOLD' }", "Retention = @{ HoldPolicy = 'PRA-HOLD'; HoldPolicyLimit = 1 }"), [Text.UTF8Encoding]::new($true))
        # A policy of the tool being deleted (still listed for hours): left out, its name not used again.
        $global:PraFake.CasePolicies.Add(@{ Name = 'PRA-HOLD-02'; Guid = [guid]::NewGuid().ToString(); CaseId = 'case-pra'; Mode = 'PendingDeletion'; Locations = [System.Collections.Generic.List[string]]::new() })
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        $sharedCloud = $compta.ExchangeGuid
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
        $policies = @($global:PraFake.CasePolicies | Where-Object { -not $_.ContainsKey('Mode') })
        @($policies | ForEach-Object { $_.Name }) | Should -Be @('PRA-HOLD', 'PRA-HOLD-03', 'PRA-HOLD-04')
        @($policies | ForEach-Object { $_.Locations.Count }) | Should -Be @(1, 1, 1)
        @($policies | ForEach-Object { $_.CaseId } | Select-Object -Unique) | Should -Be @('case-pra')
        @($global:PraFake.CaseRules) | Should -Be @('PRA-HOLD-03/PRA-HOLD-03-Rule', 'PRA-HOLD-04/PRA-HOLD-04-Rule')
        $alice.Holds | Should -Contain (Get-PraFakeHoldTag)
        $bob.Holds | Should -Contain (Get-PraFakeHoldTag 'PRA-HOLD-03')
        (Get-E2ERow 'bob@contoso.com').Holds | Should -Be (Get-PraFakeHoldTag 'PRA-HOLD-03')
        @($global:PraFake.Inactive | Where-Object { $_.ExchangeGuid -eq $sharedCloud })[0].Holds | Should -Contain (Get-PraFakeHoldTag 'PRA-HOLD-04')
        # A later run finds the policies of the tool again (Get-CaseHoldPolicy -Case lists them without locations).
        $set = Get-PcmHoldPolicySet -Policy 'PRA-HOLD' -Limit 1
        @($set.Policies | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }) | Should -Be @('PRA-HOLD=1', 'PRA-HOLD-03=1', 'PRA-HOLD-04=1')
        $set.Tags.Contains((Get-PraFakeHoldTag 'PRA-HOLD-04')) | Should -BeTrue
    }

    It 'never deletes the identity of a shared mailbox whose case hold was refused' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.CaseHoldFails = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'case hold policy PRA-HOLD'
        $compta.Type | Should -Be 'SharedMailbox'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE|EntraConnect' }) | Should -BeNullOrEmpty
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Partial'
    }

    It 'does not start a shared mailbox when the only temporary licence is kept by one that failed (review 6)' {
        $rh = Add-PraFakeUser -Upn 'rh@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('rh@contoso.com')
        Write-E2ESnapshot
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        $global:PraFake.Faults.NeverMailbox = @($compta.Id)
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'rh@contoso.com').Detail | Should -Match 'no free unit'
        $rh.CloudManaged | Should -BeFalse
        $rh.ExchangeGuid | Should -Be $global:PraFake.Ad.Where({ $_.Upn -eq 'rh@contoso.com' })[0].OnPremGuid
        @($global:PraFake.Calls | Where-Object { $_ -match "rh@|$($rh.Id)" }) | Should -BeNullOrEmpty
    }

    It 'finishes the permissions of a shared mailbox already converted even with no free unit left (review 2, finding 1)' {
        $global:PraFake.Faults.SendAsFails = $true
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Pending'
        $compta.Type | Should -Be 'SharedMailbox'
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.SendAsFails = $false
        $global:PraFake.Skus['sku-e5'].Enabled = 0
        Start-E2E -Action Convert -Batch $convert
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($compta.SendAs) | Should -Be @($bob.Id)
        (Get-E2EJournal $convert).Batch.status | Should -Be 'Complete'
    }
    It 'converts and recovers a wave given by -CsvPath (and by its former name -IdentityPath)' {
        $wave = Join-Path $global:PraE2E.Dir 'wave1.txt'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value @('# wave 1', 'alice@contoso.com', 'COMPTA@contoso.com')
        Start-E2E -Action Convert -CsvPath $wave
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com', 'compta@contoso.com')
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        # bob is only a SendAs trustee of compta: his own identity is not touched.
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser bob@' -or $_ -match [regex]::Escape($bob.Id) }) | Should -BeNullOrEmpty
        $bob.CloudManaged | Should -BeFalse
        $convert = $context.BatchId
        Close-E2E

        Set-Content -LiteralPath $wave -Encoding UTF8 -Value 'Identity', 'alice@contoso.com'
        Start-E2E -Action Recover -Batch $convert -CsvPath $wave
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        $alice.Type | Should -Be 'MailUser'
        $compta.Type | Should -Be 'SharedMailbox'
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Partial'
        Close-E2E

        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('compta@contoso.com')
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
    }

    It 'resumes a Recover wave given again by -CsvPath: the objects already rolled back are skipped with their reason, not refused' {
        $wave = Join-Path $global:PraE2E.Dir 'wave-resume.csv'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value 'IDENTITY', 'alice@contoso.com', 'COMPTA@contoso.com'
        Start-E2E -Action Convert -CsvPath $wave
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $first = Join-Path $global:PraE2E.Dir 'wave-resume-first.csv'
        Set-Content -LiteralPath $first -Encoding UTF8 -Value 'Identity', 'ALICE@contoso.com'
        Start-E2E -Action Recover -Batch $convert -CsvPath $first
        Invoke-PcmRecover
        $recover = $context.BatchId
        (Get-E2EJournal $recover).Batch.status | Should -Be 'Partial'
        Close-E2E

        Start-E2E -Action Recover -Batch $convert -CsvPath $wave
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('compta@contoso.com')
        @($context.SkippedMembers) | Should -Be @('alice@contoso.com (already rolled back by this Recover batch)')
        (Get-E2EJournal $recover).Batch.status | Should -Be 'Complete'
        Close-E2E
    }
    It 'refuses in plain words a Recover list whose objects are all rolled back already, and changes nothing' {
        $wave = Join-Path $global:PraE2E.Dir 'wave-done.csv'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value 'Identity', 'alice@contoso.com', 'compta@contoso.com'
        Start-E2E -Action Convert -CsvPath $wave
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $first = Join-Path $global:PraE2E.Dir 'wave-done-first.csv'
        Set-Content -LiteralPath $first -Encoding UTF8 -Value 'Identity', 'alice@contoso.com'
        Start-E2E -Action Recover -Batch $convert -CsvPath $first
        Invoke-PcmRecover
        Close-E2E
        Start-E2E -Action Recover -Batch $convert -CsvPath $first
        { Invoke-PcmRecover } | Should -Throw '*Every object asked for is already done (already rolled back by this Recover batch)*Nothing was changed*'
        $context.NoChangeMade | Should -BeTrue
        Close-E2E
    }
    It 'resumes a Convert wave given again by -CsvPath: the objects already finished are skipped with their reason, not refused' {
        $wave = Join-Path $global:PraE2E.Dir 'wave-convert-resume.csv'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value 'IDENTITY', 'alice@contoso.com', 'COMPTA@contoso.com'
        $original = [IO.File]::ReadAllText($global:PraE2E.Config)
        [IO.File]::WriteAllText($global:PraE2E.Config, $original.Replace("Shared = @{ SkuPartNumber = 'E5' }", "Shared = @{ SkuPartNumber = 'E5'; KeepFree = 1 }"), [Text.UTF8Encoding]::new($true))
        $global:PraFake.Skus['sku-e5'].Enabled = 2
        Start-E2E -Action Convert -CsvPath $wave
        Invoke-PcmConvert
        $convert = $context.BatchId
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2EJournal $convert).Batch.status | Should -Be 'Partial'
        Close-E2E
        [IO.File]::WriteAllText($global:PraE2E.Config, $original, [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Batch $convert -CsvPath $wave
        Invoke-PcmConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($context.SkippedMembers) | Should -Be @("alice@contoso.com (already finished in Convert batch $convert)")
        (Get-E2EJournal $convert).Batch.status | Should -Be 'Complete'
    }
    It 'refuses in plain words a Convert resume whose list holds only finished objects, and changes nothing' {
        $wave = Join-Path $global:PraE2E.Dir 'wave-convert-done.csv'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value 'Identity', 'alice@contoso.com'
        Start-E2E -Action Convert -CsvPath $wave
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Batch $convert -CsvPath $wave
        { Invoke-PcmConvert } | Should -Throw '*Every object asked for is already done (already finished in Convert batch*Nothing was changed*'
        $context.NoChangeMade | Should -BeTrue
    }
    It 'stops Convert before the next object when the window asks, then the batch resumes' {
        $stop = Join-Path $global:PraE2E.Dir 'run.stop'
        $env:PRA_STOP_FILE = $stop
        try {
            $global:PraFake.Faults.StopAfter = @{ Pattern = 'Graph POST groups/.+/members'; Path = $stop }
            Start-E2E -Action Convert
            # One user per wave: the stop is taken before the next wave (500 users in a real run).
            $context.UserWaveSize = 1
            Invoke-PcmConvert
            $convert = $context.BatchId
            @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Pending', 'Pending', 'Pending')
            (Get-E2ERow 'bob@contoso.com').Detail | Should -Match 'stop requested by the operator'
            (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'stop requested by the operator'
            $bob.CloudManaged | Should -BeFalse
            @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser bob@|assignLicense' }) | Should -BeNullOrEmpty
            (Get-E2EJournal $convert).Batch.status | Should -Be 'Partial'
            $context.NextSteps -join ' ' | Should -Match "Convert -Mode Apply -Batch $convert"
            Close-E2E

            Remove-Item -LiteralPath $stop
            $global:PraFake.Faults.StopAfter = $null
            Start-E2E -Action Convert -Batch $convert
            Invoke-PcmConvert
            @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
            @($global:PraFake.Calls | Where-Object { $_ -match 'Graph POST groups/.+/members' }).Count | Should -Be 2
            (Get-E2EJournal $convert).Batch.status | Should -Be 'Complete'
        } finally { Remove-Item Env:\PRA_STOP_FILE -ErrorAction SilentlyContinue }
    }

    It 'stops Recover before the next object: the users already switched back are finished, the scheduler is never left paused' {
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $stop = Join-Path $global:PraE2E.Dir 'run.stop'
        $env:PRA_STOP_FILE = $stop
        try {
            $global:PraFake.Faults.StopAfter = @{ Pattern = "Graph PATCH users/$($alice.Id)/onPremisesSyncBehavior"; Path = $stop }
            Start-E2E -Action Recover -Batch $convert
            $context.UserWaveSize = 1
            Invoke-PcmRecover
            (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
            $alice.Type | Should -Be 'MailUser'
            (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Pending'
            $bob.Type | Should -Be 'UserMailbox'
            $bob.CloudManaged | Should -BeTrue
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Pending'
            $compta.Type | Should -Be 'SharedMailbox'
            (Get-E2ECallIndex 'EntraConnect Pause') | Should -Be -1
            (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Partial'
            Close-E2E

            Remove-Item -LiteralPath $stop
            $global:PraFake.Faults.StopAfter = $null
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com', 'compta@contoso.com')
            @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
            (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
        } finally { Remove-Item Env:\PRA_STOP_FILE -ErrorAction SilentlyContinue }
    }

    It 'writes the events the window follows: steps, lines, the plan and the result' {
        $events = Join-Path $global:PraE2E.Dir 'run.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        try {
            Start-E2E -Action Convert -Mode Preview
            Invoke-PcmConvert
            # The audit files of a real run (Initialize-PcmAudit is not called here).
            Set-Content -LiteralPath $context.LogFile -Value 'log'
            $context.TranscriptPath = Join-Path $global:PraE2E.Dir 'run.transcript.txt'; Set-Content -LiteralPath $context.TranscriptPath -Value 'transcript'; $context['_PraTranscriptCreated'] = $true
            $null = Complete-PcmRun -Context $context
        } finally { & (Get-Module PRA.Common) { $script:EventFile = '' } }
        $records = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json })
        @($records | Where-Object kind -eq 'step' | ForEach-Object { $_.title }) | Should -Contain 'Plan and confirmation'
        @($records | Where-Object { $_.kind -eq 'item' -and $_.text -match '3 object\(s\) ready' }).Count | Should -Be 1
        @($records | Where-Object { $_.kind -eq 'item' -and $_.status -eq 'Sub' -and $_.text -match 'Planned: compta@contoso.com' }).Count | Should -Be 1
        $result = $records[-1]
        $result.kind | Should -Be 'result'
        $result.action | Should -Be 'Convert'
        $result.planned | Should -Be 3
        $result.exitCode | Should -Be 0 -Because (@($result.issues) -join '; ')
    }

    It 'reads the states in bulk exactly as page by page or object by object, in a few calls per 50 objects' {
        $script:bob.Type = 'UserMailbox'; $script:bob.Holds.Add('mbx0000000000000000000000000000abc:1'); $script:bob.Component = [guid]::NewGuid().ToString()
        Start-E2E -Action Convert -Mode Preview
        $records = @($global:PraFake.Ad | ForEach-Object { [pscustomobject]@{ object_guid = $_.ObjectGuid; immutable_id = $_.ImmutableId; user_principal_name = $_.Upn; primary_smtp_address = $_.Upn; kind = $_.Kind } })
        $records += [pscustomobject]@{ object_guid = 'g-missing'; immutable_id = 'imm-nobody'; user_principal_name = 'nobody@contoso.com'; primary_smtp_address = 'nobody@contoso.com'; kind = 'User' }
        $byRow = Get-PcmCloudStateSet -Context $context -Records $records -PageThreshold 100
        $global:PraFake.PageSize = 2; $global:PraFake.FilterCalls = 0
        $byPage = Get-PcmCloudStateSet -Context $context -Records $records -PageThreshold 0
        foreach ($record in $records) { ($byPage[$record.object_guid] | ConvertTo-Json -Depth 8 -Compress) | Should -Be ($byRow[$record.object_guid] | ConvertTo-Json -Depth 8 -Compress) }
        $global:PraFake.FilterCalls | Should -Be 4
        $alice = $byPage[($global:PraFake.Ad | Where-Object Upn -eq 'alice@contoso.com').ObjectGuid]
        $alice.RecipientType | Should -Be 'MailUser'
        @($alice.Locations | Where-Object Type -eq 'ComponentShared')[0].Guid | Should -Be $script:aliceTeams
        $alice.IsCloudManaged | Should -BeFalse
        $bob = $byPage[($global:PraFake.Ad | Where-Object Upn -eq 'bob@contoso.com').ObjectGuid]
        $bob.RecipientType | Should -Be 'UserMailbox'
        @($bob.Holds) | Should -Be @('mbx0000000000000000000000000000abc:1')
        @($bob.Locations | ForEach-Object Type) | Should -Be @('Primary', 'ComponentShared')
        $byPage['g-missing'].User | Should -BeNullOrEmpty
        # 120 objects: three filters of 50 at most per Exchange Online read.
        $ids = @(1..120 | ForEach-Object { (Add-PraFakeUser -Upn "u$_@contoso.com").Id })
        $global:PraFake.FilterCalls = 0
        $exo = Get-PcmExoStateSet -Id $ids
        $exo.Count | Should -Be 120
        $global:PraFake.FilterCalls | Should -Be 9
    }
    It 'sends again the requests of a Graph batch that were throttled, and reads one object at a time when the filters fail' {
        Start-E2E -Action Convert -Mode Preview
        $records = @($global:PraFake.Ad | ForEach-Object { [pscustomobject]@{ object_guid = $_.ObjectGuid; immutable_id = $_.ImmutableId; user_principal_name = $_.Upn; primary_smtp_address = $_.Upn; kind = $_.Kind } })
        $reference = Get-PcmCloudStateSet -Context $context -Records $records
        $script:alice.CloudManaged = $true
        $global:PraFake.Faults.BatchThrottle = 3
        $throttled = Get-PcmCloudStateSet -Context $context -Records $records
        $global:PraFake.Faults.BatchThrottle | Should -Be 0
        $throttled[($global:PraFake.Ad | Where-Object Upn -eq 'alice@contoso.com').ObjectGuid].IsCloudManaged | Should -BeTrue
        $throttled[($global:PraFake.Ad | Where-Object Upn -eq 'bob@contoso.com').ObjectGuid].IsCloudManaged | Should -BeFalse
        $script:alice.CloudManaged = $false
        $global:PraFake.Faults.ExoFilterFails = $true
        $oneByOne = Get-PcmCloudStateSet -Context $context -Records $records
        foreach ($record in $records) { ($oneByOne[$record.object_guid] | ConvertTo-Json -Depth 8 -Compress) | Should -Be ($reference[$record.object_guid] | ConvertTo-Json -Depth 8 -Compress) }
    }
    #region Run-time targeting: -Group, -SearchBase, -MaxObjects --------------------------------------------------------
    It 'takes the members of a group captured by Collect (<Form>: any letter case) and says why each other member is skipped' -Tag 'Targeting' -TestCases @(
        @{ Form = 'name'; Value = 'pra-wave1' }, @{ Form = 'upper name'; Value = 'PRA-WAVE1' }, @{ Form = 'distinguished name'; Value = 'cn=PRA-wave1,ou=GROUPS,dc=CONTOSO,dc=com' },
        @{ Form = 'SMTP address'; Value = 'PRAWAVE1@CONTOSO.com' }, @{ Form = 'GUID'; Value = 'GUID' }, @{ Form = 'sAMAccountName'; Value = 'PRA-WAVE1-SAM' }) {
        param($Form, $Value)
        $guid = Add-E2EGroup
        if ($Value -eq 'GUID') { $Value = $guid.ToUpperInvariant() }
        Start-E2E -Action Convert -Mode Preview -Group $Value
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be @('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com')
        $context.GroupTarget.Origin | Should -Match '^captured by Collect, snapshot \d+$'
        $context.ObjectsText | Should -Match '^Group PRA-Wave1 \(captured by Collect, snapshot \d+\): 6 member\(s\) including nested groups, 1 nested group\(s\), 3 skipped$'
        @($context.SkippedMembers) | Should -Contain 'PC01 (a computer, not a mailbox)'
        @($context.SkippedMembers) | Should -Contain 'Jane Contact (a contact, not a mailbox)'
        @($context.SkippedMembers | Where-Object { $_ -match '^dave@contoso\.com \(not in snapshot \d+: no mailbox, or outside the scope of the Collect\)$' }).Count | Should -Be 1
        $context.NextSteps[1] | Should -Match "-Action Convert -Mode Apply -Group '$([regex]::Escape($Value))'$"
        @($global:PraFake.Calls | Where-Object { $_ -match '^\d+ (Graph (PATCH|POST|DELETE|PUT)|Set-|Add-)' }).Count | Should -Be 0
    }
    It 'leaves a mailbox of another type out of a group, with the reason, when -Scope or Scope.Include* says so' -Tag 'Targeting' {
        $null = Add-E2EGroup
        Start-E2E -Action Convert -Mode Preview -Group 'PRA-Wave1' -Scope usersonly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be @('alice@contoso.com', 'bob@contoso.com')
        @($context.SkippedMembers) | Should -Contain 'compta@contoso.com (a shared mailbox, left out by -Scope UsersOnly)'
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("Scope = @{ Mode = 'Auto' }", "Scope = @{ Mode = 'Auto'; IncludeShared = `$false }")
        [IO.File]::WriteAllText($global:PraE2E.Config, $text, [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Check -Group 'pra-wave1'
        Invoke-PcmCheck
        @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be @('alice@contoso.com', 'bob@contoso.com')
        @($context.SkippedMembers) | Should -Contain 'compta@contoso.com (a shared mailbox, left out by Scope.IncludeShared = $false)'
    }
    It 'skips a system mailbox and a mailbox listed in Scope.ExcludeSamAccountNames' -Tag 'Targeting' {
        $null = Add-E2EGroup
        $text = [IO.File]::ReadAllText($global:PraE2E.Config).Replace("Scope = @{ Mode = 'Auto' }", "Scope = @{ Mode = 'Auto'; ExcludeSamAccountNames = @('BOB') }")
        [IO.File]::WriteAllText($global:PraE2E.Config, $text, [Text.UTF8Encoding]::new($true))
        Start-E2E -Action Convert -Mode Preview -Group 'PRA-Wave1'
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be @('alice@contoso.com', 'compta@contoso.com')
        @($context.SkippedMembers) | Should -Contain 'bob@contoso.com (excluded (a system mailbox, or listed in Scope.ExcludeSamAccountNames))'
    }
    It 'refuses a group name that matches two captured groups, and lists them' -Tag 'Targeting' {
        $null = Add-E2EGroup -Name 'Wave' -Sam 'wave-one'
        $null = Add-E2EGroup -Name 'WAVE' -Sam 'wave-two' -Members @()
        Start-E2E -Action Convert -Mode Preview -Group 'wave'
        { Invoke-PcmConvert } | Should -Throw '*2 groups of snapshot*match "wave"*'
        $context.NoChangeMade | Should -BeTrue
        @($global:PraFake.Calls) | Should -BeNullOrEmpty
    }
    It 'reads the group from Microsoft Entra ID when Collect did not capture it, maps the members by immutable ID, UPN and SMTP address, and follows the pages' -Tag 'Targeting' {
        $global:PraFake.GroupPageSize = 2
        $entra = Add-PraFakeEntraGroup -Name 'Wave-Entra' -Mail 'wave-entra@contoso.com' -Nickname 'waveentra' -Members @(
            @{ User = $script:alice }
            @{ Type = 'user'; Upn = 'BOB@CONTOSO.COM'; Name = 'Bob' }
            @{ Type = 'user'; Mail = 'COMPTA@contoso.com'; Name = 'Compta' }
            @{ Type = 'device'; Name = 'Laptop1' }
            @{ Type = 'group'; Name = 'Nested' }
            @{ Type = 'user'; Upn = 'zed@contoso.com'; Name = 'Zed' })
        Start-E2E -Action Convert -Mode Preview -Group 'WAVE-ENTRA'
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be @('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com')
        $context.GroupTarget.Origin | Should -Be 'Microsoft Entra group, read now'
        $context.ObjectsText | Should -Be 'Group Wave-Entra (Microsoft Entra group, read now): 5 member(s) including nested groups, 1 nested group(s), 2 skipped'
        @($context.SkippedMembers) | Should -Contain 'Laptop1 (a device, not a mailbox)'
        @($context.SkippedMembers | Where-Object { $_ -match '^zed@contoso\.com \(not in snapshot \d+: ' }).Count | Should -Be 1
        @($global:PraFake.Reads | Where-Object { $_ -eq "Graph GET groups/$($entra.Id)/transitiveMembers" }).Count | Should -Be 3
    }
    It 'finds a Microsoft Entra group by object ID or by its name with an apostrophe, any letter case' -Tag 'Targeting' -TestCases @(@{ Form = 'ID' }, @{ Form = 'apostrophe' }, @{ Form = 'mail' }) {
        param($Form)
        $entra = Add-PraFakeEntraGroup -Name "O'Brien Team" -Mail 'obrien@contoso.com' -Nickname 'obrien' -Members @(@{ User = $script:bob })
        $value = switch ($Form) { 'ID' { $entra.Id.ToUpperInvariant() } 'apostrophe' { "o'BRIEN team" } 'mail' { 'OBRIEN@contoso.com' } }
        Start-E2E -Action Check -Group $value
        Invoke-PcmCheck
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com')
    }
    It 'refuses a group that is neither in the snapshot nor in Microsoft Entra, and a name that matches two Entra groups' -Tag 'Targeting' {
        Start-E2E -Action Convert -Mode Preview -Group 'Nobody'
        { Invoke-PcmConvert } | Should -Throw '*Group Nobody is not in snapshot * nor in Microsoft Entra*'
        $context.NoChangeMade | Should -BeTrue
        $null = Add-PraFakeEntraGroup -Name 'Twin' -Mail 'one@contoso.com'
        $null = Add-PraFakeEntraGroup -Name 'twin' -Mail 'two@contoso.com'
        Start-E2E -Action Convert -Mode Preview -Group 'TWIN'
        { Invoke-PcmConvert } | Should -Throw '*2 Microsoft Entra groups match "TWIN"*one@contoso.com*two@contoso.com*'
    }
    It 'reads a group from Microsoft Entra ID when the snapshot was written before groups were captured' -Tag 'Targeting' {
        $global:PraE2E.NoGroupTables = $true
        $null = Add-PraFakeEntraGroup -Name 'Old-Snapshot' -Members @(@{ User = $script:alice })
        Write-E2ESnapshot
        Start-E2E -Action Check -Group 'old-snapshot'
        Invoke-PcmCheck
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        [IO.File]::ReadAllText($context.LogFile) | Should -Match 'was written before groups were captured'
    }
    It 'refuses to read a Microsoft Entra group with a sign-in that holds no group read permission' -Tag 'Targeting' {
        $null = Add-PraFakeEntraGroup -Name 'Wave-Entra' -Members @(@{ User = $script:alice })
        Mock Get-PcmGraphRole { @('User.ReadWrite.All', 'Organization.Read.All') }
        Start-E2E -Action Convert -Mode Preview -Group 'Wave-Entra'
        { Invoke-PcmConvert } | Should -Throw '*cannot read groups in Microsoft Entra ID*'
        $context.NoChangeMade | Should -BeTrue
    }
    It 'looks a group captured by Collect up with -ResolveOnly (<Value>), says what a real run says, skips nothing silently and signs in nowhere' -Tag 'Targeting' -TestCases @(@{ Value = 'pra-wave1' }, @{ Value = 'PRA-WAVE1' }) {
        param($Value)
        $global:PraE2EEvents = [System.Collections.Generic.List[object]]::new()
        Mock Write-PcmEvent { $global:PraE2EEvents.Add(@{ Kind = $Kind; Data = $Data }) }
        $null = Add-E2EGroup
        Start-E2E -Action Check -Group $Value -ResolveOnly
        Invoke-PcmResolveGroup
        $context.ObjectsText | Should -Match '^Group PRA-Wave1 \(captured by Collect, snapshot \d+\): 6 member\(s\) including nested groups, 1 nested group\(s\), 3 skipped$'
        @($context.SkippedMembers) | Should -Contain 'PC01 (a computer, not a mailbox)'
        @($context.Rows).Count | Should -Be 0
        @($global:PraFake.Calls).Count | Should -Be 0
        Should -Invoke Connect-PcmCloud -Times 0 -Exactly
        $target = @($global:PraE2EEvents | Where-Object { $_.Kind -eq 'target' })
        $target.Count | Should -Be 1
        $target[0].Data.name | Should -Be 'PRA-Wave1'
        $target[0].Data.source | Should -Be 'Snapshot'
        $target[0].Data.members | Should -Be 6
        $target[0].Data.nested | Should -Be 1
        $target[0].Data.skipped | Should -Be 3
        $target[0].Data.count | Should -Be 3
        $target[0].Data.text | Should -Be $context.ObjectsText
    }
    It 'looks a Microsoft Entra group up with -ResolveOnly like a real run, and reads and changes nothing else' -Tag 'Targeting' {
        $global:PraE2EEvents = [System.Collections.Generic.List[object]]::new()
        Mock Write-PcmEvent { $global:PraE2EEvents.Add(@{ Kind = $Kind; Data = $Data }) }
        $null = Add-PraFakeEntraGroup -Name 'Wave-Entra' -Mail 'wave-entra@contoso.com' -Members @(@{ User = $script:alice }, @{ Type = 'device'; Name = 'Laptop1' }, @{ Type = 'user'; Upn = 'zed@contoso.com'; Name = 'Zed' })
        Start-E2E -Action Check -Group 'wave-ENTRA' -ResolveOnly
        Invoke-PcmResolveGroup
        $context.ObjectsText | Should -Match '^Group Wave-Entra \(Microsoft Entra group, read now\): 3 member\(s\) including nested groups, 2 skipped$'
        @($context.Rows).Count | Should -Be 0
        Should -Invoke Connect-PcmCloud -Times 1 -Exactly
        @($global:PraFake.Calls | Where-Object { $_ -match '^\d+ (Graph (PATCH|POST|DELETE|PUT)|Set-|Add-|New-)' }).Count | Should -Be 0
        $target = @($global:PraE2EEvents | Where-Object { $_.Kind -eq 'target' })[0]
        $target.Data.source | Should -Be 'Entra'
        $target.Data.members | Should -Be 3
    }
    It 'answers -ResolveOnly with an error for a group that is nowhere, and for a name that matches two groups' -Tag 'Targeting' {
        Start-E2E -Action Check -Group 'Nobody' -ResolveOnly
        { Invoke-PcmResolveGroup } | Should -Throw '*Group Nobody is not in snapshot * nor in Microsoft Entra*'
        $context.NoChangeMade | Should -BeTrue
        $null = Add-E2EGroup -Name 'Wave' -Sam 'wave-one'
        $null = Add-E2EGroup -Name 'WAVE' -Sam 'wave-two' -Members @()
        Start-E2E -Action Check -Group 'wave' -ResolveOnly
        { Invoke-PcmResolveGroup } | Should -Throw '*2 groups of snapshot*match "wave"*'
    }
    It 'finds a group with no mailbox to take with -ResolveOnly, where a real check refuses it' -Tag 'Targeting' {
        $null = Add-E2EGroup -Name 'Only-PC' -Members @(@{ Class = 'Computer'; Name = 'PC01'; Guid = [guid]::NewGuid().ToString(); Dn = 'CN=PC01,OU=Computers,DC=contoso,DC=com' })
        Start-E2E -Action Check -Group 'only-pc' -ResolveOnly
        { Invoke-PcmResolveGroup } | Should -Not -Throw
        $context.ObjectsText | Should -Match '^Group Only-PC \(captured by Collect, snapshot \d+\): 1 member\(s\) including nested groups, 1 skipped$'
        Start-E2E -Action Check -Group 'only-pc'
        { Invoke-PcmCheck } | Should -Throw '*none of its 1 member(s) can be taken*'
    }
    It 'refuses -ResolveOnly without a group, with a batch or with another action (<Name>)' -Tag 'Targeting' -TestCases @(
        @{ Name = 'no group'; Action = 'Check'; Group = ''; Batch = '' }
        @{ Name = 'a batch'; Action = 'Check'; Group = 'g'; Batch = '1a2b3c4d' }
        @{ Name = 'Convert'; Action = 'Convert'; Group = 'g'; Batch = '' }) {
        param($Name, $Action, $Group, $Batch)
        $script:BoundOptions = @()
        Start-E2E -Action $Action -Group $Group -Batch $Batch -ResolveOnly
        { Test-PcmRunOption } | Should -Throw '*-ResolveOnly is for Check with -Group*'
    }
    It 'continues an interrupted rollback with <Given>, in any letter case, in the same Recover batch' -TestCases @(
        @{ Given = 'the Convert batch ID'; Form = 'Convert'; Upper = $false }, @{ Given = 'the Convert batch ID in capitals'; Form = 'Convert'; Upper = $true }
        @{ Given = 'the rollback batch ID'; Form = 'Recover'; Upper = $false }, @{ Given = 'the rollback batch ID in capitals'; Form = 'Recover'; Upper = $true }) {
        param($Given, $Form, $Upper)
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $other = 'mbxdeadbeef00000000000000000000beef:2'
        $alice.Holds.Add($other)
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        $rollback = $context.BatchId
        $rollback | Should -Not -Be $convert
        (Get-E2EJournal $rollback).Batch.status | Should -Be 'Partial'
        Close-E2E
        $null = $alice.Holds.Remove($other)
        $asked = if ($Form -eq 'Convert') { $convert } else { $rollback }
        if ($Upper) { $asked = $asked.ToUpperInvariant() }
        Start-E2E -Action Recover -Batch $asked
        Invoke-PcmRecover
        $context.BatchId | Should -Be $rollback
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        (Get-E2EJournal $rollback).Batch.status | Should -Be 'Complete'
        $journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal -ReadOnly
        try { [int](Invoke-PcmSql $journal "SELECT COUNT(*) FROM batch WHERE action = 'Recover'" -As Scalar) | Should -Be 1 } finally { Close-PcmStore $journal }
    }
    It 'announces in its plan the steps that its Write-PcmStep and Skip-PcmStep calls draw, for every action' -Tag 'Targeting' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $global:PraE2E.Root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
        $titlesOf = {
            param([string[]]$Functions)
            $found = [System.Collections.Generic.List[string]]::new()
            foreach ($name in $Functions) {
                $function = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
                foreach ($line in ($function.Extent.Text -split "`r?`n")) {
                    if ($line -match "(Write|Skip)-PcmStep .*-Title '([^']+)'") { $found.Add($Matches[2]) }
                    elseif ($line -match 'foreach \(\$title in @\((.+?)\)\) \{ Skip-PcmStep') { foreach ($quoted in [regex]::Matches($Matches[1], "'([^']+)'")) { $found.Add($quoted.Groups[1].Value) } }
                }
            }
            return @($found | Select-Object -Unique)
        }
        $cases = @(
            @{ Action = 'Collect'; Batch = ''; Group = ''; Resolve = $false; Functions = @('Invoke-PcmCollect') }
            @{ Action = 'Check'; Batch = ''; Group = ''; Resolve = $false; Functions = @('Read-PcmSnapshotData', 'Invoke-PcmCheck') }
            @{ Action = 'Check'; Batch = '1a2b3c4d'; Group = ''; Resolve = $false; Functions = @('Invoke-PcmCheckBatch', 'Read-PcmSnapshotData') }
            @{ Action = 'Check'; Batch = ''; Group = 'g'; Resolve = $true; Functions = @('Read-PcmSnapshotData') }
            @{ Action = 'Convert'; Batch = ''; Group = ''; Resolve = $false; Functions = @('Read-PcmSnapshotData', 'Invoke-PcmConvert') }
            @{ Action = 'Recover'; Batch = ''; Group = ''; Resolve = $false; Functions = @('Read-PcmSnapshotData', 'Invoke-PcmRecover') })
        foreach ($case in $cases) {
            if ($case.Action -eq 'Collect') { $Action = 'Collect'; $Batch = ''; $ResolveOnly = $false }
            else { Start-E2E -Action $case.Action -Batch $case.Batch -Group $case.Group -ResolveOnly:$case.Resolve }
            $planned = @(Get-PcmPlanTitle)
            $drawn = @(& $titlesOf $case.Functions)
            @($planned | Sort-Object) | Should -Be @($drawn | Sort-Object) -Because ('the plan of {0} {1}' -f $case.Action, $case.Batch)
        }
    }
    It 'refuses -SearchBase when the OU holds no object, takes sub-OUs, and accepts a distinguished name or a canonical name in any letter case' -Tag 'Targeting' -TestCases @(
        @{ Base = 'OU=Paris,DC=contoso,DC=com'; Expected = @('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com') }
        @{ Base = 'ou=STAFF, ou=paris, dc=Contoso, dc=COM'; Expected = @('alice@contoso.com', 'bob@contoso.com') }
        @{ Base = 'CONTOSO.com/paris'; Expected = @('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com') }
        @{ Base = 'contoso.com/Paris/SHARED'; Expected = @('compta@contoso.com') }
        @{ Base = 'contoso.com/Paris/Sta'; Expected = @() }
        @{ Base = 'OU=Lyon,DC=contoso,DC=com'; Expected = @() }) {
        param($Base, $Expected)
        Start-E2E -Action Check -SearchBase $Base
        if ($Expected.Count) {
            Invoke-PcmCheck
            @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be $Expected
            $context.ObjectsText | Should -Match '^OU '
        } else {
            { Invoke-PcmCheck } | Should -Throw '*-SearchBase * gives no object in snapshot*Nothing was changed*'
            $context.NoChangeMade | Should -BeTrue
        }
    }
    It 'limits the objects with -MaxObjects, says so, and goes on with the next wave when the same command is run again' -Tag 'Targeting' {
        Start-E2E -Action Convert -Mode Preview -MaxObjects 1
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        $context.LeftForLater | Should -Be 2
        $context.ObjectsText | Should -Be 'limited to the first 1 object(s) (-MaxObjects), 2 left for a later run'
        $context.NextSteps[1] | Should -Match '-Action Convert -Mode Apply -MaxObjects 1$'
        Close-E2E
        Start-E2E -Action Convert -MaxObjects 1
        Invoke-PcmConvert
        $next = @($context.NextSteps)
        ($next -join "`n") | Should -Match '2 more object\(s\) were left for a later run \(-MaxObjects\)'
        ($next -join "`n") | Should -Match '-Action Convert -MaxObjects 1\r?\n'
        $first = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -MaxObjects 1 -Scope UsersOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com')
        $context.BatchId | Should -Not -Be $first
        $script:alice.Type | Should -Be 'UserMailbox'
        $script:bob.Type | Should -Be 'UserMailbox'
        $script:compta.Type | Should -Be 'MailUser'
    }
    It 'skips a member already converted by an earlier batch and names that batch' -Tag 'Targeting' {
        $null = Add-E2EGroup
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        Invoke-PcmConvert
        $first = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Group 'PRA-Wave1' -Scope UsersOnly
        Invoke-PcmConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com')
        @($context.SkippedMembers) | Should -Contain "alice@contoso.com (already converted by Convert batch $first)"
    }
    It 'rolls back a wave of a Convert batch given by -Group: the other objects of the batch stay, a member outside the batch is skipped, and a group with no object of the batch is refused' -Tag 'Targeting' {
        $null = Add-E2EGroup -Name 'Back-First' -Members @(
            @{ Class = 'User'; Name = 'alice'; Upn = 'alice@contoso.com'; Dn = 'CN=alice' }
            @{ Class = 'User'; Name = 'carol'; Upn = 'carol@contoso.com'; Dn = 'CN=carol'; Guid = [guid]::NewGuid().ToString() })
        $null = Add-E2EGroup -Name 'Only-Alice' -Members @(@{ Class = 'User'; Name = 'alice'; Upn = 'alice@contoso.com'; Dn = 'CN=alice' })
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $convert -Group 'back-FIRST'
        Invoke-PcmRecover
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        $script:alice.Type | Should -Be 'MailUser'
        $script:bob.Type | Should -Be 'UserMailbox'
        @($context.SkippedMembers | Where-Object { $_ -match '^carol@contoso\.com \(not in snapshot' }).Count | Should -Be 1
        $context.ObjectsText | Should -Match '^Group Back-First \(captured by Collect, snapshot \d+\): 2 member\(s\)'
        Close-E2E
        # alice is rolled back: a group with only her has nothing left in this batch.
        Start-E2E -Action Recover -Batch $convert -Group 'Only-Alice'
        { Invoke-PcmRecover } | Should -Throw '*none of its 1 member(s) can be taken*Nothing was changed*'
        @($context.SkippedMembers) | Should -Contain 'alice@contoso.com (already rolled back by this Recover batch)'
    }
    It 'narrows the resume of a Convert batch with -Group, and skips a member that is not in the batch' -Tag 'Targeting' {
        $null = Add-E2EGroup -Name 'Resume-Wave' -Members @(
            @{ Class = 'User'; Name = 'alice'; Upn = 'alice@contoso.com'; Dn = 'CN=alice' }
            @{ Class = 'User'; Name = 'bob'; Upn = 'bob@contoso.com'; Dn = 'CN=bob' })
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Mode Preview -Batch $convert -Group 'resume-wave'
        { Invoke-PcmConvert } | Should -Throw '*none of its 2 member(s) can be taken*'
        @($context.SkippedMembers) | Should -Contain ("alice@contoso.com (already finished in Convert batch $convert)")
        @($context.SkippedMembers) | Should -Contain 'bob@contoso.com (not in this batch)'
    }
    #endregion

    #region Follow-up of a batch: Check -Batch, global report -------------------------------------------------------------
    It 'follows a Convert batch: Provisioned is met object by object, read-only, and the next step is the Recover command' -Tag 'FollowUp' -TestCases @(@{ Typed = 'Provisioned' }, @{ Typed = 'PROVISIONED' }, @{ Typed = '' }) {
        param($Typed)
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        $calls = $global:PraFake.Calls.Count
        Start-E2E -Action Check -Batch $convert.ToUpperInvariant() -Expect $Typed -Once
        Invoke-PcmCheckBatch
        @($context.Rows | ForEach-Object { $_.Identity } | Sort-Object) | Should -Be @('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com')
        @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Success')
        (Get-E2ERow 'alice@contoso.com').ExchangeOnline | Should -Be 'UserMailbox'
        (Get-E2ERow 'compta@contoso.com').ExchangeOnline | Should -Be 'SharedMailbox'
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'permissions of the snapshot present'
        $global:PraFake.Calls.Count | Should -Be $calls
        $context.NextSteps[0] | Should -Match "Every object of batch $convert is provisioned"
        $context.NextSteps[1] | Should -Match "-Action Recover -Batch $convert$"
        $context.BatchId | Should -Be $convert
    }
    It 'reports an object that is not provisioned yet, then waits for it with the Polling settings (no -Once) and sees it provisioned' -Tag 'FollowUp' {
        $global:PraE2EWait = @{ MailboxTimeoutMinutes = 30 }
        $global:PraFake.Faults.NeverMailbox = @($script:alice.Id)
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Pending'
        Close-E2E
        Start-E2E -Action Check -Batch $convert -Once
        Invoke-PcmCheckBatch
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'expected Provisioned: no cloud mailbox yet \(Exchange Online shows MailUser\)'
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        $context.NextSteps[0] | Should -Match '1 object\(s\) are not provisioned yet'
        $context.NextSteps[1] | Should -Match "-Action Check -Batch $convert -Expect Provisioned -Once$|-Action Check -Batch $convert -Expect Provisioned$"
        $context.NextSteps[1] | Should -Not -Match '-Once'
        @($context.FollowUp | Where-Object { -not $_.Met }).Count | Should -Be 1
        Close-E2E
        $global:PraFake.Faults.NeverMailbox = @()
        Start-E2E -Action Check -Batch $convert
        Invoke-PcmCheckBatch
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Success')
    }
    It 'follows a Recover: Deprovisioned and Retained are not met before the rollback and met after it (Convert batch or Recover batch given)' -Tag 'FollowUp' {
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Check -Batch $convert -Expect Deprovisioned -Once
        Invoke-PcmCheckBatch
        @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Error')
        (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'not a MailUser with the on-premises ExchangeGuid'
        Close-E2E
        Start-E2E -Action Check -Batch $convert -Expect Retained -Once
        Invoke-PcmCheckBatch
        (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'no case hold of the tool'
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'no inactive mailbox recorded'
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        $recover = $context.BatchId
        Close-E2E
        foreach ($given in @($convert, $recover)) {
            Start-E2E -Action Check -Batch $given -Expect deprovisioned -Once
            Invoke-PcmCheckBatch
            @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Success') -Because (($context.Rows | ForEach-Object { $_.Detail }) -join ' | ')
            $context.NextSteps[1] | Should -Match "-Action Check -Batch $given -Expect Retained( -Once)?$"
            Close-E2E
            Start-E2E -Action Check -Batch $given -Expect RETAINED -Once
            Invoke-PcmCheckBatch
            @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Success') -Because (($context.Rows | ForEach-Object { $_.Detail }) -join ' | ')
            (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'case hold stamped'
            (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'inactive mailbox kept'
            Close-E2E
        }
    }
    It 'refuses -Expect Provisioned for a Recover batch and an unknown batch, and narrows the follow-up with -Group' -Tag 'FollowUp' {
        $null = Add-E2EGroup -Name 'Follow-Wave' -Members @(@{ Class = 'User'; Name = 'alice'; Upn = 'alice@contoso.com'; Dn = 'CN=alice' }, @{ Class = 'Computer'; Name = 'PC01'; Guid = [guid]::NewGuid().ToString(); Dn = 'CN=PC01' })
        Start-E2E -Action Convert
        Invoke-PcmConvert
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PcmRecover
        $recover = $context.BatchId
        Close-E2E
        Start-E2E -Action Check -Batch $recover -Expect Provisioned -Once
        { Invoke-PcmCheckBatch } | Should -Throw "*Batch $recover is a Recover batch*"
        Close-E2E
        Start-E2E -Action Check -Batch 'deadbeef' -Once
        { Invoke-PcmCheckBatch } | Should -Throw '*Batch deadbeef not found*'
        Close-E2E
        Start-E2E -Action Check -Batch $convert -Group 'FOLLOW-wave' -Expect Deprovisioned -Once
        Invoke-PcmCheckBatch
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        @($context.SkippedMembers) | Should -Contain 'PC01 (a computer, not a mailbox)'
    }
    It 'writes the global report of the batch from the journal: every run, every object with its journey, the follow-up, and the run of a Recover wave' -Tag 'FollowUp' {
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PcmConvert
        $convert = $context.BatchId
        Write-PcmBatchRecord
        Close-E2E
        Start-E2E -Action Recover -Batch $convert -Identity 'alice@contoso.com'
        Invoke-PcmRecover
        $context.ConfigOverrides = @('Polling.IntervalSeconds: 20 -> 5')
        Write-PcmBatchRecord
        Close-E2E
        Start-E2E -Action Check -Batch $convert -Expect Provisioned -Once
        Invoke-PcmCheckBatch
        $journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal -ReadOnly
        try { $history = Get-PcmBatchHistory $journal $convert } finally { Close-PcmStore $journal }
        @($history.Runs).Count | Should -Be 2
        @($history.Batches | ForEach-Object { $_.action }) | Should -Be @('Convert', 'Recover')
        $result = Write-PcmGlobalReport -Context $context -History $history -ConvertBatchId $convert -FollowUp @($context.FollowUp)
        $result.Html | Should -Exist
        $result.Csv | Should -Exist
        $html = [IO.File]::ReadAllText($result.Html)
        $html | Should -Match 'Batch [0-9a-f]{8} from end to end'
        $html | Should -Match 'Polling.IntervalSeconds: 20 -&gt; 5'
        $html | Should -Match 'Follow-up check of this run'
        $html | Should -Not -Match 'https?://'
        $data = Read-PcmGlobalReportData -Html $html
        $alice = [array]::IndexOf(@($data.Objects | ForEach-Object { $_.Identity }), 'alice@contoso.com')
        $alice | Should -BeGreaterOrEqual 0
        @($data.Events | Where-Object { $_.Object -eq $alice -and $_.Step -eq 'MailboxReady' }).Count | Should -BeGreaterThan 0
        $rows = @(Import-Csv -LiteralPath $result.Csv -Delimiter ';')
        $rows.Count | Should -Be 2
        ($rows | Where-Object Identity -eq 'alice@contoso.com').Verdict | Should -Be 'Rolled back'
        ($rows | Where-Object Identity -eq 'bob@contoso.com').Verdict | Should -Be 'Converted'
        @(Import-Csv -LiteralPath $result.EventsCsv -Delimiter ';' | Where-Object { $_.Identity -eq 'alice@contoso.com' -and $_.Step -eq 'MailboxReady' }).Count | Should -BeGreaterThan 0
        $result.Html | Should -Match "_batch$convert`_global\.html$"
    }
    It 'records the run in the journal at the end of an Apply, and adds the global report path to the next steps' -Tag 'FollowUp' {
        Start-E2E -Action Convert -Scope UsersOnly -MaxObjects 1
        Invoke-PcmConvert
        $convert = $context.BatchId
        $context.Rows.Count | Should -Be 1
        Write-PcmBatchRecord
        $context.GlobalReportPath | Should -Match "_batch$convert`_global\.html$" -Because (@(Get-Content -LiteralPath $context.LogFile | Where-Object { $_ -match 'WARN|ERROR' }) -join ' | ')
        $context.GlobalReportPath | Should -Exist
        @($context.NextSteps)[-4] | Should -Match "^Global report of batch $convert"
        @($context.NextSteps)[-3] | Should -Be $context.GlobalReportPath
        @($context.NextSteps)[-2] | Should -Match "^Events of batch $convert"
        @($context.NextSteps)[-1] | Should -Match "_batch$convert`_global_events\.csv$"
        @($context.NextSteps)[-1] | Should -Exist
        Close-E2E
        $journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal -ReadOnly
        try {
            $run = @(Invoke-PcmSql $journal 'SELECT * FROM batch_run WHERE batch_id = @b' @{ b = $convert } -As Rows)
            $run.Count | Should -Be 1
            $run[0].action | Should -Be 'Convert'
            $run[0].source_text | Should -Match 'every object of the snapshot .* UsersOnly .* first 1'
            $run[0].exit_code | Should -Be 0
        } finally { Close-PcmStore $journal }
    }
    #endregion
}
