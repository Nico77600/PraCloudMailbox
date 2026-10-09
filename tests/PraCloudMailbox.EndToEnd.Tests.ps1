<#
.SYNOPSIS
    PRA Cloud Mailbox - end-to-end tests of Convert and Recover against an in-memory tenant (tests\FakeTenant.ps1).
.DESCRIPTION
    Runs the real functions of Invoke-PraCloudMailbox.ps1 (loaded without its main block) and of the modules, under
    StrictMode Latest, against a fake Microsoft 365 tenant that reproduces the behaviours measured in the lab. Only
    the connections, the Graph roles of the token, Entra Connect and the waits are replaced. PowerShell 7 only
    (Convert and Recover run in PowerShell 7).
.NOTES
    Author  : Nicolas Fabert
#>
#Requires -Version 5.1

Describe 'Convert and Recover end to end (fake tenant)' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
    BeforeAll {
        $root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
        Import-Module (Join-Path $root 'module\PRA2.Common.psm1') -Force
        Import-Module (Join-Path $root 'module\PRA2.Store.psm1') -Force
        Import-Module (Join-Path $root 'module\PRA2.Cloud.psm1') -Force
        . (Join-Path $PSScriptRoot 'FakeTenant.ps1')
        # The functions of the entry script, without its main block (which re-imports the modules and would drop the mocks).
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
        foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] })) {
            . ([scriptblock]::Create($function.Extent.Text))
        }
        # Short waits on demand ($global:PraE2EWait), otherwise the configured minutes.
        function Get-PraWaitSeconds {
            param([string]$Name)
            if ($global:PraE2EWait -and $global:PraE2EWait.ContainsKey($Name)) { return $global:PraE2EWait[$Name] }
            return [int]$context.Config.Polling[$Name] * 60
        }
        $global:PraE2E = @{ Root = $root }

        function Write-E2ESnapshot {
            <# A new complete snapshot of the objects in scope, as Collect would write it. #>
            $fake = $global:PraFake
            $byUpn = @{}; foreach ($ad in $fake.Ad) { $byUpn[$ad.Upn] = $ad }
            $cn = Open-Pra2Store -Path $global:PraE2E.Db -Root $global:PraE2E.Root
            try {
                $id = New-Pra2Snapshot $cn @{ RunId = [guid]::NewGuid().ToString('N'); Environment = 'TEST'; ToolVersion = 'test'; ExchangeServer = 'EX1'; ScopeJson = '{}' }
                $rows = @(foreach ($upn in $global:PraE2E.InScope) {
                        $ad = $byUpn[$upn]
                        [ordered]@{ object_guid = $ad.ObjectGuid; kind = $ad.Kind; user_principal_name = $upn; primary_smtp_address = $upn; sam_account_name = ($upn -split '@')[0]
                            exchange_guid = $ad.OnPremGuid; immutable_id = $ad.ImmutableId; recipient_type_details = "$($ad.Kind)Mailbox" }
                    })
                $null = Add-Pra2Row $cn mailbox $id $rows
                $shared = $byUpn['compta@contoso.com']
                $user = { param($upn, $right, $automap) [ordered]@{ mailbox_guid = $shared.ObjectGuid; access_right = $right; trustee = "CONTOSO\$(($upn -split '@')[0])"; trustee_kind = 'User'
                        trustee_upn = $upn; trustee_guid = $byUpn[$upn].ObjectGuid; trustee_name = ($upn -split '@')[0]; resolved = 1; auto_mapping = $automap } }
                $permissions = @(
                    (& $user 'alice@contoso.com' 'FullAccess' 1)
                    [ordered]@{ mailbox_guid = $shared.ObjectGuid; access_right = 'FullAccess'; trustee = 'CONTOSO\grp-compta'; trustee_kind = 'Group'
                        trustee_upn = $null; trustee_guid = 'grp-compta'; trustee_name = 'grp-compta'; resolved = 1; auto_mapping = $null }
                    (& $user 'bob@contoso.com' 'SendAs' $null)
                    (& $user 'alice@contoso.com' 'SendOnBehalf' $null))
                $null = Add-Pra2Row $cn permission $id $permissions
                $members = @('alice@contoso.com', 'carol@contoso.com') | ForEach-Object { [ordered]@{ group_guid = 'grp-compta'; member_kind = 'User'; member_upn = $_; member_name = ($_ -split '@')[0]; depth = 1 } }
                $null = Add-Pra2Row $cn group_member $id @($members)
                Set-Pra2SnapshotStatus $cn $id Complete @{ mailbox = $rows.Count }
            } finally { Close-Pra2Store $cn }
        }

        function Start-E2E {
            <# Sets, in the test, the script-level variables that the functions of the entry script read. #>
            param([ValidateSet('Convert', 'Recover')][string]$Action, [string]$Mode = 'Apply', [string]$Batch = '', [string]$Identity = '', [string]$Scope = 'All', [string]$IdentityPath = '')
            $config = Import-PraConfiguration -Path $global:PraE2E.Config -Root $global:PraE2E.Root
            $run = @{ Root = $global:PraE2E.Root; Version = 'test'; RunId = [guid]::NewGuid().ToString('N'); StartTime = Get-Date; Action = $Action; Mode = $Mode; Phase = ''
                CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; StepIndex = 0; StepTotal = 7; Warnings = 0
                Issues = [System.Collections.Generic.List[object]]::new(); Rows = [System.Collections.Generic.List[object]]::new()
                BackupFiles = [System.Collections.Generic.List[string]]::new(); StateFiles = [System.Collections.Generic.List[string]]::new(); Excluded = [System.Collections.Generic.List[string]]::new()
                LogFolder = $config.Logging.Folder; ReportFolder = $config.Report.Folder; LogFile = (Join-Path $global:PraE2E.Dir 'run.log'); TranscriptPath = ''; TranscriptStarted = $false
                NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''; Config = $config; VerboseEnabled = $false
                BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null }
            if ($IdentityPath) { $run.Identities = @(Read-PraIdentityFile -Path $IdentityPath) }
            $values = @{ context = $run; Action = $Action; Mode = $Mode; Batch = $Batch; Identity = $Identity; IdentityPath = $IdentityPath; Scope = $Scope; Snapshot = [long]0; Force = $true; PassThru = $false; toolVersion = 'test'; caller = $null }
            foreach ($name in $values.Keys) { Set-Variable -Name $name -Value $values[$name] -Scope 1 }
        }

        function Get-E2ERow { param([string]$Identity) @($context.Rows | Where-Object { $_.Identity -eq $Identity }) | Select-Object -First 1 }
        function Get-E2ECallIndex {
            <# Position of the first call that matches, after -After (-1 when absent). #>
            param([string]$Pattern, [int]$After = -1)
            $calls = $global:PraFake.Calls
            for ($i = $After + 1; $i -lt $calls.Count; $i++) { if ($calls[$i] -match $Pattern) { return $i } }
            return -1
        }
        function Get-E2EJournal {
            param([string]$BatchId)
            $journal = Open-Pra2Store -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal -ReadOnly
            try { return [pscustomobject]@{ Batch = (Get-Pra2Batch $journal -Id $BatchId); Items = @(Get-Pra2BatchItem $journal $BatchId); Changes = @(Get-Pra2TenantChange $journal $BatchId) } }
            finally { Close-Pra2Store $journal }
        }
        function Close-E2E { if ($context -and $context.Journal) { Close-Pra2Store $context.Journal; $context.Journal = $null } }
        function Set-E2EUsersLicensing {
            param([ValidateSet('Kiosk', 'Direct')][string]$Mode, [string]$Sku = '')
            $text = [IO.File]::ReadAllText($global:PraE2E.Config)
            $text = $text.Replace("Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }", "Users = @{ Mode = '$Mode'; SkuPartNumber = '$Sku' }")
            [IO.File]::WriteAllText($global:PraE2E.Config, $text, [Text.UTF8Encoding]::new($true))
        }

        Mock Connect-Pra2Cloud { 'Microsoft Graph and Exchange Online (fake tenant)' }
        Mock Connect-Pra2Compliance { }
        Mock Get-Pra2GraphRole { @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
        Mock Invoke-Pra2EntraConnect { Invoke-PraFakeEntraConnect -Operation $Operation }
        Mock Wait-Pra2Condition { Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
        Mock Wait-Pra2Condition -ModuleName PRA2.Cloud { Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
        Mock Start-Sleep { Step-PraFake }
        Mock Start-Sleep -ModuleName PRA2.Cloud { Step-PraFake }
        Mock Write-Host -ModuleName PRA2.Common { }
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
        $script:alice = Add-PraFakeUser -Upn 'alice@contoso.com' -Component ([guid]::NewGuid().ToString())
        $script:aliceTeams = $script:alice.Component
        $script:bob = Add-PraFakeUser -Upn 'bob@contoso.com'
        $script:carol = Add-PraFakeUser -Upn 'carol@contoso.com'
        $script:compta = Add-PraFakeUser -Upn 'compta@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope = [System.Collections.Generic.List[string]]@('alice@contoso.com', 'bob@contoso.com', 'compta@contoso.com')
        Write-E2ESnapshot
    }

    AfterEach {
        Close-E2E
        $global:PraE2EWait = @{}
    }

    AfterAll {
        foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
        Remove-Variable -Name PraFake, PraFakeCommands, PraE2E, PraE2EWait -Scope Global -ErrorAction SilentlyContinue
    }

    It 'Preview plans every object and changes nothing' {
        Start-E2E -Action Convert -Mode Preview
        Invoke-PraConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Planned', 'Planned', 'Planned')
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph (PATCH|POST|DELETE)|Set-|CaseHold|EntraConnect' }) | Should -BeNullOrEmpty
        $context.BatchId | Should -BeNullOrEmpty
    }

    It 'converts the users then the shared mailbox in the validated order' {
        Start-E2E -Action Convert
        Invoke-PraConvert
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
        # The licence is removed only once the mailbox is a SharedMailbox (lab SH2: otherwise it is disabled at once).
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
        Invoke-PraConvert
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
        Invoke-PraConvert
        (Get-E2ERow 'room1@contoso.com').Detail | Should -Match 'Room mailbox: not converted'
        (Get-E2ERow 'dave@contoso.com').Detail | Should -Match 'Already a cloud mailbox'
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($global:PraFake.Calls | Where-Object { $_ -match 'room1|dave' }) | Should -BeNullOrEmpty
    }

    It 'changes nothing when the licences are short' {
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert
        { Invoke-PraConvert } | Should -Throw '*Tenant prerequisites not met*'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph (PATCH|POST|DELETE)|Set-' }) | Should -BeNullOrEmpty
    }

    It 'resumes a batch: a user still provisioning is Pending, the next run finishes him without repeating what is done' {
        $global:PraFake.Faults.NeverMailbox = @($bob.Id)
        $global:PraE2EWait = @{ MailboxTimeoutMinutes = 1 }
        Start-E2E -Action Convert
        Invoke-PraConvert
        $batch = $context.BatchId
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Pending'
        (Get-E2EJournal $batch).Batch.status | Should -Be 'Partial'
        Close-E2E

        $global:PraFake.Faults.NeverMailbox = @()
        $global:PraE2EWait = @{}
        Start-E2E -Action Convert -Batch $batch
        Invoke-PraConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com')
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        $bob.Type | Should -Be 'UserMailbox'
        (Get-E2EJournal $batch).Batch.status | Should -Be 'Complete'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph POST groups/.+/members' }).Count | Should -Be 2
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser bob@' }).Count | Should -Be 1
    }

    It 'switches a licence group synchronised from AD to the cloud, and gives it back after the last user' {
        $global:PraFake.Group.Synced = $true
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PraConvert
        $global:PraFake.Group.CloudManaged | Should -BeTrue
        (Get-E2ECallIndex 'Graph PATCH groups/') | Should -BeLessThan (Get-E2ECallIndex 'Graph POST groups/.+/members')
        $convert = $context.BatchId
        @((Get-E2EJournal $convert).Changes | ForEach-Object { $_.kind }) | Should -Be @('GroupSoa')
        Close-E2E

        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        $global:PraFake.Group.CloudManaged | Should -BeFalse
        [bool]((Get-E2EJournal $convert).Changes[0].restored) | Should -BeTrue
    }

    It 'rolls the batch back: users on-premises under the case hold, shared mailbox inactive and recreated by Entra Connect' {
        Start-E2E -Action Convert
        Invoke-PraConvert
        $convert = $context.BatchId
        $aliceCloud = $alice.ExchangeGuid; $bobCloud = $bob.ExchangeGuid; $sharedCloud = $compta.ExchangeGuid; $oldShared = $compta.Id
        Close-E2E
        $global:PraFake.Calls.Clear()

        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        # The user was put under the case hold during the disaster (lab: user03). The policy answers "failed to be
        # deployed" while recording every change (lab 7-8 Oct): the rollback must go on.
        $global:PraFake.CasePolicy.Locations.Add($alice.Id); Step-PraFake
        $alice.Holds | Should -Contain (Get-PraFakeHoldTag)
        $global:PraFake.Faults.CaseHoldDeployError = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover 3>$null
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
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $alice.Holds.Add('mbxdeadbeef00000000000000000000beef:2')
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'other hold\(s\).*mbxdeadbeef'
        $alice.CloudManaged | Should -BeTrue
        $alice.Type | Should -Be 'UserMailbox'
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Partial'
    }

    It 'never deletes a shared identity whose case hold is not stamped' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.NoHoldStamp = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'case hold is not stamped.*NOT deleted'
        (Resolve-PraFakeUser $compta.Id).Type | Should -Be 'SharedMailbox'
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect|Graph DELETE' }) | Should -BeNullOrEmpty
    }

    It 'resumes the Entra Connect scheduler even when the identity cannot be deleted' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.DeleteFails = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ECallIndex 'EntraConnect Resume') | Should -BeGreaterThan (Get-E2ECallIndex 'EntraConnect Pause')
        $global:PraFake.Scheduler | Should -BeTrue
        @($global:PraFake.Calls | Where-Object { $_ -match 'EntraConnect Delta' }) | Should -BeNullOrEmpty
    }

    It 'never lifts the case hold of a user who is already back when Recover runs again (review 1)' {
        Start-E2E -Action Convert -Scope UsersOnly
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.NoHoldStamp = $true
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Pending', 'Pending')
        Close-E2E
        $global:PraFake.Faults.NoHoldStamp = $false
        Step-PraFake
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        @($global:PraFake.Calls | Where-Object { $_ -match 'CaseHold (Remove|Add)|HoldReleased|Graph PATCH' }) | Should -BeNullOrEmpty
        $alice.Holds | Should -Contain (Get-PraFakeHoldTag)
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        { Invoke-PraRecover } | Should -Throw '*already rolled back*'
    }

    It 'keeps the licence group in the cloud until the users of every wave are back (review 2)' {
        $global:PraFake.Group.Synced = $true
        Start-E2E -Action Convert -Identity 'alice@contoso.com'
        Invoke-PraConvert
        $first = $context.BatchId
        Close-E2E
        Start-E2E -Action Convert -Identity 'bob@contoso.com'
        Invoke-PraConvert
        $second = $context.BatchId
        @((Get-E2EJournal $second).Changes).Count | Should -Be 0
        Close-E2E
        Start-E2E -Action Recover -Batch $first
        Invoke-PraRecover
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $global:PraFake.Group.CloudManaged | Should -BeTrue
        $bob.Type | Should -Be 'UserMailbox'
        Close-E2E
        Start-E2E -Action Recover -Batch $second
        Invoke-PraRecover
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        $global:PraFake.Group.CloudManaged | Should -BeFalse
        [bool]((Get-E2EJournal $first).Changes[0].restored) | Should -BeTrue
    }

    It 'finishes a shared mailbox whose identity was deleted by an earlier run (review 4)' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.NoRecreate = $true
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Pending'
        Close-E2E
        $global:PraFake.Faults.NoRecreate = $false
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($global:PraFake.Calls | Where-Object { $_ -match 'Graph DELETE|CaseHold' }) | Should -BeNullOrEmpty
        (Get-E2ECallIndex 'Recreated compta@') | Should -BeGreaterThan -1
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
    }

    It 'rolls back like a user a shared mailbox whose Convert stopped half-way, without deleting it (review 4)' {
        $global:PraFake.Faults.NeverMailbox = @($compta.Id)
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PraConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        $compta.CloudManaged | Should -BeTrue
        $compta.Direct.Count | Should -Be 1
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        Invoke-PraConvert
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $alice.Type | Should -Be 'UserMailbox'
        @($alice.Direct | Where-Object { $_.skuId -eq 'sku-teams' }).Count | Should -Be 1
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
        $teams = @($alice.Direct | Where-Object { $_.skuId -eq 'sku-teams' })
        $teams.Count | Should -Be 1
        @($teams[0].disabledPlans) | Should -Be @('plan-kiosk')
    }

    It 'Direct mode: enables Exchange in the licence the user already had, then restores it (review 5)' {
        Set-E2EUsersLicensing -Mode Direct -Sku 'E5'
        $bob.Direct.Add(@{ skuId = 'sku-e5'; disabledPlans = @('plan-exo', 'plan-teams') })
        # Every unit is used (bob holds his): a user who already has the SKU takes no new unit (lab 8 Oct).
        $global:PraFake.Skus['sku-e5'].Enabled = 1
        Start-E2E -Action Convert -Identity 'bob@contoso.com'
        Invoke-PraConvert
        (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        @(($bob.Direct | Where-Object { $_.skuId -eq 'sku-e5' }).disabledPlans) | Should -Be @('plan-teams')
        $convert = $context.BatchId
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        Invoke-PraConvert
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
        & (Get-Module PRA2.Common) { $script:EventFile = $args[0] } $events
        try {
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PraConvert
        } finally { & (Get-Module PRA2.Common) { $script:EventFile = '' } }
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
        Invoke-PraConvert
        (Get-E2ERow 'salle-paris@contoso.com').FinalStatus | Should -Be 'Success'
        $room.Type | Should -Be 'SharedMailbox'
        $convert = $context.BatchId
        $roomCloud = $room.ExchangeGuid; $oldId = $room.Id
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        { Invoke-PraConvert } | Should -Throw '*No object can be converted*'
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
        Invoke-PraConvert
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        (Get-E2ECallIndex 'Set-MailUser accueil@') | Should -BeGreaterThan (Get-E2ECallIndex 'Set-Mailbox compta@contoso.com CustomAttribute1')
        $accueil.Type | Should -Be 'SharedMailbox'
    }

    It 'rolls back in waves: one delta cycle for the users of a wave, one scheduler pause for the shared mailboxes of a wave' {
        $accueil = Add-PraFakeUser -Upn 'accueil@contoso.com' -Kind Shared -UsageLocation ''
        $global:PraE2E.InScope.Add('accueil@contoso.com')
        Write-E2ESnapshot
        Start-E2E -Action Convert
        Invoke-PraConvert
        $convert = $context.BatchId
        $oldIds = @($compta.Id, $accueil.Id)
        Close-E2E
        $global:PraFake.Calls.Clear(); $global:PraFake.CaseHoldCalls = 0
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $events = Join-Path $global:PraE2E.Dir 'progress-recover.events.jsonl'
        & (Get-Module PRA2.Common) { $script:EventFile = $args[0] } $events
        try {
            Start-E2E -Action Recover -Batch $convert
            $context.UserWaveSize = 1
            Invoke-PraRecover
        } finally { & (Get-Module PRA2.Common) { $script:EventFile = '' } }
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
        # A policy of the tool being deleted (still listed for hours, lab 9 Oct): left out, its name not used again.
        $global:PraFake.CasePolicies.Add(@{ Name = 'PRA-HOLD-02'; Guid = [guid]::NewGuid().ToString(); CaseId = 'case-pra'; Mode = 'PendingDeletion'; Locations = [System.Collections.Generic.List[string]]::new() })
        Start-E2E -Action Convert
        Invoke-PraConvert
        $convert = $context.BatchId
        $sharedCloud = $compta.ExchangeGuid
        Close-E2E
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        $set = Get-Pra2HoldPolicySet -Policy 'PRA-HOLD' -Limit 1
        @($set.Policies | ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }) | Should -Be @('PRA-HOLD=1', 'PRA-HOLD-03=1', 'PRA-HOLD-04=1')
        $set.Tags.Contains((Get-PraFakeHoldTag 'PRA-HOLD-04')) | Should -BeTrue
    }

    It 'never deletes the identity of a shared mailbox whose case hold was refused' {
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.CaseHoldFails = $true
        $global:PraFake.Calls.Clear()
        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
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
        Invoke-PraConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
        (Get-E2ERow 'rh@contoso.com').Detail | Should -Match 'no free unit'
        $rh.CloudManaged | Should -BeFalse
        $rh.ExchangeGuid | Should -Be $global:PraFake.Ad.Where({ $_.Upn -eq 'rh@contoso.com' })[0].OnPremGuid
        @($global:PraFake.Calls | Where-Object { $_ -match "rh@|$($rh.Id)" }) | Should -BeNullOrEmpty
    }

    It 'finishes the permissions of a shared mailbox already converted even with no free unit left (review 2, finding 1)' {
        $global:PraFake.Faults.SendAsFails = $true
        Start-E2E -Action Convert -Scope SharedOnly
        Invoke-PraConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Pending'
        $compta.Type | Should -Be 'SharedMailbox'
        $convert = $context.BatchId
        Close-E2E
        $global:PraFake.Faults.SendAsFails = $false
        $global:PraFake.Skus['sku-e5'].Enabled = 0
        Start-E2E -Action Convert -Batch $convert
        Invoke-PraConvert
        (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        @($compta.SendAs) | Should -Be @($bob.Id)
        (Get-E2EJournal $convert).Batch.status | Should -Be 'Complete'
    }
    It 'converts and recovers a wave given by -IdentityPath, and reports the identities it cannot find' {
        $wave = Join-Path $global:PraE2E.Dir 'wave1.txt'
        Set-Content -LiteralPath $wave -Encoding UTF8 -Value @('# wave 1', 'alice@contoso.com', 'COMPTA@contoso.com', 'ghost@contoso.com', 'alice@contoso.com')
        Start-E2E -Action Convert -IdentityPath $wave
        Invoke-PraConvert
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com', 'compta@contoso.com')
        @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
        # bob is only a SendAs trustee of compta: his own identity is not touched.
        @($global:PraFake.Calls | Where-Object { $_ -match 'Set-MailUser bob@' -or $_ -match [regex]::Escape($bob.Id) }) | Should -BeNullOrEmpty
        $bob.CloudManaged | Should -BeFalse
        @($context.WarningList | Where-Object { $_.Message -match '1 identity\(ies\) of the list not in snapshot .*ghost@contoso.com' }).Count | Should -Be 1
        $convert = $context.BatchId
        Close-E2E

        Set-Content -LiteralPath $wave -Encoding UTF8 -Value 'Identity', 'alice@contoso.com'
        Start-E2E -Action Recover -Batch $convert -IdentityPath $wave
        Invoke-PraRecover
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('alice@contoso.com')
        $alice.Type | Should -Be 'MailUser'
        $compta.Type | Should -Be 'SharedMailbox'
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Partial'
        Close-E2E

        Start-E2E -Action Recover -Batch $convert
        Invoke-PraRecover
        @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('compta@contoso.com')
        (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
    }

    It 'stops Convert before the next object when the window asks, then the batch resumes' {
        $stop = Join-Path $global:PraE2E.Dir 'run.stop'
        $env:PRA_STOP_FILE = $stop
        try {
            $global:PraFake.Faults.StopAfter = @{ Pattern = 'Graph POST groups/.+/members'; Path = $stop }
            Start-E2E -Action Convert
            # One user per wave: the stop is taken before the next wave (500 users in a real run).
            $context.UserWaveSize = 1
            Invoke-PraConvert
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
            Invoke-PraConvert
            @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success', 'Success')
            @($global:PraFake.Calls | Where-Object { $_ -match 'Graph POST groups/.+/members' }).Count | Should -Be 2
            (Get-E2EJournal $convert).Batch.status | Should -Be 'Complete'
        } finally { Remove-Item Env:\PRA_STOP_FILE -ErrorAction SilentlyContinue }
    }

    It 'stops Recover before the next object: the users already switched back are finished, the scheduler is never left paused' {
        Start-E2E -Action Convert
        Invoke-PraConvert
        $convert = $context.BatchId
        Close-E2E
        $stop = Join-Path $global:PraE2E.Dir 'run.stop'
        $env:PRA_STOP_FILE = $stop
        try {
            $global:PraFake.Faults.StopAfter = @{ Pattern = "Graph PATCH users/$($alice.Id)/onPremisesSyncBehavior"; Path = $stop }
            Start-E2E -Action Recover -Batch $convert
            $context.UserWaveSize = 1
            Invoke-PraRecover
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
            Invoke-PraRecover
            @($context.Rows | ForEach-Object { $_.Identity }) | Should -Be @('bob@contoso.com', 'compta@contoso.com')
            @($context.Rows | ForEach-Object { $_.FinalStatus }) | Should -Be @('Success', 'Success')
            (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
        } finally { Remove-Item Env:\PRA_STOP_FILE -ErrorAction SilentlyContinue }
    }

    It 'writes the events the window follows: steps, lines, the plan and the result' {
        $events = Join-Path $global:PraE2E.Dir 'run.events.jsonl'
        & (Get-Module PRA2.Common) { $script:EventFile = $args[0] } $events
        try {
            Start-E2E -Action Convert -Mode Preview
            Invoke-PraConvert
            # The audit files of a real run (Initialize-PraAudit is not called here).
            Set-Content -LiteralPath $context.LogFile -Value 'log'
            $context.TranscriptPath = Join-Path $global:PraE2E.Dir 'run.transcript.txt'; Set-Content -LiteralPath $context.TranscriptPath -Value 'transcript'; $context['_PraTranscriptCreated'] = $true
            $null = Complete-PraRun -Context $context
        } finally { & (Get-Module PRA2.Common) { $script:EventFile = '' } }
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
        $byRow = Get-Pra2CloudStateSet -Context $context -Records $records -PageThreshold 100
        $global:PraFake.PageSize = 2; $global:PraFake.FilterCalls = 0
        $byPage = Get-Pra2CloudStateSet -Context $context -Records $records -PageThreshold 0
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
        $exo = Get-Pra2ExoStateSet -Id $ids
        $exo.Count | Should -Be 120
        $global:PraFake.FilterCalls | Should -Be 9
    }
    It 'sends again the requests of a Graph batch that were throttled, and reads one object at a time when the filters fail' {
        Start-E2E -Action Convert -Mode Preview
        $records = @($global:PraFake.Ad | ForEach-Object { [pscustomobject]@{ object_guid = $_.ObjectGuid; immutable_id = $_.ImmutableId; user_principal_name = $_.Upn; primary_smtp_address = $_.Upn; kind = $_.Kind } })
        $reference = Get-Pra2CloudStateSet -Context $context -Records $records
        $script:alice.CloudManaged = $true
        $global:PraFake.Faults.BatchThrottle = 3
        $throttled = Get-Pra2CloudStateSet -Context $context -Records $records
        $global:PraFake.Faults.BatchThrottle | Should -Be 0
        $throttled[($global:PraFake.Ad | Where-Object Upn -eq 'alice@contoso.com').ObjectGuid].IsCloudManaged | Should -BeTrue
        $throttled[($global:PraFake.Ad | Where-Object Upn -eq 'bob@contoso.com').ObjectGuid].IsCloudManaged | Should -BeFalse
        $script:alice.CloudManaged = $false
        $global:PraFake.Faults.ExoFilterFails = $true
        $oneByOne = Get-Pra2CloudStateSet -Context $context -Records $records
        foreach ($record in $records) { ($oneByOne[$record.object_guid] | ConvertTo-Json -Depth 8 -Compress) | Should -Be ($reference[$record.object_guid] | ConvertTo-Json -Depth 8 -Compress) }
    }}
