<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the outcomes and of the batches of Check, Convert and Recover (Pester 5+, offline).
.DESCRIPTION
    Runs the real functions of Invoke-PraCloudMailbox.ps1 (loaded without its main block) and of the modules, under StrictMode Latest, against the in-memory tenant of
    tests\FakeTenant.ps1: a Check -Batch that follows every object of the batch, a Convert that is not resumed after its Recover, an Apply that stops half-way, the licence mode of
    the batch at Recover, a shared mailbox whose identity was deleted by an earlier run, the units of the licence of the shared mailboxes, and the smaller rules of the same functions.
    PowerShell 7 only.
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

Describe 'Outcomes and batches of Check, Convert and Recover (fake tenant)' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
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
                            alias = ($upn -split '@')[0]; display_name = $upn; hidden_from_address_lists = 1
                            name = ($upn -split '@')[0]
                            distinguished_name = 'CN={0},OU={1},OU=Paris,DC=contoso,DC=com' -f ($upn -split '@')[0], $(if ($ad.Kind -eq 'User') { 'Staff' } else { 'Shared' })
                            organizational_unit = 'contoso.com/Paris/{0}' -f $(if ($ad.Kind -eq 'User') { 'Staff' } else { 'Shared' })
                            email_addresses_json = (ConvertTo-Json -InputObject @("SMTP:$upn", "smtp:proxy-$upn", "X500:/o=CONTOSO/ou=Legacy/cn=$upn") -Compress)
                            legacy_exchange_dn = "/o=CONTOSO/ou=Old/cn=$upn"; resource_settings_json = $null }
                    })
                $null = Add-PcmRow $cn mailbox $id $rows
                $shared = $byUpn['compta@contoso.com']
                $permissions = @([ordered]@{ mailbox_guid = $shared.ObjectGuid; access_right = 'FullAccess'; trustee = 'CONTOSO\alice'; trustee_kind = 'User'
                        trustee_upn = 'alice@contoso.com'; trustee_guid = $byUpn['alice@contoso.com'].ObjectGuid; trustee_name = 'alice'; resolved = 1; auto_mapping = 1 })
                $null = Add-PcmRow $cn permission $id $permissions
                Set-PcmSnapshotStatus $cn $id Complete @{ mailbox = $rows.Count }
            } finally { Close-PcmStore $cn }
        }

        function Start-E2E {
            <# Sets, in the test, the script-level variables that the functions of the entry script read. #>
            param([ValidateSet('Check', 'Convert', 'Recover')][string]$Action, [string]$Mode = 'Apply', [string]$Batch = '', [string]$Identity = '', [string]$Scope = 'All', [string]$Expect = '', [switch]$Once)
            Close-E2E
            $config = Import-PcmConfiguration -Path $global:PraE2E.Config -Root $global:PraE2E.Root
            $run = @{ Root = $global:PraE2E.Root; Version = 'test'; RunId = [guid]::NewGuid().ToString('N'); StartTime = Get-Date; Action = $Action; Mode = $Mode; Phase = ''
                CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; StepIndex = 0; StepTotal = 7; Warnings = 0
                Issues = [System.Collections.Generic.List[object]]::new(); Rows = [System.Collections.Generic.List[object]]::new()
                BackupFiles = [System.Collections.Generic.List[string]]::new(); StateFiles = [System.Collections.Generic.List[string]]::new(); Excluded = [System.Collections.Generic.List[string]]::new()
                LogFolder = $config.Logging.Folder; ReportFolder = $config.Report.Folder; LogFile = (Join-Path $global:PraE2E.Dir 'run.log'); TranscriptPath = ''; TranscriptStarted = $false
                NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''; Config = $config; VerboseEnabled = $false
                BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null
                SkippedMembers = [System.Collections.Generic.List[string]]::new(); ObjectsText = ''; ConfigOverrides = @(); GroupTarget = $null; LeftForLater = 0; CloudDescription = ''; GlobalReportPath = '' }
            $run.Approval = New-PcmApprovalCallback $script:E2ECaller $run
            $run.ConfirmPreference = 'None'
            $run.FallbackNext = @(); $run.NextOnFailure = $false; $run.NoChangeMade = $false; $run.Cancelled = $false
            Set-PcmWriteApproval -Approval $run.Approval
            $values = @{ context = $run; Action = $Action; Mode = $Mode; Batch = $Batch; Identity = $Identity; CsvPath = ''; SearchBase = ''; Group = ''; MaxObjects = 0
                Expect = $Expect; Once = [bool]$Once; ResolveOnly = $false; BaseConfigPath = ''; Scope = $Scope; Snapshot = [long]0; Force = $true; PassThru = $false; toolVersion = 'test'; caller = $null
                effectiveMode = $Mode; ConfigPath = $global:PraE2E.Config }
            foreach ($name in $values.Keys) { Set-Variable -Name $name -Value $values[$name] -Scope 1 }
        }

        function Get-E2ERow { param([string]$Identity) @($context.Rows | Where-Object { $_.Identity -eq $Identity }) | Select-Object -First 1 }
        function Get-E2EJournal {
            param([string]$BatchId)
            $journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal -ReadOnly
            try {
                return [pscustomobject]@{ Batch = (Get-PcmBatch $journal -Id $BatchId); Items = @(Get-PcmBatchItem $journal $BatchId)
                    Events = @(Invoke-PcmSql $journal 'SELECT step, outcome, detail FROM batch_event WHERE batch_id = @b ORDER BY time_utc, rowid' @{ b = $BatchId } -As Rows) }
            }
            finally { Close-PcmStore $journal }
        }
        function Close-E2E {
            $runContext = Get-Variable context -ValueOnly -ErrorAction SilentlyContinue
            if ($runContext -and $runContext.Journal) { Close-PcmStore $runContext.Journal; $runContext.Journal = $null }
        }
        function Set-E2EConfigText {
            param([string]$Old, [string]$New)
            $text = [IO.File]::ReadAllText($global:PraE2E.Config)
            if (-not $text.Contains($Old)) { throw "The test configuration has no '$Old'." }
            [IO.File]::WriteAllText($global:PraE2E.Config, $text.Replace($Old, $New), [Text.UTF8Encoding]::new($true))
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
        foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "function:$name" -ErrorAction SilentlyContinue }
        Remove-Variable -Name PraFake, PraFakeCommands, PraFakeSetUser, PraE2E, PraE2EWait, PraSlept -Scope Global -ErrorAction SilentlyContinue
    }

    Context 'Check -Batch follows every object of the batch' {
        It 'reports an object that Convert never started as not converted, and never says that every object is provisioned' {
            $global:PraFake.Skus['sku-e5'].Enabled = 2
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            (Get-E2EJournal $convert).Batch.status | Should -Be 'Partial'
            Close-E2E
            Start-E2E -Action Check -Batch $convert -Once
            Invoke-PcmCheckBatch
            @($context.Rows).Count | Should -Be 3
            (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
            (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
            (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'expected Provisioned: not converted: Convert never started this object, nothing was changed on it \(journal: not started: no free unit of E5'
            $outcome = Get-PcmOutcome $context
            $outcome.Status | Should -Be 'Failed'
            $outcome.ExitCode | Should -Be 1
            ($context.NextSteps -join "`n") | Should -Not -Match 'Every object'
            $context.NextSteps[0] | Should -Match '1 object\(s\) are not provisioned yet'
            @($context.FollowUp | Where-Object { -not $_.Met }).Count | Should -Be 1
        }
        It 'does not say that every object is provisioned while the Convert batch is not complete' {
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            $journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal
            try { Set-PcmBatch $journal $convert 'Partial' @{} } finally { Close-PcmStore $journal }
            Close-E2E
            Start-E2E -Action Check -Batch $convert -Once
            Invoke-PcmCheckBatch
            @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Success')
            (Get-PcmOutcome $context).ExitCode | Should -Be 1
            ($context.NextSteps -join "`n") | Should -Not -Match 'Every object'
            ($context.NextSteps -join "`n") | Should -Match 'Resume the batch first'
        }
        It 'proposes no resume of a Convert batch that a Recover already rolled back, and a new Convert instead' {
            $global:PraFake.Skus['sku-e5'].Enabled = 2
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            Close-E2E
            Start-E2E -Action Check -Batch $convert -Once
            Invoke-PcmCheckBatch
            $steps = $context.NextSteps -join "`n"
            $steps | Should -Match 'Recover already rolled back objects of Convert batch'
            $steps | Should -Not -Match "-Action Convert -Mode Apply -Batch $convert"
            $steps | Should -Not -Match "-Action Convert -Batch $convert"
        }
    }

    Context 'Convert resumed after a Recover' {
        It 'is refused with a plain message and the right next step, and changes nothing' {
            $global:PraFake.Skus['sku-e5'].Enabled = 2
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            $recover = $context.BatchId
            Close-E2E
            $global:PraFake.Skus['sku-e5'].Enabled = 10
            $calls = $global:PraFake.Calls.Count
            Start-E2E -Action Convert -Batch $convert
            { Invoke-PcmConvert } | Should -Throw "Convert batch $convert cannot be resumed: Recover batch $recover (complete) already rolled back objects of it*Nothing was changed."
            $context.NoChangeMade | Should -BeTrue
            $context.NextSteps[0] | Should -Match 'run a new Convert for them'
            $context.NextSteps[1] | Should -Match '-Action Convert'
            $context.NextSteps[1] | Should -Not -Match '-Batch'
            $global:PraFake.Calls.Count | Should -Be $calls
            $script:compta.Type | Should -Be 'MailUser'
        }
        It 'is refused whatever the letter case of the batch ID' {
            $global:PraFake.Skus['sku-e5'].Enabled = 2
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            Close-E2E
            Start-E2E -Action Convert -Batch $convert.ToUpperInvariant()
            { Invoke-PcmConvert } | Should -Throw '*cannot be resumed*'
        }
    }

    Context 'A run that does not reach its end' {
        It 'leaves the rows of an Apply that are still Planned as Pending, never as a success' {
            Mock Invoke-PcmUserPlanSet { throw 'stopped half-way' }
            Start-E2E -Action Convert
            { Invoke-PcmConvert } | Should -Throw 'stopped half-way'
            @($context.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Planned')
            $outcome = Get-PcmOutcome $context
            $outcome.Counts.Pending | Should -Be 3
            $outcome.Counts.Planned | Should -Be 0
            $outcome.Counts.Success | Should -Be 0
            $outcome.Status | Should -Be 'Pending'
            $outcome.ExitCode | Should -Be 2
            @($outcome.Rows | ForEach-Object { $_.FinalStatus } | Sort-Object -Unique) | Should -Be @('Pending')
            $outcome.Rows[0].Detail | Should -Match 'the run ended before this object was done'
        }
        It 'keeps the rows as Planned when the run ended before its batch was written (a cancelled or refused Apply)' {
            Start-E2E -Action Convert
            [void]$context.Rows.Add((New-PcmRow ([pscustomobject]@{ user_principal_name = 'x@contoso.com'; kind = 'User'; primary_smtp_address = 'x@contoso.com'; object_guid = 'g'; exchange_guid = '' })))
            $context.Cancelled = $true
            $outcome = Get-PcmOutcome $context
            $outcome.Counts.Planned | Should -Be 1
            $outcome.Counts.Pending | Should -Be 0
            $outcome.Status | Should -Be 'Cancelled'
        }
        It 'keeps the rows of a Preview as Planned' {
            Start-E2E -Action Convert -Mode Preview
            Invoke-PcmConvert
            $outcome = Get-PcmOutcome $context
            $outcome.Counts.Planned | Should -Be 3
            $outcome.Counts.Pending | Should -Be 0
            $outcome.Status | Should -Be 'Planned'
            $outcome.ExitCode | Should -Be 0
        }
        It 'marks the batch that the run left open as partial, with a note, instead of running for ever' {
            Mock Invoke-PcmUserPlanSet { throw 'stopped half-way' }
            Start-E2E -Action Convert
            { Invoke-PcmConvert } | Should -Throw
            $batch = $context.BatchId
            (Get-E2EJournal $batch).Batch.status | Should -Be 'Running'
            Write-PcmBatchRecord
            $after = (Get-E2EJournal $batch).Batch
            $after.status | Should -Be 'Partial'
            $after.note | Should -Match 'ended before it finished'
        }
        It 'does not change the status of a batch that the run completed' {
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $context.RunReachedEnd = $true
            Write-PcmBatchRecord
            (Get-E2EJournal $context.BatchId).Batch.status | Should -Be 'Complete'
        }
    }

    Context 'An object matched by UPN only' {
        It 'is an error, with a plain reason, when its mailbox GUID also differs, and nothing is changed on it' {
            $ad = $global:PraFake.Ad | Where-Object { $_.Upn -eq 'bob@contoso.com' }
            $ad.ImmutableId = [Convert]::ToBase64String([guid]::NewGuid().ToByteArray())
            $ad.OnPremGuid = [guid]::NewGuid().ToString()
            Write-E2ESnapshot
            $calls = $global:PraFake.Calls.Count
            Start-E2E -Action Convert -Identity 'bob@contoso.com'
            { Invoke-PcmConvert } | Should -Throw 'No object can be converted*'
            $row = Get-E2ERow 'bob@contoso.com'
            $row.FinalStatus | Should -Be 'Error'
            $row.Detail | Should -Match 'Not converted: .*probably not this person.*found by UPN only.*Nothing is changed on it'
            $global:PraFake.Calls.Count | Should -Be $calls
            $script:bob.CloudManaged | Should -BeFalse
            $script:bob.Type | Should -Be 'MailUser'
        }
        It 'stays a warning when only one of the two differs' {
            $config = Import-PcmConfiguration -Path $global:PraE2E.Config -Root $global:PraE2E.Root
            $record = [pscustomobject]@{ kind = 'User'; object_guid = 'g'; exchange_guid = 'aaaaaaaa-0000-0000-0000-000000000001'; immutable_id = 'x'; user_principal_name = 'u@contoso.com' }
            $user = [pscustomobject]@{ id = 'id1'; onPremisesSyncEnabled = $true; onPremisesImmutableId = 'y'; usageLocation = 'FR'; assignedPlans = @(); assignedLicenses = @(); serviceProvisioningErrors = @() }
            $state = { param($matched, $guid) [pscustomobject]@{ MatchedBy = $matched; User = $user; IsCloudManaged = $false; RecipientType = 'MailUser'; ExchangeGuid = $guid; Holds = @(); Locations = @(); TagValue = '' } }
            $codes = { param($findings, $level) @($findings | Where-Object { $_.Level -eq $level } | ForEach-Object Code) }
            $upnOnly = Get-PcmReadiness -Config $config -Record $record -State (& $state 'UPN' 'aaaaaaaa-0000-0000-0000-000000000001')
            & $codes $upnOnly 'Error' | Should -Not -Contain 'ENTRA_WRONG_OBJECT'
            $guidOnly = Get-PcmReadiness -Config $config -Record $record -State (& $state 'immutableId' 'cccccccc-0000-0000-0000-000000000003')
            & $codes $guidOnly 'Error' | Should -Not -Contain 'ENTRA_WRONG_OBJECT'
            $both = Get-PcmReadiness -Config $config -Record $record -State (& $state 'UPN' 'cccccccc-0000-0000-0000-000000000003')
            & $codes $both 'Error' | Should -Contain 'ENTRA_WRONG_OBJECT'
            & $codes $both 'Warn' | Should -Contain 'GUID_MISMATCH'
        }
    }

    Context 'Convert: write failures in the middle of a wave' {
        It 'reports the user whose source of authority cannot be moved and goes on with the others' {
            $global:PraFake.Faults.SoaFails = @($script:alice.Id)
            Start-E2E -Action Convert -Scope UsersOnly
            Invoke-PcmConvert
            (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Error'
            (Get-E2ERow 'alice@contoso.com').Detail | Should -Match 'source of authority not transferred \(403\)'
            (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
            $journal = Get-E2EJournal $context.BatchId
            $journal.Batch.status | Should -Be 'Partial'
            @($journal.Items | Where-Object { $_.identity -eq 'alice@contoso.com' })[0].status | Should -Be 'Failed'
        }
        It 'puts the Exchange plan present on an object in the row once' {
            $script:bob.Direct.Add(@{ skuId = 'sku-e5'; disabledPlans = @() })
            Start-E2E -Action Convert -Identity 'bob@contoso.com'
            { Invoke-PcmConvert } | Should -Throw
            $detail = (Get-E2ERow 'bob@contoso.com').Detail
            $detail | Should -Match 'EXCHANGE|Exchange Online mailbox plan is already enabled'
            ([regex]::Matches($detail, 'An Exchange Online mailbox plan is already enabled')).Count | Should -Be 1
        }
    }

    Context 'Recover: the licence mode of the batch' {
        It 'stops before any change when Licensing.Users.Mode is no longer the mode of the batch, and says what to do' {
            Start-E2E -Action Convert -Scope UsersOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Set-E2EConfigText "Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }" "Users = @{ Mode = 'Direct'; SkuPartNumber = 'E5' }"
            $calls = $global:PraFake.Calls.Count
            Start-E2E -Action Recover -Batch $convert
            { Invoke-PcmRecover } | Should -Throw "Convert batch $convert gave the Exchange plan of its users with Licensing.Users.Mode = Group, and the configuration now says Direct*Nothing was changed."
            $context.NoChangeMade | Should -BeTrue
            $context.NextSteps[0] | Should -Match 'Set Licensing.Users.Mode back to Group'
            $context.NextSteps[1] | Should -Match "-Action Recover -Mode Apply -Batch $convert"
            $global:PraFake.Calls.Count | Should -Be $calls
            $script:alice.Type | Should -Be 'UserMailbox'
        }
        It 'refuses the Preview as well' {
            Start-E2E -Action Convert -Scope UsersOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Set-E2EConfigText "Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }" "Users = @{ Mode = 'Kiosk' }"
            Start-E2E -Action Recover -Mode Preview -Batch $convert
            { Invoke-PcmRecover } | Should -Throw '*Licensing.Users.Mode = Group, and the configuration now says Kiosk*'
        }
        It 'does not mind the letter case of the mode' {
            Start-E2E -Action Convert -Scope UsersOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Set-E2EConfigText "Users = @{ Mode = 'Group';" "Users = @{ Mode = 'gROUP';"
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Success'
            Get-PcmBatchUsersMode -Row ([pscustomobject]@{ scope_json = '{"UsersMode":"Direct"}' }) | Should -Be 'Direct'
            Get-PcmBatchUsersMode -Row ([pscustomobject]@{ scope_json = '{"Scope":"All"}' }) | Should -Be ''
            Get-PcmBatchUsersMode -Row ([pscustomobject]@{ scope_json = 'not json' }) | Should -Be ''
        }
        It 'does not refuse a Recover of shared mailboxes only, whatever the mode now' {
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Set-E2EConfigText "Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }" "Users = @{ Mode = 'Kiosk' }"
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
        }
    }

    Context 'Recover: the user wave' {
        It 'says that the cloud mailbox is not under the case hold yet when the user is not back' {
            Start-E2E -Action Convert -Scope UsersOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            $global:PraFake.Faults.NeverBack = @($script:alice.Id)
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            $detail = (Get-E2ERow 'alice@contoso.com').Detail
            (Get-E2ERow 'alice@contoso.com').FinalStatus | Should -Be 'Error'
            $detail | Should -Match 'not a MailUser with the on-premises ExchangeGuid .* yet \(Exchange Online shows UserMailbox\)'
            $detail | Should -Match 'the cloud mailbox is not under the case hold yet'
            (Get-E2ERow 'bob@contoso.com').FinalStatus | Should -Be 'Success'
        }
    }

    Context 'Recover: a shared mailbox whose identity was deleted' {
        It 'journals the deletion, and a re-run goes on with the wait for the inactive mailbox instead of saying that there is nothing to roll back' {
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            $global:PraFake.Faults.NoInactive = $true
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            $recover = $context.BatchId
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
            (Get-E2ERow 'compta@contoso.com').Detail | Should -Match 'identity deleted but mailbox .* not inactive after 15 min'
            $journal = Get-E2EJournal $recover
            @($journal.Events | Where-Object step -eq 'Deleting').Count | Should -Be 1
            @($journal.Events | Where-Object step -eq 'Deleted').Count | Should -Be 1
            $journal.Items[0].step | Should -Be 'Deleted'
            ($journal.Items[0].result_json | ConvertFrom-Json).cloudGuid | Should -Not -BeNullOrEmpty
            @($global:PraFake.Recycle.Keys).Count | Should -Be 1
            $global:PraFake.Scheduler | Should -BeTrue
            Close-E2E
            $global:PraFake.Faults.NoInactive = $false
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            $row = Get-E2ERow 'compta@contoso.com'
            $row.FinalStatus | Should -Be 'Success'
            $row.Detail | Should -Match 'recreated from AD'
            $context.BatchId | Should -Be $recover
            (Get-E2EJournal $recover).Batch.status | Should -Be 'Complete'
            @($global:PraFake.Recycle.Keys).Count | Should -Be 0
            @($global:PraFake.Inactive).Count | Should -Be 1
        }
        It 'does not take an unreadable recycle bin for a deleted identity: the step Purged is not written and the next run finishes' {
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            $global:PraFake.Faults.RecycleReadFails = $true
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            $recover = $context.BatchId
            $row = Get-E2ERow 'compta@contoso.com'
            $row.FinalStatus | Should -Be 'Error'
            $row.Detail | Should -Match 'the Entra ID recycle bin could not be read \(403\).*not confirmed as deleted permanently'
            $journal = Get-E2EJournal $recover
            @($journal.Events | Where-Object step -eq 'Purged').Count | Should -Be 0
            $journal.Items[0].step | Should -Be 'Inactive'
            @($global:PraFake.Recycle.Keys).Count | Should -Be 1
            $global:PraFake.Scheduler | Should -BeTrue
            Close-E2E
            $global:PraFake.Faults.RecycleReadFails = $false
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
            @((Get-E2EJournal $recover).Events | Where-Object step -eq 'Purged').Count | Should -Be 1
        }
    }

    Context 'Convert: the units of the licence of the shared mailboxes' {
        It 'selects the objects that need no unit even when no unit is free, and the others while a unit is free' {
            $entries = @('a', 'b', 'c') | ForEach-Object { [pscustomobject]@{ EntraId = $_ } }
            $needs = @{ a = $true; b = $false; c = $true }
            $none = Select-PcmSharedWave -Candidates $entries -NeedsUnit $needs -Free 0
            @($none.Wave | ForEach-Object EntraId) | Should -Be @('b')
            @($none.Waiting | ForEach-Object EntraId) | Should -Be @('a', 'c')
            $one = Select-PcmSharedWave -Candidates $entries -NeedsUnit $needs -Free 1
            @($one.Wave | ForEach-Object EntraId) | Should -Be @('a', 'b')
            @($one.Waiting | ForEach-Object EntraId) | Should -Be @('c')
            $negative = Select-PcmSharedWave -Candidates $entries -NeedsUnit $needs -Free -3
            @($negative.Wave | ForEach-Object EntraId) | Should -Be @('b')
            @((Select-PcmSharedWave -Candidates @() -NeedsUnit @{} -Free 5).Wave).Count | Should -Be 0
        }
        It 'waits for the unit that the previous wave gave back instead of failing the next shared mailbox' {
            $null = Add-PraFakeUser -Upn 'rh@contoso.com' -Kind Shared -UsageLocation ''
            $global:PraE2E.InScope.Add('rh@contoso.com')
            Write-E2ESnapshot
            Set-E2EConfigText "Shared = @{ SkuPartNumber = 'E5' }" "Shared = @{ SkuPartNumber = 'E5'; Parallel = 1 }"
            $global:PraFake.Skus['sku-e5'].Enabled = 1
            $global:PraFake.Faults.ReleaseLagReads = 3
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PcmConvert
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Success'
            (Get-E2ERow 'rh@contoso.com').FinalStatus | Should -Be 'Success'
            $journal = Get-E2EJournal $context.BatchId
            @($journal.Events | Where-Object step -eq 'NoFreeUnit').Count | Should -Be 0
            $journal.Batch.status | Should -Be 'Complete'
        }
        It 'still fails the shared mailbox that has no unit, once, with its reason, when no unit comes back' {
            $null = Add-PraFakeUser -Upn 'rh@contoso.com' -Kind Shared -UsageLocation ''
            $global:PraE2E.InScope.Add('rh@contoso.com')
            Write-E2ESnapshot
            Set-E2EConfigText "Shared = @{ SkuPartNumber = 'E5' }" "Shared = @{ SkuPartNumber = 'E5'; Parallel = 1 }"
            $global:PraFake.Skus['sku-e5'].Enabled = 1
            $global:PraFake.Faults.NeverMailbox = @($script:compta.Id)
            Start-E2E -Action Convert -Scope SharedOnly
            Invoke-PcmConvert
            (Get-E2ERow 'compta@contoso.com').FinalStatus | Should -Be 'Error'
            (Get-E2ERow 'rh@contoso.com').FinalStatus | Should -Be 'Error'
            (Get-E2ERow 'rh@contoso.com').Detail | Should -Match '^not started: no free unit of E5'
            @((Get-E2EJournal $context.BatchId).Events | Where-Object step -eq 'NoFreeUnit').Count | Should -Be 1
        }
    }

    Context 'The licence capacity of a licence group' {
        It 'is the smallest number of free units among its licences, not their sum' {
            $plan = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; servicePlanId = 'p-exo' })
            $skus = @(
                [pscustomobject]@{ skuId = 'big'; skuPartNumber = 'BIG'; consumedUnits = 0; prepaidUnits = [pscustomobject]@{ enabled = 100 }; servicePlans = $plan }
                [pscustomobject]@{ skuId = 'small'; skuPartNumber = 'SMALL'; consumedUnits = 8; prepaidUnits = [pscustomobject]@{ enabled = 10 }; servicePlans = $plan })
            $group = [pscustomobject]@{ displayName = 'LIC'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'big'; disabledPlans = @() }, [pscustomobject]@{ skuId = 'small'; disabledPlans = @() }) }
            $tenant = [pscustomobject]@{ Skus = $skus; Organization = $null; LicenceGroup = $group; LicenceGroupCloudManaged = $true }
            $config = Import-PcmConfiguration -Path $global:PraE2E.Config -Root $global:PraE2E.Root
            $roles = @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All')
            $fits = @(Test-PcmTenantReadiness -Config $config -Roles $roles -Tenant $tenant -UserCount 2 | Where-Object Code -eq 'LICENCE_CAPACITY')[0]
            $fits.Level | Should -Be 'Ok'
            $fits.Message | Should -Match '2 free unit\(s\) for 2 user\(s\)'
            $short = @(Test-PcmTenantReadiness -Config $config -Roles $roles -Tenant $tenant -UserCount 3 | Where-Object Code -eq 'LICENCE_CAPACITY')[0]
            $short.Level | Should -Be 'Error'
            $short.Message | Should -Match '2 free unit\(s\) for 3 user\(s\)'
        }
    }

    Context 'The index of the converted objects' {
        It 'answers from one read of the journal, whatever the letter case of the object ID' {
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Start-E2E -Action Check
            $context.Journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal
            $guid = [string](@($context.Rows) + @((Get-PcmBatchItem $context.Journal $convert)))[0].object_guid
            (Find-PcmConvertedBatch -ObjectGuid $guid.ToUpperInvariant()).id | Should -Be $convert
            Close-PcmStore $context.Journal
            (Find-PcmConvertedBatch -ObjectGuid $guid.ToLowerInvariant()).id | Should -Be $convert
            Find-PcmConvertedBatch -ObjectGuid ([guid]::NewGuid().ToString()) | Should -BeNullOrEmpty
            $context.Journal = $null
        }
        It 'leaves out an object that a Recover rolled back, and keeps an object that is not rolled back yet' {
            Start-E2E -Action Convert
            Invoke-PcmConvert
            $convert = $context.BatchId
            Close-E2E
            Start-E2E -Action Recover -Batch $convert -Mode Apply
            $context.Journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal
            $guids = @(Get-PcmBatchItem $context.Journal $convert | ForEach-Object { [string]$_.object_guid })
            @($guids | Where-Object { Find-PcmConvertedBatch -ObjectGuid $_ }).Count | Should -Be $guids.Count
            Close-E2E
            Start-E2E -Action Recover -Batch $convert
            Invoke-PcmRecover
            Close-E2E
            Start-E2E -Action Check
            $context.Journal = Open-PcmStore -Path $context.Config.Store.JournalPath -Root $global:PraE2E.Root -Kind Journal
            @($guids | Where-Object { Find-PcmConvertedBatch -ObjectGuid $_ }).Count | Should -Be 0
        }
    }

    Context 'Throttled requests of a Graph batch' {
        BeforeEach { $global:PraSlept = [System.Collections.Generic.List[int]]::new() }
        It 'waits for the Retry-After of the answer, at least the usual pause' {
            Mock Start-Sleep -ModuleName PRA.Cloud { $global:PraSlept.Add([int]$Seconds); Step-PraFake }
            $global:PraFake.Faults.BatchThrottle = 1
            $global:PraFake.Faults.BatchThrottleRetryAfter = 17
            Start-E2E -Action Check
            $answers = Invoke-PcmGraphBatch -Requests @(@{ Method = 'GET'; Url = ('/users/{0}/onPremisesSyncBehavior?$select=isCloudManaged' -f $script:alice.Id) })
            $answers[0].Ok | Should -BeTrue
            @($global:PraSlept) | Should -Contain 17
        }
        It 'keeps the usual pause when the answer gives none or a smaller one, and never waits more than two minutes' {
            Mock Start-Sleep -ModuleName PRA.Cloud { $global:PraSlept.Add([int]$Seconds); Step-PraFake }
            Start-E2E -Action Check
            $global:PraFake.Faults.BatchThrottle = 1
            $null = Invoke-PcmGraphBatch -Requests @(@{ Method = 'GET'; Url = ('/users/{0}/onPremisesSyncBehavior?$select=isCloudManaged' -f $script:alice.Id) })
            @($global:PraSlept) | Should -Be @(5)
            $global:PraSlept.Clear()
            $global:PraFake.Faults.BatchThrottle = 1
            $global:PraFake.Faults.BatchThrottleRetryAfter = 3
            $null = Invoke-PcmGraphBatch -Requests @(@{ Method = 'GET'; Url = ('/users/{0}/onPremisesSyncBehavior?$select=isCloudManaged' -f $script:alice.Id) })
            @($global:PraSlept) | Should -Be @(5)
            $global:PraSlept.Clear()
            $global:PraFake.Faults.BatchThrottle = 1
            $global:PraFake.Faults.BatchThrottleRetryAfter = 900
            $null = Invoke-PcmGraphBatch -Requests @(@{ Method = 'GET'; Url = ('/users/{0}/onPremisesSyncBehavior?$select=isCloudManaged' -f $script:alice.Id) })
            @($global:PraSlept) | Should -Be @(120)
        }
        It 'reads the Retry-After header whatever its letter case and its container' {
            $read = { param($response) & (Get-Module PRA.Cloud) { param($r) Get-PcmRetryAfterSecond -Response $r } $response }
            & $read ([pscustomobject]@{ headers = @{ 'retry-after' = '9' } }) | Should -Be 9
            & $read ([pscustomobject]@{ headers = [pscustomobject]@{ 'RETRY-AFTER' = '11' } }) | Should -Be 11
            & $read ([pscustomobject]@{ headers = @{ 'Retry-After' = 'soon' } }) | Should -Be 0
            & $read ([pscustomobject]@{ status = 429 }) | Should -Be 0
            & $read $null | Should -Be 0
        }
    }
}
