<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the follow-up of a batch (expectations, global report), of the report page, of the settings of a run (-BaseConfigPath) and of the
    interactive sign-in. Pester 5+, offline, both editions. The follow-up end to end (Check -Batch against the fake tenant) is in PraCloudMailbox.EndToEnd.Tests.ps1.
.NOTES
    Author  : Nicolas Fabert
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    Import-Module (Join-Path $script:Root 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Store.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Cloud.psm1') -Force
    . (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
    function global:Connect-MgGraph { [CmdletBinding()] param($ClientId, $TenantId, $CertificateThumbprint, $ContextScope, [switch]$NoWelcome, $Scopes) }
    function global:Disconnect-MgGraph { [CmdletBinding()] param() }
    function global:Connect-ExchangeOnline { [CmdletBinding()] param($AppId, $CertificateThumbprint, $Organization, $ShowBanner, $CommandName, $UserPrincipalName, [switch]$DisableWAM) }
    function global:Disconnect-ExchangeOnline { [CmdletBinding(SupportsShouldProcess)] param() Write-Host 'Disconnected successfully !' }
    function global:Connect-IPPSSession { [CmdletBinding()] param($AppId, $CertificateThumbprint, $Organization, $ShowBanner, $CommandName, $UserPrincipalName, [switch]$EnableSearchOnlySession) }
    function script:New-FollowConfig {
        param([hashtable]$Cloud = @{}, [hashtable]$Licensing = @{ Users = @{ Mode = 'Group' } }, [hashtable]$Scope = @{ ResourceMode = 'SynchronizedShared' })
        $base = @{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; AppId = 'app'; CertificateThumbprint = 'ABCDEF0123456789'; SignIn = 'Certificate'; UserPrincipalName = '' }
        foreach ($key in $Cloud.Keys) { $base[$key] = $Cloud[$key] }
        return @{ Cloud = $base; Licensing = $Licensing; Scope = $Scope; Retention = @{ TagAttribute = 'CustomAttribute1'; TagValue = 'Converted'; HoldPolicy = '' } }
    }
}

AfterAll {
    foreach ($name in @('Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext', 'Reset-PraFakeGraphTransport', 'Connect-MgGraph', 'Disconnect-MgGraph', 'Connect-ExchangeOnline', 'Disconnect-ExchangeOnline', 'Connect-IPPSSession')) {
        Remove-Item -Path "function:$name" -ErrorAction SilentlyContinue
    }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Expectations of a batch (Check -Batch), pure rules' {
    BeforeAll {
        $script:config = @{ Retention = @{ TagAttribute = 'CustomAttribute1'; TagValue = 'Converted' } }
        $script:userItem = [pscustomobject]@{ kind = 'User'; object_guid = 'g1'; identity = 'alice@contoso.com' }
        $script:sharedItem = [pscustomobject]@{ kind = 'Shared'; object_guid = 'g2'; identity = 'compta@contoso.com' }
        $script:record = [pscustomobject]@{ exchange_guid = 'aaaaaaaa-0000-0000-0000-000000000001' }
        function script:New-Cloud {
            param([string]$Type, [string]$Guid = '77777777-0000-0000-0000-000000000007', $Cloud = $true, [bool]$Synced = $false, [string]$Tag = '', [string[]]$Holds = @())
            [pscustomobject]@{ RecipientType = $Type; ExchangeGuid = $Guid; IsCloudManaged = $Cloud; User = [pscustomobject]@{ onPremisesSyncEnabled = $(if ($Synced) { $true } else { $null }) }; TagValue = $Tag; Holds = $Holds }
        }
    }
    It 'Provisioned, user: every point is checked and named in plain words' {
        $good = Test-PcmBatchExpectation -Expect Provisioned -Item $script:userItem -Record $script:record -State (New-Cloud 'UserMailbox') -Config $script:config -Facts @{ HasPlan = $true }
        $good.Met | Should -BeTrue
        $good.Stage | Should -Be 'UserMailbox'
        $good.Detail | Should -Match 'cloud mailbox present; source of authority in the cloud; Exchange plan enabled'
        $late = Test-PcmBatchExpectation -Expect Provisioned -Item $script:userItem -Record $script:record -State (New-Cloud 'MailUser' -Cloud $true) -Config $script:config -Facts @{ HasPlan = $false }
        $late.Met | Should -BeFalse
        $late.Detail | Should -Match 'no cloud mailbox yet \(Exchange Online shows MailUser\)'
        $late.Detail | Should -Match 'no Exchange plan enabled'
        (Test-PcmBatchExpectation -Expect Provisioned -Item $script:userItem -Record $script:record -State (New-Cloud 'UserMailbox' -Cloud $false) -Config $script:config -Facts @{ HasPlan = $true }).Detail | Should -Match 'the source of authority is not in the cloud'
        (Test-PcmBatchExpectation -Expect Provisioned -Item $script:userItem -Record $script:record -State $null -Config $script:config).Detail | Should -Match 'no Entra ID object found'
    }
    It 'Provisioned, shared mailbox: type, tag and permissions (journal and live read-back)' {
        $ok = Test-PcmBatchExpectation -Expect Provisioned -Item $script:sharedItem -Record $script:record -State (New-Cloud 'SharedMailbox' -Tag 'Converted') -Config $script:config -Facts @{ PermissionsFailed = @(); PermissionsMissing = @() }
        $ok.Met | Should -BeTrue
        $ok.Detail | Should -Match 'permissions of the snapshot present'
        $bad = Test-PcmBatchExpectation -Expect Provisioned -Item $script:sharedItem -Record $script:record -State (New-Cloud 'SharedMailbox' -Tag '') -Config $script:config -Facts @{ PermissionsFailed = @('SendAs bob: refused'); PermissionsMissing = @('FullAccess alice@contoso.com') }
        $bad.Met | Should -BeFalse
        $bad.Detail | Should -Match 'CustomAttribute1 is not Converted'
        $bad.Detail | Should -Match 'permissions that Convert could not grant: SendAs bob: refused'
        $bad.Detail | Should -Match 'permissions of the snapshot not found: FullAccess alice@contoso.com'
    }
    It 'Deprovisioned: MailUser with the on-premises GUID, synchronised again, the source of authority back in AD' {
        $back = Test-PcmBatchExpectation -Expect Deprovisioned -Item $script:userItem -Record $script:record -State (New-Cloud 'MailUser' -Guid 'AAAAAAAA-0000-0000-0000-000000000001' -Cloud $false -Synced $true) -Config $script:config
        $back.Met | Should -BeTrue
        $not = Test-PcmBatchExpectation -Expect Deprovisioned -Item $script:userItem -Record $script:record -State (New-Cloud 'UserMailbox' -Cloud $true) -Config $script:config
        $not.Met | Should -BeFalse
        $not.Detail | Should -Match 'not a MailUser with the on-premises ExchangeGuid'
        $not.Detail | Should -Match 'not synchronised from AD yet'
        $not.Detail | Should -Match 'still in the cloud'
    }
    It 'Retained: the case hold for a user, the inactive mailbox for a shared mailbox, nothing for an unprotected native resource' {
        (Test-PcmBatchExpectation -Expect Retained -Item $script:userItem -Record $script:record -State (New-Cloud 'MailUser') -Config $script:config -Facts @{ HoldStamped = $true }).Met | Should -BeTrue
        (Test-PcmBatchExpectation -Expect Retained -Item $script:userItem -Record $script:record -State (New-Cloud 'MailUser') -Config $script:config -Facts @{ HoldStamped = $false }).Detail | Should -Match 'no case hold of the tool'
        (Test-PcmBatchExpectation -Expect Retained -Item $script:sharedItem -Record $script:record -State (New-Cloud 'MailUser') -Config $script:config -Facts @{ InactiveFound = $true }).Met | Should -BeTrue
        (Test-PcmBatchExpectation -Expect Retained -Item $script:sharedItem -Record $script:record -State (New-Cloud 'MailUser') -Config $script:config -Facts @{ InactiveFound = $false }).Detail | Should -Match 'the inactive mailbox is not there'
        (Test-PcmBatchExpectation -Expect Retained -Item $script:sharedItem -Record $script:record -State (New-Cloud 'MailUser') -Config $script:config).Detail | Should -Match 'no inactive mailbox recorded'
        $native = Test-PcmBatchExpectation -Expect Retained -Item ([pscustomobject]@{ kind = 'Room'; object_guid = 'g3' }) -Record $null -State $null -Config $script:config -Facts @{ Native = $true; NativeMode = 'None'; NativeExists = $false }
        $native.Met | Should -BeTrue
        $native.Detail | Should -Match 'by design'
    }
    It 'native resource: Provisioned needs the owned mailbox, Deprovisioned needs the identity gone' {
        $item = [pscustomobject]@{ kind = 'Room'; object_guid = 'g3' }
        (Test-PcmBatchExpectation -Expect Provisioned -Item $item -Record $null -State $null -Config $script:config -Facts @{ Native = $true; NativeExists = $true; NativeType = 'RoomMailbox'; CloudType = 'RoomMailbox' }).Met | Should -BeTrue
        (Test-PcmBatchExpectation -Expect Provisioned -Item $item -Record $null -State $null -Config $script:config -Facts @{ Native = $true; NativeExists = $false; CloudType = 'RoomMailbox' }).Met | Should -BeFalse
        (Test-PcmBatchExpectation -Expect Deprovisioned -Item $item -Record $null -State $null -Config $script:config -Facts @{ Native = $true; NativeExists = $false }).Met | Should -BeTrue
        (Test-PcmBatchExpectation -Expect Deprovisioned -Item $item -Record $null -State $null -Config $script:config -Facts @{ Native = $true; NativeExists = $true }).Detail | Should -Match 'still there'
    }
    It 'says that a state that could not be read is not met' {
        $failure = [System.Management.Automation.ErrorRecord]::new([Exception]::new('Graph refused'), 'x', 'NotSpecified', $null)
        (Test-PcmBatchExpectation -Expect Provisioned -Item $script:userItem -Record $script:record -State $failure -Config $script:config).Detail | Should -Match 'the state could not be read: Graph refused'
    }
}

Describe 'Global report of a batch' {
    BeforeAll {
        function script:New-History {
            $convert = [pscustomobject]@{ id = 'c0nvert1'; action = 'Convert'; convert_batch = ''; status = 'Partial'; created_utc = '2026-11-01T08:00:00.0000000Z'; updated_utc = '2026-11-01T09:00:00.0000000Z'; snapshot_id = 4; account = 'CONTOSO\admin'; computer = 'ADM1'; tool_version = '2.0.0' }
            $recover = [pscustomobject]@{ id = 'r3cover1'; action = 'Recover'; convert_batch = 'c0nvert1'; status = 'Complete'; created_utc = '2026-11-05T08:00:00.0000000Z'; updated_utc = '2026-11-05T09:00:00.0000000Z'; snapshot_id = 4; account = 'CONTOSO\admin'; computer = 'ADM1'; tool_version = '2.0.0' }
            $item = { param($batch, $guid, $identity, $kind, $status, $step, $message) [pscustomobject]@{ batch_id = $batch; object_guid = $guid; identity = $identity; kind = $kind; status = $status; step = $step; message = $message } }
            $event = { param($batch, $guid, $time, $step, $outcome, $detail) [pscustomobject]@{ batch_id = $batch; object_guid = $guid; time_utc = $time; step = $step; outcome = $outcome; detail = $detail } }
            [pscustomobject]@{
                Batches = @($convert, $recover)
                Items = @((& $item 'c0nvert1' 'g1' 'alice@contoso.com' 'User' 'Done' 'MailboxReady' ''), (& $item 'c0nvert1' 'g2' 'bob@contoso.com' 'User' 'Failed' 'ConvertUser' 'not added to the licence group'),
                    (& $item 'c0nvert1' 'g3' '=cmd@contoso.com' 'User' 'Done' 'MailboxReady' ''), (& $item 'r3cover1' 'g1' 'alice@contoso.com' 'User' 'Done' 'Hold' ''))
                Events = @((& $event 'c0nvert1' 'g1' '2026-11-01T08:05:00.0000000Z' 'GuidCleared' 'Ok' 'ExchangeGuid cleared'), (& $event 'c0nvert1' 'g1' '2026-11-01T08:10:00.0000000Z' 'MailboxReady' 'Ok' 'UserMailbox <b>x</b>'),
                    (& $event 'c0nvert1' 'g2' '2026-11-01T08:06:00.0000000Z' 'ConvertUser' 'Fail' 'not added to the licence group'), (& $event 'r3cover1' 'g1' '2026-11-05T08:30:00.0000000Z' 'Hold' 'Ok' 'case hold stamped'),
                    (& $event 'c0nvert1' '' '2026-11-01T08:00:30.0000000Z' 'GroupSoaCloud' 'Ok' 'licence group managed in the cloud'))
                Runs = @([pscustomobject]@{ run_id = 'r1'; batch_id = 'c0nvert1'; action = 'Convert'; mode = 'Apply'; started_utc = '2026-11-01T08:00:00.0000000Z'; ended_utc = '2026-11-01T09:00:00.0000000Z'; status = 'Failed'; exit_code = 1; done_count = 2; pending_count = 0; error_count = 1; skipped_count = 0
                        source_text = 'group PRA-Wave1 (captured by Collect, snapshot 4)'; overrides_text = 'Polling.IntervalSeconds: 20 -> 5 | Cloud.SignIn: Certificate -> Interactive' })
            }
        }
    }
    It 'builds the verdict of each object from the stages and the events, journey in time order' {
        $model = Get-PcmGlobalModel -History (New-History) -Version '2.0.0' -Environment 'TEST'
        $model.Counts.Converted | Should -Be 1
        $model.Counts.RolledBack | Should -Be 1
        $model.Counts.Failed | Should -Be 1
        ($model.Objects | Where-Object Identity -eq 'alice@contoso.com').Verdict | Should -Be 'Rolled back'
        ($model.Objects | Where-Object Identity -eq 'bob@contoso.com').Verdict | Should -Be 'Failed'
        ($model.Objects | Where-Object Identity -eq 'bob@contoso.com').Note | Should -Be 'not added to the licence group'
        @(($model.Objects | Where-Object Identity -eq 'alice@contoso.com').Events | ForEach-Object { $_.Step }) | Should -Be @('GuidCleared', 'MailboxReady', 'Hold')
        @($model.BatchEvents).Count | Should -Be 1
    }
    It 'writes one HTML page with no script, every value encoded, the runs, the settings of each run, and a CSV with a formula guard' {
        $folder = Join-Path $TestDrive 'global'
        $context = @{ ReportFolder = $folder; Version = '2.0.0'; Config = @{ Environment = 'TEST' }; Warnings = 0; LogFile = '' }
        $result = Write-PcmGlobalReport -Context $context -History (New-History) -ConvertBatchId 'c0nvert1' -FollowUp @([pscustomobject]@{ Identity = 'alice@contoso.com'; Expect = 'Deprovisioned'; Met = $true; Detail = 'MailUser' })
        $result.Html | Should -Match 'PraCloudMailbox_TEST_batchc0nvert1_global\.html$'
        $html = [IO.File]::ReadAllText($result.Html)
        $html | Should -Not -Match '<script'
        $html | Should -Match 'Batch c0nvert1 from end to end'
        $html | Should -Match 'UserMailbox &lt;b&gt;x&lt;/b&gt;'
        $html | Should -Not -Match 'UserMailbox <b>'
        $html | Should -Match 'Polling.IntervalSeconds: 20 -&gt; 5'
        $html | Should -Match 'group PRA-Wave1 \(captured by Collect, snapshot 4\)'
        $html | Should -Match 'Follow-up check of this run'
        $rows = @(Import-Csv -LiteralPath $result.Csv -Delimiter ';')
        $rows.Count | Should -Be 3
        ($rows | Where-Object Verdict -eq 'Converted').Identity | Should -Be "'=cmd@contoso.com"
        ($rows | Where-Object Identity -eq 'alice@contoso.com').Journey | Should -Match 'Hold: case hold stamped'
        # The same two files are replaced by the next run.
        $again = Write-PcmGlobalReport -Context $context -History (New-History) -ConvertBatchId 'c0nvert1'
        $again.Html | Should -Be $result.Html
    }
    It 'writes the journey only for the objects that are not simply done when there are more than 500 objects' {
        $history = New-History
        $many = foreach ($number in 1..501) { [pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = ('x{0}' -f $number); identity = ('u{0:000}@contoso.com' -f $number); kind = 'User'; status = 'Done'; step = 'MailboxReady'; message = '' } }
        $history.Items = @($history.Items) + @($many)
        $history.Events = @($history.Events) + @([pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = 'x7'; time_utc = '2026-11-01T08:00:00.0000000Z'; step = 'MailboxReady'; outcome = 'Ok'; detail = 'simple object' })
        $context = @{ ReportFolder = (Join-Path $TestDrive 'global2'); Version = '2.0.0'; Config = @{ Environment = 'TEST' }; Warnings = 0; LogFile = '' }
        $result = Write-PcmGlobalReport -Context $context -History $history -ConvertBatchId 'c0nvert1'
        $html = [IO.File]::ReadAllText($result.Html)
        $html | Should -Match 'More than 500 objects'
        $html | Should -Not -Match 'Journey of u007@contoso.com'
        $html | Should -Match 'Journey of bob@contoso.com'
    }
    It 'keeps the runs of a batch in the journal, and reads an older journal without the table' {
        $path = Join-Path $TestDrive 'journal.db'
        $cn = Open-PcmStore -Path $path -Root $script:Root -Kind Journal
        try {
            $id = New-PcmBatch $cn Convert @{ SnapshotId = 1; Environment = 'T' }
            Add-PcmBatchRun $cn $id @{ RunId = 'run-a'; Action = 'Convert'; Mode = 'Apply'; Started = [datetime]::UtcNow.AddMinutes(-5); Account = 'a'; ToolVersion = '2.0.0'; Source = 'every object of the snapshot'; Overrides = ''; Status = 'Success'; ExitCode = 0; Done = 3; Pending = 0; Errors = 0; Skipped = 0; LogFile = 'x.log' }
            Add-PcmBatchRun $cn $id @{ RunId = 'run-b'; Action = 'Convert'; Mode = 'Apply'; Started = [datetime]::UtcNow; Account = 'a'; ToolVersion = '2.0.0'; Source = 'group PRA'; Overrides = 'A: 1 -> 2'; Status = 'Pending'; ExitCode = 2; Done = 1; Pending = 2; Errors = 0; Skipped = 0; LogFile = 'y.log' }
            $history = Get-PcmBatchHistory $cn $id.ToUpperInvariant()
            @($history.Runs | ForEach-Object { $_.run_id }) | Should -Be @('run-a', 'run-b')
            $null = Invoke-PcmSql $cn 'DROP TABLE batch_run'
            @((Get-PcmBatchHistory $cn $id).Runs).Count | Should -Be 0
        } finally { Close-PcmStore $cn }
    }
}

Describe 'Report page: data-driven template' {
    BeforeAll {
        function script:New-ReportContext {
            param([object[]]$Rows)
            $run = @{ Root = $script:Root; Version = '2.0.0'; RunId = 'r1'; StartTime = Get-Date; Action = 'Convert'; Mode = 'Apply'; Config = @{ Environment = 'TEST'; Cloud = @{ Organization = 'contoso.onmicrosoft.com' } }
                Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); NextSteps = @('pwsh -File "x.ps1" -Action Convert'); BatchId = 'abcd1234'; SnapshotLabel = '7'; Server = 'EX1'
                ObjectsText = 'Group PRA-Wave1 (captured by Collect, snapshot 7): 3 member(s) including nested groups, 1 skipped'; SkippedMembers = (New-Object 'Collections.Generic.List[string]')
                ConfigOverrides = @('Polling.IntervalSeconds: 20 -> 5'); ExitCode = 0; CloudConnectFailed = $false; ResultStatus = '' }
            $run.SkippedMembers.Add('PC01 (a computer, not a mailbox)')
            foreach ($row in $Rows) { [void]$run.Rows.Add($row) }
            return $run
        }
    }
    It 'stores the rows, the run errors and the warnings as compressed chunks and reads them back exactly, whatever the characters' {
        $rows = @(
            [ordered]@{ Identity = 'alice@contoso.com'; Kind = 'User'; PrimarySmtpAddress = 'alice@contoso.com'; ObjectGuid = 'g1'; Action = 'Convert'; ExchangeOnline = 'UserMailbox'; FinalStatus = 'Success'; Detail = 'plain' }
            [ordered]@{ Identity = '</script><script>alert(1)</script>'; Kind = 'Shared'; ObjectGuid = 'g2'; Action = 'Convert'; FinalStatus = 'Error'; Detail = ("quote `" back\slash`r`nline " + [char]0x2028 + ' <b>&amp;</b>') })
        $context = New-ReportContext -Rows $rows
        Add-PcmIssue $context 'Stopped at step "x": <boom>'
        $context['WarningList'] = [System.Collections.Generic.List[object]]::new()
        [void]$context['WarningList'].Add([pscustomobject]@{ Timestamp = ''; Phase = 'x'; Message = 'a warning' })
        $outcome = Get-PcmOutcome $context
        $html = Get-PcmReportHtml -Context $context -Outcome $outcome
        ([regex]::Matches($html, '<script')).Count | Should -BeGreaterThan 0
        $html | Should -Not -Match '</script><script>alert'
        $html | Should -Match 'Group PRA-Wave1 \(captured by Collect, snapshot 7\)'
        $html | Should -Match 'PC01 \(a computer, not a mailbox\)'
        $html | Should -Match 'Polling.IntervalSeconds: 20 -&gt; 5'
        $data = Read-PcmReportData -Html $html
        @($data.Rows).Count | Should -Be 3
        $data.Rows[0].Identity | Should -Be 'alice@contoso.com'
        $data.Rows[1].Identity | Should -Be '</script><script>alert(1)</script>'
        $data.Rows[1].Detail | Should -Be ("quote `" back\slash`r`nline " + [char]0x2028 + ' <b>&amp;</b>')
        $data.Rows[0].Permissions | Should -BeNullOrEmpty
        @($data.Issues).Count | Should -Be 1
        $data.Issues[0].Message | Should -Match '<boom>'
        @($data.Warnings).Count | Should -Be 1
        $data.Header.rows.count | Should -Be 3
        $data.Rows[2].Detail | Should -Match '^\[Run/Execution\] Stopped at step'
    }
    It 'streams a report with thousands of rows in chunks and keeps the memory use flat' {
        $rows = foreach ($number in 1..2500) { [ordered]@{ Identity = ('user{0}@contoso.com' -f $number); Kind = 'User'; ObjectGuid = [guid]::NewGuid().ToString(); Action = 'Check'; FinalStatus = 'Success'; Detail = ('detail {0}' -f ($number % 7)) } }
        $context = New-ReportContext -Rows @($rows)
        $outcome = Get-PcmOutcome $context
        $writer = New-Object IO.StringWriter
        Write-PcmReportHtml -Context $context -Outcome $outcome -Writer $writer -ChunkRows 1000
        $html = $writer.ToString()
        ([regex]::Matches($html, 'data-kind="rows"')).Count | Should -Be 3
        (Read-PcmReportData -Html $html).Rows.Count | Should -Be 2500
    }
    It 'the template page script is valid and reads the PCM fields (no field of the other tool)' {
        $template = [IO.File]::ReadAllText((Join-Path $script:Root 'templates\Report.template.html'))
        foreach ($marker in @('{{TITLE}}', '{{SUBTITLE}}', '{{VERSION}}', '{{META}}', '{{TILES}}', '{{NEXT}}', '{{DATA_NOTICE}}', '{{DATA}}')) { $template | Should -Match ([regex]::Escape($marker)) }
        $template | Should -Match 'PRA Cloud Mailbox'
        $template | Should -Not -Match 'SamAccountName|ADVerified|PRA Remote Mailbox'
        foreach ($field in @('Identity', 'PrimarySmtpAddress', 'ExchangeOnline', 'TeamsStorage', 'Licence', 'Holds', 'Permissions', 'FinalStatus')) { $template | Should -Match $field }
    }
}

Describe 'Settings of a run that differ from the configuration file (-BaseConfigPath)' {
    It 'lists what differs, key by key, and ignores the letter case and the list order of nothing but the text' {
        $base = @{ Environment = 'PROD'; Polling = @{ IntervalSeconds = 20; MailboxTimeoutMinutes = 30 }; Collect = @{ TargetGroups = @('A', 'B') }; Cloud = @{ SignIn = 'Certificate' }; _Path = 'x' }
        $run = @{ Environment = 'prod'; Polling = @{ IntervalSeconds = 5; MailboxTimeoutMinutes = 30 }; Collect = @{ TargetGroups = @('A', 'C') }; Cloud = @{ SignIn = 'INTERACTIVE' }; _Path = 'y' }
        $difference = @(Get-PcmConfigurationDifference -Base $base -Effective $run)
        @($difference | ForEach-Object { $_.Key }) | Should -Be @('Cloud.SignIn', 'Collect.TargetGroups', 'Polling.IntervalSeconds')
        ($difference | Where-Object Key -eq 'Polling.IntervalSeconds').Text | Should -Be 'Polling.IntervalSeconds: 20 -> 5'
        ($difference | Where-Object Key -eq 'Collect.TargetGroups').Text | Should -Be 'Collect.TargetGroups: A, B -> A, C'
        @(Get-PcmConfigurationDifference -Base $base -Effective $base).Count | Should -Be 0
    }
    It 'shows the differences in the summary card and keeps them in the result' {
        $cards = [System.Collections.Generic.List[object]]::new()
        Mock Write-PcmSummary -ModuleName PRA.Common { $cards.Add($Values) }
        Mock Write-Host -ModuleName PRA.Common { }
        $run = @{ Root = $script:Root; Version = 't'; RunId = 'r'; StartTime = Get-Date; Action = 'Check'; Mode = 'Preview'; Phase = ''; CurrentPhase = 'x'; CurrentOperation = 'x'; CurrentIdentity = ''; Warnings = 0
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
            LogFolder = $TestDrive; ReportFolder = $TestDrive; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''
            Config = @{}; BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null; ObjectsText = 'One object: a@contoso.com'; ConfigOverrides = @(1..6 | ForEach-Object { "Key$_`: a -> b" }) }
        $result = Complete-PcmRun -Context $run
        $values = $cards[0]
        $values['Objects'][1] | Should -Be 'One object: a@contoso.com'
        @($values['Settings'][1])[0] | Should -Be '6 setting(s) of this run differ from the configuration file:'
        @($values['Settings'][1]).Count | Should -Be 6
        @($values['Settings'][1])[-1] | Should -Match '^\.\.\. and 2 more'
        $result.ConfigOverrides.Count | Should -Be 6
    }
}

Describe 'Interactive sign-in of an administrator' {
    BeforeEach {
        Reset-PraFakeGraphTransport
        & (Get-Module PRA.Cloud) { $script:GraphRequestOwner = $null; $script:GraphTransportReady = $false }
        $global:PraFakeGraphRequest.MaxRetry = 3
        $script:captured = @{}
        Mock Import-Module -ModuleName PRA.Cloud { }
        Mock Test-PcmNonInteractiveProcess -ModuleName PRA.Cloud { $false }
        Mock Connect-MgGraph -ModuleName PRA.Cloud {
            $script:captured.Scopes = @($Scopes); $script:captured.Tenant = $TenantId; $script:captured.ContextScope = $ContextScope; $script:captured.AppId = (Get-Variable -Name ClientId -ValueOnly -ErrorAction SilentlyContinue)
            $global:PraFakeGraphAuth = [pscustomobject]@{ Account = 'admin@contoso.com'; Scopes = @($Scopes) }
        }
        Mock Disconnect-MgGraph -ModuleName PRA.Cloud { $global:PraFakeGraphAuth = $null }
        Mock Connect-ExchangeOnline -ModuleName PRA.Cloud {
            $value = { param($name) Get-Variable -Name $name -ValueOnly -ErrorAction SilentlyContinue }
            $script:captured.Exo = @{ UserPrincipalName = (& $value 'UserPrincipalName'); AppId = (& $value 'AppId'); CertificateThumbprint = (& $value 'CertificateThumbprint'); DisableWAM = [bool](& $value 'DisableWAM') }
            $script:captured.ExoCommands = @($CommandName)
        }
        Mock Connect-IPPSSession -ModuleName PRA.Cloud {
            $value = { param($name) Get-Variable -Name $name -ValueOnly -ErrorAction SilentlyContinue }
            $script:captured.Ipps = @{ UserPrincipalName = (& $value 'UserPrincipalName'); AppId = (& $value 'AppId'); EnableSearchOnlySession = [bool](& $value 'EnableSearchOnlySession'); CommandName = @($CommandName) }
        }
    }
    It 'signs in to Graph with the delegated permissions and to Exchange Online with the administrator, and never with a certificate or -DisableWAM' {
        $context = @{ Config = (New-FollowConfig -Cloud @{ SignIn = 'interactive'; UserPrincipalName = 'admin@contoso.com'; AppId = ''; CertificateThumbprint = '' }); Approval = { param($Target, $Operation) $true } }
        $text = Connect-PcmCloud -Context $context
        $text | Should -Match 'interactive sign-in as admin@contoso.com'
        $script:captured.ContextScope | Should -BeExactly 'Process'
        $script:captured.Tenant | Should -Be '33333333-3333-3333-3333-333333333333'
        $script:captured.AppId | Should -BeNullOrEmpty
        $script:captured.Scopes | Should -Contain 'User.ReadWrite.All'
        $script:captured.Scopes | Should -Contain 'GroupMember.ReadWrite.All'
        $script:captured.Exo.UserPrincipalName | Should -Be 'admin@contoso.com'
        $script:captured.Exo.AppId | Should -BeNullOrEmpty
        $script:captured.Exo.CertificateThumbprint | Should -BeNullOrEmpty
        $script:captured.Exo.DisableWAM | Should -BeFalse
        $script:captured.ExoCommands | Should -Contain 'Set-Mailbox'
        $context['ExoConnected'] | Should -BeTrue
        Disconnect-PcmCloud -Context $context
    }
    It 'signs in to Security and Compliance with the same administrator, and the search-only session' {
        $context = @{ Config = (New-FollowConfig -Cloud @{ SignIn = 'Interactive'; UserPrincipalName = 'admin@contoso.com' }); Approval = { param($Target, $Operation) $true } }
        Connect-PcmCompliance -Context $context
        $script:captured.Ipps.UserPrincipalName | Should -Be 'admin@contoso.com'
        $script:captured.Ipps.AppId | Should -BeNullOrEmpty
        $script:captured.Ipps.EnableSearchOnlySession | Should -BeTrue
        $script:captured.Ipps.CommandName | Should -Contain 'Set-CaseHoldPolicy'
    }
    It 'asks for the account that Graph signed in when no administrator is given' {
        $context = @{ Config = (New-FollowConfig -Cloud @{ SignIn = 'Interactive' }); Approval = { param($Target, $Operation) $true } }
        $null = Connect-PcmCloud -Context $context
        $script:captured.Exo.UserPrincipalName | Should -Be 'admin@contoso.com'
        Disconnect-PcmCloud -Context $context
    }
    It 'signs out of Exchange Online without the line of the module on the console' {
        $context = @{ Config = (New-FollowConfig -Cloud @{ SignIn = 'Interactive' }); Approval = { param($Target, $Operation) $true } }
        $null = Connect-PcmCloud -Context $context
        $context['ExoConnected'] | Should -BeTrue
        $printed = @(Disconnect-PcmCloud -Context $context 6>&1)
        $printed.Count | Should -Be 0
        $context['ExoConnected'] | Should -BeFalse
    }
    It 'refuses, in plain words, to sign in when PowerShell was started with -NonInteractive' {
        Mock Test-PcmNonInteractiveProcess -ModuleName PRA.Cloud { $true }
        $context = @{ Config = (New-FollowConfig -Cloud @{ SignIn = 'Interactive' }); Approval = { param($Target, $Operation) $true } }
        { Connect-PcmCloud -Context $context } | Should -Throw '*started with -NonInteractive*Open PowerShell 7 normally*'
        Should -Invoke Connect-MgGraph -ModuleName PRA.Cloud -Times 0 -Exactly
    }
    It 'detects -NonInteractive in any accepted abbreviation and letter case' -TestCases @(
        @{ Line = @('pwsh', '-NonInteractive', '-File', 'x'); Expected = $true }, @{ Line = @('pwsh', '-noninteractive'); Expected = $true }, @{ Line = @('pwsh', '-NONI', '-File'); Expected = $true }
        @{ Line = @('pwsh', '/NonInteractive'); Expected = $true }, @{ Line = @('pwsh', '-NoProfile', '-File', 'x', '-Action', 'Check'); Expected = $false }, @{ Line = @('pwsh', '-No'); Expected = $false }
        @{ Line = @('pwsh', '-NoLogo'); Expected = $false }) {
        param($Line, $Expected)
        Test-PcmNonInteractiveProcess -CommandLine $Line | Should -Be $Expected
    }
    It 'derives the delegated permissions from the licensing mode: the read-only group permission only when no write one is asked' -TestCases @(
        @{ Mode = 'Group'; Expected = @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
        @{ Mode = 'kiosk'; Expected = @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.Read.All') }
        @{ Mode = 'DIRECT'; Expected = @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.Read.All') }) {
        param($Mode, $Expected)
        @(Get-PcmDelegatedScope -Config (New-FollowConfig -Licensing @{ Users = @{ Mode = $Mode } })) | Should -Be $Expected
    }
    It 'reports the missing delegated permissions like the missing application permissions, in each wording' {
        $roles = @('User.ReadWrite.All', 'Organization.Read.All')
        $tenant = [pscustomobject]@{ Skus = @(); Organization = $null; LicenceGroup = $null; LicenceGroupCloudManaged = $null; OrgHolds = @() }
        $config = @{ Licensing = @{ Users = @{ Mode = 'Kiosk' }; Shared = @{ SkuPartNumber = '' } } }
        $app = Test-PcmTenantReadiness -Config $config -Roles $roles -Tenant $tenant -SkipUserLicensing
        $delegated = Test-PcmTenantReadiness -Config $config -Roles $roles -Tenant $tenant -SkipUserLicensing -SignIn Interactive
        ($app | Where-Object Code -eq 'APP_PERMISSION').Message | Should -Match '^Missing Graph application permission\(s\): User-OnPremisesSyncBehavior\.ReadWrite\.All'
        ($delegated | Where-Object Code -eq 'APP_PERMISSION').Message | Should -Match '^Missing Graph delegated permission\(s\): User-OnPremisesSyncBehavior\.ReadWrite\.All .*consent'
        ($delegated | Where-Object Code -eq 'APP_PERMISSION').Level | Should -Be 'Error'
    }
    It 'compares the Microsoft Entra roles of the administrator with what an action needs, letter case ignored' {
        (Test-PcmAdminRole -Action Convert -Roles @('global administrator')).Level | Should -Be 'Ok'
        $some = @(Test-PcmAdminRole -Action Convert -Roles @('USER ADMINISTRATOR', 'Exchange Administrator'))
        @($some | Where-Object Level -eq 'Warn').Count | Should -Be 1
        $some[0].Message | Should -Match 'Hybrid Identity Administrator, Global Administrator'
        @(Test-PcmAdminRole -Action Check -Roles @('Global Reader'))[0].Level | Should -Be 'Ok'
        @(Test-PcmAdminRole -Action Convert -Roles $null)[0].Level | Should -Be 'Info'
        (Get-PcmRoleNeed -Action Recover).Count | Should -Be 4
    }
    It 'reads the roles of the account through Graph, and nothing when they cannot be read' {
        Mock Invoke-PcmGraph -ModuleName PRA.Cloud { [pscustomobject]@{ value = @([pscustomobject]@{ displayName = 'User Administrator' }, [pscustomobject]@{ displayName = 'Exchange Administrator' }) } }
        @(Get-PcmAdminRole) | Should -Be @('User Administrator', 'Exchange Administrator')
        Mock Invoke-PcmGraph -ModuleName PRA.Cloud { throw 'Forbidden' }
        Get-PcmAdminRole | Should -BeNullOrEmpty
    }
    It 'knows which Graph permissions can read groups' {
        Test-PcmGroupReadRole -Roles @('User.ReadWrite.All') | Should -BeFalse
        Test-PcmGroupReadRole -Roles @('groupmember.read.all') | Should -BeTrue
        Test-PcmGroupReadRole -Roles @('GroupMember.ReadWrite.All') | Should -BeTrue
        Test-PcmGroupReadRole -Roles @() | Should -BeFalse
    }
    It 'reads Cloud.SignIn and Cloud.UserPrincipalName in any letter case and refuses Interactive without a tenant' {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $folder
        $path = Join-Path $folder 's.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; cloud = @{ signin = 'INTERACTIVE'; tenantid = '33333333-3333-3333-3333-333333333333'; userprincipalname = ' admin@contoso.com ' } }"
        $config = Import-PcmConfiguration -Path $path -Root $folder
        $config.Cloud.SignIn | Should -BeExactly 'Interactive'
        $config.Cloud.UserPrincipalName | Should -BeExactly 'admin@contoso.com'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Cloud = @{ SignIn = 'interactive' } }"
        { Import-PcmConfiguration -Path $path -Root $folder } | Should -Throw '*Cloud.SignIn = Interactive requires Cloud.TenantId*'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Cloud = @{ SignIn = 'Password' } }"
        { Import-PcmConfiguration -Path $path -Root $folder } | Should -Throw '*Cloud.SignIn must be Certificate*or Interactive*'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Cloud = @{ SignIn = 'Interactive'; TenantId = '33333333-3333-3333-3333-333333333333'; AppId = 'only-the-app' } }"
        { Import-PcmConfiguration -Path $path -Root $folder } | Should -Not -Throw
    }
}

Describe 'The console keeps the format of the tool: the warnings of the Microsoft modules go to the log' {
    BeforeAll {
        function global:Get-OrganizationConfig { [CmdletBinding()] param() Write-Warning 'Le systeme n''a pas pu obtenir les parametres de la fonctionnalite de resilience'; [pscustomobject]@{ Name = 'contoso' } }
        function global:Get-PcmQuietProbe { param() Write-Warning 'a warning of the tool itself' }
    }
    AfterAll {
        foreach ($name in 'Get-OrganizationConfig', 'Get-PcmQuietProbe') { Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue }
    }
    It 'keeps the warning of an Exchange Online command off the console, keeps its result, and logs the warning once at the sign-out' {
        $null = Read-PcmModuleWarning
        $records = @(& (Get-Module PRA.Cloud) { Get-OrganizationConfig; Get-OrganizationConfig } 3>&1)
        @($records | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }).Count | Should -Be 0
        @($records | Where-Object { $_.Name -eq 'contoso' }).Count | Should -Be 2
        $script:logged = [System.Collections.Generic.List[string]]::new()
        Mock Write-PcmLog -ModuleName PRA.Cloud { $script:logged.Add([string]$Message) }
        Disconnect-PcmCloud -Context @{ ExoConnected = $false }
        @($script:logged | Where-Object { $_ -like 'Microsoft 365 module warning: Le systeme*resilience' }).Count | Should -Be 1
        @(Read-PcmModuleWarning).Count | Should -Be 0
    }
    It 'adds the warnings that the entry script collected, and never silences a command of the tool itself' {
        Add-PcmModuleWarning -Warning @('from the entry script', $null)
        @(Read-PcmModuleWarning) | Should -Be @('from the entry script')
        (Get-PcmModuleQuietDefault).Keys | Where-Object { $_ -like '*Pcm*' } | Should -BeNullOrEmpty
        (Get-PcmModuleQuietDefault -VariableName 'X')['Get-Mailbox:WarningVariable'] | Should -Be '+X'
        @(& (Get-Module PRA.Cloud) { Get-PcmQuietProbe } 3>&1 | Where-Object { $_ -is [System.Management.Automation.WarningRecord] }).Count | Should -Be 1
    }
}
