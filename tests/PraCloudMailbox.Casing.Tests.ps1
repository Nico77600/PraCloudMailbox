<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the letter case (nothing an operator types is case sensitive) and of the command line of the engine
    (usage, confirmation, next steps, strict targets, module checks). Pester 5+, offline.
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
    Import-Module (Join-Path $script:Root 'module\PRA.Collect.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Cloud.psm1') -Force
    $script:Gui = Import-Module (Join-Path $script:Root 'module\PRA.Gui.psm1') -Force -PassThru
    . (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
    $script:Entry = Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Entry, [ref]$null, [ref]$null)
    foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                $_.Name -in @('ConvertTo-PcmIdentityKey', 'Read-PcmIdentityFile', 'Get-PcmIdentityList', 'Select-PcmIdentity', 'ConvertTo-PcmLiteral', 'Get-PcmCommandLine',
                    'Get-PcmFallbackNext', 'New-PcmApprovalCallback', 'Confirm-PcmApply', 'Test-PcmRunOption', 'Get-PcmTargetKind') })) {
        . ([scriptblock]::Create($function.Extent.Text))
    }
    $script:EntrySource = [IO.File]::ReadAllText($script:Entry)
    $script:HostExe = (Get-Process -Id $PID).Path

    function script:New-CaseRecord {
        param([string]$Upn, [string]$Smtp, [string]$Sam, [string]$Guid, [string]$Dn)
        [pscustomobject]@{ user_principal_name = $Upn; primary_smtp_address = $Smtp; sam_account_name = $Sam; object_guid = $Guid; distinguished_name = $Dn; kind = 'User' }
    }
    function script:New-CaseConfigFile {
        param([string]$Folder, [string]$Extra = '')
        $null = New-Item -ItemType Directory -Path $Folder -Force
        $path = Join-Path $Folder 'case.config.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value @"
@{
    Environment = 'TEST'
    Scope = @{ Mode = 'Auto' }
    Store = @{ Path = '$Folder\data\snapshot.db'; JournalPath = '$Folder\data\journal.db'; BackupFolder = '$Folder\data\backup' }
    Cloud = @{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; AppId = 'app'; CertificateThumbprint = 'ABCDEF0123456789' }
    EntraConnect = @{ Mode = 'Manual'; Server = 'aadc' }
    Logging = @{ Folder = '$Folder\logs' }
    Report = @{ Folder = '$Folder\reports' }
$Extra
}

AfterAll {
    Restore-PcmGlobalState $script:GlobalBefore
}
"@
        return $path
    }
    function script:Start-CaseProcess {
        <# Runs the entry script in a child process; stdin stays open so that a prompt would hang (the test then fails on the timeout). #>
        param([string]$Exe, [string[]]$Arguments, [int]$TimeoutSeconds = 120)
        $start = [Diagnostics.ProcessStartInfo]::new($Exe)
        $start.UseShellExecute = $false; $start.RedirectStandardInput = $true; $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true; $start.CreateNoWindow = $true
        $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script:Entry) + $Arguments
        $start.Arguments = ($arguments | ForEach-Object { if ($_ -match '[\s"]') { '"' + ($_ -replace '"', '\"') + '"' } else { $_ } }) -join ' '
        $process = [Diagnostics.Process]::Start($start)
        $outTask = $process.StandardOutput.ReadToEndAsync(); $errTask = $process.StandardError.ReadToEndAsync()
        $finished = $process.WaitForExit($TimeoutSeconds * 1000)
        if (-not $finished) { try { $process.Kill() } catch { $null = $_ } }
        [pscustomobject]@{ Finished = $finished; ExitCode = $(if ($finished) { $process.ExitCode } else { -1 }); Output = ($outTask.Result + "`n" + $errTask.Result) }
    }
}

Describe 'Case: configuration values' {
    BeforeEach { $script:folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $script:folder }

    It 'accepts every closed-list value in any letter case and keeps the official spelling' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value @"
@{
    environment = 'X'
    exchange = @{ connectionmode = 'remote'; server = 'ex1' }
    scope = @{ mode = 'group'; groupdn = 'CN=g,DC=contoso,DC=com'; resourcemode = 'cloudresource' }
    cloud = @{ tenantid = '33333333-AAAA-BBBB-CCCC-333333333333'; defaultusagelocation = 'fr' }
    licensing = @{ users = @{ mode = 'KIOSK'; groupid = '11111111-AAAA-BBBB-CCCC-111111111111' } }
    retention = @{ tagattribute = 'customattribute15' }
    entraconnect = @{ mode = 'MANUAL'; server = 'aadc' }
    resources = @{ licensinggroupid = '22222222-AAAA-BBBB-CCCC-222222222222' }
}
"@
        $config = Import-PcmConfiguration -Path $path -Root $script:folder
        $config.Exchange.Server | Should -BeExactly 'ex1'
        $config.Exchange.ContainsKey('ConnectionMode') | Should -BeFalse
        $config.Scope.Mode | Should -BeExactly 'Group'
        $config.Scope.ResourceMode | Should -BeExactly 'CloudResource'
        $config.Licensing.Users.Mode | Should -BeExactly 'Kiosk'
        $config.Retention.TagAttribute | Should -BeExactly 'CustomAttribute15'
        $config.EntraConnect.Mode | Should -BeExactly 'Manual'
        $config.Cloud.DefaultUsageLocation | Should -BeExactly 'FR'
        $config.Cloud.TenantId | Should -BeExactly '33333333-aaaa-bbbb-cccc-333333333333'
        $config.Licensing.Users.GroupId | Should -BeExactly '11111111-aaaa-bbbb-cccc-111111111111'
        $config.Resources.LicensingGroupId | Should -BeExactly '22222222-aaaa-bbbb-cccc-222222222222'
    }
    It 'still refuses a value that is not in the list, and one with an invisible character' {
        $path = Join-Path $script:folder 'bad.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Groupe' } }"
        { Import-PcmConfiguration -Path $path -Root $script:folder } | Should -Throw '*Scope.Mode must be*'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value ("@{ Scope = @{ Mode = 'Auto" + [char]0x200B + "' } }")
        { Import-PcmConfiguration -Path $path -Root $script:folder } | Should -Throw '*Scope.Mode must be*'
    }
    It 'has the Exchange Online module minimum in the shipped configuration and checks it' {
        $config = Import-PcmConfiguration -Path (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1') -Root $script:Root
        $config.Exo.MinModuleVersion | Should -Be '3.10.1'
        $path = Join-Path $script:folder 'exo.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Exo = @{ MinModuleVersion = 'three' } }"
        { Import-PcmConfiguration -Path $path -Root $script:folder } | Should -Throw '*Exo.MinModuleVersion*'
    }
}

Describe 'Case: helpers' {
    It 'keeps the first spelling and the order of distinct texts, letter case ignored' {
        @('Alice@Contoso.com', 'bob@contoso.com', 'ALICE@contoso.com', 'alice@contoso.com', 'Bob@Contoso.com', 'carol@contoso.com' | Get-PcmDistinctText) -join ',' |
            Should -BeExactly 'Alice@Contoso.com,bob@contoso.com,carol@contoso.com'
        @() | Get-PcmDistinctText | Should -BeNullOrEmpty
    }
    It 'compares texts by character code, letter case ignored, and invisible characters count' {
        Test-PcmSameText 'Apply' 'APPLY' | Should -BeTrue
        Test-PcmSameText 'Apply' ('Apply' + [char]0x200B) | Should -BeFalse
        Test-PcmSameText 'Apply' ('Apply' + [char]1) | Should -BeFalse
        Test-PcmSameText $null $null | Should -BeTrue
        Test-PcmSameText $null '' | Should -BeFalse
    }
    It 'finds the official spelling of a value of a closed list' {
        Resolve-PcmOfficialValue 'usersonly' @('All', 'UsersOnly', 'SharedOnly') | Should -BeExactly 'UsersOnly'
        Resolve-PcmOfficialValue 'users' @('All', 'UsersOnly') | Should -BeNullOrEmpty
    }
    It 'reads the tool version in one place' {
        Get-PcmToolVersion | Should -Be '2.0.0'
        $script:EntrySource | Should -Match '(?m)^\$toolVersion = Get-PcmToolVersion'
        $script:EntrySource | Should -Not -Match "(?m)^\`$toolVersion = '"
    }
}

Describe 'Case: identities' {
    BeforeAll {
        $script:alice = New-CaseRecord -Upn 'Alice@Contoso.com' -Smtp 'alice.smith@contoso.com' -Sam 'ASmith' -Guid 'a0000000-0000-0000-0000-00000000000a' -Dn 'CN=Alice Smith,OU=Staff,DC=contoso,DC=com'
        $script:bob = New-CaseRecord -Upn 'bob@contoso.com' -Smtp 'bob@contoso.com' -Sam 'bob' -Guid 'b0000000-0000-0000-0000-00000000000b' -Dn 'CN=Bob,OU=Staff,DC=contoso,DC=com'
        $script:keys = { param($r) @([string]$r.user_principal_name, [string]$r.primary_smtp_address, [string]$r.sam_account_name, [string]$r.object_guid, [string]$r.distinguished_name) }
        $script:context = @{ Identities = @(); NoChangeMade = $false; Issues = (New-Object 'Collections.Generic.List[object]') }
    }
    It 'finds an object by <Kind> in any letter case' -TestCases @(
        @{ Kind = 'UPN'; Typed = 'ALICE@CONTOSO.COM' }, @{ Kind = 'UPN'; Typed = 'alice@contoso.com' }
        @{ Kind = 'SMTP address'; Typed = 'ALICE.SMITH@Contoso.Com' }, @{ Kind = 'sAMAccountName'; Typed = 'asmith' }, @{ Kind = 'sAMAccountName'; Typed = 'ASMITH' }
        @{ Kind = 'GUID in capitals'; Typed = 'A0000000-0000-0000-0000-00000000000A' }, @{ Kind = 'GUID in lower case'; Typed = 'a0000000-0000-0000-0000-00000000000a' }
        @{ Kind = 'GUID with braces'; Typed = '{A0000000-0000-0000-0000-00000000000A}' }, @{ Kind = 'GUID without hyphens'; Typed = 'A000000000000000000000000000000A' }
        @{ Kind = 'DN'; Typed = 'cn=alice smith,ou=staff,dc=CONTOSO,dc=com' }) {
        param($Kind, $Typed)
        $Identity = $Typed
        $context.Identities = @()
        $found = @(Select-PcmIdentity -Items @($script:alice, $script:bob) -Where 'the snapshot' -Keys $script:keys)
        $found.Count | Should -Be 1
        $found[0].sam_account_name | Should -BeExactly 'ASmith'
    }
    It 'finds the objects of a list in any letter case' {
        $Identity = ''
        $context.Identities = @('BOB@contoso.COM', 'asmith')
        @(Select-PcmIdentity -Items @($script:alice, $script:bob) -Where 'the snapshot' -Keys $script:keys).Count | Should -Be 2
        $context.Identities = @()
    }
    It 'refuses the run when an identity asked for is not there, and lists up to 10 of them' {
        $Identity = 'nobody@contoso.com'
        { Select-PcmIdentity -Items @($script:alice) -Where 'snapshot 7' -Keys $script:keys } | Should -Throw '*nobody@contoso.com is not in snapshot 7. Nothing was changed*'
        $Identity = ''
        $context.Identities = @(1..12 | ForEach-Object { "ghost$_@contoso.com" }) + @('alice@contoso.com')
        $message = try { Select-PcmIdentity -Items @($script:alice) -Where 'snapshot 7' -Keys $script:keys; '' } catch { $_.Exception.Message }
        $message | Should -Match '12 of the 13 identities'
        $message | Should -Match 'ghost1@contoso.com'
        $message | Should -Match 'ghost10@contoso.com'
        $message | Should -Not -Match 'ghost11@contoso.com'
        $message | Should -Match '\.\.\.'
        $context.NoChangeMade | Should -BeTrue
        $context.Identities = @(); $context.NoChangeMade = $false
    }
}

Describe 'Case: -CsvPath files' {
    BeforeEach { $script:file = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.txt') }

    It 'reads a list, comments and blank lines ignored, spelling kept' {
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @('# wave', 'Alice@Contoso.com', '', '  bob@contoso.com  ')
        Read-PcmIdentityFile -Path $script:file | Should -Be @('Alice@Contoso.com', 'bob@contoso.com')
    }
    It 'reads a CSV whose Identity column is written <Header> (delimiter <Delimiter>)' -TestCases @(
        @{ Header = 'Identity'; Delimiter = ',' }, @{ Header = 'identity'; Delimiter = ',' }, @{ Header = 'IDENTITY'; Delimiter = ';' }, @{ Header = '"iDentity"'; Delimiter = ';' }) {
        param($Header, $Delimiter)
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @(('{0}{1}Note' -f $Header, $Delimiter), ('compta@contoso.com{0}shared' -f $Delimiter), ('Carol@Contoso.com{0}' -f $Delimiter))
        Read-PcmIdentityFile -Path $script:file | Should -Be @('compta@contoso.com', 'Carol@Contoso.com')
    }
    It 'refuses a list that names an object twice, letter case ignored, with the line numbers' {
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @('# wave', 'alice@contoso.com', 'bob@contoso.com', '  ALICE@contoso.com')
        { Read-PcmIdentityFile -Path $script:file } | Should -Throw '*line 4 lists ALICE@contoso.com again (already on line 2)*'
    }
    It 'sees a GUID written in two notations as the same object' {
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @('A0000000-0000-0000-0000-00000000000A', '{a0000000-0000-0000-0000-00000000000a}')
        { Read-PcmIdentityFile -Path $script:file } | Should -Throw '*line 2 lists*again (already on line 1)*'
    }
    It 'refuses a CSV row without identity, with its line number' {
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @('Identity;Note', 'alice@contoso.com;a', ';empty', 'bob@contoso.com;b')
        { Read-PcmIdentityFile -Path $script:file } | Should -Throw '*line 3 has no valid identity*'
    }
    It 'refuses a line with a control character' {
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @('alice@contoso.com', ('bob@contoso.com' + [char]7 + 'x'))
        { Read-PcmIdentityFile -Path $script:file } | Should -Throw '*line 2 has no valid identity*'
    }
    It 'says nothing is done and shows at most 10 problems' {
        Set-Content -LiteralPath $script:file -Encoding UTF8 -Value @(1..14 | ForEach-Object { 'same@contoso.com' })
        $message = try { Read-PcmIdentityFile -Path $script:file; '' } catch { $_.Exception.Message }
        $message | Should -Match 'nothing was done'
        $message | Should -Match 'and 3 more'
    }
}

Describe 'Case: batch IDs and journal values' {
    BeforeAll {
        $script:journalPath = Join-Path $TestDrive 'journal.db'
        $script:cn = Open-PcmStore -Path $script:journalPath -Root $script:Root -Kind Journal
        $script:batchId = New-PcmBatch $script:cn Convert @{ SnapshotId = 1; Environment = 'TEST'; ToolVersion = 'test' }
        $script:guid = 'a0000000-0000-0000-0000-00000000000a'
        Set-PcmBatchItem $script:cn $script:batchId $script:guid @{ kind = 'User'; identity = 'alice@contoso.com'; status = 'Planned' }
    }
    AfterAll { Close-PcmStore $script:cn }

    It 'finds a batch by its ID in capitals' {
        (Get-PcmBatch $script:cn -Id $script:batchId.ToUpperInvariant()).id | Should -BeExactly $script:batchId
        (Get-PcmBatch $script:cn -Id $script:batchId.Substring(0, 4).ToUpperInvariant()).id | Should -BeExactly $script:batchId
    }
    It 'reads and updates the rows of a batch whatever the letter case of its ID and of the object GUID' {
        @(Get-PcmBatchItem $script:cn $script:batchId.ToUpperInvariant()).Count | Should -Be 1
        Set-PcmBatchItem $script:cn $script:batchId.ToUpperInvariant() $script:guid.ToUpperInvariant() @{ step = 'GuidCleared'; status = 'Running' }
        $items = @(Get-PcmBatchItem $script:cn $script:batchId)
        $items.Count | Should -Be 1
        $items[0].step | Should -BeExactly 'GuidCleared'
        Set-PcmBatch $script:cn $script:batchId.ToUpperInvariant() Partial @{}
        (Get-PcmBatch $script:cn -Id $script:batchId).status | Should -BeExactly 'Partial'
    }
    It 'updates a tenant change once, whatever the letter case of its target' {
        Set-PcmTenantChange $script:cn $script:batchId 'GroupSoa' '11111111-AAAA-BBBB-CCCC-111111111111' -OriginalJson '{}'
        Set-PcmTenantChange $script:cn $script:batchId.ToUpperInvariant() 'GroupSoa' '11111111-aaaa-bbbb-cccc-111111111111' -Restored
        $changes = @(Get-PcmTenantChange $script:cn $script:batchId)
        $changes.Count | Should -Be 1
        [bool]$changes[0].restored | Should -BeTrue
    }
}

Describe 'Case: window confirmation' {
    It 'accepts the name of the action in any letter case, spaces around it ignored (<Typed>)' -TestCases @(
        @{ Typed = 'convert' }, @{ Typed = ' Convert ' }, @{ Typed = 'CONVERT' }, @{ Typed = "`tConVert`r`n" }) {
        param($Typed)
        & $script:Gui { param($t) Test-PcmGuiConfirmation -Typed $t -Word 'CONVERT' } $Typed | Should -BeTrue
    }
    It 'refuses a typed word that differs, also by an invisible character (<Name>)' -TestCases @(
        @{ Name = 'other word'; Typed = 'recover' }, @{ Name = 'empty'; Typed = '' }, @{ Name = 'zero-width space'; Typed = ('conv' + [char]0x200B + 'ert') }
        @{ Name = 'zero-width at the end'; Typed = ('Convert' + [char]0x200B) }, @{ Name = 'control character'; Typed = ('Convert' + [char]1) }, @{ Name = 'soft hyphen'; Typed = ('Conv' + [char]0xAD + 'ert') }) {
        param($Name, $Typed)
        & $script:Gui { param($t) Test-PcmGuiConfirmation -Typed $t -Word 'CONVERT' } $Typed | Should -BeFalse
    }
    It 'asks for the same word everywhere in the inline confirmation, and compares what is typed with the helper only' {
        $words = & $script:Gui { $script:GuiActionWord }
        foreach ($action in 'Collect', 'Convert', 'Recover') { $words[$action] | Should -BeExactly $action.ToUpperInvariant() }
        $words.ContainsKey('Check') | Should -BeFalse
        $view = [IO.File]::ReadAllText((Join-Path $script:Root 'module\PRA.Gui.View.ps1'))
        $run = [IO.File]::ReadAllText((Join-Path $script:Root 'module\PRA.Gui.Run.ps1'))
        # The word comes from one table, for the prompt next to the box and for the sentence of the summary.
        $view | Should -Match '\$word = \$script:GuiActionWord\[\$(g\.Action|action)\]'
        $view | Should -Match 'Type \{0\} to confirm \(any letter case\)\.[^'']*'' -f \$word'
        $view | Should -Match 'Type \$word below \(capitals or not\)'
        $run | Should -Match '\$word = \$script:GuiActionWord\[\$request\.Action\]'
        # Run is unlocked, and the run is accepted, by the helper that ignores the letter case and the spaces around the word.
        $view | Should -Match 'Test-PcmGuiConfirmation -Typed \(\[string\]\$g\.Controls\.ConfirmBox\.Text\) -Word \$word'
        $run | Should -Match 'Test-PcmGuiConfirmation -Typed \(\[string\]\$c\.ConfirmBox\.Text\) -Word \$word'
        ($view + $run) | Should -Not -Match 'ConfirmBox\.Text\s+-c(eq|ne)'
        ($view + $run) | Should -Not -Match 'ConfirmBox\.Text\)?\s*\.Equals\('
    }
    It 'builds the sets of GUIDs of the window with an explicit case-insensitive comparer' {
        $back = & $script:Gui {
            $summary = Get-PcmGuiBatchSummary -State ([pscustomobject]@{
                    Items = @([pscustomobject]@{ batch_id = 'c1'; object_guid = 'AAA'; kind = 'User'; status = 'Done'; step = 'Mailbox'; identity = 'a'; message = '' }
                        [pscustomobject]@{ batch_id = 'r1'; object_guid = 'aaa'; kind = 'User'; status = 'Done'; step = 'RecoverUser'; identity = 'a'; message = '' })
                    Batches = @([pscustomobject]@{ id = 'c1'; action = 'Convert'; convert_batch = ''; status = 'Partial'; created_utc = '2026-01-01T00:00:00Z'; account = 'x' }
                        [pscustomobject]@{ id = 'r1'; action = 'Recover'; convert_batch = 'C1'; status = 'Complete'; created_utc = '2026-01-02T00:00:00Z'; account = 'x' }) })
            @($summary)[0].Back.Contains('AaA')
        }
        $back | Should -BeTrue
    }
}

Describe 'Case: Collect removes duplicates that differ by letter case' {
    BeforeAll {
        foreach ($name in 'Get-Group', 'Get-CalendarProcessing', 'Get-MailboxCalendarConfiguration', 'Get-User', 'Get-DistributionGroup') {
            Set-Item -Path "function:global:$name" -Value { param($Identity, $DomainController, $Filter, $ResultSize) }
        }
    }
    AfterAll { foreach ($name in 'Get-Group', 'Get-CalendarProcessing', 'Get-MailboxCalendarConfiguration', 'Get-User', 'Get-DistributionGroup') { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue } }

    It 'keeps one SMTP address of a booking list, the first spelling, in order' {
        $mailbox = [pscustomobject]@{ PrimarySmtpAddress = 'room@contoso.com'; RecipientTypeDetails = 'RoomMailbox'; DistinguishedName = 'CN=room,OU=PRA,DC=contoso,DC=com'; ResourceCapacity = 4; ResourceCustom = @()
            RequireSenderAuthenticationEnabled = $true; ModerationEnabled = $false; SendModerationNotifications = 'Never'
            AcceptMessagesOnlyFromSendersOrMembers = @(); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @() }
        $booking = [pscustomobject]@{ AutomateProcessing = 'AutoAccept'; AllBookInPolicy = $false; AllRequestInPolicy = $true; AllRequestOutOfPolicy = $false
            ResourceDelegates = @('CN=d1', 'CN=d2', 'CN=d3'); BookInPolicy = @(); RequestInPolicy = @(); RequestOutOfPolicy = @(); ScheduleOnlyDuringWorkHours = $false }
        $hours = [pscustomobject]@{ WorkingHoursStartTime = '08:00:00'; WorkingHoursEndTime = '18:00:00'; WorkDays = 'Weekdays'; WorkingHoursTimeZone = 'Romance Standard Time' }
        $script:booking = $booking; $script:hours = $hours
        Mock Get-CalendarProcessing -ModuleName PRA.Collect { $script:booking }
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Collect { $script:hours }
        Mock Get-User -ModuleName PRA.Collect { [pscustomobject]@{ Office = ''; StreetAddress = ''; City = ''; StateOrProvince = ''; PostalCode = ''; CountryOrRegion = ''; Phone = ''; Notes = '' } }
        Mock Get-DistributionGroup -ModuleName PRA.Collect { @() }
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect {
            [pscustomobject]@{ Resolved = $true; Smtp = @{ 'CN=d1' = 'Delegate@Contoso.com'; 'CN=d2' = 'delegate@contoso.com'; 'CN=d3' = 'other@contoso.com' }[$Trustee] }
        }
        $context = @{ Config = @{ Exchange = @{ DomainController = '' } }; Issues = (New-Object 'Collections.Generic.List[object]'); Warnings = 0; LogFile = ''; Excluded = (New-Object 'Collections.Generic.List[string]') }
        $settings = Get-PcmResourceSettings -Context $context -Mailbox $mailbox
        $settings.Calendar.ResourceDelegates | Should -Be @('Delegate@Contoso.com', 'other@contoso.com')
    }
    It 'keeps one DN of a group whose members differ by letter case' {
        Mock Get-Group -ModuleName PRA.Collect { if ($Identity -eq 'g1') { [pscustomobject]@{ Members = @('CN=A,DC=x', 'cn=a,dc=X', 'CN=B,DC=x') } } else { throw 'not a group' } }
        $context = @{ Config = @{ Exchange = @{ DomainController = '' } } }
        @(Get-PcmGroupMemberDn -Context $context -Identity 'g1') | Should -Be @('CN=A,DC=x', 'CN=B,DC=x')
    }
    It 'keeps one address of a delivery restriction, the first spelling' {
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect { [pscustomobject]@{ Resolved = $true; Smtp = $(if ($Trustee -eq 'a') { 'Boss@Contoso.com' } else { 'boss@contoso.com' }) } }
        $recipient = [pscustomobject]@{ PrimarySmtpAddress = 'x@contoso.com'; RequireSenderAuthenticationEnabled = $true; ModerationEnabled = $false; SendModerationNotifications = 'Never'
            AcceptMessagesOnlyFromSendersOrMembers = @('a', 'b'); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @() }
        $delivery = & (Get-Module PRA.Collect) { param($r) Get-PcmDeliverySettings -Context @{} -Recipient $r } $recipient
        @($delivery.AcceptMessagesOnlyFromSendersOrMembers) | Should -Be @('Boss@Contoso.com')
    }
}

Describe 'Case: the entry script run as a command' {
    It 'stops at once without -Action, with the usage and exit code 1, and never waits for an answer' {
        $run = Start-CaseProcess -Exe $script:HostExe -Arguments @() -TimeoutSeconds 60
        $run.Finished | Should -BeTrue -Because 'the script must not prompt'
        $run.ExitCode | Should -Be 1
        $run.Output | Should -Match '-Action is missing: nothing was done'
        $run.Output | Should -Match 'Collect, Check, Convert, Recover or Schedule'
    }
    It 'shows the official spelling of the values typed in lower case, and still ends with the next step on exit 1' -Skip:($PSVersionTable.PSEdition -ne 'Core') {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $config = New-CaseConfigFile -Folder $folder
        $run = Start-CaseProcess -Exe $script:HostExe -Arguments @('-action', 'check', '-scope', 'sharedonly', '-identity', 'Alice@Contoso.com', '-configpath', $config)
        $run.Finished | Should -BeTrue
        $run.ExitCode | Should -Be 1
        $run.Output | Should -Match 'Check'
        $run.Output | Should -Match 'SharedOnly'
        $run.Output | Should -Not -MatchExactly 'sharedonly'
        $log = @(Get-ChildItem -LiteralPath (Join-Path $folder 'logs') -Filter '*.log' | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
        $log | Should -Match 'Summary - Next:'
        $log | Should -Match ('Summary - Next: \.\\Invoke-PraCloudMailbox\.ps1 -Action Check -Identity ''Alice@Contoso.com'' -Scope SharedOnly -ConfigPath ''{0}''' -f [regex]::Escape($config))
    }
    It 'shows Collect and Preview in their official spelling when typed in lower case ' {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $config = New-CaseConfigFile -Folder $folder
        $run = Start-CaseProcess -Exe $script:HostExe -Arguments @('-action', 'collect', '-mode', 'preview', '-configpath', $config)
        $run.Finished | Should -BeTrue
        $run.ExitCode | Should -Be 1
        $run.Output | Should -MatchExactly 'Collect'
        $run.Output | Should -MatchExactly 'Preview \(nothing is changed\)'
    }
    It 'refuses a ValidateSet value with an invisible character' {
        $run = Start-CaseProcess -Exe $script:HostExe -Arguments @('-Action', 'Check', '-Mode', ('Preview' + [char]0x200B))
        $run.ExitCode | Should -Not -Be 0
    }
}

Describe 'Engine: command line of the next steps' {
    BeforeAll {
        $script:ConfigPathBound = $false; $script:ConfirmBound = $false
        $context = @{ Root = 'C:\Tools' }
    }
    It 'writes the full path of the script, pwsh as the engine and quoted literals' {
        $Identity = "O'Brien@contoso.com"; $CsvPath = ''; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'UsersOnly'; $Snapshot = [long]7
        $line = Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -BatchId '1a2b3c4d' -KeepFilters
        $line | Should -Match '^\.\\Invoke-PraCloudMailbox\.ps1 -Action Convert -Mode Apply -Batch 1a2b3c4d -Identity ''O''''Brien@contoso.com'' -Scope UsersOnly$'
        (Get-PcmCommandLine -ActionName 'Check' -KeepFilters) | Should -Match "-Identity 'O''Brien@contoso.com' -Scope UsersOnly -Snapshot 7$"
        (Get-PcmCommandLine -ActionName 'Collect' -ModeName 'Apply') | Should -Match '^\.\\Invoke-PraCloudMailbox\.ps1 -Action Collect -Mode Apply$'
    }
    It 'keeps -CsvPath, and -ConfigPath only when it is not the default file' {
        $Identity = ''; $CsvPath = 'D:\waves\wave 1.txt'; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'All'; $Snapshot = [long]0; $ConfigPath = 'D:\cfg\my.psd1'
        $script:ConfigPathBound = $false
        (Get-PcmCommandLine -ActionName 'Recover' -BatchId 'abcd1234' -KeepFilters) | Should -Not -Match 'ConfigPath'
        (Get-PcmCommandLine -ActionName 'Recover' -BatchId 'abcd1234' -KeepFilters) | Should -Match "-CsvPath 'D:\\waves\\wave 1.txt'$"
        $script:ConfigPathBound = $true
        (Get-PcmCommandLine -ActionName 'Recover' -BatchId 'abcd1234' -KeepFilters) | Should -Match "-CsvPath 'D:\\waves\\wave 1.txt' -ConfigPath 'D:\\cfg\\my.psd1'$"
        (Get-PcmCommandLine -ActionName 'Check' -OtherServer) | Should -Match '^\.\\Invoke-PraCloudMailbox\.ps1 -Action Check$'
        $script:ConfigPathBound = $false
    }
    It 'proposes after a failure what was kept and the command to run again (<Case>)' -TestCases @(
        @{ Case = 'Preview'; Action = 'Convert'; Mode = 'Preview'; BatchId = ''; Text = 'nothing was changed'; Command = '-Action Convert$' }
        @{ Case = 'Check'; Action = 'Check'; Mode = 'Preview'; BatchId = ''; Text = 'Check only reads'; Command = '-Action Check$' }
        @{ Case = 'Collect'; Action = 'Collect'; Mode = 'Apply'; BatchId = ''; Text = 'Collect only reads'; Command = '^\.\\Invoke-PraCloudMailbox\.ps1 -Action Collect -Mode Apply$' }
        @{ Case = 'Apply before the journal'; Action = 'Convert'; Mode = 'Apply'; BatchId = ''; Text = 'stopped before it wrote its journal'; Command = '-Action Convert -Mode Apply$' }
        @{ Case = 'Apply with a batch'; Action = 'Convert'; Mode = 'Apply'; BatchId = 'f00dbabe'; Text = 'kept in the journal \(batch f00dbabe\)'; Command = '-Action Convert -Mode Apply -Batch f00dbabe$' }) {
        param($Case, $Action, $Mode, $BatchId, $Text, $Command)
        $effectiveMode = $Mode; $Identity = ''; $CsvPath = ''; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'All'; $Snapshot = [long]0; $Batch = ''; $ConfigPath = 'C:\x.psd1'
        $context = @{ Root = 'C:\Tools'; BatchId = $BatchId; Cancelled = $false; NoChangeMade = $false }
        $next = @(Get-PcmFallbackNext)
        $next.Count | Should -Be 2
        $next[0] | Should -Match $Text
        $next[1] | Should -Match $Command
    }
    It 'says that the batch is kept, and that the command continues it, when a resumed Recover stopped before taking its batch back (the sign-in)' {
        $effectiveMode = 'Apply'; $Action = 'Recover'; $Identity = ''; $CsvPath = ''; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'All'; $Snapshot = [long]0; $Batch = 'ABCD1234'; $ConfigPath = 'C:\x.psd1'
        $context = @{ Root = 'C:\Tools'; BatchId = ''; Cancelled = $false; NoChangeMade = $false; ResumeBatchId = 'ee11aa22' }
        $next = @(Get-PcmFallbackNext)
        $next.Count | Should -Be 2
        $next[0] | Should -Match 'The Recover batch ee11aa22 already exists and is kept in the journal'
        $next[0] | Should -Match 'the command below continues it'
        $next[0] | Should -Not -Match 'stopped before it wrote its journal|Nothing was changed'
        $next[1] | Should -Match '-Action Recover -Mode Apply -Batch ABCD1234$'
        $context.ResumeBatchId = ''
        @(Get-PcmFallbackNext)[0] | Should -Match 'stopped before it wrote its journal'
    }
    It 'repeats -ResolveOnly in the command to run again after a Check -Group -ResolveOnly that failed' {
        $effectiveMode = 'Preview'; $Action = 'Check'; $Identity = ''; $CsvPath = ''; $SearchBase = ''; $Group = 'PRA Wave 1'; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'All'; $Snapshot = [long]0; $Batch = ''; $ConfigPath = 'C:\x.psd1'; $ResolveOnly = $true
        $context = @{ Root = 'C:\Tools'; BatchId = ''; Cancelled = $false; NoChangeMade = $false }
        $next = @(Get-PcmFallbackNext)
        $next[0] | Should -Match 'Check only reads'
        $next[1] | Should -Match "^\.\\Invoke-PraCloudMailbox\.ps1 -Action Check -Group 'PRA Wave 1' -ResolveOnly$"
        $ResolveOnly = $false
        @(Get-PcmFallbackNext)[1] | Should -Not -Match 'ResolveOnly'
    }
    It 'proposes the Convert batch of a Recover that failed, with the options of the run' {
        $effectiveMode = 'Apply'; $Action = 'Recover'; $Identity = ''; $CsvPath = 'D:\w.txt'; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'SharedOnly'; $Snapshot = [long]0; $Batch = 'ABCD1234'; $ConfigPath = 'C:\x.psd1'
        $context = @{ Root = 'C:\Tools'; BatchId = '99998888'; Cancelled = $false; NoChangeMade = $false }
        $next = @(Get-PcmFallbackNext)
        $next[0] | Should -Match 'batch 99998888'
        $next[1] | Should -Match '-Action Recover -Mode Apply -Batch ABCD1234 -CsvPath ''D:\\w.txt'' -Scope SharedOnly$'
    }
    It 'says that nothing was written after a refused confirmation' {
        $effectiveMode = 'Apply'; $Identity = ''; $CsvPath = ''; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'All'; $Snapshot = [long]0; $Batch = ''; $ConfigPath = 'C:\x.psd1'; $Action = 'Recover'
        $context = @{ Root = 'C:\Tools'; BatchId = ''; Cancelled = $true }
        @(Get-PcmFallbackNext)[0] | Should -Match 'Nothing was written'
    }
    It 'prints the next step of a failed run, and the errors of the objects first (5, then a count, then the stop line)' {
        $cards = [System.Collections.Generic.List[object]]::new()
        Mock Write-PcmSummary -ModuleName PRA.Common { $cards.Add($Values) }
        Mock Write-Host -ModuleName PRA.Common { }
        $run = @{ Root = $script:Root; Version = 't'; RunId = 'r'; StartTime = Get-Date; Action = 'Convert'; Mode = 'Apply'; Phase = ''; CurrentPhase = 'Convert'; CurrentOperation = 'x'; CurrentIdentity = ''; Warnings = 0
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
            LogFolder = $TestDrive; ReportFolder = $TestDrive; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''
            Config = @{}; BatchId = ''; SnapshotLabel = ''; NextSteps = @('a step that no longer applies'); Journal = $null; FallbackNext = @('What was kept.', 'pwsh -File "x.ps1" -Action Convert -Mode Apply') }
        1..7 | ForEach-Object { Add-PcmIssue $run ("object error $_") }
        Add-PcmIssue $run "Stopped at step 'Convert' (operation x): the run stops"
        $result = Complete-PcmRun -Context $run
        $result.ExitCode | Should -Be 1
        $result.NextSteps | Should -Be @('What was kept.', 'pwsh -File "x.ps1" -Action Convert -Mode Apply')
        $values = $cards[0]
        $values.Contains('Next') | Should -BeTrue
        @($values['Next'][1]) | Should -Be @('What was kept.', 'pwsh -File "x.ps1" -Action Convert -Mode Apply')
        $lines = @($values['Error'][1])
        $lines[1..5] | Should -Be @('object error 1', 'object error 2', 'object error 3', 'object error 4', 'object error 5')
        $lines[6] | Should -Match '^\.\.\. and \d+ more, see the report$'
        $lines[-1] | Should -Match '^Stopped at step'
        $lines.Count | Should -Be 8
    }
    It 'counts only the error of the audit that could not start (no audit of a log that never existed)' {
        $cards = [System.Collections.Generic.List[object]]::new()
        Mock Write-PcmSummary -ModuleName PRA.Common { $cards.Add($Values) }
        Mock Write-Host -ModuleName PRA.Common { }
        $run = @{ Root = $script:Root; Version = 't'; RunId = 'r'; StartTime = Get-Date; Action = 'Collect'; Mode = 'Preview'; Phase = ''; CurrentPhase = 'Start'; CurrentOperation = 'Read-Configuration'; CurrentIdentity = ''; Warnings = 0
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
            LogFolder = $TestDrive; ReportFolder = $TestDrive; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''
            Config = @{}; BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null; FallbackNext = @('fallback') }
        Add-PcmIssue $run 'Audit could not start: The folder X:\logs cannot be created: access denied'
        $result = Complete-PcmRun -Context $run
        $result.ExitCode | Should -Be 1
        @($run.Issues | Where-Object { [string]$_.Message -match 'Audit file' }).Count | Should -Be 0
        @($cards[0]['Error'][1]) | Should -Be @('Audit could not start: The folder X:\logs cannot be created: access denied')
    }
    It 'keeps the own next step of a run that asked for it on failure' {
        $cards = [System.Collections.Generic.List[object]]::new()
        Mock Write-PcmSummary -ModuleName PRA.Common { $cards.Add($Values) }
        Mock Write-Host -ModuleName PRA.Common { }
        $run = @{ Root = $script:Root; Version = 't'; RunId = 'r'; StartTime = Get-Date; Action = 'Convert'; Mode = 'Apply'; Phase = ''; CurrentPhase = 'Convert'; CurrentOperation = 'x'; CurrentIdentity = ''; Warnings = 0
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
            LogFolder = $TestDrive; ReportFolder = $TestDrive; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''
            Config = @{}; BatchId = ''; SnapshotLabel = ''; NextSteps = @('resume it'); NextOnFailure = $true; Journal = $null; FallbackNext = @('fallback') }
        Add-PcmIssue $run 'one object failed'
        (Complete-PcmRun -Context $run).NextSteps | Should -Be @('resume it')
    }
}

Describe 'Engine: confirmation and approval of the changes' {
    BeforeAll {
        $script:ConfigPathBound = $false
        function script:New-FakeCmdlet {
            $cmdlet = [pscustomobject]@{ Answer = $true; Asked = (New-Object 'Collections.Generic.List[string]'); Preference = New-Object 'Collections.Generic.List[string]'; Continue = $true }
            $cmdlet | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($Target, $Operation) $this.Asked.Add("$Operation|$Target"); $this.Preference.Add([string]$ConfirmPreference); $this.Answer }
            $cmdlet | Add-Member -MemberType ScriptMethod -Name ShouldContinue -Value { param($Question, $Title) $this.Asked.Add("question|$Question"); $this.Continue }
            return $cmdlet
        }
    }
    AfterAll {
        foreach ($name in @('Get-MgContext', 'Get-MgRequestContext', 'Set-MgRequestContext', 'Reset-PraFakeGraphTransport', 'Invoke-MgGraphRequest')) { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
        Set-PcmWriteApproval -Approval $null
    }
    It 'has a High ConfirmImpact, supports ShouldProcess and does not make -Action mandatory' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Entry, [ref]$null, [ref]$null)
        $binding = $ast.ParamBlock.Attributes | Where-Object { $_.TypeName.Name -eq 'CmdletBinding' }
        $named = @{}; foreach ($argument in $binding.NamedArguments) { $named[$argument.ArgumentName] = $argument.Argument.Extent.Text }
        $named['ConfirmImpact'] | Should -Match 'High'
        $named['SupportsShouldProcess'] | Should -Match 'true'
        $action = $ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq 'Action' }
        @($action.Attributes | Where-Object { $_.TypeName.Name -eq 'Parameter' } | ForEach-Object { $_.NamedArguments } | Where-Object { $_.ArgumentName -eq 'Mandatory' }).Count | Should -Be 0
    }
    It 'asks ShouldProcess before a change in Apply and never writes in a Preview' {
        $cmdlet = New-FakeCmdlet
        $approval = New-PcmApprovalCallback $cmdlet @{ Mode = 'Apply'; ConfirmPreference = 'High' }
        & $approval 'users/u1' 'Microsoft Graph PATCH' | Should -BeTrue
        $cmdlet.Asked | Should -Be @('Microsoft Graph PATCH|users/u1')
        $cmdlet.Preference | Should -Be @('High')
        $preview = New-PcmApprovalCallback $cmdlet @{ Mode = 'preview'; ConfirmPreference = 'None' }
        { & $preview 'users/u1' 'Microsoft Graph PATCH' } | Should -Throw '*A Preview never writes*'
        $cmdlet.Asked.Count | Should -Be 1
    }
    It 'asks nothing when the run confirmation preference is None (after the one confirmation or with -Force), and -WhatIf refuses the change' {
        function Invoke-ProbeApproval {
            [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
            param([hashtable]$Run)
            $approval = New-PcmApprovalCallback $PSCmdlet $Run
            & $approval 'users/u1' 'Microsoft Graph PATCH'
        }
        Invoke-ProbeApproval -Run @{ Mode = 'Apply'; ConfirmPreference = 'None' } | Should -BeTrue
        Invoke-ProbeApproval -Run @{ Mode = 'Apply'; ConfirmPreference = 'None' } -WhatIf | Should -BeFalse
        { Invoke-ProbeApproval -Run @{ Mode = 'Preview'; ConfirmPreference = 'None' } } | Should -Throw '*A Preview never writes*'
    }
    It 'asks one confirmation for the run, none with -Force, -Confirm or -Confirm:$false, and then lets the changes through' {
        $cmdlet = New-FakeCmdlet
        $caller = $cmdlet
        $context = @{ Mode = 'Apply'; ConfirmPreference = 'High'; Cancelled = $false }
        $Force = [switch]$false; $script:ConfirmBound = $false
        Mock Write-PcmLog { }
        Confirm-PcmApply 'Convert 2 objects?'
        $cmdlet.Asked | Should -Be @('question|Convert 2 objects?')
        $context.ConfirmPreference | Should -Be 'None'
        $cmdlet.Asked.Clear(); $context.ConfirmPreference = 'High'
        $Force = [switch]$true
        Confirm-PcmApply 'Convert 2 objects?'
        $Force = [switch]$false; $script:ConfirmBound = $true
        Confirm-PcmApply 'Convert 2 objects?'
        $cmdlet.Asked.Count | Should -Be 0
        $script:ConfirmBound = $false
    }
    It 'stops the run, marked as cancelled, when the operator says no' {
        $cmdlet = New-FakeCmdlet; $cmdlet.Continue = $false
        $caller = $cmdlet
        $context = @{ Mode = 'Apply'; ConfirmPreference = 'High'; Cancelled = $false }
        $Force = [switch]$false; $script:ConfirmBound = $false
        { Confirm-PcmApply 'Convert 2 objects?' } | Should -Throw '*Cancelled by the operator: nothing was changed*'
        $context.Cancelled | Should -BeTrue
        $context.ConfirmPreference | Should -Be 'High'
    }
    It 'asks the approval of the run before every Microsoft Graph change, never for a read, and sends nothing when it is refused' {
        Reset-PraFakeGraphTransport
        $script:asked = [System.Collections.Generic.List[string]]::new()
        $script:wire = [System.Collections.Generic.List[string]]::new()
        if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { function global:Invoke-MgGraphRequest { param($Method, $Uri, $Body, $ContentType, $OutputType, $Headers) } }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { $script:wire.Add("$Method $Uri"); if ($Uri -match 'batch') { [pscustomobject]@{ responses = @([pscustomobject]@{ id = '0'; status = 204; body = $null }) } } else { [pscustomobject]@{ value = @() } } }
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $script:asked.Add("$Operation|$Target"); $true }
        $null = Invoke-PcmGraph -Uri 'v1.0/users/u1?$select=id'
        $null = Invoke-PcmGraph -Method POST -Uri 'v1.0/users/u1/checkMemberGroups' -Body @{ groupIds = @('g') }
        $script:asked.Count | Should -Be 0
        $null = Invoke-PcmGraph -Method PATCH -Uri 'v1.0/users/u1?x=1' -Body @{ usageLocation = 'FR' }
        $script:asked | Should -Be @('Microsoft Graph PATCH|users/u1')
        $script:asked.Clear()
        $null = Invoke-PcmGraphBatch -Requests @(@{ Method = 'GET'; Url = '/users/u1' }, @{ Method = 'PATCH'; Url = '/users/u2'; Body = @{ a = 1 } }, @{ Method = 'DELETE'; Url = '/users/u3' })
        $script:asked | Should -Be @('Microsoft Graph PATCH|users/u2', 'Microsoft Graph DELETE|users/u3')
        $script:wire.Clear()
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $false }
        { Invoke-PcmGraph -Method DELETE -Uri 'v1.0/users/u1' } | Should -Throw '*Operation refused*'
        { Invoke-PcmGraphBatch -Requests @(@{ Method = 'DELETE'; Url = '/users/u3' }) } | Should -Throw '*Operation refused*'
        $script:wire.Count | Should -Be 0
        Set-PcmWriteApproval -Approval $null
        { Invoke-PcmGraph -Method PATCH -Uri 'v1.0/users/u1' -Body @{ a = 1 } } | Should -Throw '*No approval is registered*'
        $script:wire.Count | Should -Be 0
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $true }
    }
    It 'refuses every Exchange Online and case hold change without an approval' {
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $false }
        { Confirm-PcmWrite -Target 'x' -Operation 'y' } | Should -Throw '*Operation refused: y on x*'
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $true }
        { Confirm-PcmWrite -Target 'x' -Operation 'y' } | Should -Not -Throw
        $source = [IO.File]::ReadAllText((Join-Path $script:Root 'module\PRA.Cloud.psm1')) + $script:EntrySource
        foreach ($cmdlet in 'Set-MailUser', 'Set-Mailbox', 'New-Mailbox', 'New-DistributionGroup', 'Set-DistributionGroup', 'Add-DistributionGroupMember', 'New-MailContact', 'Set-MailContact', 'Set-CalendarProcessing',
            'Set-MailboxCalendarConfiguration', 'Set-User', 'Add-MailboxPermission', 'Add-RecipientPermission', 'Set-CaseHoldPolicy', 'New-CaseHoldPolicy') {
            $asts = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
            $calls = @($asts.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq $cmdlet }, $true))
            foreach ($call in $calls) {
                $function = $call.Parent; while ($function -and $function -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $function = $function.Parent }
                $function | Should -Not -BeNullOrEmpty
                $function.Extent.Text | Should -Match 'Confirm-PcmWrite|Invoke-PcmHoldChange' -Because "$cmdlet is called in $($function.Name)"
            }
        }
    }
}

Describe 'Engine: cloud modules and versions' {
    BeforeAll {
        function script:New-Mod { param($Name, $Version) [pscustomobject]@{ Name = $Name; Version = [version]$Version } }
    }
    It 'accepts Microsoft.Graph.Authentication 2.41.1 and ExchangeOnlineManagement 3.10.1' {
        $report = Get-PcmCloudModuleReport -Available @((New-Mod 'Microsoft.Graph.Authentication' '2.41.1'), (New-Mod 'ExchangeOnlineManagement' '3.10.1')) -Loaded @()
        $report.Ok | Should -BeTrue
        $report.Graph.Version | Should -Be '2.41.1'
        $report.Exo.Version | Should -Be '3.10.1'
    }
    It 'gives ready Install-Module commands, one module per command, when a module is missing or too old' {
        $report = Get-PcmCloudModuleReport -Available @((New-Mod 'Microsoft.Graph.Authentication' '2.40.0'), (New-Mod 'ExchangeOnlineManagement' '3.9.0')) -Loaded @()
        $report.Ok | Should -BeFalse
        $report.Graph.Message | Should -Match 'found 2\.40\.0'
        $report.Graph.Message | Should -Match 'Install-Module Microsoft\.Graph\.Authentication -MinimumVersion 2\.41\.1 -Scope CurrentUser'
        $report.Exo.Message | Should -Match 'found 3\.9\.0'
        $report.Exo.Message | Should -Match 'Install-Module ExchangeOnlineManagement -MinimumVersion 3\.10\.1 -Scope CurrentUser'
        foreach ($message in $report.Messages) { ([regex]::Matches($message, 'Install-Module')).Count | Should -Be 1 }
        (Get-PcmCloudModuleReport -Available @() -Loaded @()).Messages.Count | Should -Be 2
    }
    It 'takes the highest installed version, and never goes under the built-in Exchange Online minimum' {
        $report = Get-PcmCloudModuleReport -MinExo '3.1.0' -Available @((New-Mod 'Microsoft.Graph.Authentication' '2.41.1'), (New-Mod 'Microsoft.Graph.Authentication' '2.50.0'), (New-Mod 'ExchangeOnlineManagement' '3.9.9')) -Loaded @()
        $report.Graph.Version | Should -Be '2.50.0'
        $report.Exo.Ok | Should -BeFalse
        (Get-PcmCloudModuleReport -MinExo '3.20.0' -Available @((New-Mod 'Microsoft.Graph.Authentication' '2.41.1'), (New-Mod 'ExchangeOnlineManagement' '3.10.1')) -Loaded @()).Exo.Ok | Should -BeFalse
    }
    It 'needs the Microsoft Graph modules already loaded in the session to be of one version, the minimum or later' {
        $available = @((New-Mod 'Microsoft.Graph.Authentication' '2.41.1'), (New-Mod 'ExchangeOnlineManagement' '3.10.1'))
        (Get-PcmCloudModuleReport -Available $available -Loaded @((New-Mod 'Microsoft.Graph.Authentication' '2.41.1'), (New-Mod 'Microsoft.Graph.Users' '2.41.1'))).Ok | Should -BeTrue
        $mixed = Get-PcmCloudModuleReport -Available $available -Loaded @((New-Mod 'Microsoft.Graph.Authentication' '2.41.1'), (New-Mod 'Microsoft.Graph.Users' '2.30.0'))
        $mixed.Ok | Should -BeFalse
        $mixed.Graph.Message | Should -Match 'several versions are loaded'
        (Get-PcmCloudModuleReport -Available $available -Loaded @((New-Mod 'Microsoft.Graph.Users' '2.30.0'))).Graph.Message | Should -Match 'needs version 2\.41\.1 or later'
    }
    It 'stops the cloud actions before any change when the modules are missing, also in a Preview' {
        $source = $script:EntrySource
        $source | Should -Match 'if \(\$Action -notin @\(''Collect'', ''Schedule''\)\) \{ Test-PcmCloudModuleAvailable \}'
        $source | Should -Match 'Nothing was changed\.'' -f \(\$report\.Messages'
    }
    It 'asks for PowerShell 7.5 or later for every action, Collect included' {
        $script:EntrySource | Should -Match "\[version\]'7\.5'"
        $script:EntrySource | Should -Not -Match '7\.4'
        $script:EntrySource | Should -Match '#Requires -Version 7\.5'
        $script:EntrySource | Should -Not -Match 'Windows PowerShell 5\.1 \(powershell\.exe\)'
        [IO.File]::ReadAllText((Join-Path $script:Root 'module\PRA.Gui.psm1')) | Should -Not -Match '7\.4'
    }
}

Describe 'Targets: options of the run' {
    BeforeEach {
        $script:BoundOptions = @()
        $Identity = ''; $CsvPath = ''; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $Expect = ''; $Once = $false; $ResolveOnly = $false; $Batch = ''; $Action = 'Convert'
        $context = @{ NoChangeMade = $false }
    }
    It 'refuses to name the objects in two ways: <First> with <Second>' -TestCases @(
        @{ First = '-Identity'; Second = '-CsvPath'; Set = { $Identity = 'a@contoso.com'; $CsvPath = 'wave.csv' } }
        @{ First = '-Identity'; Second = '-SearchBase'; Set = { $Identity = 'a@contoso.com'; $SearchBase = 'OU=Paris,DC=contoso,DC=com' } }
        @{ First = '-Identity'; Second = '-Group'; Set = { $Identity = 'a@contoso.com'; $Group = 'PRA-Wave1' } }
        @{ First = '-CsvPath'; Second = '-SearchBase'; Set = { $CsvPath = 'wave.csv'; $SearchBase = 'contoso.com/Paris' } }
        @{ First = '-CsvPath'; Second = '-Group'; Set = { $CsvPath = 'wave.csv'; $Group = 'PRA-Wave1' } }
        @{ First = '-SearchBase'; Second = '-Group'; Set = { $SearchBase = 'contoso.com/Paris'; $Group = 'PRA-Wave1' } }) {
        param($First, $Second, $Set)
        . $Set
        { Test-PcmRunOption } | Should -Throw "*Choose one way to say which objects: $First and $Second cannot be used together*"
        $context.NoChangeMade | Should -BeTrue
    }
    It 'accepts one way to name the objects with -Scope and -MaxObjects, for every action but Collect' -TestCases @(
        @{ Name = 'Identity'; Set = { $Identity = 'a@contoso.com' } }, @{ Name = 'CsvPath'; Set = { $CsvPath = 'wave.csv' } }
        @{ Name = 'SearchBase'; Set = { $SearchBase = 'OU=Paris,DC=contoso,DC=com' } }, @{ Name = 'Group'; Set = { $Group = 'pra-wave1' } }) {
        param($Name, $Set)
        . $Set
        $MaxObjects = 20
        { Test-PcmRunOption } | Should -Not -Throw
        (Get-PcmTargetKind) | Should -Be $(if ($Name -eq 'CsvPath') { 'Csv' } else { $Name })
    }
    It 'refuses -CsvPath, -SearchBase, -Group and -MaxObjects for Collect, but not -Identity' {
        $Action = 'Collect'
        $Identity = 'a@contoso.com'
        { Test-PcmRunOption } | Should -Not -Throw
        $Identity = ''; $Group = 'PRA-Wave1'
        { Test-PcmRunOption } | Should -Throw '*are for Check, Convert and Recover*Collect reads the scope of the configuration*'
        $Group = ''; $MaxObjects = 5
        { Test-PcmRunOption } | Should -Throw '*are for Check, Convert and Recover*'
    }
    It 'refuses an empty value or a control character in -SearchBase, -Group and -CsvPath' {
        $script:BoundOptions = @('Group')
        $Group = '   '
        { Test-PcmRunOption } | Should -Throw '*-Group is empty*'
        $script:BoundOptions = @(); $Group = ''
        $SearchBase = "OU=Paris`r`n,DC=contoso,DC=com"
        { Test-PcmRunOption } | Should -Throw '*-SearchBase holds a control character*'
    }
    It 'uses -Expect and -Once only with Check -Batch, whatever their letter case' -TestCases @(@{ Used = 'Expect' }, @{ Used = 'Once' }) {
        param($Used)
        if ($Used -eq 'Expect') { $Expect = 'Provisioned' } else { $Once = $true }
        $Action = 'Check'
        { Test-PcmRunOption } | Should -Throw '*are for Check with -Batch*'
        $Action = 'convert'; $Batch = 'abcd1234'
        { Test-PcmRunOption } | Should -Throw '*are for Check with -Batch*'
        $Action = 'check'
        { Test-PcmRunOption } | Should -Not -Throw
    }
    It 'accepts -Expect in any letter case and keeps the official spelling' {
        $parameter = (Get-Command $script:Entry).Parameters['Expect']
        $set = @($parameter.Attributes | Where-Object { $_ -is [Management.Automation.ValidateSetAttribute] })[0]
        @($set.ValidValues) | Should -Be @('Provisioned', 'Deprovisioned', 'Retained')
        $script:EntrySource | Should -Match 'Set-Variable -Name \$parameterInfo\.Name -Value \(\[string\]\$official\[0\]\)'
        (Get-Command $script:Entry).Parameters['CsvPath'].Aliases | Should -Contain 'IdentityPath'
    }
}

Describe 'Targets: command line keeps the source and the filters' {
    BeforeAll { $script:ConfigPathBound = $false; $context = @{ Root = 'C:\Tools' } }
    BeforeEach { $Identity = ''; $CsvPath = ''; $SearchBase = ''; $Group = ''; $MaxObjects = 0; $BaseConfigPath = ''; $Scope = 'All'; $Snapshot = [long]0; $ConfigPath = 'C:\x.psd1' }
    It 'keeps -Group, -SearchBase and -MaxObjects, quoted as literals, after the source and before -Snapshot' {
        $Group = "O'Brien team"; $MaxObjects = 25; $Scope = 'UsersOnly'; $Snapshot = [long]4
        (Get-PcmCommandLine -ActionName 'Convert' -ModeName 'Apply' -KeepFilters) | Should -Match "-Action Convert -Mode Apply -Group 'O''Brien team' -Scope UsersOnly -MaxObjects 25 -Snapshot 4$"
        $Group = ''; $SearchBase = 'OU=Paris,DC=contoso,DC=com'
        (Get-PcmCommandLine -ActionName 'Check' -KeepFilters) | Should -Match "-Action Check -SearchBase 'OU=Paris,DC=contoso,DC=com' -Scope UsersOnly -MaxObjects 25 -Snapshot 4$"
        (Get-PcmCommandLine -ActionName 'Recover' -BatchId 'abcd1234' -KeepFilters) | Should -Match "-Batch abcd1234 -SearchBase 'OU=Paris,DC=contoso,DC=com' -Scope UsersOnly -MaxObjects 25$"
        (Get-PcmCommandLine -ActionName 'Recover' -BatchId 'abcd1234') | Should -Not -Match 'SearchBase|MaxObjects'
    }
    It 'shows the original configuration file (-BaseConfigPath) in the next steps, never the generated one' {
        $ConfigPath = 'C:\run\generated.psd1'; $BaseConfigPath = 'D:\cfg\original.psd1'
        $script:ConfigPathBound = $true
        $line = Get-PcmCommandLine -ActionName 'Check' -KeepFilters
        $line | Should -Match "-ConfigPath 'D:\\cfg\\original.psd1'$"
        $line | Should -Not -Match 'generated'
        $script:ConfigPathBound = $false
    }
}

Describe 'Targets: organizational unit, group members and records' {
    It 'tells whether an object is under an OU, sub-OUs included, with a DN or a canonical name in any letter case' -TestCases @(
        @{ Base = 'OU=Paris,DC=contoso,DC=com'; Dn = 'CN=Ann,OU=Staff,OU=Paris,DC=contoso,DC=com'; Ou = 'contoso.com/Paris/Staff'; Expected = $true }
        @{ Base = 'ou=paris, dc=CONTOSO , dc=com'; Dn = 'CN=Ann,OU=Staff,OU=Paris,DC=contoso,DC=com'; Ou = ''; Expected = $true }
        @{ Base = 'OU=ParisNord,DC=contoso,DC=com'; Dn = 'CN=Ann,OU=Staff,OU=Paris,DC=contoso,DC=com'; Ou = 'contoso.com/Paris/Staff'; Expected = $false }
        @{ Base = 'OU=Staff,OU=Paris,DC=contoso,DC=com'; Dn = ''; Ou = 'contoso.com/Paris/Staff'; Expected = $true }
        @{ Base = 'OU=Paris,DC=contoso,DC=com'; Dn = ''; Ou = 'contoso.com/ParisNord'; Expected = $false }
        @{ Base = 'CONTOSO.COM/paris'; Dn = ''; Ou = 'contoso.com/Paris/Staff'; Expected = $true }
        @{ Base = 'contoso.com/Paris/'; Dn = ''; Ou = 'contoso.com/Paris'; Expected = $true }
        @{ Base = 'contoso.com/Paris'; Dn = 'CN=Ann,OU=Paris,DC=contoso,DC=com'; Ou = ''; Expected = $true }
        @{ Base = 'contoso.com/Par'; Dn = 'CN=Ann,OU=Paris,DC=contoso,DC=com'; Ou = 'contoso.com/Paris'; Expected = $false }
        @{ Base = 'contoso.com/Lyon'; Dn = ''; Ou = ''; Expected = $false }
        @{ Base = ''; Dn = 'CN=Ann,OU=Paris,DC=contoso,DC=com'; Ou = ''; Expected = $false }) {
        param($Base, $Dn, $Ou, $Expected)
        Test-PcmInSearchBase -SearchBase $Base -DistinguishedName $Dn -OrganizationalUnit $Ou | Should -Be $Expected
    }
    It 'turns a distinguished name into its canonical name, and keeps an escaped comma' {
        ConvertFrom-PcmDnToCanonical 'CN=Ann,OU=Staff,OU=Paris,DC=contoso,DC=com' | Should -Be 'contoso.com/Paris/Staff/Ann'
        ConvertFrom-PcmDnToCanonical 'OU=Paris,DC=contoso,DC=com' | Should -Be 'contoso.com/Paris'
        ConvertTo-PcmDnKey 'cn=Ann\, Bob , ou=Staff' | Should -Be 'cn=Ann\, Bob,ou=Staff'
    }
    It 'finds the mailbox of a group member by GUID, immutable ID, UPN, then SMTP address, letter case ignored but for the immutable ID' {
        $ann = [pscustomobject]@{ object_guid = 'a0000000-0000-0000-0000-00000000000a'; immutable_id = 'AbCdEf=='; user_principal_name = 'Ann@Contoso.com'; primary_smtp_address = 'ann.s@contoso.com' }
        $bob = [pscustomobject]@{ object_guid = 'b0000000-0000-0000-0000-00000000000b'; immutable_id = 'abcdef=='; user_principal_name = 'bob@contoso.com'; primary_smtp_address = 'bob@contoso.com' }
        $index = New-PcmRecordIndex -Records @($ann, $bob)
        (Find-PcmRecord -Index $index -Guid '{A0000000-0000-0000-0000-00000000000A}').user_principal_name | Should -Be 'Ann@Contoso.com'
        (Find-PcmRecord -Index $index -ImmutableId 'AbCdEf==').user_principal_name | Should -Be 'Ann@Contoso.com'
        (Find-PcmRecord -Index $index -ImmutableId 'abcdef==').user_principal_name | Should -Be 'bob@contoso.com'
        Find-PcmRecord -Index $index -ImmutableId 'ABCDEF==' | Should -BeNullOrEmpty
        (Find-PcmRecord -Index $index -Upn 'ANN@CONTOSO.COM').object_guid | Should -Be 'a0000000-0000-0000-0000-00000000000a'
        (Find-PcmRecord -Index $index -Smtp 'ANN.S@contoso.com').object_guid | Should -Be 'a0000000-0000-0000-0000-00000000000a'
        (Find-PcmRecord -Index $index -ImmutableId 'nope' -Upn 'bob@CONTOSO.com').object_guid | Should -Be 'b0000000-0000-0000-0000-00000000000b'
        Find-PcmRecord -Index $index -Upn 'carol@contoso.com' | Should -BeNullOrEmpty
    }
    It 'knows the Exchange system mailboxes' {
        foreach ($name in @('HealthMailbox0123', 'SystemMailbox{1f05a927-2}', 'Microsoft Exchange Approval Assistant', 'DiscoverySearchMailbox {D919BA05}', 'Migration.8f3e7716', 'FederatedEmail.4c1f4d8b', 'SM_abc')) {
            Test-PcmSystemMailboxName $name | Should -BeTrue -Because $name
        }
        Test-PcmSystemMailboxName 'Alice' | Should -BeFalse
        Test-PcmSystemMailboxName '' | Should -BeFalse
    }
    It 'reads Collect.TargetGroups in any letter case, drops empty and repeated names, and refuses a control character' {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $folder
        $path = Join-Path $folder 'g.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; collect = @{ targetgroups = @('PRA-Wave1', ' pra-WAVE1 ', '', 'CN=G2,DC=contoso,DC=com') } }"
        $config = Import-PcmConfiguration -Path $path -Root $folder
        @($config.Collect.TargetGroups) | Should -Be @('PRA-Wave1', 'CN=G2,DC=contoso,DC=com')
        Set-Content -LiteralPath $path -Encoding UTF8 -Value ("@{ Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Collect = @{ TargetGroups = @('a" + [char]0x0007 + "b') } }")
        { Import-PcmConfiguration -Path $path -Root $folder } | Should -Throw '*Collect.TargetGroups holds a value with a control character*'
    }
}

Describe 'Case: no new case-sensitive comparison' {
    BeforeAll {
        # Every case-sensitive operator or ordinal comparer of the runtime code is listed here with the reason it is a machine identifier, never something
        # an operator types. A new one fails this test: it must be a comparison that ignores the letter case, or be added here with its reason.
        $script:Pattern = '(?<![\w-])-c(eq|ne|match|notmatch|like|notlike|in|notin|contains|notcontains)\b|StringComparer\]::Ordinal\b|StringComparison\]::Ordinal\b|-CaseSensitive\b'
        $script:Allowed = @(
            @{ File = 'Invoke-PraCloudMailbox.ps1'; Text = '[string]$found[0].Name -cne $name'; Count = 1; Reason = 'Routing contact: the creation name carries the random ownership token written in the journal; ownership proof, not an operator value.' }
            @{ File = 'Invoke-PraCloudMailbox.ps1'; Text = '$marker -cne [string]$Intent.token'; Count = 1; Reason = 'Ownership marker against the random token of the journal.' }
            @{ File = 'Invoke-PraCloudMailbox.ps1'; Text = '-cne (Get-PcmResourceProtectionMode $intent)'; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (None, CaseHold): another spelling means the journal was not written by the tool.' }
            @{ File = 'Invoke-PraCloudMailbox.ps1'; Text = '$before -ceq $after'; Count = 1; Reason = 'Ownership fields of the journal (token, creation name, marker attribute, type); the GUID fields are compared with Test-PcmSameText just above.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '$byImmutable = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)'; Count = 1; Reason = 'onPremisesImmutableId is base64: two identities can differ only by letter case.' }
            @{ File = 'module\PRA.Common.psm1'; Text = 'ByImmutable = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)'; Count = 1; Reason = 'onPremisesImmutableId (immutable_id) is base64: two group members can differ only by letter case, so it is matched exactly (GUID, UPN and SMTP address are matched with the case ignored).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = 'HashSet[string]]::new([StringComparer]::Ordinal)'; Count = 5; Reason = 'Immutable IDs (base64, case significant), the filter predicates that may hold one, and the exact next links already visited.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cnotmatch ("^PRA-'; Count = 2; Reason = 'Creation name of a native resource or room list: it embeds the random ownership token of the journal.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cnotin @(''None'',''CaseHold'')'; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (protectionMode).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cnotin @(''None'', ''CaseHold'')'; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (protectionMode).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cnotin @('''', ''Intent'', ''Issuing'', ''Issued'', ''Verified'')'; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (deletion progress).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cne ''CloudResource'''; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (lifecycle).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '(Get-PcmResourceProtectionMode $Intent) -cne ''None'''; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (protectionMode).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '(Get-PcmUnprotectedDeletionProgress $Intent) -cne ''Verified'''; Count = 1; Reason = 'Closed vocabulary written by the tool in the journal (deletion progress).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '[string]$Record.kind -cnotin @(''Room'',''Equipment'')'; Count = 1; Reason = 'Kind written by Collect from a fixed map; part of the tombstone proof against the journal.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cne "$($Record.kind)Mailbox"'; Count = 1; Reason = 'Cloud type written by the tool in the journal, tombstone proof.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-ceq [string]$Intent.cloudType'; Count = 1; Reason = 'Recipient type name of the tombstone, against the cloud type written by the tool in the journal.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cmatch ''^SMTP:'''; Count = 1; Reason = 'The prefix SMTP: (primary address) against smtp: (secondary) is a machine flag whose letter case has a meaning.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = 'ManagementObjectNotFoundException'; Count = 4; Reason = '.NET exception type names and Exchange error identifiers (machine identifiers, never typed).' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = 'Sort-Object -CaseSensitive'; Count = 1; Reason = 'Order of generated filter clauses only (stable output); nothing is compared.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '$parsedUri.Scheme -cne ''https'''; Count = 1; Reason = 'Next link of a Graph page: the scheme of a parsed URI is a machine value; guard against redirection.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '$parsedUri.AbsolutePath -cne $query.Path'; Count = 1; Reason = 'Next link of a Graph page must keep the exact path of the query: guard against redirection.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '$query.Path -ceq ''/v1.0/users'''; Count = 1; Reason = 'The path of a query built by the tool itself.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '$userFacts[$user.id] -cne $facts'; Count = 1; Reason = 'JSON of machine facts returned by Graph for one user ID.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '-cnotin $Inventory.GraphImmutableIds'; Count = 1; Reason = 'onPremisesImmutableId is base64: case significant.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '''onPremisesImmutableId'' '''') -ceq $immutableId'; Count = 1; Reason = 'onPremisesImmutableId is base64: case significant.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '[string]$recipient.Name -cne [string]$Intent.creationName'; Count = 1; Reason = 'Creation name with the random ownership token of the journal.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '[string]$candidate.Name -cne [string]$Intent.creationName'; Count = 1; Reason = 'Creation name with the random ownership token of the journal.' }
            @{ File = 'module\PRA.Cloud.psm1'; Text = '$marker -cne [string]$Intent.token'; Count = 2; Reason = 'Ownership marker against the random token of the journal.' }
            @{ File = 'module\PRA.Common.psm1'; Text = '[string]::Equals($_, ''.'', [StringComparison]::Ordinal) -or [string]::Equals($_, ''..'', [StringComparison]::Ordinal)'; Count = 1; Reason = 'Device-namespace path segments: file system syntax, not an operator value.' }
            @{ File = 'module\PRA.Common.psm1'; Text = '$existing.StartsWith(''\\'', [StringComparison]::Ordinal)'; Count = 1; Reason = 'UNC prefix of a path resolved by Windows: file system syntax, not an operator value.' }
            @{ File = 'module\PRA.Common.psm1'; Text = '$physical.StartsWith(''\\?\'', [StringComparison]::Ordinal)'; Count = 1; Reason = 'Device-namespace prefix of a path resolved by Windows: file system syntax, not an operator value.' }
            @{ File = 'module\PRA.Common.psm1'; Text = '[string]::Equals($first, $Canonical, [StringComparison]::Ordinal)'; Count = 2; Reason = 'The canonical folder names data, logs and reports are the exact names that the integrity verifier (ordinal) accepts at the root of the package: a machine identifier of the package layout.' }
            @{ File = 'module\PRA.Gui.psm1'; Text = '.StartsWith('; Count = 9; Reason = 'Prefixes of the names of the controls of the window (R_, Mode_, Fl_, Th_, C_Scope_, S_Expect_, S_Wave_, T_, Nav, F_): identifiers written in the XAML, never typed by an operator.' }
            @{ File = 'module\PRA.Gui.Process.ps1'; Text = '[Collections.IComparer][StringComparer]::Ordinal)'; Count = 1; Reason = 'The environment block given to CreateProcess is sorted by upper-cased variable name, by UTF-16 code unit: a Windows requirement on names of variables, not a comparison of what an operator types.' }
        )
        # The integrity bootstrap and the verifier compare signature states, thumbprints and SHA256 hashes: machine identifiers that stay exact, by design.
        $files = Get-ChildItem -LiteralPath $script:Root -Recurse -File | Where-Object { $_.Extension -in @('.ps1', '.psm1') -and $_.Name -ne 'Test-PackageIntegrity.ps1' -and $_.FullName -notmatch '\\(tests|lib|logs|reports|data|docs|\.git)\\' }
        $script:Hits = @(foreach ($file in $files) {
                $relative = $file.FullName.Substring($script:Root.Length + 1)
                $number = 0
                $inBootstrap = $false
                foreach ($line in [IO.File]::ReadAllLines($file.FullName)) {
                    $number++
                    if ($line.StartsWith('# >>> PRA package integrity bootstrap')) { $inBootstrap = $true }
                    if ($inBootstrap) { if ($line.StartsWith('# <<< PRA package integrity bootstrap')) { $inBootstrap = $false }; continue }
                    if ($line -match $script:Pattern) { [pscustomobject]@{ File = $relative; Line = $number; Text = $line.Trim() } }
                }
            })
    }
    It 'finds only listed comparisons' {
        $unlisted = @($script:Hits | Where-Object {
                $hit = $_
                -not @($script:Allowed | Where-Object { $_.File -eq $hit.File -and $hit.Text.Contains($_.Text) }).Count
            } | ForEach-Object { '{0}:{1}: {2}' -f $_.File, $_.Line, $_.Text })
        $unlisted | Should -BeNullOrEmpty -Because 'a case-sensitive comparison must be justified in the allow-list'
    }
    It 'lists no comparison that is not in the code any more' {
        $stale = @($script:Allowed | Where-Object {
                $entry = $_
                @($script:Hits | Where-Object { $_.File -eq $entry.File -and $_.Text.Contains($entry.Text) }).Count -ne $entry.Count
            } | ForEach-Object { '{0}: {1} (expected {2})' -f $_.File, $_.Text, $_.Count })
        $stale | Should -BeNullOrEmpty
    }
}
