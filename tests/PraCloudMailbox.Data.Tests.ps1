<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the snapshot store, the link between a batch and its snapshot, the size of a Collect, the hash files of a Collect, the groups and the lists
    read by Collect, and the Script mode of Entra Connect (Pester 5+). No Exchange, no cloud: Collect runs in a child process against synthetic Exchange commands.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
#>
#Requires -Version 7.5

BeforeDiscovery {
    $script:AnsiPage = [Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
    $script:AnsiUsable = $script:AnsiPage -ne 65001 -and $script:AnsiPage -ne 0
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    Import-Module (Join-Path $script:Root 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Store.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Collect.psm1') -Force
    $script:Entry = Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    $script:AnsiPage = [Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
    $script:Work = Join-Path ([IO.Path]::GetTempPath()) ('PCM-data-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $null = New-Item -ItemType Directory -Path $script:Work -Force
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($script:Entry, [ref]$null, [ref]$null)
    foreach ($function in @($ast.EndBlock.Statements | Where-Object { $_ -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $_.Name -in @('ConvertTo-PcmIdentityKey', 'Read-PcmIdentityFile') })) {
        . ([scriptblock]::Create($function.Extent.Text))
    }

    function script:New-DataConfig {
        param([Parameter(Mandatory)][string]$Folder, [string]$StoreExtra = '', [string]$CollectExtra = '', [string]$BackupFolder = '')
        $null = New-Item -ItemType Directory -Path $Folder -Force
        if (-not $BackupFolder) { $BackupFolder = Join-Path $Folder 'backup' }
        $text = @"
@{
    Environment = 'TEST'
    Exchange = @{ Server = 'ex1.contoso.com' }
    Scope = @{ Mode = 'Auto'; IncludeUsers = `$true; IncludeShared = `$true }
    Collect = @{ SharedPermissions = `$true; $CollectExtra }
    Store = @{ Path = '$Folder\db\pra.db'; BackupFolder = '$BackupFolder'; JournalPath = '$Folder\db\journal.db'; $StoreExtra }
    EntraConnect = @{ Mode = 'Manual' }
    Logging = @{ Folder = '$Folder\logs' }
    Report = @{ Folder = '$Folder\reports' }
}
"@
        $path = Join-Path $Folder 'config.psd1'
        Set-Content -LiteralPath $path -Value $text -Encoding UTF8
        $path
    }

    function script:Invoke-Entry {
        <# Runs the entry script in a child PowerShell with synthetic Exchange commands: PCM_TEST_MAILBOXES user mailboxes (default 3), PCM_TEST_EXCHANGE_FAILS = an Exchange that fails. #>
        param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string[]]$Arguments, [int]$Mailboxes = 3, [switch]$ExchangeFails)
        $runner = Join-Path $Folder ('run-{0}.ps1' -f [guid]::NewGuid().ToString('N').Substring(0, 6))
        $body = @'
$ErrorActionPreference = 'Stop'
function global:New-PSSession { param($ConfigurationName, $ConnectionUri, $Authentication) [pscustomobject]@{ Id = 1; Name = 'synthetic' } }
function global:Import-PSSession { param($Session, $CommandName, [switch]$DisableNameChecking, [switch]$AllowClobber) New-Module -Name PcmSyntheticExchange -ScriptBlock { } }
function global:Remove-PSSession { param($Session) }
function global:Get-ExchangeServer { @([pscustomobject]@{ Name = 'ex1'; AdminDisplayVersion = 'Version 15.2 (Build 1748.10)' }) }
function global:Set-ADServerSettings { param($ViewEntireForest) }
function global:Get-Mailbox {
    param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController)
    if ($env:PCM_TEST_EXCHANGE_FAILS) { throw 'The WS-Management service cannot process the request: the server did not answer.' }
    $count = [int]$env:PCM_TEST_MAILBOXES
    @($(if ($count -gt 0) { 1..$count }) | ForEach-Object {
        $sam = 'user{0:00}' -f $_
        [pscustomobject]@{ Guid = [guid]::NewGuid(); Identity = $sam; Name = $sam; DisplayName = $sam; Alias = $sam; SamAccountName = $sam
            UserPrincipalName = "$sam@contoso.com"; DistinguishedName = "CN=$sam,OU=PRA,DC=contoso,DC=com"; OrganizationalUnit = 'contoso.com/PRA'; PrimarySmtpAddress = "$sam@contoso.com"
            WindowsEmailAddress = "$sam@contoso.com"; LegacyExchangeDN = "/o=C/cn=$sam"; ExchangeGuid = [guid]::NewGuid(); ArchiveGuid = [guid]::Empty; ArchiveStatus = 'None'
            RecipientTypeDetails = 'UserMailbox'; HiddenFromAddressListsEnabled = $false; Database = 'DB01'; EmailAddresses = @("SMTP:$sam@contoso.com"); GrantSendOnBehalfTo = @() }
    })
}
function global:Get-Module {
    [CmdletBinding()]
    param([string[]]$Name, [switch]$ListAvailable, [switch]$SkipEditionCheck)
    if ($ListAvailable) {
        foreach ($module in @(@{ Name = 'Microsoft.Graph.Authentication'; Version = '2.41.1' }, @{ Name = 'ExchangeOnlineManagement'; Version = '3.10.1' })) {
            if (-not $Name -or $Name -contains $module.Name) { [pscustomobject]@{ Name = $module.Name; Version = [version]$module.Version } }
        }
    } else { Microsoft.PowerShell.Core\Get-Module @PSBoundParameters }
}
__CALL__
exit $LASTEXITCODE
'@
        $literal = (@($Arguments | ForEach-Object { if ($_ -like '-*') { $_ } else { "'" + $_.Replace("'", "''") + "'" } }) -join ' ')
        $body = $body.Replace('__CALL__', ('& ''' + $script:Entry.Replace("'", "''") + "' " + $literal))
        [IO.File]::WriteAllText($runner, $body, (New-Object Text.UTF8Encoding($true)))
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        $env:PCM_TEST_MAILBOXES = [string]$Mailboxes
        $env:PCM_TEST_EXCHANGE_FAILS = $(if ($ExchangeFails) { '1' } else { '' })
        try { $output = & $script:Pwsh -NoProfile -NonInteractive -File $runner 2>&1 | Out-String }
        finally { $ErrorActionPreference = $previous; $env:PCM_TEST_MAILBOXES = ''; $env:PCM_TEST_EXCHANGE_FAILS = '' }
        [pscustomobject]@{ Code = $LASTEXITCODE; Text = $output }
    }

    function script:Get-LogText {
        param([Parameter(Mandatory)][string]$Folder)
        $files = @(Get-ChildItem -LiteralPath (Join-Path $Folder 'logs') -Filter '*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime)
        if (-not $files.Count) { return '' }
        [IO.File]::ReadAllText($files[-1].FullName)
    }

    function script:Get-CompleteSnapshot {
        param([Parameter(Mandatory)][string]$Path)
        $connection = Open-PcmStore -Path $Path -Root $script:Root -ReadOnly
        try { @(Invoke-PcmSql $connection "SELECT id, run_id, status FROM snapshot ORDER BY id" -As Rows) } finally { Close-PcmStore $connection }
    }

    function script:Get-DbDigest {
        param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][long]$Id)
        $connection = Open-PcmStore -Path $Path -Root $script:Root -ReadOnly
        try {
            Get-PcmSnapshotDigest -Mailbox @(Read-PcmTable $connection mailbox $Id) -Permission @(Read-PcmTable $connection permission $Id) -Member @(Read-PcmTable $connection group_member $Id)
        } finally { Close-PcmStore $connection }
    }

    function script:New-ConvertBatch {
        <# A Convert batch in the journal of a tool folder, with the snapshot reference that the test gives. #>
        param([Parameter(Mandatory)][string]$Folder, [long]$SnapshotId = 1, [string]$RunId = '', [string]$Hash = '', [string]$Step = 'GuidCleared', [string]$Status = 'Done')
        $journal = Open-PcmStore -Path (Join-Path $Folder 'db\journal.db') -Root $script:Root -Kind Journal
        try {
            $id = New-PcmBatch $journal Convert @{ SnapshotId = $SnapshotId; SnapshotRunId = $RunId; SnapshotHash = $Hash; Environment = 'TEST'; ToolVersion = 'test' }
            Set-PcmBatchItem $journal $id 'a0000000-0000-0000-0000-00000000000a' @{ kind = 'User'; identity = 'user01@contoso.com'; entra_id = 'e1'; step = $Step; status = $Status }
            $id
        } finally { Close-PcmStore $journal }
    }

    function script:New-SeededFolder {
        <# A tool folder in which a Collect of $Mailboxes mailboxes wrote snapshot 1. #>
        param([Parameter(Mandatory)][string]$Name, [int]$Mailboxes = 3, [string]$StoreExtra = '', [string]$CollectExtra = '')
        $folder = Join-Path $script:Work $Name
        $config = New-DataConfig -Folder $folder -StoreExtra $StoreExtra -CollectExtra $CollectExtra
        $run = Invoke-Entry -Folder $folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-ConfigPath', $config, '-Force') -Mailboxes $Mailboxes
        if ($run.Code -ne 0) { throw "The seed Collect failed: $($run.Text)" }
        [pscustomobject]@{ Folder = $folder; Config = $config; Db = (Join-Path $folder 'db\pra.db'); Hash = (Join-Path $folder 'db\pra.db.sha256'); Journal = (Join-Path $folder 'db\journal.db') }
    }
}

AfterAll {
    foreach ($path in @($script:Work, (Split-Path $script:Root -Parent))) { if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue } }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Journal schema 3: the snapshot of a batch and the index on the objects' {
    It 'keeps the run and the content digest of the snapshot in the batch, and finds them again' {
        $path = Join-Path $TestDrive 'schema3.db'
        $cn = Open-PcmStore -Path $path -Root $script:Root -Kind Journal
        try {
            $id = New-PcmBatch $cn Convert @{ SnapshotId = 7; SnapshotRunId = 'run-abc'; SnapshotHash = ('ab' * 32); Environment = 'TEST'; ToolVersion = 't' }
            $row = Get-PcmBatch $cn -Id $id
            $row.snapshot_id | Should -Be 7
            $row.snapshot_run_id | Should -BeExactly 'run-abc'
            $row.snapshot_hash | Should -BeExactly ('ab' * 32)
            (Invoke-PcmSql $cn "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar) | Should -Be '3'
        } finally { Close-PcmStore $cn }
    }
    It 'has an index on batch_item (object_guid)' {
        $cn = Open-PcmStore -Path (Join-Path $TestDrive 'index.db') -Root $script:Root -Kind Journal
        try {
            @(Invoke-PcmSql $cn "PRAGMA index_info(ix_batch_item_object)" -As Rows | ForEach-Object { $_.name }) | Should -Be @('object_guid')
            @(Invoke-PcmSql $cn "PRAGMA index_list(batch_item)" -As Rows | ForEach-Object { $_.name }) | Should -Contain 'ix_batch_item_object'
        } finally { Close-PcmStore $cn }
    }
    It 'reads a journal of schema 2 without loss, and gives it the new columns and the index when a command that writes opens it' {
        $path = Join-Path $TestDrive 'schema2.db'
        $cn = Open-PcmStore -Path $path -Root $script:Root -Kind Journal
        try {
            $id = New-PcmBatch $cn Convert @{ SnapshotId = 4; Environment = 'TEST'; ToolVersion = 't' }
            Set-PcmBatchItem $cn $id 'a0000000-0000-0000-0000-00000000000a' @{ kind = 'User'; identity = 'alice@contoso.com'; step = 'Done'; status = 'Done' }
            $null = Invoke-PcmSql $cn 'DROP INDEX ix_batch_item_object'
            $null = Invoke-PcmSql $cn 'ALTER TABLE batch DROP COLUMN snapshot_run_id'
            $null = Invoke-PcmSql $cn 'ALTER TABLE batch DROP COLUMN snapshot_hash'
            $null = Invoke-PcmSql $cn "UPDATE meta SET value = '2' WHERE key = 'schema_version'"
            $script:oldBatch = $id
        } finally { Close-PcmStore $cn }
        $read = Open-PcmStore -Path $path -Root $script:Root -Kind Journal -ReadOnly
        try {
            $old = Get-PcmBatch $read -Id $script:oldBatch
            $old.snapshot_id | Should -Be 4
            $old.PSObject.Properties['snapshot_run_id'] | Should -BeNullOrEmpty
            @(Get-PcmBatchItem $read $script:oldBatch).Count | Should -Be 1
        } finally { Close-PcmStore $read }
        $write = Open-PcmStore -Path $path -Root $script:Root -Kind Journal
        try {
            $migrated = Get-PcmBatch $write -Id $script:oldBatch
            $migrated.snapshot_id | Should -Be 4
            $migrated.snapshot_run_id | Should -BeNullOrEmpty
            @(Invoke-PcmSql $write 'PRAGMA table_info(batch)' -As Rows | ForEach-Object { $_.name }) | Should -Contain 'snapshot_hash'
            @(Invoke-PcmSql $write 'PRAGMA index_list(batch_item)' -As Rows | ForEach-Object { $_.name }) | Should -Contain 'ix_batch_item_object'
            @(Get-PcmBatchItem $write $script:oldBatch).Count | Should -Be 1
        } finally { Close-PcmStore $write }
    }
    It 'refuses a batch ID with a line end after it (the whole text is checked)' {
        $cn = Open-PcmStore -Path (Join-Path $TestDrive 'ids.db') -Root $script:Root -Kind Journal
        try {
            $id = New-PcmBatch $cn Convert @{ SnapshotId = 1 }
            (Get-PcmBatch $cn -Id $id).id | Should -BeExactly $id
            { Get-PcmBatch $cn -Id ($id + "`n") } | Should -Throw '*Invalid batch ID*'
        } finally { Close-PcmStore $cn }
    }
}

Describe 'The digest of a snapshot' {
    BeforeAll {
        function script:New-DigestSnapshot {
            param([string]$Path, [string]$Display = 'Alice', [string]$Trustee = 'CONTOSO\bob', [int]$Pad = 0)
            $cn = Open-PcmStore -Path $Path -Root $script:Root
            try {
                if ($Pad) { for ($i = 0; $i -lt $Pad; $i++) { $null = New-PcmSnapshot $cn @{ RunId = "pad$i" } } }
                $id = New-PcmSnapshot $cn @{ RunId = 'run-' + [guid]::NewGuid().ToString('N'); Environment = 'T'; ToolVersion = 't'; ExchangeServer = 'EX'; ScopeJson = '{}' }
                $mailbox = [ordered]@{ object_guid = 'a0000000-0000-0000-0000-00000000000a'; kind = 'User'; recipient_type_details = 'UserMailbox'; sam_account_name = 'alice'; user_principal_name = 'alice@contoso.com'
                    name = 'alice'; display_name = $Display; alias = 'alice'; distinguished_name = 'CN=alice'; organizational_unit = 'x'; primary_smtp_address = 'alice@contoso.com'; windows_email_address = 'alice@contoso.com'
                    legacy_exchange_dn = '/o=C/cn=alice'; exchange_guid = 'b0000000-0000-0000-0000-00000000000b'; archive_guid = $null; archive_status = $null; hidden_from_address_lists = $false
                    immutable_id = 'abc='; database = 'DB'; email_addresses_json = '[]'; custom_attributes_json = '{}'; grant_send_on_behalf_json = '[]'; item_count = $null; total_item_size_bytes = $null; resource_settings_json = $null }
                $null = Add-PcmRow $cn mailbox $id @($mailbox)
                $permission = [ordered]@{ mailbox_guid = 'a0000000-0000-0000-0000-00000000000a'; access_right = 'FullAccess'; trustee = $Trustee; trustee_guid = $null; trustee_kind = 'User'; trustee_upn = $null
                    trustee_smtp = $null; trustee_name = $null; trustee_dn = $null; auto_mapping = 1; resolved = 1; note = '' }
                $null = Add-PcmRow $cn permission $id @($permission)
                Set-PcmSnapshotStatus $cn $id Complete @{}
                [pscustomobject]@{ Id = $id; Digest = (Get-PcmSnapshotDigest -Mailbox @(Read-PcmTable $cn mailbox $id) -Permission @(Read-PcmTable $cn permission $id) -Member @(Read-PcmTable $cn group_member $id)) }
            } finally { Close-PcmStore $cn }
        }
    }
    It 'is 64 lower-case hexadecimal characters, the same for the same content whatever the number and the run of the snapshot' {
        $first = New-DigestSnapshot -Path (Join-Path $TestDrive 'd1.db')
        $second = New-DigestSnapshot -Path (Join-Path $TestDrive 'd2.db') -Pad 3
        $first.Digest | Should -Match '^[0-9a-f]{64}\z'
        $second.Id | Should -Not -Be $first.Id
        $second.Digest | Should -BeExactly $first.Digest
    }
    It 'changes when a mailbox changes or a permission changes' {
        $base = New-DigestSnapshot -Path (Join-Path $TestDrive 'd3.db')
        (New-DigestSnapshot -Path (Join-Path $TestDrive 'd4.db') -Display 'Alicia').Digest | Should -Not -BeExactly $base.Digest
        (New-DigestSnapshot -Path (Join-Path $TestDrive 'd5.db') -Trustee 'CONTOSO\carol').Digest | Should -Not -BeExactly $base.Digest
    }
    It 'is the digest of nothing for an empty snapshot, and does not fail' {
        Get-PcmSnapshotDigest | Should -Match '^[0-9a-f]{64}\z'
    }
}

Describe 'A snapshot that a batch still needs is never pruned' {
    BeforeAll {
        function script:New-FiveSnapshots {
            param([string]$Path)
            $cn = Open-PcmStore -Path $Path -Root $script:Root
            try {
                foreach ($number in 1..5) {
                    $id = New-PcmSnapshot $cn @{ RunId = "run-$number"; Environment = 'T'; ToolVersion = 't'; ExchangeServer = 'EX'; ScopeJson = '{}' }
                    Set-PcmSnapshotStatus $cn $id Complete @{}
                }
            } finally { Close-PcmStore $cn }
        }
        function script:New-InUseJournal {
            param([string]$Path, [string]$RunId = 'run-1', [long]$SnapshotId = 1)
            $cn = Open-PcmStore -Path $Path -Root $script:Root -Kind Journal
            try {
                $batch = New-PcmBatch $cn Convert @{ SnapshotId = $SnapshotId; SnapshotRunId = $RunId; Environment = 'T'; ToolVersion = 't' }
                Set-PcmBatchItem $cn $batch 'a0000000-0000-0000-0000-000000000001' @{ kind = 'User'; identity = 'a'; step = 'Done'; status = 'Done' }
                Set-PcmBatchItem $cn $batch 'a0000000-0000-0000-0000-000000000002' @{ kind = 'User'; identity = 'b'; step = 'Done'; status = 'Done' }
                Set-PcmBatch $cn $batch Partial @{}
                $batch
            } finally { Close-PcmStore $cn }
        }
        function script:Add-RollBack {
            param([string]$Path, [string]$ConvertBatch, [string[]]$Objects)
            $cn = Open-PcmStore -Path $Path -Root $script:Root -Kind Journal
            try {
                $recover = New-PcmBatch $cn Recover @{ ConvertBatch = $ConvertBatch; SnapshotId = 1; Environment = 'T'; ToolVersion = 't' }
                foreach ($guid in $Objects) { Set-PcmBatchItem $cn $recover $guid @{ kind = 'User'; identity = 'x'; step = 'Done'; status = 'Done' } }
                Set-PcmBatch $cn $recover Complete @{}
            } finally { Close-PcmStore $cn }
        }
    }
    It 'lists no snapshot in use when there is no journal, and keeps the old behaviour' {
        $inUse = Get-PcmSnapshotInUse -JournalPath (Join-Path $TestDrive 'none.db') -Root $script:Root
        $inUse.Readable | Should -BeTrue
        @($inUse.Items).Count | Should -Be 0
        $db = Join-Path $TestDrive 'plain.db'
        New-FiveSnapshots $db
        $cn = Open-PcmStore -Path $db -Root $script:Root
        try {
            Remove-PcmOldSnapshot $cn 2 -Protect @($inUse.Items) | Should -Be 3
            @(Get-PcmSnapshotList $cn | ForEach-Object { $_.id }) | Should -Be @(5, 4)
        } finally { Close-PcmStore $cn }
    }
    It 'keeps the snapshot of a Convert batch that is not rolled back, and says which one it kept' {
        $db = Join-Path $TestDrive 'keep.db'; $journal = Join-Path $TestDrive 'keep-journal.db'
        New-FiveSnapshots $db
        $batch = New-InUseJournal -Path $journal
        $inUse = Get-PcmSnapshotInUse -JournalPath $journal -Root $script:Root
        $inUse.Readable | Should -BeTrue
        @($inUse.Items).Count | Should -Be 1
        $inUse.Items[0].Batch | Should -BeExactly $batch
        $cn = Open-PcmStore -Path $db -Root $script:Root
        try {
            $kept = @()
            Remove-PcmOldSnapshot $cn 2 -Protect $inUse.Items -Kept ([ref]$kept) | Should -Be 2
            @($kept) | Should -Be @(1)
            @(Get-PcmSnapshotList $cn | ForEach-Object { $_.id } | Sort-Object) | Should -Be @(1, 4, 5)
            (Get-PcmSnapshot $cn).id | Should -Be 5
        } finally { Close-PcmStore $cn }
    }
    It 'deletes it as soon as every converted object is rolled back, and not before (a wave that rolled back one object of two keeps it)' {
        $journal = Join-Path $TestDrive 'wave-journal.db'
        $batch = New-InUseJournal -Path $journal
        Add-RollBack -Path $journal -ConvertBatch $batch -Objects @('a0000000-0000-0000-0000-000000000001')
        @((Get-PcmSnapshotInUse -JournalPath $journal -Root $script:Root).Items).Count | Should -Be 1
        Add-RollBack -Path $journal -ConvertBatch $batch -Objects @('A0000000-0000-0000-0000-000000000002')
        @((Get-PcmSnapshotInUse -JournalPath $journal -Root $script:Root).Items).Count | Should -Be 0
    }
    It 'finds the snapshot by its run when the numbers differ, and by its number for a batch without a run' {
        $db = Join-Path $TestDrive 'byrun.db'
        New-FiveSnapshots $db
        $cn = Open-PcmStore -Path $db -Root $script:Root
        try {
            $kept = @()
            Remove-PcmOldSnapshot $cn 1 -Protect @([pscustomobject]@{ SnapshotId = 99; RunId = 'RUN-2' }) -Kept ([ref]$kept) | Should -Be 3
            @($kept) | Should -Be @(2)
        } finally { Close-PcmStore $cn }
        $cn = Open-PcmStore -Path (Join-Path $TestDrive 'bynumber.db') -Root $script:Root
        try {
            foreach ($number in 1..4) { $id = New-PcmSnapshot $cn @{ RunId = "r$number" }; Set-PcmSnapshotStatus $cn $id Complete @{} }
            $kept = @()
            Remove-PcmOldSnapshot $cn 1 -Protect @([pscustomobject]@{ SnapshotId = 3; RunId = '' }) -Kept ([ref]$kept) | Should -Be 2
            @($kept) | Should -Be @(3)
        } finally { Close-PcmStore $cn }
    }
    It 'counts a Convert batch that is still running as not rolled back, even before it has an object' {
        $journal = Join-Path $TestDrive 'running-journal.db'
        $cn = Open-PcmStore -Path $journal -Root $script:Root -Kind Journal
        try { $null = New-PcmBatch $cn Convert @{ SnapshotId = 2; SnapshotRunId = 'run-2'; Environment = 'T'; ToolVersion = 't' } } finally { Close-PcmStore $cn }
        @((Get-PcmSnapshotInUse -JournalPath $journal -Root $script:Root).Items).Count | Should -Be 1
    }
    It 'reads a journal of schema 2 (no run recorded): the number protects the snapshot' {
        $journal = Join-Path $TestDrive 'old-journal.db'
        $null = New-InUseJournal -Path $journal -RunId '' -SnapshotId 3
        $cn = Open-PcmStore -Path $journal -Root $script:Root -Kind Journal
        try {
            $null = Invoke-PcmSql $cn 'ALTER TABLE batch DROP COLUMN snapshot_run_id'
            $null = Invoke-PcmSql $cn 'ALTER TABLE batch DROP COLUMN snapshot_hash'
            $null = Invoke-PcmSql $cn "UPDATE meta SET value = '2' WHERE key = 'schema_version'"
        } finally { Close-PcmStore $cn }
        $inUse = Get-PcmSnapshotInUse -JournalPath $journal -Root $script:Root
        $inUse.Readable | Should -BeTrue
        $inUse.Items[0].SnapshotId | Should -Be 3
        $inUse.Items[0].RunId | Should -BeExactly ''
    }
    It 'says the journal cannot be read when the file is not a journal, so that the caller deletes nothing' {
        $journal = Join-Path $TestDrive 'broken-journal.db'
        [IO.File]::WriteAllBytes($journal, [byte[]](1..200))
        $inUse = Get-PcmSnapshotInUse -JournalPath $journal -Root $script:Root
        $inUse.Readable | Should -BeFalse
        $inUse.Message | Should -Not -BeNullOrEmpty
    }
}

Describe 'Snapshot ID of a group: letter case of non-ASCII letters' {
    BeforeAll {
        $script:groupDb = Join-Path $TestDrive 'groups.db'
        $cn = Open-PcmStore -Path $script:groupDb -Root $script:Root
        try {
            $script:groupSnapshot = New-PcmSnapshot $cn @{ RunId = 'g'; Environment = 'T'; ToolVersion = 't'; ExchangeServer = 'EX'; ScopeJson = '{}' }
            $null = Add-PcmRow $cn target_group $script:groupSnapshot @(
                [ordered]@{ group_guid = '11111111-aaaa-4aaa-8aaa-000000000001'; name = 'Équipe-Paris'; sam_account_name = 'equipe-paris'; distinguished_name = 'CN=Équipe-Paris,OU=Groupes,DC=contoso,DC=com'
                    primary_smtp_address = 'equipe.paris@contoso.com'; display_name = 'Équipe Paris'; member_count = 2 })
            Set-PcmSnapshotStatus $cn $script:groupSnapshot Complete @{}
        } finally { Close-PcmStore $cn }
    }
    It 'finds the group named <Written> (<Value>)' -TestCases @(
        @{ Written = 'in lower case'; Value = 'équipe-paris' }, @{ Written = 'in capitals'; Value = 'ÉQUIPE-PARIS' }, @{ Written = 'as the display name in lower case'; Value = 'équipe paris' },
        @{ Written = 'by its distinguished name in lower case'; Value = 'cn=équipe-paris,ou=groupes,dc=contoso,dc=com' }, @{ Written = 'with a decomposed accent'; Value = "E$([char]0x0301)quipe-Paris" },
        @{ Written = 'by its account name in capitals'; Value = 'EQUIPE-PARIS' }, @{ Written = 'by its address'; Value = 'EQUIPE.PARIS@CONTOSO.COM' },
        @{ Written = 'by its GUID in capitals'; Value = '11111111-AAAA-4AAA-8AAA-000000000001' }) {
        param($Written, $Value)
        $cn = Open-PcmStore -Path $script:groupDb -Root $script:Root -ReadOnly
        try {
            $found = Find-PcmTargetGroup $cn $script:groupSnapshot $Value
            $found.HasTables | Should -BeTrue
            @($found.Groups).Count | Should -Be 1
        } finally { Close-PcmStore $cn }
    }
    It 'does not find a group that is not there' {
        $cn = Open-PcmStore -Path $script:groupDb -Root $script:Root -ReadOnly
        try { @((Find-PcmTargetGroup $cn $script:groupSnapshot 'Équipe-Lyon').Groups).Count | Should -Be 0 } finally { Close-PcmStore $cn }
    }
}

Describe 'Lists written by a person' {
    It 'finds the delimiter of a header with the column Identity anywhere, in any letter case, quoted or not' -TestCases @(
        @{ Header = 'Identity'; Expected = ',' }, @{ Header = 'Name,Identity,Note'; Expected = ',' }, @{ Header = 'Name;Identity'; Expected = ';' }, @{ Header = '"IDENTITY";Name'; Expected = ';' },
        @{ Header = 'Name, "identity" ,x'; Expected = ',' }, @{ Header = 'Name,Note'; Expected = '' }, @{ Header = 'Identity2,Name'; Expected = '' }, @{ Header = ''; Expected = '' }) {
        param($Header, $Expected)
        Get-PcmCsvDelimiter $Header | Should -BeExactly $Expected
    }
    It 'reads a UTF-8 file with or without a byte order mark, and UTF-16' {
        $text = "Identity`r`nÉlodie@contoso.com`r`n"
        foreach ($case in @(@('utf8bom', (New-Object Text.UTF8Encoding($true))), @('utf8', (New-Object Text.UTF8Encoding($false))), @('utf16', [Text.Encoding]::Unicode))) {
            $path = Join-Path $TestDrive ('enc-' + $case[0] + '.csv')
            [IO.File]::WriteAllText($path, $text, $case[1])
            Read-PcmCsvText -Path $path | Should -BeExactly $text
        }
    }
    It 'reads a file that is not UTF-8 (Excel, Notepad ANSI) with the ANSI code page, so that the accents are kept' -Skip:(-not $script:AnsiUsable) {
        $text = "Identity`r`nÉlodie@contoso.com`r`n"
        $path = Join-Path $TestDrive 'ansi.csv'
        [IO.File]::WriteAllBytes($path, [Text.Encoding]::GetEncoding($script:AnsiPage).GetBytes($text))
        Read-PcmCsvText -Path $path | Should -BeExactly $text
    }
    It 'reads a -CsvPath list whose Identity column is not the first one, with semicolons (any letter case)' {
        $path = Join-Path $TestDrive 'wave-anywhere.csv'
        Set-Content -LiteralPath $path -Value "Name;IDENTITY;Note`r`nAlice;alice@contoso.com;a`r`nBob;Bob@Contoso.com;b`r`n" -Encoding UTF8
        Read-PcmIdentityFile -Path $path | Should -Be @('alice@contoso.com', 'Bob@Contoso.com')
    }
    It 'reads a -CsvPath list saved as ANSI' -Skip:(-not $script:AnsiUsable) {
        $path = Join-Path $TestDrive 'wave-ansi.csv'
        [IO.File]::WriteAllBytes($path, [Text.Encoding]::GetEncoding($script:AnsiPage).GetBytes("Identity;Name`r`nÉlodie@contoso.com;Élodie`r`n"))
        Read-PcmIdentityFile -Path $path | Should -Be @('Élodie@contoso.com')
    }
}

Describe 'Collect reads the objects of a Csv scope and of a group' {
    BeforeAll {
        function script:New-Config {
            @{ Environment = 'TEST'; Exchange = @{ Server = 'ex1.contoso.com'; DomainController = '' }
                Scope = @{ Mode = 'Csv'; SearchBase = ''; GroupDN = ''; CsvPath = ''; IncludeUsers = $true; IncludeShared = $true; IncludeRoom = $false; IncludeEquipment = $false; ExcludeSamAccountNames = @() }
                Collect = @{ SharedPermissions = $true; ExpandGroupTrustees = $true; ExcludeTrustees = @(); TargetGroups = @() } }
        }
        function script:New-Context {
            param([hashtable]$Config)
            @{ Config = $Config; Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]')
                Excluded = (New-Object 'Collections.Generic.List[string]'); LogFile = ''; Warnings = 0; Server = '' }
        }
        function global:Get-Mailbox { param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController) }
        function global:Get-Group { param($Identity, $DomainController) }
        function script:New-FakeMailbox {
            param([string]$Sam, [string]$Type = 'UserMailbox')
            [pscustomobject]@{ Guid = [guid]::NewGuid(); Name = $Sam; SamAccountName = $Sam; DistinguishedName = "CN=$Sam,OU=PRA,DC=contoso,DC=com"; RecipientTypeDetails = $Type; PrimarySmtpAddress = "$Sam@contoso.com" }
        }
    }
    AfterAll { foreach ($name in 'Get-Mailbox', 'Get-Group') { Remove-Item -LiteralPath "Function:\$name" -Force -ErrorAction SilentlyContinue } }

    It 'takes the identities of a CSV file with semicolons and the Identity column in the middle' {
        $path = Join-Path $TestDrive 'scope.csv'
        Set-Content -LiteralPath $path -Value "Name;Identity;Note`r`nAlice;alice@contoso.com;a`r`nBob;bob@contoso.com;b`r`n" -Encoding UTF8
        Mock Get-Mailbox -ModuleName PRA.Collect { New-FakeMailbox -Sam (([string]$Identity -split '@')[0]) }
        $config = New-Config; $config.Scope.CsvPath = $path
        @(Get-PcmScopeMailbox -Context (New-Context $config) | ForEach-Object { [string]$_.SamAccountName }) | Should -Be @('alice', 'bob')
    }
    It 'takes the identities of a CSV file saved as ANSI' -Skip:(-not $script:AnsiUsable) {
        $path = Join-Path $TestDrive 'scope-ansi.csv'
        [IO.File]::WriteAllBytes($path, [Text.Encoding]::GetEncoding($script:AnsiPage).GetBytes("Identity;Name`r`nélodie@contoso.com;Élodie`r`n"))
        $script:asked = New-Object 'Collections.Generic.List[string]'
        Mock Get-Mailbox -ModuleName PRA.Collect { $script:asked.Add([string]$Identity); New-FakeMailbox -Sam 'elodie' }
        $config = New-Config; $config.Scope.CsvPath = $path
        $null = Get-PcmScopeMailbox -Context (New-Context $config)
        $script:asked[0] | Should -BeExactly 'élodie@contoso.com'
    }
    It 'says in plain words that the Identity column is missing, or that the file is empty' {
        $config = New-Config
        $path = Join-Path $TestDrive 'scope-bad.csv'
        Set-Content -LiteralPath $path -Value "Name,Note`r`nAlice,a`r`n" -Encoding UTF8
        $config.Scope.CsvPath = $path
        { Get-PcmScopeMailbox -Context (New-Context $config) } | Should -Throw '*column Identity missing*'
        $empty = Join-Path $TestDrive 'scope-empty.csv'
        Set-Content -LiteralPath $empty -Value '' -Encoding UTF8
        $config.Scope.CsvPath = $empty
        { Get-PcmScopeMailbox -Context (New-Context $config) } | Should -Throw '*the file is empty*'
    }
    It 'stops, and does not take a member for a member without mailbox, when Exchange fails to read it (Group scope)' {
        $config = New-Config; $config.Scope.Mode = 'Group'; $config.Scope.GroupDN = 'CN=G1'
        Mock Get-Group -ModuleName PRA.Collect { if ($Identity -eq 'CN=G1') { [pscustomobject]@{ Members = @('CN=u1', 'CN=u2') } } else { throw "The operation couldn't be performed because object '$Identity' couldn't be found on 'dc01.contoso.com'." } }
        Mock Get-Mailbox -ModuleName PRA.Collect { if ($Identity -eq 'CN=u1') { throw 'The WS-Management service cannot process the request: the server did not answer.' } else { New-FakeMailbox -Sam 'u2' } }
        { Get-PcmScopeMailbox -Context (New-Context $config) } | Should -Throw '*the member CN=u1 of the group CN=G1 could not be read*'
    }
    It 'still reports a member that Exchange says does not exist as a member without mailbox (Group scope)' {
        $config = New-Config; $config.Scope.Mode = 'Group'; $config.Scope.GroupDN = 'CN=G1'
        Mock Get-Group -ModuleName PRA.Collect { if ($Identity -eq 'CN=G1') { [pscustomobject]@{ Members = @('CN=u1', 'CN=gone') } } else { throw "The operation couldn't be performed because object '$Identity' couldn't be found on 'dc01.contoso.com'." } }
        Mock Get-Mailbox -ModuleName PRA.Collect { if ($Identity -eq 'CN=u1') { New-FakeMailbox -Sam 'u1' } else { throw "The operation couldn't be performed because object '$Identity' couldn't be found on 'dc01.contoso.com'." } }
        Mock Get-PcmDirectoryObject -ModuleName PRA.Collect { $null }
        $context = New-Context $config
        @(Get-PcmScopeMailbox -Context $context | ForEach-Object { [string]$_.SamAccountName }) | Should -Be @('u1')
        @($context.Excluded) | Should -Contain 'CN=gone: not a mailbox'
    }
    It 'stops when a nested member cannot be read to tell whether it is a group, and walks the others (<Case>)' -TestCases @(
        @{ Case = 'a member that is not a group'; Fault = "The operation couldn't be performed because object 'CN=u1' couldn't be found on 'dc01.contoso.com'."; Throws = $false }
        @{ Case = 'a read that fails'; Fault = 'The WS-Management service cannot process the request: the server did not answer.'; Throws = $true }) {
        param($Case, $Fault, $Throws)
        $script:fault = $Fault
        Mock Get-Group -ModuleName PRA.Collect {
            if ($Identity -eq 'CN=G1') { [pscustomobject]@{ Members = @('CN=u1') } } else { throw $script:fault }
        }
        $context = New-Context (New-Config)
        if ($Throws) { { Get-PcmGroupMemberDn -Context $context -Identity 'CN=G1' } | Should -Throw '*could not be read to tell whether it is a group*' }
        else { @(Get-PcmGroupMemberDn -Context $context -Identity 'CN=G1') | Should -Be @('CN=u1') }
    }
    It 'tells an object that does not exist from a read that failed' -TestCases @(
        @{ Message = "The operation couldn't be performed because object 'x' couldn't be found on 'dc'."; Expected = $true }
        @{ Message = 'Group x not found'; Expected = $true }
        @{ Message = 'There is no such object on the server.'; Expected = $true }
        @{ Message = 'The server is not operational.'; Expected = $false }
        @{ Message = 'The network path was not found.'; Expected = $false }
        @{ Message = 'The WS-Management service cannot process the request.'; Expected = $false }) {
        param($Message, $Expected)
        $record = try { throw $Message } catch { $_ }
        Test-PcmNotFoundError -ErrorRecord $record | Should -Be $Expected
    }
}

Describe 'Collect: the groups of the snapshot are read to the end' {
    BeforeAll {
        function script:New-Config {
            @{ Environment = 'TEST'; Exchange = @{ Server = 'ex1.contoso.com'; DomainController = '' }
                Scope = @{ Mode = 'Auto'; SearchBase = ''; GroupDN = ''; CsvPath = ''; IncludeUsers = $true; IncludeShared = $true; IncludeRoom = $false; IncludeEquipment = $false; ExcludeSamAccountNames = @() }
                Collect = @{ SharedPermissions = $true; ExpandGroupTrustees = $true; ExcludeTrustees = @(); TargetGroups = @('CN=G1') } }
        }
        function script:New-Context {
            param([hashtable]$Config)
            @{ Config = $Config; Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]')
                Excluded = (New-Object 'Collections.Generic.List[string]'); LogFile = ''; Warnings = 0; Server = '' }
        }
        function global:Get-Group { param($Identity, $DomainController) }
        function global:Get-Recipient { param($Identity, $DomainController) }
        function global:Get-MailboxPermission { param($Identity, $DomainController) }
        function global:Get-ADPermission { param($Identity, $DomainController) }
    }
    BeforeEach {
        Mock Get-Group -ModuleName PRA.Collect { [pscustomobject]@{ Name = 'G1'; SamAccountName = 'g1'; DistinguishedName = 'CN=G1'; DisplayName = 'G1'; Guid = '11111111-aaaa-4aaa-8aaa-000000000001'; Members = @('CN=u1', 'CN=ghost') } }
        Mock Get-Recipient -ModuleName PRA.Collect { throw 'not found' }
    }
    AfterAll { foreach ($name in 'Get-Group', 'Get-Recipient', 'Get-MailboxPermission', 'Get-ADPermission') { Remove-Item -LiteralPath "Function:\$name" -Force -ErrorAction SilentlyContinue } }

    It 'counts a member that does not exist any more as a warning and keeps it as Unknown, so that a Complete snapshot does not hide it' {
        Mock Get-PcmDirectoryObject -ModuleName PRA.Collect {
            if ($DistinguishedName -eq 'CN=u1') { [pscustomobject]@{ Class = 'User'; Name = 'u1'; Sam = ''; Upn = 'u1@contoso.com'; Smtp = 'u1@contoso.com'; DisplayName = ''; Dn = 'CN=u1'; Guid = 'aaaaaaaa-0000-4000-8000-000000000001' } }
            else { $null }
        }
        $result = Get-PcmTargetGroup -Context (New-Context (New-Config))
        @($result.Groups).Count | Should -Be 1
        @($result.Members | Where-Object { $_.member_class -eq 'Unknown' }).Count | Should -Be 1
        @($result.Warnings).Count | Should -Be 1
        $result.Warnings[0] | Should -BeLike '*1 member(s) do not exist in Active Directory any more and are stored as Unknown*CN=ghost*'
    }
    It 'does not capture the group when a member cannot be read for another reason than its absence, and says so' {
        Mock Get-PcmDirectoryObject -ModuleName PRA.Collect { throw 'The server is not operational.' }
        $result = Get-PcmTargetGroup -Context (New-Context (New-Config))
        @($result.Groups).Count | Should -Be 0
        @($result.Members).Count | Should -Be 0
        $result.Warnings[0] | Should -Match 'its members could not all be read \(The server is not operational'
    }
    It 'reads the object strictly: absent is $null, a failed read is an error, and the loose read stays $null (<Case>)' -TestCases @(
        @{ Case = 'does not exist'; Message = 'There is no such object on the server.'; Strict = $true; Throws = $false }
        @{ Case = 'fails, strict'; Message = 'The server is not operational.'; Strict = $true; Throws = $true }
        @{ Case = 'fails, loose'; Message = 'The server is not operational.'; Strict = $false; Throws = $false }) {
        param($Case, $Message, $Strict, $Throws)
        $script:adsiFault = $Message
        Mock New-Object -ModuleName PRA.Collect -ParameterFilter { $TypeName -eq 'System.DirectoryServices.DirectoryEntry' } { throw $script:adsiFault }
        $context = New-Context (New-Config)
        if ($Throws) { { Get-PcmDirectoryObject -Context $context -DistinguishedName 'CN=x' -Strict } | Should -Throw '*not operational*' }
        elseif ($Strict) { Get-PcmDirectoryObject -Context $context -DistinguishedName 'CN=x' -Strict | Should -BeNullOrEmpty }
        else { Get-PcmDirectoryObject -Context $context -DistinguishedName 'CN=x' | Should -BeNullOrEmpty }
    }
    It 'warns, in the shared mailbox permissions, about a member of a group trustee that cannot be read' {
        $shared = [pscustomobject]@{ Guid = 's1'; PrimarySmtpAddress = 'compta@contoso.com'; DistinguishedName = 'CN=compta,OU=PRA,DC=contoso,DC=com'; GrantSendOnBehalfTo = @() }
        Mock Get-MailboxPermission -ModuleName PRA.Collect { @([pscustomobject]@{ User = 'CONTOSO\Team'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }) }
        Mock Get-ADPermission -ModuleName PRA.Collect { @() }
        Mock Get-PcmDelegateListLink -ModuleName PRA.Collect { @() }
        Mock Get-PcmGroupMemberDn -ModuleName PRA.Collect { @('CN=ghost,OU=PRA,DC=contoso,DC=com', 'CN=u1,OU=PRA,DC=contoso,DC=com') }
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect {
            switch -Wildcard ($Trustee) {
                'CONTOSO\Team' { [pscustomobject]@{ Resolved = $true; Guid = 'g1'; Kind = 'Group'; Upn = $null; Smtp = 'team@contoso.com'; Name = 'Team'; Dn = 'CN=Team'; Note = '' } }
                'CN=u1*' { [pscustomobject]@{ Resolved = $true; Guid = 'u1'; Kind = 'User'; Upn = 'u1@contoso.com'; Smtp = 'u1@contoso.com'; Name = 'u1'; Dn = $Trustee; Note = '' } }
                default { [pscustomobject]@{ Resolved = $false; Guid = $null; Kind = 'Unknown'; Upn = $null; Smtp = $null; Name = $null; Dn = $null; Note = 'not found (deleted account or orphan SID)' } }
            }
        }
        $result = Get-PcmSharedPermission -Context (New-Context (New-Config)) -Mailbox $shared
        @($result.Members).Count | Should -Be 2
        @($result.Warnings | Where-Object { $_ -like '*CN=ghost*of group Team*stored as Unknown*' }).Count | Should -Be 1
    }
}

Describe 'Collect: the path of an object in Active Directory' {
    It 'escapes the "/" of a distinguished name, names the domain controller when there is one' -TestCases @(
        @{ Dn = 'CN=Sales/Marketing,OU=PRA,DC=contoso,DC=com'; Dc = ''; Expected = 'LDAP://CN=Sales\/Marketing,OU=PRA,DC=contoso,DC=com' }
        @{ Dn = 'CN=Sales/Marketing,OU=PRA,DC=contoso,DC=com'; Dc = 'dc01.contoso.com'; Expected = 'LDAP://dc01.contoso.com/CN=Sales\/Marketing,OU=PRA,DC=contoso,DC=com' }
        @{ Dn = 'CN=alice,OU=PRA,DC=contoso,DC=com'; Dc = ''; Expected = 'LDAP://CN=alice,OU=PRA,DC=contoso,DC=com' }) {
        param($Dn, $Dc, $Expected)
        Get-PcmLdapPath -DistinguishedName $Dn -DomainController $Dc | Should -BeExactly $Expected
    }
    It 'reads the AutoMapping of a mailbox whose name holds a "/" through the escaped path' {
        $script:paths = New-Object 'Collections.Generic.List[string]'
        Mock New-Object -ModuleName PRA.Collect -ParameterFilter { $TypeName -eq 'System.DirectoryServices.DirectoryEntry' } {
            $script:paths.Add([string]$ArgumentList[0])
            $entry = [pscustomobject]@{ Properties = @{ msExchDelegateListLink = @('CN=bob,OU=PRA,DC=contoso,DC=com') } }
            Add-Member -InputObject $entry -MemberType ScriptMethod -Name Dispose -Value { } -PassThru
        }
        $context = @{ Config = @{ Exchange = @{ DomainController = '' } } }
        @(Get-PcmDelegateListLink -Context $context -DistinguishedName 'CN=Sales/Marketing,OU=PRA,DC=contoso,DC=com') | Should -Be @('CN=bob,OU=PRA,DC=contoso,DC=com')
        $script:paths[0] | Should -BeExactly 'LDAP://CN=Sales\/Marketing,OU=PRA,DC=contoso,DC=com'
    }
}

Describe 'Collect: the forest is read as a whole' {
    BeforeAll {
        function script:New-Config { @{ Exchange = @{ Server = 'ex1.contoso.com'; DomainController = '' } } }
        function script:New-Context { @{ Config = (New-Config); Issues = (New-Object 'Collections.Generic.List[object]'); Excluded = (New-Object 'Collections.Generic.List[string]'); Server = '' } }
        function global:New-PSSession { [CmdletBinding()] param($ConfigurationName, $ConnectionUri, $Authentication) [pscustomobject]@{ Id = 9 } }
        function global:Import-PSSession { [CmdletBinding()] param($Session, $CommandName, [switch]$DisableNameChecking, [switch]$AllowClobber) New-Module -Name PcmDataFakeExchange -ScriptBlock { } }
        function global:Remove-PSSession { [CmdletBinding()] param($Session) $script:removed = $true }
        function global:Get-ExchangeServer { param() }
    }
    AfterAll {
        foreach ($name in 'New-PSSession', 'Import-PSSession', 'Remove-PSSession', 'Get-ExchangeServer', 'Set-ADServerSettings') { Remove-Item -LiteralPath "Function:\$name" -Force -ErrorAction SilentlyContinue }
        Remove-Module PcmDataFakeExchange -Force -ErrorAction SilentlyContinue
    }
    BeforeEach { $script:removed = $false; Remove-Item -LiteralPath 'Function:\Set-ADServerSettings' -Force -ErrorAction SilentlyContinue }

    It 'stops, closes the session and says what to check when Exchange refuses ViewEntireForest' {
        function global:Set-ADServerSettings { param($ViewEntireForest) throw 'Access to the cmdlet is denied.' }
        $context = New-Context
        $message = try { $null = Connect-PcmExchange -Context $context; '' } catch { $_.Exception.Message }
        $message | Should -BeLike 'Collect cannot read the whole forest on ex1.contoso.com: Access to the cmdlet is denied.*View-Only Organization Management*Nothing was written.'
        $script:removed | Should -BeTrue
        $context['ExchangeSession'] | Should -BeNullOrEmpty
    }
    It 'stops when the command is not available to the account' {
        $message = try { $null = Connect-PcmExchange -Context (New-Context); '' } catch { $_.Exception.Message }
        $message | Should -BeLike 'Collect cannot read the whole forest on ex1.contoso.com: the command Set-ADServerSettings is not available to this account*'
        $script:removed | Should -BeTrue
    }
    It 'goes on when ViewEntireForest is set' {
        function global:Set-ADServerSettings { param($ViewEntireForest) $script:forest = $ViewEntireForest }
        (Connect-PcmExchange -Context (New-Context)) | Should -BeLike 'remote session to ex1.contoso.com (Kerberos)*'
        $script:forest | Should -BeTrue
    }
}

Describe 'Collect refuses a snapshot that lost its mailboxes' {
    BeforeAll {
        $script:size = New-SeededFolder -Name 'size' -Mailboxes 10
        function script:Get-SizeRun {
            param([int]$Mailboxes, [string]$Folder = $script:size.Folder, [string]$Config = $script:size.Config, [switch]$Preview)
            $arguments = @('-Action', 'Collect', '-ConfigPath', $Config)
            if (-not $Preview) { $arguments += @('-Mode', 'Apply', '-Force') }
            Invoke-Entry -Folder $Folder -Arguments $arguments -Mailboxes $Mailboxes
        }
    }
    It 'has written snapshot 1 with 10 mailboxes' {
        @(Get-CompleteSnapshot $script:size.Db).Count | Should -Be 1
    }
    It 'fails, writes nothing and keeps the last complete snapshot and its hash file when the mailboxes drop by more than 50 percent' {
        $before = (Read-PcmHashFile -Path $script:size.Hash).Hash
        $run = Get-SizeRun -Mailboxes 4
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'RESULT: FAIL|Collect stopped'
        $run.Text | Should -Match 'Collect found 4 mailbox\(es\), 60% fewer than the last complete snapshot of the same scope \(snapshot 1, 10 mailbox\(es\)\)'
        $run.Text | Should -Match 'Collect\.MaxDropPercent \(50%\)'
        $run.Text | Should -Match 'Nothing was changed on Exchange'
        @(Get-CompleteSnapshot $script:size.Db).Count | Should -Be 1
        (Read-PcmHashFile -Path $script:size.Hash).Hash | Should -BeExactly $before
        (Test-PcmSnapshotHash -Path $script:size.Db).State | Should -Be 'Verified'
        @(Get-ChildItem -LiteralPath (Join-Path $script:size.Folder 'backup') -Filter '*.db').Count | Should -Be 1
    }
    It 'fails when Exchange returns no mailbox at all, and says what to check' {
        $run = Get-SizeRun -Mailboxes 0
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'Collect found no mailbox in the scope'
        $run.Text | Should -Match 'The last complete snapshot is snapshot 1, with 10 mailbox\(es\)'
        $run.Text | Should -Match 'set Collect\.MaxDropPercent to 100'
        @(Get-CompleteSnapshot $script:size.Db).Count | Should -Be 1
    }
    It 'shows the same refusal as a warning in a Preview, and exits 0' {
        $run = Get-SizeRun -Mailboxes 0 -Preview
        $run.Code | Should -Be 0 -Because $run.Text
        (Get-LogText -Folder $script:size.Folder) | Should -Match 'In Apply mode this Collect would be refused: Collect found no mailbox'
    }
    It 'writes the snapshot when the drop is not above the limit, and the hash file matches it' {
        $run = Get-SizeRun -Mailboxes 6
        $run.Code | Should -Be 0 -Because $run.Text
        @(Get-CompleteSnapshot $script:size.Db).Count | Should -Be 2
        (Test-PcmSnapshotHash -Path $script:size.Db).State | Should -Be 'Verified'
    }
    It 'does not compare a snapshot of one object with a full one' {
        $run = Invoke-Entry -Folder $script:size.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-Identity', 'user01@contoso.com', '-ConfigPath', $script:size.Config) -Mailboxes 1
        $run.Code | Should -Be 0 -Because $run.Text
        @(Get-CompleteSnapshot $script:size.Db).Count | Should -Be 3
    }
    It 'accepts any size with Collect.MaxDropPercent = 100 (also an empty snapshot)' {
        $folder = Join-Path $script:Work 'size100'
        $config = New-DataConfig -Folder $folder -CollectExtra 'MaxDropPercent = 100'
        (Invoke-Entry -Folder $folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $config) -Mailboxes 10).Code | Should -Be 0
        $run = Invoke-Entry -Folder $folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $config) -Mailboxes 0
        $run.Code | Should -Be 0 -Because $run.Text
        @(Get-CompleteSnapshot (Join-Path $folder 'db\pra.db')).Count | Should -Be 2
    }
    It 'takes the limit in the configuration, any letter case in the key' {
        $folder = Join-Path $script:Work 'size20'
        $config = New-DataConfig -Folder $folder -CollectExtra 'maxdroppercent = 20'
        (Invoke-Entry -Folder $folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $config) -Mailboxes 10).Code | Should -Be 0
        $run = Invoke-Entry -Folder $folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $config) -Mailboxes 7
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'Collect\.MaxDropPercent \(20%\)'
    }
    It 'refuses a limit above 100 in the configuration' {
        $folder = Join-Path $script:Work 'size101'
        $config = New-DataConfig -Folder $folder -CollectExtra 'MaxDropPercent = 101'
        $run = Invoke-Entry -Folder $folder -Arguments @('-Action', 'Collect', '-ConfigPath', $config) -Mailboxes 3
        $run.Code | Should -Be 1
        $run.Text | Should -Match 'Collect\.MaxDropPercent must be a whole number between 0 and 100'
    }
}

Describe 'A failed Collect never leaves the good snapshot without its hash file' {
    It 'keeps the hash file when Exchange fails before anything is written' {
        $seed = New-SeededFolder -Name 'failexchange'
        $before = (Read-PcmHashFile -Path $seed.Hash).Hash
        $run = Invoke-Entry -Folder $seed.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $seed.Config) -ExchangeFails
        $run.Code | Should -Be 1 -Because $run.Text
        (Read-PcmHashFile -Path $seed.Hash).Hash | Should -BeExactly $before
        (Test-PcmSnapshotHash -Path $seed.Db).State | Should -Be 'Verified'
    }
    It 'writes the hash file again for the database as it is when the Collect fails after it opened it' {
        $seed = New-SeededFolder -Name 'failbackup'
        $blocker = Join-Path $script:Work 'a-file-not-a-folder'
        [IO.File]::WriteAllText($blocker, 'x')
        $config = New-DataConfig -Folder $seed.Folder -BackupFolder $blocker
        $run = Invoke-Entry -Folder $seed.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $config) -Mailboxes 3
        $run.Code | Should -Be 1 -Because $run.Text
        Test-Path -LiteralPath $seed.Hash | Should -BeTrue
        (Test-PcmSnapshotHash -Path $seed.Db).State | Should -Be 'Verified'
        @(Get-CompleteSnapshot $seed.Db).Count | Should -Be 2
    }
    It 'does not bless a database that did not match its hash file before the failed Collect' {
        $seed = New-SeededFolder -Name 'failtampered'
        $connection = Open-PcmStore -Path $seed.Db -Root $script:Root
        try { $null = Invoke-PcmSql $connection "UPDATE snapshot SET exchange_server = 'changed.contoso.com'" } finally { Close-PcmStore $connection }
        $stale = (Read-PcmHashFile -Path $seed.Hash).Hash
        $blocker = Join-Path $script:Work 'a-file-not-a-folder-2'
        [IO.File]::WriteAllText($blocker, 'x')
        $config = New-DataConfig -Folder $seed.Folder -BackupFolder $blocker
        $null = Invoke-Entry -Folder $seed.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $config) -Mailboxes 3
        (Read-PcmHashFile -Path $seed.Hash).Hash | Should -BeExactly $stale
        (Test-PcmSnapshotHash -Path $seed.Db).State | Should -Be 'Changed'
    }
}

Describe 'Collect keeps the snapshot of a batch that is not rolled back' {
    It 'does not delete it beyond Store.KeepSnapshots, says so, and deletes it once the batch is rolled back' {
        $seed = New-SeededFolder -Name 'prune' -StoreExtra 'KeepSnapshots = 2'
        $snapshot = (Get-CompleteSnapshot $seed.Db)[0]
        $batch = New-ConvertBatch -Folder $seed.Folder -SnapshotId $snapshot.id -RunId $snapshot.run_id -Hash (Get-DbDigest $seed.Db $snapshot.id)
        foreach ($number in 1..3) {
            (Invoke-Entry -Folder $seed.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $seed.Config) -Mailboxes (3 + $number)).Code | Should -Be 0
        }
        @(Get-CompleteSnapshot $seed.Db | ForEach-Object { $_.id }) | Should -Be @(1, 3, 4)
        $text = (Get-LogText -Folder $seed.Folder)
        $text | Should -Match '1 old snapshot\(s\) kept although older than Store.KeepSnapshots = 2: a Convert batch that is not rolled back yet needs them \(snapshot 1\)'
        $journal = Open-PcmStore -Path $seed.Journal -Root $script:Root -Kind Journal
        try {
            $recover = New-PcmBatch $journal Recover @{ ConvertBatch = $batch; SnapshotId = 1; Environment = 'TEST'; ToolVersion = 'test' }
            Set-PcmBatchItem $journal $recover 'a0000000-0000-0000-0000-00000000000a' @{ kind = 'User'; identity = 'x'; step = 'Done'; status = 'Done' }
            Set-PcmBatch $journal $recover Complete @{}
        } finally { Close-PcmStore $journal }
        (Invoke-Entry -Folder $seed.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $seed.Config) -Mailboxes 6).Code | Should -Be 0
        @(Get-CompleteSnapshot $seed.Db | ForEach-Object { $_.id }) | Should -Be @(4, 5)
    }
    It 'deletes nothing, and says so, when the journal cannot be read' {
        $seed = New-SeededFolder -Name 'prune-broken' -StoreExtra 'KeepSnapshots = 1'
        [IO.File]::WriteAllBytes($seed.Journal, [byte[]](1..200))
        $run = Invoke-Entry -Folder $seed.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-Force', '-ConfigPath', $seed.Config) -Mailboxes 3
        $run.Code | Should -Be 0 -Because $run.Text
        @(Get-CompleteSnapshot $seed.Db).Count | Should -Be 2
        (Get-LogText -Folder $seed.Folder) | Should -Match 'could not be read .*: no old snapshot is deleted in this run'
    }
}

Describe 'Convert and Recover in Apply mode need a snapshot they can verify' {
    BeforeAll { $script:verify = New-SeededFolder -Name 'verify' }
    BeforeEach {
        Copy-Item -LiteralPath $script:verify.Db -Destination ($script:verify.Db + '.bak') -Force
    }
    AfterEach {
        if (Test-Path -LiteralPath ($script:verify.Db + '.bak')) { Move-Item -LiteralPath ($script:verify.Db + '.bak') -Destination $script:verify.Db -Force }
        $null = Write-PcmHashFile -Path $script:verify.Db
    }
    It 'stops Convert -Mode Apply, with nothing changed, when there is no hash file (<Spelling>)' -TestCases @(@{ Spelling = 'official' }, @{ Spelling = 'lower case' }) {
        param($Spelling)
        Remove-Item -LiteralPath $script:verify.Hash -Force
        $arguments = if ($Spelling -eq 'lower case') { @('-action', 'convert', '-mode', 'apply', '-force', '-configpath', $script:verify.Config) } else { @('-Action', 'Convert', '-Mode', 'Apply', '-Force', '-ConfigPath', $script:verify.Config) }
        $run = Invoke-Entry -Folder $script:verify.Folder -Arguments $arguments
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'cannot prove it is the file written by Collect'
        $run.Text | Should -Match 'Nothing was changed\. If you know that this file is the one Collect wrote, run the same command again with -AcceptUnverifiedSnapshot'
        $journal = Open-PcmStore -Path $script:verify.Journal -Root $script:Root -Kind Journal -ReadOnly
        try { [int](Invoke-PcmSql $journal 'SELECT COUNT(*) FROM batch' -As Scalar) | Should -Be 0 } finally { Close-PcmStore $journal }
    }
    It 'only warns in a Preview of Convert and in a Check' {
        Remove-Item -LiteralPath $script:verify.Hash -Force
        foreach ($arguments in @(@('-Action', 'Convert', '-ConfigPath', $script:verify.Config), @('-Action', 'Check', '-ConfigPath', $script:verify.Config))) {
            $run = Invoke-Entry -Folder $script:verify.Folder -Arguments $arguments
            $run.Text | Should -Not -Match 'If you know that this file is the one Collect wrote'
            (Get-LogText -Folder $script:verify.Folder) | Should -Match 'cannot prove it is the file written by Collect'
        }
    }
    It 'goes past the check of the snapshot with -AcceptUnverifiedSnapshot' {
        Remove-Item -LiteralPath $script:verify.Hash -Force
        $run = Invoke-Entry -Folder $script:verify.Folder -Arguments @('-Action', 'Convert', '-Mode', 'Apply', '-Force', '-AcceptUnverifiedSnapshot', '-ConfigPath', $script:verify.Config)
        $run.Text | Should -Not -Match 'If you know that this file is the one Collect wrote'
        $run.Text | Should -Match 'Integrity +Skipped on request \(-AcceptUnverifiedSnapshot\)'
        $run.Text | Should -Match 'Snapshot 1 of'
    }
    It 'stops Recover -Mode Apply, with nothing changed, when there is no hash file' {
        $snapshot = (Get-CompleteSnapshot $script:verify.Db)[0]
        $batch = New-ConvertBatch -Folder $script:verify.Folder -SnapshotId $snapshot.id -RunId $snapshot.run_id -Hash (Get-DbDigest $script:verify.Db $snapshot.id)
        Remove-Item -LiteralPath $script:verify.Hash -Force
        $run = Invoke-Entry -Folder $script:verify.Folder -Arguments @('-Action', 'Recover', '-Mode', 'Apply', '-Force', '-Batch', $batch, '-ConfigPath', $script:verify.Config)
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'Nothing was changed\. If you know that this file is the one Collect wrote, run the same command again with -AcceptUnverifiedSnapshot'
    }
}

Describe 'A batch is tied to its snapshot' {
    BeforeAll {
        $script:tie = New-SeededFolder -Name 'tie'
        $script:tieSnapshot = (Get-CompleteSnapshot $script:tie.Db)[0]
        $script:tieDigest = Get-DbDigest $script:tie.Db $script:tieSnapshot.id
        function script:Get-TieRun {
            param([string]$Action = 'Check', [Parameter(Mandatory)][string]$Batch, [string[]]$Extra = @())
            $arguments = if ($Action -eq 'Check') { @('-Action', 'Check', '-Batch', $Batch, '-Once', '-ConfigPath', $script:tie.Config) }
                elseif ($Action -eq 'Recover') { @('-Action', 'Recover', '-Mode', 'Apply', '-Force', '-Batch', $Batch, '-ConfigPath', $script:tie.Config) }
                else { @('-Action', 'Convert', '-Mode', 'Apply', '-Force', '-Batch', $Batch, '-ConfigPath', $script:tie.Config) }
            Invoke-Entry -Folder $script:tie.Folder -Arguments ($arguments + $Extra)
        }
    }
    It 'records the run and the digest of the snapshot in the batch that a Convert creates' {
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId $script:tieSnapshot.id -RunId $script:tieSnapshot.run_id -Hash $script:tieDigest
        $journal = Open-PcmStore -Path $script:tie.Journal -Root $script:Root -Kind Journal -ReadOnly
        try {
            $row = Get-PcmBatch $journal -Id $batch
            $row.snapshot_run_id | Should -BeExactly $script:tieSnapshot.run_id
            $row.snapshot_hash | Should -BeExactly $script:tieDigest
        } finally { Close-PcmStore $journal }
    }
    It 'follows a batch whose snapshot is the one it was made from, and does not stop at the snapshot (<Action>)' -TestCases @(@{ Action = 'Check' }, @{ Action = 'Recover' }, @{ Action = 'Convert' }) {
        param($Action)
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId $script:tieSnapshot.id -RunId $script:tieSnapshot.run_id -Hash $script:tieDigest
        $run = Get-TieRun -Action $Action -Batch $batch
        $run.Text | Should -Match 'Snapshot file verified'
        $run.Text | Should -Not -Match 'is not the snapshot that batch'
        $run.Text | Should -Not -Match 'which is not in'
        $run.Text | Should -Match 'Snapshot 1 of'
    }
    It 'stops, with nothing changed, when the content of the snapshot is not the one of the batch (<Action>)' -TestCases @(@{ Action = 'Check' }, @{ Action = 'Recover' }, @{ Action = 'Convert' }) {
        param($Action)
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId $script:tieSnapshot.id -RunId $script:tieSnapshot.run_id -Hash ('ab' * 32)
        $run = Get-TieRun -Action $Action -Batch $batch
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match ('Snapshot 1 of .* is not the snapshot that batch {0} was made from' -f $batch)
        $run.Text | Should -Match 'Restore the right database from Store\.BackupFolder'
        $run.Text | Should -Match '-AcceptUnverifiedSnapshot'
    }
    It 'goes on with a warning when the content differs and -AcceptUnverifiedSnapshot is given' {
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId $script:tieSnapshot.id -RunId $script:tieSnapshot.run_id -Hash ('ab' * 32)
        $run = Get-TieRun -Action Check -Batch $batch -Extra @('-AcceptUnverifiedSnapshot')
        $run.Text | Should -Not -Match 'is not the snapshot that batch .* was made from: the mailboxes'
        (Get-LogText -Folder $script:tie.Folder) | Should -Match 'is not the snapshot that batch .* was made from; the run goes on because of -AcceptUnverifiedSnapshot'
    }
    It 'stops and points to Store.BackupFolder when the snapshot of the batch is not in the database any more (<Action>)' -TestCases @(@{ Action = 'Check' }, @{ Action = 'Recover' }, @{ Action = 'Convert' }) {
        param($Action)
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId 57 -RunId 'a-run-that-was-pruned' -Hash $script:tieDigest
        $run = Get-TieRun -Action $Action -Batch $batch
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match ('Batch {0} was made from snapshot 57 \(Collect run a-run-that-was-pruned\), which is not in' -f $batch)
        $run.Text | Should -Match 'Store\.BackupFolder'
        $run.Text | Should -Match 'PraCloudMailbox_<Environment>_snapshot<number>_<date>\.db'
    }
    It 'does not take another snapshot for the one of the batch when the database was recreated (the number is the same, the run is not)' {
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId $script:tieSnapshot.id -RunId 'run-of-another-database' -Hash $script:tieDigest
        $run = Get-TieRun -Action Recover -Batch $batch
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'which is not in'
    }
    It 'points to Store.BackupFolder when -Snapshot names a snapshot that is not there' {
        $run = Invoke-Entry -Folder $script:tie.Folder -Arguments @('-Action', 'Check', '-Snapshot', '57', '-ConfigPath', $script:tie.Config)
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'Snapshot 57 not found in .*pra\.db'
        $run.Text | Should -Match 'Store\.BackupFolder'
    }
    It 'still reads a batch made before batches kept their snapshot, by its number, with a warning' {
        $batch = New-ConvertBatch -Folder $script:tie.Folder -SnapshotId $script:tieSnapshot.id
        $run = Get-TieRun -Action Check -Batch $batch
        $run.Text | Should -Match 'Snapshot 1 of'
        $run.Text | Should -Not -Match 'is not the snapshot that batch'
        (Get-LogText -Folder $script:tie.Folder) | Should -Match ('Batch {0} was made before batches kept the content of their snapshot' -f $batch)
    }
}

Describe 'Entra Connect Script mode: the signed sample runs as it is' {
    BeforeAll {
        $script:sample = Join-Path $script:Root 'config\EntraConnect.sample.ps1'
        $script:invoked = New-Object 'Collections.Generic.List[object]'
        function global:Import-Clixml { param([string]$LiteralPath) $script:invoked.Add(@('Import-Clixml', $LiteralPath)); [pscredential]::new('CONTOSO\svc-pra-sync', (ConvertTo-SecureString 'x' -AsPlainText -Force)) }
        function global:Invoke-Command { param($ComputerName, $Credential, $ScriptBlock) $script:invoked.Add(@('Invoke-Command', $ComputerName, $Credential.UserName)); 'SyncCycleEnabled=True InProgress=False' }
        function script:New-EntraContext {
            param([string]$ScriptPath, [string]$Server = '', [string]$CredentialPath = '')
            @{ Config = @{ EntraConnect = @{ Mode = 'Script'; Server = $Server; ScriptPath = $ScriptPath; CredentialPath = $CredentialPath; MinVersion = '2.5.76.0' } }; ConvertBatchId = '' }
        }
        Import-Module (Join-Path $script:Root 'module\PRA.Cloud.psm1') -Force
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $true }
    }
    AfterAll {
        Set-PcmWriteApproval -Approval $null
        foreach ($name in 'Import-Clixml', 'Invoke-Command') { Remove-Item -LiteralPath "Function:\$name" -Force -ErrorAction SilentlyContinue }
    }
    BeforeEach { $script:invoked.Clear() }

    It 'gives the server and the credential file of the configuration to the unmodified sample' {
        $credential = Join-Path $TestDrive 'aadc.cred.xml'
        [IO.File]::WriteAllText($credential, '<Objs/>')
        $text = Invoke-PcmEntraConnect -Context (New-EntraContext -ScriptPath $script:sample -Server 'aadc01.contoso.com' -CredentialPath $credential) -Operation Pause -Confirm:$false
        $text | Should -BeLike 'Pause on aadc01.contoso.com: SyncCycleEnabled=True InProgress=False'
        $script:invoked[0][1] | Should -BeExactly $credential
        $script:invoked[1][1] | Should -BeExactly 'aadc01.contoso.com'
        $script:invoked[1][2] | Should -BeExactly 'CONTOSO\svc-pra-sync'
    }
    It 'uses EntraConnect.cred.xml beside the sample when no credential file is configured' {
        $beside = Join-Path $script:Root 'config\EntraConnect.cred.xml'
        [IO.File]::WriteAllText($beside, '<Objs/>')
        try { $null = Invoke-PcmEntraConnect -Context (New-EntraContext -ScriptPath $script:sample -Server 'aadc01.contoso.com') -Operation Delta -Confirm:$false }
        finally { Remove-Item -LiteralPath $beside -Force -ErrorAction SilentlyContinue }
        $script:invoked[0][1] | Should -BeExactly $beside
    }
    It 'says in plain words what is missing: the server, then the credential file, with the exit code of the script' {
        $credential = Join-Path $TestDrive 'there.cred.xml'
        [IO.File]::WriteAllText($credential, '<Objs/>')
        { Invoke-PcmEntraConnect -Context (New-EntraContext -ScriptPath $script:sample -Server '' -CredentialPath $credential) -Operation Pause -Confirm:$false } |
            Should -Throw '*ended with exit code 1: Pause on  failed: EntraConnect.Server is empty: set it to the host name of the Entra Connect server in the configuration file.*'
        { Invoke-PcmEntraConnect -Context (New-EntraContext -ScriptPath $script:sample -Server 'aadc01.contoso.com' -CredentialPath (Join-Path $TestDrive 'missing.xml')) -Operation Pause -Confirm:$false } |
            Should -Throw '*the credential file * does not exist: create it with Get-Credential | Export-Clixml, or set EntraConnect.CredentialPath.*'
    }
    It 'calls a script of the administrator that does not declare -Server and -CredentialPath with -Operation only, as before' {
        $own = Join-Path $TestDrive 'own.ps1'
        Set-Content -LiteralPath $own -Value 'param([Parameter(Mandatory)][string]$Operation) "own $Operation"' -Encoding UTF8
        Invoke-PcmEntraConnect -Context (New-EntraContext -ScriptPath $own -Server 'aadc01.contoso.com' -CredentialPath 'C:\x.xml') -Operation Resume -Confirm:$false | Should -BeExactly 'own Resume'
    }
    It 'gives them to a script of the administrator that declares them' {
        $own = Join-Path $TestDrive 'own2.ps1'
        Set-Content -LiteralPath $own -Value 'param([string]$Operation, [string]$Server, [string]$CredentialPath) "own $Operation $Server $CredentialPath"' -Encoding UTF8
        Invoke-PcmEntraConnect -Context (New-EntraContext -ScriptPath $own -Server 'aadc01.contoso.com' -CredentialPath 'C:\x.xml') -Operation Resume -Confirm:$false | Should -BeExactly 'own Resume aadc01.contoso.com C:\x.xml'
    }
    It 'declares -Server and -CredentialPath in the sample, with no server written in it' {
        $text = [IO.File]::ReadAllText($script:sample)
        $text | Should -Match '\[string\]\$Server'
        $text | Should -Match '\[string\]\$CredentialPath'
        $text | Should -Not -Match "aadc01\.contoso\.com'"
    }
    It 'reads the new settings of the configuration, in any letter case, and resolves the credential path from the tool folder' {
        $path = Join-Path $TestDrive 'entra.psd1'
        Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Scope = @{ Mode = 'Auto' }; entraconnect = @{ mode = 'script'; server = 'aadc'; scriptpath = '.\config\EntraConnect.sample.ps1'; credentialpath = '.\config\aadc.cred.xml' } }"
        $config = Import-PcmConfiguration -Path $path -Root $script:Root
        $config.EntraConnect.Mode | Should -BeExactly 'Script'
        $config.EntraConnect.CredentialPath | Should -BeExactly (Join-Path $script:Root 'config\aadc.cred.xml')
        (Import-PcmConfiguration -Path (Join-Path $script:Repo 'package\config\PraCloudMailbox.config.psd1') -Root $script:Root).EntraConnect.CredentialPath | Should -BeExactly ''
        (Import-PcmConfiguration -Path (Join-Path $script:Repo 'package\config\PraCloudMailbox.config.psd1') -Root $script:Root).Collect.MaxDropPercent | Should -Be 50
    }
}

Describe 'The folders of the store are created private' {
    BeforeAll {
        function script:Get-FolderSids { param([string]$Path) @((Get-Acl -LiteralPath $Path).Access | ForEach-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } | Sort-Object -Unique) }
    }
    It 'creates the folder of a database with protected permissions: the account, SYSTEM and the Administrators only' {
        $folder = Join-Path $TestDrive 'private-db\data'
        $cn = Open-PcmStore -Path (Join-Path $folder 'pra.db') -Root $script:Root
        Close-PcmStore $cn
        (Get-Acl -LiteralPath $folder).AreAccessRulesProtected | Should -BeTrue
        $sids = Get-FolderSids $folder
        $sids | Should -Contain 'S-1-5-18'
        $sids | Should -Contain 'S-1-5-32-544'
        $sids.Count | Should -BeLessOrEqual 3
    }
    It 'creates the folder of a journal private as well' {
        $folder = Join-Path $TestDrive 'private-journal\data'
        $cn = Open-PcmStore -Path (Join-Path $folder 'journal.db') -Root $script:Root -Kind Journal
        Close-PcmStore $cn
        (Get-Acl -LiteralPath $folder).AreAccessRulesProtected | Should -BeTrue
    }
    It 'creates the backup folder private' {
        $cn = Open-PcmStore -Path (Join-Path $TestDrive 'private-backup\pra.db') -Root $script:Root
        try {
            $folder = Join-Path $TestDrive 'private-backup\copies'
            $null = Backup-PcmStore -Connection $cn -Folder $folder -Name 'copy.db' -Confirm:$false
            (Get-Acl -LiteralPath $folder).AreAccessRulesProtected | Should -BeTrue
            Join-Path $folder 'copy.db' | Should -Exist
        } finally { Close-PcmStore $cn }
    }
    It 'leaves an existing folder as the administrator set it' {
        $folder = Join-Path $TestDrive 'existing-db'
        $null = New-Item -ItemType Directory -Path $folder
        $cn = Open-PcmStore -Path (Join-Path $folder 'pra.db') -Root $script:Root
        Close-PcmStore $cn
        (Get-Acl -LiteralPath $folder).AreAccessRulesProtected | Should -BeFalse
    }
    It 'stops with the folder and the reason when the folder cannot be created' {
        $file = Join-Path $TestDrive 'a-file-not-a-folder'
        Set-Content -LiteralPath $file -Value 'x'
        { Open-PcmStore -Path (Join-Path $file 'pra.db') -Root $script:Root } | Should -Throw '*cannot be created*'
        { Open-PcmStore -Path (Join-Path $TestDrive 'bad|name\pra.db') -Root $script:Root } | Should -Throw '*cannot be created*'
    }
}
