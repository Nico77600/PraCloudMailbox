<#
.SYNOPSIS
    Legacy stores: what earlier releases wrote stays readable (offline; synthetic data only).
.DESCRIPTION
    tests\fixtures\legacy holds three stores written, byte for byte, by the store module of the releases that created them (tools\New-LegacyFixtures.ps1):
    a snapshot of schema 1 (release 1.0.0), a snapshot of schema 2 and a journal of schema 1 (release 1.2.0). Their SHA-256 are listed in fixtures.sha256 and checked
    first. Every test works on a copy: the committed files never change.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
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
    $script:Gui = Import-Module (Join-Path $script:Root 'module\PRA.Gui.psm1') -Force -PassThru
    $script:Legacy = Join-Path $PSScriptRoot 'fixtures\legacy'
    $script:Work = Join-Path ([IO.Path]::GetTempPath()) ('PL-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $null = New-Item -ItemType Directory -Path $script:Work -Force
    function script:Copy-Legacy([string]$Name) {
        $target = Join-Path $script:Work ('{0}-{1}' -f [guid]::NewGuid().ToString('N').Substring(0, 6), $Name)
        Copy-Item -LiteralPath (Join-Path $script:Legacy $Name) -Destination $target
        $target
    }
}

AfterAll {
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue
    Remove-PcmTestPackage $script:Root
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'The frozen stores' {
    It 'match the SHA-256 that was written with them' {
        $lines = @([IO.File]::ReadAllLines((Join-Path $script:Legacy 'fixtures.sha256')) | Where-Object { $_.Trim() })
        $lines.Count | Should -Be 3
        foreach ($line in $lines) {
            $hash, $name = $line -split '\s+', 2
            (Get-FileHash -LiteralPath (Join-Path $script:Legacy $name.Trim()) -Algorithm SHA256).Hash.ToLowerInvariant() | Should -BeExactly $hash -Because $name
        }
    }
    It 'carry the schema versions of the releases that wrote them' {
        $cn = Open-PcmStore -Path (Copy-Legacy 'snapshot-schema1.db') -Root $script:Root -ReadOnly
        try { Invoke-PcmSql $cn "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar | Should -Be '1' } finally { Close-PcmStore $cn }
        $cn = Open-PcmStore -Path (Copy-Legacy 'snapshot-schema2.db') -Root $script:Root -ReadOnly
        try { Invoke-PcmSql $cn "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar | Should -Be '2' } finally { Close-PcmStore $cn }
        $cn = Open-PcmStore -Path (Copy-Legacy 'journal-schema1.db') -Root $script:Root -Kind Journal -ReadOnly
        try { Invoke-PcmSql $cn "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar | Should -Be '1' } finally { Close-PcmStore $cn }
    }
}

Describe 'Snapshots of schema 1 and 2 stay readable' {
    It 'reads the snapshot of schema <Schema> read-only: the last complete snapshot, its mailboxes, its permissions' -ForEach @(
        @{ Schema = 1; Name = 'snapshot-schema1.db'; Mailboxes = 3 }, @{ Schema = 2; Name = 'snapshot-schema2.db'; Mailboxes = 4 }
    ) {
        $cn = Open-PcmStore -Path (Copy-Legacy $Name) -Root $script:Root -ReadOnly
        try {
            $snapshot = Get-PcmSnapshot $cn
            $snapshot.status | Should -Be 'Complete'
            $snapshot.exchange_server | Should -Be 'SRV-EXCH01'
            $rows = @(Read-PcmTable $cn mailbox $snapshot.id)
            $rows.Count | Should -Be $Mailboxes
            @($rows | Where-Object { $_.kind -eq 'Shared' }).Count | Should -Be 1
            ($rows | Where-Object { $_.user_principal_name -eq 'adele.vance@contoso.com' }).immutable_id | Should -Not -BeNullOrEmpty
            @(Read-PcmTable $cn permission $snapshot.id).Count | Should -Be 1
            Test-PcmTable $cn target_group | Should -BeFalse
        } finally { Close-PcmStore $cn }
    }
    It 'reads the resource settings of a room in a schema 2 snapshot' {
        $cn = Open-PcmStore -Path (Copy-Legacy 'snapshot-schema2.db') -Root $script:Root -ReadOnly
        try {
            $room = @(Read-PcmTable $cn mailbox (Get-PcmSnapshot $cn).id | Where-Object { $_.kind -eq 'Room' })[0]
            ($room.resource_settings_json | ConvertFrom-Json).Capacity | Should -Be 8
        } finally { Close-PcmStore $cn }
    }
    It 'opens a snapshot of schema <Schema> for writing: the missing parts are added and the old rows stay' -ForEach @(
        @{ Schema = 1; Name = 'snapshot-schema1.db'; Mailboxes = 3 }, @{ Schema = 2; Name = 'snapshot-schema2.db'; Mailboxes = 4 }
    ) {
        $path = Copy-Legacy $Name
        $cn = Open-PcmStore -Path $path -Root $script:Root
        try {
            Test-PcmTable $cn target_group | Should -BeTrue
            $id = (Get-PcmSnapshot $cn).id
            @(Read-PcmTable $cn mailbox $id).Count | Should -Be $Mailboxes
            (Invoke-PcmSql $cn 'PRAGMA table_info(mailbox)' -As Rows | ForEach-Object name) | Should -Contain 'resource_settings_json'
            @((Find-PcmTargetGroup $cn $id 'PRA-Wave1').Groups).Count | Should -Be 0
        } finally { Close-PcmStore $cn }
    }
    It 'shows the objects of the snapshot in the window state, and a group asked of an old snapshot is looked up in Microsoft Entra' {
        $config = Join-Path $script:Work 'legacy.config.psd1'
        $path = Copy-Legacy 'snapshot-schema1.db'
        [IO.File]::WriteAllText($config, "@{ Environment = 'PROD'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Store = @{ Path = '$($path.Replace("'", "''"))'; JournalPath = '$((Join-Path $script:Work 'none-journal.db').Replace("'", "''"))'; BackupFolder = '$((Join-Path $script:Work 'backup').Replace("'", "''"))' }; Logging = @{ Folder = '$((Join-Path $script:Work 'logs').Replace("'", "''"))' }; Report = @{ Folder = '$((Join-Path $script:Work 'reports').Replace("'", "''"))' } }", [Text.UTF8Encoding]::new($true))
        $state = Read-PcmGuiState -Root $script:Root -ConfigPath $config
        $state.ConfigError | Should -BeNullOrEmpty
        $state.SnapshotError | Should -BeNullOrEmpty
        @($state.Objects).Count | Should -Be 3
        $lookup = & $script:Gui { param($s) Get-PcmGuiGroupInSnapshot -State $s -Value 'PRA-Wave1' } $state
        [bool]$lookup.HasTables | Should -BeFalse
        [int]$lookup.Count | Should -Be 0
    }
}

Describe 'Journal of schema 1 stays readable' {
    It 'reads the batches, the objects, the events and the tenant changes of a journal of schema 1' {
        $cn = Open-PcmStore -Path (Copy-Legacy 'journal-schema1.db') -Root $script:Root -Kind Journal -ReadOnly
        try {
            $convert = Get-PcmBatch $cn -Action Convert
            $convert.status | Should -Be 'Complete'
            @(Get-PcmBatchItem $cn $convert.id).Count | Should -Be 2
            (Invoke-PcmSql $cn 'SELECT count(*) FROM batch_event WHERE batch_id = @b' @{ b = $convert.id } -As Scalar) | Should -Be 2
            @(Get-PcmTenantChange $cn $convert.id).Count | Should -Be 1
            $recover = Get-PcmBatch $cn -Action Recover
            $recover.convert_batch | Should -Be $convert.id
            @(Get-PcmBatchItem $cn $recover.id).Count | Should -Be 1
        } finally { Close-PcmStore $cn }
    }
    It 'opens a journal of schema 1 for writing, adds the runs of a batch and keeps the old batches' {
        $cn = Open-PcmStore -Path (Copy-Legacy 'journal-schema1.db') -Root $script:Root -Kind Journal
        try {
            $convert = Get-PcmBatch $cn -Action Convert
            Add-PcmBatchRun $cn $convert.id @{ RunId = 'run-1'; Action = 'Convert'; Mode = 'Preview'; Started = [datetime]::UtcNow; Account = 'CONTOSO\pra.operator'; ToolVersion = '2.0.0'; Status = 'Complete'; ExitCode = 0 }
            @((Get-PcmBatchHistory $cn $convert.id).Runs).Count | Should -Be 1
            @(Get-PcmBatchItem $cn $convert.id).Count | Should -Be 2
        } finally { Close-PcmStore $cn }
    }
    It 'follows the batches of a journal of schema 1 in the window state' {
        $config = Join-Path $script:Work 'legacy-journal.config.psd1'
        $path = Copy-Legacy 'journal-schema1.db'
        $snapshot = Copy-Legacy 'snapshot-schema2.db'
        [IO.File]::WriteAllText($config, "@{ Environment = 'PROD'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Store = @{ Path = '$($snapshot.Replace("'", "''"))'; JournalPath = '$($path.Replace("'", "''"))'; BackupFolder = '$((Join-Path $script:Work 'backup').Replace("'", "''"))' }; Logging = @{ Folder = '$((Join-Path $script:Work 'logs').Replace("'", "''"))' }; Report = @{ Folder = '$((Join-Path $script:Work 'reports').Replace("'", "''"))' } }", [Text.UTF8Encoding]::new($true))
        $state = Read-PcmGuiState -Root $script:Root -ConfigPath $config
        $state.JournalError | Should -BeNullOrEmpty
        @($state.Batches).Count | Should -Be 2
        @(Get-PcmGuiBatchEntry -State $state -Kind Follow).Count | Should -Be 2
    }
}
