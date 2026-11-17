<#
.SYNOPSIS
    PRA Cloud Mailbox - storage module: the SQLite database that keeps the Collect snapshots.

.DESCRIPTION
    Collect (Exchange on-premises) writes one snapshot per run: the mailboxes in scope, the shared mailbox
    permissions, and optionally the contacts and distribution groups. The cloud actions read the last
    complete snapshot, when Exchange and Active Directory are gone.

    One file, one engine: System.Data.SQLite 1.0.119 (SQLite 3.46.1, public domain), netstandard2.0 build
    for PowerShell 7 with the native x64 SQLite.Interop.dll (lib\sqlite).

    A snapshot is written inside its own rows with the status Running, then marked Complete at the end:
    an interrupted Collect never replaces the last good snapshot. Old snapshots are pruned (Store.KeepSnapshots)
    and every Apply writes a consistent copy of the database (VACUUM INTO) to Store.BackupFolder.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
#>
#Requires -Version 7.5
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# >>> PRA package integrity bootstrap: identical in every runtime PowerShell file (checked by the tests).
# It authenticates module\Test-PackageIntegrity.ps1 before running it; there is no unsigned or bypass path.
& {
    param([Parameter(Mandatory)][string]$PackageRoot)
    $verifier = [IO.Path]::GetFullPath([IO.Path]::Combine($PackageRoot, 'module', 'Test-PackageIntegrity.ps1'))
    $cursor = $verifier
    while ($cursor) {
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if (-not [string]::Equals([string]$signature.Status, 'Valid', [StringComparison]::Ordinal) -or $null -eq $signature.SignerCertificate -or
        -not [string]::Equals([string]$signature.SignerCertificate.Thumbprint, '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6', [StringComparison]::Ordinal) -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
    }
    & $verifier -Root $PackageRoot
} ([IO.Path]::GetDirectoryName($PSScriptRoot))
# <<< PRA package integrity bootstrap
Import-Module (Join-Path $PSScriptRoot 'PRA.Common.psm1') -ErrorAction Stop
$script:SchemaVersion = 3
$script:SqliteLoaded = $false

# Snapshot schema 2 adds optional resource settings and schema 3 the groups captured as run targets (target_group, target_member); older snapshots remain readable.
$script:Schema = @(
    'CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)'
    @'
CREATE TABLE IF NOT EXISTS snapshot (
    id INTEGER PRIMARY KEY AUTOINCREMENT, run_id TEXT NOT NULL,
    status TEXT NOT NULL CHECK (status IN ('Running','Complete','Failed')),
    started_utc TEXT NOT NULL, finished_utc TEXT, environment TEXT, computer TEXT, account TEXT, tool_version TEXT,
    exchange_server TEXT, scope_json TEXT, counts_json TEXT, note TEXT)
'@
    @'
CREATE TABLE IF NOT EXISTS mailbox (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, object_guid TEXT NOT NULL,
    kind TEXT NOT NULL CHECK (kind IN ('User','Shared','Room','Equipment')), recipient_type_details TEXT,
    sam_account_name TEXT, user_principal_name TEXT, name TEXT, display_name TEXT, alias TEXT, distinguished_name TEXT,
    organizational_unit TEXT, primary_smtp_address TEXT, windows_email_address TEXT, legacy_exchange_dn TEXT,
    exchange_guid TEXT, archive_guid TEXT, archive_status TEXT, hidden_from_address_lists INTEGER, immutable_id TEXT,
    database TEXT, email_addresses_json TEXT, custom_attributes_json TEXT, grant_send_on_behalf_json TEXT,
    item_count INTEGER, total_item_size_bytes INTEGER, resource_settings_json TEXT,
    PRIMARY KEY (snapshot_id, object_guid))
'@
    @'
CREATE TABLE IF NOT EXISTS permission (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, mailbox_guid TEXT NOT NULL,
    access_right TEXT NOT NULL CHECK (access_right IN ('FullAccess','SendAs','SendOnBehalf')),
    trustee TEXT NOT NULL, trustee_guid TEXT, trustee_kind TEXT, trustee_upn TEXT, trustee_smtp TEXT, trustee_name TEXT,
    trustee_dn TEXT, auto_mapping INTEGER, resolved INTEGER NOT NULL DEFAULT 0, note TEXT)
'@
    @'
CREATE TABLE IF NOT EXISTS group_member (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, group_guid TEXT NOT NULL,
    member_guid TEXT, member_kind TEXT, member_upn TEXT, member_smtp TEXT, member_name TEXT, via_group_guid TEXT, depth INTEGER)
'@
    @'
CREATE TABLE IF NOT EXISTS contact (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, object_guid TEXT NOT NULL,
    name TEXT, display_name TEXT, alias TEXT, external_email_address TEXT, primary_smtp_address TEXT, legacy_exchange_dn TEXT,
    hidden_from_address_lists INTEGER, organizational_unit TEXT, email_addresses_json TEXT, custom_attributes_json TEXT,
    PRIMARY KEY (snapshot_id, object_guid))
'@
    @'
CREATE TABLE IF NOT EXISTS distribution_group (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, object_guid TEXT NOT NULL,
    kind TEXT NOT NULL CHECK (kind IN ('Distribution','Security','Dynamic')), name TEXT, display_name TEXT, alias TEXT,
    primary_smtp_address TEXT, legacy_exchange_dn TEXT, hidden_from_address_lists INTEGER, organizational_unit TEXT,
    email_addresses_json TEXT, managed_by_json TEXT, settings_json TEXT, recipient_filter TEXT, member_count INTEGER,
    PRIMARY KEY (snapshot_id, object_guid))
'@
    @'
CREATE TABLE IF NOT EXISTS dl_member (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, group_guid TEXT NOT NULL,
    member_guid TEXT, member_kind TEXT, member_name TEXT, member_smtp TEXT, member_upn TEXT)
'@
    @'
CREATE TABLE IF NOT EXISTS collect_issue (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, level TEXT NOT NULL, object TEXT,
    operation TEXT, message TEXT)
'@
    @'
CREATE TABLE IF NOT EXISTS target_group (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, group_guid TEXT NOT NULL,
    name TEXT, sam_account_name TEXT, distinguished_name TEXT, primary_smtp_address TEXT, display_name TEXT, member_count INTEGER,
    PRIMARY KEY (snapshot_id, group_guid))
'@
    @'
CREATE TABLE IF NOT EXISTS target_member (
    snapshot_id INTEGER NOT NULL REFERENCES snapshot(id) ON DELETE CASCADE, group_guid TEXT NOT NULL,
    member_guid TEXT, member_class TEXT, member_name TEXT, member_upn TEXT, member_smtp TEXT, member_dn TEXT, via_group_guid TEXT, depth INTEGER)
'@
    'CREATE INDEX IF NOT EXISTS ix_permission ON permission (snapshot_id, mailbox_guid)'
    'CREATE INDEX IF NOT EXISTS ix_group_member ON group_member (snapshot_id, group_guid)'
    'CREATE INDEX IF NOT EXISTS ix_dl_member ON dl_member (snapshot_id, group_guid)'
    'CREATE INDEX IF NOT EXISTS ix_target_member ON target_member (snapshot_id, group_guid)'
)
$script:Tables = @('mailbox','permission','group_member','contact','distribution_group','dl_member','collect_issue','target_group','target_member')

# Journal of the cloud actions (Convert, Recover): a separate file on the cloud admin server, so that a new
# Collect copied over the snapshot database never erases a batch. Every step of every object is recorded:
# a run interrupted (session cut, reboot) is resumed with the same -Batch.
$script:JournalSchemaVersion = 3
# Journal schema 3 adds snapshot_run_id and snapshot_hash to batch (the snapshot a batch was made from, by its run and by its content) and an index on batch_item (object_guid); older journals stay readable.
# The column tenant_id of batch (the tenant a batch was made in) is added to a journal that lacks it, without a change of the schema version: a journal without it is read as a batch of no known tenant.
$script:JournalSchema = @(
    'CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)'
    @'
CREATE TABLE IF NOT EXISTS batch (
    id TEXT PRIMARY KEY, action TEXT NOT NULL CHECK (action IN ('Convert','Recover')), convert_batch TEXT,
    status TEXT NOT NULL CHECK (status IN ('Running','Complete','Partial','Failed')),
    created_utc TEXT NOT NULL, updated_utc TEXT, snapshot_id INTEGER, environment TEXT, computer TEXT, account TEXT,
    tool_version TEXT, scope_json TEXT, counts_json TEXT, note TEXT, snapshot_run_id TEXT, snapshot_hash TEXT, tenant_id TEXT)
'@
    @'
CREATE TABLE IF NOT EXISTS batch_item (
    batch_id TEXT NOT NULL REFERENCES batch(id) ON DELETE CASCADE, object_guid TEXT NOT NULL,
    kind TEXT NOT NULL, identity TEXT, entra_id TEXT, step TEXT,
    status TEXT NOT NULL CHECK (status IN ('Planned','Running','Done','Failed','Skipped')),
    original_json TEXT, result_json TEXT, message TEXT, updated_utc TEXT,
    PRIMARY KEY (batch_id, object_guid))
'@
    @'
CREATE TABLE IF NOT EXISTS batch_event (
    batch_id TEXT NOT NULL REFERENCES batch(id) ON DELETE CASCADE, object_guid TEXT, time_utc TEXT NOT NULL,
    step TEXT NOT NULL, outcome TEXT NOT NULL CHECK (outcome IN ('Ok','Fail','Info')), detail TEXT)
'@
    @'
CREATE TABLE IF NOT EXISTS tenant_change (
    batch_id TEXT NOT NULL REFERENCES batch(id) ON DELETE CASCADE, kind TEXT NOT NULL, target_id TEXT NOT NULL,
    original_json TEXT, restored INTEGER NOT NULL DEFAULT 0, updated_utc TEXT, PRIMARY KEY (batch_id, kind, target_id))
'@
    @'
CREATE TABLE IF NOT EXISTS batch_run (
    run_id TEXT PRIMARY KEY, batch_id TEXT NOT NULL REFERENCES batch(id) ON DELETE CASCADE, action TEXT NOT NULL, mode TEXT,
    started_utc TEXT NOT NULL, ended_utc TEXT, account TEXT, computer TEXT, tool_version TEXT, source_text TEXT, overrides_text TEXT,
    status TEXT, exit_code INTEGER, done_count INTEGER, pending_count INTEGER, error_count INTEGER, skipped_count INTEGER, log_file TEXT)
'@
    'CREATE INDEX IF NOT EXISTS ix_batch_event ON batch_event (batch_id, object_guid)'
    'CREATE INDEX IF NOT EXISTS ix_batch_item_object ON batch_item (object_guid)'
)

function Import-PcmSqlite {
    <#
    .SYNOPSIS
        Loads System.Data.SQLite (once per process).
    .PARAMETER Root
        Tool folder (holds lib\sqlite).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)
    if ($script:SqliteLoaded) { return }
    if (-not [Environment]::Is64BitProcess) { throw 'The SQLite engine of the tool is 64-bit: run a 64-bit PowerShell.' }
    $already = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.Data.SQLite' } | Select-Object -First 1
    if (-not $already) {
        $dll = Join-Path $Root 'lib\sqlite\netstandard2.0\System.Data.SQLite.dll'
        if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) { throw "SQLite library missing: $dll (copy the whole lib folder of the tool)." }
        Add-Type -LiteralPath $dll
    }
    $script:SqliteLoaded = $true
}

function Open-PcmStore {
    <#
    .SYNOPSIS
        Opens (and creates, unless -ReadOnly) the SQLite database, and checks or creates the schema.
    .OUTPUTS
        System.Data.SQLite.SQLiteConnection (open). Close it with Close-PcmStore.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root, [switch]$ReadOnly, [ValidateSet('Snapshot','Journal')][string]$Kind = 'Snapshot')
    $null = Import-PcmSqlite -Root $Root
    if ($ReadOnly -and -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw $(if ($Kind -eq 'Journal') { "Journal not found: $Path (it is written by Convert)." } else { "Database not found: $Path. Run -Action Collect -Mode Apply first (it writes the database), or copy the database file of the last snapshot here." })
    }
    $schema = if ($Kind -eq 'Journal') { $script:JournalSchema } else { $script:Schema }
    $expected = if ($Kind -eq 'Journal') { $script:JournalSchemaVersion } else { $script:SchemaVersion }
    $folder = Split-Path -Parent $Path
    if (-not $ReadOnly -and $folder) { Initialize-PcmPrivateFolder -Path $folder }
    $builder = New-Object System.Data.SQLite.SQLiteConnectionStringBuilder
    $builder.DataSource = $Path
    $builder.ForeignKeys = $true
    $builder.FailIfMissing = [bool]$ReadOnly
    $builder.ReadOnly = [bool]$ReadOnly
    $connection = New-Object System.Data.SQLite.SQLiteConnection($builder.ToString())
    $connection.Open()
    try {
        $version = $null; $storedKind = $null
        $hasMeta = Invoke-PcmSql $connection "SELECT count(*) FROM sqlite_master WHERE type = 'table' AND name = 'meta'" -As Scalar
        if ([int]$hasMeta -gt 0) {
            $version = Invoke-PcmSql $connection "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar
            $storedKind = Invoke-PcmSql $connection "SELECT value FROM meta WHERE key = 'kind'" -As Scalar
        }
        $hasVersion = $null -ne $version -and $version -isnot [DBNull]
        $actualKind = if ($null -ne $storedKind -and $storedKind -isnot [DBNull]) { [string]$storedKind } elseif ($hasVersion) { 'Snapshot' } else { '' }
        if ($actualKind -and $actualKind -ne $Kind) { throw "$Path is a $($actualKind.ToLowerInvariant()) database, not a $($Kind.ToLowerInvariant()) database." }
        if ($hasVersion -and [int]$version -gt $expected) {
            throw "Database schema version $version is newer than this tool (version $expected): use the same tool version as the Collect."
        }
        if (-not $ReadOnly) {
            foreach ($statement in $schema) { $null = Invoke-PcmSql $connection $statement }
            if ($Kind -eq 'Snapshot') {
                $columns = @(Invoke-PcmSql $connection 'PRAGMA table_info(mailbox)' -As Rows)
                if ('resource_settings_json' -notin @($columns | ForEach-Object { $_.name })) {
                    $null = Invoke-PcmSql $connection 'ALTER TABLE mailbox ADD COLUMN resource_settings_json TEXT'
                }
            } else {
                $columns = @(Invoke-PcmSql $connection 'PRAGMA table_info(batch)' -As Rows | ForEach-Object { [string]$_.name })
                foreach ($column in @('snapshot_run_id', 'snapshot_hash', 'tenant_id')) {
                    if ($column -notin $columns) { $null = Invoke-PcmSql $connection ('ALTER TABLE batch ADD COLUMN {0} TEXT' -f $column) }
                }
            }
            $null = Invoke-PcmSql $connection "INSERT OR REPLACE INTO meta (key, value) VALUES ('schema_version', @v)" @{ v = [string]$expected }
            $null = Invoke-PcmSql $connection "INSERT OR REPLACE INTO meta (key, value) VALUES ('kind', @k)" @{ k = $Kind }
        } elseif ($null -eq $version -or $version -is [DBNull]) {
            throw "Not a PRA Cloud Mailbox database (no schema version): $Path"
        }
    } catch { $connection.Dispose(); throw }
    return $connection
}

function Close-PcmStore {
    [CmdletBinding()]
    param([AllowNull()][object]$Connection)
    if ($null -ne $Connection) { try { $Connection.Close() } finally { $Connection.Dispose() } }
}

function ConvertTo-PcmDbValue {
    <# PowerShell value -> SQLite parameter value (null, booleans as 0/1, dates in ISO 8601 UTC). #>
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return [DBNull]::Value }
    if ($Value -is [bool]) { return [int]$Value }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o') }
    if ($Value -is [guid]) { return $Value.ToString() }
    return $Value
}

function Invoke-PcmSql {
    <#
    .SYNOPSIS
        Runs one SQL statement with named parameters (@name). Values are always bound, never concatenated.
    .PARAMETER As
        NonQuery (rows changed), Scalar (first value) or Rows (PSCustomObject per row).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Sql, [hashtable]$Parameters = @{},
        [ValidateSet('NonQuery','Scalar','Rows')][string]$As = 'NonQuery', [object]$Transaction)
    $command = $Connection.CreateCommand()
    try {
        $command.CommandText = $Sql
        if ($Transaction) { $command.Transaction = $Transaction }
        foreach ($key in $Parameters.Keys) { [void]$command.Parameters.AddWithValue('@' + $key, (ConvertTo-PcmDbValue $Parameters[$key])) }
        switch ($As) {
            'NonQuery' { return $command.ExecuteNonQuery() }
            'Scalar' { return $command.ExecuteScalar() }
            'Rows' {
                $reader = $command.ExecuteReader()
                try {
                    $rows = [System.Collections.Generic.List[object]]::new()
                    while ($reader.Read()) {
                        $row = [ordered]@{}
                        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                            $value = $reader.GetValue($i)
                            $row[$reader.GetName($i)] = if ($value -is [DBNull]) { $null } else { $value }
                        }
                        [void]$rows.Add([pscustomobject]$row)
                    }
                    return $rows.ToArray()
                } finally { $reader.Dispose() }
            }
        }
    } finally { $command.Dispose() }
}

function New-PcmSnapshot {
    <#
    .SYNOPSIS
        Creates a snapshot row with the status Running and returns its ID.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][hashtable]$Info)
    if (-not $PSCmdlet.ShouldProcess('database', 'New snapshot')) { return 0 }
    $values = @{
        run_id = [string]$Info['RunId']; started = [datetime]::UtcNow; environment = [string]$Info['Environment']; computer = $env:COMPUTERNAME
        account = [string]$Info['Account']; tool_version = [string]$Info['ToolVersion']; exchange_server = [string]$Info['ExchangeServer']
        scope_json = [string]$Info['ScopeJson']
    }
    $null = Invoke-PcmSql $Connection @'
INSERT INTO snapshot (run_id, status, started_utc, environment, computer, account, tool_version, exchange_server, scope_json)
VALUES (@run_id, 'Running', @started, @environment, @computer, @account, @tool_version, @exchange_server, @scope_json)
'@ $values
    return [long](Invoke-PcmSql $Connection 'SELECT last_insert_rowid()' -As Scalar)
}

function Set-PcmSnapshotStatus {
    <# Marks a snapshot Complete or Failed, with its counters (JSON). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][long]$Id,
        [Parameter(Mandatory)][ValidateSet('Complete','Failed')][string]$Status, [hashtable]$Counts = @{}, [string]$Note = '')
    if (-not $PSCmdlet.ShouldProcess("snapshot $Id", "Set status $Status")) { return }
    $changed = Invoke-PcmSql $Connection 'UPDATE snapshot SET status = @s, finished_utc = @f, counts_json = @c, note = @n WHERE id = @id' @{
        s = $Status; f = [datetime]::UtcNow; c = ($Counts | ConvertTo-Json -Compress); n = $Note; id = $Id }
    if ($changed -ne 1) { throw "Snapshot $Id not found." }
}

function Add-PcmRow {
    <#
    .SYNOPSIS
        Inserts rows of one table in a single transaction. Every row gets snapshot_id = SnapshotId.
    .PARAMETER Rows
        Dictionaries whose keys are column names (all rows must have the same keys).
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateScript({ $_ -in $script:Tables })][string]$Table,
        [Parameter(Mandatory)][long]$SnapshotId, [AllowEmptyCollection()][object[]]$Rows)
    $list = @($Rows | Where-Object { $null -ne $_ })
    if (-not $list.Count) { return 0 }
    if (-not $PSCmdlet.ShouldProcess($Table, "Insert $($list.Count) row(s)")) { return 0 }
    $columns = @($list[0].Keys | ForEach-Object { [string]$_ })
    foreach ($column in $columns) { if ($column -notmatch '^[a-z_]+$' -or $column -eq 'snapshot_id') { throw "Invalid column name for $Table : $column" } }
    $sql = 'INSERT INTO {0} (snapshot_id, {1}) VALUES (@snapshot_id, {2})' -f $Table, ($columns -join ', '), (($columns | ForEach-Object { '@' + $_ }) -join ', ')
    $transaction = $Connection.BeginTransaction()
    $command = $Connection.CreateCommand()
    try {
        $command.Transaction = $transaction
        $command.CommandText = $sql
        [void]$command.Parameters.AddWithValue('@snapshot_id', $SnapshotId)
        $parameters = @{}
        foreach ($column in $columns) { $parameters[$column] = $command.Parameters.AddWithValue('@' + $column, [DBNull]::Value) }
        foreach ($row in $list) {
            if (@($row.Keys).Count -ne $columns.Count) { throw "Row with a different set of columns for $Table." }
            foreach ($column in $columns) { $parameters[$column].Value = ConvertTo-PcmDbValue $row[$column] }
            [void]$command.ExecuteNonQuery()
        }
        $transaction.Commit()
    } catch { $transaction.Rollback(); throw }
    finally { $command.Dispose(); $transaction.Dispose() }
    return $list.Count
}

function Get-PcmSnapshot {
    <#
    .SYNOPSIS
        One snapshot: by run (the run that wrote it, letter case ignored), by ID, or the last Complete one. Returns $null when there is none.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [long]$Id = 0, [string]$RunId = '')
    $rows = if ($RunId) { Invoke-PcmSql $Connection 'SELECT * FROM snapshot WHERE run_id = @r COLLATE NOCASE ORDER BY id DESC LIMIT 1' @{ r = $RunId } -As Rows }
        elseif ($Id -gt 0) { Invoke-PcmSql $Connection 'SELECT * FROM snapshot WHERE id = @id' @{ id = $Id } -As Rows }
        else { Invoke-PcmSql $Connection "SELECT * FROM snapshot WHERE status = 'Complete' ORDER BY id DESC LIMIT 1" -As Rows }
    $row = $rows | Select-Object -First 1
    if ($row) { return $row }
    return $null
}

function Get-PcmSnapshotList {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection)
    return Invoke-PcmSql $Connection 'SELECT id, status, started_utc, finished_utc, environment, computer, exchange_server, counts_json FROM snapshot ORDER BY id DESC' -As Rows
}

function Read-PcmTable {
    <# Every row of one table for a snapshot. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateScript({ $_ -in $script:Tables })][string]$Table,
        [Parameter(Mandatory)][long]$SnapshotId)
    return Invoke-PcmSql $Connection ("SELECT * FROM {0} WHERE snapshot_id = @id ORDER BY rowid" -f $Table) @{ id = $SnapshotId } -As Rows
}

function Test-PcmTable {
    <# Does the database hold this table? A snapshot written by an older version of the tool has no target_group and no target_member; a journal of an older version has no batch_run. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Table)
    return [int](Invoke-PcmSql $Connection "SELECT count(*) FROM sqlite_master WHERE type = 'table' AND name = @n" @{ n = $Table } -As Scalar) -gt 0
}

function Find-PcmTargetGroup {
    <#
    .SYNOPSIS
        The groups captured by Collect in a snapshot that a value names: name, sAMAccountName, distinguished name, SMTP address, display name or GUID,
        with the letter case ignored. Returns HasTables = $false for a snapshot written before the capture existed.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][long]$SnapshotId, [Parameter(Mandatory)][string]$Value)
    if (-not (Test-PcmTable $Connection 'target_group')) { return [pscustomobject]@{ HasTables = $false; Groups = @() } }
    $key = ([string]$Value).Trim()
    $guid = [guid]::Empty
    if ([guid]::TryParse($key, [ref]$guid)) { $key = $guid.ToString() }
    $wanted = $key.Normalize([Text.NormalizationForm]::FormC)
    $groups = @(Invoke-PcmSql $Connection 'SELECT * FROM target_group WHERE snapshot_id = @id ORDER BY rowid' @{ id = $SnapshotId } -As Rows | Where-Object {
            $group = $_
            foreach ($column in @('name', 'sam_account_name', 'distinguished_name', 'primary_smtp_address', 'display_name', 'group_guid')) {
                $text = [string]$group.$column
                if ($text -and [string]::Equals($text.Normalize([Text.NormalizationForm]::FormC), $wanted, [StringComparison]::OrdinalIgnoreCase)) { return $true }
            }
            $false
        })
    return [pscustomobject]@{ HasTables = $true; Groups = $groups }
}

function Remove-PcmOldSnapshot {
    <#
    .SYNOPSIS
        Keeps the last Keep Complete snapshots (and anything newer); deletes the older ones and their rows, except the snapshots in -Protect.
    .PARAMETER Protect
        Snapshots that a batch not yet rolled back still needs (Get-PcmSnapshotInUse): each has SnapshotId and RunId. A snapshot is kept when its run is named or, for an entry without a run, its ID.
    .PARAMETER Kept
        Receives the IDs of the old snapshots that were kept because of -Protect.
    .OUTPUTS
        Number of snapshots deleted.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateRange(1, 1000)][int]$Keep, [AllowEmptyCollection()][object[]]$Protect = @(), [ref]$Kept)
    $limit = Invoke-PcmSql $Connection "SELECT id FROM snapshot WHERE status = 'Complete' ORDER BY id DESC LIMIT 1 OFFSET @k" @{ k = ($Keep - 1) } -As Scalar
    if ($null -eq $limit -or $limit -is [DBNull]) { return 0 }
    if (-not $PSCmdlet.ShouldProcess('database', "Delete snapshots older than $limit")) { return 0 }
    $needed = [System.Collections.Generic.List[long]]::new()
    foreach ($old in @(Invoke-PcmSql $Connection 'SELECT id, run_id FROM snapshot WHERE id < @limit ORDER BY id' @{ limit = [long]$limit } -As Rows)) {
        foreach ($entry in @($Protect | Where-Object { $null -ne $_ })) {
            $run = [string](Get-PcmStoreValue $entry 'RunId' '')
            $id = [long](Get-PcmStoreValue $entry 'SnapshotId' 0)
            if (($run -and [string]::Equals($run, [string]$old.run_id, [StringComparison]::OrdinalIgnoreCase)) -or (-not $run -and $id -eq [long]$old.id)) { $needed.Add([long]$old.id); break }
        }
    }
    if ($null -ne $Kept) { $Kept.Value = $needed.ToArray() }
    if (-not $needed.Count) { return [int](Invoke-PcmSql $Connection 'DELETE FROM snapshot WHERE id < @limit' @{ limit = [long]$limit }) }
    $deleted = 0
    foreach ($old in @(Invoke-PcmSql $Connection 'SELECT id FROM snapshot WHERE id < @limit ORDER BY id' @{ limit = [long]$limit } -As Rows)) {
        if ([long]$old.id -in $needed) { continue }
        $deleted += [int](Invoke-PcmSql $Connection 'DELETE FROM snapshot WHERE id = @id' @{ id = [long]$old.id })
    }
    return $deleted
}

function Get-PcmStoreValue {
    <# One property of an object or one key of a dictionary, or the default when it is missing (a journal row of an older version lacks the newest columns). #>
    param([AllowNull()][object]$Object, [Parameter(Mandatory)][string]$Name, [AllowNull()][object]$Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } else { return $Default } }
    $property = $Object.PSObject.Properties[$Name]
    if ($property -and $null -ne $property.Value -and $property.Value -isnot [DBNull]) { return $property.Value }
    return $Default
}

function Test-PcmStoreColumn {
    <# Does the table hold this column? A journal written by an older version has no snapshot_run_id and no snapshot_hash. #>
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Table, [Parameter(Mandatory)][string]$Column)
    if ($Table -notmatch '^[a-z_]+\z') { throw "Invalid table name: $Table" }
    return @(Invoke-PcmSql $Connection ('PRAGMA table_info({0})' -f $Table) -As Rows | Where-Object { [string]::Equals([string]$_.name, $Column, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0
}

function Get-PcmSnapshotDigest {
    <#
    .SYNOPSIS
        SHA-256 of the content of one snapshot: the mailboxes, the permissions and the group members that Convert and Recover read, in the order of the database, without the snapshot number.
    .DESCRIPTION
        The hash of the database file changes with every Collect (the file holds several snapshots), so a batch keeps this digest instead: it proves that the snapshot found later is the one the batch was made from.
    .OUTPUTS
        64 lower-case hexadecimal characters.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowEmptyCollection()][object[]]$Mailbox = @(), [AllowEmptyCollection()][object[]]$Permission = @(), [AllowEmptyCollection()][object[]]$Member = @())
    $hash = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
    $utf8 = New-Object Text.UTF8Encoding($false)
    try {
        foreach ($set in @(@('mailbox', $Mailbox), @('permission', $Permission), @('group_member', $Member))) {
            $hash.AppendData($utf8.GetBytes(('[{0}]{1}' -f $set[0], [char]10)))
            $builder = [System.Text.StringBuilder]::new()
            foreach ($row in @($set[1] | Where-Object { $null -ne $_ })) {
                foreach ($property in $row.PSObject.Properties) {
                    if ($property.Name -eq 'snapshot_id') { continue }
                    $text = if ($null -eq $property.Value) { '' } else { [string]$property.Value }
                    [void]$builder.Append($property.Name).Append('=').Append($text).Append([char]0x1F)
                }
                [void]$builder.Append([char]0x1E)
                if ($builder.Length -gt 1MB) { $hash.AppendData($utf8.GetBytes($builder.ToString())); [void]$builder.Clear() }
            }
            $hash.AppendData($utf8.GetBytes($builder.ToString()))
        }
        return ([BitConverter]::ToString($hash.GetHashAndReset()) -replace '-', '').ToLowerInvariant()
    } finally { $hash.Dispose() }
}

function Get-PcmSnapshotInUse {
    <#
    .SYNOPSIS
        The snapshots that a Convert batch not yet rolled back still needs: Collect must not delete them.
    .DESCRIPTION
        A Convert batch is rolled back when every object that it converted has a Done item in a Recover batch of it. A Convert batch that is still running counts as not rolled back while it has no object yet or an object not finished (a resume reads the snapshot again).
        The journal is opened read-only; there is no journal yet when no batch was ever made. Readable is $false, with the reason in Message, when the journal exists but cannot be read:
        the caller must then delete nothing.
    .OUTPUTS
        Readable, Message, Items (SnapshotId, RunId, Batch).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$JournalPath, [Parameter(Mandatory)][string]$Root)
    $result = [pscustomobject]@{ Readable = $true; Message = ''; Items = @() }
    if (-not [IO.File]::Exists($JournalPath)) { return $result }
    $connection = $null
    try {
        $connection = Open-PcmStore -Path $JournalPath -Root $Root -ReadOnly -Kind Journal
        $run = if (Test-PcmStoreColumn $connection 'batch' 'snapshot_run_id') { 'b.snapshot_run_id' } else { 'NULL' }
        $result.Items = @(Invoke-PcmSql $connection @"
SELECT b.id AS batch, b.snapshot_id AS snapshot_id, $run AS run_id FROM batch b
WHERE b.action = 'Convert' AND b.snapshot_id IS NOT NULL AND (
    EXISTS (
        SELECT 1 FROM batch_item i WHERE i.batch_id = b.id AND i.status <> 'Skipped' AND i.step IS NOT NULL AND i.step <> '' AND NOT EXISTS (
            SELECT 1 FROM batch_item ri JOIN batch r ON r.id = ri.batch_id
            WHERE r.action = 'Recover' AND r.convert_batch = b.id COLLATE NOCASE AND ri.object_guid = i.object_guid COLLATE NOCASE AND ri.status = 'Done'))
    OR (b.status = 'Running' AND (NOT EXISTS (SELECT 1 FROM batch_item i WHERE i.batch_id = b.id)
        OR EXISTS (SELECT 1 FROM batch_item i WHERE i.batch_id = b.id AND i.status NOT IN ('Done', 'Skipped')))))
"@ -As Rows | ForEach-Object { [pscustomobject]@{ SnapshotId = [long]$_.snapshot_id; RunId = [string]$_.run_id; Batch = [string]$_.batch } })
    } catch {
        $result.Readable = $false
        $result.Message = $_.Exception.Message
    } finally { Close-PcmStore $connection }
    return $result
}

function Backup-PcmStore {
    <#
    .SYNOPSIS
        Writes a consistent copy of the database (VACUUM INTO) and returns its path.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name)
    if ($Name -notmatch '^[A-Za-z0-9_.-]+\z') { throw "Invalid backup file name: $Name" }
    Initialize-PcmPrivateFolder -Path $Folder
    $target = Join-Path $Folder $Name
    if (Test-Path -LiteralPath $target) { throw "Backup file already exists: $target" }
    if (-not $PSCmdlet.ShouldProcess($target, 'Backup database')) { return '' }
    $null = Invoke-PcmSql $Connection 'VACUUM INTO @t' @{ t = $target }
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw "Backup not written: $target" }
    return $target
}

#region Snapshot hash -----------------------------------------------------------------------------

function Get-PcmFileHash {
    <# SHA-256 of a file as 64 lower-case hexadecimal characters; the file may be open for writing by someone else. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)
    $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        $algorithm = [Security.Cryptography.SHA256]::Create()
        try { return ([BitConverter]::ToString($algorithm.ComputeHash($stream)) -replace '-', '').ToLowerInvariant() }
        finally { $algorithm.Dispose() }
    } finally { $stream.Dispose() }
}

function Write-PcmHashFile {
    <#
    .SYNOPSIS
        Writes <file>.sha256 beside a file: the hash, two spaces and the file name, like the transfer packages of the other tools. The file is written whole under a
        temporary name and then moved over the old one, so a reader never sees half a hash file.
    .OUTPUTS
        The path of the hash file.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$Path)
    $hash = Get-PcmFileHash -Path $Path
    $target = $Path + '.sha256'
    $temporary = '{0}.{1}.tmp' -f $target, [guid]::NewGuid().ToString('N').Substring(0, 8)
    try {
        [IO.File]::WriteAllBytes($temporary, (New-Object Text.UTF8Encoding($false)).GetBytes(('{0}  {1}{2}' -f $hash, [IO.Path]::GetFileName($Path), "`n")))
        [IO.File]::Move($temporary, $target, $true)
    } catch { try { [IO.File]::Delete($temporary) } catch { $null = $_ }; throw }
    return $target
}

function Remove-PcmHashFile {
    <# Deletes <file>.sha256 when it exists (a file that is about to change must not keep a hash that no longer fits it). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $target = $Path + '.sha256'
    if ([IO.File]::Exists($target)) { [IO.File]::Delete($target) }
}

function Read-PcmHashFile {
    <#
    .SYNOPSIS
        Reads a .sha256 file: Valid, Hash (lower case) and FileName (empty when the file gives none). Accepts '<hash>  <name>' and '<hash> *<name>', any letter case, any line ending, a
        byte order mark. A file that holds no 64-digit hexadecimal hash is not Valid.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $text = [IO.File]::ReadAllText($Path).TrimStart([char]0xFEFF).Trim()
    $line = @($text -split '\r?\n' | Where-Object { $_.Trim() }) | Select-Object -First 1
    $match = if ($line) { [regex]::Match([string]$line, '^\s*([0-9A-Fa-f]{64})(?:\s+\*?(.+?))?\s*$') } else { $null }
    if (-not $match -or -not $match.Success) { return [pscustomobject]@{ Valid = $false; Hash = ''; FileName = '' } }
    return [pscustomobject]@{ Valid = $true; Hash = $match.Groups[1].Value.ToLowerInvariant(); FileName = $match.Groups[2].Value.Trim() }
}

function Test-PcmSnapshotHash {
    <#
    .SYNOPSIS
        Does the snapshot file still match the hash that Collect wrote beside it?
    .DESCRIPTION
        State is Verified (the hash of the file is the one in <file>.sha256), NotVerifiable (there is no .sha256 beside the file: an old snapshot, or a copy made without it) or
        Changed (the hash differs, or the .sha256 holds no hash). Expected and Actual are the two hashes, HashFile the path of the .sha256, Message the plain sentence for the operator
        (the warning of NotVerifiable, the reason to stop of Changed). The file is only read. The journal is written by Convert and Recover themselves and is not covered.
    .PARAMETER BackupFolder
        Named in the message: the consistent copies that Collect writes there have a .sha256 of their own.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$BackupFolder = '')
    $name = [IO.Path]::GetFileName($Path)
    $hashFile = $Path + '.sha256'
    $copy = if ($BackupFolder) { " (the copy in $BackupFolder has its own .sha256)" } else { '' }
    $state = [ordered]@{ State = 'NotVerifiable'; Expected = ''; Actual = ''; HashFile = $hashFile; Message = '' }
    if (-not [IO.File]::Exists($Path)) {
        $state.Message = "The snapshot file $Path does not exist."
        return [pscustomobject]$state
    }
    if (-not [IO.File]::Exists($hashFile)) {
        $state.Message = "The snapshot file $Path has no $name.sha256 beside it, so the tool cannot prove it is the file written by Collect (old snapshot, or the file was copied without it). Copy the .sha256 file with it next time."
        return [pscustomobject]$state
    }
    $read = $null
    try { $read = Read-PcmHashFile -Path $hashFile } catch { $read = [pscustomobject]@{ Valid = $false; Hash = ''; FileName = '' } }
    if (-not $read.Valid) {
        $state.State = 'Changed'
        $state.Message = "The snapshot file $Path cannot be verified: $name.sha256 beside it does not hold a SHA-256 hash. Nothing was changed. Restore the .sha256 file written by Collect$copy, or, if you accept the snapshot as it is, run the same command again with -AcceptUnverifiedSnapshot."
        return [pscustomobject]$state
    }
    $state.Expected = $read.Hash
    $state.Actual = Get-PcmFileHash -Path $Path
    if ([string]::Equals($state.Actual, $state.Expected, [StringComparison]::OrdinalIgnoreCase)) { $state.State = 'Verified'; return [pscustomobject]$state }
    $state.State = 'Changed'
    $state.Message = "The snapshot file $Path was changed after its last Collect: its SHA-256 is $($state.Actual) and the file $name.sha256 written by Collect says $($state.Expected). Nothing was changed. Restore the file from your external backup$copy, or, if you know why it changed and accept it, run the same command again with -AcceptUnverifiedSnapshot."
    return [pscustomobject]$state
}
#endregion

#region Lists written by a person ----------------------------------------------------------------

function Get-PcmCsvDelimiter {
    <# The delimiter of a header line that holds the column Identity (comma or semicolon, the column anywhere, letter case and quotes ignored): ',' or ';'. '' when the line has no such column. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([AllowNull()][AllowEmptyString()][string]$Header, [string]$Column = 'Identity')
    $name = [regex]::Escape($Column)
    if ($Header -match ('(?i)(^|,)\s*"?{0}"?\s*(,|\z)' -f $name)) { return ',' }
    if ($Header -match ('(?i)(^|;)\s*"?{0}"?\s*(;|\z)' -f $name)) { return ';' }
    return ''
}
#endregion

#region Journal (cloud actions) -------------------------------------------------------------------

function New-PcmBatchId {
    <# Short batch ID shown to the operator (8 hex characters). #>
    return [guid]::NewGuid().ToString('N').Substring(0, 8)
}

function New-PcmBatch {
    <#
    .SYNOPSIS
        Creates a batch (Convert or Recover) with the status Running and returns its ID.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateSet('Convert','Recover')][string]$Action, [hashtable]$Info = @{})
    $id = New-PcmBatchId
    if (-not $PSCmdlet.ShouldProcess("batch $id", "New $Action batch")) { return '' }
    $null = Invoke-PcmSql $Connection @'
INSERT INTO batch (id, action, convert_batch, status, created_utc, updated_utc, snapshot_id, environment, computer, account, tool_version, scope_json, snapshot_run_id, snapshot_hash, tenant_id)
VALUES (@id, @action, @convert, 'Running', @now, @now, @snapshot, @environment, @computer, @account, @tool, @scope, @run, @hash, @tenant)
'@ @{ id = $id; action = $Action; convert = [string]$Info['ConvertBatch']; now = [datetime]::UtcNow; snapshot = $Info['SnapshotId']
        environment = [string]$Info['Environment']; computer = $env:COMPUTERNAME; account = [string]$Info['Account']; tool = [string]$Info['ToolVersion']
        scope = [string]$Info['ScopeJson']; run = [string]$Info['SnapshotRunId']; hash = [string]$Info['SnapshotHash']; tenant = [string]$Info['TenantId'] }
    return $id
}

function Set-PcmBatch {
    <# Status and counters of a batch. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][ValidateSet('Running','Complete','Partial','Failed')][string]$Status, [hashtable]$Counts = @{}, [string]$Note = '')
    if (-not $PSCmdlet.ShouldProcess("batch $Id", "Set status $Status")) { return }
    $changed = Invoke-PcmSql $Connection 'UPDATE batch SET status = @s, updated_utc = @u, counts_json = @c, note = @n WHERE id = @id COLLATE NOCASE' @{
        s = $Status; u = [datetime]::UtcNow; c = ($Counts | ConvertTo-Json -Compress); n = $Note; id = $Id }
    if ($changed -ne 1) { throw "Batch $Id not found." }
}

function Get-PcmBatch {
    <#
    .SYNOPSIS
        One batch by ID (the full 8 characters, or a unique prefix of 4 or more), or the last batch of an action.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [string]$Id, [ValidateSet('Convert','Recover')][string]$Action)
    if ($Id) {
        if ($Id -notmatch '^[0-9a-fA-F]{4,8}\z') { throw "Invalid batch ID: $Id (8 hexadecimal characters, printed at the end of Convert)." }
        $rows = @(Invoke-PcmSql $Connection 'SELECT * FROM batch WHERE id LIKE @p ORDER BY created_utc DESC' @{ p = ($Id.ToLowerInvariant() + '%') } -As Rows)
        if ($rows.Count -gt 1) { throw "Batch ID $Id is ambiguous: give more characters." }
        if ($rows.Count -eq 1) { return $rows[0] }
        return $null
    }
    $filter = if ($Action) { 'WHERE action = @a' } else { '' }
    $rows = @(Invoke-PcmSql $Connection ("SELECT * FROM batch {0} ORDER BY created_utc DESC LIMIT 1" -f $filter) @{ a = $Action } -As Rows)
    if ($rows.Count) { return $rows[0] }
    return $null
}

function Set-PcmBatchItem {
    <#
    .SYNOPSIS
        Creates or updates the journal row of one object. Only the given values change.
    .PARAMETER Values
        Keys among kind, identity, entra_id, step, status, original_json, result_json, message.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$ObjectGuid,
        [Parameter(Mandatory)][hashtable]$Values)
    $allowed = @('kind','identity','entra_id','step','status','original_json','result_json','message')
    foreach ($key in $Values.Keys) { if ($key -notin $allowed) { throw "Invalid journal column: $key" } }
    if (-not $PSCmdlet.ShouldProcess("$BatchId/$ObjectGuid", 'Journal')) { return }
    $exists = [int](Invoke-PcmSql $Connection 'SELECT count(*) FROM batch_item WHERE batch_id = @b COLLATE NOCASE AND object_guid = @o COLLATE NOCASE' @{ b = $BatchId; o = $ObjectGuid } -As Scalar)
    $parameters = @{ b = $BatchId; o = $ObjectGuid; u = [datetime]::UtcNow }
    foreach ($key in $Values.Keys) { $parameters[$key] = $Values[$key] }
    if ($exists) {
        $sets = @($Values.Keys | ForEach-Object { '{0} = @{0}' -f $_ }) + 'updated_utc = @u'
        $null = Invoke-PcmSql $Connection ('UPDATE batch_item SET {0} WHERE batch_id = @b COLLATE NOCASE AND object_guid = @o COLLATE NOCASE' -f ($sets -join ', ')) $parameters
    } else {
        if (-not $Values.ContainsKey('kind') -or -not $Values.ContainsKey('status')) { throw 'A new journal row needs kind and status.' }
        $columns = @($Values.Keys)
        $null = Invoke-PcmSql $Connection ('INSERT INTO batch_item (batch_id, object_guid, {0}, updated_utc) VALUES (@b, @o, {1}, @u)' -f ($columns -join ', '), (($columns | ForEach-Object { '@' + $_ }) -join ', ')) $parameters
    }
}

function Get-PcmBatchItem {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId)
    return Invoke-PcmSql $Connection 'SELECT * FROM batch_item WHERE batch_id = @b COLLATE NOCASE ORDER BY kind DESC, identity' @{ b = $BatchId } -As Rows
}

function Add-PcmBatchEvent {
    <# One line of the batch history (every cloud change and every wait result). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [string]$ObjectGuid = '',
        [Parameter(Mandatory)][string]$Step, [ValidateSet('Ok','Fail','Info')][string]$Outcome = 'Ok', [string]$Detail = '')
    if (-not $PSCmdlet.ShouldProcess($BatchId, "Event $Step")) { return }
    $null = Invoke-PcmSql $Connection 'INSERT INTO batch_event (batch_id, object_guid, time_utc, step, outcome, detail) VALUES (@b, @o, @t, @s, @r, @d)' @{
        b = $BatchId; o = $ObjectGuid; t = [datetime]::UtcNow; s = $Step; r = $Outcome; d = $Detail }
}

function Set-PcmTenantChange {
    <# Tenant-level change and restoration. UpdateOriginal explicitly advances a journaled room-list creation intent; legacy originals stay immutable. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$TargetId, [string]$OriginalJson, [switch]$Restored, [switch]$UpdateOriginal)
    if (-not $PSCmdlet.ShouldProcess("$Kind $TargetId", 'Journal')) { return }
    $exists = [int](Invoke-PcmSql $Connection 'SELECT count(*) FROM tenant_change WHERE batch_id = @b COLLATE NOCASE AND kind = @k AND target_id = @t COLLATE NOCASE' @{ b = $BatchId; k = $Kind; t = $TargetId } -As Scalar)
    if ($exists) {
        $setOriginal = if ($UpdateOriginal) { ', original_json = @o' } else { '' }
        $null = Invoke-PcmSql $Connection ("UPDATE tenant_change SET restored = @r, updated_utc = @u{0} WHERE batch_id = @b COLLATE NOCASE AND kind = @k AND target_id = @t COLLATE NOCASE" -f $setOriginal) @{
            r = [bool]$Restored; u = [datetime]::UtcNow; b = $BatchId; k = $Kind; t = $TargetId; o = $OriginalJson }
    }
    else { $null = Invoke-PcmSql $Connection 'INSERT INTO tenant_change (batch_id, kind, target_id, original_json, restored, updated_utc) VALUES (@b, @k, @t, @o, @r, @u)' @{ b = $BatchId; k = $Kind; t = $TargetId; o = $OriginalJson; r = [bool]$Restored; u = [datetime]::UtcNow } }
}

function Get-PcmTenantChange {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId)
    return Invoke-PcmSql $Connection 'SELECT * FROM tenant_change WHERE batch_id = @b COLLATE NOCASE' @{ b = $BatchId } -As Rows
}

function Add-PcmBatchRun {
    <# One run of a batch (Convert, a resume of Convert, a Recover wave): when, who, with which source and settings, and how it ended. Written by the run itself, at its end. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][hashtable]$Run)
    if (-not $PSCmdlet.ShouldProcess($BatchId, 'Journal run')) { return }
    $null = Invoke-PcmSql $Connection @'
INSERT OR REPLACE INTO batch_run (run_id, batch_id, action, mode, started_utc, ended_utc, account, computer, tool_version, source_text, overrides_text, status, exit_code,
    done_count, pending_count, error_count, skipped_count, log_file)
VALUES (@run_id, @batch_id, @action, @mode, @started, @ended, @account, @computer, @tool, @source, @overrides, @status, @exit, @done, @pending, @errors, @skipped, @log)
'@ @{ run_id = [string]$Run['RunId']; batch_id = $BatchId; action = [string]$Run['Action']; mode = [string]$Run['Mode']; started = [datetime]$Run['Started']; ended = [datetime]::UtcNow
        account = [string]$Run['Account']; computer = $env:COMPUTERNAME; tool = [string]$Run['ToolVersion']; source = [string]$Run['Source']; overrides = [string]$Run['Overrides']
        status = [string]$Run['Status']; exit = [int]$Run['ExitCode']; done = [int]$Run['Done']; pending = [int]$Run['Pending']; errors = [int]$Run['Errors']; skipped = [int]$Run['Skipped']; log = [string]$Run['LogFile'] }
}

function Get-PcmBatchHistory {
    <#
    .SYNOPSIS
        Everything the journal knows about one Convert batch and the Recover batches that rolled it back: the batches, their objects, their events and their runs.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$ConvertBatchId)
    $batches = @(Invoke-PcmSql $Connection 'SELECT * FROM batch WHERE id = @c COLLATE NOCASE OR convert_batch = @c COLLATE NOCASE ORDER BY created_utc, id' @{ c = $ConvertBatchId } -As Rows)
    $items = [System.Collections.Generic.List[object]]::new(); $events = [System.Collections.Generic.List[object]]::new(); $runs = [System.Collections.Generic.List[object]]::new()
    $hasRuns = Test-PcmTable $Connection 'batch_run'
    foreach ($batch in $batches) {
        foreach ($row in @(Invoke-PcmSql $Connection 'SELECT * FROM batch_item WHERE batch_id = @b ORDER BY identity' @{ b = $batch.id } -As Rows)) { $items.Add($row) }
        foreach ($row in @(Invoke-PcmSql $Connection 'SELECT rowid AS seq, * FROM batch_event WHERE batch_id = @b ORDER BY time_utc, rowid' @{ b = $batch.id } -As Rows)) { $events.Add($row) }
        if ($hasRuns) { foreach ($row in @(Invoke-PcmSql $Connection 'SELECT * FROM batch_run WHERE batch_id = @b ORDER BY started_utc' @{ b = $batch.id } -As Rows)) { $runs.Add($row) } }
    }
    return [pscustomobject]@{ Batches = $batches; Items = $items.ToArray(); Events = $events.ToArray(); Runs = $runs.ToArray() }
}
#endregion

Export-ModuleMember -Function Import-PcmSqlite, Open-PcmStore, Close-PcmStore, Invoke-PcmSql, New-PcmSnapshot, Set-PcmSnapshotStatus,
    Add-PcmRow, Get-PcmSnapshot, Get-PcmSnapshotList, Read-PcmTable, Test-PcmTable, Find-PcmTargetGroup, Remove-PcmOldSnapshot, Get-PcmSnapshotDigest, Get-PcmSnapshotInUse, Backup-PcmStore, Get-PcmFileHash, Write-PcmHashFile, Remove-PcmHashFile, Read-PcmHashFile, Test-PcmSnapshotHash,
    Get-PcmCsvDelimiter,
    New-PcmBatch, Set-PcmBatch, Get-PcmBatch, Set-PcmBatchItem, Get-PcmBatchItem, Add-PcmBatchEvent, Set-PcmTenantChange, Get-PcmTenantChange, Add-PcmBatchRun, Get-PcmBatchHistory

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDX2WGjsxrXj08p
# 0dsmyY66hhgVo2wcZt9V2dY9m21xZaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC5wIEk2kIy
# JRYSL2JqnAhFJnbwFVdQAAlScPi8K4A0DzANBgkqhkiG9w0BAQEFAASCAgCXB5f1
# 2lwRDBJpm0kCq20KGC8rHcdpEXPTIjduf456c/9cWTZUYNw+jxqFUSF+/DfJt14/
# Uf0ElRJsnbCX3FFFdrFR23Ynm4NiGHvs1O32gCr5hVl9CNx0cUr37P3701SRFlND
# f9Lykv4Uq0O40q9S4jtKtRsQyBJfYJ1nU68iqqNkw4rv9aujT5pgUNrqEKYpnJe9
# eBuYAjA8/7iJDFH98T83rqU6xYkz8S8ddc1X0N1yASiJIgkzb4nbgbXFYcR1hT3L
# LNNYXe13uUf47ZPO8BER7X8DxVZ2XEtNHefrJ5jrq3re5fXNn39G7lJl15T+FQtr
# 8o7jqgik1/LoRF5AsbUyB+RjP8cWRH9dmhIc7SGeW6NEEhkqy5rnus/NYg02QsnR
# 3y8tEIGXsDgh9R+mKRZhjVmpL5bqQH7OBejyBFgs1/JNib210s5h4O1mtVqBpBkn
# UuAlaYUjtcD5Fp3K9FRawYpy/Cc7qX3GeWRNT31Ue7IdGNnkrZL2LOQCN2/7fqBw
# mLyuE+1wpscc1i3jH89rIIXJ6d6tqxsRVUnpJ5627MydfsUQEemB4gSTEYNaVw70
# dmjF8sgJIUiQ/rhQQlb3wVtCUkPwyNGJXKIUpFGGLUyqEpD+UGjueaeuV/OdM28j
# gRw9BfaHj0ARoaEp9PZAwcK/kgqm01qNpWNEt6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEICDfaTrkwVMmqeeA+pCAOaNdDZnK5bYqBemUX5qloUMdAggap78/SksI
# JBgPMjAyNjExMTUyMDA2NDJaMAMCAQECEDMFP+ie/TKRkbNfaPa23u2gghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTIwMDY0MlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIKJ8
# jGR5D1hsLQvof9U+ZGE/yr7x3QE9hyQYKzZCEY2gMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAgjblT769UokVphf80VA6+REJqQIab06UVHx63Nk/h6kCIHLz
# khA/0PdVEY3y46gppgwdTFMWdE5iRtM7Wu1KDpD/
# SIG # End signature block
