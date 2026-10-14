<#
.SYNOPSIS
    PRA Cloud Mailbox - storage module: the SQLite database that keeps the Collect snapshots.

.DESCRIPTION
    Collect (Windows PowerShell 5.1, Exchange on-premises) writes one snapshot per run: the mailboxes in
    scope, the shared mailbox permissions, and optionally the contacts and distribution groups. The cloud
    actions (PowerShell 7) read the last complete snapshot, when Exchange and Active Directory are gone.

    One file, one engine for both editions: System.Data.SQLite 1.0.119 (SQLite 3.46.1, public domain),
    net46 build for Windows PowerShell 5.1 and netstandard2.0 build for PowerShell 7, with the same
    native x64 SQLite.Interop.dll (lib\sqlite).

    A snapshot is written inside its own rows with the status Running, then marked Complete at the end:
    an interrupted Collect never replaces the last good snapshot. Old snapshots are pruned (Store.KeepSnapshots)
    and every Apply writes a consistent copy of the database (VACUUM INTO) to Store.BackupFolder.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.0
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SchemaVersion = 2
$script:SqliteFlavor = ''

# Snapshot schema 2 adds optional resource settings; schema 1 remains readable.
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
    'CREATE INDEX IF NOT EXISTS ix_permission ON permission (snapshot_id, mailbox_guid)'
    'CREATE INDEX IF NOT EXISTS ix_group_member ON group_member (snapshot_id, group_guid)'
    'CREATE INDEX IF NOT EXISTS ix_dl_member ON dl_member (snapshot_id, group_guid)'
)
$script:Tables = @('mailbox','permission','group_member','contact','distribution_group','dl_member','collect_issue')

# Journal of the cloud actions (Convert, Recover): a separate file on the cloud admin server, so that a new
# Collect copied over the snapshot database never erases a batch. Every step of every object is recorded:
# a run interrupted (session cut, reboot) is resumed with the same -Batch.
$script:JournalSchemaVersion = 1
$script:JournalSchema = @(
    'CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)'
    @'
CREATE TABLE IF NOT EXISTS batch (
    id TEXT PRIMARY KEY, action TEXT NOT NULL CHECK (action IN ('Convert','Recover')), convert_batch TEXT,
    status TEXT NOT NULL CHECK (status IN ('Running','Complete','Partial','Failed')),
    created_utc TEXT NOT NULL, updated_utc TEXT, snapshot_id INTEGER, environment TEXT, computer TEXT, account TEXT,
    tool_version TEXT, scope_json TEXT, counts_json TEXT, note TEXT)
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
    'CREATE INDEX IF NOT EXISTS ix_batch_event ON batch_event (batch_id, object_guid)'
)

function Import-Pra2Sqlite {
    <#
    .SYNOPSIS
        Loads System.Data.SQLite for this edition (once per process).
    .PARAMETER Root
        Tool folder (holds lib\sqlite).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Root)
    if ($script:SqliteFlavor) { return $script:SqliteFlavor }
    if (-not [Environment]::Is64BitProcess) { throw 'The SQLite engine of the tool is 64-bit: run a 64-bit PowerShell.' }
    $flavor = if ($PSVersionTable.PSEdition -eq 'Desktop') { 'net46' } else { 'netstandard2.0' }
    $already = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.Data.SQLite' } | Select-Object -First 1
    if (-not $already) {
        $dll = Join-Path $Root "lib\sqlite\$flavor\System.Data.SQLite.dll"
        if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) { throw "SQLite library missing: $dll (copy the whole lib folder of the tool)." }
        Add-Type -LiteralPath $dll
    }
    $script:SqliteFlavor = $flavor
    return $flavor
}

function Open-Pra2Store {
    <#
    .SYNOPSIS
        Opens (and creates, unless -ReadOnly) the SQLite database, and checks or creates the schema.
    .OUTPUTS
        System.Data.SQLite.SQLiteConnection (open). Close it with Close-Pra2Store.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root, [switch]$ReadOnly, [ValidateSet('Snapshot','Journal')][string]$Kind = 'Snapshot')
    $null = Import-Pra2Sqlite -Root $Root
    if ($ReadOnly -and -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw $(if ($Kind -eq 'Journal') { "Journal not found: $Path (it is written by Convert)." } else { "Database not found: $Path. Run -Action Collect -Mode Apply on the Exchange side first, then copy the file here." })
    }
    $schema = if ($Kind -eq 'Journal') { $script:JournalSchema } else { $script:Schema }
    $expected = if ($Kind -eq 'Journal') { $script:JournalSchemaVersion } else { $script:SchemaVersion }
    $folder = Split-Path -Parent $Path
    if (-not $ReadOnly -and $folder) { $null = New-Item -ItemType Directory -Path $folder -Force -WhatIf:$false -Confirm:$false }
    $builder = New-Object System.Data.SQLite.SQLiteConnectionStringBuilder
    $builder.DataSource = $Path
    $builder.ForeignKeys = $true
    $builder.FailIfMissing = [bool]$ReadOnly
    $builder.ReadOnly = [bool]$ReadOnly
    $connection = New-Object System.Data.SQLite.SQLiteConnection($builder.ToString())
    $connection.Open()
    try {
        $version = $null; $storedKind = $null
        $hasMeta = Invoke-Pra2Sql $connection "SELECT count(*) FROM sqlite_master WHERE type = 'table' AND name = 'meta'" -As Scalar
        if ([int]$hasMeta -gt 0) {
            $version = Invoke-Pra2Sql $connection "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar
            $storedKind = Invoke-Pra2Sql $connection "SELECT value FROM meta WHERE key = 'kind'" -As Scalar
        }
        $hasVersion = $null -ne $version -and $version -isnot [DBNull]
        $actualKind = if ($null -ne $storedKind -and $storedKind -isnot [DBNull]) { [string]$storedKind } elseif ($hasVersion) { 'Snapshot' } else { '' }
        if ($actualKind -and $actualKind -ne $Kind) { throw "$Path is a $($actualKind.ToLowerInvariant()) database, not a $($Kind.ToLowerInvariant()) database." }
        if ($hasVersion -and [int]$version -gt $expected) {
            throw "Database schema version $version is newer than this tool (version $expected): use the same tool version as the Collect."
        }
        if (-not $ReadOnly) {
            foreach ($statement in $schema) { $null = Invoke-Pra2Sql $connection $statement }
            if ($Kind -eq 'Snapshot') {
                $columns = @(Invoke-Pra2Sql $connection 'PRAGMA table_info(mailbox)' -As Rows)
                if ('resource_settings_json' -notin @($columns | ForEach-Object { $_.name })) {
                    $null = Invoke-Pra2Sql $connection 'ALTER TABLE mailbox ADD COLUMN resource_settings_json TEXT'
                }
            }
            $null = Invoke-Pra2Sql $connection "INSERT OR REPLACE INTO meta (key, value) VALUES ('schema_version', @v)" @{ v = [string]$expected }
            $null = Invoke-Pra2Sql $connection "INSERT OR REPLACE INTO meta (key, value) VALUES ('kind', @k)" @{ k = $Kind }
        } elseif ($null -eq $version -or $version -is [DBNull]) {
            throw "Not a PRA Cloud Mailbox database (no schema version): $Path"
        }
    } catch { $connection.Dispose(); throw }
    return $connection
}

function Close-Pra2Store {
    [CmdletBinding()]
    param([AllowNull()][object]$Connection)
    if ($null -ne $Connection) { try { $Connection.Close() } finally { $Connection.Dispose() } }
}

function ConvertTo-Pra2DbValue {
    <# PowerShell value -> SQLite parameter value (null, booleans as 0/1, dates in ISO 8601 UTC). #>
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return [DBNull]::Value }
    if ($Value -is [bool]) { return [int]$Value }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o') }
    if ($Value -is [guid]) { return $Value.ToString() }
    return $Value
}

function Invoke-Pra2Sql {
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
        foreach ($key in $Parameters.Keys) { [void]$command.Parameters.AddWithValue('@' + $key, (ConvertTo-Pra2DbValue $Parameters[$key])) }
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

function New-Pra2Snapshot {
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
    $null = Invoke-Pra2Sql $Connection @'
INSERT INTO snapshot (run_id, status, started_utc, environment, computer, account, tool_version, exchange_server, scope_json)
VALUES (@run_id, 'Running', @started, @environment, @computer, @account, @tool_version, @exchange_server, @scope_json)
'@ $values
    return [long](Invoke-Pra2Sql $Connection 'SELECT last_insert_rowid()' -As Scalar)
}

function Set-Pra2SnapshotStatus {
    <# Marks a snapshot Complete or Failed, with its counters (JSON). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][long]$Id,
        [Parameter(Mandatory)][ValidateSet('Complete','Failed')][string]$Status, [hashtable]$Counts = @{}, [string]$Note = '')
    if (-not $PSCmdlet.ShouldProcess("snapshot $Id", "Set status $Status")) { return }
    $changed = Invoke-Pra2Sql $Connection 'UPDATE snapshot SET status = @s, finished_utc = @f, counts_json = @c, note = @n WHERE id = @id' @{
        s = $Status; f = [datetime]::UtcNow; c = ($Counts | ConvertTo-Json -Compress); n = $Note; id = $Id }
    if ($changed -ne 1) { throw "Snapshot $Id not found." }
}

function Add-Pra2Row {
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
            foreach ($column in $columns) { $parameters[$column].Value = ConvertTo-Pra2DbValue $row[$column] }
            [void]$command.ExecuteNonQuery()
        }
        $transaction.Commit()
    } catch { $transaction.Rollback(); throw }
    finally { $command.Dispose(); $transaction.Dispose() }
    return $list.Count
}

function Get-Pra2Snapshot {
    <#
    .SYNOPSIS
        One snapshot: by ID, or the last Complete one. Returns $null when there is none.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [long]$Id = 0)
    $rows = if ($Id -gt 0) { Invoke-Pra2Sql $Connection 'SELECT * FROM snapshot WHERE id = @id' @{ id = $Id } -As Rows }
        else { Invoke-Pra2Sql $Connection "SELECT * FROM snapshot WHERE status = 'Complete' ORDER BY id DESC LIMIT 1" -As Rows }
    $row = $rows | Select-Object -First 1
    if ($row) { return $row }
    return $null
}

function Get-Pra2SnapshotList {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection)
    return Invoke-Pra2Sql $Connection 'SELECT id, status, started_utc, finished_utc, environment, computer, exchange_server, counts_json FROM snapshot ORDER BY id DESC' -As Rows
}

function Read-Pra2Table {
    <# Every row of one table for a snapshot. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateScript({ $_ -in $script:Tables })][string]$Table,
        [Parameter(Mandatory)][long]$SnapshotId)
    return Invoke-Pra2Sql $Connection ("SELECT * FROM {0} WHERE snapshot_id = @id ORDER BY rowid" -f $Table) @{ id = $SnapshotId } -As Rows
}

function Remove-Pra2OldSnapshot {
    <#
    .SYNOPSIS
        Keeps the last Keep Complete snapshots (and anything newer); deletes the older ones and their rows.
    .OUTPUTS
        Number of snapshots deleted.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateRange(1, 1000)][int]$Keep)
    $limit = Invoke-Pra2Sql $Connection "SELECT id FROM snapshot WHERE status = 'Complete' ORDER BY id DESC LIMIT 1 OFFSET @k" @{ k = ($Keep - 1) } -As Scalar
    if ($null -eq $limit -or $limit -is [DBNull]) { return 0 }
    if (-not $PSCmdlet.ShouldProcess('database', "Delete snapshots older than $limit")) { return 0 }
    return [int](Invoke-Pra2Sql $Connection 'DELETE FROM snapshot WHERE id < @limit' @{ limit = [long]$limit })
}

function Backup-Pra2Store {
    <#
    .SYNOPSIS
        Writes a consistent copy of the database (VACUUM INTO) and returns its path.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name)
    if ($Name -notmatch '^[A-Za-z0-9_.-]+$') { throw "Invalid backup file name: $Name" }
    $null = New-Item -ItemType Directory -Path $Folder -Force -WhatIf:$false -Confirm:$false
    $target = Join-Path $Folder $Name
    if (Test-Path -LiteralPath $target) { throw "Backup file already exists: $target" }
    if (-not $PSCmdlet.ShouldProcess($target, 'Backup database')) { return '' }
    $null = Invoke-Pra2Sql $Connection 'VACUUM INTO @t' @{ t = $target }
    if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { throw "Backup not written: $target" }
    return $target
}

#region Journal (cloud actions) -------------------------------------------------------------------

function New-Pra2BatchId {
    <# Short batch ID shown to the operator (8 hex characters). #>
    return [guid]::NewGuid().ToString('N').Substring(0, 8)
}

function New-Pra2Batch {
    <#
    .SYNOPSIS
        Creates a batch (Convert or Recover) with the status Running and returns its ID.
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][ValidateSet('Convert','Recover')][string]$Action, [hashtable]$Info = @{})
    $id = New-Pra2BatchId
    if (-not $PSCmdlet.ShouldProcess("batch $id", "New $Action batch")) { return '' }
    $null = Invoke-Pra2Sql $Connection @'
INSERT INTO batch (id, action, convert_batch, status, created_utc, updated_utc, snapshot_id, environment, computer, account, tool_version, scope_json)
VALUES (@id, @action, @convert, 'Running', @now, @now, @snapshot, @environment, @computer, @account, @tool, @scope)
'@ @{ id = $id; action = $Action; convert = [string]$Info['ConvertBatch']; now = [datetime]::UtcNow; snapshot = $Info['SnapshotId']
        environment = [string]$Info['Environment']; computer = $env:COMPUTERNAME; account = [string]$Info['Account']; tool = [string]$Info['ToolVersion']
        scope = [string]$Info['ScopeJson'] }
    return $id
}

function Set-Pra2Batch {
    <# Status and counters of a batch. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][ValidateSet('Running','Complete','Partial','Failed')][string]$Status, [hashtable]$Counts = @{}, [string]$Note = '')
    if (-not $PSCmdlet.ShouldProcess("batch $Id", "Set status $Status")) { return }
    $changed = Invoke-Pra2Sql $Connection 'UPDATE batch SET status = @s, updated_utc = @u, counts_json = @c, note = @n WHERE id = @id' @{
        s = $Status; u = [datetime]::UtcNow; c = ($Counts | ConvertTo-Json -Compress); n = $Note; id = $Id }
    if ($changed -ne 1) { throw "Batch $Id not found." }
}

function Get-Pra2Batch {
    <#
    .SYNOPSIS
        One batch by ID (the full 8 characters, or a unique prefix of 4 or more), or the last batch of an action.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [string]$Id, [ValidateSet('Convert','Recover')][string]$Action)
    if ($Id) {
        if ($Id -notmatch '^[0-9a-fA-F]{4,8}$') { throw "Invalid batch ID: $Id (8 hexadecimal characters, printed at the end of Convert)." }
        $rows = @(Invoke-Pra2Sql $Connection 'SELECT * FROM batch WHERE id LIKE @p ORDER BY created_utc DESC' @{ p = ($Id.ToLowerInvariant() + '%') } -As Rows)
        if ($rows.Count -gt 1) { throw "Batch ID $Id is ambiguous: give more characters." }
        if ($rows.Count -eq 1) { return $rows[0] }
        return $null
    }
    $filter = if ($Action) { 'WHERE action = @a' } else { '' }
    $rows = @(Invoke-Pra2Sql $Connection ("SELECT * FROM batch {0} ORDER BY created_utc DESC LIMIT 1" -f $filter) @{ a = $Action } -As Rows)
    if ($rows.Count) { return $rows[0] }
    return $null
}

function Set-Pra2BatchItem {
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
    $exists = [int](Invoke-Pra2Sql $Connection 'SELECT count(*) FROM batch_item WHERE batch_id = @b AND object_guid = @o' @{ b = $BatchId; o = $ObjectGuid } -As Scalar)
    $parameters = @{ b = $BatchId; o = $ObjectGuid; u = [datetime]::UtcNow }
    foreach ($key in $Values.Keys) { $parameters[$key] = $Values[$key] }
    if ($exists) {
        $sets = @($Values.Keys | ForEach-Object { '{0} = @{0}' -f $_ }) + 'updated_utc = @u'
        $null = Invoke-Pra2Sql $Connection ('UPDATE batch_item SET {0} WHERE batch_id = @b AND object_guid = @o' -f ($sets -join ', ')) $parameters
    } else {
        if (-not $Values.ContainsKey('kind') -or -not $Values.ContainsKey('status')) { throw 'A new journal row needs kind and status.' }
        $columns = @($Values.Keys)
        $null = Invoke-Pra2Sql $Connection ('INSERT INTO batch_item (batch_id, object_guid, {0}, updated_utc) VALUES (@b, @o, {1}, @u)' -f ($columns -join ', '), (($columns | ForEach-Object { '@' + $_ }) -join ', ')) $parameters
    }
}

function Get-Pra2BatchItem {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId)
    return Invoke-Pra2Sql $Connection 'SELECT * FROM batch_item WHERE batch_id = @b ORDER BY kind DESC, identity' @{ b = $BatchId } -As Rows
}

function Add-Pra2BatchEvent {
    <# One line of the batch history (every cloud change and every wait result). #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [string]$ObjectGuid = '',
        [Parameter(Mandatory)][string]$Step, [ValidateSet('Ok','Fail','Info')][string]$Outcome = 'Ok', [string]$Detail = '')
    if (-not $PSCmdlet.ShouldProcess($BatchId, "Event $Step")) { return }
    $null = Invoke-Pra2Sql $Connection 'INSERT INTO batch_event (batch_id, object_guid, time_utc, step, outcome, detail) VALUES (@b, @o, @t, @s, @r, @d)' @{
        b = $BatchId; o = $ObjectGuid; t = [datetime]::UtcNow; s = $Step; r = $Outcome; d = $Detail }
}

function Set-Pra2TenantChange {
    <# Tenant-level change and restoration. UpdateOriginal explicitly advances a journaled room-list creation intent; legacy originals stay immutable. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$TargetId, [string]$OriginalJson, [switch]$Restored, [switch]$UpdateOriginal)
    if (-not $PSCmdlet.ShouldProcess("$Kind $TargetId", 'Journal')) { return }
    $exists = [int](Invoke-Pra2Sql $Connection 'SELECT count(*) FROM tenant_change WHERE batch_id = @b AND kind = @k AND target_id = @t' @{ b = $BatchId; k = $Kind; t = $TargetId } -As Scalar)
    if ($exists) {
        $setOriginal = if ($UpdateOriginal) { ', original_json = @o' } else { '' }
        $null = Invoke-Pra2Sql $Connection ("UPDATE tenant_change SET restored = @r, updated_utc = @u{0} WHERE batch_id = @b AND kind = @k AND target_id = @t" -f $setOriginal) @{
            r = [bool]$Restored; u = [datetime]::UtcNow; b = $BatchId; k = $Kind; t = $TargetId; o = $OriginalJson }
    }
    else { $null = Invoke-Pra2Sql $Connection 'INSERT INTO tenant_change (batch_id, kind, target_id, original_json, restored, updated_utc) VALUES (@b, @k, @t, @o, @r, @u)' @{ b = $BatchId; k = $Kind; t = $TargetId; o = $OriginalJson; r = [bool]$Restored; u = [datetime]::UtcNow } }
}

function Get-Pra2TenantChange {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId)
    return Invoke-Pra2Sql $Connection 'SELECT * FROM tenant_change WHERE batch_id = @b' @{ b = $BatchId } -As Rows
}
#endregion

Export-ModuleMember -Function Import-Pra2Sqlite, Open-Pra2Store, Close-Pra2Store, Invoke-Pra2Sql, New-Pra2Snapshot, Set-Pra2SnapshotStatus,
    Add-Pra2Row, Get-Pra2Snapshot, Get-Pra2SnapshotList, Read-Pra2Table, Remove-Pra2OldSnapshot, Backup-Pra2Store,
    New-Pra2Batch, Set-Pra2Batch, Get-Pra2Batch, Set-Pra2BatchItem, Get-Pra2BatchItem, Add-Pra2BatchEvent, Set-Pra2TenantChange, Get-Pra2TenantChange

# SIG # Begin signature block
# MIIghQYJKoZIhvcNAQcCoIIgdjCCIHICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCChfyB1yBzb/5kG
# VVkDm853VOAEsNlusVk62ktl1TRmHaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYISbDCCEmgCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBctBy3f1QA
# 36xONSM/RACSZupNDNjjGNVWOZnjJgN5KDANBgkqhkiG9w0BAQEFAASCAgCr1kP8
# M5iNLtIFqCkf4lyIy6Ey2uXl42a4Y/MYFqTnj6/H07yRW2klXupconVLW8d9IPKe
# tVIwgDWs0ompV2veRJGwBOpp6dIkJh8bRA17XV6b6qPw+6FDU/m6CPrpjnnCEWMs
# SKw7cgu4+4nu0hUUYpFitWwqHrxpOodg/1YmI4rqsNiFeuyAuVL2MmbXGZvNNFDO
# euZAjNH8UbvoGxvE7R6M/Tt33rRy/uWdCn+aiwa1XTFJ3avL29fVWGC0yPbk2mJ3
# 2TN333JLLBLlmESq1kdhBoMrgWDLX6rYFj8tE+A7bMguBFmter/uxltV7AAIylFU
# E82z/mmRQ6pLw8rh5UjIJo3jtMSVQd1yOvJXiZ6dPWRhM3lEMJo2pPhFLLYmrAEf
# WpepBsWgOW6EO79ugRCM8JS9jtaeix50X551IFkmNagnGKsyxYz1jD4qh386ivTy
# 8DcTzrzKuXzz0gjTrBL0iXQRszwJ3TUuWZybIbfFX4beXHoXm/xVdm/69LwiFLj3
# 8X5nWpZKsdvh//i1eCwWSb4naQhkMJoQIls+qvpU1/D5yKCeqU8uCiQ4wIrPjD/s
# WBX43NNQ6YDLD+vJsg71C9eQc5w3RPA1/YO65/xM+NuYaolFYREIzRo6XYO4DKh9
# qWt1R6hknjlduNTn8680JemC8tTVKAxDZy3Z+aGCDykwgg8lBgorBgEEAYI3AwMB
# MYIPFTCCDxEGCSqGSIb3DQEHAqCCDwIwgg7+AgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIPXZ0Ii5IjaOLvvmPHcaUd/nyDFLIwU7Gkqilg3V3WKZAghX2/bqXOy9
# nRgPMjAyNjEwMTQxNTIzNDJaMAMCAQECEEeSWHAVzL7xCZ9w0g1s//agggwAMIIE
# /DCCAuSgAwIBAgIQH2sWYtIuG2xd8cDBoGAOODANBgkqhkiG9w0BAQsFADBzMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS8wLQYDVQQDDCZTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSBSMTAeFw0yNTAyMTgxNjMyMDJaFw0zNDExMTIxODUwMDVa
# MG4xCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3Rv
# bjERMA8GA1UECgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1w
# aW5nIFVuaXQgMjAyNSBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABBv7UVHH
# b+ZVluxPXlfE3M6tg0Xnq8dic+O3vPOCRpalUM1vO9A+GRzSVVjyygHhYrBw62XL
# Fh1kv7e+yRd/aTajggFaMIIBVjAfBgNVHSMEGDAWgBQMnRAljpqnG5mHQ88IfuG9
# gZD0zzBRBggrBgEFBQcBAQRFMEMwQQYIKwYBBQUHMAKGNWh0dHA6Ly9jZXJ0LnNz
# bC5jb20vU1NMLmNvbS10aW1lU3RhbXBpbmctSS1SU0EtUjEuY2VyMFEGA1UdIARK
# MEgwPAYMKwYBBAGCqTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNz
# bC5jb20vcmVwb3NpdG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUH
# AwgwRgYDVR0fBD8wPTA7oDmgN4Y1aHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29t
# LXRpbWVTdGFtcGluZy1JLVJTQS1SMS5jcmwwHQYDVR0OBBYEFM582cAEgMUkEGoJ
# 6hyrJT0R/ajSMA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAgHN1
# JIW9ZlNip07FfY9r1LJwLKTl1h/y9r1BjDgHMEz7U7J7mAy32JxA1KLyjCCJqNku
# RfhDHGFEfZtjTOru7AQMDSDiUDvCCc3jaz5py9jGvwIe5npkLbXqXhA+SqCzG719
# USigwhHkYr7v8NiXoaY0Xu03p0BiZZreZlpngmGB9+86N94FqIFx2I5Z5iuu7d4j
# 7aAOZBNxfU9jFDIIGDtWfWeCpF0Q8rLE/3DHDlzJKs5Il3i/VoOl7rTU938oSrxM
# www6GDBlKX8rGHxvkjgIb9MXkVglJSDHsNAZDeCS/inaxKHA+ChuB8K4OxIxa6k1
# e+eHBypakAY3wuN7w7PlPmhN6k5IdqZn3HZHs2VjoR841Z9wmVPFdGDdkBZ7XOmb
# 5OZUaKufsvtER0VbY2DkzaqAaItt1kc+I+FUz7PiVz2PUzpegkftRRxvrPyIL08b
# lJTnbMQ0XrqDN0rEfuTLv4XhhnKyJaXhxCUdjO/cRumMzZ35QTUMItBWy1xGp0iy
# oFVnAya2pPsVMuAI+sy1zxlOS9l5iSxvKJ8gpnpOgjqYTa7u1eyZo+4JlcGoiiR1
# 7LthMTrF1q62tO5xOpI8txUZ6gtKGtkJV6wdj+vZLcwKZ2M93GqVpOo6UBlGjx1s
# KsVZ7AIUzJZli4NJ0M2KYcrfNnlebkYKpDnSqWwwggb8MIIE5KADAgECAhBtUhhw
# h+gjTYVgANCAj5NWMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTE5MTExMzE4NTAwNVoXDTM0MTExMjE4NTAwNVowczELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCu
# URAT0vk8IKAghd7JUBxkyeH9xek0/wp/MUjoclrFXqhh/fGH91Fc+7fm0MHCE7A+
# wmOiqBj9ODrJAYGq3rm33jCnHSsCBNWAQYyoauLq8IjqsS1JlXL29qDNMMdwZ8UN
# zQS7vWZMDJ40JSGNphMGTIA2qn2bohGtgRc4p1395ESypUOaGvJ3t0FNL3BuKmb6
# YctMcQUF2sqooMzd89h0E6ujdvBDo6ZwNnWoxj7YmfWjSXg33A5GuY9ym4QZM5OE
# Vgo8ebz/B+gyhyCLNNhh4Mb/4xvCTCMVmNYrBviGgdPZYrym8Zb84TQCmSuX0JlL
# La6WK1aO6qlwISbb9bVGh866ekKblC/XRP20gAu1CjvcYciUgNTrGFg8f8AJgQPO
# Cc1/CCdaJSYwhJpSdheKOnQgESgNmYZPhFOC6IKaMAUXk5U1tjTcFCgFvvArXtK4
# azAWUOO1Y3fdldIBL6LjkzLUCYJNkFXqhsBVcPMuB0nUDWvLJfPimstjJ8lF4S6E
# CxWnlWi7OElVwTnt1GtRqeY9ydvvGLntU+FecK7DbqHDUd366UreMkSBtzevAc9a
# qoZPnjVMjvFqV1pYOjzmTiVHZtAc80bAfFe5LLfJzPI6DntNyqobpwTevQpHqPDN
# 9qqNO83r3kaw8A9j+HZiSw2AX5cGdQP0kG0vhzfgBwIDAQABo4IBgTCCAX0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTCBgwYIKwYBBQUHAQEEdzB1MFEGCCsGAQUFBzAChkVodHRwOi8vd3d3LnNzbC5j
# b20vcmVwb3NpdG9yeS9TU0xjb21Sb290Q2VydGlmaWNhdGlvbkF1dGhvcml0eVJT
# QS5jcnQwIAYIKwYBBQUHMAGGFGh0dHA6Ly9vY3Nwcy5zc2wuY29tMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwOwYDVR0fBDQwMjAwoC6gLIYq
# aHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0GA1Ud
# DgQWBBQMnRAljpqnG5mHQ88IfuG9gZD0zzAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAJIZdQ2mWkLPGQfZ8vyU+sCb8BXpRJZaL3Ez3VDlE3uZk3cP
# xPtybVfLuqaci0W6SB22JTMttCiQMnIVOsXWnIuAbD/aFTcUkTLBI3xys+wEajzX
# aXJYWACDS47BRjDtYlDW14gLJxf8W6DQoH3jHDGGy8kGJFOlDKG7/YrK7UGfHtBA
# EDVe6lyZ+FtCsrk7dD/IiL/+Q3Q6SFASJLQ2XI89ihFugdYL77CiDNXrI2MFspQG
# swXEAGpHuaQDTHUp/LdR3TyrIsLlnzoLskUGswF/KF8+kpWUiKJNC4rPWtNrxlbX
# YRGgdEdx8SMjUTDClldcrknlFxbqHsVmr9xkT2QtFmG+dEq1v5fsIK0vHaHrWjMM
# maJ9i+4qGJSD0stYfQ6v0PddT7EpGxGd867Ada6FZyHwbuQSadMb0K0P0OC2r7rw
# qBUe0BaMqTa6LWzWItgBjGcObXeMxmbQqlEz2YtAcErkZvh0WABDDE4U8GyV/32F
# daAvJgTfe9MiL2nSBioYe/g5mHUSWAay/Ip1RQmQCvmF9sNfqlhJwkjy/1U1ibUk
# TIUBX3HgymyQvqQTZLLys6pL2tCdWcjI9YuLw30rgZm8+K387L7ycUvqrmQ3ZJlu
# jHl3r1hgV76s3WwMPgKk1bAEFMj+rRXimSC+Ev30hXZdqyMdl/il5Ksd0vhGMYIC
# WDCCAlQCAQEwgYcwczELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEvMC0GA1UEAwwmU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgUjECEB9rFmLSLhtsXfHAwaBg
# DjgwCwYJYIZIAWUDBAIBoIIBYTAaBgkqhkiG9w0BCQMxDQYLKoZIhvcNAQkQAQQw
# HAYJKoZIhvcNAQkFMQ8XDTI2MTAxNDE1MjM0MlowKAYJKoZIhvcNAQk0MRswGTAL
# BglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIOfjx+zYL56G
# +go2wqZULjUElxQcr1Ho57oFCcNRfLauMIHJBgsqhkiG9w0BCRACLzGBuTCBtjCB
# szCBsAQgVCr5oWqNci5mEUl4iumUwYqaruWmXLNEolSa+Wx5x4swgYswd6R1MHMx
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjER
# MA8GA1UECgwIU1NMIENvcnAxLzAtBgNVBAMMJlNTTC5jb20gVGltZXN0YW1waW5n
# IElzc3VpbmcgUlNBIENBIFIxAhAfaxZi0i4bbF3xwMGgYA44MAoGCCqGSM49BAMC
# BEcwRQIhAIq73p1UQi5A47n/BzUi9csSGw4TmGMeiRi2HMtaY9ELAiA8V5j7DeTg
# nb84bu0F0euaeSnEVvx7Zqc0EWWR2eLudw==
# SIG # End signature block
