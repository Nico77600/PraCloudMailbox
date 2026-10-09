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
    Version : 1.2.0
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:SchemaVersion = 1
$script:SqliteFlavor = ''

# Tables of schema version 1. Column names are the keys of the rows written by Add-Pra2Row.
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
    item_count INTEGER, total_item_size_bytes INTEGER,
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
    <# Tenant-level change of a batch (e.g. the source of authority of the licence group) and whether it was restored. #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([Parameter(Mandatory)][object]$Connection, [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$Kind,
        [Parameter(Mandatory)][string]$TargetId, [string]$OriginalJson, [switch]$Restored)
    if (-not $PSCmdlet.ShouldProcess("$Kind $TargetId", 'Journal')) { return }
    $exists = [int](Invoke-Pra2Sql $Connection 'SELECT count(*) FROM tenant_change WHERE batch_id = @b AND kind = @k AND target_id = @t' @{ b = $BatchId; k = $Kind; t = $TargetId } -As Scalar)
    if ($exists) { $null = Invoke-Pra2Sql $Connection 'UPDATE tenant_change SET restored = @r, updated_utc = @u WHERE batch_id = @b AND kind = @k AND target_id = @t' @{ r = [bool]$Restored; u = [datetime]::UtcNow; b = $BatchId; k = $Kind; t = $TargetId } }
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
