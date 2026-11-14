<#
.SYNOPSIS
    The synthetic world of the console pictures and tests: a configuration, a snapshot of 20 users and 4 shared mailboxes, a group captured by Collect, a journal with
    batches in every state, and the scenario of the synthetic engine (repository tool, never shipped). Synthetic names only (contoso.com, Adele Vance...).
.NOTES
    Needs PRA.Common.psm1 and PRA.Store.psm1 imported, and the lib\sqlite folder of the tool under ToolRoot.
#>
$script:FixtureUsers = @('Adele Vance', 'Alex Wilber', 'Megan Bowen', 'Diego Siciliani', 'Grady Archie', 'Isaiah Langford', 'Lee Gu', 'Lynne Robbins', 'Miriam Graham', 'Nestor Wilke',
    'Patti Fernandez', 'Johanna Lorenz', 'Joni Sherman', 'Henrietta Mueller', 'Pradeep Gupta', 'Christie Cline', 'Debra Berger', 'Irvin Sayers', 'Allan Deyoung', 'Emily Braun')
$script:FixtureShared = @('Accounting', 'Helpdesk', 'Reception', 'Sales')

function Get-FixtureGuid([int]$Number) { return ('a0000000-0000-0000-0000-{0:D12}' -f $Number) }

function New-ConsoleFixture {
    <#
        Writes the configuration, the snapshot database and the journal under ToolRoot, and the scenario of the synthetic engine next to the synthetic script (not with -NoScenario: the real script verifies its folder, where no foreign file may be).
        Returns Config (path), Snapshot (id), Users, Shared (rows), Batches (ids).
    #>
    param([Parameter(Mandatory)][string]$ToolRoot, [switch]$NoSnapshot, [switch]$NoScenario, [int]$SnapshotAgeHours = 14)
    $data = Join-Path $ToolRoot 'data'
    $null = New-Item -ItemType Directory -Path $data, (Join-Path $ToolRoot 'reports'), (Join-Path $ToolRoot 'logs') -Force
    $config = Join-Path $ToolRoot 'config\PraCloudMailbox.config.psd1'
    $null = New-Item -ItemType Directory -Path (Split-Path $config -Parent) -Force
    $text = @"
@{
    Environment = 'PROD'
    Exchange = @{ Server = 'SRV-EXCH01.contoso.com' }
    Scope = @{ Mode = 'Auto' }
    EntraConnect = @{ Mode = 'Manual' }
    Licensing = @{ Shared = @{ SkuPartNumber = 'EXCHANGESTANDARD' } }
    Store = @{ Path = '$data\PraCloudMailbox.db'; JournalPath = '$data\PraCloudMailbox-journal.db'; BackupFolder = '$data\backup'; MaxSnapshotAgeDays = 2 }
    Cloud = @{ TenantId = '00000000-0000-0000-0000-000000000000'; Organization = 'contoso.onmicrosoft.com'; AppId = '11111111-2222-3333-4444-555555555555'; CertificateThumbprint = '0A1B2C3D4E5F60718293A4B5C6D7E8F901234567' }
    Logging = @{ Folder = '$ToolRoot\logs' }
    Report = @{ Enabled = `$true; Folder = '$ToolRoot\reports' }
}
"@
    [IO.File]::WriteAllText($config, $text, (New-Object Text.UTF8Encoding($true)))
    $objects = [Collections.Generic.List[object]]::new()
    $number = 0
    foreach ($name in $script:FixtureUsers) {
        $number++
        $local = ($name.ToLowerInvariant() -replace ' ', '.')
        $unit = if ($number -le 12) { @('contoso.com/Paris/Users', 'OU=Users,OU=Paris') } else { @('contoso.com/Lyon', 'OU=Lyon') }
        $objects.Add([pscustomobject]@{ guid = (Get-FixtureGuid $number); upn = "$local@contoso.com"; name = $name; kind = 'User'; shared = $false; ou = $unit[0]; dn = ('CN={0},{1},DC=contoso,DC=com' -f $name, $unit[1]); sam = ($local -replace '\.', '') ; warning = ''; blocked = '' })
    }
    foreach ($name in $script:FixtureShared) {
        $number++
        $local = $name.ToLowerInvariant()
        $objects.Add([pscustomobject]@{ guid = (Get-FixtureGuid $number); upn = "$local@contoso.com"; name = $name; kind = 'Shared'; shared = $true; ou = 'contoso.com/Shared'; dn = ('CN={0},OU=Shared,DC=contoso,DC=com' -f $name); sam = $local; warning = ''; blocked = '' })
    }
    $result = [ordered]@{ Config = $config; Snapshot = 0; Users = $script:FixtureUsers; Shared = $script:FixtureShared; Objects = $objects.ToArray(); Batches = @() }
    if (-not $NoSnapshot) {
        $connection = Open-PcmStore -Path (Join-Path $data 'PraCloudMailbox.db') -Root $ToolRoot
        try {
            # The pictures show snapshot 42: the counter of the database starts there.
            $null = Invoke-PcmSql $connection "INSERT INTO sqlite_sequence (name, seq) VALUES ('snapshot', 41)"
            $sid = New-PcmSnapshot $connection @{ RunId = 'run-0042'; Environment = 'PROD'; ToolVersion = '2.0.0'; ExchangeServer = 'SRV-EXCH01'; ScopeJson = '{}' }
            $rows = @($objects | ForEach-Object {
                    [ordered]@{ object_guid = $_.guid; kind = $_.kind; recipient_type_details = $(if ($_.shared) { 'SharedMailbox' } else { 'UserMailbox' }); sam_account_name = $_.sam; user_principal_name = $_.upn; name = $_.name; display_name = $_.name
                        alias = $_.sam; distinguished_name = $_.dn; organizational_unit = $_.ou; primary_smtp_address = $_.upn
                        email_addresses_json = (@(('SMTP:' + $_.upn), ('smtp:{0}@contoso.com' -f $_.sam)) | ConvertTo-Json -Compress) }
                })
            $null = Add-PcmRow $connection mailbox $sid $rows
            $groupGuid = 'b0000000-0000-0000-0000-000000000001'
            $null = Add-PcmRow $connection target_group $sid @([ordered]@{ group_guid = $groupGuid; name = 'PRA-Wave1'; sam_account_name = 'PRA-Wave1'; distinguished_name = 'CN=PRA-Wave1,OU=Groups,DC=contoso,DC=com'; primary_smtp_address = ''; display_name = 'PRA-Wave1'; member_count = 26 })
            $members = @($objects | ForEach-Object { [ordered]@{ group_guid = $groupGuid; member_guid = $_.guid; member_class = 'User'; member_name = $_.name; member_upn = $_.upn; member_smtp = $_.upn; member_dn = $_.dn; via_group_guid = ''; depth = 1 } })
            $members += [ordered]@{ group_guid = $groupGuid; member_guid = 'c0000000-0000-0000-0000-000000000001'; member_class = 'Contact'; member_name = 'Print Room'; member_upn = ''; member_smtp = 'print.room@contoso.com'; member_dn = ''; via_group_guid = ''; depth = 1 }
            $members += [ordered]@{ group_guid = $groupGuid; member_guid = 'c0000000-0000-0000-0000-000000000002'; member_class = 'Computer'; member_name = 'KIOSK-01'; member_upn = ''; member_smtp = ''; member_dn = ''; via_group_guid = ''; depth = 2 }
            $null = Add-PcmRow $connection target_member $sid $members
            Set-PcmSnapshotStatus $connection $sid Complete @{ mailbox = $objects.Count; permission = 0 }
            $finished = [datetime]::UtcNow.AddHours(-$SnapshotAgeHours)
            $null = Invoke-PcmSql $connection 'UPDATE snapshot SET finished_utc = @f, started_utc = @s, computer = @c WHERE id = @id' @{ f = $finished; s = $finished.AddMinutes(-9); c = 'SRV-EXCH01'; id = $sid }
            $result.Snapshot = $sid
        } finally { Close-PcmStore $connection }
        $null = Write-PcmHashFile -Path (Join-Path $data 'PraCloudMailbox.db')
    }
    $journal = Open-PcmStore -Path (Join-Path $data 'PraCloudMailbox-journal.db') -Root $ToolRoot -Kind Journal
    try {
        $batch = {
            param($Id, $Action, $Convert, $Status, $Hours, $Snapshot)
            $null = Invoke-PcmSql $journal 'INSERT INTO batch (id, action, convert_batch, status, created_utc, updated_utc, snapshot_id, environment, computer, account, tool_version) VALUES (@id, @a, @c, @s, @t, @t, @n, @e, @m, @u, @v)' @{ id = $Id; a = $Action; c = $Convert; s = $Status; t = [datetime]::UtcNow.AddHours(-$Hours); n = $Snapshot; e = 'PROD'; m = 'SRV-PRA01'; u = 'CONTOSO\pra.operator'; v = '2.0.0' }
        }
        $item = {
            param($Batch, $Object, $Step, $Status)
            $null = Invoke-PcmSql $journal 'INSERT INTO batch_item (batch_id, object_guid, kind, identity, step, status, updated_utc) VALUES (@b, @o, @k, @i, @s, @t, @u)' @{ b = $Batch; o = $Object.guid; k = $Object.kind; i = $Object.upn; s = $Step; t = $Status; u = [datetime]::UtcNow }
        }
        & $batch '1a2b3c4d' 'Convert' '' 'Complete' 2 $result.Snapshot
        foreach ($o in $objects) { & $item '1a2b3c4d' $o 'Mailbox' 'Done' }
        & $batch '9c8d7e6f' 'Convert' '' 'Complete' 30 41
        foreach ($o in @($objects | Select-Object -First 4)) { & $item '9c8d7e6f' $o 'Mailbox' 'Done' }
        & $batch '5e6f7a8b' 'Convert' '' 'Partial' 5 $result.Snapshot
        foreach ($o in @($objects | Select-Object -First 20)) { & $item '5e6f7a8b' $o $(if ($o -in @($objects | Select-Object -First 12)) { 'Mailbox' } else { '' }) $(if ($o -in @($objects | Select-Object -First 12)) { 'Done' } else { 'Planned' }) }
        & $batch 'ab12cd34' 'Recover' '9c8d7e6f' 'Partial' 28 41
        foreach ($o in @($objects | Select-Object -First 2)) { & $item 'ab12cd34' $o 'RecoverUser' 'Done' }
        $result.Batches = @('1a2b3c4d', '9c8d7e6f', '5e6f7a8b', 'ab12cd34')
    } finally { Close-PcmStore $journal }
    $scenario = [ordered]@{ batchId = '1a2b3c4d'; holdAfter = 0
        objects = @($objects | ForEach-Object { [ordered]@{ guid = $_.guid; upn = $_.upn; shared = $_.shared; warning = $(if ($_.name -eq 'Alex Wilber') { 'Found by UPN only: onPremisesImmutableId differs from the objectGUID (7Gk2x0Q9yEOnZ4mF1aB3cA==); check the source anchor.' } else { '' }); blocked = '' } })
        skipped = @([ordered]@{ upn = 'print.room@contoso.com'; kind = 'Contact' }, [ordered]@{ upn = 'kiosk-01'; kind = 'Computer' }) }
    if (-not $NoScenario) { [IO.File]::WriteAllText((Join-Path $ToolRoot 'scenario.json'), ($scenario | ConvertTo-Json -Depth 6), (New-Object Text.UTF8Encoding($false))) }
    return [pscustomobject]$result
}
