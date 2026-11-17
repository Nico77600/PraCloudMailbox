#Requires -Version 7.5
<#
.SYNOPSIS
    Writes the legacy stores of tests\fixtures\legacy with the store module of the releases that created them (repository tool, not in the package).

.DESCRIPTION
    The stores that earlier releases wrote must stay readable. They are made here, once, by the very store module of those releases (read from the git history with
    git show), with synthetic data only (contoso.com), and committed byte-exact:
        snapshot-schema1.db   snapshot schema 1 and journal schema 1: release 1.0.0 (no resource settings, no target groups)
        snapshot-schema2.db   snapshot schema 2 (resource settings of rooms) written by release 1.2.0
        journal-schema1.db    journal schema 1 (batches, objects, events, tenant changes) written by release 1.2.0
    fixtures.sha256 lists the SHA-256 of each file: the tests check the files against it before they use them. The SQLite library is the one of package\lib\sqlite.

.PARAMETER OutputFolder
    Default: tests\fixtures\legacy.

.PARAMETER ComputerName
    The computer name the releases record in the stores. Default: PRA-HOST, so that the files never carry the name of the computer that makes them.

.EXAMPLE
    .\tools\New-LegacyFixtures.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[CmdletBinding()]
param([string]$OutputFolder, [ValidatePattern('^[A-Za-z0-9-]{1,15}$')][string]$ComputerName = 'PRA-HOST')

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'tests\fixtures\legacy' }
$null = New-Item -ItemType Directory -Path $OutputFolder -Force
$work = Join-Path ([IO.Path]::GetTempPath()) ('legacy-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = New-Item -ItemType Directory -Path $work -Force
$library = Join-Path $root 'package'
$pwsh = Join-Path $PSHOME 'pwsh.exe'

function Get-StoreModule([string]$Reference, [string]$Name) {
    $folder = Join-Path $work $Name
    $null = New-Item -ItemType Directory -Path $folder -Force
    $path = Join-Path $folder 'PRA2.Store.psm1'
    $previous = [Console]::OutputEncoding
    [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
    try { $text = & git -C $root show "${Reference}:module/PRA2.Store.psm1" }
    finally { [Console]::OutputEncoding = $previous }
    if ($LASTEXITCODE -ne 0) { throw "Cannot read the store module of $Reference from git." }
    [IO.File]::WriteAllText($path, ($text -join "`n"), [Text.UTF8Encoding]::new($false))
    return $path
}

$mailboxes = @(
    @{ guid = 'a0000000-0000-0000-0000-000000000001'; kind = 'User'; type = 'UserMailbox'; sam = 'adele.vance'; upn = 'adele.vance@contoso.com'; name = 'Adele Vance' }
    @{ guid = 'a0000000-0000-0000-0000-000000000002'; kind = 'User'; type = 'UserMailbox'; sam = 'alex.wilber'; upn = 'alex.wilber@contoso.com'; name = 'Alex Wilber' }
    @{ guid = 'a0000000-0000-0000-0000-000000000003'; kind = 'Shared'; type = 'SharedMailbox'; sam = 'accounting'; upn = 'accounting@contoso.com'; name = 'Accounting' }
)

$snapshotScript = @'
param([string]$Module, [string]$Root, [string]$Output, [string]$Release, [switch]$WithResource)
$ErrorActionPreference = 'Stop'
Import-Module $Module -Force
$cn = Open-Pra2Store -Path $Output -Root $Root
try {
    $id = New-Pra2Snapshot $cn @{ RunId = 'run-legacy'; Environment = 'PROD'; Account = 'CONTOSO\pra.operator'; ToolVersion = $Release; ExchangeServer = 'SRV-EXCH01'; ScopeJson = '{"Mode":"Auto"}' }
    $mailboxes = $env:PCM_LEGACY_MAILBOXES | ConvertFrom-Json
    $rows = @(foreach ($m in $mailboxes) {
            $row = [ordered]@{ object_guid = $m.guid; kind = $m.kind; recipient_type_details = $m.type; sam_account_name = $m.sam; user_principal_name = $m.upn; name = $m.name; display_name = $m.name
                alias = $m.sam; distinguished_name = ('CN={0},OU=Paris,DC=contoso,DC=com' -f $m.name); organizational_unit = 'contoso.com/Paris'; primary_smtp_address = $m.upn
                legacy_exchange_dn = ('/o=CONTOSO/ou=Exchange Administrative Group/cn=Recipients/cn={0}' -f $m.sam); exchange_guid = [guid]::NewGuid().ToString(); hidden_from_address_lists = 0
                immutable_id = [Convert]::ToBase64String(([guid]$m.guid).ToByteArray()); email_addresses_json = ('["SMTP:{0}","smtp:{1}@contoso.local"]' -f $m.upn, $m.sam)
                custom_attributes_json = '{}'; grant_send_on_behalf_json = '[]' }
            if ($WithResource) { $row['resource_settings_json'] = $null }
            $row
        })
    if ($WithResource) {
        $rows += [ordered]@{ object_guid = 'a0000000-0000-0000-0000-000000000004'; kind = 'Room'; recipient_type_details = 'RoomMailbox'; sam_account_name = 'room-paris'; user_principal_name = 'room.paris@contoso.com'
            name = 'Room Paris'; display_name = 'Room Paris'; alias = 'room-paris'; distinguished_name = 'CN=Room Paris,OU=Paris,DC=contoso,DC=com'; organizational_unit = 'contoso.com/Paris'
            primary_smtp_address = 'room.paris@contoso.com'; legacy_exchange_dn = '/o=CONTOSO/cn=Recipients/cn=room-paris'; exchange_guid = [guid]::NewGuid().ToString(); hidden_from_address_lists = 0
            immutable_id = [Convert]::ToBase64String(([guid]'a0000000-0000-0000-0000-000000000004').ToByteArray()); email_addresses_json = '["SMTP:room.paris@contoso.com"]'
            custom_attributes_json = '{}'; grant_send_on_behalf_json = '[]'; resource_settings_json = '{"SchemaVersion":2,"Capacity":8,"ResourceCustom":[],"Calendar":{"AutomateProcessing":"AutoAccept"}}' }
    }
    $null = Add-Pra2Row $cn mailbox $id $rows
    $null = Add-Pra2Row $cn permission $id @([ordered]@{ mailbox_guid = 'a0000000-0000-0000-0000-000000000003'; access_right = 'FullAccess'; trustee = 'CONTOSO\adele.vance'; trustee_guid = 'a0000000-0000-0000-0000-000000000001'
            trustee_kind = 'User'; trustee_upn = 'adele.vance@contoso.com'; trustee_smtp = 'adele.vance@contoso.com'; trustee_name = 'Adele Vance'; trustee_dn = 'CN=Adele Vance,OU=Paris,DC=contoso,DC=com'; auto_mapping = 1; resolved = 1; note = '' })
    Set-Pra2SnapshotStatus $cn $id Complete @{ mailbox = $rows.Count; permission = 1 }
} finally { Close-Pra2Store $cn }
'@

$journalScript = @'
param([string]$Module, [string]$Root, [string]$Output, [string]$Release)
$ErrorActionPreference = 'Stop'
Import-Module $Module -Force
$cn = Open-Pra2Store -Path $Output -Root $Root -Kind Journal
try {
    $batch = New-Pra2Batch $cn Convert @{ SnapshotId = 1; Environment = 'PROD'; Account = 'CONTOSO\pra.operator'; ToolVersion = $Release; ScopeJson = '{"Identity":"","Scope":"All"}' }
    Set-Pra2BatchItem $cn $batch 'a0000000-0000-0000-0000-000000000001' @{ kind = 'Mailbox'; identity = 'adele.vance@contoso.com'; step = 'Done'; status = 'Done'; message = '' }
    Set-Pra2BatchItem $cn $batch 'a0000000-0000-0000-0000-000000000003' @{ kind = 'Shared'; identity = 'accounting@contoso.com'; step = 'Permissions'; status = 'Done'; message = '' }
    Add-Pra2BatchEvent $cn $batch 'a0000000-0000-0000-0000-000000000001' 'Licence' 'Ok' 'Exchange plan assigned'
    Add-Pra2BatchEvent $cn $batch 'a0000000-0000-0000-0000-000000000003' 'Permissions' 'Ok' 'FullAccess granted'
    Set-Pra2TenantChange $cn $batch 'LicenceUnit' 'EXCHANGESTANDARD' '{"consumed":10}'
    Set-Pra2Batch $cn $batch Complete @{ done = 2; error = 0 } 'legacy fixture'
    $recover = New-Pra2Batch $cn Recover @{ ConvertBatch = $batch; SnapshotId = 1; Environment = 'PROD'; Account = 'CONTOSO\pra.operator'; ToolVersion = $Release; ScopeJson = '{}' }
    Set-Pra2BatchItem $cn $recover 'a0000000-0000-0000-0000-000000000001' @{ kind = 'RecoverUser'; identity = 'adele.vance@contoso.com'; step = 'Done'; status = 'Done'; message = '' }
    Set-Pra2Batch $cn $recover Complete @{ done = 1 } ''
    $batch
} finally { Close-Pra2Store $cn }
'@
$snapshotFile = Join-Path $work 'snapshot.ps1'; $journalFile = Join-Path $work 'journal.ps1'
[IO.File]::WriteAllText($snapshotFile, $snapshotScript, [Text.UTF8Encoding]::new($true))
[IO.File]::WriteAllText($journalFile, $journalScript, [Text.UTF8Encoding]::new($true))
$previousComputerName = $env:COMPUTERNAME
$env:COMPUTERNAME = $ComputerName
$env:PCM_LEGACY_MAILBOXES = ($mailboxes | ForEach-Object { [pscustomobject]$_ } | ConvertTo-Json -Compress)

function Invoke-Child([string]$File, [string[]]$Arguments) {
    $output = & $pwsh -NoProfile -NonInteractive -File $File @Arguments 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Legacy fixture generation failed ($File $($Arguments -join " ")):`n$output" }
}

$plan = @(
    @{ Name = 'snapshot-schema1.db'; Reference = '27f659f'; Release = '1.0.0'; Kind = 'Snapshot'; Resource = $false }
    @{ Name = 'snapshot-schema2.db'; Reference = 'eecaf6f'; Release = '1.2.0'; Kind = 'Snapshot'; Resource = $true }
    @{ Name = 'journal-schema1.db'; Reference = 'eecaf6f'; Release = '1.2.0'; Kind = 'Journal'; Resource = $false }
)
foreach ($item in $plan) {
    $module = Get-StoreModule $item.Reference ($item.Name -replace '\.db$', '')
    $target = Join-Path $OutputFolder $item.Name
    if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
    if ($item.Kind -eq 'Snapshot') { Invoke-Child $snapshotFile (@('-Module', $module, '-Root', $library, '-Output', $target, '-Release', $item.Release) + $(if ($item.Resource) { @('-WithResource') } else { @() })) }
    else { Invoke-Child $journalFile @('-Module', $module, '-Root', $library, '-Output', $target, '-Release', $item.Release) }
}
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
$lines = foreach ($item in $plan) { '{0}  {1}' -f (Get-FileHash -LiteralPath (Join-Path $OutputFolder $item.Name) -Algorithm SHA256).Hash.ToLowerInvariant(), $item.Name }
[IO.File]::WriteAllText((Join-Path $OutputFolder 'fixtures.sha256'), (($lines -join "`n") + "`n"), [Text.UTF8Encoding]::new($false))
Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item Env:\PCM_LEGACY_MAILBOXES -ErrorAction SilentlyContinue
$env:COMPUTERNAME = $previousComputerName
Write-Host ('{0} legacy stores written to {1}' -f $plan.Count, $OutputFolder)
