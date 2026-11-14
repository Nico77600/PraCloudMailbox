<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the snapshot check (Pester 5+): the .sha256 that Collect writes beside the snapshot file and its backup copy, and the check that
    Check, Convert and Recover make before they read the snapshot. No Exchange and no cloud: Collect runs in a child process against synthetic Exchange commands.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
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
    $script:Entry = Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    $script:Work = Join-Path ([IO.Path]::GetTempPath()) ('PCM-snap-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $null = New-Item -ItemType Directory -Path $script:Work -Force

    function script:New-SnapshotConfig {
        param([Parameter(Mandatory)][string]$Folder)
        $null = New-Item -ItemType Directory -Path $Folder -Force
        $text = @"
@{
    Environment = 'TEST'
    Exchange = @{ Server = 'ex1.contoso.com' }
    Scope = @{ Mode = 'Auto'; IncludeUsers = `$true; IncludeShared = `$true }
    Collect = @{ SharedPermissions = `$true }
    Store = @{ Path = '$Folder\db\pra.db'; BackupFolder = '$Folder\backup'; JournalPath = '$Folder\db\journal.db' }
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
        <# Runs the entry script in a child PowerShell with synthetic Exchange commands (Collect) or a synthetic module list (the cloud actions). #>
        param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string[]]$Arguments)
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
    $make = { param($sam, $type) [pscustomobject]@{ Guid = [guid]::NewGuid(); Identity = $sam; Name = $sam; DisplayName = $sam; Alias = $sam; SamAccountName = $sam
        UserPrincipalName = "$sam@contoso.com"; DistinguishedName = "CN=$sam,OU=PRA,DC=contoso,DC=com"; OrganizationalUnit = 'contoso.com/PRA'; PrimarySmtpAddress = "$sam@contoso.com"
        WindowsEmailAddress = "$sam@contoso.com"; LegacyExchangeDN = "/o=C/cn=$sam"; ExchangeGuid = [guid]::NewGuid(); ArchiveGuid = [guid]::Empty; ArchiveStatus = 'None'
        RecipientTypeDetails = $type; HiddenFromAddressListsEnabled = $false; Database = 'DB01'; EmailAddresses = @("SMTP:$sam@contoso.com"); GrantSendOnBehalfTo = @() } }
    @((& $make 'user01' 'UserMailbox'), (& $make 'compta' 'SharedMailbox')) }
function global:Get-MailboxPermission { param($Identity, $DomainController) @([pscustomobject]@{ User = 'CONTOSO\user01'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }) }
function global:Get-ADPermission { param($Identity, $DomainController) @() }
function global:Get-Recipient { param($Identity, $DomainController) [pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u01'; PrimarySmtpAddress = 'user01@contoso.com'; Name = 'user01'; DistinguishedName = 'CN=user01,OU=PRA,DC=contoso,DC=com' } }
function global:Get-User { param($Identity, $DomainController) [pscustomobject]@{ UserPrincipalName = 'user01@contoso.com' } }
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
        try { $output = & $script:Pwsh -NoProfile -NonInteractive -File $runner 2>&1 | Out-String }
        finally { $ErrorActionPreference = $previous }
        [pscustomobject]@{ Code = $LASTEXITCODE; Text = $output }
    }

    function script:Get-ReportText {
        param([Parameter(Mandatory)][string]$Folder)
        $files = @(Get-ChildItem -LiteralPath (Join-Path $Folder 'logs') -Filter '*.log' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime)
        if (-not $files.Count) { return '' }
        [IO.File]::ReadAllText($files[-1].FullName)
    }

    # One Collect writes the snapshot that the other tests copy; its folder name has a space on purpose.
    $script:Seed = Join-Path $script:Work 'seed folder'
    $script:SeedConfig = New-SnapshotConfig -Folder $script:Seed
    $script:SeedRun = Invoke-Entry -Folder $script:Seed -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-ConfigPath', $script:SeedConfig, '-Force')
    $script:SeedDb = Join-Path $script:Seed 'db\pra.db'

    function script:Copy-Seed {
        <# A copy of the seed tool folder with its snapshot (and the hash files), with a configuration that points at the copy. #>
        param([Parameter(Mandatory)][string]$Name)
        $folder = Join-Path $script:Work $Name
        $null = New-Item -ItemType Directory -Path $folder -Force
        $config = New-SnapshotConfig -Folder $folder
        $null = New-Item -ItemType Directory -Path (Join-Path $folder 'db') -Force
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:Seed 'db') -File | Where-Object { $_.Name -like 'pra.db*' }) { Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $folder 'db') }
        $null = New-Item -ItemType Directory -Path (Join-Path $folder 'backup') -Force
        Get-ChildItem -LiteralPath (Join-Path $script:Seed 'backup') -File | Copy-Item -Destination (Join-Path $folder 'backup')
        [pscustomobject]@{ Folder = $folder; Config = $config; Db = (Join-Path $folder 'db\pra.db'); Hash = (Join-Path $folder 'db\pra.db.sha256') }
    }

    function script:Edit-SnapshotRow {
        <# A real change to the snapshot file that leaves it a valid database. #>
        param([Parameter(Mandatory)][string]$SnapshotFile)
        $connection = Open-PcmStore -Path $SnapshotFile -Root $script:Root
        try { $null = Invoke-PcmSql $connection "UPDATE snapshot SET exchange_server = 'changed.contoso.com'" } finally { Close-PcmStore $connection }
    }
}

AfterAll {
    foreach ($path in @($script:Work, (Split-Path $script:Root -Parent))) { if ($path -and (Test-Path -LiteralPath $path)) { Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue } }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Snapshot hash file: format, reading and checking' {
    BeforeAll {
        $script:folder = Join-Path $TestDrive 'path with spaces'
        $null = New-Item -ItemType Directory -Path $script:folder
        $script:file = Join-Path $script:folder 'Snapshot File.db'
        [IO.File]::WriteAllBytes($script:file, [byte[]](1..200))
        $script:expected = (Get-FileHash -LiteralPath $script:file -Algorithm SHA256).Hash.ToLowerInvariant()
    }
    It 'writes the .sha256 file beside a file as the lower-case hash, two spaces, the file name and a line end' {
        $hashFile = Write-PcmHashFile -Path $script:file
        $hashFile | Should -Be ($script:file + '.sha256')
        $text = [IO.File]::ReadAllText($hashFile)
        $text | Should -BeExactly ("{0}  Snapshot File.db`n" -f $script:expected)
        [IO.File]::ReadAllBytes($hashFile)[0] | Should -Not -Be 0xEF
        @(Get-ChildItem -LiteralPath $script:folder -Filter '*.tmp').Count | Should -Be 0
    }
    It 'gives the same hash as Get-FileHash, for a file that another process has open for writing' {
        $stream = [IO.File]::Open($script:file, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::ReadWrite)
        try { Get-PcmFileHash -Path $script:file | Should -BeExactly $script:expected } finally { $stream.Dispose() }
    }
    It 'replaces an older hash file whole' {
        [IO.File]::WriteAllText($script:file + '.sha256', 'old')
        $null = Write-PcmHashFile -Path $script:file
        (Read-PcmHashFile -Path ($script:file + '.sha256')).Hash | Should -BeExactly $script:expected
    }
    It 'reads <Case>, whatever the letter case, the line ending and the byte order mark' -TestCases @(
        @{ Case = 'two spaces and the name'; Text = "{0}  Snapshot File.db`n" }
        @{ Case = 'the binary marker'; Text = "{0} *Snapshot File.db" }
        @{ Case = 'upper-case hex and an upper-case name'; Text = "{0}  SNAPSHOT FILE.DB`r`n"; Upper = $true }
        @{ Case = 'the hash alone'; Text = '{0}' }
        @{ Case = 'a byte order mark'; Text = "$([char]0xFEFF){0}  x.db`r`n" }) {
        param($Case, $Text, $Upper)
        $hex = if ($Upper) { $script:expected.ToUpperInvariant() } else { $script:expected }
        $path = Join-Path $script:folder 'read.sha256'
        [IO.File]::WriteAllText($path, ($Text -f $hex))
        $read = Read-PcmHashFile -Path $path
        $read.Valid | Should -BeTrue
        $read.Hash | Should -BeExactly $script:expected
    }
    It 'does not accept a file without a 64-digit hash' -TestCases @(@{ Text = '' }, @{ Text = 'not a hash  x.db' }, @{ Text = ('ab' * 31) + '  x.db' }, @{ Text = ('zz' * 32) + '  x.db' }) {
        param($Text)
        $path = Join-Path $script:folder 'bad.sha256'
        [IO.File]::WriteAllText($path, $Text)
        (Read-PcmHashFile -Path $path).Valid | Should -BeFalse
    }
    It 'says Verified when the file matches its hash, in any letter case of the hash and of the name' -TestCases @(@{ Upper = $false }, @{ Upper = $true }) {
        param($Upper)
        $hex = if ($Upper) { $script:expected.ToUpperInvariant() } else { $script:expected }
        [IO.File]::WriteAllText($script:file + '.sha256', "$hex  SNAPSHOT file.DB`n")
        $state = Test-PcmSnapshotHash -Path $script:file
        $state.State | Should -Be 'Verified'
        $state.Expected | Should -BeExactly $script:expected
        $state.Actual | Should -BeExactly $script:expected
    }
    It 'says Changed, with the exact sentence, when the file no longer matches' {
        $null = Write-PcmHashFile -Path $script:file
        $before = [IO.File]::ReadAllBytes($script:file)
        $changed = $before.Clone(); $changed[10] = 99
        [IO.File]::WriteAllBytes($script:file, $changed)
        $state = Test-PcmSnapshotHash -Path $script:file -BackupFolder 'D:\PRA\backup'
        [IO.File]::WriteAllBytes($script:file, $before)
        $state.State | Should -Be 'Changed'
        $state.Actual | Should -Not -BeExactly $state.Expected
        $state.Message | Should -BeExactly ("The snapshot file {0} was changed after its last Collect: its SHA-256 is {1} and the file Snapshot File.db.sha256 written by Collect says {2}. Nothing was changed. Restore the file from your external backup (the copy in D:\PRA\backup has its own .sha256), or, if you know why it changed and accept it, run the same command again with -AcceptUnverifiedSnapshot." -f $script:file, $state.Actual, $state.Expected)
    }
    It 'says NotVerifiable, with a warning and no stop, when there is no hash file' {
        Remove-PcmHashFile -Path $script:file
        Test-Path -LiteralPath ($script:file + '.sha256') | Should -BeFalse
        $state = Test-PcmSnapshotHash -Path $script:file
        $state.State | Should -Be 'NotVerifiable'
        $state.Message | Should -BeExactly ("The snapshot file {0} has no Snapshot File.db.sha256 beside it, so the tool cannot prove it is the file written by Collect (old snapshot, or the file was copied without it). Copy the .sha256 file with it next time." -f $script:file)
    }
    It 'stops (Changed) on a hash file that holds no hash, and says so in plain words' {
        [IO.File]::WriteAllText($script:file + '.sha256', 'garbage')
        $state = Test-PcmSnapshotHash -Path $script:file
        $state.State | Should -Be 'Changed'
        $state.Message | Should -BeLike '*cannot be verified*does not hold a SHA-256 hash*Nothing was changed*-AcceptUnverifiedSnapshot*'
    }
}

Describe 'Collect writes the hash files' {
    It 'ends with RESULT: PASS and writes the snapshot' {
        $script:SeedRun.Code | Should -Be 0 -Because $script:SeedRun.Text
        Test-Path -LiteralPath $script:SeedDb | Should -BeTrue
    }
    It 'writes the .sha256 of the live snapshot file and of its backup copy, matching the files, in a folder with spaces' {
        $hashFile = $script:SeedDb + '.sha256'
        Test-Path -LiteralPath $hashFile | Should -BeTrue
        $read = Read-PcmHashFile -Path $hashFile
        $read.FileName | Should -Be 'pra.db'
        $read.Hash | Should -BeExactly (Get-FileHash -LiteralPath $script:SeedDb -Algorithm SHA256).Hash.ToLowerInvariant()
        $backup = @(Get-ChildItem -LiteralPath (Join-Path $script:Seed 'backup') -Filter '*.db')
        $backup.Count | Should -Be 1
        $backupRead = Read-PcmHashFile -Path ($backup[0].FullName + '.sha256')
        $backupRead.FileName | Should -Be $backup[0].Name
        $backupRead.Hash | Should -BeExactly (Get-FileHash -LiteralPath $backup[0].FullName -Algorithm SHA256).Hash.ToLowerInvariant()
        $script:SeedRun.Text | Should -Match 'SHA-256 file written'
    }
    It 'leaves no temporary file and no SQLite journal beside the snapshot file' {
        @(Get-ChildItem -LiteralPath (Join-Path $script:Seed 'db') -File | Where-Object { $_.Name -notin @('pra.db', 'pra.db.sha256') }).Count | Should -Be 0
    }
    It 'writes nothing in Preview' {
        $copy = Copy-Seed 'preview'
        Remove-Item -LiteralPath $copy.Db, $copy.Hash -Force
        Remove-Item -LiteralPath (Join-Path $copy.Folder 'backup') -Recurse -Force
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Collect', '-ConfigPath', $copy.Config)
        $run.Code | Should -Be 0 -Because $run.Text
        Test-Path -LiteralPath $copy.Db | Should -BeFalse
        Test-Path -LiteralPath $copy.Hash | Should -BeFalse
    }
    It 'refreshes the hash files at the next Collect and leaves a matching pair' {
        $copy = Copy-Seed 'recollect'
        $old = (Read-PcmHashFile -Path $copy.Hash).Hash
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Collect', '-Mode', 'Apply', '-ConfigPath', $copy.Config, '-Force')
        $run.Code | Should -Be 0 -Because $run.Text
        $new = (Read-PcmHashFile -Path $copy.Hash).Hash
        $new | Should -Not -BeExactly $old
        $new | Should -BeExactly (Get-FileHash -LiteralPath $copy.Db -Algorithm SHA256).Hash.ToLowerInvariant()
        (Test-PcmSnapshotHash -Path $copy.Db).State | Should -Be 'Verified'
        @(Get-ChildItem -LiteralPath (Join-Path $copy.Folder 'backup') -Filter '*.db.sha256').Count | Should -Be 2
    }
}

Describe 'Check, Convert and Recover check the snapshot file before they read it' {
    It 'stops Check with exit code 1 and nothing changed when the snapshot file was changed after its Collect (<Spelling>)' -TestCases @(@{ Spelling = 'official' }, @{ Spelling = 'lower case' }) {
        param($Spelling)
        $copy = Copy-Seed ('changed-check-' + $Spelling.Replace(' ', ''))
        Edit-SnapshotRow -SnapshotFile $copy.Db
        $arguments = if ($Spelling -eq 'lower case') { @('-action', 'check', '-configpath', $copy.Config) } else { @('-Action', 'Check', '-ConfigPath', $copy.Config) }
        $run = Invoke-Entry -Folder $copy.Folder -Arguments $arguments
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match ([regex]::Escape(('The snapshot file {0} was changed after its last Collect' -f $copy.Db)))
        $run.Text | Should -Match 'written by Collect says'
        $run.Text | Should -Match 'Nothing was changed\. Restore the file from your external backup'
        $run.Text | Should -Match 'run the same command again with -AcceptUnverifiedSnapshot'
        $run.Text | Should -Not -Match 'Microsoft 365 connection.*(Connected|Connecting)'
        Test-Path -LiteralPath (Join-Path $copy.Folder 'db\journal.db') | Should -BeFalse
        $run.Text | Should -Match 'Check only reads: nothing was changed'
    }
    It 'stops Convert -Mode Apply before anything is written, with the same message' {
        $copy = Copy-Seed 'changed-convert'
        Edit-SnapshotRow -SnapshotFile $copy.Db
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Convert', '-Mode', 'Apply', '-Force', '-ConfigPath', $copy.Config)
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'was changed after its last Collect'
        $journal = Join-Path $copy.Folder 'db\journal.db'
        if (Test-Path -LiteralPath $journal) {
            $connection = Open-PcmStore -Path $journal -Root $script:Root -Kind Journal -ReadOnly
            try { [int](Invoke-PcmSql $connection 'SELECT COUNT(*) FROM batch' -As Scalar) | Should -Be 0 } finally { Close-PcmStore $connection }
        }
    }
    It 'checks the file in the one place where Check, Check -Batch, Check -ResolveOnly, Convert (and its resume) and Recover read the snapshot, before the database is opened' {
        $ast = [Management.Automation.Language.Parser]::ParseFile($script:Entry, [ref]$null, [ref]$null)
        $functions = @{}
        foreach ($function in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $functions[$function.Name] = $function.Extent.Text }
        foreach ($name in 'Invoke-PcmCheck', 'Invoke-PcmResolveGroup', 'Invoke-PcmCheckBatch', 'Invoke-PcmConvert', 'Invoke-PcmRecover') { $functions[$name] | Should -Match 'Read-PcmSnapshotData' -Because $name }
        $read = $functions['Read-PcmSnapshotData']
        $read.IndexOf('Confirm-PcmSnapshotHash') | Should -BeGreaterThan 0
        $read.IndexOf('Confirm-PcmSnapshotHash') | Should -BeLessThan $read.IndexOf('Open-PcmStore')
        $others = @($functions.Keys | Where-Object { $_ -notin @('Read-PcmSnapshotData', 'Invoke-PcmCollect', 'Get-PcmCollectSizeProblem') -and $functions[$_] -match 'Open-PcmStore[^\r\n]*Store\.Path' })
        $others | Should -BeNullOrEmpty
    }
    It 'stops when the hash file was altered instead of the snapshot' {
        $copy = Copy-Seed 'altered-hash'
        [IO.File]::WriteAllText($copy.Hash, ('{0}  pra.db' -f ('ab' * 32)))
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Check', '-ConfigPath', $copy.Config)
        $run.Code | Should -Be 1 -Because $run.Text
        $run.Text | Should -Match 'was changed after its last Collect'
    }
    It 'goes on with a warning when there is no hash file (an old snapshot)' {
        $copy = Copy-Seed 'no-hash'
        Remove-Item -LiteralPath $copy.Hash -Force
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Check', '-ConfigPath', $copy.Config)
        $run.Text | Should -Not -Match 'was changed after its last Collect'
        $log = Get-ReportText -Folder $copy.Folder
        $log | Should -Match ([regex]::Escape(('The snapshot file {0} has no pra.db.sha256 beside it, so the tool cannot prove it is the file written by Collect' -f $copy.Db)))
        $log | Should -Match 'Snapshot 1 of'
    }
    It 'reads a verified snapshot and says so' {
        $copy = Copy-Seed 'verified'
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Check', '-ConfigPath', $copy.Config)
        $run.Text | Should -Match 'Snapshot file verified'
        $run.Text | Should -Not -Match 'was changed after its last Collect'
    }
    It 'accepts a changed snapshot with -AcceptUnverifiedSnapshot (<Spelling>), says so everywhere and keeps the switch in the next steps' -TestCases @(@{ Spelling = '-AcceptUnverifiedSnapshot' }, @{ Spelling = '-acceptunverifiedsnapshot' }) {
        param($Spelling)
        $copy = Copy-Seed ('accepted-' + $Spelling.Substring(1, 3))
        Edit-SnapshotRow -SnapshotFile $copy.Db
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Check', '-ConfigPath', $copy.Config, $Spelling)
        $run.Text | Should -Not -Match 'was changed after its last Collect: its SHA-256'
        $run.Text | Should -Match 'Integrity +Skipped on request \(-AcceptUnverifiedSnapshot\)'
        $log = Get-ReportText -Folder $copy.Folder
        $log | Should -Match 'Snapshot check skipped on request \(-AcceptUnverifiedSnapshot\): the file does not match its \.sha256\.'
        $log | Should -Match 'Snapshot 1 of'
        $run.Text | Should -Match 'Integrity +skipped on request \(-AcceptUnverifiedSnapshot\)'
        $run.Text | Should -Match '-AcceptUnverifiedSnapshot'
    }
    It 'ignores -AcceptUnverifiedSnapshot for Collect and keeps it out of the Collect command line' {
        $copy = Copy-Seed 'collect-switch'
        $run = Invoke-Entry -Folder $copy.Folder -Arguments @('-Action', 'Collect', '-ConfigPath', $copy.Config, '-AcceptUnverifiedSnapshot')
        $run.Code | Should -Be 0 -Because $run.Text
        $run.Text | Should -Not -Match 'Integrity +(verified|skipped|failed|not verifiable)'
        ($run.Text -split "`n" | Where-Object { $_ -match 'Action Collect -Mode Apply' }) -join ' ' | Should -Not -Match 'AcceptUnverified'
    }
}

Describe 'The console shows the state of the snapshot file' {
    BeforeAll {
        $script:modules = @([pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = [version]'2.41.1' }, [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = [version]'3.10.1' })
        $script:db = Join-Path $TestDrive 'ui\snap.db'
        $null = New-Item -ItemType Directory -Path (Split-Path $script:db) -Force
        [IO.File]::WriteAllBytes($script:db, [byte[]](1..64))
        function script:Get-SnapshotLight {
            $state = [pscustomobject]@{ Config = @{ Store = @{ Path = $script:db; BackupFolder = (Join-Path $TestDrive 'ui\backup'); MaxSnapshotAgeDays = 7 } }; SnapshotExists = $true; SnapshotError = ''
                Snapshot = [pscustomobject]@{ id = 3; finished_utc = [datetime]::UtcNow.AddHours(-2).ToString('o') } }
            $info = & $script:Gui { param($s) Get-PcmGuiSnapshotInfo -State $s } $state
            $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.5.1'; Modules = $script:modules; SignIn = 'Interactive'; Tenant = '11111111-1111-1111-1111-111111111111'; Snapshot = $info
                Task = [pscustomobject]@{ Found = $false; Text = '' } }
            [pscustomobject]@{ Info = $info; Light = @($detect.Lights | Where-Object { $_.Name -eq 'Snapshot' })[0] }
        }
    }
    It 'shows Verified (green) when the file matches its hash' {
        $null = Write-PcmHashFile -Path $script:db
        $r = Get-SnapshotLight
        $r.Info.Integrity | Should -Be 'Verified'
        $r.Light.Status | Should -Be 'Ok'
        $r.Light.Detail | Should -BeLike '*Verified: it matches its .sha256 written by Collect.*'
    }
    It 'shows Not verifiable (amber) when there is no hash file' {
        Remove-PcmHashFile -Path $script:db
        $r = Get-SnapshotLight
        $r.Info.Integrity | Should -Be 'Not verifiable'
        $r.Light.Status | Should -Be 'Warn'
        $r.Light.Detail | Should -BeLike '*no .sha256 beside it*cannot prove it is the file written by Collect*'
    }
    It 'shows Changed (red) with the remedy when the file does not match its hash' {
        $null = Write-PcmHashFile -Path $script:db
        [IO.File]::WriteAllBytes($script:db, [byte[]](1..65))
        $r = Get-SnapshotLight
        $r.Info.Integrity | Should -Be 'Changed'
        $r.Light.Status | Should -Be 'Fail'
        $r.Light.Detail | Should -BeLike '*was changed after its last Collect*Restore the snapshot file from your external backup*'
    }
    It 'shows Changed too when the hash file holds no hash' {
        [IO.File]::WriteAllText($script:db + '.sha256', 'nothing')
        (Get-SnapshotLight).Light.Status | Should -Be 'Fail'
    }
}
