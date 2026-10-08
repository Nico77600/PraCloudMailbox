<#
.SYNOPSIS
    PRA Cloud Mailbox - offline tests (Pester 5+). No Exchange, Active Directory or cloud access: the
    Exchange cmdlets are synthetic functions, the cloud rules are pure functions.
.DESCRIPTION
    Run in both editions (Collect = Windows PowerShell 5.1, cloud actions = PowerShell 7):
        powershell.exe -NoProfile -File .\tests\Invoke-TestGate.ps1
        pwsh -NoProfile -File .\tests\Invoke-TestGate.ps1
.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
#>
#Requires -Version 5.1

BeforeAll {
    $script:Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    Import-Module (Join-Path $script:Root 'module\PRA2.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA2.Store.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA2.Collect.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA2.Cloud.psm1') -Force

    function script:New-TestConfig {
        param([hashtable]$Override = @{})
        $config = @{
            Environment = 'TEST'
            Exchange = @{ ConnectionMode = 'Local'; Server = ''; DomainController = '' }
            Scope = @{ Mode = 'OU'; SearchBase = 'OU=PRA,DC=contoso,DC=com'; GroupDN = ''; CsvPath = ''; IncludeUsers = $true; IncludeShared = $true; IncludeRoom = $false; IncludeEquipment = $false; ExcludeSamAccountNames = @() }
            Collect = @{ SharedPermissions = $true; ExpandGroupTrustees = $true; ExcludeTrustees = @('Administrator'); MailboxStatistics = $false; Contacts = $false; ContactsSearchBase = ''; DistributionGroups = $false; DistributionGroupsSearchBase = ''; DynamicDistributionGroups = $false }
            Store = @{ Path = '.\data\test.db'; KeepSnapshots = 10; BackupFolder = '.\data\backup'; MaxSnapshotAgeDays = 7; JournalPath = '.\data\journal.db' }
            Cloud = @{ TenantId = ''; Organization = ''; AppId = ''; CertificateThumbprint = ''; DefaultUsageLocation = '' }
            Polling = @{ IntervalSeconds = 20; MailboxTimeoutMinutes = 30; SyncTimeoutMinutes = 30; HoldTimeoutMinutes = 30 }
            Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111'; SkuPartNumber = '' }; Shared = @{ SkuPartNumber = 'E5' } }
            Retention = @{ TagAttribute = 'CustomAttribute1'; TagValue = 'Converted'; HoldPolicy = '' }
            EntraConnect = @{ Mode = 'Manual'; Server = ''; ScriptPath = ''; MinVersion = '2.5.76.0' }
            Logging = @{ Folder = '.\logs' }
            Report = @{ Enabled = $true; Folder = '.\reports' }
        }
        foreach ($key in $Override.Keys) { $config[$key] = $Override[$key] }
        return $config
    }
    function script:ConvertTo-Psd1 {
        param([object]$Value, [int]$Indent = 0)
        $pad = ' ' * $Indent
        if ($Value -is [hashtable]) {
            $lines = @('@{')
            foreach ($key in $Value.Keys) { $lines += ('{0}    {1} = {2}' -f $pad, $key, (ConvertTo-Psd1 $Value[$key] ($Indent + 4))) }
            return ($lines + ($pad + '}')) -join "`r`n"
        }
        if ($Value -is [array]) { return '@(' + (($Value | ForEach-Object { ConvertTo-Psd1 $_ }) -join ', ') + ')' }
        if ($Value -is [bool]) { return $(if ($Value) { '$true' } else { '$false' }) }
        if ($Value -is [int]) { return [string]$Value }
        return "'" + ([string]$Value).Replace("'", "''") + "'"
    }
    function script:New-TestContext {
        param([hashtable]$Config = (New-TestConfig))
        return @{ Config = $Config; Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]')
            Excluded = (New-Object 'Collections.Generic.List[string]'); LogFile = ''; Warnings = 0; Server = '' }
    }

    # Synthetic Exchange cmdlets (Collect is tested without Exchange).
    function global:Get-Mailbox { param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController) }
    function global:Get-MailboxStatistics { param($Identity, $DomainController) }
    function global:Get-MailboxPermission { param($Identity, $DomainController) }
    function global:Get-ADPermission { param($Identity, $DomainController) }
    function global:Get-Recipient { param($Identity, $DomainController) }
    function global:Get-User { param($Identity, $DomainController) }
    function global:Get-Group { param($Identity, $DomainController) }
    function global:Get-MailContact { param($ResultSize, $OrganizationalUnit, $DomainController) }
    function global:Get-DistributionGroup { param($ResultSize, $OrganizationalUnit, $DomainController) }
    function global:Get-DistributionGroupMember { param($Identity, $ResultSize, $DomainController) }
    function global:Get-DynamicDistributionGroup { param($ResultSize, $OrganizationalUnit, $DomainController) }
    function global:Get-ExchangeServer { param() }

    function script:New-FakeMailbox {
        param([string]$Sam, [string]$Type = 'UserMailbox', [string]$Guid = ([guid]::NewGuid().ToString()), [object[]]$OnBehalf = @())
        return [pscustomobject]@{
            Guid = $Guid; Identity = $Sam; Name = $Sam; DisplayName = "Display $Sam"; Alias = $Sam; SamAccountName = $Sam
            UserPrincipalName = "$Sam@contoso.com"; DistinguishedName = "CN=$Sam,OU=PRA,DC=contoso,DC=com"; OrganizationalUnit = 'contoso.com/PRA'
            PrimarySmtpAddress = "$Sam@contoso.com"; WindowsEmailAddress = "$Sam@contoso.com"; LegacyExchangeDN = "/o=Contoso/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=$Sam"
            ExchangeGuid = [guid]::NewGuid().ToString(); ArchiveGuid = [guid]::Empty.ToString(); ArchiveStatus = 'None'
            RecipientTypeDetails = $Type; HiddenFromAddressListsEnabled = ($Type -ne 'UserMailbox'); Database = 'DB01'
            EmailAddresses = @("SMTP:$Sam@contoso.com", "smtp:$Sam@contoso.mail.onmicrosoft.com", "X500:/o=Old/cn=$Sam")
            GrantSendOnBehalfTo = $OnBehalf; CustomAttribute1 = ''; CustomAttribute2 = 'Paris'
        }
    }
}

AfterAll {
    foreach ($name in @('Get-Mailbox','Get-MailboxStatistics','Get-MailboxPermission','Get-ADPermission','Get-Recipient','Get-User','Get-Group','Get-MailContact',
            'Get-DistributionGroup','Get-DistributionGroupMember','Get-DynamicDistributionGroup','Get-ExchangeServer')) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
}

Describe 'Configuration' {
    BeforeEach { $script:folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $script:folder }

    It 'accepts the shipped configuration file' {
        $config = Import-PraConfiguration -Path (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1') -Root $script:Root
        $config.Store.Path | Should -Be (Join-Path $script:Root 'data\PraCloudMailbox.db')
        $config.Licensing.Users.Mode | Should -Be 'Group'
    }
    It 'adds the default values of missing settings' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        $config = Import-PraConfiguration -Path $path -Root $script:folder
        $config.Store.KeepSnapshots | Should -Be 10
        $config.Collect.ExpandGroupTrustees | Should -BeTrue
        $config.Retention.TagValue | Should -Be 'Converted'
    }
    It 'refuses an unknown setting' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto'; IncludeUser = `$true } }" -Encoding UTF8
        { Import-PraConfiguration -Path $path -Root $script:folder } | Should -Throw '*unknown setting Scope.IncludeUser*'
    }
    It 'refuses <Case>' -ForEach @(
        @{ Case = 'OU mode without SearchBase'; Text = "Scope = @{ Mode = 'OU'; SearchBase = '' }"; Message = '*requires Scope.SearchBase*' }
        @{ Case = 'Remote mode without server'; Text = "Exchange = @{ ConnectionMode = 'Remote' }; Scope = @{ Mode = 'Auto' }"; Message = '*requires Exchange.Server*' }
        @{ Case = 'an AppId without certificate'; Text = "Scope = @{ Mode = 'Auto' }; Cloud = @{ AppId = 'x' }"; Message = '*go together*' }
        @{ Case = 'an invalid tag attribute'; Text = "Scope = @{ Mode = 'Auto' }; Retention = @{ TagAttribute = 'extensionAttribute1' }"; Message = '*CustomAttribute1 to CustomAttribute15*' }
        @{ Case = 'Direct mode without SKU'; Text = "Scope = @{ Mode = 'Auto' }; Licensing = @{ Users = @{ Mode = 'Direct' } }"; Message = '*requires Licensing.Users.SkuPartNumber*' }
        @{ Case = 'no mailbox type'; Text = "Scope = @{ Mode = 'Auto'; IncludeUsers = `$false; IncludeShared = `$false }"; Message = '*at least one mailbox type*' }
        @{ Case = 'a text instead of a number'; Text = "Scope = @{ Mode = 'Auto' }; Store = @{ KeepSnapshots = '5' }"; Message = '*whole number*' }
    ) {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value ("@{ Environment = 'X'; $Text }") -Encoding UTF8
        { Import-PraConfiguration -Path $path -Root $script:folder } | Should -Throw $Message
    }
}

Describe 'SQLite store' {
    BeforeEach {
        $script:db = Join-Path $TestDrive ("store-{0}.db" -f [guid]::NewGuid().ToString('N'))
        $script:cn = Open-Pra2Store -Path $script:db -Root $script:Root
    }
    AfterEach { Close-Pra2Store $script:cn }

    It 'writes and reads a complete snapshot with Unicode values' {
        $id = New-Pra2Snapshot $script:cn @{ RunId = 'r1'; Environment = 'TEST'; ToolVersion = '0.1.0'; ExchangeServer = 'EX1'; ScopeJson = '{}' }
        $rows = @([ordered]@{ object_guid = 'g1'; kind = 'Shared'; user_principal_name = 'compta@contoso.com'; display_name = 'Comptabilité é ü'; hidden_from_address_lists = $true; item_count = $null })
        Add-Pra2Row $script:cn mailbox $id $rows | Should -Be 1
        Set-Pra2SnapshotStatus $script:cn $id Complete @{ mailbox = 1 }
        $snap = Get-Pra2Snapshot $script:cn
        $snap.id | Should -Be $id
        $snap.status | Should -Be 'Complete'
        $read = @(Read-Pra2Table $script:cn mailbox $id)
        $read[0].display_name | Should -Be 'Comptabilité é ü'
        $read[0].hidden_from_address_lists | Should -Be 1
        $read[0].item_count | Should -BeNullOrEmpty
    }
    It 'never returns a Running or Failed snapshot as the last complete one' {
        $first = New-Pra2Snapshot $script:cn @{ RunId = 'r1' }
        Set-Pra2SnapshotStatus $script:cn $first Complete
        $null = New-Pra2Snapshot $script:cn @{ RunId = 'r2' }
        $third = New-Pra2Snapshot $script:cn @{ RunId = 'r3' }
        Set-Pra2SnapshotStatus $script:cn $third Failed
        (Get-Pra2Snapshot $script:cn).id | Should -Be $first
    }
    It 'rejects a value outside the allowed kinds' {
        $id = New-Pra2Snapshot $script:cn @{ RunId = 'r1' }
        { Add-Pra2Row $script:cn mailbox $id @([ordered]@{ object_guid = 'g'; kind = 'Room2' }) } | Should -Throw
        @(Read-Pra2Table $script:cn mailbox $id).Count | Should -Be 0
    }
    It 'rejects an invalid column name (no SQL injection through the keys)' {
        $id = New-Pra2Snapshot $script:cn @{ RunId = 'r1' }
        { Add-Pra2Row $script:cn mailbox $id @([ordered]@{ 'object_guid) --' = 'g'; kind = 'User' }) } | Should -Throw '*Invalid column name*'
    }
    It 'keeps the last N complete snapshots and deletes their rows' {
        foreach ($n in 1..4) {
            $id = New-Pra2Snapshot $script:cn @{ RunId = "r$n" }
            $null = Add-Pra2Row $script:cn permission $id @([ordered]@{ mailbox_guid = 'g'; access_right = 'SendAs'; trustee = 'x'; resolved = $true })
            Set-Pra2SnapshotStatus $script:cn $id Complete
        }
        Remove-Pra2OldSnapshot $script:cn 2 | Should -Be 2
        @(Get-Pra2SnapshotList $script:cn).Count | Should -Be 2
        [int](Invoke-Pra2Sql $script:cn 'SELECT count(*) FROM permission' -As Scalar) | Should -Be 2
    }
    It 'writes a consistent backup copy and refuses to overwrite it' {
        $id = New-Pra2Snapshot $script:cn @{ RunId = 'r1' }
        Set-Pra2SnapshotStatus $script:cn $id Complete
        $backup = Backup-Pra2Store $script:cn (Join-Path $TestDrive 'bk') 'copy.db'
        Test-Path -LiteralPath $backup | Should -BeTrue
        { Backup-Pra2Store $script:cn (Join-Path $TestDrive 'bk') 'copy.db' } | Should -Throw '*already exists*'
        $copy = Open-Pra2Store -Path $backup -Root $script:Root -ReadOnly
        try { (Get-Pra2Snapshot $copy).id | Should -Be $id } finally { Close-Pra2Store $copy }
    }
    It 'opens read-only without creating anything and refuses writes' {
        { Open-Pra2Store -Path (Join-Path $TestDrive 'missing.db') -Root $script:Root -ReadOnly } | Should -Throw '*Database not found*'
        $ro = Open-Pra2Store -Path $script:db -Root $script:Root -ReadOnly
        try { { Invoke-Pra2Sql $ro "INSERT INTO meta (key, value) VALUES ('x', 'y')" } | Should -Throw } finally { Close-Pra2Store $ro }
    }
    It 'refuses a database written by a newer schema' {
        $null = Invoke-Pra2Sql $script:cn "UPDATE meta SET value = '99' WHERE key = 'schema_version'"
        Close-Pra2Store $script:cn
        { Open-Pra2Store -Path $script:db -Root $script:Root -ReadOnly } | Should -Throw '*newer than this tool*'
        $script:cn = Open-Pra2Store -Path (Join-Path $TestDrive ("other-{0}.db" -f [guid]::NewGuid().ToString('N'))) -Root $script:Root
    }
}

Describe 'Collect: mailboxes' {
    It 'maps a mailbox to a row of the mailbox table' {
        $mailbox = New-FakeMailbox -Sam 'compta' -Type 'SharedMailbox' -OnBehalf @('CN=user01,OU=PRA,DC=contoso,DC=com')
        $record = ConvertTo-Pra2MailboxRecord -Mailbox $mailbox -Statistics ([pscustomobject]@{ ItemCount = 42; TotalItemSize = '1.5 MB (1,572,864 bytes)' })
        $record.kind | Should -Be 'Shared'
        $record.immutable_id | Should -Be ([Convert]::ToBase64String(([guid]$mailbox.Guid).ToByteArray()))
        @(($record.email_addresses_json | ConvertFrom-Json) | ForEach-Object { $_ }) | Should -Contain 'X500:/o=Old/cn=compta'
        ($record.custom_attributes_json | ConvertFrom-Json).CustomAttribute2 | Should -Be 'Paris'
        $record.archive_guid | Should -BeNullOrEmpty
        $record.item_count | Should -Be 42
        $record.total_item_size_bytes | Should -Be 1572864
        @(($record.grant_send_on_behalf_json | ConvertFrom-Json) | ForEach-Object { $_ }).Count | Should -Be 1
    }
    It 'refuses an unsupported mailbox type' {
        { ConvertTo-Pra2MailboxRecord -Mailbox (New-FakeMailbox -Sam 'x' -Type 'LinkedMailbox') } | Should -Throw '*Unsupported mailbox type*'
    }
    It 'applies -Scope to the configured mailbox types' {
        $scope = (New-TestConfig).Scope
        Get-Pra2RecipientType -Scope $scope -Filter All | Should -Be @('UserMailbox','SharedMailbox')
        Get-Pra2RecipientType -Scope $scope -Filter UsersOnly | Should -Be @('UserMailbox')
        Get-Pra2RecipientType -Scope $scope -Filter SharedOnly | Should -Be @('SharedMailbox')
    }
    It 'reads the OU, excludes system and configured mailboxes, removes duplicates and sorts' {
        $config = New-TestConfig
        $config.Scope.ExcludeSamAccountNames = @('skip')
        $context = New-TestContext $config
        $dup = New-FakeMailbox -Sam 'bob'
        Mock -ModuleName PRA2.Collect Get-Mailbox { @((New-FakeMailbox -Sam 'zoe'), $dup, $dup, (New-FakeMailbox -Sam 'HealthMailbox0123'), (New-FakeMailbox -Sam 'skip')) }
        $result = @(Get-Pra2ScopeMailbox -Context $context)
        ($result | ForEach-Object SamAccountName) | Should -Be @('bob','zoe')
        $context.Excluded.Count | Should -Be 2
        Should -Invoke -ModuleName PRA2.Collect Get-Mailbox -Times 1 -ParameterFilter { $OrganizationalUnit -eq 'OU=PRA,DC=contoso,DC=com' -and $ResultSize -eq 'Unlimited' }
    }
    It 'refuses -Identity on a mailbox type outside the scope' {
        $config = New-TestConfig
        $config.Scope.IncludeShared = $false
        Mock -ModuleName PRA2.Collect Get-Mailbox { New-FakeMailbox -Sam 'compta' -Type 'SharedMailbox' }
        { Get-Pra2ScopeMailbox -Context (New-TestContext $config) -Identity 'compta' } | Should -Throw '*not a mailbox type in scope*'
    }
    It 'expands a scope group recursively and survives a membership loop' {
        $config = New-TestConfig
        $config.Scope.Mode = 'Group'; $config.Scope.GroupDN = 'CN=G1'
        Mock -ModuleName PRA2.Collect Get-Group {
            switch ($Identity) {
                'CN=G1' { [pscustomobject]@{ Members = @('CN=u1', 'CN=G2') } }
                'CN=G2' { [pscustomobject]@{ Members = @('CN=u2', 'CN=G1') } }
                default { throw 'not a group' }
            }
        }
        Mock -ModuleName PRA2.Collect Get-Mailbox { New-FakeMailbox -Sam ($Identity -replace 'CN=', '') }
        $result = @(Get-Pra2ScopeMailbox -Context (New-TestContext $config))
        ($result | ForEach-Object SamAccountName) | Should -Be @('u1','u2')
    }
}

Describe 'Collect: shared mailbox permissions' {
    BeforeEach {
        $script:context = New-TestContext
        $script:shared = New-FakeMailbox -Sam 'compta' -Type 'SharedMailbox' -OnBehalf @('CN=user02,OU=PRA,DC=contoso,DC=com')
        Mock -ModuleName PRA2.Collect Get-MailboxPermission {
            @([pscustomobject]@{ User = 'CONTOSO\user01'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess','ReadPermission'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\Domain Admins'; AccessRights = @('FullAccess'); IsInherited = $true; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\Administrator'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\denied'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $true }
              [pscustomobject]@{ User = 'CONTOSO\GRP-Compta'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'S-1-5-21-1-2-3-4567'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false })
        }
        Mock -ModuleName PRA2.Collect Get-ADPermission {
            @([pscustomobject]@{ User = 'CONTOSO\user01'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\user01'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\ghost'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false })
        }
        Mock -ModuleName PRA2.Collect Get-Recipient {
            switch -Wildcard ($Identity) {
                '*user01*' { [pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u01'; PrimarySmtpAddress = 'user01@contoso.com'; Name = 'user01'; DistinguishedName = 'CN=user01,OU=PRA,DC=contoso,DC=com' } }
                '*user02*' { [pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u02'; PrimarySmtpAddress = 'user02@contoso.com'; Name = 'user02'; DistinguishedName = 'CN=user02,OU=PRA,DC=contoso,DC=com' } }
                '*user03*' { [pscustomobject]@{ RecipientType = 'MailUser'; Guid = 'u03'; PrimarySmtpAddress = 'user03@contoso.com'; Name = 'user03'; DistinguishedName = 'CN=user03,OU=PRA,DC=contoso,DC=com' } }
                default { throw 'not found' }
            }
        }
        Mock -ModuleName PRA2.Collect Get-User {
            switch -Wildcard ($Identity) {
                '*user0*' { [pscustomobject]@{ UserPrincipalName = (($Identity -replace '^CN=([^,]+),.*$', '$1') + '@contoso.com'); Guid = 'x'; Name = 'x'; DistinguishedName = $Identity } }
                default { throw 'not found' }
            }
        }
        Mock -ModuleName PRA2.Collect Get-Group {
            switch -Wildcard ($Identity) {
                '*GRP-Compta*' { [pscustomobject]@{ Guid = 'g1'; Name = 'GRP-Compta'; DistinguishedName = 'CN=GRP-Compta,OU=PRA,DC=contoso,DC=com'; Members = @('CN=user03,OU=PRA,DC=contoso,DC=com') } }
                default { throw 'not a group' }
            }
        }
        Mock -ModuleName PRA2.Collect Get-Pra2DelegateListLink { @('CN=user01,OU=PRA,DC=contoso,DC=com') }
    }

    It 'keeps explicit grants only, without system, admin, inherited, denied, duplicate or self entries' {
        $result = Get-Pra2SharedPermission -Context $script:context -Mailbox $script:shared
        $rights = @($result.Permissions | ForEach-Object { '{0} {1}' -f $_.access_right, $_.trustee })
        $rights | Should -Be @('FullAccess CONTOSO\user01', 'FullAccess CONTOSO\GRP-Compta', 'SendAs CONTOSO\user01', 'SendAs CONTOSO\ghost', 'SendOnBehalf CN=user02,OU=PRA,DC=contoso,DC=com')
    }
    It 'resolves trustees to UPN / SMTP and reads AutoMapping' {
        $result = Get-Pra2SharedPermission -Context $script:context -Mailbox $script:shared
        $fullAccess = @($result.Permissions | Where-Object { $_.access_right -eq 'FullAccess' -and $_.trustee_kind -eq 'User' })[0]
        $fullAccess.trustee_upn | Should -Be 'user01@contoso.com'
        $fullAccess.trustee_smtp | Should -Be 'user01@contoso.com'
        $fullAccess.auto_mapping | Should -BeTrue
        @($result.Permissions | Where-Object { $_.access_right -eq 'SendOnBehalf' })[0].trustee_upn | Should -Be 'user02@contoso.com'
    }
    It 'expands a group trustee to its members' {
        $result = Get-Pra2SharedPermission -Context $script:context -Mailbox $script:shared
        @($result.Members).Count | Should -Be 1
        $result.Members[0].group_guid | Should -Be 'g1'
        $result.Members[0].member_upn | Should -Be 'user03@contoso.com'
    }
    It 'reports a trustee that cannot be resolved' {
        $result = Get-Pra2SharedPermission -Context $script:context -Mailbox $script:shared
        $ghost = @($result.Permissions | Where-Object { $_.trustee -eq 'CONTOSO\ghost' })[0]
        $ghost.resolved | Should -BeFalse
        @($result.Warnings | Where-Object { $_ -like '*ghost*' }).Count | Should -Be 1
    }
    It 'excludes trustees with wildcards from Collect.ExcludeTrustees' {
        $script:context.Config.Collect.ExcludeTrustees = @('GRP-*')
        Test-Pra2ExcludedTrustee -Context $script:context -Trustee 'CONTOSO\GRP-Compta' | Should -BeTrue
        Test-Pra2ExcludedTrustee -Context $script:context -Trustee 'CONTOSO\user01' | Should -BeFalse
        Test-Pra2ExcludedTrustee -Context $script:context -Trustee 'CONTOSO\Exchange Trusted Subsystem' | Should -BeTrue
    }
}

Describe 'Collect: end to end (synthetic Exchange)' -Tag 'Desktop' {
    It 'writes a snapshot in Apply and nothing in Preview' -Skip:($PSVersionTable.PSEdition -ne 'Desktop') {
        $folder = Join-Path $TestDrive 'e2e'
        $null = New-Item -ItemType Directory $folder
        $config = New-TestConfig
        $config.Store.Path = (Join-Path $folder 'db\pra.db'); $config.Store.BackupFolder = (Join-Path $folder 'backup')
        $config.Logging.Folder = (Join-Path $folder 'logs'); $config.Report.Folder = (Join-Path $folder 'reports')
        $configPath = Join-Path $folder 'config.psd1'
        Set-Content -LiteralPath $configPath -Value (ConvertTo-Psd1 $config) -Encoding UTF8
        $stubs = @'
function global:Get-ExchangeServer { @([pscustomobject]@{ Name = $env:COMPUTERNAME; AdminDisplayVersion = 'Version 15.2 (Build 1748.10)' }) }
function global:Get-Mailbox { param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController)
    $make = { param($sam, $type) [pscustomobject]@{ Guid = [guid]::NewGuid().ToString(); Identity = $sam; Name = $sam; DisplayName = $sam; Alias = $sam; SamAccountName = $sam
        UserPrincipalName = "$sam@contoso.com"; DistinguishedName = "CN=$sam,OU=PRA,DC=contoso,DC=com"; OrganizationalUnit = 'contoso.com/PRA'; PrimarySmtpAddress = "$sam@contoso.com"
        WindowsEmailAddress = "$sam@contoso.com"; LegacyExchangeDN = "/o=C/cn=$sam"; ExchangeGuid = [guid]::NewGuid().ToString(); ArchiveGuid = [guid]::Empty.ToString(); ArchiveStatus = 'None'
        RecipientTypeDetails = $type; HiddenFromAddressListsEnabled = $false; Database = 'DB01'; EmailAddresses = @("SMTP:$sam@contoso.com"); GrantSendOnBehalfTo = @() } }
    @((& $make 'user01' 'UserMailbox'), (& $make 'compta' 'SharedMailbox')) }
function global:Get-MailboxPermission { param($Identity, $DomainController) @([pscustomobject]@{ User = 'CONTOSO\user01'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }) }
function global:Get-ADPermission { param($Identity, $DomainController) @() }
function global:Get-Recipient { param($Identity, $DomainController) [pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u01'; PrimarySmtpAddress = 'user01@contoso.com'; Name = 'user01'; DistinguishedName = 'CN=user01,OU=PRA,DC=contoso,DC=com' } }
function global:Get-User { param($Identity, $DomainController) [pscustomobject]@{ UserPrincipalName = 'user01@contoso.com' } }
'@
        $runner = Join-Path $folder 'run.ps1'
        Set-Content -LiteralPath $runner -Encoding UTF8 -Value @"
`$ErrorActionPreference = 'Stop'
$stubs
& '$(Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')' -Action Collect -Mode `$args[0] -ConfigPath '$configPath' -Force -PassThru
"@
        $preview = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner Preview 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($preview | Out-String)
        Test-Path -LiteralPath $config.Store.Path | Should -BeFalse
        $apply = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner Apply 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($apply | Out-String)
        $cn = Open-Pra2Store -Path $config.Store.Path -Root $script:Root -ReadOnly
        try {
            $snap = Get-Pra2Snapshot $cn
            $snap.status | Should -Be 'Complete'
            @(Read-Pra2Table $cn mailbox $snap.id).Count | Should -Be 2
            @(Read-Pra2Table $cn permission $snap.id | Where-Object { $_.trustee_upn -eq 'user01@contoso.com' }).Count | Should -Be 1
        } finally { Close-Pra2Store $cn }
        @(Get-ChildItem -LiteralPath $config.Store.BackupFolder -Filter '*.db').Count | Should -Be 1
        @(Get-ChildItem -LiteralPath $config.Report.Folder -Filter '*.html').Count | Should -Be 2
        (Get-Content -LiteralPath (Get-ChildItem -LiteralPath $config.Logging.Folder -Filter '*.log' | Sort-Object LastWriteTime | Select-Object -Last 1).FullName -Raw) | Should -Match 'RESULT: PASS'
    }
}

Describe 'Cloud rules: object readiness' {
    BeforeAll {
        $script:config = New-TestConfig
        $script:userRecord = [pscustomobject]@{ kind = 'User'; object_guid = 'g'; exchange_guid = 'aaaaaaaa-0000-0000-0000-000000000001'; immutable_id = 'x'; user_principal_name = 'u@contoso.com' }
        $script:sharedRecord = [pscustomobject]@{ kind = 'Shared'; object_guid = 's'; exchange_guid = 'bbbbbbbb-0000-0000-0000-000000000002'; immutable_id = 'y'; user_principal_name = 's@contoso.com' }
        function script:New-State {
            param([hashtable]$User = @{}, [hashtable]$Other = @{})
            if ($Other.ContainsKey('User') -and $null -eq $Other['User']) { $u = $null }
            else { $u = @{ id = 'id1'; onPremisesSyncEnabled = $true; onPremisesImmutableId = 'x'; usageLocation = 'FR'; assignedPlans = @(); assignedLicenses = @(); serviceProvisioningErrors = @() }
            foreach ($k in $User.Keys) { $u[$k] = $User[$k] }; $u = [pscustomobject]$u }
            $s = @{ MatchedBy = 'immutableId'; User = $u; IsCloudManaged = $false; RecipientType = 'MailUser'; ExchangeGuid = 'aaaaaaaa-0000-0000-0000-000000000001'; Holds = @(); Locations = @(); TagValue = '' }
            foreach ($k in $Other.Keys) { if ($k -ne 'User') { $s[$k] = $Other[$k] } }
            return [pscustomobject]$s
        }
        function script:Get-Codes { param($Findings, [string]$Level) @($Findings | Where-Object { -not $Level -or $_.Level -eq $Level } | ForEach-Object Code) }
    }

    It 'a synchronised MailUser with the on-premises GUID is ready' {
        $findings = Get-Pra2Readiness -Config $script:config -Record $script:userRecord -State (New-State)
        Get-Codes $findings 'Error' | Should -BeNullOrEmpty
        Get-Codes $findings 'Ok' | Should -Contain 'EXO_MAILUSER'
        Get-Codes $findings | Should -Contain 'NO_TEAMS_STORAGE'
    }
    It 'announces the promotion of the Teams storage' {
        $state = New-State -Other @{ Locations = @([pscustomobject]@{ Type = 'ComponentShared'; Guid = 'c' }) }
        Get-Codes (Get-Pra2Readiness -Config $script:config -Record $script:userRecord -State $state) | Should -Contain 'TEAMS_STORAGE'
    }
    It 'blocks <Case>' -ForEach @(
        @{ Case = 'a missing Entra ID object'; UserProps = @{}; StateProps = @{ User = $null }; Code = 'ENTRA_NOT_FOUND' }
        @{ Case = 'a cloud-only object'; UserProps = @{ onPremisesSyncEnabled = $null }; StateProps = @{}; Code = 'ENTRA_NOT_SYNCED' }
        @{ Case = 'a missing Exchange Online recipient'; UserProps = @{}; StateProps = @{ RecipientType = '' }; Code = 'EXO_NOT_FOUND' }
        @{ Case = 'an existing cloud mailbox'; UserProps = @{}; StateProps = @{ RecipientType = 'UserMailbox' }; Code = 'ALREADY_CLOUD_MAILBOX' }
        @{ Case = 'an Exchange plan already enabled'; UserProps = @{ assignedPlans = @([pscustomobject]@{ service = 'exchange'; capabilityStatus = 'Enabled' }) }; StateProps = @{}; Code = 'EXCHANGE_PLAN_PRESENT' }
    ) {
        $state = New-State -User $UserProps -Other $StateProps
        Get-Codes (Get-Pra2Readiness -Config $script:config -Record $script:userRecord -State $state) 'Error' | Should -Contain $Code
    }
    It 'warns about <Case>' -ForEach @(
        @{ Case = 'an object already cloud-managed'; UserProps = @{}; StateProps = @{ IsCloudManaged = $true }; Code = 'ALREADY_CLOUD_MANAGED' }
        @{ Case = 'a GUID already cleared'; UserProps = @{}; StateProps = @{ ExchangeGuid = '00000000-0000-0000-0000-000000000000' }; Code = 'GUID_CLEARED' }
        @{ Case = 'a GUID that differs'; UserProps = @{}; StateProps = @{ ExchangeGuid = 'cccccccc-0000-0000-0000-000000000003' }; Code = 'GUID_MISMATCH' }
        @{ Case = 'an object found by UPN only'; UserProps = @{}; StateProps = @{ MatchedBy = 'UPN' }; Code = 'ENTRA_MATCHED_BY_UPN' }
        @{ Case = 'a hold on a user'; UserProps = @{}; StateProps = @{ Holds = @('UniH1234') }; Code = 'HOLD_PRESENT' }
        @{ Case = 'a tagged user'; UserProps = @{}; StateProps = @{ TagValue = 'Converted' }; Code = 'USER_TAGGED' }
        @{ Case = 'provisioning errors'; UserProps = @{ serviceProvisioningErrors = @([pscustomobject]@{ isResolved = $false }) }; StateProps = @{}; Code = 'PROVISIONING_ERROR' }
    ) {
        $state = New-State -User $UserProps -Other $StateProps
        $findings = Get-Pra2Readiness -Config $script:config -Record $script:userRecord -State $state
        Get-Codes $findings 'Warn' | Should -Contain $Code
        Get-Codes $findings 'Error' | Should -BeNullOrEmpty
    }
    It 'only informs about a hold on a shared mailbox (it is wanted before the identity is deleted)' {
        $state = New-State -Other @{ Holds = @('UniH1234'); ExchangeGuid = 'bbbbbbbb-0000-0000-0000-000000000002' }
        $findings = Get-Pra2Readiness -Config $script:config -Record $script:sharedRecord -State $state
        @($findings | Where-Object Code -eq 'HOLD_PRESENT')[0].Level | Should -Be 'Info'
    }
    It 'refuses a <Kind> mailbox (Convert would turn it into a shared mailbox)' -ForEach @(@{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        $record = [pscustomobject]@{ kind = $Kind; object_guid = 'r'; exchange_guid = 'bbbbbbbb-0000-0000-0000-000000000002'; immutable_id = 'z'; user_principal_name = 'room@contoso.com' }
        $state = New-State -Other @{ ExchangeGuid = 'bbbbbbbb-0000-0000-0000-000000000002' }
        Get-Codes (Get-Pra2Readiness -Config $script:config -Record $record -State $state) 'Error' | Should -Contain 'KIND_NOT_SUPPORTED'
        Get-Codes (Get-Pra2Readiness -Config $script:config -Record $script:sharedRecord -State $state) 'Error' | Should -Not -Contain 'KIND_NOT_SUPPORTED'
    }
    It 'detects the Exchange mailbox plans by ID: Foundation and other Exchange-backed plans give no mailbox' {
        $skus = @([pscustomobject]@{ skuId = 's'; servicePlans = @(
                    [pscustomobject]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; servicePlanId = 'exo' }
                    [pscustomobject]@{ servicePlanName = 'EXCHANGE_S_FOUNDATION'; servicePlanId = 'foundation' }
                    [pscustomobject]@{ servicePlanName = 'EXCHANGE_ANALYTICS'; servicePlanId = 'analytics' }) })
        Get-Pra2MailboxPlanId -Skus $skus | Should -Be @('exo')
        Get-Pra2SkuExchangePlan -Sku ([pscustomobject]@{ servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_FOUNDATION'; servicePlanId = 'f' }) }) | Should -BeNullOrEmpty
        $foundation = New-State -User @{ assignedPlans = @([pscustomobject]@{ service = 'exchange'; servicePlanId = 'foundation'; capabilityStatus = 'Enabled' }) }
        Get-Codes (Get-Pra2Readiness -Config $script:config -Record $script:userRecord -State $foundation -Skus $skus) | Should -Not -Contain 'EXCHANGE_PLAN_PRESENT'
        $mailbox = New-State -User @{ assignedPlans = @([pscustomobject]@{ service = 'exchange'; servicePlanId = 'exo'; capabilityStatus = 'Enabled' }) }
        Get-Codes (Get-Pra2Readiness -Config $script:config -Record $script:userRecord -State $mailbox -Skus $skus) 'Error' | Should -Contain 'EXCHANGE_PLAN_PRESENT'
    }
    It 'refuses an object without usageLocation only when no default country is configured' {
        $state = New-State -User @{ usageLocation = '' }
        $config = New-TestConfig
        Get-Codes (Get-Pra2Readiness -Config $config -Record $script:userRecord -State $state) 'Error' | Should -Contain 'NO_USAGE_LOCATION'
        $config.Cloud.DefaultUsageLocation = 'FR'
        $findings = Get-Pra2Readiness -Config $config -Record $script:userRecord -State $state
        Get-Codes $findings 'Info' | Should -Contain 'NO_USAGE_LOCATION'
        Get-Codes $findings 'Error' | Should -BeNullOrEmpty
    }
    It 'checks the Kiosk plan of each user in Kiosk mode' {
        $config = New-TestConfig
        $config.Licensing.Users.Mode = 'Kiosk'
        $skus = @([pscustomobject]@{ skuId = 't1'; skuPartNumber = 'Microsoft_Teams_Enterprise_New'; servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_DESKLESS'; servicePlanId = 'k' }) })
        $with = New-State -User @{ assignedLicenses = @([pscustomobject]@{ skuId = 't1'; disabledPlans = @('k') }) }
        Get-Codes (Get-Pra2Readiness -Config $config -Record $script:userRecord -State $with -Skus $skus) 'Ok' | Should -Contain 'KIOSK_AVAILABLE'
        Get-Codes (Get-Pra2Readiness -Config $config -Record $script:userRecord -State (New-State) -Skus $skus) 'Error' | Should -Contain 'NO_KIOSK_PLAN'
    }
}

Describe 'Cloud rules: tenant and licences' {
    BeforeAll {
        $script:skus = @(
            [pscustomobject]@{ skuId = 'e5'; skuPartNumber = 'E5'; consumedUnits = 24; prepaidUnits = [pscustomobject]@{ enabled = 25 }
                servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; servicePlanId = 'p-exo' }, [pscustomobject]@{ servicePlanName = 'TEAMS1'; servicePlanId = 'p-teams' }) }
            [pscustomobject]@{ skuId = 'tm'; skuPartNumber = 'TEAMS'; consumedUnits = 25; prepaidUnits = [pscustomobject]@{ enabled = 25 }
                servicePlans = @([pscustomobject]@{ servicePlanName = 'TEAMS1'; servicePlanId = 'p-teams' }) }
        )
        $script:roles = @('User.ReadWrite.All','User-OnPremisesSyncBehavior.ReadWrite.All','LicenseAssignment.ReadWrite.All','Organization.Read.All','GroupMember.ReadWrite.All')
        function script:New-Tenant { param($Group, $CloudManaged = $true)
            [pscustomobject]@{ Skus = $script:skus; Organization = $null; LicenceGroup = $Group; LicenceGroupCloudManaged = $CloudManaged } }
        $script:group = [pscustomobject]@{ displayName = 'PRA-LIC-EXO'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @('p-teams') }) }
    }

    It 'reports the capacity of the licence group' {
        $findings = Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $script:group) -UserCount 1 -SharedCount 1
        @($findings | Where-Object Level -eq 'Error').Count | Should -Be 0
        @($findings | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Message | Should -Match '1 free unit'
    }
    It 'fails when the group has fewer free units than users' {
        $findings = Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $script:group) -UserCount 3
        @($findings | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Level | Should -Be 'Error'
    }
    It 'fails when the group assigns no Exchange plan' {
        $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @('p-exo') }) }
        $findings = Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $group) -UserCount 1
        @($findings | Where-Object { $_.Code -eq 'LICENCE_GROUP' -and $_.Level -eq 'Error' }).Count | Should -Be 1
    }
    It 'needs the group SOA permission when the licence group is synchronised from AD' {
        $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $true; assignedLicenses = $script:group.assignedLicenses }
        $findings = Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $group $false) -UserCount 1
        @($findings | Where-Object Code -eq 'LICENCE_GROUP_SOA').Count | Should -Be 1
        @($findings | Where-Object { $_.Code -eq 'APP_PERMISSION' -and $_.Level -eq 'Error' -and $_.Message -like '*Group-OnPremisesSyncBehavior*' }).Count | Should -Be 1
    }
    It 'lists the missing application permissions' {
        $findings = Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles @('User.ReadWrite.All') -Tenant (New-Tenant $script:group) -UserCount 1
        @($findings | Where-Object { $_.Code -eq 'APP_PERMISSION' })[0].Message | Should -Match 'User-OnPremisesSyncBehavior'
    }
    It 'fails when no unit is free for the temporary licence of shared mailboxes' {
        $config = New-TestConfig
        $config.Licensing.Shared.SkuPartNumber = 'TEAMS'
        $findings = Test-Pra2TenantReadiness -Config $config -Roles $script:roles -Tenant (New-Tenant $script:group) -UserCount 0 -SharedCount 2
        @($findings | Where-Object { $_.Code -eq 'SHARED_LICENCE' -and $_.Level -eq 'Error' }).Count | Should -Be 1
    }
    It 'warns about retention policies applied to every mailbox (they would block the user Recover)' {
        $tenant = New-Tenant $script:group
        @(Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant $tenant -UserCount 1 | Where-Object Code -eq 'ORG_HOLD').Count | Should -Be 0
        $tenant | Add-Member -NotePropertyName OrgHolds -NotePropertyValue @('mbxaaaa:2')
        $finding = @(Test-Pra2TenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant $tenant -UserCount 1 | Where-Object Code -eq 'ORG_HOLD')[0]
        $finding.Level | Should -Be 'Warn'
        $finding.Message | Should -Match 'mbxaaaa:2'
    }
}

Describe 'Licences restored as before Convert (mocked Graph)' {
    BeforeAll {
        if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { function global:Invoke-MgGraphRequest { param($Method, $Uri, $Body, $ContentType, $OutputType) } }
        function script:New-Licence { param([object[]]$States) [pscustomobject]@{ licenseAssignmentStates = $States } }
    }
    AfterAll { Remove-Item -Path 'function:global:Invoke-MgGraphRequest' -ErrorAction SilentlyContinue }
    BeforeEach { $script:sent = [System.Collections.Generic.List[object]]::new() }
    It 'enables a plan inside the existing direct assignment and keeps the other disabled plans' {
        Mock Get-Pra2UserLicence -ModuleName PRA2.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = $null; disabledPlans = @('kiosk', 'yammer') }) }
        Mock Set-Pra2Licence -ModuleName PRA2.Cloud { $script:sent.Add(@{ Sku = $SkuId; Operation = $Operation; Disabled = @(Get-Variable -Name DisabledPlans -ValueOnly -ErrorAction SilentlyContinue | ForEach-Object { $_ }) }) }
        Enable-Pra2LicencePlan -UserId 'u' -SkuId 'teams' -PlanIds @('kiosk') -DisabledPlans @('ignored') | Should -Be 'Enabled'
        $script:sent[0].Disabled | Should -Be @('yammer')
    }
    It 'adds a direct assignment next to a group one, and changes nothing when the plan is already enabled' {
        Mock Get-Pra2UserLicence -ModuleName PRA2.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = 'g'; disabledPlans = @('kiosk') }) }
        Mock Set-Pra2Licence -ModuleName PRA2.Cloud { $script:sent.Add(@{ Sku = $SkuId; Operation = $Operation; Disabled = @(Get-Variable -Name DisabledPlans -ValueOnly -ErrorAction SilentlyContinue | ForEach-Object { $_ }) }) }
        Enable-Pra2LicencePlan -UserId 'u' -SkuId 'teams' -PlanIds @('kiosk') -DisabledPlans @('x') | Should -Be 'Assigned'
        $script:sent[0].Disabled | Should -Be @('x')
        Mock Get-Pra2UserLicence -ModuleName PRA2.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = $null; disabledPlans = @('yammer') }) }
        Enable-Pra2LicencePlan -UserId 'u' -SkuId 'teams' -PlanIds @('kiosk') | Should -Be 'Present'
        $script:sent.Count | Should -Be 1
    }
    It 'restores the original disabled plans of a licence the user had, and removes one he did not have' {
        Mock Get-Pra2UserLicence -ModuleName PRA2.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = $null; disabledPlans = @('yammer') }) }
        Mock Set-Pra2Licence -ModuleName PRA2.Cloud { $script:sent.Add(@{ Sku = $SkuId; Operation = $Operation; Disabled = @(Get-Variable -Name DisabledPlans -ValueOnly -ErrorAction SilentlyContinue | ForEach-Object { $_ }) }) }
        Restore-Pra2Licence -UserId 'u' -SkuId 'teams' -Original @([pscustomobject]@{ skuId = 'teams'; disabledPlans = @('yammer', 'kiosk') }) | Should -Be 'Restored'
        $script:sent[0].Operation | Should -Be 'Add'
        $script:sent[0].Disabled | Should -Be @('kiosk', 'yammer')
        Restore-Pra2Licence -UserId 'u' -SkuId 'teams' -Original @() | Should -Be 'Removed'
        $script:sent[1].Operation | Should -Be 'Remove'
        Restore-Pra2Licence -UserId 'u' -SkuId 'other' -Original @() | Should -Be 'Absent'
        Restore-Pra2Licence -UserId 'u' -SkuId 'teams' -Original @([pscustomobject]@{ skuId = 'teams'; disabledPlans = @('yammer') }) | Should -Be 'Unchanged'
        $script:sent.Count | Should -Be 2
    }
}

Describe 'Entra Connect operator script' {
    It 'accepts a script that ends without exit (StrictMode) and refuses a non-zero exit code' {
        $folder = Join-Path $TestDrive 'ec'; $null = New-Item -ItemType Directory $folder
        $ok = Join-Path $folder 'ok.ps1'; Set-Content -LiteralPath $ok -Value "param(`$Operation) 'scheduler paused'" -Encoding UTF8
        $bad = Join-Path $folder 'bad.ps1'; Set-Content -LiteralPath $bad -Value "param(`$Operation) 'failed'; exit 3" -Encoding UTF8
        $context = @{ Config = @{ EntraConnect = @{ Mode = 'Script'; ScriptPath = $bad; Server = 'aadc' } } }
        { Invoke-Pra2EntraConnect -Context $context -Operation Pause } | Should -Throw '*exit code 3*'
        $context.Config.EntraConnect.ScriptPath = $ok
        & { Set-StrictMode -Version Latest; Invoke-Pra2EntraConnect -Context $context -Operation Pause } | Should -Be 'scheduler paused'
    }
}

Describe 'Cloud rules: shared mailbox trustees' {
    It 'counts the rights and flags trustees unknown in Exchange Online' {
        $record = [pscustomobject]@{ object_guid = 's1' }
        $permissions = @(
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'FullAccess'; trustee = 'C\u1'; trustee_kind = 'User'; trustee_upn = 'u1@c.com'; trustee_smtp = ''; trustee_guid = 'u1'; resolved = 1 }
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'SendAs'; trustee = 'C\grp'; trustee_kind = 'Group'; trustee_upn = ''; trustee_smtp = ''; trustee_guid = 'g1'; resolved = 1 }
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'SendOnBehalf'; trustee = 'C\old'; trustee_kind = 'Unknown'; trustee_upn = ''; trustee_smtp = ''; trustee_guid = ''; resolved = 0 }
            [pscustomobject]@{ mailbox_guid = 'other'; access_right = 'FullAccess'; trustee = 'C\u9'; trustee_kind = 'User'; trustee_upn = 'u9@c.com'; trustee_smtp = ''; trustee_guid = 'u9'; resolved = 1 }
        )
        $members = @([pscustomobject]@{ group_guid = 'g1'; member_kind = 'User'; member_upn = 'gone@c.com'; member_smtp = '' })
        $finding = @(Get-Pra2TrusteeFinding -Record $record -Permissions $permissions -Members $members -Lookup { param($id) $id -eq 'u1@c.com' })[0]
        $finding.Level | Should -Be 'Warn'
        $finding.Message | Should -Match 'FullAccess 1 · SendAs 1 · SendOnBehalf 1'
        $finding.Message | Should -Match 'gone@c\.com \(not found in Exchange Online\)'
        $finding.Message | Should -Match 'C\\old \(not resolved at Collect\)'
        $finding.Message | Should -Not -Match 'u9'
    }
}

Describe 'Results and report' {
    It 'computes the exit code: errors = 1, pending = 2, done = 0' {
        $context = @{ Rows = (New-Object 'Collections.Generic.List[object]'); Issues = (New-Object 'Collections.Generic.List[object]'); ExitCode = 0; Mode = 'Preview'; CloudConnectFailed = $false }
        [void]$context.Rows.Add([ordered]@{ Identity = 'a'; FinalStatus = 'Success' })
        (Get-PraOutcome $context).ExitCode | Should -Be 0
        [void]$context.Rows.Add([ordered]@{ Identity = 'b'; FinalStatus = 'Pending' })
        (Get-PraOutcome $context).ExitCode | Should -Be 2
        [void]$context.Rows.Add([ordered]@{ Identity = 'c'; FinalStatus = 'Error' })
        (Get-PraOutcome $context).ExitCode | Should -Be 1
    }
}

Describe 'Entry script' {
    It 'uses no $PSScriptRoot in the parameter defaults (empty in Windows PowerShell 5.1 with -File, e.g. a scheduled task)' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
        $defaults = @($ast.ParamBlock.Parameters | Where-Object { $_.DefaultValue } | ForEach-Object { $_.DefaultValue.Extent.Text })
        @($defaults | Where-Object { $_ -match 'PSScriptRoot|PSCommandPath|MyInvocation' }) | Should -BeNullOrEmpty
    }
    It 'creates no generic collection with New-Object (@() of the PSObject wrapper throws "Argument types do not match")' {
        $sources = Get-ChildItem -LiteralPath $script:Root -Recurse -File | Where-Object { $_.Extension -in @('.ps1', '.psm1') -and $_.FullName -notmatch '\\(lib|logs|reports|data|tests)\\' }
        @($sources | Select-String -Pattern 'New-Object\s+[''"]?System\.Collections\.Generic') | Should -BeNullOrEmpty
    }
    It 'reads no first element with @(...)[0] (IndexOutOfRange on an empty result under StrictMode Latest)' {
        $sources = Get-ChildItem -LiteralPath $script:Root -Recurse -File | Where-Object { $_.Extension -in @('.ps1', '.psm1') -and $_.FullName -notmatch '\\(lib|logs|reports|data|tests)\\' }
        @($sources | Select-String -Pattern '@\([^\r\n]*\)\[0\]') | ForEach-Object { '{0}:{1}' -f $_.Filename, $_.LineNumber } | Should -BeNullOrEmpty
    }
    It 'starts with powershell.exe -File and reports the missing Exchange cmdlets' -Skip:($PSVersionTable.PSEdition -ne 'Desktop') {
        $folder = Join-Path $TestDrive 'file-mode'
        $null = New-Item -ItemType Directory $folder
        $config = New-TestConfig
        $config.Logging.Folder = (Join-Path $folder 'logs'); $config.Report.Folder = (Join-Path $folder 'reports'); $config.Store.Path = (Join-Path $folder 'db.db')
        $configPath = Join-Path $folder 'c.psd1'
        Set-Content -LiteralPath $configPath -Value (ConvertTo-Psd1 $config) -Encoding UTF8
        $output = & powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1') -Action Collect -ConfigPath $configPath 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 1
        $output | Should -Not -Match 'Join-Path'
        $output | Should -Match 'Exchange cmdlets not available|RESULT: FAIL'
    }
}

Describe 'Journal of the cloud actions' {
    BeforeEach {
        $script:jpath = Join-Path $TestDrive ("journal-{0}.db" -f [guid]::NewGuid().ToString('N'))
        $script:j = Open-Pra2Store -Path $script:jpath -Root $script:Root -Kind Journal
    }
    AfterEach { Close-Pra2Store $script:j }

    It 'creates a batch, records items and events, and finds it by prefix' {
        $id = New-Pra2Batch $script:j Convert @{ SnapshotId = 3; Environment = 'TEST'; ToolVersion = '0.2.0' }
        $id | Should -Match '^[0-9a-f]{8}$'
        Set-Pra2BatchItem $script:j $id 'g1' @{ kind = 'User'; identity = 'u1@c.com'; entra_id = 'e1'; status = 'Planned'; original_json = '{"isCloudManaged":false}' }
        Set-Pra2BatchItem $script:j $id 'g1' @{ step = 'SoaCloud'; status = 'Running' }
        Add-Pra2BatchEvent $script:j $id -ObjectGuid 'g1' -Step 'SoaCloud' -Detail 'isCloudManaged = true'
        $item = @(Get-Pra2BatchItem $script:j $id)[0]
        $item.step | Should -Be 'SoaCloud'
        $item.status | Should -Be 'Running'
        $item.original_json | Should -Be '{"isCloudManaged":false}'
        (Get-Pra2Batch $script:j -Id $id.Substring(0, 5)).id | Should -Be $id
        (Get-Pra2Batch $script:j -Action Convert).snapshot_id | Should -Be 3
        [int](Invoke-Pra2Sql $script:j 'SELECT count(*) FROM batch_event' -As Scalar) | Should -Be 1
    }
    It 'refuses an invalid batch ID, a journal column outside the list, and a new item without kind and status' {
        { Get-Pra2Batch $script:j -Id "x'; DROP" } | Should -Throw '*Invalid batch ID*'
        $id = New-Pra2Batch $script:j Recover @{ ConvertBatch = 'abcd1234' }
        { Set-Pra2BatchItem $script:j $id 'g' @{ kind = 'User'; status = 'Planned'; 'step = 1 --' = 'x' } } | Should -Throw '*Invalid journal column*'
        { Set-Pra2BatchItem $script:j $id 'g' @{ step = 'x' } } | Should -Throw '*needs kind and status*'
    }
    It 'records a tenant change and its restoration' {
        $id = New-Pra2Batch $script:j Convert
        Set-Pra2TenantChange $script:j $id 'GroupSoa' 'grp1' -OriginalJson '{"isCloudManaged":false}'
        @(Get-Pra2TenantChange $script:j $id)[0].restored | Should -Be 0
        Set-Pra2TenantChange $script:j $id 'GroupSoa' 'grp1' -Restored
        @(Get-Pra2TenantChange $script:j $id)[0].restored | Should -Be 1
    }
    It 'never opens a snapshot database as a journal, nor a journal as a snapshot database' {
        $snapshotPath = Join-Path $TestDrive ("snap-{0}.db" -f [guid]::NewGuid().ToString('N'))
        Close-Pra2Store (Open-Pra2Store -Path $snapshotPath -Root $script:Root)
        { Open-Pra2Store -Path $snapshotPath -Root $script:Root -Kind Journal } | Should -Throw '*snapshot database, not a journal*'
        { Open-Pra2Store -Path $script:jpath -Root $script:Root -ReadOnly } | Should -Throw '*journal database, not a snapshot*'
    }
}

Describe 'Convert helpers (pure)' {
    It 'expands the shared mailbox permissions to one grant per right and user' {
        $record = [pscustomobject]@{ object_guid = 's1' }
        $permissions = @(
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'FullAccess'; trustee = 'C\u1'; trustee_kind = 'User'; trustee_upn = 'u1@c.com'; trustee_smtp = ''; trustee_guid = 'u1'; trustee_name = 'u1'; resolved = 1; auto_mapping = 0 }
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'FullAccess'; trustee = 'C\grp'; trustee_kind = 'Group'; trustee_upn = ''; trustee_smtp = ''; trustee_guid = 'g1'; trustee_name = 'grp'; resolved = 1; auto_mapping = $null }
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'SendAs'; trustee = 'C\old'; trustee_kind = 'Unknown'; trustee_upn = ''; trustee_smtp = ''; trustee_guid = ''; trustee_name = ''; resolved = 0; auto_mapping = $null }
            [pscustomobject]@{ mailbox_guid = 's1'; access_right = 'SendOnBehalf'; trustee = 'CN=u2'; trustee_kind = 'User'; trustee_upn = ''; trustee_smtp = 'u2@c.com'; trustee_guid = 'u2'; trustee_name = 'u2'; resolved = 1; auto_mapping = $null }
        )
        $members = @(
            [pscustomobject]@{ group_guid = 'g1'; member_kind = 'User'; member_upn = 'u1@c.com'; member_smtp = '' }
            [pscustomobject]@{ group_guid = 'g1'; member_kind = 'User'; member_upn = 'u3@c.com'; member_smtp = '' }
        )
        $grants = @(Get-Pra2SharedGrant -Record $record -Permissions $permissions -Members $members)
        ($grants | ForEach-Object { '{0} {1} {2}' -f $_.Right, $_.Trustee, $_.AutoMapping }) | Should -Be @('FullAccess u1@c.com False', 'FullAccess u3@c.com True', 'SendOnBehalf u2@c.com ')
    }
    It 'computes a Kiosk assignment that only adds the Kiosk plan to the existing licence' {
        $skus = @([pscustomobject]@{ skuId = 't1'; skuPartNumber = 'TEAMS'; servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_DESKLESS'; servicePlanId = 'k' }, [pscustomobject]@{ servicePlanName = 'X'; servicePlanId = 'x' }) })
        $licence = [pscustomobject]@{ licenseAssignmentStates = @([pscustomobject]@{ skuId = 't1'; assignedByGroup = 'grp'; disabledPlans = @('k', 'x') }) }
        $kiosk = Get-Pra2KioskAssignment -Licence $licence -Skus $skus
        $kiosk.SkuId | Should -Be 't1'
        $kiosk.DisabledPlans | Should -Be @('x')
        $kiosk.DirectWithKiosk | Should -BeFalse
        $licence.licenseAssignmentStates += [pscustomobject]@{ skuId = 't1'; assignedByGroup = $null; disabledPlans = @('x') }
        (Get-Pra2KioskAssignment -Licence $licence -Skus $skus).DirectWithKiosk | Should -BeTrue
        Get-Pra2KioskAssignment -Licence ([pscustomobject]@{ licenseAssignmentStates = @() }) -Skus $skus | Should -BeNullOrEmpty
    }
    It 'leaves only the Exchange plan enabled for a licence given for the mailbox' {
        $sku = [pscustomobject]@{ servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; servicePlanId = 'e'; appliesTo = 'User' }, [pscustomobject]@{ servicePlanName = 'TEAMS1'; servicePlanId = 't'; appliesTo = 'User' }, [pscustomobject]@{ servicePlanName = 'COMPANY'; servicePlanId = 'c'; appliesTo = 'Company' }) }
        Get-Pra2DisabledPlanForExchangeOnly -Sku $sku | Should -Be @('t')
    }
    It 'waits for a condition and reports the elapsed time' {
        $script:n = 0
        $result = Wait-Pra2Condition -Test { $script:n++; $script:n -ge 2 } -TimeoutSeconds 10 -IntervalSeconds 1
        $result.Ok | Should -BeTrue
        (Wait-Pra2Condition -Test { $false } -TimeoutSeconds 1 -IntervalSeconds 1).Ok | Should -BeFalse
    }
}

Describe 'Cloud calls (mocked Graph)' {
    BeforeAll {
        if (-not (Get-Command Invoke-MgGraphRequest -ErrorAction SilentlyContinue)) { function global:Invoke-MgGraphRequest { param($Method, $Uri, $Body, $ContentType, $OutputType) } }
        function script:New-GraphError {
            param([string]$Code, [string]$Message)
            $record = [System.Management.Automation.ErrorRecord]::new([Exception]::new('Response status code does not indicate success: BadRequest (Bad Request).'), 'Graph', 'InvalidOperation', $null)
            $record.ErrorDetails = [System.Management.Automation.ErrorDetails]::new((@{ error = @{ code = $Code; message = $Message } } | ConvertTo-Json -Compress))
            return $record
        }
    }
    AfterAll { Remove-Item -Path 'function:global:Invoke-MgGraphRequest' -ErrorAction SilentlyContinue }
    It 'puts the Graph error code and message in the exception (not only "BadRequest")' {
        Mock Invoke-MgGraphRequest -ModuleName PRA2.Cloud { throw (New-GraphError 'Request_BadRequest' 'Invalid usage location.') }
        { Invoke-Pra2Graph -Method POST -Uri 'v1.0/users/u1/assignLicense?x=1' -Body @{} } | Should -Throw '*POST v1.0/users/u1/assignLicense failed*Request_BadRequest: Invalid usage location.*'
    }
    It 'retries a licence assignment refused right after the usageLocation change' {
        $script:calls = 0
        Mock Start-Sleep -ModuleName PRA2.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA2.Cloud { $script:calls++; if ($script:calls -lt 3) { throw (New-GraphError 'Request_BadRequest' 'License assignment cannot be done for user with invalid usage location.') } }
        Set-Pra2Licence -UserId 'u1' -SkuId 's1' -Operation Add -DisabledPlans @('p1')
        $script:calls | Should -Be 3
    }
    It 'does not retry a licence removal' {
        $script:calls = 0
        Mock Start-Sleep -ModuleName PRA2.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA2.Cloud { $script:calls++; throw (New-GraphError 'Request_BadRequest' 'x') }
        { Set-Pra2Licence -UserId 'u1' -SkuId 's1' -Operation Remove } | Should -Throw
        $script:calls | Should -Be 1
    }
}

Describe 'Configuration' {
    It 'refuses <Case>' -ForEach @(
        @{ Case = 'Remoting without server'; Text = "Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Remoting'; Server = '' }"; Message = '*requires EntraConnect.Server*' }
        @{ Case = 'Script without path'; Text = "Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Script'; Server = 'x' }"; Message = '*requires EntraConnect.ScriptPath*' }
        @{ Case = 'a bad country'; Text = "Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'x' }; Cloud = @{ DefaultUsageLocation = 'France' }"; Message = '*two-letter country*' }
        @{ Case = 'a too short polling interval'; Text = "Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'x' }; Polling = @{ IntervalSeconds = 1 }"; Message = '*IntervalSeconds*' }
    ) {
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $folder
        $path = Join-Path $folder 'c.psd1'
        Set-Content -LiteralPath $path -Value ("@{ Environment = 'X'; $Text }") -Encoding UTF8
        { Import-PraConfiguration -Path $path -Root $folder } | Should -Throw $Message
    }
}