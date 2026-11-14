<#
.SYNOPSIS
    PRA Cloud Mailbox - Collect against the objects of a remote Exchange session (Pester 5+). No Exchange, no Active Directory, no network.
.DESCRIPTION
    Collect reads Exchange through Import-PSSession, so every object it gets is a deserialized copy: properties kept, methods lost, sizes as text,
    enumerations as numbers that print as their name, directory references as ADObjectId copies, collections as ArrayLists. The fixtures
    (tests\DeserializedFixture.ps1) build objects of private types that mimic the Exchange ones and send them through PSSerializer, the conversion
    of a remote session. Every code path of Collect is fed twice, with the live objects and with their deserialized copies, and must give the same result.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
#>
#Requires -Version 7.5

BeforeDiscovery {
    . (Join-Path $PSScriptRoot 'DeserializedFixture.ps1')
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo =[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    . (Join-Path $PSScriptRoot 'DeserializedFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    Import-Module (Join-Path $script:Root 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Store.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Collect.psm1') -Force

    # The Exchange commands that Collect calls (the remote session imports them; here they are empty functions that the tests mock).
    function global:Get-Mailbox { param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController) }
    function global:Get-MailboxStatistics { param($Identity, $DomainController) }
    function global:Get-CalendarProcessing { param($Identity, $DomainController) }
    function global:Get-MailboxCalendarConfiguration { param($Identity, $DomainController) }
    function global:Get-MailboxPermission { param($Identity, $DomainController) }
    function global:Get-ADPermission { param($Identity, $DomainController) }
    function global:Get-Recipient { param($Identity, $DomainController) }
    function global:Get-User { param($Identity, $DomainController) }
    function global:Get-Group { param($Identity, $DomainController) }
    function global:Get-MailContact { param($ResultSize, $OrganizationalUnit, $DomainController) }
    function global:Get-DistributionGroup { param($ResultSize, $OrganizationalUnit, $DomainController, $Filter) }
    function global:Get-DistributionGroupMember { param($Identity, $ResultSize, $DomainController) }
    function global:Get-DynamicDistributionGroup { param($ResultSize, $OrganizationalUnit, $DomainController) }
    function global:Get-ExchangeServer { param() }
    function global:Set-ADServerSettings { param($ViewEntireForest) }

    function script:New-TestConfig {
        $config = @{
            Environment = 'TEST'
            Exchange = @{ Server = 'ex1.contoso.com'; DomainController = '' }
            Scope = @{ Mode = 'OU'; SearchBase = 'OU=PRA,DC=contoso,DC=com'; GroupDN = ''; CsvPath = ''; IncludeUsers = $true; IncludeShared = $true; IncludeRoom = $false; IncludeEquipment = $false; ExcludeSamAccountNames = @(); ConvertRooms = $false }
            Collect = @{ SharedPermissions = $true; ExpandGroupTrustees = $true; ExcludeTrustees = @('Administrator'); MailboxStatistics = $false; Contacts = $false; ContactsSearchBase = ''; DistributionGroups = $false; DistributionGroupsSearchBase = ''; DynamicDistributionGroups = $false; TargetGroups = @() }
        }
        return $config
    }
    function script:New-TestContext {
        param([hashtable]$Config = (New-TestConfig))
        return @{ Config = $Config; Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]')
            Excluded = (New-Object 'Collections.Generic.List[string]'); LogFile = ''; Warnings = 0; Server = '' }
    }
    # What the engine receives: the object as it is (live) or its deserialized copy (what a remote session returns).
    function global:Use-Fixture([string]$Kind, $Value) { if ($Kind -eq 'Deserialized') { return (ConvertTo-DeserializedFixture $Value) } else { return $Value } }
    function script:ConvertTo-Comparable($Value) { return (ConvertTo-Json -InputObject $Value -Depth 10 -Compress) }
}

AfterAll {
    foreach ($name in @('Get-Mailbox', 'Get-MailboxStatistics', 'Get-CalendarProcessing', 'Get-MailboxCalendarConfiguration', 'Get-MailboxPermission', 'Get-ADPermission', 'Get-Recipient', 'Get-User', 'Get-Group',
            'Get-MailContact', 'Get-DistributionGroup', 'Get-DistributionGroupMember', 'Get-DynamicDistributionGroup', 'Get-ExchangeServer', 'Set-ADServerSettings')) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'The fixtures are what a remote session returns' {
    It 'loses the methods, keeps the properties and prints the text of the original' {
        $deserialized = ConvertTo-DeserializedFixture (New-FixtureMailbox -Sam 'alice' -Type SharedMailbox)
        $deserialized.PSObject.TypeNames[0] | Should -BeExactly 'Deserialized.Microsoft.Exchange.Data.Directory.Management.Mailbox'
        $size = (ConvertTo-DeserializedFixture ([pscustomobject]@{ Size = (New-FixtureSize 1288490188) })).Size
        $size.PSObject.Methods['ToBytes'] | Should -BeNullOrEmpty
        [string]$size | Should -BeExactly '1.2 GB (1,288,490,188 bytes)'
        $deserialized.OrganizationalUnit.PSObject.TypeNames[0] | Should -BeLike 'Deserialized.*'
        [string]$deserialized.OrganizationalUnit | Should -BeExactly 'contoso.com/PRA'
        $deserialized.OrganizationalUnit.DistinguishedName | Should -BeExactly 'OU=PRA,DC=contoso,DC=com'
        [string]$deserialized.RecipientTypeDetails | Should -BeExactly 'SharedMailbox'
        $deserialized.RecipientTypeDetails.GetType().FullName | Should -BeExactly 'System.Int32'
        $deserialized.EmailAddresses.GetType().FullName | Should -BeExactly 'System.Collections.ArrayList'
        [string]$deserialized.EmailAddresses[0] | Should -BeExactly 'SMTP:alice@contoso.com'
        $deserialized.EmailAddresses[0].PSObject.TypeNames[0] | Should -BeLike 'Deserialized.*'
    }
}

Describe 'Collect: sizes' {
    It 'reads <Name> as <Bytes> bytes' -ForEach @(
        @{ Name = 'a live ByteQuantifiedSize'; Value = (New-FixtureSize 1288490188); Bytes = 1288490188 }
        @{ Name = 'its deserialized copy (text only, no ToBytes)'; Value = (ConvertTo-DeserializedFixture ([pscustomobject]@{ S = (New-FixtureSize 1288490188) })).S; Bytes = 1288490188 }
        @{ Name = 'the text with commas'; Value = '1.2 GB (1,288,490,188 bytes)'; Bytes = 1288490188 }
        @{ Name = 'the text with no-break spaces'; Value = ('1,2 GB (1{0}288{0}490{0}188 bytes)' -f [char]0x00A0); Bytes = 1288490188 }
        @{ Name = 'the text with narrow no-break spaces'; Value = ('1,2 GB (1{0}288{0}490{0}188 bytes)' -f [char]0x202F); Bytes = 1288490188 }
        @{ Name = 'the text with dots'; Value = '12 KB (12.288 bytes)'; Bytes = 12288 }
        @{ Name = 'an empty mailbox'; Value = '0 B (0 bytes)'; Bytes = 0 }
        @{ Name = 'a number'; Value = [long]2048; Bytes = 2048 }
        @{ Name = 'an integer'; Value = 512; Bytes = 512 }
        @{ Name = 'a deserialized size of zero'; Value = (ConvertTo-DeserializedFixture ([pscustomobject]@{ S = (New-FixtureSize 0) })).S; Bytes = 0 }
    ) {
        ConvertTo-PcmSizeBytes $Value | Should -Be $Bytes
    }
    It 'reads <Name> as unknown' -ForEach @(
        @{ Name = 'Unlimited'; Value = 'Unlimited' }
        @{ Name = 'a deserialized unlimited size'; Value = (ConvertTo-DeserializedFixture ([pscustomobject]@{ S = [PcmFixture.UnlimitedSize]::new() })).S }
        @{ Name = 'nothing'; Value = $null }
        @{ Name = 'a text without a byte count'; Value = 'about one gigabyte' }
    ) {
        ConvertTo-PcmSizeBytes $Value | Should -BeNullOrEmpty
    }
}

Describe 'Collect: references and plain values (internal helpers)' {
    It 'reads the distinguished name of <Name>' -ForEach @(
        @{ Name = 'a string'; Value = '  CN=u1,OU=PRA,DC=contoso,DC=com '; Expected = 'CN=u1,OU=PRA,DC=contoso,DC=com' }
        @{ Name = 'an ADObjectId'; Value = (New-FixtureAdId 'CN=u1,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/u1'); Expected = 'CN=u1,OU=PRA,DC=contoso,DC=com' }
        @{ Name = 'a deserialized ADObjectId'; Value = (ConvertTo-DeserializedFixture ([pscustomobject]@{ R = (New-FixtureAdId 'CN=u1,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/u1') })).R; Expected = 'CN=u1,OU=PRA,DC=contoso,DC=com' }
        @{ Name = 'a deserialized reference that has no distinguished name (its text)'; Value = (ConvertTo-DeserializedFixture ([pscustomobject]@{ R = ([PcmFixture.NameOnlyReference]@{ Name = 'u9'; Text = 'CN=u9,OU=PRA,DC=contoso,DC=com' }) })).R; Expected = 'CN=u9,OU=PRA,DC=contoso,DC=com' }
        @{ Name = 'nothing'; Value = $null; Expected = '' }
    ) {
        InModuleScope PRA.Collect -Parameters @{ Value = $Value; Expected = $Expected } { ConvertTo-PcmDistinguishedName $Value | Should -BeExactly $Expected }
    }
    It 'turns <Name> into plain data' -ForEach @(
        @{ Name = 'a string'; Make = { 'text' }; Expected = '"text"' }
        @{ Name = 'a boolean'; Make = { $true }; Expected = 'true' }
        @{ Name = 'a number'; Make = { 12 }; Expected = '12' }
        @{ Name = 'a live enumeration'; Make = { [PcmFixture.AutomateProcessing]::AutoAccept }; Expected = '"AutoAccept"' }
        @{ Name = 'a deserialized enumeration (a number that prints as its name)'; Make = { (ConvertTo-DeserializedFixture ([pscustomobject]@{ V = [PcmFixture.AutomateProcessing]::AutoAccept })).V }; Expected = '"AutoAccept"' }
        @{ Name = 'a deserialized flags enumeration'; Make = { (ConvertTo-DeserializedFixture ([pscustomobject]@{ V = [PcmFixture.WorkDays]'Monday, Tuesday' })).V }; Expected = '"Monday, Tuesday"' }
        @{ Name = 'a deserialized size'; Make = { (ConvertTo-DeserializedFixture ([pscustomobject]@{ V = (New-FixtureSize 1024) })).V }; Expected = '"1.0 KB (1,024 bytes)"' }
        @{ Name = 'a duration'; Make = { [timespan]'08:30:00' }; Expected = '"08:30:00"' }
        @{ Name = 'a date'; Make = { [datetime]::SpecifyKind([datetime]'2026-01-02T03:04:05', 'Utc') }; Expected = '"2026-01-02T03:04:05.0000000Z"' }
        @{ Name = 'a GUID'; Make = { [guid]'11111111-aaaa-4aaa-8aaa-000000000001' }; Expected = '"11111111-aaaa-4aaa-8aaa-000000000001"' }
        @{ Name = 'a deserialized multi-valued property'; Make = { (ConvertTo-DeserializedFixture ([pscustomobject]@{ V = (New-FixtureMulti @((New-FixtureProxy 'SMTP' 'a@contoso.com'), (New-FixtureProxy 'smtp' 'b@contoso.com'))) })).V }; Expected = '["SMTP:a@contoso.com","smtp:b@contoso.com"]' }
        @{ Name = 'nothing'; Make = { $null }; Expected = 'null' }
    ) {
        $value = & $Make
        $plain = & (Get-Module PRA.Collect) { ConvertTo-PcmPlainValue $args[0] } $value
        ConvertTo-Json -InputObject $plain -Compress -Depth 4 | Should -BeExactly $Expected
    }
}

Describe 'Collect: the mailbox record' {
    It 'is the same for the live mailbox and its deserialized copy (<Type>)' -ForEach @(@{ Type = 'UserMailbox' }, @{ Type = 'SharedMailbox' }) {
        $mailbox = New-FixtureMailbox -Sam 'Alice' -Type $Type -OnBehalf @((New-FixtureAdId 'CN=Bob,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/Bob'))
        $mailbox.ArchiveGuid = [guid]'33333333-aaaa-4aaa-8aaa-000000000003'; $mailbox.ArchiveStatus = 'Active'
        $statistics = [pscustomobject]@{ ItemCount = 1200; TotalItemSize = (New-FixtureSize 1288490188) }
        $live = ConvertTo-PcmMailboxRecord -Mailbox $mailbox -Statistics $statistics
        $copy = ConvertTo-PcmMailboxRecord -Mailbox (ConvertTo-DeserializedFixture $mailbox) -Statistics (ConvertTo-DeserializedFixture $statistics)
        ConvertTo-Comparable $copy | Should -BeExactly (ConvertTo-Comparable $live)
        $copy.kind | Should -BeExactly $(if ($Type -eq 'UserMailbox') { 'User' } else { 'Shared' })
        $copy.recipient_type_details | Should -BeExactly $Type
        $copy.organizational_unit | Should -BeExactly 'contoso.com/PRA'
        $copy.database | Should -BeExactly 'DB01'
        $copy.primary_smtp_address | Should -BeExactly 'Alice@contoso.com'
        $copy.object_guid | Should -Be $mailbox.Guid.ToString()
        $copy.immutable_id | Should -Be ([Convert]::ToBase64String($mailbox.Guid.ToByteArray()))
        $copy.archive_guid | Should -Be '33333333-aaaa-4aaa-8aaa-000000000003'
        $copy.archive_status | Should -BeExactly 'Active'
        $copy.hidden_from_address_lists | Should -Be ($Type -ne 'UserMailbox')
        @($copy.email_addresses_json | ConvertFrom-Json) | Should -Be @('SMTP:Alice@contoso.com', 'smtp:Alice@contoso.mail.onmicrosoft.com', 'X500:/o=Old/cn=Alice')
        ($copy.custom_attributes_json | ConvertFrom-Json).CustomAttribute2 | Should -BeExactly 'Paris'
        @($copy.grant_send_on_behalf_json | ConvertFrom-Json) | Should -Be @('contoso.com/PRA/Bob')
        $copy.item_count | Should -Be 1200
        $copy.total_item_size_bytes | Should -Be 1288490188
    }
    It 'keeps no archive and no statistics when Exchange gives none' {
        $mailbox = ConvertTo-DeserializedFixture (New-FixtureMailbox -Sam 'carol')
        $record = ConvertTo-PcmMailboxRecord -Mailbox $mailbox
        $record.archive_guid | Should -BeNullOrEmpty
        $record.item_count | Should -BeNullOrEmpty
        $record.total_item_size_bytes | Should -BeNullOrEmpty
        $record.resource_settings_json | Should -BeNullOrEmpty
    }
    It 'stops on a mailbox type that is not collected' {
        $mailbox = New-FixtureMailbox -Sam 'dave'
        $mailbox.RecipientTypeDetails = [PcmFixture.RecipientTypeDetails]::RoomList
        { ConvertTo-PcmMailboxRecord -Mailbox (ConvertTo-DeserializedFixture $mailbox) } | Should -Throw '*Unsupported mailbox type RoomList*'
    }
}

Describe 'Collect: the mailboxes in scope' {
    BeforeEach {
        $script:people = @((New-FixtureMailbox -Sam 'user02'), (New-FixtureMailbox -Sam 'User01'), (New-FixtureMailbox -Sam 'compta' -Type SharedMailbox), (New-FixtureMailbox -Sam 'HealthMailbox0123' -Type UserMailbox))
        $script:peopleByDn = @{}; foreach ($person in $script:people) { $script:peopleByDn[$person.DistinguishedName] = $person }
    }
    It 'lists the mailboxes of the whole organisation sorted by sAMAccountName, without the system mailboxes and the excluded accounts (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        Mock Get-Mailbox -ModuleName PRA.Collect { Use-Fixture $script:Kind @($script:people + $script:people[0]) }
        $config = New-TestConfig; $config.Scope.Mode = 'Auto'; $config.Scope.SearchBase = ''; $config.Scope.ExcludeSamAccountNames = @('COMPTA')
        $context = New-TestContext $config
        $result = @(Get-PcmScopeMailbox -Context $context)
        @($result | ForEach-Object { [string]$_.SamAccountName }) | Should -Be @('User01', 'user02')
        $context.Excluded | Should -Contain 'HealthMailbox0123: system mailbox'
        $context.Excluded | Should -Contain 'compta: Scope.ExcludeSamAccountNames'
        Should -Invoke Get-Mailbox -ModuleName PRA.Collect -Times 1 -ParameterFilter { [string]$ResultSize -eq 'Unlimited' -and @($RecipientTypeDetails) -contains 'UserMailbox' -and -not $OrganizationalUnit }
    }
    It 'lists the mailboxes of an organizational unit (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        Mock Get-Mailbox -ModuleName PRA.Collect { Use-Fixture $script:Kind $script:people[0..1] }
        $config = New-TestConfig; $config.Scope.Mode = 'ou'
        @(Get-PcmScopeMailbox -Context (New-TestContext $config)).Count | Should -Be 2
        Should -Invoke Get-Mailbox -ModuleName PRA.Collect -Times 1 -ParameterFilter { $OrganizationalUnit -eq 'OU=PRA,DC=contoso,DC=com' }
    }
    It 'takes one mailbox by identity, and refuses a type that is out of scope (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        Mock Get-Mailbox -ModuleName PRA.Collect { Use-Fixture $script:Kind $script:people[2] }
        $context = New-TestContext
        @(Get-PcmScopeMailbox -Context $context -Identity 'compta@contoso.com').Count | Should -Be 1
        { Get-PcmScopeMailbox -Context $context -Identity 'compta@contoso.com' -Filter UsersOnly } | Should -Throw '*compta@contoso.com is a SharedMailbox: not a mailbox type in scope (UserMailbox)*'
    }
    It 'takes the members of a group, nested groups walked, and says why a member is not collected (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        $groups = @{
            'CN=G1,OU=Groups,DC=contoso,DC=com' = @{ Members = @($script:people[1].DistinguishedName, 'CN=G2,OU=Groups,DC=contoso,DC=com', 'CN=PC1,OU=PRA,DC=contoso,DC=com', $script:people[2].DistinguishedName) }
            'CN=G2,OU=Groups,DC=contoso,DC=com' = @{ Members = @($script:people[0].DistinguishedName, $script:people[1].DistinguishedName, 'CN=G1,OU=Groups,DC=contoso,DC=com') }
        }
        $script:groupFixture = $groups
        Mock Get-Group -ModuleName PRA.Collect {
            $entry = $script:groupFixture[[string]$Identity]
            if (-not $entry) { throw "Group $Identity not found" }
            Use-Fixture $script:Kind ([pscustomobject]@{ Name = ($Identity -split ',')[0]; DistinguishedName = [string]$Identity; Members = (New-FixtureMulti @($entry.Members | ForEach-Object { New-FixtureAdId $_ ('contoso.com/x/' + $_) })) })
        }
        Mock Get-Mailbox -ModuleName PRA.Collect { if ($script:peopleByDn.ContainsKey([string]$Identity)) { Use-Fixture $script:Kind $script:peopleByDn[[string]$Identity] } }
        Mock Get-PcmDirectoryObject -ModuleName PRA.Collect { [pscustomobject]@{ Class = 'Computer'; Name = 'PC1'; Sam = ''; Upn = ''; Smtp = ''; DisplayName = ''; Dn = $DistinguishedName; Guid = '' } }
        $config = New-TestConfig; $config.Scope.Mode = 'group'; $config.Scope.GroupDN = 'CN=G1,OU=Groups,DC=contoso,DC=com'; $config.Scope.IncludeShared = $false
        $context = New-TestContext $config
        $result = @(Get-PcmScopeMailbox -Context $context)
        @($result | ForEach-Object { [string]$_.SamAccountName }) | Should -Be @('User01', 'user02')
        $context.Excluded | Should -Contain 'compta: SharedMailbox not in scope'
        $context.Excluded | Should -Contain 'PC1: a computer, not a mailbox'
    }
    It 'takes the objects of a CSV file, any letter case in the header and the identities (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        $path = Join-Path $TestDrive ('list-{0}.csv' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $path -Value "identity`r`nUSER01@CONTOSO.COM`r`n compta@contoso.com `r`n`r`n" -Encoding UTF8
        Mock Get-Mailbox -ModuleName PRA.Collect { $sam = ([string]$Identity -split '@')[0]; Use-Fixture $script:Kind (@($script:people | Where-Object { $_.SamAccountName -eq $sam }) | Select-Object -First 1) }
        $config = New-TestConfig; $config.Scope.Mode = 'CSV'; $config.Scope.CsvPath = $path; $config.Scope.IncludeShared = $false
        $context = New-TestContext $config
        @(Get-PcmScopeMailbox -Context $context | ForEach-Object { [string]$_.SamAccountName }) | Should -Be @('User01')
        $context.Excluded | Should -Contain 'compta@contoso.com: SharedMailbox not in scope'
    }
}

Describe 'Collect: group members' {
    BeforeEach {
        $script:tree = @{
            'CN=G1,OU=Groups,DC=contoso,DC=com' = @((New-FixtureAdId 'CN=u1,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/u1'), (New-FixtureAdId 'CN=G2,OU=Groups,DC=contoso,DC=com' 'contoso.com/Groups/G2'),
                ([PcmFixture.NameOnlyReference]@{ Name = 'u9'; Text = 'CN=u9,OU=PRA,DC=contoso,DC=com' }), 'CN=u2,OU=PRA,DC=contoso,DC=com')
            'CN=G2,OU=Groups,DC=contoso,DC=com' = @((New-FixtureAdId 'CN=G1,OU=Groups,DC=contoso,DC=com' 'contoso.com/Groups/G1'), (New-FixtureAdId 'CN=u1,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/u1'), (New-FixtureAdId 'CN=u3,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/u3'))
        }
        Mock Get-Group -ModuleName PRA.Collect {
            $members = $script:tree[[string]$Identity]
            if (-not $members) { throw "Group $Identity not found" }
            Use-Fixture $script:Kind ([pscustomobject]@{ Name = ($Identity -split ',')[0]; DistinguishedName = [string]$Identity; Members = (New-FixtureMulti $members) })
        }
    }
    It 'lists the user members once, walks the nested groups and ignores a loop, whatever the shape of the references (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        $result = @(Get-PcmGroupMemberDn -Context (New-TestContext) -Identity 'CN=G1,OU=Groups,DC=contoso,DC=com')
        $result | Sort-Object | Should -Be @('CN=u1,OU=PRA,DC=contoso,DC=com', 'CN=u2,OU=PRA,DC=contoso,DC=com', 'CN=u3,OU=PRA,DC=contoso,DC=com', 'CN=u9,OU=PRA,DC=contoso,DC=com')
    }
}

Describe 'Collect: trustees' {
    BeforeEach {
        Mock Get-Recipient -ModuleName PRA.Collect {
            $recipients = @{
                'CONTOSO\alice' = @{ Type = 'UserMailbox'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000a1'; Smtp = 'alice@contoso.com'; Name = 'alice'; Dn = 'CN=alice,OU=PRA,DC=contoso,DC=com' }
                'CN=alice,OU=PRA,DC=contoso,DC=com' = @{ Type = 'UserMailbox'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000a1'; Smtp = 'alice@contoso.com'; Name = 'alice'; Dn = 'CN=alice,OU=PRA,DC=contoso,DC=com' }
                'CONTOSO\finance' = @{ Type = 'MailUniversalSecurityGroup'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000f1'; Smtp = 'finance@contoso.com'; Name = 'finance'; Dn = 'CN=finance,OU=Groups,DC=contoso,DC=com' }
                'CONTOSO\jane' = @{ Type = 'MailContact'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000c1'; Smtp = 'jane@partner.com'; Name = 'jane'; Dn = 'CN=jane,OU=PRA,DC=contoso,DC=com' }
            }
            $entry = $recipients[[string]$Identity]
            if (-not $entry) { throw "Recipient $Identity not found" }
            Use-Fixture $script:Kind ([pscustomobject]@{ RecipientType = [PcmFixture.RecipientType]$entry.Type; Guid = [guid]$entry.Guid; PrimarySmtpAddress = $entry.Smtp; Name = $entry.Name; DistinguishedName = $entry.Dn })
        }
        Mock Get-User -ModuleName PRA.Collect {
            if ([string]$Identity -in @('CN=alice,OU=PRA,DC=contoso,DC=com', 'CONTOSO\bob')) { Use-Fixture $script:Kind ([pscustomobject]@{ UserPrincipalName = 'alice@contoso.com'; Guid = [guid]'11111111-aaaa-4aaa-8aaa-0000000000b1'; Name = 'bob'; DistinguishedName = 'CN=bob,OU=PRA,DC=contoso,DC=com' }) }
            else { throw "User $Identity not found" }
        }
        Mock Get-Group -ModuleName PRA.Collect {
            if ([string]$Identity -eq 'CONTOSO\legacy') { Use-Fixture $script:Kind ([pscustomobject]@{ Guid = [guid]'11111111-aaaa-4aaa-8aaa-0000000000d1'; Name = 'legacy'; DistinguishedName = 'CN=legacy,OU=Groups,DC=contoso,DC=com' }) }
            else { throw "Group $Identity not found" }
        }
    }
    It 'resolves <Trustee> to <Kind> (<Variant>)' -ForEach @(
        foreach ($variant in 'Live', 'Deserialized') {
            @{ Variant = $variant; Trustee = 'CONTOSO\alice'; Kind = 'User'; Resolved = $true; Smtp = 'alice@contoso.com'; Upn = 'alice@contoso.com' }
            @{ Variant = $variant; Trustee = 'CONTOSO\finance'; Kind = 'Group'; Resolved = $true; Smtp = 'finance@contoso.com'; Upn = $null }
            @{ Variant = $variant; Trustee = 'CONTOSO\jane'; Kind = 'Contact'; Resolved = $true; Smtp = 'jane@partner.com'; Upn = $null }
            @{ Variant = $variant; Trustee = 'CONTOSO\bob'; Kind = 'User'; Resolved = $true; Smtp = $null; Upn = 'alice@contoso.com' }
            @{ Variant = $variant; Trustee = 'CONTOSO\legacy'; Kind = 'Group'; Resolved = $true; Smtp = $null; Upn = $null }
            @{ Variant = $variant; Trustee = 'S-1-5-21-100-200-300-9999'; Kind = 'Unknown'; Resolved = $false; Smtp = $null; Upn = $null }
        }
    ) {
        $script:Kind = $Variant
        $result = Resolve-PcmTrustee -Context (New-TestContext) -Trustee $Trustee
        $result.Kind | Should -BeExactly $Kind
        $result.Resolved | Should -Be $Resolved
        $result.Smtp | Should -Be $Smtp
        $result.Upn | Should -Be $Upn
        if ($Resolved) { $result.Guid | Should -Match '^[0-9a-f]{8}-' }
    }
}

Describe 'Collect: the permissions of a shared mailbox' {
    BeforeEach {
        $script:shared = New-FixtureMailbox -Sam 'compta' -Type SharedMailbox -OnBehalf @((New-FixtureAdId 'CN=alice,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/alice'), (New-FixtureAdId 'CN=bob,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/bob'))
        $script:rightsFixture = @(
            @{ User = 'CONTOSO\alice'; Rights = @('FullAccess', 'ReadPermission'); Inherited = $false; Deny = $false }
            @{ User = 'CONTOSO\finance'; Rights = @('FullAccess'); Inherited = $false; Deny = $false }
            @{ User = 'NT AUTHORITY\SELF'; Rights = @('FullAccess'); Inherited = $false; Deny = $false }
            @{ User = 'S-1-5-21-100-200-300-9999'; Rights = @('FullAccess'); Inherited = $false; Deny = $false }
            @{ User = 'CONTOSO\Administrator'; Rights = @('FullAccess'); Inherited = $false; Deny = $false }
            @{ User = 'CONTOSO\jane'; Rights = @('FullAccess'); Inherited = $true; Deny = $false }
            @{ User = 'CONTOSO\jane'; Rights = @('FullAccess'); Inherited = $false; Deny = $true }
            @{ User = 'CONTOSO\jane'; Rights = @('ReadPermission'); Inherited = $false; Deny = $false }
        )
        Mock Get-MailboxPermission -ModuleName PRA.Collect {
            Use-Fixture $script:Kind @(foreach ($ace in $script:rightsFixture) {
                    [pscustomobject]@{ User = (New-FixturePrincipal $ace.User); AccessRights = (New-FixtureMulti @($ace.Rights | ForEach-Object { [PcmFixture.MailboxRights]$_ })); IsInherited = $ace.Inherited; Deny = $ace.Deny }
                })
        }
        Mock Get-ADPermission -ModuleName PRA.Collect {
            Use-Fixture $script:Kind @(
                [pscustomobject]@{ User = (New-FixturePrincipal 'CONTOSO\alice'); ExtendedRights = (New-FixtureMulti @((New-FixturePrincipal 'Send-As'))); IsInherited = $false; Deny = $false }
                [pscustomobject]@{ User = (New-FixturePrincipal 'CONTOSO\bob'); ExtendedRights = (New-FixtureMulti @((New-FixturePrincipal 'Receive-As'))); IsInherited = $false; Deny = $false }
                [pscustomobject]@{ User = (New-FixturePrincipal 'CONTOSO\alice'); ExtendedRights = (New-FixtureMulti @((New-FixturePrincipal 'Send-As'))); IsInherited = $true; Deny = $false }
            )
        }
        Mock Get-PcmDelegateListLink -ModuleName PRA.Collect { @('CN=alice,OU=PRA,DC=contoso,DC=com') }
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect {
            $known = @{
                'CONTOSO\alice' = @{ Kind = 'User'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000a1'; Upn = 'alice@contoso.com'; Smtp = 'alice@contoso.com'; Name = 'alice'; Dn = 'CN=alice,OU=PRA,DC=contoso,DC=com' }
                'CN=alice,OU=PRA,DC=contoso,DC=com' = @{ Kind = 'User'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000a1'; Upn = 'alice@contoso.com'; Smtp = 'alice@contoso.com'; Name = 'alice'; Dn = 'CN=alice,OU=PRA,DC=contoso,DC=com' }
                'CN=bob,OU=PRA,DC=contoso,DC=com' = @{ Kind = 'User'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000b1'; Upn = 'bob@contoso.com'; Smtp = 'bob@contoso.com'; Name = 'bob'; Dn = 'CN=bob,OU=PRA,DC=contoso,DC=com' }
                'CONTOSO\bob' = @{ Kind = 'User'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000b1'; Upn = 'bob@contoso.com'; Smtp = 'bob@contoso.com'; Name = 'bob'; Dn = 'CN=bob,OU=PRA,DC=contoso,DC=com' }
                'CONTOSO\finance' = @{ Kind = 'Group'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000f1'; Upn = $null; Smtp = 'finance@contoso.com'; Name = 'finance'; Dn = 'CN=finance,OU=Groups,DC=contoso,DC=com' }
                'CN=carol,OU=PRA,DC=contoso,DC=com' = @{ Kind = 'User'; Guid = '11111111-aaaa-4aaa-8aaa-0000000000c9'; Upn = 'carol@contoso.com'; Smtp = 'carol@contoso.com'; Name = 'carol'; Dn = 'CN=carol,OU=PRA,DC=contoso,DC=com' }
            }
            $entry = $known[[string]$Trustee]
            if ($entry) { [pscustomobject]@{ Resolved = $true; Guid = $entry.Guid; Kind = $entry.Kind; Upn = $entry.Upn; Smtp = $entry.Smtp; Name = $entry.Name; Dn = $entry.Dn; Note = '' } }
            else { [pscustomobject]@{ Resolved = $false; Guid = $null; Kind = 'Unknown'; Upn = $null; Smtp = $null; Name = $null; Dn = $null; Note = 'not found (deleted account or orphan SID)' } }
        }
        Mock Get-Group -ModuleName PRA.Collect {
            if ([string]$Identity -eq 'CN=finance,OU=Groups,DC=contoso,DC=com') { Use-Fixture $script:Kind ([pscustomobject]@{ Name = 'finance'; DistinguishedName = [string]$Identity; Members = (New-FixtureMulti @((New-FixtureAdId 'CN=carol,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/carol'))) }) }
            else { throw "Group $Identity not found" }
        }
    }
    It 'reads FullAccess, SendAs and SendOnBehalf, drops the inherited, denied, built-in and excluded entries, and expands the group (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        $result = Get-PcmSharedPermission -Context (New-TestContext) -Mailbox (Use-Fixture $script:Kind $script:shared)
        $rows = @($result.Permissions | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.access_right, $_.trustee, $_.trustee_kind, $_.auto_mapping })
        $rows | Should -Be @('FullAccess|CONTOSO\alice|User|True', 'FullAccess|CONTOSO\finance|Group|False', 'SendAs|CONTOSO\alice|User|', 'SendOnBehalf|CN=alice,OU=PRA,DC=contoso,DC=com|User|', 'SendOnBehalf|CN=bob,OU=PRA,DC=contoso,DC=com|User|')
        @($result.Members | ForEach-Object { '{0}|{1}' -f $_.member_upn, $_.member_kind }) | Should -Be @('carol@contoso.com|User')
        @($result.Warnings).Count | Should -Be 0
    }
    It 'gives the same rows for the live objects and their deserialized copies' {
        $live = Get-PcmSharedPermission -Context (New-TestContext) -Mailbox $script:shared
        $copy = $null
        $script:Kind = 'Deserialized'
        $copy = Get-PcmSharedPermission -Context (New-TestContext) -Mailbox (ConvertTo-DeserializedFixture $script:shared)
        ConvertTo-Comparable $copy.Permissions | Should -BeExactly (ConvertTo-Comparable $live.Permissions)
    }
    It 'warns about a trustee that cannot be resolved (an unknown account)' {
        $script:rightsFixture = @(@{ User = 'CONTOSO\ghost'; Rights = @('FullAccess'); Inherited = $false; Deny = $false })
        $result = Get-PcmSharedPermission -Context (New-TestContext) -Mailbox (ConvertTo-DeserializedFixture $script:shared)
        @($result.Warnings | Where-Object { $_ -like '*FullAccess trustee CONTOSO\ghost not found*' }).Count | Should -Be 1
    }
}

Describe 'Collect: target groups with the members of a remote session' {
    BeforeEach {
        $script:g1 = '11111111-aaaa-4aaa-8aaa-000000000001'; $script:g2 = '22222222-aaaa-4aaa-8aaa-000000000002'
        $script:dir = @{
            'CN=u1,OU=PRA,DC=contoso,DC=com' = @{ Class = 'User'; Name = 'u1'; Guid = 'aaaaaaaa-0000-4000-8000-000000000001'; Upn = 'u1@contoso.com'; Smtp = 'u1@contoso.com' }
            'CN=u3,OU=PRA,DC=contoso,DC=com' = @{ Class = 'User'; Name = 'u3'; Guid = 'aaaaaaaa-0000-4000-8000-000000000003'; Upn = 'u3@contoso.com'; Smtp = 'u3@contoso.com' }
            'CN=PC1,OU=PRA,DC=contoso,DC=com' = @{ Class = 'Computer'; Name = 'PC1'; Guid = 'aaaaaaaa-0000-4000-8000-0000000000c1'; Upn = ''; Smtp = '' }
            'CN=Jane,OU=PRA,DC=contoso,DC=com' = @{ Class = 'Contact'; Name = 'Jane'; Guid = 'aaaaaaaa-0000-4000-8000-0000000000c2'; Upn = ''; Smtp = 'jane@partner.com' }
            'CN=G2,OU=Groups,DC=contoso,DC=com' = @{ Class = 'Group'; Name = 'G2'; Guid = $script:g2; Upn = ''; Smtp = '' }
            'CN=G1,OU=Groups,DC=contoso,DC=com' = @{ Class = 'Group'; Name = 'PRA-Wave1'; Guid = $script:g1; Upn = ''; Smtp = '' }
        }
        $script:members = @{
            'CN=G1,OU=Groups,DC=contoso,DC=com' = @('CN=u1,OU=PRA,DC=contoso,DC=com', 'CN=PC1,OU=PRA,DC=contoso,DC=com', 'CN=Jane,OU=PRA,DC=contoso,DC=com', 'CN=G2,OU=Groups,DC=contoso,DC=com')
            'CN=G2,OU=Groups,DC=contoso,DC=com' = @('CN=u3,OU=PRA,DC=contoso,DC=com', 'CN=G1,OU=Groups,DC=contoso,DC=com', 'CN=u1,OU=PRA,DC=contoso,DC=com')
        }
        Mock Get-Group -ModuleName PRA.Collect {
            $id = [string]$Identity
            if ($id -in @('PRA-Wave1')) { $id = 'CN=G1,OU=Groups,DC=contoso,DC=com' }
            if (-not $script:members.ContainsKey($id)) { throw "Group $Identity not found" }
            $entry = $script:dir[$id]
            Use-Fixture $script:Kind ([pscustomobject]@{ Name = $entry.Name; SamAccountName = 'wave1'; DistinguishedName = $id; DisplayName = 'PRA Wave 1'; Guid = [guid]$entry.Guid
                    Members = (New-FixtureMulti @($script:members[$id] | ForEach-Object { New-FixtureAdId $_ ('contoso.com/x/' + $_) })) })
        }
        Mock Get-Recipient -ModuleName PRA.Collect {
            if ([string]$Identity -in @('CN=G1,OU=Groups,DC=contoso,DC=com', 'wave1@contoso.com')) { Use-Fixture $script:Kind ([pscustomobject]@{ RecipientType = [PcmFixture.RecipientType]::MailUniversalDistributionGroup; PrimarySmtpAddress = 'wave1@contoso.com'; DistinguishedName = 'CN=G1,OU=Groups,DC=contoso,DC=com' }) }
            else { throw "Recipient $Identity not found" }
        }
        Mock Get-PcmDirectoryObject -ModuleName PRA.Collect {
            $entry = $script:dir[$DistinguishedName]
            if (-not $entry) { return $null }
            [pscustomobject]@{ Class = $entry.Class; Name = $entry.Name; Sam = ''; Upn = [string]$entry.Upn; Smtp = [string]$entry.Smtp; DisplayName = ''; Dn = $DistinguishedName; Guid = $entry.Guid }
        }
    }
    It 'captures the group, its nested group, a loop and the members that are no mailbox, as it does for live objects (<Kind>, group given as <Written>)' -ForEach @(
        foreach ($kind in 'Live', 'Deserialized') {
            @{ Kind = $kind; Written = 'name in upper case'; Value = 'PRA-WAVE1' }
            @{ Kind = $kind; Written = 'distinguished name'; Value = 'CN=G1,OU=Groups,DC=contoso,DC=com' }
            @{ Kind = $kind; Written = 'SMTP address'; Value = 'WAVE1@contoso.com' }
        }
    ) {
        $script:Kind = $Kind
        $config = New-TestConfig; $config.Collect.TargetGroups = @($Value)
        $result = Get-PcmTargetGroup -Context (New-TestContext $config)
        @($result.Warnings).Count | Should -Be 0
        @($result.Groups).Count | Should -Be 1
        $result.Groups[0].group_guid | Should -Be $script:g1
        $result.Groups[0].primary_smtp_address | Should -Be 'wave1@contoso.com'
        $result.Groups[0].member_count | Should -Be 4
        $classes = @{}; foreach ($row in $result.Members) { $classes[$row.member_dn] = $row.member_class }
        $classes['CN=u1,OU=PRA,DC=contoso,DC=com'] | Should -Be 'User'
        $classes['CN=PC1,OU=PRA,DC=contoso,DC=com'] | Should -Be 'Computer'
        $classes['CN=Jane,OU=PRA,DC=contoso,DC=com'] | Should -Be 'Contact'
        $classes['CN=G2,OU=Groups,DC=contoso,DC=com'] | Should -Be 'Group'
        @($result.Members | Where-Object { $_.member_dn -eq 'CN=u1,OU=PRA,DC=contoso,DC=com' }).Count | Should -Be 1
    }
    It 'translates the canonical names that a remote session gives for the members (<Kind>)' -ForEach @(@{ Kind = 'Live' }, @{ Kind = 'Deserialized' }) {
        $script:Kind = $Kind
        $script:canonical = @{}
        foreach ($dn in $script:dir.Keys) { $script:canonical['CONTOSO.COM/x/' + (($dn -split ',')[0] -replace '^CN=', '')] = $dn }
        $script:members = @{
            'CN=G1,OU=Groups,DC=contoso,DC=com' = @('contoso.com/x/u1', 'contoso.com/x/PC1', 'contoso.com/x/Jane', 'contoso.com/x/G2')
            'CN=G2,OU=Groups,DC=contoso,DC=com' = @('contoso.com/x/u3', 'contoso.com/x/G1')
        }
        Mock Get-Group -ModuleName PRA.Collect {
            $id = [string]$Identity
            if ($id -in @('PRA-Wave1')) { $id = 'CN=G1,OU=Groups,DC=contoso,DC=com' }
            if (-not $script:members.ContainsKey($id)) { throw "Group $Identity not found" }
            $entry = $script:dir[$id]
            Use-Fixture $script:Kind ([pscustomobject]@{ Name = $entry.Name; SamAccountName = 'wave1'; DistinguishedName = $id; DisplayName = 'PRA Wave 1'; Guid = [guid]$entry.Guid
                    Members = (New-FixtureMulti @($script:members[$id])) })
        }
        Mock ConvertFrom-PcmCanonicalName -ModuleName PRA.Collect {
            foreach ($key in $script:canonical.Keys) { if ([string]::Equals($key, $CanonicalName, [StringComparison]::OrdinalIgnoreCase)) { return $script:canonical[$key] } }
            return ''
        }
        $config = New-TestConfig; $config.Collect.TargetGroups = @('pra-wave1')
        $result = Get-PcmTargetGroup -Context (New-TestContext $config)
        @($result.Warnings).Count | Should -Be 0
        $result.Groups[0].member_count | Should -Be 4
        @($result.Members | ForEach-Object { $_.member_dn }) | Should -Contain 'CN=u3,OU=PRA,DC=contoso,DC=com'
        @($result.Members | Where-Object { $_.member_dn -notmatch '^CN=' }).Count | Should -Be 0
        Should -Invoke Get-PcmDirectoryObject -ModuleName PRA.Collect -ParameterFilter { $DistinguishedName -match '/' } -Times 0
    }
    It 'does not capture a group with a member that Active Directory cannot translate, and says so' {
        $script:Kind = 'Deserialized'
        $script:members = @{ 'CN=G1,OU=Groups,DC=contoso,DC=com' = @('contoso.com/x/u1', 'contoso.com/x/gone') }
        Mock Get-Group -ModuleName PRA.Collect {
            $entry = $script:dir['CN=G1,OU=Groups,DC=contoso,DC=com']
            Use-Fixture $script:Kind ([pscustomobject]@{ Name = $entry.Name; SamAccountName = 'wave1'; DistinguishedName = 'CN=G1,OU=Groups,DC=contoso,DC=com'; DisplayName = 'PRA Wave 1'; Guid = [guid]$entry.Guid
                    Members = (New-FixtureMulti @($script:members['CN=G1,OU=Groups,DC=contoso,DC=com'])) })
        }
        Mock ConvertFrom-PcmCanonicalName -ModuleName PRA.Collect { if ($CanonicalName -eq 'contoso.com/x/u1') { 'CN=u1,OU=PRA,DC=contoso,DC=com' } else { '' } }
        $config = New-TestConfig; $config.Collect.TargetGroups = @('PRA-Wave1')
        $result = Get-PcmTargetGroup -Context (New-TestContext $config)
        @($result.Groups).Count | Should -Be 0
        $result.Warnings[0] | Should -Match 'contoso\.com/x/gone cannot be found in Active Directory'
    }
    It 'gives the same rows for the live objects and their deserialized copies' {
        $config = New-TestConfig; $config.Collect.TargetGroups = @('PRA-Wave1')
        $script:Kind = 'Live'
        $live = Get-PcmTargetGroup -Context (New-TestContext $config)
        $script:Kind = 'Deserialized'
        $copy = Get-PcmTargetGroup -Context (New-TestContext $config)
        ConvertTo-Comparable $copy.Groups | Should -BeExactly (ConvertTo-Comparable $live.Groups)
        ConvertTo-Comparable $copy.Members | Should -BeExactly (ConvertTo-Comparable $live.Members)
    }
}

Describe 'Collect: native resource settings from a remote session' {
    BeforeEach {
        $script:room = New-FixtureMailbox -Sam 'room' -Type RoomMailbox
        $script:room | Add-Member -NotePropertyName ResourceCapacity -NotePropertyValue 12
        $script:room | Add-Member -NotePropertyName ResourceCustom -NotePropertyValue (New-FixtureMulti @('Video', 'Whiteboard'))
        $script:room | Add-Member -NotePropertyName RequireSenderAuthenticationEnabled -NotePropertyValue $true
        $script:room | Add-Member -NotePropertyName ModerationEnabled -NotePropertyValue $true
        $script:room | Add-Member -NotePropertyName SendModerationNotifications -NotePropertyValue ([PcmFixture.ModerationNotifications]::Internal)
        $script:room | Add-Member -NotePropertyName AcceptMessagesOnlyFromSendersOrMembers -NotePropertyValue (New-FixtureMulti @((New-FixtureAdId 'CN=bookers,OU=Groups,DC=contoso,DC=com' 'contoso.com/Groups/bookers')))
        $script:room | Add-Member -NotePropertyName RejectMessagesFromSendersOrMembers -NotePropertyValue (New-FixtureMulti @())
        $script:room | Add-Member -NotePropertyName ModeratedBy -NotePropertyValue (New-FixtureMulti @((New-FixtureAdId 'CN=delegate,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/delegate')))
        $script:room | Add-Member -NotePropertyName BypassModerationFromSendersOrMembers -NotePropertyValue (New-FixtureMulti @())
        $script:booking = [pscustomobject]@{
            AutomateProcessing = [PcmFixture.AutomateProcessing]::AutoAccept; AllBookInPolicy = $false; AllRequestInPolicy = $true; AllRequestOutOfPolicy = $false
            ResourceDelegates = (New-FixtureMulti @((New-FixtureAdId 'CN=delegate,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/delegate'))); BookInPolicy = (New-FixtureMulti @())
            RequestInPolicy = (New-FixtureMulti @((New-FixtureAdId 'CN=bookers,OU=Groups,DC=contoso,DC=com' 'contoso.com/Groups/bookers'))); RequestOutOfPolicy = (New-FixtureMulti @())
            ScheduleOnlyDuringWorkHours = $true; AllowConflicts = $false; MaximumDurationInMinutes = 120; BookingWindowInDays = 180; AdditionalResponse = 'Booking policy'
        }
        $script:hours = [pscustomobject]@{ WorkingHoursStartTime = [timespan]'08:30:00'; WorkingHoursEndTime = [timespan]'18:00:00'
            WorkDays = [PcmFixture.WorkDays]'Monday, Wednesday, Friday'; WorkingHoursTimeZone = (New-FixtureTimeZone 'Romance Standard Time' '(UTC+01:00) Brussels, Copenhagen, Madrid, Paris') }
        $script:roomUser = [pscustomobject]@{ Office = 'Building A / room 101'; StreetAddress = '1 Test Street'; City = 'Paris'; StateOrProvince = ''; PostalCode = '75001'; CountryOrRegion = 'France'; Phone = '+33100000000'; Notes = 'Booking reception' }
        $script:roomList = [pscustomobject]@{
            Guid = [guid]'44444444-aaaa-4aaa-8aaa-000000000004'; Name = 'building-a'; DisplayName = 'Building A'; Alias = 'building-a'; PrimarySmtpAddress = 'building-a@contoso.com'
            HiddenFromAddressListsEnabled = $false; EmailAddresses = (New-FixtureMulti @((New-FixtureProxy 'SMTP' 'building-a@contoso.com'))); LegacyExchangeDN = '/o=Contoso/cn=building-a'
            MemberJoinRestriction = [PcmFixture.JoinRestriction]::Closed; MemberDepartRestriction = [PcmFixture.DepartRestriction]::Open
            RequireSenderAuthenticationEnabled = $false; ModerationEnabled = $false; SendModerationNotifications = [PcmFixture.ModerationNotifications]::Never
            AcceptMessagesOnlyFromSendersOrMembers = (New-FixtureMulti @()); RejectMessagesFromSendersOrMembers = (New-FixtureMulti @()); ModeratedBy = (New-FixtureMulti @()); BypassModerationFromSendersOrMembers = (New-FixtureMulti @())
        }
        Mock Get-CalendarProcessing -ModuleName PRA.Collect { Use-Fixture $script:Kind $script:booking }
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Collect { Use-Fixture $script:Kind $script:hours }
        Mock Get-User -ModuleName PRA.Collect { Use-Fixture $script:Kind $script:roomUser }
        Mock Get-DistributionGroup -ModuleName PRA.Collect { Use-Fixture $script:Kind @($script:roomList) }
        Mock Get-DistributionGroupMember -ModuleName PRA.Collect { Use-Fixture $script:Kind @([pscustomobject]@{ Guid = $script:room.Guid; Name = 'room' }) }
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect {
            $smtp = switch -Regex ([string]$Trustee) { 'delegate' { 'delegate@contoso.com' } 'bookers' { 'bookers@contoso.com' } default { $null } }
            [pscustomobject]@{ Resolved = [bool]$smtp; Smtp = $smtp }
        }
    }
    It 'captures the same settings from the live objects and their deserialized copies' {
        $script:Kind = 'Live'
        $live = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:room
        $script:Kind = 'Deserialized'
        $copy = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox (ConvertTo-DeserializedFixture $script:room)
        ConvertTo-Comparable $copy | Should -BeExactly (ConvertTo-Comparable $live)
        $copy.Capacity | Should -Be 12
        $copy.ResourceCustom | Should -Be @('Video', 'Whiteboard')
        $copy.Calendar.AutomateProcessing | Should -BeExactly 'AutoAccept'
        $copy.Calendar.MaximumDurationInMinutes | Should -Be 120
        $copy.Calendar.BookingWindowInDays | Should -Be 180
        $copy.Calendar.ResourceDelegates | Should -Be @('delegate@contoso.com')
        $copy.Calendar.RequestInPolicy | Should -Be @('bookers@contoso.com')
        $copy.CalendarConfiguration.WorkingHoursStartTime | Should -Be '08:30:00'
        $copy.CalendarConfiguration.WorkingHoursEndTime | Should -Be '18:00:00'
        $copy.CalendarConfiguration.WorkingHoursDays | Should -Be @('Monday', 'Wednesday', 'Friday')
        $copy.CalendarConfiguration.WorkingHoursTimeZone | Should -Be 'Romance Standard Time'
        $copy.Delivery.SendModerationNotifications | Should -BeExactly 'Internal'
        $copy.Delivery.ModerationEnabled | Should -BeTrue
        $copy.Delivery.ModeratedBy | Should -Be @('delegate@contoso.com')
        $copy.Delivery.AcceptMessagesOnlyFromSendersOrMembers | Should -Be @('bookers@contoso.com')
        @($copy.RoomLists).Count | Should -Be 1
        $copy.RoomLists[0].MemberJoinRestriction | Should -BeExactly 'Closed'
        $copy.RoomLists[0].PrimarySmtpAddress | Should -Be 'building-a@contoso.com'
        $copy.RoomLists[0].EmailAddresses | Should -Be @('SMTP:building-a@contoso.com')
        $copy.Directory.Office | Should -Be 'Building A / room 101'
    }
    It 'reads a deserialized WorkDays of <Days>' -ForEach @(@{ Days = 'Weekdays' }, @{ Days = 'WeekendDays' }, @{ Days = 'AllDays' }) {
        $script:hours.WorkDays = [PcmFixture.WorkDays]$Days
        $script:Kind = 'Deserialized'
        $settings = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox (ConvertTo-DeserializedFixture $script:room)
        $settings.CalendarConfiguration.WorkingHoursDays | Should -Be @($Days)
    }
    It 'stores the calendar of a deserialized copy as JSON a later run can read back' {
        $script:Kind = 'Deserialized'
        $mailbox = ConvertTo-DeserializedFixture $script:room
        $settings = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $mailbox
        $stored = (ConvertTo-PcmMailboxRecord -Mailbox $mailbox -ResourceSettings $settings).resource_settings_json | ConvertFrom-Json
        $stored.SchemaVersion | Should -Be 2
        $stored.Calendar.AutomateProcessing | Should -BeExactly 'AutoAccept'
        $stored.Calendar.AllRequestInPolicy | Should -BeTrue
        @($stored.CalendarConfiguration.WorkingHoursDays).Count | Should -Be 3
    }
}

Describe 'Collect: contacts and distribution groups from a remote session' {
    BeforeEach {
        Mock Get-MailContact -ModuleName PRA.Collect {
            $contact = [pscustomobject]@{ Guid = [guid]'55555555-aaaa-4aaa-8aaa-000000000005'; Name = 'Jane'; DisplayName = 'Jane Partner'; Alias = 'jane'; ExternalEmailAddress = (New-FixtureProxy 'SMTP' 'jane@partner.com')
                PrimarySmtpAddress = 'jane@partner.com'; LegacyExchangeDN = '/o=Contoso/cn=jane'; HiddenFromAddressListsEnabled = $false; OrganizationalUnit = (New-FixtureAdId 'OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA')
                EmailAddresses = (New-FixtureMulti @((New-FixtureProxy 'SMTP' 'jane@partner.com'), (New-FixtureProxy 'smtp' 'j@partner.com'))); CustomAttribute3 = 'Partner' }
            Use-Fixture $script:Kind @($contact)
        }
        Mock Get-DistributionGroup -ModuleName PRA.Collect {
            $group = [pscustomobject]@{ Guid = [guid]'66666666-aaaa-4aaa-8aaa-000000000006'; Name = 'All'; DisplayName = 'Everybody'; Alias = 'all'; PrimarySmtpAddress = 'all@contoso.com'; DistinguishedName = 'CN=All,OU=Groups,DC=contoso,DC=com'
                LegacyExchangeDN = '/o=Contoso/cn=all'; HiddenFromAddressListsEnabled = $false; OrganizationalUnit = (New-FixtureAdId 'OU=Groups,DC=contoso,DC=com' 'contoso.com/Groups'); RecipientTypeDetails = [PcmFixture.RecipientType]::MailUniversalSecurityGroup
                EmailAddresses = (New-FixtureMulti @((New-FixtureProxy 'SMTP' 'all@contoso.com'))); ManagedBy = (New-FixtureMulti @((New-FixtureAdId 'CN=alice,OU=PRA,DC=contoso,DC=com' 'contoso.com/PRA/alice')))
                RequireSenderAuthenticationEnabled = $true; MemberJoinRestriction = [PcmFixture.JoinRestriction]::Closed; MemberDepartRestriction = [PcmFixture.DepartRestriction]::Open
                ModerationEnabled = $false; SendModerationNotifications = [PcmFixture.ModerationNotifications]::Never; AcceptMessagesOnlyFromSendersOrMembers = (New-FixtureMulti @())
                MaxSendSize = [PcmFixture.UnlimitedSize]::new(); MaxReceiveSize = (New-FixtureSize 10485760); Notes = 'all staff' }
            Use-Fixture $script:Kind @($group)
        }
        Mock Get-DistributionGroupMember -ModuleName PRA.Collect {
            Use-Fixture $script:Kind @([pscustomobject]@{ Guid = [guid]'11111111-aaaa-4aaa-8aaa-0000000000a1'; RecipientType = [PcmFixture.RecipientType]::UserMailbox; Name = 'alice'; PrimarySmtpAddress = 'alice@contoso.com' })
        }
        Mock Get-DynamicDistributionGroup -ModuleName PRA.Collect {
            Use-Fixture $script:Kind @([pscustomobject]@{ Guid = [guid]'77777777-aaaa-4aaa-8aaa-000000000007'; Name = 'Dyn'; DisplayName = 'Dynamic'; Alias = 'dyn'; PrimarySmtpAddress = 'dyn@contoso.com'; LegacyExchangeDN = '/o=Contoso/cn=dyn'
                    HiddenFromAddressListsEnabled = $false; OrganizationalUnit = (New-FixtureAdId 'OU=Groups,DC=contoso,DC=com' 'contoso.com/Groups'); EmailAddresses = (New-FixtureMulti @((New-FixtureProxy 'SMTP' 'dyn@contoso.com')))
                    ManagedBy = (New-FixtureMulti @()); RecipientFilter = "(RecipientType -eq 'UserMailbox')" })
        }
    }
    It 'reads the mail contacts the same from the live objects and their deserialized copies' {
        $script:Kind = 'Live'; $live = @(Get-PcmContact -Context (New-TestContext))
        $script:Kind = 'Deserialized'; $copy = @(Get-PcmContact -Context (New-TestContext))
        ConvertTo-Comparable $copy | Should -BeExactly (ConvertTo-Comparable $live)
        $copy[0].organizational_unit | Should -BeExactly 'contoso.com/PRA'
        $copy[0].external_email_address | Should -BeExactly 'SMTP:jane@partner.com'
        @($copy[0].email_addresses_json | ConvertFrom-Json) | Should -Be @('SMTP:jane@partner.com', 'smtp:j@partner.com')
        ($copy[0].custom_attributes_json | ConvertFrom-Json).CustomAttribute3 | Should -BeExactly 'Partner'
    }
    It 'reads the distribution groups the same from the live objects and their deserialized copies' {
        $config = New-TestConfig; $config.Collect.DynamicDistributionGroups = $true
        $script:Kind = 'Live'; $live = Get-PcmDistributionGroup -Context (New-TestContext $config)
        $script:Kind = 'Deserialized'; $copy = Get-PcmDistributionGroup -Context (New-TestContext $config)
        ConvertTo-Comparable $copy | Should -BeExactly (ConvertTo-Comparable $live)
        @($copy.Groups).Count | Should -Be 2
        $copy.Groups[0].kind | Should -BeExactly 'Security'
        $copy.Groups[0].organizational_unit | Should -BeExactly 'contoso.com/Groups'
        @($copy.Groups[0].managed_by_json | ConvertFrom-Json) | Should -Be @('contoso.com/PRA/alice')
        $settings = $copy.Groups[0].settings_json | ConvertFrom-Json
        $settings.MemberJoinRestriction | Should -BeExactly 'Closed'
        $settings.MaxSendSize | Should -BeExactly 'Unlimited'
        $settings.MaxReceiveSize | Should -BeExactly '10.0 MB (10,485,760 bytes)'
        $copy.Groups[1].kind | Should -BeExactly 'Dynamic'
        @($copy.Members).Count | Should -Be 1
        $copy.Members[0].member_kind | Should -BeExactly 'UserMailbox'
    }
}

Describe 'Collect: the remote session to Exchange' {
    BeforeEach {
        $script:sessionLog = [Collections.Generic.List[string]]::new()
        $script:session = [pscustomobject]@{ Id = 7 }
        $script:importedFrom = $null
        function global:New-PSSession { [CmdletBinding()] param($ConfigurationName, $ConnectionUri, $Authentication) $script:sessionLog.Add(('{0}|{1}|{2}' -f $ConfigurationName, $ConnectionUri, $Authentication)); if ($script:sessionFails) { throw $script:sessionFails }; $script:session }
        function global:Import-PSSession { [CmdletBinding()] param($Session, $CommandName, [switch]$DisableNameChecking, [switch]$AllowClobber) $script:importedFrom = $Session; $script:importedNames = @($CommandName); New-Module -Name PcmFakeExchange -ScriptBlock { } }
        function global:Remove-PSSession { [CmdletBinding()] param($Session) $script:sessionLog.Add('removed') }
        $script:sessionFails = $null
        Mock Get-ExchangeServer -ModuleName PRA.Collect { ConvertTo-DeserializedFixture @([pscustomobject]@{ Name = 'EX1'; AdminDisplayVersion = ([PcmFixture.NameOnlyReference]@{ Name = 'v'; Text = 'Version 15.2 (Build 1544.4)' }) }, [pscustomobject]@{ Name = 'EX2'; AdminDisplayVersion = 'other' }) }
        Mock Set-ADServerSettings -ModuleName PRA.Collect { }
    }
    AfterEach {
        foreach ($name in 'New-PSSession', 'Import-PSSession', 'Remove-PSSession') { Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue }
        Remove-Module PcmFakeExchange -Force -ErrorAction SilentlyContinue
    }
    It 'opens a Kerberos session to the PowerShell endpoint of the Exchange server, imports the Exchange commands and reports the version' {
        $config = New-TestConfig; $config.Exchange.Server = 'EX1.contoso.com'
        $context = New-TestContext $config
        $how = Connect-PcmExchange -Context $context
        $script:sessionLog[0] | Should -BeExactly 'Microsoft.Exchange|http://EX1.contoso.com/PowerShell/|Kerberos'
        $script:importedFrom.Id | Should -Be 7
        $script:importedNames | Should -Contain 'Get-Mailbox'
        $script:importedNames | Should -Contain 'Get-DistributionGroupMember'
        $context.Server | Should -BeExactly 'EX1.contoso.com'
        $how | Should -BeLike 'remote session to EX1.contoso.com (Kerberos) *Version 15.2 (Build 1544.4)'
        Should -Invoke Set-ADServerSettings -ModuleName PRA.Collect -Times 1 -ParameterFilter { $ViewEntireForest -eq $true }
    }
    It 'reports no version when the server is not in the list, and still connects' {
        $config = New-TestConfig; $config.Exchange.Server = 'other.contoso.com'
        (Connect-PcmExchange -Context (New-TestContext $config)) | Should -BeExactly 'remote session to other.contoso.com (Kerberos)'
    }
    It 'removes the session when the Exchange side is closed' {
        $context = New-TestContext
        $null = Connect-PcmExchange -Context $context
        Disconnect-PcmExchange -Context $context
        $script:sessionLog | Should -Contain 'removed'
        $context['ExchangeSession'] | Should -BeNullOrEmpty
        { Disconnect-PcmExchange -Context $context } | Should -Not -Throw
    }
    It 'refuses an empty Exchange.Server in plain words, before any session' {
        $config = New-TestConfig; $config.Exchange.Server = '  '
        $message = $null
        try { $null = Connect-PcmExchange -Context (New-TestContext $config) } catch { $message = $_.Exception.Message }
        $message | Should -BeExactly 'Exchange.Server is empty. Collect reads Exchange through a remote session: set Exchange.Server to the host name of an Exchange server (for example EX01.contoso.com) in the configuration file.'
        $script:sessionLog.Count | Should -Be 0
    }
    It 'says what to check when the session cannot be opened, and leaves nothing open' {
        $script:sessionFails = 'The WinRM client cannot process the request.'
        $context = New-TestContext
        $message = $null
        try { $null = Connect-PcmExchange -Context $context } catch { $message = $_.Exception.Message }
        $message | Should -BeExactly 'Cannot open the remote session to Exchange on ex1.contoso.com: The WinRM client cannot process the request. Collect needs a domain-joined computer, an account allowed to use Exchange remote PowerShell, and an Exchange server that answers on http://ex1.contoso.com/PowerShell/ with Kerberos.'
        $context['ExchangeSession'] | Should -BeNullOrEmpty
    }
    It 'removes the session and says what to check when the commands cannot be imported' {
        function global:Import-PSSession { [CmdletBinding()] param($Session, $CommandName, [switch]$DisableNameChecking, [switch]$AllowClobber) throw 'Access denied by the remote endpoint.' }
        $context = New-TestContext
        $message = $null
        try { $null = Connect-PcmExchange -Context $context } catch { $message = $_.Exception.Message }
        $message | Should -BeLike 'Cannot open the remote session to Exchange on ex1.contoso.com: Access denied by the remote endpoint.*'
        $script:sessionLog | Should -Contain 'removed'
    }
}
