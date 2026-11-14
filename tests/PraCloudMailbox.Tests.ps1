<#
.SYNOPSIS
    PRA Cloud Mailbox - offline tests (Pester 5+). No Exchange, Active Directory or cloud access: the
    Exchange cmdlets are synthetic functions, the cloud rules are pure functions.
.DESCRIPTION
    Run in PowerShell 7.5 or later:
        pwsh -NoProfile -File .\tests\Invoke-TestGate.ps1
.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
#>
#Requires -Version 7.5

BeforeAll {
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    Import-Module (Join-Path $script:Root 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Store.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Collect.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Cloud.psm1') -Force
    . (Join-Path $PSScriptRoot 'GraphRequestContext.Fake.ps1')
    . (Join-Path $PSScriptRoot 'DeserializedFixture.ps1')
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    $script:RemoteStubs = $global:PcmRemoteStubs

    function script:New-TestConfig {
        param([hashtable]$Override = @{})
        $config = @{
            Environment = 'TEST'
            Exchange = @{ Server = 'ex1.contoso.com'; DomainController = '' }
            Scope = @{ Mode = 'OU'; SearchBase = 'OU=PRA,DC=contoso,DC=com'; GroupDN = ''; CsvPath = ''; IncludeUsers = $true; IncludeShared = $true; IncludeRoom = $false; IncludeEquipment = $false; ExcludeSamAccountNames = @(); ConvertRooms = $false }
            Collect = @{ SharedPermissions = $true; ExpandGroupTrustees = $true; ExcludeTrustees = @('Administrator'); MailboxStatistics = $false; Contacts = $false; ContactsSearchBase = ''; DistributionGroups = $false; DistributionGroupsSearchBase = ''; DynamicDistributionGroups = $false }
            Store = @{ Path = '.\data\test.db'; KeepSnapshots = 10; BackupFolder = '.\data\backup'; MaxSnapshotAgeDays = 7; JournalPath = '.\data\journal.db' }
            Cloud = @{ TenantId = ''; Organization = ''; AppId = ''; CertificateThumbprint = ''; DefaultUsageLocation = '' }
            Polling = @{ IntervalSeconds = 20; MailboxTimeoutMinutes = 30; SyncTimeoutMinutes = 30; HoldTimeoutMinutes = 30 }
            Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111'; SkuPartNumber = '' }; Shared = @{ SkuPartNumber = 'E5'; Parallel = 100 } }
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
            RequireSenderAuthenticationEnabled = $true; ModerationEnabled = $false; SendModerationNotifications = 'Never'
            AcceptMessagesOnlyFromSendersOrMembers = @(); RejectMessagesFromSendersOrMembers = @(); ModeratedBy = @(); BypassModerationFromSendersOrMembers = @()
        }
    }
    function script:New-FakeRoomList {
        $group = New-FakeMailbox -Sam 'building-a' -Type RoomList
        $group | Add-Member -NotePropertyName MemberJoinRestriction -NotePropertyValue Closed
        $group | Add-Member -NotePropertyName MemberDepartRestriction -NotePropertyValue Closed
        return $group
    }
}

AfterAll {
    foreach ($name in @('Get-Mailbox','Get-MailboxStatistics','Get-CalendarProcessing','Get-MailboxCalendarConfiguration','Get-MailboxPermission','Get-ADPermission','Get-Recipient','Get-User','Get-Group','Get-MailContact',
            'Get-DistributionGroup','Get-DistributionGroupMember','Get-DynamicDistributionGroup','Get-ExchangeServer')) {
        Remove-Item -Path "function:global:$name" -ErrorAction SilentlyContinue
    }
}

Describe 'Configuration' {
    BeforeEach { $script:folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory $script:folder }

    It 'accepts the shipped configuration file' {
        $config = Import-PcmConfiguration -Path (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1') -Root $script:Root
        $config.Store.Path | Should -Be (Join-Path $script:Root 'data\PraCloudMailbox.db')
        $config.Licensing.Users.Mode | Should -Be 'Group'
    }
    It 'adds the default values of missing settings' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        $config = Import-PcmConfiguration -Path $path -Root $script:folder
        $config.Store.KeepSnapshots | Should -Be 10
        $config.Collect.ExpandGroupTrustees | Should -BeTrue
        $config.Retention.TagValue | Should -Be 'Converted'
        $config.Scope.ResourceMode | Should -Be 'SynchronizedShared'
        $config.Resources.OnPremisesReady | Should -BeFalse
        $config.Resources.AllowUnprotectedDeletion | Should -BeFalse
        $config.Resources.Parallel | Should -Be 100
    }
    It 'keeps the former behaviour when KeepFree and Recover.WaveSize are absent' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' }; Licensing = @{ Shared = @{ Parallel = 40 } } }" -Encoding UTF8
        $config = Import-PcmConfiguration -Path $path -Root $script:folder
        $config.Licensing.Shared.KeepFree | Should -Be 0
        $config.Recover.WaveSize | Should -Be 0
        Get-PcmRecoverWaveSize $config | Should -Be 40
    }
    It 'gives Recover its own wave size, independent from Licensing.Shared.Parallel' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' }; Licensing = @{ Shared = @{ Parallel = 40; KeepFree = 5 } }; recover = @{ wavesize = 7 } }" -Encoding UTF8
        $config = Import-PcmConfiguration -Path $path -Root $script:folder
        $config.Licensing.Shared.KeepFree | Should -Be 5
        $config.Recover.WaveSize | Should -Be 7
        Get-PcmRecoverWaveSize $config | Should -Be 7
        $config.Licensing.Shared.Parallel | Should -Be 40
    }
    It 'reads an empty Recover.WaveSize (<Value>) like an absent one' -ForEach @(@{ Value = "''" }, @{ Value = "'  '" }, @{ Value = '$null' }, @{ Value = '0' }) {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' }; Licensing = @{ Shared = @{ Parallel = 30 } }; Recover = @{ WaveSize = $Value } }" -Encoding UTF8
        Get-PcmRecoverWaveSize (Import-PcmConfiguration -Path $path -Root $script:folder) | Should -Be 30
    }
    It 'lists KeepFree and Recover.WaveSize among the settings a run changes' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        $base = Import-PcmConfiguration -Path $path -Root $script:folder
        $run = Import-PcmConfiguration -Path $path -Root $script:folder
        $run.Licensing.Shared.KeepFree = 10; $run.Recover.WaveSize = 20
        @(Get-PcmConfigurationDifference -Base $base -Effective $run | ForEach-Object { $_.Text }) | Should -Be @('Licensing.Shared.KeepFree: 0 -> 10', 'Recover.WaveSize: 0 -> 20')
    }
    It 'refuses an unknown setting' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto'; IncludeUser = `$true } }" -Encoding UTF8
        { Import-PcmConfiguration -Path $path -Root $script:folder } | Should -Throw '*unknown setting Scope.IncludeUser*'
    }
    It 'refuses the former Exchange.ConnectionMode Local (<Value>) with the way to migrate' -ForEach @(@{ Value = 'Local' }, @{ Value = 'LOCAL' }, @{ Value = ' local ' }) {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Exchange = @{ ConnectionMode = '$Value'; Server = 'ex1.contoso.com' }; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        $message = $null
        try { $null = Import-PcmConfiguration -Path $path -Root $script:folder } catch { $message = $_.Exception.Message }
        $message | Should -BeExactly "Configuration: Exchange.ConnectionMode = 'Local' (the Exchange Management Shell of this server) is not supported any more. Collect now always reads Exchange through a remote session: remove ConnectionMode and set Exchange.Server to the host name of an Exchange server."
    }
    It 'accepts the former Exchange.ConnectionMode Remote (<Value>) and ignores it' -ForEach @(@{ Value = 'Remote' }, @{ Value = 'REMOTE' }, @{ Value = ' remote ' }) {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Exchange = @{ ConnectionMode = '$Value'; Server = 'ex1.contoso.com' }; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        $config = Import-PcmConfiguration -Path $path -Root $script:folder
        $config.Exchange.ContainsKey('ConnectionMode') | Should -BeFalse
        $config.Exchange.Server | Should -BeExactly 'ex1.contoso.com'
    }
    It 'refuses any other Exchange.ConnectionMode, and needs none' {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Exchange = @{ ConnectionMode = 'Hybrid' }; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        $message = $null
        try { $null = Import-PcmConfiguration -Path $path -Root $script:folder } catch { $message = $_.Exception.Message }
        $message | Should -BeExactly 'Configuration: Exchange.ConnectionMode is no longer a setting (Collect always reads Exchange through a remote session): remove it and set Exchange.Server.'
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' } }" -Encoding UTF8
        (Import-PcmConfiguration -Path $path -Root $script:folder).Exchange.Server | Should -BeExactly ''
    }
    It 'refuses <Case>' -ForEach @(
        @{ Case = 'OU mode without SearchBase'; Text = "Scope = @{ Mode = 'OU'; SearchBase = '' }"; Message = '*requires Scope.SearchBase*' }
        @{ Case = 'an AppId without certificate'; Text = "Scope = @{ Mode = 'Auto' }; Cloud = @{ AppId = 'x' }"; Message = '*go together*' }
        @{ Case = 'an invalid tag attribute'; Text = "Scope = @{ Mode = 'Auto' }; Retention = @{ TagAttribute = 'extensionAttribute1' }"; Message = '*CustomAttribute1 to CustomAttribute15*' }
        @{ Case = 'Direct mode without SKU'; Text = "Scope = @{ Mode = 'Auto' }; Licensing = @{ Users = @{ Mode = 'Direct' } }"; Message = '*requires Licensing.Users.SkuPartNumber*' }
        @{ Case = 'no mailbox type'; Text = "Scope = @{ Mode = 'Auto'; IncludeUsers = `$false; IncludeShared = `$false }"; Message = '*at least one mailbox type*' }
        @{ Case = 'a text instead of a number'; Text = "Scope = @{ Mode = 'Auto' }; Store = @{ KeepSnapshots = '5' }"; Message = '*whole number*' }
        @{ Case = 'a KeepFree below zero'; Text = "Scope = @{ Mode = 'Auto' }; Licensing = @{ Shared = @{ KeepFree = -1 } }"; Message = '*Licensing.Shared.KeepFree*' }
        @{ Case = 'a KeepFree given as text'; Text = "Scope = @{ Mode = 'Auto' }; Licensing = @{ Shared = @{ KeepFree = ' 10 ' } }"; Message = '*Licensing.Shared.KeepFree must be a whole number*' }
        @{ Case = 'a fractional KeepFree'; Text = "Scope = @{ Mode = 'Auto' }; Licensing = @{ Shared = @{ KeepFree = 1.5 } }"; Message = '*Licensing.Shared.KeepFree*' }
        @{ Case = 'a Recover wave size above 1000'; Text = "Scope = @{ Mode = 'Auto' }; Recover = @{ WaveSize = 1001 } "; Message = '*Recover.WaveSize*' }
        @{ Case = 'a Recover wave size below zero'; Text = "Scope = @{ Mode = 'Auto' }; Recover = @{ WaveSize = -5 } "; Message = '*Recover.WaveSize*' }
        @{ Case = 'a Recover wave size given as text'; Text = "Scope = @{ Mode = 'Auto' }; Recover = @{ WaveSize = ' 40 ' } "; Message = '*Recover.WaveSize must be a whole number*' }
        @{ Case = 'a fractional Recover wave size'; Text = "Scope = @{ Mode = 'Auto' }; Recover = @{ WaveSize = 2.5 } "; Message = '*Recover.WaveSize*' }
        @{ Case = 'an unknown resource mode'; Text = "Scope = @{ Mode = 'Auto'; ResourceMode = 'Automatic' }"; Message = '*Scope.ResourceMode*' }
        @{ Case = 'an invalid resource group'; Text = "Scope = @{ Mode = 'Auto' }; Resources = @{ LicensingGroupId = 'invalid' }"; Message = '*Resources.LicensingGroupId*' }
        @{ Case = 'a resource routing URL'; Text = "Scope = @{ Mode = 'Auto' }; Resources = @{ OnPremisesRoutingDomain = 'https://contoso.com' }"; Message = '*Resources.OnPremisesRoutingDomain*' }
        @{ Case = 'an invalid resource wave size'; Text = "Scope = @{ Mode = 'Auto' }; Resources = @{ Parallel = 0 }"; Message = '*Resources.Parallel*' }
        @{ Case = 'a fractional resource wave size'; Text = "Scope = @{ Mode = 'Auto' }; Resources = @{ Parallel = 1.5 }"; Message = '*Resources.Parallel*' }
        @{ Case = 'text deletion consent'; Text = "Scope = @{ Mode = 'Auto' }; Resources = @{ AllowUnprotectedDeletion = 'true' }"; Message = '*Resources.AllowUnprotectedDeletion must be*' }
        @{ Case = 'numeric routing consent'; Text = "Scope = @{ Mode = 'Auto' }; Resources = @{ OnPremisesReady = 1 }"; Message = '*Resources.OnPremisesReady must be*' }
    ) {
        $path = Join-Path $script:folder 'c.psd1'
        Set-Content -LiteralPath $path -Value ("@{ Environment = 'X'; $Text }") -Encoding UTF8
        { Import-PcmConfiguration -Path $path -Root $script:folder } | Should -Throw $Message
    }
    It 'ships both native resource consents false and the packager resets any true value without touching comments (<Case>)' -ForEach @(
        @{ Case = 'block'; Text = "Resources = @{`r`n    OnPremisesReady = `$true`r`n    AllowUnprotectedDeletion = `$True # keep`r`n    Parallel = 100`r`n}" }
        @{ Case = 'inline'; Text = "Resources = @{ LicensingGroupId = ''; OnPremisesReady = `$true; AllowUnprotectedDeletion = `$true; Parallel = 100 }" }
        @{ Case = 'already false'; Text = "Resources = @{ OnPremisesReady = `$false; AllowUnprotectedDeletion = `$false }" }
    ) {
        $shipped = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1')
        $shipped.Resources.OnPremisesReady | Should -BeFalse
        $shipped.Resources.AllowUnprotectedDeletion | Should -BeFalse
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Repo 'tools\New-PraCloudPackage.ps1'), [ref]$null, [ref]$null)
        $function = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Reset-PraPackageConsent' }, $true) | Select-Object -First 1
        $function | Should -Not -BeNullOrEmpty
        . ([scriptblock]::Create($function.Extent.Text))
        $text = "@{ Environment = 'X'; # OnPremisesReady = `$true. AllowUnprotectedDeletion = `$true`r`n Scope = @{ Mode = 'Auto' }; $Text }"
        $result = Reset-PraPackageConsent -Text $text
        $path = Join-Path $script:folder 'packaged.psd1'
        Set-Content -LiteralPath $path -Value $result -Encoding UTF8
        $packaged = Import-PowerShellDataFile -LiteralPath $path
        $packaged.Resources.OnPremisesReady | Should -BeFalse
        $packaged.Resources.AllowUnprotectedDeletion | Should -BeFalse
        $result | Should -Match ([regex]::Escape('# OnPremisesReady = $true. AllowUnprotectedDeletion = $true'))
        if ($Case -ne 'already false') { $result | Should -Not -BeExactly $text }
    }
    It 'resets the packaged shipped configuration even if both consents were true' {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Repo 'tools\New-PraCloudPackage.ps1'), [ref]$null, [ref]$null)
        $function = $ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Reset-PraPackageConsent' }, $true) | Select-Object -First 1
        . ([scriptblock]::Create($function.Extent.Text))
        $text = [IO.File]::ReadAllText((Join-Path $script:Root 'config\PraCloudMailbox.config.psd1')).Replace('OnPremisesReady        = $false', 'OnPremisesReady        = $true').Replace('AllowUnprotectedDeletion = $false', 'AllowUnprotectedDeletion = $true')
        $text | Should -Match 'OnPremisesReady\s*=\s*\$true\s'
        $text | Should -Match 'AllowUnprotectedDeletion\s*=\s*\$true\s'
        $path = Join-Path $script:folder 'shipped-true.psd1'
        Set-Content -LiteralPath $path -Value (Reset-PraPackageConsent -Text $text) -Encoding UTF8
        $packaged = Import-PowerShellDataFile -LiteralPath $path
        $packaged.Resources.OnPremisesReady | Should -BeFalse
        $packaged.Resources.AllowUnprotectedDeletion | Should -BeFalse
    }
}

Describe 'SQLite store' {
    BeforeEach {
        $script:db = Join-Path $TestDrive ("store-{0}.db" -f [guid]::NewGuid().ToString('N'))
        $script:cn = Open-PcmStore -Path $script:db -Root $script:Root
    }
    AfterEach { Close-PcmStore $script:cn }

    It 'writes and reads a complete snapshot with Unicode values' {
        $id = New-PcmSnapshot $script:cn @{ RunId = 'r1'; Environment = 'TEST'; ToolVersion = '0.1.0'; ExchangeServer = 'EX1'; ScopeJson = '{}' }
        $rows = @([ordered]@{ object_guid = 'g1'; kind = 'Shared'; user_principal_name = 'compta@contoso.com'; display_name = 'Comptabilité é ü'; hidden_from_address_lists = $true; item_count = $null })
        Add-PcmRow $script:cn mailbox $id $rows | Should -Be 1
        Set-PcmSnapshotStatus $script:cn $id Complete @{ mailbox = 1 }
        $snap = Get-PcmSnapshot $script:cn
        $snap.id | Should -Be $id
        $snap.status | Should -Be 'Complete'
        $read = @(Read-PcmTable $script:cn mailbox $id)
        $read[0].display_name | Should -Be 'Comptabilité é ü'
        $read[0].hidden_from_address_lists | Should -Be 1
        $read[0].item_count | Should -BeNullOrEmpty
    }
    It 'never returns a Running or Failed snapshot as the last complete one' {
        $first = New-PcmSnapshot $script:cn @{ RunId = 'r1' }
        Set-PcmSnapshotStatus $script:cn $first Complete
        $null = New-PcmSnapshot $script:cn @{ RunId = 'r2' }
        $third = New-PcmSnapshot $script:cn @{ RunId = 'r3' }
        Set-PcmSnapshotStatus $script:cn $third Failed
        (Get-PcmSnapshot $script:cn).id | Should -Be $first
    }
    It 'rejects a value outside the allowed kinds' {
        $id = New-PcmSnapshot $script:cn @{ RunId = 'r1' }
        { Add-PcmRow $script:cn mailbox $id @([ordered]@{ object_guid = 'g'; kind = 'Room2' }) } | Should -Throw
        @(Read-PcmTable $script:cn mailbox $id).Count | Should -Be 0
    }
    It 'rejects an invalid column name (no SQL injection through the keys)' {
        $id = New-PcmSnapshot $script:cn @{ RunId = 'r1' }
        { Add-PcmRow $script:cn mailbox $id @([ordered]@{ 'object_guid) --' = 'g'; kind = 'User' }) } | Should -Throw '*Invalid column name*'
    }
    It 'keeps the last N complete snapshots and deletes their rows' {
        foreach ($n in 1..4) {
            $id = New-PcmSnapshot $script:cn @{ RunId = "r$n" }
            $null = Add-PcmRow $script:cn permission $id @([ordered]@{ mailbox_guid = 'g'; access_right = 'SendAs'; trustee = 'x'; resolved = $true })
            Set-PcmSnapshotStatus $script:cn $id Complete
        }
        Remove-PcmOldSnapshot $script:cn 2 | Should -Be 2
        @(Get-PcmSnapshotList $script:cn).Count | Should -Be 2
        [int](Invoke-PcmSql $script:cn 'SELECT count(*) FROM permission' -As Scalar) | Should -Be 2
    }
    It 'writes a consistent backup copy and refuses to overwrite it' {
        $id = New-PcmSnapshot $script:cn @{ RunId = 'r1' }
        Set-PcmSnapshotStatus $script:cn $id Complete
        $backup = Backup-PcmStore $script:cn (Join-Path $TestDrive 'bk') 'copy.db'
        Test-Path -LiteralPath $backup | Should -BeTrue
        { Backup-PcmStore $script:cn (Join-Path $TestDrive 'bk') 'copy.db' } | Should -Throw '*already exists*'
        $copy = Open-PcmStore -Path $backup -Root $script:Root -ReadOnly
        try { (Get-PcmSnapshot $copy).id | Should -Be $id } finally { Close-PcmStore $copy }
    }
    It 'opens read-only without creating anything and refuses writes' {
        { Open-PcmStore -Path (Join-Path $TestDrive 'missing.db') -Root $script:Root -ReadOnly } | Should -Throw '*Database not found*'
        $ro = Open-PcmStore -Path $script:db -Root $script:Root -ReadOnly
        try { { Invoke-PcmSql $ro "INSERT INTO meta (key, value) VALUES ('x', 'y')" } | Should -Throw } finally { Close-PcmStore $ro }
    }
    It 'refuses a database written by a newer schema' {
        $null = Invoke-PcmSql $script:cn "UPDATE meta SET value = '99' WHERE key = 'schema_version'"
        Close-PcmStore $script:cn
        { Open-PcmStore -Path $script:db -Root $script:Root -ReadOnly } | Should -Throw '*newer than this tool*'
        $script:cn = Open-PcmStore -Path (Join-Path $TestDrive ("other-{0}.db" -f [guid]::NewGuid().ToString('N'))) -Root $script:Root
    }
    It 'reads schema 1 without modification and migrates it only on a writable open' {
        $id = New-PcmSnapshot $script:cn @{ RunId = 'legacy' }
        $null = Add-PcmRow $script:cn mailbox $id @(@{ object_guid = 'legacy'; kind = 'Room' })
        Set-PcmSnapshotStatus $script:cn $id Complete
        $null = Invoke-PcmSql $script:cn 'ALTER TABLE mailbox DROP COLUMN resource_settings_json'
        $null = Invoke-PcmSql $script:cn "UPDATE meta SET value = '1' WHERE key = 'schema_version'"
        Close-PcmStore $script:cn
        $script:cn = $null
        $read = Open-PcmStore -Path $script:db -Root $script:Root -ReadOnly
        try {
            @(Read-PcmTable $read mailbox $id).Count | Should -Be 1
            (Invoke-PcmSql $read "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar) | Should -Be '1'
            @(Invoke-PcmSql $read 'PRAGMA table_info(mailbox)' -As Rows | ForEach-Object name) | Should -Not -Contain 'resource_settings_json'
        } finally { Close-PcmStore $read }
        $script:cn = Open-PcmStore -Path $script:db -Root $script:Root
        (Invoke-PcmSql $script:cn "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar) | Should -Be '3'
        @(Read-PcmTable $script:cn mailbox $id)[0].resource_settings_json | Should -BeNullOrEmpty
    }
    It 'round-trips resource settings, and the journal holds the runs of a batch (schema 2)' {
        $id = New-PcmSnapshot $script:cn @{ RunId = 'resource' }
        $settings = '{"SchemaVersion":1,"Capacity":12,"Calendar":{"AutomateProcessing":"AutoAccept"}}'
        $null = Add-PcmRow $script:cn mailbox $id @(@{ object_guid = 'room'; kind = 'Room'; resource_settings_json = $settings })
        (@(Read-PcmTable $script:cn mailbox $id)[0].resource_settings_json | ConvertFrom-Json).Capacity | Should -Be 12
        $journal = Open-PcmStore -Path (Join-Path $TestDrive 'resource-journal.db') -Root $script:Root -Kind Journal
        try { (Invoke-PcmSql $journal "SELECT value FROM meta WHERE key = 'schema_version'" -As Scalar) | Should -Be '2' }
        finally { Close-PcmStore $journal }
    }
}

Describe 'Collect: mailboxes' {
    It 'maps a mailbox to a row of the mailbox table' {
        $mailbox = New-FakeMailbox -Sam 'compta' -Type 'SharedMailbox' -OnBehalf @('CN=user01,OU=PRA,DC=contoso,DC=com')
        $record = ConvertTo-PcmMailboxRecord -Mailbox $mailbox -Statistics ([pscustomobject]@{ ItemCount = 42; TotalItemSize = '1.5 MB (1,572,864 bytes)' })
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
        { ConvertTo-PcmMailboxRecord -Mailbox (New-FakeMailbox -Sam 'x' -Type 'LinkedMailbox') } | Should -Throw '*Unsupported mailbox type*'
    }
    It 'applies -Scope to the configured mailbox types' {
        $scope = (New-TestConfig).Scope
        Get-PcmRecipientType -Scope $scope -Filter All | Should -Be @('UserMailbox','SharedMailbox')
        Get-PcmRecipientType -Scope $scope -Filter UsersOnly | Should -Be @('UserMailbox')
        Get-PcmRecipientType -Scope $scope -Filter SharedOnly | Should -Be @('SharedMailbox')
    }
    It 'reads the OU, excludes system and configured mailboxes, removes duplicates and sorts' {
        $config = New-TestConfig
        $config.Scope.ExcludeSamAccountNames = @('skip')
        $context = New-TestContext $config
        $dup = New-FakeMailbox -Sam 'bob'
        Mock -ModuleName PRA.Collect Get-Mailbox { @((New-FakeMailbox -Sam 'zoe'), $dup, $dup, (New-FakeMailbox -Sam 'HealthMailbox0123'), (New-FakeMailbox -Sam 'skip')) }
        $result = @(Get-PcmScopeMailbox -Context $context)
        ($result | ForEach-Object SamAccountName) | Should -Be @('bob','zoe')
        $context.Excluded.Count | Should -Be 2
        Should -Invoke -ModuleName PRA.Collect Get-Mailbox -Times 1 -ParameterFilter { $OrganizationalUnit -eq 'OU=PRA,DC=contoso,DC=com' -and $ResultSize -eq 'Unlimited' }
    }
    It 'refuses -Identity on a mailbox type outside the scope' {
        $config = New-TestConfig
        $config.Scope.IncludeShared = $false
        Mock -ModuleName PRA.Collect Get-Mailbox { New-FakeMailbox -Sam 'compta' -Type 'SharedMailbox' }
        { Get-PcmScopeMailbox -Context (New-TestContext $config) -Identity 'compta' } | Should -Throw '*not a mailbox type in scope*'
    }
    It 'expands a scope group recursively and survives a membership loop' {
        $config = New-TestConfig
        $config.Scope.Mode = 'Group'; $config.Scope.GroupDN = 'CN=G1'
        Mock -ModuleName PRA.Collect Get-Group {
            switch ($Identity) {
                'CN=G1' { [pscustomobject]@{ Members = @('CN=u1', 'CN=G2') } }
                'CN=G2' { [pscustomobject]@{ Members = @('CN=u2', 'CN=G1') } }
                default { throw 'not a group' }
            }
        }
        Mock -ModuleName PRA.Collect Get-Mailbox { New-FakeMailbox -Sam ($Identity -replace 'CN=', '') }
        $result = @(Get-PcmScopeMailbox -Context (New-TestContext $config))
        ($result | ForEach-Object SamAccountName) | Should -Be @('u1','u2')
    }
}

Describe 'Collect: native resource settings' {
    BeforeEach {
        $script:resource = New-FakeMailbox -Sam 'room' -Type 'RoomMailbox'
        $script:resource | Add-Member -NotePropertyName ResourceCapacity -NotePropertyValue 12
        $script:resource | Add-Member -NotePropertyName ResourceCustom -NotePropertyValue @('Video')
        $script:booking = [pscustomobject]@{
            AutomateProcessing = 'AutoAccept'; AllBookInPolicy = $false; AllRequestInPolicy = $true; AllRequestOutOfPolicy = $false
            ResourceDelegates = @('CN=delegate'); BookInPolicy = @(); RequestInPolicy = @('CN=bookers'); RequestOutOfPolicy = @()
            ScheduleOnlyDuringWorkHours = $true; AllowConflicts = $false; MaximumDurationInMinutes = 120; AdditionalResponse = 'Booking policy'
        }
        $script:hours = [pscustomobject]@{ WorkingHoursStartTime = '08:30:00'; WorkingHoursEndTime = '18:00:00'
            WorkDays = 'Monday, Tuesday, Wednesday, Thursday, Friday'; WorkingHoursTimeZone = [pscustomobject]@{ Id = 'Romance Standard Time' } }
        Mock Get-CalendarProcessing -ModuleName PRA.Collect { $script:booking }
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Collect { $script:hours }
        $script:resourceUser = [pscustomobject]@{ Office = 'Building A / room 101'; StreetAddress = '1 Test Street'; City = 'Paris'
            StateOrProvince = ''; PostalCode = '75001'; CountryOrRegion = 'France'; Phone = '+33100000000'; Notes = 'Booking reception' }
        Mock Get-User -ModuleName PRA.Collect { $script:resourceUser }
        Mock Get-DistributionGroup -ModuleName PRA.Collect { @() }
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect {
            [pscustomobject]@{ Resolved = $true; Smtp = $(if ($Trustee -eq 'CN=delegate') { 'delegate@contoso.com' } else { 'bookers@contoso.com' }) }
        }
    }
    It 'captures resource type settings, SMTP booking trustees and working hours' {
        $config = New-TestConfig
        $config.Exchange.DomainController = 'dc.contoso.com'
        $settings = Get-PcmResourceSettings -Context (New-TestContext $config) -Mailbox $script:resource
        $settings.Capacity | Should -Be 12
        $settings.ResourceCustom | Should -Be @('Video')
        $settings.Calendar.ResourceDelegates | Should -Be @('delegate@contoso.com')
        $settings.Calendar.RequestInPolicy | Should -Be @('bookers@contoso.com')
        $settings.Calendar.AllBookInPolicy | Should -BeFalse
        $settings.Calendar.MaximumDurationInMinutes | Should -Be 120
        $settings.CalendarConfiguration.WorkingHoursStartTime | Should -Be '08:30:00'
        $settings.CalendarConfiguration.WorkingHoursTimeZone | Should -Be 'Romance Standard Time'
        $settings.CalendarConfiguration.WorkingHoursDays.Count | Should -Be 5
        $settings.SchemaVersion | Should -Be 2
        $settings.Directory.Office | Should -Be 'Building A / room 101'
        $settings.Directory.City | Should -Be 'Paris'
        $settings.Directory.CountryOrRegion | Should -Be 'France'
        $settings.Directory.StateOrProvince | Should -Be ''
        $settings.RoomLists.Count | Should -Be 0
        $row = ConvertTo-PcmMailboxRecord -Mailbox $script:resource -ResourceSettings $settings
        ($row.resource_settings_json | ConvertFrom-Json).Calendar.ResourceDelegates | Should -Be @('delegate@contoso.com')
        $script:resource.GrantSendOnBehalfTo = @('CN=general-delegate')
        @((ConvertTo-PcmMailboxRecord -Mailbox $script:resource -ResourceSettings $settings).grant_send_on_behalf_json | ConvertFrom-Json | ForEach-Object { $_ }).Count | Should -Be 0
        Should -Invoke Get-CalendarProcessing -ModuleName PRA.Collect -Times 1 -ParameterFilter { $DomainController -eq 'dc.contoso.com' }
        Should -Invoke Get-MailboxCalendarConfiguration -ModuleName PRA.Collect -Times 1 -ParameterFilter { $DomainController -eq 'dc.contoso.com' }
        Should -Invoke Get-User -ModuleName PRA.Collect -Times 1 -ParameterFilter { $DomainController -eq 'dc.contoso.com' }
    }
    It 'does not invent settings for a legacy snapshot row' {
        (ConvertTo-PcmMailboxRecord -Mailbox $script:resource).resource_settings_json | Should -BeNullOrEmpty
        Should -Invoke Get-CalendarProcessing -ModuleName PRA.Collect -Times 0
    }
    It 'maps native WorkDays flags to the unchanged portable snapshot (<Days>)' -TestCases @(
        @{ Days = 'Weekdays' }, @{ Days = 'WeekendDays' }, @{ Days = 'AllDays' }) {
        param($Days)
        $script:hours.WorkDays = $Days
        $script:hours.WorkingHoursStartTime = [timespan]'08:00:00'
        $script:hours.WorkingHoursEndTime = [timespan]'18:00:00'
        $script:hours.WorkingHoursTimeZone = 'Romance Standard Time'
        $settings = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource
        $settings.CalendarConfiguration.WorkingHoursDays | Should -Be @($Days)
        $settings.CalendarConfiguration.Keys | Should -Not -Contain 'WorkDays'
        $row = ConvertTo-PcmMailboxRecord -Mailbox $script:resource -ResourceSettings $settings
        $stored = $row.resource_settings_json | ConvertFrom-Json
        $stored.SchemaVersion | Should -Be 2
        $stored.CalendarConfiguration.PSObject.Properties.Name | Should -Be @('WorkingHoursStartTime','WorkingHoursEndTime','WorkingHoursDays','WorkingHoursTimeZone')
        ($stored.CalendarConfiguration.WorkingHoursDays -is [array]) | Should -BeTrue
        $stored.CalendarConfiguration.WorkingHoursDays | Should -Be @($Days)
        $stored.CalendarConfiguration.WorkingHoursStartTime | Should -Be '08:00:00'
        $stored.CalendarConfiguration.WorkingHoursEndTime | Should -Be '18:00:00'
        $stored.CalendarConfiguration.WorkingHoursTimeZone | Should -Be 'Romance Standard Time'
        $script:hours.PSObject.Properties.Name | Should -Not -Contain 'WorkingHoursDays'
    }
    It 'preserves individual native WorkDays array values' {
        $script:hours.WorkDays = @('Monday','Wednesday','Friday')
        $settings = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource
        $settings.CalendarConfiguration.WorkingHoursDays | Should -Be @('Monday','Wednesday','Friday')
    }
    It 'refuses absent or unreadable native WorkDays without a snapshot-name fallback (<Fault>)' -TestCases @(
        @{ Fault = 'Missing' }, @{ Fault = 'Null' }, @{ Fault = 'Unreadable' },
        @{ Fault = 'Empty' }, @{ Fault = 'EmptyArray' }, @{ Fault = 'Unknown' }) {
        param($Fault)
        $script:hours | Add-Member -NotePropertyName WorkingHoursDays -NotePropertyValue 'Weekdays'
        switch ($Fault) {
            Missing { $script:hours.PSObject.Properties.Remove('WorkDays') }
            Null { $script:hours.WorkDays = $null }
            Unreadable {
                $script:hours.PSObject.Properties.Remove('WorkDays')
                $script:hours | Add-Member -MemberType ScriptProperty -Name WorkDays -Value { throw 'WorkDays getter unavailable' }
            }
            Empty { $script:hours.WorkDays = '' }
            EmptyArray { $script:hours.WorkDays = @() }
            Unknown { $script:hours.WorkDays = 'UnknownDays' }
        }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*WorkDays*'
    }
    It 'fails on unreadable native calendar configuration instead of manufacturing working hours' {
        Mock Get-MailboxCalendarConfiguration -ModuleName PRA.Collect { throw 'calendar configuration unavailable' }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*calendar configuration unavailable*'
    }
    It 'stops when calendar settings cannot be read instead of substituting auto-accept' {
        Mock Get-CalendarProcessing -ModuleName PRA.Collect { throw 'calendar unavailable' }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*calendar unavailable*'
    }
    It 'refuses unresolved booking trustees rather than opening the policy to everyone' {
        Mock Resolve-PcmTrustee -ModuleName PRA.Collect { [pscustomobject]@{ Resolved = $false; Smtp = '' } }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*has no resolved SMTP address*'
    }
    It 'refuses missing mandatory booking controls and working hours' {
        $script:booking.PSObject.Properties.Remove('AllBookInPolicy')
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*AllBookInPolicy missing*'
        $script:booking | Add-Member -NotePropertyName AllBookInPolicy -NotePropertyValue $false
        $script:hours.WorkingHoursTimeZone = ''
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*empty working-hours time zone*'
    }
    It 'captures invitation sender restrictions and moderation as SMTP controls, not mailbox rights' {
        $script:resource.AcceptMessagesOnlyFromSendersOrMembers = @('CN=bookers')
        $script:resource.RejectMessagesFromSendersOrMembers = @('CN=blocked')
        $script:resource.ModerationEnabled = $true
        $script:resource.ModeratedBy = @('CN=delegate')
        $script:resource.BypassModerationFromSendersOrMembers = @('CN=bookers')
        $settings = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource
        $settings.Delivery.RequireSenderAuthenticationEnabled | Should -BeTrue
        $settings.Delivery.ModerationEnabled | Should -BeTrue
        $settings.Delivery.AcceptMessagesOnlyFromSendersOrMembers | Should -Be @('bookers@contoso.com')
        $settings.Delivery.ModeratedBy | Should -Be @('delegate@contoso.com')
    }
    It 'refuses missing or unresolved invitation restrictions (<Fault>)' -TestCases @(
        @{ Fault = 'Missing' }, @{ Fault = 'Unresolved' }, @{ Fault = 'Boolean' }, @{ Fault = 'Moderation' }) {
        param($Fault)
        if ($Fault -eq 'Missing') { $script:resource.PSObject.Properties.Remove('RequireSenderAuthenticationEnabled') }
        elseif ($Fault -eq 'Boolean') { $script:resource.RequireSenderAuthenticationEnabled = $null }
        elseif ($Fault -eq 'Moderation') { $script:resource.ModerationEnabled = $true }
        else {
            $script:resource.RejectMessagesFromSendersOrMembers = @('CN=missing')
            Mock Resolve-PcmTrustee -ModuleName PRA.Collect { [pscustomobject]@{ Resolved = $false; Smtp = '' } }
        }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw '*Resource delivery restriction*'
    }
    It 'records explicitly unset recipient restrictions as empty arrays' {
        $script:resource.RejectMessagesFromSendersOrMembers = $null
        $settings = Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource
        ($settings.Delivery.RejectMessagesFromSendersOrMembers -is [array]) | Should -BeTrue
        $settings.Delivery.RejectMessagesFromSendersOrMembers.Count | Should -Be 0
    }
    It 'does not silently omit unreadable descriptive properties (<Fault>)' -TestCases @(
        @{ Fault = 'Read' }, @{ Fault = 'Missing' }) {
        param($Fault)
        if ($Fault -eq 'Read') { Mock Get-User -ModuleName PRA.Collect { throw 'directory unavailable' } }
        else { $script:resourceUser.PSObject.Properties.Remove('Office') }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw
    }
    It 'captures matching room lists once per Collect independently of distribution-group collection' {
        $group = New-FakeRoomList
        Mock Get-DistributionGroup -ModuleName PRA.Collect { $group }
        Mock Get-DistributionGroupMember -ModuleName PRA.Collect { [pscustomobject]@{ Guid = $script:resource.Guid } }
        $context = New-TestContext
        $settings = Get-PcmResourceSettings -Context $context -Mailbox $script:resource
        $second = Get-PcmResourceSettings -Context $context -Mailbox $script:resource
        $settings.RoomLists.Count | Should -Be 1
        $settings.RoomLists[0].PrimarySmtpAddress | Should -Be 'building-a@contoso.com'
        $second.RoomLists[0].HiddenFromAddressListsEnabled | Should -BeTrue
        $settings.RoomLists[0].EmailAddresses | Should -Contain 'X500:/o=Old/cn=building-a'
        $settings.RoomLists[0].LegacyExchangeDN | Should -Not -BeNullOrEmpty
        $settings.RoomLists[0].Delivery.RequireSenderAuthenticationEnabled | Should -BeTrue
        $settings.RoomLists[0].MemberJoinRestriction | Should -Be Closed
        $row = [pscustomobject](ConvertTo-PcmMailboxRecord -Mailbox $script:resource -ResourceSettings $settings)
        $portable = Read-PcmResourceSettings -Record $row
        $portable.RoomLists[0].EmailAddresses | Should -Contain 'X500:/o=Old/cn=building-a'
        $portable.RoomLists[0].Delivery.RequireSenderAuthenticationEnabled | Should -BeTrue
        Should -Invoke Get-DistributionGroup -ModuleName PRA.Collect -Times 1 -Exactly -ParameterFilter { $Filter -eq "RecipientTypeDetails -eq 'RoomList'" -and $ResultSize -eq 'Unlimited' }
        Should -Invoke Get-DistributionGroupMember -ModuleName PRA.Collect -Times 1 -Exactly
    }
    It 'does not assign equipment or unrelated rooms to a room list' {
        $script:resource.RecipientTypeDetails = 'EquipmentMailbox'
        (Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource).RoomLists.Count | Should -Be 0
        Should -Invoke Get-DistributionGroup -ModuleName PRA.Collect -Times 0 -Exactly
        $script:resource.RecipientTypeDetails = 'RoomMailbox'
        Mock Get-DistributionGroup -ModuleName PRA.Collect { New-FakeRoomList }
        Mock Get-DistributionGroupMember -ModuleName PRA.Collect { [pscustomobject]@{ Guid = 'another-room' } }
        (Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource).RoomLists.Count | Should -Be 0
    }
    It 'fails explicitly on room-list membership read or identity failures (<Fault>)' -TestCases @(
        @{ Fault = 'Read' }, @{ Fault = 'MissingGuid' }) {
        param($Fault)
        Mock Get-DistributionGroup -ModuleName PRA.Collect { New-FakeRoomList }
        if ($Fault -eq 'Read') { Mock Get-DistributionGroupMember -ModuleName PRA.Collect { throw 'members unavailable' } }
        else { Mock Get-DistributionGroupMember -ModuleName PRA.Collect { [pscustomobject]@{ Name = 'unreadable member' } } }
        { Get-PcmResourceSettings -Context (New-TestContext) -Mailbox $script:resource } | Should -Throw
    }
}

Describe 'Collect: target groups' {
    BeforeAll {
        $script:g1 = '11111111-aaaa-4aaa-8aaa-000000000001'; $script:g2 = '22222222-aaaa-4aaa-8aaa-000000000002'
        $script:dir = @{
            'CN=u1' = @{ Class = 'User'; Name = 'u1'; Guid = 'aaaaaaaa-0000-4000-8000-000000000001'; Upn = 'u1@contoso.com'; Smtp = 'u1@contoso.com' }
            'CN=u3' = @{ Class = 'User'; Name = 'u3'; Guid = 'aaaaaaaa-0000-4000-8000-000000000003'; Upn = 'u3@contoso.com'; Smtp = 'u3@contoso.com' }
            'CN=PC1' = @{ Class = 'Computer'; Name = 'PC1'; Guid = 'aaaaaaaa-0000-4000-8000-0000000000c1'; Upn = ''; Smtp = '' }
            'CN=Jane' = @{ Class = 'Contact'; Name = 'Jane'; Guid = 'aaaaaaaa-0000-4000-8000-0000000000c2'; Upn = ''; Smtp = 'jane@partner.com' }
            'CN=G2' = @{ Class = 'Group'; Name = 'G2'; Guid = $script:g2; Upn = ''; Smtp = '' }
        }
    }
    BeforeEach {
        Mock -ModuleName PRA.Collect Get-Group {
            switch -Regex ($Identity) {
                '^(CN=G1|PRA-Wave1)$' { [pscustomobject]@{ Name = 'PRA-Wave1'; SamAccountName = 'wave1'; DistinguishedName = 'CN=G1'; DisplayName = 'PRA Wave 1'; Guid = $script:g1; Members = @('CN=u1', 'CN=PC1', 'CN=Jane', 'CN=G2', 'CN=Gone') } }
                '^CN=G2$' { [pscustomobject]@{ Name = 'G2'; SamAccountName = 'g2'; DistinguishedName = 'CN=G2'; DisplayName = 'G2'; Guid = $script:g2; Members = @('CN=u3', 'CN=G1', 'CN=u1') } }
                '^CN=Broken$' { [pscustomobject]@{ Name = 'Broken'; SamAccountName = 'broken'; DistinguishedName = 'CN=Broken'; DisplayName = 'Broken'; Guid = [guid]::NewGuid().ToString(); Members = @('CN=G3') } }
                '^CN=G3$' { throw 'Access denied' }
                default { throw "Group $Identity not found" }
            }
        }
        Mock -ModuleName PRA.Collect Get-Recipient {
            if ($Identity -in @('CN=G1', 'wave1@contoso.com')) { [pscustomobject]@{ RecipientType = 'MailUniversalDistributionGroup'; PrimarySmtpAddress = 'wave1@contoso.com'; DistinguishedName = 'CN=G1' } }
            else { throw "Recipient $Identity not found" }
        }
        Mock -ModuleName PRA.Collect Get-PcmDirectoryObject {
            if ($DistinguishedName -eq 'CN=G1') { return [pscustomobject]@{ Class = 'Group'; Name = 'PRA-Wave1'; Sam = 'wave1'; Upn = ''; Smtp = ''; DisplayName = ''; Dn = 'CN=G1'; Guid = $script:g1 } }
            if ($DistinguishedName -eq 'CN=G3') { return [pscustomobject]@{ Class = 'Group'; Name = 'G3'; Sam = 'g3'; Upn = ''; Smtp = ''; DisplayName = ''; Dn = 'CN=G3'; Guid = [guid]::NewGuid().ToString() } }
            $entry = $script:dir[$DistinguishedName]
            if (-not $entry) { return $null }
            [pscustomobject]@{ Class = $entry.Class; Name = $entry.Name; Sam = ''; Upn = [string]$entry.Upn; Smtp = [string]$entry.Smtp; DisplayName = ''; Dn = $DistinguishedName; Guid = $entry.Guid }
        }
    }
    It 'captures a group with every member, class kept, nested groups walked, a loop and a duplicate ignored (the group is named <Written>)' -TestCases @(
        @{ Written = 'name in lower case'; Value = 'pra-wave1' }, @{ Written = 'name in upper case'; Value = 'PRA-WAVE1' }, @{ Written = 'distinguished name'; Value = 'CN=G1' }, @{ Written = 'SMTP address'; Value = 'WAVE1@contoso.com' }) {
        param($Written, $Value)
        $config = New-TestConfig
        $config.Collect.TargetGroups = @($Value, 'Pra-Wave1')
        $result = Get-PcmTargetGroup -Context (New-TestContext $config)
        @($result.Groups).Count | Should -Be 1
        $result.Groups[0].group_guid | Should -Be $script:g1
        $result.Groups[0].name | Should -Be 'PRA-Wave1'
        $result.Groups[0].sam_account_name | Should -Be 'wave1'
        $result.Groups[0].primary_smtp_address | Should -Be 'wave1@contoso.com'
        $result.Groups[0].member_count | Should -Be 5
        @($result.Warnings).Count | Should -Be 0
        $classes = @{}; foreach ($row in $result.Members) { $classes[$row.member_dn] = $row }
        $classes['CN=u1'].member_class | Should -Be 'User'
        $classes['CN=PC1'].member_class | Should -Be 'Computer'
        $classes['CN=Jane'].member_class | Should -Be 'Contact'
        $classes['CN=Gone'].member_class | Should -Be 'Unknown'
        $classes['CN=G2'].member_class | Should -Be 'Group'
        $classes['CN=u3'].via_group_guid | Should -Be $script:g2
        $classes['CN=u3'].depth | Should -Be 2
        $classes['CN=u1'].depth | Should -Be 1
        @($result.Members | Where-Object { $_.member_dn -eq 'CN=u1' }).Count | Should -Be 1
        @($result.Members | Where-Object { $_.member_class -eq 'Group' }).Count | Should -Be 1
        @($result.Members | Where-Object { $_.group_guid -ne $script:g1 }).Count | Should -Be 0
    }
    It 'captures the group of Scope.GroupDN too, and warns about a configured group that is not found without stopping' {
        $config = New-TestConfig
        $config.Scope.Mode = 'Group'; $config.Scope.GroupDN = 'cn=g1'
        $config.Collect.TargetGroups = @('Missing Group')
        $result = Get-PcmTargetGroup -Context (New-TestContext $config)
        @($result.Groups).Count | Should -Be 1
        $result.Groups[0].name | Should -Be 'PRA-Wave1'
        @($result.Warnings).Count | Should -Be 1
        $result.Warnings[0] | Should -Match 'Missing Group not found in Active Directory'
    }
    It 'does not capture a group whose members cannot all be read, and says so' {
        $config = New-TestConfig
        $config.Collect.TargetGroups = @('CN=Broken')
        $result = Get-PcmTargetGroup -Context (New-TestContext $config)
        @($result.Groups).Count | Should -Be 0
        @($result.Members).Count | Should -Be 0
        $result.Warnings[0] | Should -Match 'its members could not all be read'
    }
    It 'reports the members of the Group scope that are not mailboxes, with the reason' {
        $config = New-TestConfig
        $config.Scope.Mode = 'Group'; $config.Scope.GroupDN = 'CN=G1'
        Mock -ModuleName PRA.Collect Get-Mailbox { if ($Identity -eq 'CN=u1') { New-FakeMailbox -Sam 'u1' } elseif ($Identity -eq 'CN=u3') { New-FakeMailbox -Sam 'u3' -Type 'RoomMailbox' } }
        $context = New-TestContext $config
        $result = @(Get-PcmScopeMailbox -Context $context)
        ($result | ForEach-Object SamAccountName) | Should -Be @('u1')
        @($context.Excluded) | Should -Contain 'PC1: a computer, not a mailbox'
        @($context.Excluded) | Should -Contain 'Jane: a contact, not a mailbox'
        @($context.Excluded) | Should -Contain 'u3: RoomMailbox not in scope'
        @($context.Excluded | Where-Object { $_ -like 'CN=Gone: not a mailbox' }).Count | Should -Be 1
    }
    It 'writes the captured groups to the snapshot and reads them back whatever the letter case of the question' {
        $path = Join-Path $TestDrive 'groups.db'
        $cn = Open-PcmStore -Path $path -Root $script:Root
        try {
            $id = New-PcmSnapshot $cn @{ RunId = 'r'; Environment = 'T'; ToolVersion = 't'; ExchangeServer = 'EX'; ScopeJson = '{}' }
            $config = New-TestConfig
            $config.Collect.TargetGroups = @('CN=G1')
            $captured = Get-PcmTargetGroup -Context (New-TestContext $config)
            (Add-PcmRow $cn target_group $id @($captured.Groups)) | Should -Be 1
            (Add-PcmRow $cn target_member $id @($captured.Members)) | Should -Be 6
            Set-PcmSnapshotStatus $cn $id Complete @{}
            foreach ($value in @('PRA-WAVE1', 'pra wave 1', 'WAVE1', 'cn=g1', 'WAVE1@CONTOSO.COM', $script:g1.ToUpperInvariant(), ('{' + $script:g1 + '}'))) {
                $found = Find-PcmTargetGroup $cn $id $value
                $found.HasTables | Should -BeTrue
                @($found.Groups).Count | Should -Be 1 -Because $value
            }
            @((Find-PcmTargetGroup $cn $id 'nothing').Groups).Count | Should -Be 0
            @((Read-PcmTable $cn target_member $id) | Where-Object { $_.member_class -eq 'Computer' }).Count | Should -Be 1
        } finally { Close-PcmStore $cn }
    }
    It 'reads a snapshot written before the groups were captured: no table, no group' {
        $path = Join-Path $TestDrive 'old.db'
        $cn = Open-PcmStore -Path $path -Root $script:Root
        try {
            $id = New-PcmSnapshot $cn @{ RunId = 'r'; Environment = 'T'; ToolVersion = 't'; ExchangeServer = 'EX'; ScopeJson = '{}' }
            Set-PcmSnapshotStatus $cn $id Complete @{}
            $null = Invoke-PcmSql $cn 'DROP TABLE target_member'; $null = Invoke-PcmSql $cn 'DROP TABLE target_group'
        } finally { Close-PcmStore $cn }
        $read = Open-PcmStore -Path $path -Root $script:Root -ReadOnly
        try {
            $found = Find-PcmTargetGroup $read 1 'anything'
            $found.HasTables | Should -BeFalse
            @($found.Groups).Count | Should -Be 0
        } finally { Close-PcmStore $read }
    }
}

Describe 'Collect: shared mailbox permissions' {
    BeforeEach {
        $script:context = New-TestContext
        $script:shared = New-FakeMailbox -Sam 'compta' -Type 'SharedMailbox' -OnBehalf @('CN=user02,OU=PRA,DC=contoso,DC=com')
        Mock -ModuleName PRA.Collect Get-MailboxPermission {
            @([pscustomobject]@{ User = 'CONTOSO\user01'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; AccessRights = @('FullAccess','ReadPermission'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\Domain Admins'; AccessRights = @('FullAccess'); IsInherited = $true; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\Administrator'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\denied'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $true }
              [pscustomobject]@{ User = 'CONTOSO\GRP-Compta'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'S-1-5-21-1-2-3-4567'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false })
        }
        Mock -ModuleName PRA.Collect Get-ADPermission {
            @([pscustomobject]@{ User = 'CONTOSO\user01'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\user01'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'CONTOSO\ghost'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false }
              [pscustomobject]@{ User = 'NT AUTHORITY\SELF'; ExtendedRights = @('Send-As'); IsInherited = $false; Deny = $false })
        }
        Mock -ModuleName PRA.Collect Get-Recipient {
            switch -Wildcard ($Identity) {
                '*user01*' { [pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u01'; PrimarySmtpAddress = 'user01@contoso.com'; Name = 'user01'; DistinguishedName = 'CN=user01,OU=PRA,DC=contoso,DC=com' } }
                '*user02*' { [pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u02'; PrimarySmtpAddress = 'user02@contoso.com'; Name = 'user02'; DistinguishedName = 'CN=user02,OU=PRA,DC=contoso,DC=com' } }
                '*user03*' { [pscustomobject]@{ RecipientType = 'MailUser'; Guid = 'u03'; PrimarySmtpAddress = 'user03@contoso.com'; Name = 'user03'; DistinguishedName = 'CN=user03,OU=PRA,DC=contoso,DC=com' } }
                default { throw 'not found' }
            }
        }
        Mock -ModuleName PRA.Collect Get-User {
            switch -Wildcard ($Identity) {
                '*user0*' { [pscustomobject]@{ UserPrincipalName = (($Identity -replace '^CN=([^,]+),.*$', '$1') + '@contoso.com'); Guid = 'x'; Name = 'x'; DistinguishedName = $Identity } }
                default { throw 'not found' }
            }
        }
        Mock -ModuleName PRA.Collect Get-Group {
            switch -Wildcard ($Identity) {
                '*GRP-Compta*' { [pscustomobject]@{ Guid = 'g1'; Name = 'GRP-Compta'; DistinguishedName = 'CN=GRP-Compta,OU=PRA,DC=contoso,DC=com'; Members = @('CN=user03,OU=PRA,DC=contoso,DC=com') } }
                default { throw 'not a group' }
            }
        }
        Mock -ModuleName PRA.Collect Get-PcmDelegateListLink { @('CN=user01,OU=PRA,DC=contoso,DC=com') }
    }

    It 'keeps explicit grants only, without system, admin, inherited, denied, duplicate or self entries' {
        $result = Get-PcmSharedPermission -Context $script:context -Mailbox $script:shared
        $rights = @($result.Permissions | ForEach-Object { '{0} {1}' -f $_.access_right, $_.trustee })
        $rights | Should -Be @('FullAccess CONTOSO\user01', 'FullAccess CONTOSO\GRP-Compta', 'SendAs CONTOSO\user01', 'SendAs CONTOSO\ghost', 'SendOnBehalf CN=user02,OU=PRA,DC=contoso,DC=com')
    }
    It 'resolves trustees to UPN / SMTP and reads AutoMapping' {
        $result = Get-PcmSharedPermission -Context $script:context -Mailbox $script:shared
        $fullAccess = @($result.Permissions | Where-Object { $_.access_right -eq 'FullAccess' -and $_.trustee_kind -eq 'User' })[0]
        $fullAccess.trustee_upn | Should -Be 'user01@contoso.com'
        $fullAccess.trustee_smtp | Should -Be 'user01@contoso.com'
        $fullAccess.auto_mapping | Should -BeTrue
        @($result.Permissions | Where-Object { $_.access_right -eq 'SendOnBehalf' })[0].trustee_upn | Should -Be 'user02@contoso.com'
    }
    It 'expands a group trustee to its members' {
        $result = Get-PcmSharedPermission -Context $script:context -Mailbox $script:shared
        @($result.Members).Count | Should -Be 1
        $result.Members[0].group_guid | Should -Be 'g1'
        $result.Members[0].member_upn | Should -Be 'user03@contoso.com'
    }
    It 'reports a trustee that cannot be resolved' {
        $result = Get-PcmSharedPermission -Context $script:context -Mailbox $script:shared
        $ghost = @($result.Permissions | Where-Object { $_.trustee -eq 'CONTOSO\ghost' })[0]
        $ghost.resolved | Should -BeFalse
        @($result.Warnings | Where-Object { $_ -like '*ghost*' }).Count | Should -Be 1
    }
    It 'excludes trustees with wildcards from Collect.ExcludeTrustees' {
        $script:context.Config.Collect.ExcludeTrustees = @('GRP-*')
        Test-PcmExcludedTrustee -Context $script:context -Trustee 'CONTOSO\GRP-Compta' | Should -BeTrue
        Test-PcmExcludedTrustee -Context $script:context -Trustee 'CONTOSO\user01' | Should -BeFalse
        Test-PcmExcludedTrustee -Context $script:context -Trustee 'CONTOSO\Exchange Trusted Subsystem' | Should -BeTrue
    }
}

Describe 'Collect: end to end (synthetic Exchange, remote session)' {
    It 'collects native WorkDays into portable snapshots or stops before writing (<Days>, <Fault>)' -TestCases @(
        @{ Days = 'Weekdays'; Fault = '' }, @{ Days = 'WeekendDays'; Fault = '' }, @{ Days = 'AllDays'; Fault = '' },
        @{ Days = 'Weekdays'; Fault = 'Missing' }, @{ Days = 'Weekdays'; Fault = 'Null' },
        @{ Days = 'Weekdays'; Fault = 'Empty' }, @{ Days = 'Weekdays'; Fault = 'Read' }) {
        param($Days, $Fault)
        $folder = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory $folder
        $config = New-TestConfig
        $config.Scope.IncludeUsers = $false; $config.Scope.IncludeShared = $false
        $config.Scope.IncludeRoom = $true; $config.Scope.IncludeEquipment = $true; $config.Scope.ResourceMode = 'CloudResource'
        $config.Store.Path = Join-Path $folder 'db\pra.db'; $config.Store.BackupFolder = Join-Path $folder 'backup'
        $config.Logging.Folder = Join-Path $folder 'logs'; $config.Report.Folder = Join-Path $folder 'reports'
        $configPath = Join-Path $folder 'config.psd1'
        Set-Content -LiteralPath $configPath -Value (ConvertTo-Psd1 $config) -Encoding UTF8
        $mailboxes = @(foreach ($kind in @('Room','Equipment')) {
                $mailbox = New-FakeMailbox -Sam $kind.ToLowerInvariant() -Type "${kind}Mailbox"
                $mailbox | Add-Member -NotePropertyName ResourceCapacity -NotePropertyValue 12
                $mailbox | Add-Member -NotePropertyName ResourceCustom -NotePropertyValue @()
                $mailbox
            })
        $hours = [pscustomobject]@{ WorkDays = $Days; WorkingHoursStartTime = [timespan]'08:00:00'
            WorkingHoursEndTime = [timespan]'18:00:00'; WorkingHoursTimeZone = 'Romance Standard Time' }
        if ($Fault) { $hours | Add-Member -NotePropertyName WorkingHoursDays -NotePropertyValue 'Weekdays' }
        switch ($Fault) {
            Missing { $hours.PSObject.Properties.Remove('WorkDays') }
            Null { $hours.WorkDays = $null }
            Empty { $hours.WorkDays = '' }
        }
        $fixturePath = Join-Path $folder 'synthetic-exchange.xml'
        Export-Clixml -LiteralPath $fixturePath -InputObject @{
            Mailboxes = $mailboxes; Hours = $hours; Fault = $Fault
            Calendar = [pscustomobject]@{ AutomateProcessing = 'AutoAccept'; AllBookInPolicy = $false
                AllRequestInPolicy = $true; AllRequestOutOfPolicy = $false; ScheduleOnlyDuringWorkHours = $true
                ResourceDelegates = @(); BookInPolicy = @(); RequestInPolicy = @(); RequestOutOfPolicy = @() }
            User = [pscustomobject]@{ Office = ''; StreetAddress = ''; City = ''; StateOrProvince = ''
                PostalCode = ''; CountryOrRegion = ''; Phone = ''; Notes = '' }
        }
        $stubs = @'
function global:Get-ExchangeServer { @([pscustomobject]@{ Name = $env:COMPUTERNAME; AdminDisplayVersion = 'Version 15.2 (Build 2562.17)' }) }
function global:Get-Mailbox { param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController) $global:PraSyntheticCollect.Mailboxes }
function global:Get-CalendarProcessing { param($Identity, $DomainController) $global:PraSyntheticCollect.Calendar }
function global:Get-MailboxCalendarConfiguration {
    param($Identity, $DomainController)
    if ($global:PraSyntheticCollect.Fault -eq 'Read') { throw 'WorkDays calendar read unavailable' }
    $global:PraSyntheticCollect.Hours
}
function global:Get-User { param($Identity, $DomainController) $global:PraSyntheticCollect.User }
function global:Get-MailboxPermission { param($Identity, $DomainController) @() }
function global:Get-ADPermission { param($Identity, $DomainController) @() }
function global:Get-Recipient { param($Identity, $DomainController) throw 'Unexpected recipient read in synthetic Collect' }
function global:Get-DistributionGroup { param($Filter, $ResultSize, $DomainController) @() }
'@
        $runner = Join-Path $folder 'run.ps1'
        Set-Content -LiteralPath $runner -Encoding UTF8 -Value @"
`$ErrorActionPreference = 'Stop'
`$global:PcmSessionLog = '$(Join-Path $folder 'session.log')'
$($script:RemoteStubs)
`$global:PraSyntheticCollect = Import-Clixml -LiteralPath '$fixturePath'
$stubs
`$result = & '$(Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')' -Action Collect -Mode Apply -ConfigPath '$configPath' -Force -PassThru
`$result
exit `$result.ExitCode
"@
        $output = & $script:Pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner 2>&1
        if ($Fault) {
            $LASTEXITCODE | Should -Be 1 -Because ($output | Out-String)
            ($output | Out-String) | Should -Match 'Collect stopped before writing the snapshot'
            ($output | Out-String) | Should -Match 'WorkDays'
            Test-Path -LiteralPath $config.Store.Path | Should -BeFalse
            Test-Path -LiteralPath $config.Store.BackupFolder | Should -BeFalse
        } else {
            $LASTEXITCODE | Should -Be 0 -Because ($output | Out-String)
            $cn = Open-PcmStore -Path $config.Store.Path -Root $script:Root -ReadOnly
            try {
                $snap = Get-PcmSnapshot $cn
                $snap.status | Should -Be 'Complete'
                $rows = @(Read-PcmTable $cn mailbox $snap.id)
                $rows.Count | Should -Be 2
                $rows.kind | Sort-Object | Should -Be @('Equipment','Room')
                foreach ($row in $rows) {
                    $stored = $row.resource_settings_json | ConvertFrom-Json
                    $stored.SchemaVersion | Should -Be 2
                    $stored.CalendarConfiguration.PSObject.Properties.Name | Should -Not -Contain 'WorkDays'
                    $stored.CalendarConfiguration.WorkingHoursDays | Should -Be @($Days)
                    $stored.CalendarConfiguration.WorkingHoursStartTime | Should -Be '08:00:00'
                    $stored.CalendarConfiguration.WorkingHoursEndTime | Should -Be '18:00:00'
                    $stored.CalendarConfiguration.WorkingHoursTimeZone | Should -Be 'Romance Standard Time'
                }
            } finally { Close-PcmStore $cn }
        }
    }
    It 'writes a snapshot in Apply and nothing in Preview, through a remote session whose objects are deserialized' {
        $folder = Join-Path $TestDrive 'e2e'
        $null = New-Item -ItemType Directory $folder
        $config = New-TestConfig
        $config.Store.Path = (Join-Path $folder 'db\pra.db'); $config.Store.BackupFolder = (Join-Path $folder 'backup')
        $config.Logging.Folder = (Join-Path $folder 'logs'); $config.Report.Folder = (Join-Path $folder 'reports')
        $configPath = Join-Path $folder 'config.psd1'
        Set-Content -LiteralPath $configPath -Value (ConvertTo-Psd1 $config) -Encoding UTF8
        $stubs = @'
function global:Get-ExchangeServer { ConvertTo-PcmDeserialized @([pscustomobject]@{ Name = 'ex1'; AdminDisplayVersion = 'Version 15.2 (Build 1748.10)' }) }
function global:Get-Mailbox { param($Identity, $OrganizationalUnit, $ResultSize, $RecipientTypeDetails, $DomainController)
    $make = { param($sam, $type) [pscustomobject]@{ Guid = [guid]::NewGuid().ToString(); Identity = $sam; Name = $sam; DisplayName = $sam; Alias = $sam; SamAccountName = $sam
        UserPrincipalName = "$sam@contoso.com"; DistinguishedName = "CN=$sam,OU=PRA,DC=contoso,DC=com"; OrganizationalUnit = 'contoso.com/PRA'; PrimarySmtpAddress = "$sam@contoso.com"
        WindowsEmailAddress = "$sam@contoso.com"; LegacyExchangeDN = "/o=C/cn=$sam"; ExchangeGuid = [guid]::NewGuid().ToString(); ArchiveGuid = [guid]::Empty.ToString(); ArchiveStatus = 'None'
        RecipientTypeDetails = $type; HiddenFromAddressListsEnabled = $false; Database = 'DB01'; EmailAddresses = @("SMTP:$sam@contoso.com"); GrantSendOnBehalfTo = @() } }
    ConvertTo-PcmDeserialized @((& $make 'user01' 'UserMailbox'), (& $make 'compta' 'SharedMailbox')) }
function global:Get-MailboxPermission { param($Identity, $DomainController) ConvertTo-PcmDeserialized @([pscustomobject]@{ User = 'CONTOSO\user01'; AccessRights = @('FullAccess'); IsInherited = $false; Deny = $false }) }
function global:Get-ADPermission { param($Identity, $DomainController) @() }
function global:Get-Recipient { param($Identity, $DomainController) ConvertTo-PcmDeserialized ([pscustomobject]@{ RecipientType = 'UserMailbox'; Guid = 'u01'; PrimarySmtpAddress = 'user01@contoso.com'; Name = 'user01'; DistinguishedName = 'CN=user01,OU=PRA,DC=contoso,DC=com' }) }
function global:Get-User { param($Identity, $DomainController) ConvertTo-PcmDeserialized ([pscustomobject]@{ UserPrincipalName = 'user01@contoso.com' }) }
'@
        $runner = Join-Path $folder 'run.ps1'
        Set-Content -LiteralPath $runner -Encoding UTF8 -Value @"
`$ErrorActionPreference = 'Stop'
`$global:PcmSessionLog = '$(Join-Path $folder 'session.log')'
$($script:RemoteStubs)
$stubs
& '$(Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')' -Action Collect -Mode `$args[0] -ConfigPath '$configPath' -Force -PassThru
"@
        $preview = & $script:Pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner Preview 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($preview | Out-String)
        Test-Path -LiteralPath $config.Store.Path | Should -BeFalse
        $apply = & $script:Pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner Apply 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($apply | Out-String)
        ($apply | Out-String) | Should -Match 'remote session to ex1\.contoso\.com \(Kerberos\)'
        $session = @(Get-Content -LiteralPath (Join-Path $folder 'session.log'))
        $session[0] | Should -BeExactly 'Microsoft.Exchange|http://ex1.contoso.com/PowerShell/|Kerberos'
        $session | Should -Contain 'removed'
        $cn = Open-PcmStore -Path $config.Store.Path -Root $script:Root -ReadOnly
        try {
            $snap = Get-PcmSnapshot $cn
            $snap.status | Should -Be 'Complete'
            @(Read-PcmTable $cn mailbox $snap.id).Count | Should -Be 2
            @(Read-PcmTable $cn permission $snap.id | Where-Object { $_.trustee_upn -eq 'user01@contoso.com' }).Count | Should -Be 1
        } finally { Close-PcmStore $cn }
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
        $findings = Get-PcmReadiness -Config $script:config -Record $script:userRecord -State (New-State)
        Get-Codes $findings 'Error' | Should -BeNullOrEmpty
        Get-Codes $findings 'Ok' | Should -Contain 'EXO_MAILUSER'
        Get-Codes $findings | Should -Contain 'NO_TEAMS_STORAGE'
    }
    It 'announces the promotion of the Teams storage' {
        $state = New-State -Other @{ Locations = @([pscustomobject]@{ Type = 'ComponentShared'; Guid = 'c' }) }
        Get-Codes (Get-PcmReadiness -Config $script:config -Record $script:userRecord -State $state) | Should -Contain 'TEAMS_STORAGE'
    }
    It 'blocks <Case>' -ForEach @(
        @{ Case = 'a missing Entra ID object'; UserProps = @{}; StateProps = @{ User = $null }; Code = 'ENTRA_NOT_FOUND' }
        @{ Case = 'a cloud-only object'; UserProps = @{ onPremisesSyncEnabled = $null }; StateProps = @{}; Code = 'ENTRA_NOT_SYNCED' }
        @{ Case = 'a missing Exchange Online recipient'; UserProps = @{}; StateProps = @{ RecipientType = '' }; Code = 'EXO_NOT_FOUND' }
        @{ Case = 'an existing cloud mailbox'; UserProps = @{}; StateProps = @{ RecipientType = 'UserMailbox' }; Code = 'ALREADY_CLOUD_MAILBOX' }
        @{ Case = 'an Exchange plan already enabled'; UserProps = @{ assignedPlans = @([pscustomobject]@{ service = 'exchange'; capabilityStatus = 'Enabled' }) }; StateProps = @{}; Code = 'EXCHANGE_PLAN_PRESENT' }
    ) {
        $state = New-State -User $UserProps -Other $StateProps
        Get-Codes (Get-PcmReadiness -Config $script:config -Record $script:userRecord -State $state) 'Error' | Should -Contain $Code
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
        $findings = Get-PcmReadiness -Config $script:config -Record $script:userRecord -State $state
        Get-Codes $findings 'Warn' | Should -Contain $Code
        Get-Codes $findings 'Error' | Should -BeNullOrEmpty
    }
    It 'only informs about a hold on a shared mailbox (it is wanted before the identity is deleted)' {
        $state = New-State -Other @{ Holds = @('UniH1234'); ExchangeGuid = 'bbbbbbbb-0000-0000-0000-000000000002' }
        $findings = Get-PcmReadiness -Config $script:config -Record $script:sharedRecord -State $state
        @($findings | Where-Object Code -eq 'HOLD_PRESENT')[0].Level | Should -Be 'Info'
    }
    It 'refuses a <Kind> mailbox by default (Convert would turn it into a shared mailbox)' -ForEach @(@{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        $record = [pscustomobject]@{ kind = $Kind; object_guid = 'r'; exchange_guid = 'bbbbbbbb-0000-0000-0000-000000000002'; immutable_id = 'z'; user_principal_name = 'room@contoso.com' }
        $state = New-State -Other @{ ExchangeGuid = 'bbbbbbbb-0000-0000-0000-000000000002' }
        Get-Codes (Get-PcmReadiness -Config $script:config -Record $record -State $state) 'Error' | Should -Contain 'KIND_NOT_SUPPORTED'
        Get-Codes (Get-PcmReadiness -Config $script:config -Record $script:sharedRecord -State $state) 'Error' | Should -Not -Contain 'KIND_NOT_SUPPORTED'
    }
    It 'converts a <Kind> mailbox like a shared one when Scope.ConvertRooms is set' -ForEach @(@{ Kind = 'Room' }, @{ Kind = 'Equipment' }) {
        $config = New-TestConfig -Override @{ Scope = @{ Mode = 'OU'; SearchBase = 'OU=PRA,DC=contoso,DC=com'; GroupDN = ''; CsvPath = ''; IncludeUsers = $true; IncludeShared = $true; IncludeRoom = $true; IncludeEquipment = $true; ExcludeSamAccountNames = @(); ConvertRooms = $true } }
        $record = [pscustomobject]@{ kind = $Kind; object_guid = 'r'; exchange_guid = 'bbbbbbbb-0000-0000-0000-000000000002'; immutable_id = 'z'; user_principal_name = 'room@contoso.com' }
        $state = New-State -Other @{ ExchangeGuid = 'bbbbbbbb-0000-0000-0000-000000000002' }
        $findings = Get-PcmReadiness -Config $config -Record $record -State $state
        Get-Codes $findings 'Error' | Should -Not -Contain 'KIND_NOT_SUPPORTED'
        Get-Codes $findings | Should -Contain 'ROOM_CONVERT_ENABLED'
    }
    It 'detects the Exchange mailbox plans by ID: Foundation and other Exchange-backed plans give no mailbox' {
        $skus = @([pscustomobject]@{ skuId = 's'; servicePlans = @(
                    [pscustomobject]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; servicePlanId = 'exo' }
                    [pscustomobject]@{ servicePlanName = 'EXCHANGE_S_FOUNDATION'; servicePlanId = 'foundation' }
                    [pscustomobject]@{ servicePlanName = 'EXCHANGE_ANALYTICS'; servicePlanId = 'analytics' }) })
        Get-PcmMailboxPlanId -Skus $skus | Should -Be @('exo')
        Get-PcmSkuExchangePlan -Sku ([pscustomobject]@{ servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_FOUNDATION'; servicePlanId = 'f' }) }) | Should -BeNullOrEmpty
        $foundation = New-State -User @{ assignedPlans = @([pscustomobject]@{ service = 'exchange'; servicePlanId = 'foundation'; capabilityStatus = 'Enabled' }) }
        Get-Codes (Get-PcmReadiness -Config $script:config -Record $script:userRecord -State $foundation -Skus $skus) | Should -Not -Contain 'EXCHANGE_PLAN_PRESENT'
        $mailbox = New-State -User @{ assignedPlans = @([pscustomobject]@{ service = 'exchange'; servicePlanId = 'exo'; capabilityStatus = 'Enabled' }) }
        Get-Codes (Get-PcmReadiness -Config $script:config -Record $script:userRecord -State $mailbox -Skus $skus) 'Error' | Should -Contain 'EXCHANGE_PLAN_PRESENT'
    }
    It 'refuses an object without usageLocation only when no default country is configured' {
        $state = New-State -User @{ usageLocation = '' }
        $config = New-TestConfig
        Get-Codes (Get-PcmReadiness -Config $config -Record $script:userRecord -State $state) 'Error' | Should -Contain 'NO_USAGE_LOCATION'
        $config.Cloud.DefaultUsageLocation = 'FR'
        $findings = Get-PcmReadiness -Config $config -Record $script:userRecord -State $state
        Get-Codes $findings 'Info' | Should -Contain 'NO_USAGE_LOCATION'
        Get-Codes $findings 'Error' | Should -BeNullOrEmpty
    }
    It 'checks the Kiosk plan of each user in Kiosk mode' {
        $config = New-TestConfig
        $config.Licensing.Users.Mode = 'Kiosk'
        $skus = @([pscustomobject]@{ skuId = 't1'; skuPartNumber = 'Microsoft_Teams_Enterprise_New'; servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_DESKLESS'; servicePlanId = 'k' }) })
        $with = New-State -User @{ assignedLicenses = @([pscustomobject]@{ skuId = 't1'; disabledPlans = @('k') }) }
        Get-Codes (Get-PcmReadiness -Config $config -Record $script:userRecord -State $with -Skus $skus) 'Ok' | Should -Contain 'KIOSK_AVAILABLE'
        Get-Codes (Get-PcmReadiness -Config $config -Record $script:userRecord -State (New-State) -Skus $skus) 'Error' | Should -Contain 'NO_KIOSK_PLAN'
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
        $findings = Test-PcmTenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $script:group) -UserCount 1 -SharedCount 1
        @($findings | Where-Object Level -eq 'Error').Count | Should -Be 0
        @($findings | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Message | Should -Match '1 free unit'
    }
    It 'fails when the group has fewer free units than users' {
        $findings = Test-PcmTenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $script:group) -UserCount 3
        @($findings | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Level | Should -Be 'Error'
    }
    It 'fails when the group assigns no Exchange plan' {
        $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @('p-exo') }) }
        $findings = Test-PcmTenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $group) -UserCount 1
        @($findings | Where-Object { $_.Code -eq 'LICENCE_GROUP' -and $_.Level -eq 'Error' }).Count | Should -Be 1
    }
    It 'needs the group SOA permission when the licence group is synchronised from AD' {
        $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $true; assignedLicenses = $script:group.assignedLicenses }
        $findings = Test-PcmTenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant (New-Tenant $group $false) -UserCount 1
        @($findings | Where-Object Code -eq 'LICENCE_GROUP_SOA').Count | Should -Be 1
        @($findings | Where-Object { $_.Code -eq 'APP_PERMISSION' -and $_.Level -eq 'Error' -and $_.Message -like '*Group-OnPremisesSyncBehavior*' }).Count | Should -Be 1
    }
    It 'lists the missing application permissions' {
        $findings = Test-PcmTenantReadiness -Config (New-TestConfig) -Roles @('User.ReadWrite.All') -Tenant (New-Tenant $script:group) -UserCount 1
        @($findings | Where-Object { $_.Code -eq 'APP_PERMISSION' })[0].Message | Should -Match 'User-OnPremisesSyncBehavior'
    }
    It 'fails when no unit is free for the temporary licence of shared mailboxes' {
        $config = New-TestConfig
        $config.Licensing.Shared.SkuPartNumber = 'TEAMS'
        $findings = Test-PcmTenantReadiness -Config $config -Roles $script:roles -Tenant (New-Tenant $script:group) -UserCount 0 -SharedCount 2
        @($findings | Where-Object { $_.Code -eq 'SHARED_LICENCE' -and $_.Level -eq 'Error' }).Count | Should -Be 1
    }
    Context 'units kept free for the shared mailboxes' {
        BeforeAll {
            $script:plan = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_STANDARD'; servicePlanId = 'p-exo' })
            $script:bigSkus = @(
                [pscustomobject]@{ skuId = 'std'; skuPartNumber = 'EXCHANGESTANDARD'; consumedUnits = 0; prepaidUnits = [pscustomobject]@{ enabled = 50 }; servicePlans = $script:plan }
                [pscustomobject]@{ skuId = 'e5'; skuPartNumber = 'E5'; consumedUnits = 0; prepaidUnits = [pscustomobject]@{ enabled = 25 }; servicePlans = $script:plan }
            )
            $script:bigTenant = [pscustomobject]@{ Skus = $script:bigSkus; Organization = $null; LicenceGroup = $script:group; LicenceGroupCloudManaged = $true }
        }
        BeforeEach { $script:keepConfig = New-TestConfig; $script:keepConfig.Licensing.Shared.SkuPartNumber = 'EXCHANGESTANDARD'; $script:keepConfig.Licensing.Shared.Parallel = 100 }
        It 'shrinks the waves by the kept units' {
            $script:keepConfig.Licensing.Shared.KeepFree = 10
            $finding = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 0 -SharedCount 200 | Where-Object Code -eq 'SHARED_LICENCE')[0]
            $finding.Level | Should -Be 'Ok'
            $finding.Message | Should -BeLike 'EXCHANGESTANDARD: 50 free unit(s), 10 kept free (Licensing.Shared.KeepFree); the shared mailboxes are converted in waves of 40*'
        }
        It 'keeps the waves at Licensing.Shared.Parallel when enough units are free' {
            $script:keepConfig.Licensing.Shared.KeepFree = 10; $script:keepConfig.Licensing.Shared.Parallel = 5
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 0 -SharedCount 200 | Where-Object Code -eq 'SHARED_LICENCE')[0].Message | Should -Match 'waves of 5 '
        }
        It 'is an error when the kept units are all the free units (<Keep>)' -ForEach @(@{ Keep = 50 }, @{ Keep = 80 }) {
            $script:keepConfig.Licensing.Shared.KeepFree = $Keep
            $finding = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 0 -SharedCount 3 | Where-Object Code -eq 'SHARED_LICENCE')[0]
            $finding.Level | Should -Be 'Error'
            $finding.Message | Should -Match 'no unit is left'
            $finding.Message | Should -Match 'Licensing.Shared.KeepFree'
        }
        It 'does not change anything when KeepFree is 0 or absent' {
            $finding = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 0 -SharedCount 3 | Where-Object Code -eq 'SHARED_LICENCE')[0]
            $finding.Message | Should -BeLike 'EXCHANGESTANDARD: 50 free unit(s); the shared mailboxes are converted in waves of 50*'
            $script:keepConfig.Licensing.Shared.Remove('KeepFree')
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 0 -SharedCount 3 | Where-Object Code -eq 'SHARED_LICENCE')[0].Level | Should -Be 'Ok'
        }
        It 'spells the licence as the tenant does, whatever the letter case typed' {
            Get-PcmSkuSpelling -Name 'exchangestandard' -Tenant $script:bigTenant | Should -BeExactly 'EXCHANGESTANDARD'
            Get-PcmSkuSpelling -Name 'Unknown' -Tenant $script:bigTenant | Should -BeExactly 'Unknown'
            Get-PcmSkuSpelling -Name '' -Tenant $script:bigTenant | Should -BeExactly ''
            $script:keepConfig.Licensing.Shared.SkuPartNumber = 'exchangestandard'
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 0 -SharedCount 3 | Where-Object Code -eq 'SHARED_LICENCE')[0].Message | Should -BeLike 'EXCHANGESTANDARD: 50 free unit(s);*'
        }
        It 'counts the units left after the users of the same run take theirs (licence group, same licence): 25 free, 20 users, waves of 5' {
            $script:keepConfig.Licensing.Shared.SkuPartNumber = 'E5'
            $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @() }) }
            $tenant = [pscustomobject]@{ Skus = $script:bigSkus; Organization = $null; LicenceGroup = $group; LicenceGroupCloudManaged = $true }
            $finding = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $tenant -UserCount 20 -SharedCount 50 | Where-Object Code -eq 'SHARED_LICENCE')[0]
            $finding.Level | Should -Be 'Ok'
            $finding.Message | Should -BeLike 'E5: 25 free unit(s), 20 taken by the users of this run; the shared mailboxes are converted in waves of 5 *'
            $units = Get-PcmSharedUnitCount -Config $script:keepConfig -Tenant $tenant -UserCount 20
            $units.Wave | Should -Be 5
            $units.UserDraw | Should -Be 20
            (Get-PcmSharedUnitCount -Config $script:keepConfig -Tenant $tenant -UserCount 0).Wave | Should -Be 25
        }
        It 'counts the units left after the users and after KeepFree (Direct licence)' {
            $script:keepConfig.Licensing.Users.Mode = 'Direct'; $script:keepConfig.Licensing.Users.SkuPartNumber = 'E5'
            $script:keepConfig.Licensing.Shared.SkuPartNumber = 'E5'; $script:keepConfig.Licensing.Shared.KeepFree = 2
            (Get-PcmSharedUnitCount -Config $script:keepConfig -Tenant $script:bigTenant -UserCount 3).Wave | Should -Be 20
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 3 -SharedCount 50 | Where-Object Code -eq 'SHARED_LICENCE')[0].Message |
                Should -BeLike 'E5: 25 free unit(s), 3 taken by the users of this run, 2 kept free (Licensing.Shared.KeepFree); the shared mailboxes are converted in waves of 20 *'
        }
        It 'takes nothing from the shared licence when the users draw from another one' {
            $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @() }) }
            $tenant = [pscustomobject]@{ Skus = $script:bigSkus; Organization = $null; LicenceGroup = $group; LicenceGroupCloudManaged = $true }
            (Get-PcmSharedUnitCount -Config $script:keepConfig -Tenant $tenant -UserCount 20).Wave | Should -Be 50
            $finding = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $tenant -UserCount 20 -SharedCount 3 | Where-Object Code -eq 'SHARED_LICENCE')[0]
            $finding.Message | Should -BeLike 'EXCHANGESTANDARD: 50 free unit(s); the shared mailboxes are converted in waves of 50*'
        }
        It 'warns, and does not announce a wave, when the users of the run take every unit that can be used' {
            $script:keepConfig.Licensing.Shared.SkuPartNumber = 'E5'
            $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @() }) }
            $tenant = [pscustomobject]@{ Skus = $script:bigSkus; Organization = $null; LicenceGroup = $group; LicenceGroupCloudManaged = $true }
            $finding = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $tenant -UserCount 25 -SharedCount 2 | Where-Object Code -eq 'SHARED_LICENCE')[0]
            $finding.Level | Should -Be 'Warn'
            $finding.Message | Should -Match 'take every unit that can be used'
            $finding.Message | Should -Not -Match 'waves of'
            (Get-PcmSharedUnitCount -Config $script:keepConfig -Tenant $tenant -UserCount 25).Wave | Should -Be 0
        }
        It 'makes the users leave the kept units free when they draw from the same licence (group)' {
            $script:keepConfig.Licensing.Shared.SkuPartNumber = 'E5'
            $group = [pscustomobject]@{ displayName = 'G'; onPremisesSyncEnabled = $null; assignedLicenses = @([pscustomobject]@{ skuId = 'e5'; disabledPlans = @() }) }
            $tenant = [pscustomobject]@{ Skus = $script:bigSkus; Organization = $null; LicenceGroup = $group; LicenceGroupCloudManaged = $true }
            $script:keepConfig.Licensing.Shared.KeepFree = 5
            $ok = @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $tenant -UserCount 20 -SharedCount 1 | Where-Object Code -eq 'LICENCE_CAPACITY')[0]
            $ok.Level | Should -Be 'Ok'
            $ok.Message | Should -Match '5 of them kept free'
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $tenant -UserCount 21 -SharedCount 1 | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Level | Should -Be 'Error'
        }
        It 'makes the users leave the kept units free when they draw from the same licence (Direct)' {
            $script:keepConfig.Licensing.Users.Mode = 'Direct'; $script:keepConfig.Licensing.Users.SkuPartNumber = 'E5'
            $script:keepConfig.Licensing.Shared.SkuPartNumber = 'E5'; $script:keepConfig.Licensing.Shared.KeepFree = 5
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 20 -SharedCount 0 | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Level | Should -Be 'Ok'
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 21 -SharedCount 0 | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Level | Should -Be 'Error'
        }
        It 'does not hold units back from users who draw from another licence' {
            $script:keepConfig.Licensing.Users.Mode = 'Direct'; $script:keepConfig.Licensing.Users.SkuPartNumber = 'E5'
            $script:keepConfig.Licensing.Shared.KeepFree = 30
            @(Test-PcmTenantReadiness -Config $script:keepConfig -Roles $script:roles -Tenant $script:bigTenant -UserCount 25 -SharedCount 0 | Where-Object Code -eq 'LICENCE_CAPACITY')[0].Level | Should -Be 'Ok'
        }
    }
    It 'warns about retention policies applied to every mailbox (they would block the user Recover)' {
        $tenant = New-Tenant $script:group
        @(Test-PcmTenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant $tenant -UserCount 1 | Where-Object Code -eq 'ORG_HOLD').Count | Should -Be 0
        $tenant | Add-Member -NotePropertyName OrgHolds -NotePropertyValue @('mbxaaaa:2')
        $finding = @(Test-PcmTenantReadiness -Config (New-TestConfig) -Roles $script:roles -Tenant $tenant -UserCount 1 | Where-Object Code -eq 'ORG_HOLD')[0]
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
    BeforeEach { Reset-PraFakeGraphTransport; $script:sent = [System.Collections.Generic.List[object]]::new() }
    It 'enables a plan inside the existing direct assignment and keeps the other disabled plans' {
        Mock Get-PcmUserLicence -ModuleName PRA.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = $null; disabledPlans = @('kiosk', 'yammer') }) }
        Mock Set-PcmLicence -ModuleName PRA.Cloud { $script:sent.Add(@{ Sku = $SkuId; Operation = $Operation; Disabled = @(Get-Variable -Name DisabledPlans -ValueOnly -ErrorAction SilentlyContinue | ForEach-Object { $_ }) }) }
        Enable-PcmLicencePlan -UserId 'u' -SkuId 'teams' -PlanIds @('kiosk') -DisabledPlans @('ignored') | Should -Be 'Enabled'
        $script:sent[0].Disabled | Should -Be @('yammer')
    }
    It 'adds a direct assignment next to a group one, and changes nothing when the plan is already enabled' {
        Mock Get-PcmUserLicence -ModuleName PRA.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = 'g'; disabledPlans = @('kiosk') }) }
        Mock Set-PcmLicence -ModuleName PRA.Cloud { $script:sent.Add(@{ Sku = $SkuId; Operation = $Operation; Disabled = @(Get-Variable -Name DisabledPlans -ValueOnly -ErrorAction SilentlyContinue | ForEach-Object { $_ }) }) }
        Enable-PcmLicencePlan -UserId 'u' -SkuId 'teams' -PlanIds @('kiosk') -DisabledPlans @('x') | Should -Be 'Assigned'
        $script:sent[0].Disabled | Should -Be @('x')
        Mock Get-PcmUserLicence -ModuleName PRA.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = $null; disabledPlans = @('yammer') }) }
        Enable-PcmLicencePlan -UserId 'u' -SkuId 'teams' -PlanIds @('kiosk') | Should -Be 'Present'
        $script:sent.Count | Should -Be 1
    }
    It 'restores the original disabled plans of a licence the user had, and removes one he did not have' {
        Mock Get-PcmUserLicence -ModuleName PRA.Cloud { New-Licence @([pscustomobject]@{ skuId = 'teams'; assignedByGroup = $null; disabledPlans = @('yammer') }) }
        Mock Set-PcmLicence -ModuleName PRA.Cloud { $script:sent.Add(@{ Sku = $SkuId; Operation = $Operation; Disabled = @(Get-Variable -Name DisabledPlans -ValueOnly -ErrorAction SilentlyContinue | ForEach-Object { $_ }) }) }
        Restore-PcmLicence -UserId 'u' -SkuId 'teams' -Original @([pscustomobject]@{ skuId = 'teams'; disabledPlans = @('yammer', 'kiosk') }) | Should -Be 'Restored'
        $script:sent[0].Operation | Should -Be 'Add'
        $script:sent[0].Disabled | Should -Be @('kiosk', 'yammer')
        Restore-PcmLicence -UserId 'u' -SkuId 'teams' -Original @() | Should -Be 'Removed'
        $script:sent[1].Operation | Should -Be 'Remove'
        Restore-PcmLicence -UserId 'u' -SkuId 'other' -Original @() | Should -Be 'Absent'
        Restore-PcmLicence -UserId 'u' -SkuId 'teams' -Original @([pscustomobject]@{ skuId = 'teams'; disabledPlans = @('yammer') }) | Should -Be 'Unchanged'
        $script:sent.Count | Should -Be 2
    }
}

Describe 'Entra Connect operator script' {
    It 'accepts a script that ends without exit (StrictMode) and refuses a non-zero exit code' {
        $folder = Join-Path $TestDrive 'ec'; $null = New-Item -ItemType Directory $folder
        $ok = Join-Path $folder 'ok.ps1'; Set-Content -LiteralPath $ok -Value "param(`$Operation) 'scheduler paused'" -Encoding UTF8
        $bad = Join-Path $folder 'bad.ps1'; Set-Content -LiteralPath $bad -Value "param(`$Operation) 'failed'; exit 3" -Encoding UTF8
        $context = @{ Config = @{ EntraConnect = @{ Mode = 'Script'; ScriptPath = $bad; Server = 'aadc' } } }
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $true }
        { Invoke-PcmEntraConnect -Context $context -Operation Pause } | Should -Throw '*exit code 3*'
        $context.Config.EntraConnect.ScriptPath = $ok
        & { Set-StrictMode -Version Latest; Invoke-PcmEntraConnect -Context $context -Operation Pause } | Should -Be 'scheduler paused'
    }
}

Describe 'Entra Connect: a connector that is busy' {
    BeforeAll {
        function global:Get-ADSyncScheduler { [pscustomobject]@{ SyncCycleEnabled = $true; SyncCycleInProgress = $false } }
        function global:Set-ADSyncScheduler { }
        function global:Add-EcSleep { param([int]$Seconds) $global:EcSleeps.Add($Seconds) }
        function global:Start-ADSyncSyncCycle {
            [CmdletBinding()] param([string]$PolicyType)
            $global:EcStarts++
            if ($global:EcStarts -le $global:EcBusyAnswers) { throw [System.InvalidOperationException]::new($global:EcRawText) }
        }
    }
    AfterAll { foreach ($name in 'Get-ADSyncScheduler', 'Set-ADSyncScheduler', 'Start-ADSyncSyncCycle', 'Add-EcSleep') { Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue } }
    BeforeEach {
        $global:EcStarts = 0; $global:EcBusyAnswers = 0; $global:EcSleeps = [System.Collections.Generic.List[int]]::new()
        $global:EcRawText = "Retry operation: the connector is busy.`n   at Microsoft.IdentityManagement.Foo.Run() in C:\x.cs:line 12"
        # The commands of the remote block run here: the values the block takes from the caller ($using:) are put in the text, the module import is left out and the waits are noted.
        Mock Invoke-Command -ModuleName PRA.Cloud {
            $text = $ScriptBlock.ToString().Replace('$using:Operation', "'Delta'").Replace('$using:busyAttempts', [string](& (Get-Module PRA.Cloud) { $script:EntraBusyAttempts })).Replace('$using:busySpacing', [string](& (Get-Module PRA.Cloud) { $script:EntraBusySpacingSeconds }))
            $text = $text.Replace('Import-Module ADSync -ErrorAction Stop', '').Replace('Start-Sleep -Seconds', 'Add-EcSleep')
            & ([scriptblock]::Create($text))
        }
        Set-PcmWriteApproval -Approval { param($Target, $Operation) $true }
        $script:ecContext = @{ BatchId = '9f8e7d6c'; ConvertBatchId = 'ab12cd34'; Config = @{ EntraConnect = @{ Mode = 'Remoting'; Server = 'aadc' } } }
    }
    AfterEach { Set-PcmWriteApproval -Approval $null }
    It 'starts the cycle once when the connector is free' {
        Invoke-PcmEntraConnect -Context $script:ecContext -Operation Delta | Should -Match 'SyncCycleEnabled=True'
        $global:EcStarts | Should -Be 1
    }
    It 'tries again, one minute apart, when the connector is busy, and goes on when it answers' {
        $global:EcBusyAnswers = 2
        Invoke-PcmEntraConnect -Context $script:ecContext -Operation Delta | Should -Match 'SyncCycleEnabled=True'
        $global:EcStarts | Should -Be 3
        @($global:EcSleeps | Where-Object { $_ -eq 60 }).Count | Should -Be 2
    }
    It 'stops after three attempts with a plain message that names the Convert batch to run again: no .NET text, no stack line' {
        $global:EcBusyAnswers = 99
        $failure = try { Invoke-PcmEntraConnect -Context $script:ecContext -Operation Delta } catch { $_ }
        $failure.Exception.Message | Should -BeExactly 'Entra Connect is busy with another synchronization cycle: wait for it to finish, then run Recover -Batch ab12cd34 again.'
        $global:EcStarts | Should -Be 3
        $failure.Exception.Message | Should -Not -Match 'Microsoft\.IdentityManagement|\s+at |Retry operation'
    }
    It 'names a placeholder instead of the batch when the run has no Convert batch yet' {
        $global:EcBusyAnswers = 99; $script:ecContext.ConvertBatchId = ''
        { Invoke-PcmEntraConnect -Context $script:ecContext -Operation Delta } | Should -Throw '*run Recover -Batch <id> again.'
    }
    It 'does not try again for another error' {
        $global:EcBusyAnswers = 99; $global:EcRawText = 'Access is denied.'
        { Invoke-PcmEntraConnect -Context $script:ecContext -Operation Delta } | Should -Throw '*Access is denied*'
        $global:EcStarts | Should -Be 1
    }
    It 'keeps the number of attempts and the spacing in two settings of the module: 3 and 60 seconds' {
        & (Get-Module PRA.Cloud) { $script:EntraBusyAttempts } | Should -Be 3
        & (Get-Module PRA.Cloud) { $script:EntraBusySpacingSeconds } | Should -Be 60
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
        $finding = @(Get-PcmTrusteeFinding -Record $record -Permissions $permissions -Members $members -Lookup { param($id) $id -eq 'u1@c.com' })[0]
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
        (Get-PcmOutcome $context).ExitCode | Should -Be 0
        [void]$context.Rows.Add([ordered]@{ Identity = 'b'; FinalStatus = 'Pending' })
        (Get-PcmOutcome $context).ExitCode | Should -Be 2
        [void]$context.Rows.Add([ordered]@{ Identity = 'c'; FinalStatus = 'Error' })
        (Get-PcmOutcome $context).ExitCode | Should -Be 1
    }
}

Describe 'Rough duration estimate' {
    It 'is empty when there is nothing to convert or roll back' {
        (Get-PcmDurationEstimate -UserWaves 0 -SharedRounds 0).Text | Should -BeNullOrEmpty
    }
    It 'adds 2 min per user wave and 3 min per shared pipeline round by default' {
        $estimate = Get-PcmDurationEstimate -UserWaves 10 -SharedRounds 4
        $estimate.Seconds | Should -Be (10 * 120 + 4 * 180)
        $estimate.Text | Should -Match 'min|h'
    }
    It 'accepts other typical durations' {
        (Get-PcmDurationEstimate -UserWaves 2 -SharedRounds 0 -UserWaveSeconds 30).Seconds | Should -Be 60
    }
}

Describe 'Entry script' {
    It 'uses no $PSScriptRoot in the parameter defaults (the default configuration path is computed in the script body)' {
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
    It 'starts with pwsh -File and stops in plain words when Exchange.Server is empty' {
        $folder = Join-Path $TestDrive 'file-mode'
        $null = New-Item -ItemType Directory $folder
        $config = New-TestConfig
        $config.Exchange.Server = ''
        $config.Logging.Folder = (Join-Path $folder 'logs'); $config.Report.Folder = (Join-Path $folder 'reports'); $config.Store.Path = (Join-Path $folder 'db.db')
        $configPath = Join-Path $folder 'c.psd1'
        Set-Content -LiteralPath $configPath -Value (ConvertTo-Psd1 $config) -Encoding UTF8
        $output = & $script:Pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File (Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1') -Action Collect -ConfigPath $configPath 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 1
        $output | Should -Not -Match 'Join-Path'
        $output | Should -Match 'Exchange\.Server is empty\. Collect reads Exchange through a remote session'
        $output | Should -Match 'RESULT: FAIL'
    }
    It 'stops in plain words when the remote session to Exchange cannot be opened, and leaves no snapshot' {
        $folder = Join-Path $TestDrive 'session-refused'
        $null = New-Item -ItemType Directory $folder
        $config = New-TestConfig
        $config.Logging.Folder = (Join-Path $folder 'logs'); $config.Report.Folder = (Join-Path $folder 'reports'); $config.Store.Path = (Join-Path $folder 'db.db')
        $configPath = Join-Path $folder 'c.psd1'
        Set-Content -LiteralPath $configPath -Value (ConvertTo-Psd1 $config) -Encoding UTF8
        $runner = Join-Path $folder 'run.ps1'
        Set-Content -LiteralPath $runner -Encoding UTF8 -Value @"
`$ErrorActionPreference = 'Stop'
function global:New-PSSession { [CmdletBinding()] param(`$ConfigurationName, `$ConnectionUri, `$Authentication) throw 'Access is denied.' }
`$result = & '$(Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')' -Action Collect -Mode Apply -ConfigPath '$configPath' -Force -PassThru
exit `$result.ExitCode
"@
        $output = & $script:Pwsh -NoProfile -NonInteractive -ExecutionPolicy Bypass -File $runner 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 1
        $output | Should -Match 'Cannot open the remote session to Exchange on ex1\.contoso\.com: Access is denied\. Collect needs a domain-joined computer'
        $output | Should -Match 'http://ex1\.contoso\.com/PowerShell/ with Kerberos'
        Test-Path -LiteralPath (Join-Path $folder 'db.db') | Should -BeFalse
    }
}

Describe 'Journal of the cloud actions' {
    BeforeEach {
        $script:jpath = Join-Path $TestDrive ("journal-{0}.db" -f [guid]::NewGuid().ToString('N'))
        $script:j = Open-PcmStore -Path $script:jpath -Root $script:Root -Kind Journal
    }
    AfterEach { Close-PcmStore $script:j }

    It 'creates a batch, records items and events, and finds it by prefix' {
        $id = New-PcmBatch $script:j Convert @{ SnapshotId = 3; Environment = 'TEST'; ToolVersion = '0.2.0' }
        $id | Should -Match '^[0-9a-f]{8}$'
        Set-PcmBatchItem $script:j $id 'g1' @{ kind = 'User'; identity = 'u1@c.com'; entra_id = 'e1'; status = 'Planned'; original_json = '{"isCloudManaged":false}' }
        Set-PcmBatchItem $script:j $id 'g1' @{ step = 'SoaCloud'; status = 'Running' }
        Add-PcmBatchEvent $script:j $id -ObjectGuid 'g1' -Step 'SoaCloud' -Detail 'isCloudManaged = true'
        $item = @(Get-PcmBatchItem $script:j $id)[0]
        $item.step | Should -Be 'SoaCloud'
        $item.status | Should -Be 'Running'
        $item.original_json | Should -Be '{"isCloudManaged":false}'
        (Get-PcmBatch $script:j -Id $id.Substring(0, 5)).id | Should -Be $id
        (Get-PcmBatch $script:j -Action Convert).snapshot_id | Should -Be 3
        [int](Invoke-PcmSql $script:j 'SELECT count(*) FROM batch_event' -As Scalar) | Should -Be 1
    }
    It 'refuses an invalid batch ID, a journal column outside the list, and a new item without kind and status' {
        { Get-PcmBatch $script:j -Id "x'; DROP" } | Should -Throw '*Invalid batch ID*'
        $id = New-PcmBatch $script:j Recover @{ ConvertBatch = 'abcd1234' }
        { Set-PcmBatchItem $script:j $id 'g' @{ kind = 'User'; status = 'Planned'; 'step = 1 --' = 'x' } } | Should -Throw '*Invalid journal column*'
        { Set-PcmBatchItem $script:j $id 'g' @{ step = 'x' } } | Should -Throw '*needs kind and status*'
    }
    It 'records a tenant change and its restoration' {
        $id = New-PcmBatch $script:j Convert
        Set-PcmTenantChange $script:j $id 'GroupSoa' 'grp1' -OriginalJson '{"isCloudManaged":false}'
        @(Get-PcmTenantChange $script:j $id)[0].restored | Should -Be 0
        Set-PcmTenantChange $script:j $id 'GroupSoa' 'grp1' -Restored
        @(Get-PcmTenantChange $script:j $id)[0].restored | Should -Be 1
        @(Get-PcmTenantChange $script:j $id)[0].original_json | Should -Be '{"isCloudManaged":false}'
    }
    It 'updates a room-list intent explicitly without changing legacy tenant-change originals' {
        $id = New-PcmBatch $script:j Convert
        Set-PcmTenantChange $script:j $id 'NativeRoomList' 'list1' -OriginalJson '{"creationAttempted":false}'
        Set-PcmTenantChange $script:j $id 'NativeRoomList' 'list1' -OriginalJson '{"creationAttempted":true}' -UpdateOriginal
        @(Get-PcmTenantChange $script:j $id)[0].original_json | Should -Be '{"creationAttempted":true}'
        Set-PcmTenantChange $script:j $id 'NativeRoomList' 'list1' -OriginalJson '{"unexpected":true}' -Restored
        @(Get-PcmTenantChange $script:j $id)[0].original_json | Should -Be '{"creationAttempted":true}'
        @(Get-PcmTenantChange $script:j $id).Count | Should -Be 1
    }
    It 'never opens a snapshot database as a journal, nor a journal as a snapshot database' {
        $snapshotPath = Join-Path $TestDrive ("snap-{0}.db" -f [guid]::NewGuid().ToString('N'))
        Close-PcmStore (Open-PcmStore -Path $snapshotPath -Root $script:Root)
        { Open-PcmStore -Path $snapshotPath -Root $script:Root -Kind Journal } | Should -Throw '*snapshot database, not a journal*'
        { Open-PcmStore -Path $script:jpath -Root $script:Root -ReadOnly } | Should -Throw '*journal database, not a snapshot*'
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
        $grants = @(Get-PcmSharedGrant -Record $record -Permissions $permissions -Members $members)
        ($grants | ForEach-Object { '{0} {1} {2}' -f $_.Right, $_.Trustee, $_.AutoMapping }) | Should -Be @('FullAccess u1@c.com False', 'FullAccess u3@c.com True', 'SendOnBehalf u2@c.com ')
    }
    It 'computes a Kiosk assignment that only adds the Kiosk plan to the existing licence' {
        $skus = @([pscustomobject]@{ skuId = 't1'; skuPartNumber = 'TEAMS'; servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_DESKLESS'; servicePlanId = 'k' }, [pscustomobject]@{ servicePlanName = 'X'; servicePlanId = 'x' }) })
        $licence = [pscustomobject]@{ licenseAssignmentStates = @([pscustomobject]@{ skuId = 't1'; assignedByGroup = 'grp'; disabledPlans = @('k', 'x') }) }
        $kiosk = Get-PcmKioskAssignment -Licence $licence -Skus $skus
        $kiosk.SkuId | Should -Be 't1'
        $kiosk.DisabledPlans | Should -Be @('x')
        $kiosk.DirectWithKiosk | Should -BeFalse
        $licence.licenseAssignmentStates += [pscustomobject]@{ skuId = 't1'; assignedByGroup = $null; disabledPlans = @('x') }
        (Get-PcmKioskAssignment -Licence $licence -Skus $skus).DirectWithKiosk | Should -BeTrue
        Get-PcmKioskAssignment -Licence ([pscustomobject]@{ licenseAssignmentStates = @() }) -Skus $skus | Should -BeNullOrEmpty
    }
    It 'leaves only the Exchange plan enabled for a licence given for the mailbox' {
        $sku = [pscustomobject]@{ servicePlans = @([pscustomobject]@{ servicePlanName = 'EXCHANGE_S_ENTERPRISE'; servicePlanId = 'e'; appliesTo = 'User' }, [pscustomobject]@{ servicePlanName = 'TEAMS1'; servicePlanId = 't'; appliesTo = 'User' }, [pscustomobject]@{ servicePlanName = 'COMPANY'; servicePlanId = 'c'; appliesTo = 'Company' }) }
        Get-PcmDisabledPlanForExchangeOnly -Sku $sku | Should -Be @('t')
    }
    It 'waits for a condition and reports the elapsed time' {
        $script:n = 0
        $result = Wait-PcmCondition -Test { $script:n++; $script:n -ge 2 } -TimeoutSeconds 10 -IntervalSeconds 1
        $result.Ok | Should -BeTrue
        (Wait-PcmCondition -Test { $false } -TimeoutSeconds 1 -IntervalSeconds 1).Ok | Should -BeFalse
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
    BeforeEach { Reset-PraFakeGraphTransport }
    AfterAll { Remove-Item -Path 'function:global:Invoke-MgGraphRequest' -ErrorAction SilentlyContinue }
    It 'puts the Graph error code and message in the exception (not only "BadRequest")' {
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { throw (New-GraphError 'Request_BadRequest' 'Invalid usage location.') }
        { Invoke-PcmGraph -Method POST -Uri 'v1.0/users/u1/assignLicense?x=1' -Body @{} } | Should -Throw '*POST v1.0/users/u1/assignLicense failed*Request_BadRequest: Invalid usage location.*'
    }
    It 'retries a licence assignment refused right after the usageLocation change' {
        $script:calls = 0
        Mock Start-Sleep -ModuleName PRA.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { $script:calls++; if ($script:calls -lt 3) { throw (New-GraphError 'Request_BadRequest' 'License assignment cannot be done for user with invalid usage location.') } }
        Set-PcmLicence -UserId 'u1' -SkuId 's1' -Operation Add -DisabledPlans @('p1')
        $script:calls | Should -Be 3
    }
    It 'does not retry a licence removal' {
        $script:calls = 0
        Mock Start-Sleep -ModuleName PRA.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { $script:calls++; throw (New-GraphError 'Request_BadRequest' 'x') }
        { Set-PcmLicence -UserId 'u1' -SkuId 's1' -Operation Remove } | Should -Throw
        $script:calls | Should -Be 1
    }
    It 'sends a request again when the connection failed without an answer (a timeout), never a refused one' {
        $script:calls = 0
        Mock Start-Sleep -ModuleName PRA.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud {
            $script:calls++
            if ($script:calls -lt 3) { throw [Exception]::new('A connection attempt failed (graph.microsoft.com:443)', [Net.Sockets.SocketException]::new(10060)) }
            [pscustomobject]@{ id = 'u1' }
        }
        (Invoke-PcmGraph -Uri 'v1.0/users/u1').id | Should -Be 'u1'
        $script:calls | Should -Be 3
        & (Get-Module PRA.Cloud) { Test-PcmNetworkFailure ([TimeoutException]::new('x')) } | Should -BeTrue
        & (Get-Module PRA.Cloud) { Test-PcmNetworkFailure ([InvalidOperationException]::new('BadRequest')) } | Should -BeFalse
        $script:calls = 0
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { $script:calls++; throw [Exception]::new('down', [IO.IOException]::new('reset')) }
        { Invoke-PcmGraph -Uri 'v1.0/users/u1' } | Should -Throw '*down*'
        $script:calls | Should -Be 5
    }
    It 'does not automatically retry an unprotected DELETE with an ambiguous response (<Fault>)' -TestCases @(
        @{ Fault = 'Network' }, @{ Fault = 'Gateway'; Status = 502 }, @{ Fault = 'Unavailable'; Status = 503 }, @{ Fault = 'Timeout'; Status = 504 }) {
        param($Fault, $Status)
        Mock Start-Sleep -ModuleName PRA.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud {
            if ($Fault -eq 'Network') { throw [TimeoutException]::new('ambiguous deletion') }
            $exception = [Exception]::new('ambiguous deletion')
            $exception | Add-Member -NotePropertyName Response -NotePropertyValue ([pscustomobject]@{ StatusCode = $Status })
            throw $exception
        }
        { Invoke-PcmGraph -Method DELETE -Uri 'v1.0/users/u1' -AllowNotFound -NoRetry } | Should -Throw '*ambiguous deletion*'
        Should -Invoke Invoke-MgGraphRequest -ModuleName PRA.Cloud -Times 1 -Exactly
        Should -Invoke Start-Sleep -ModuleName PRA.Cloud -Times 0 -Exactly
    }
    It 'retains five outer attempts for ordinary <Method> requests' -TestCases @(
        @{ Method = 'GET' }, @{ Method = 'PATCH' }, @{ Method = 'DELETE' }) {
        param($Method)
        Mock Start-Sleep -ModuleName PRA.Cloud { }
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud {
            $exception = [Exception]::new('unavailable')
            $exception | Add-Member Response ([pscustomobject]@{ StatusCode = 503 })
            throw $exception
        }
        { Invoke-PcmGraph -Method $Method -Uri 'v1.0/users/u1' } | Should -Throw '*unavailable*'
        Should -Invoke Invoke-MgGraphRequest -ModuleName PRA.Cloud -Times 5 -Exactly
    }
    It 'fails closed before sending when <Fault> prevents transport verification' -TestCases @(
        @{ Fault = 'NotInitialized' }, @{ Fault = 'ReadFails' }, @{ Fault = 'MissingContext' }, @{ Fault = 'MissingProperty' }, @{ Fault = 'NullRetry' }, @{ Fault = 'NonzeroRetry' }) {
        param($Fault)
        Mock Invoke-MgGraphRequest -ModuleName PRA.Cloud { throw 'Must not send' }
        switch ($Fault) {
            NotInitialized { & (Get-Module PRA.Cloud) { $script:GraphTransportReady = $false } }
            ReadFails { Mock Get-MgRequestContext -ModuleName PRA.Cloud { throw 'Readback failed' } }
            MissingContext { $global:PraFakeGraphRequest = $null }
            MissingProperty { $global:PraFakeGraphRequest = [pscustomobject]@{} }
            NullRetry { $global:PraFakeGraphRequest.MaxRetry = $null }
            NonzeroRetry { $global:PraFakeGraphRequest.MaxRetry = 3 }
        }
        { Invoke-PcmGraph -Method DELETE -Uri 'v1.0/users/u1' -NoRetry } | Should -Throw
        Should -Invoke Invoke-MgGraphRequest -ModuleName PRA.Cloud -Times 0 -Exactly
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
        { Import-PcmConfiguration -Path $path -Root $folder } | Should -Throw $Message
    }
}

Describe 'Link with the window: events, stop, questions, waves' {
    BeforeAll {
        $script:linkFolder = Join-Path $TestDrive 'link'
        $null = New-Item -ItemType Directory -Path $script:linkFolder
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
        foreach ($function in @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -in @('Read-PcmIdentityFile', 'ConvertTo-PcmIdentityKey') }, $true))) {
            . ([scriptblock]::Create($function.Extent.Text))
        }
    }
    AfterEach {
        & (Get-Module PRA.Common) { $script:EventFile = ''; $script:QuestionCount = 0 }
        Remove-Item Env:\PRA_STOP_FILE -ErrorAction SilentlyContinue
    }

    It 'writes nothing when the run does not belong to the window' {
        Write-PcmEvent 'item' @{ text = 'x' }
        @(Get-ChildItem -LiteralPath $script:linkFolder).Count | Should -Be 0
    }
    It 'Write-PcmProgress emits done/total/percent, with a rough time left once something is done' {
        $events = Join-Path $script:linkFolder 'progress.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        $context = New-TestContext
        Write-PcmProgress -Context $context -Phase 'Users' -Done 0 -Total 3000 -ElapsedSeconds 0
        Write-PcmProgress -Context $context -Phase 'Users' -Done 500 -Total 3000 -ElapsedSeconds 600
        Write-PcmProgress -Context $context -Phase 'Users' -Done 3000 -Total 3000 -ElapsedSeconds 3600
        $progress = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json } | Where-Object kind -eq 'progress')
        $progress.Count | Should -Be 3
        $progress[0].done | Should -Be 0; $progress[0].percent | Should -Be 0; $progress[0].eta | Should -BeNullOrEmpty
        $progress[1].done | Should -Be 500; $progress[1].percent | Should -Be 17; $progress[1].eta | Should -Match 'left at this pace'
        $progress[2].done | Should -Be 3000; $progress[2].percent | Should -Be 100; $progress[2].eta | Should -BeNullOrEmpty
    }
    It 'Write-PcmProgress does nothing for an empty total' {
        $events = Join-Path $script:linkFolder 'progress-empty.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        Write-PcmProgress -Context (New-TestContext) -Phase 'Users' -Done 0 -Total 0
        Test-Path -LiteralPath $events | Should -BeFalse
    }
    It 'asks the operator through the event file and reads the answer; a stop request answers no' {
        $events = Join-Path $script:linkFolder 'ask.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        Set-Content -LiteralPath "$events.answer-1" -Value 'yes'
        Request-PcmOperator -Question 'Pause the scheduler. Done?' | Should -BeTrue
        $env:PRA_STOP_FILE = Join-Path $script:linkFolder 'run.stop'
        Set-Content -LiteralPath $env:PRA_STOP_FILE -Value 'stop'
        Test-PcmStopRequest | Should -BeTrue
        Request-PcmOperator -Question 'Resume the scheduler. Done?' | Should -BeFalse
        $asked = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json } | Where-Object kind -eq 'ask')
        @($asked | ForEach-Object { $_.text }) | Should -Be @('Pause the scheduler. Done?', 'Resume the scheduler. Done?')
        $asked[1].answer | Should -Be "$events.answer-2"
    }
    It 'Entra Connect Manual mode asks the window when the run belongs to it' {
        $events = Join-Path $script:linkFolder 'manual.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        $env:PRA_EVENT_FILE = $events
        try {
            Set-Content -LiteralPath "$events.answer-1" -Value 'yes'
            $context = New-TestContext
            Invoke-PcmEntraConnect -Context $context -Operation Pause | Should -Be 'confirmed by the operator'
            Set-Content -LiteralPath "$events.answer-2" -Value 'no'
            { Invoke-PcmEntraConnect -Context $context -Operation Resume } | Should -Throw '*not confirmed by the operator*'
        } finally { Remove-Item Env:\PRA_EVENT_FILE -ErrorAction SilentlyContinue }
    }
    It 'announces the steps of the run in a plan, then gives each step its id, its state when it ends, and its objects with the name and the step they belong to' {
        $events = Join-Path $script:linkFolder 'plan.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        Mock Write-Host -ModuleName PRA.Common { }
        $context = New-TestContext
        Write-PcmPlan -Context $context -Titles @('Read', 'Work', 'Report', 'Never reached')
        $context.StepTotal | Should -Be 4
        Write-PcmStep -Context $context -Title 'Read' -Icon Plan
        $context.CurrentIdentity = 'alice@contoso.com'
        Write-PcmItem -Context $context -Status Ok -Text 'Ready'
        $context.CurrentIdentity = ''
        Write-PcmItem -Context $context -Status Warn -Text 'A warning of the step'
        Write-PcmStep -Context $context -Title 'Work' -Icon Plan
        Write-PcmStepCount -Context $context -Titles @('Work', 'Report', 'Not in the plan') -Count 3
        Write-PcmStepCount -Context $context -Titles @('Work') -Count 0
        Write-PcmLog -Context $context -Message 'The step failed' -Level Error
        Write-PcmStep -Context $context -Title 'Report' -Icon Plan
        Write-PcmStepState -Context $context -State skipped
        Write-PcmStep -Context $context -Title 'A step that is not in the plan' -Icon Plan
        Write-PcmStepState -Context $context
        $all = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json })
        $plan = @($all | Where-Object kind -eq 'plan')
        $plan.Count | Should -Be 1
        @($plan[0].steps | ForEach-Object { '{0}={1}' -f $_.id, $_.label }) | Should -Be @('s1=Read', 's2=Work', 's3=Report', 's4=Never reached')
        $steps = @($all | Where-Object kind -eq 'step')
        @($steps | ForEach-Object { if ($_.PSObject.Properties['state']) { '{0}:{1}' -f $_.id, $_.state } else { '{0}>{1}' -f $_.index, $(if ($_.PSObject.Properties['id']) { $_.id } else { '-' }) } }) |
            Should -Be @('1>s1', 's1:warn', '2>s2', 's2:fail', '3>s3', 's3:skipped', '4>-')
        $items = @($all | Where-Object kind -eq 'item')
        $items[0].name | Should -Be 'alice@contoso.com'; $items[0].identity | Should -Be 'alice@contoso.com'; $items[0].stepId | Should -Be 's1'
        $items[1].name | Should -Be ''; $items[1].stepId | Should -Be 's1'
        $count = @($all | Where-Object kind -eq 'count')
        $count.Count | Should -Be 1
        $count[0].count | Should -Be 3
        @($count[0].stepIds) | Should -Be @('s2', 's3')
    }
    It 'Wait-PcmCondition tells the window which wait it is, which attempt this is, how many there can be and when the next check comes' {
        $events = Join-Path $script:linkFolder 'wait.events.jsonl'
        & (Get-Module PRA.Common) { $script:EventFile = $args[0] } $events
        $result = Wait-PcmCondition -Test { $false } -TimeoutSeconds 1 -IntervalSeconds 1 -Phase 'Waiting for the test'
        $result.Ok | Should -BeFalse
        $progress = @(Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json } | Where-Object kind -eq 'progress')
        $progress.Count | Should -Be 1
        $progress[0].phase | Should -Be 'Waiting for the test'
        $progress[0].attempt | Should -Be 1
        $progress[0].maxAttempts | Should -Be 2
        $progress[0].nextInSeconds | Should -Be 1
        (Wait-PcmCondition -Test { 'ready' } -TimeoutSeconds 1).Ok | Should -BeTrue
        @(Get-Content -LiteralPath $events -Encoding UTF8).Count | Should -Be 1
    }
    It 'reads a wave from a list or a CSV file' {
        $list = Join-Path $script:linkFolder 'wave.txt'
        Set-Content -LiteralPath $list -Encoding UTF8 -Value @('# managers', 'alice@contoso.com', '', '  Bob@Contoso.com ', 'carol@contoso.com')
        Read-PcmIdentityFile -Path $list | Should -Be @('alice@contoso.com', 'Bob@Contoso.com', 'carol@contoso.com')
        $csv = Join-Path $script:linkFolder 'wave.csv'
        Set-Content -LiteralPath $csv -Encoding UTF8 -Value @('Identity;Note', 'compta@contoso.com;shared', 'carol@contoso.com;')
        Read-PcmIdentityFile -Path $csv | Should -Be @('compta@contoso.com', 'carol@contoso.com')
        Set-Content -LiteralPath $list -Encoding UTF8 -Value '# nothing'
        { Read-PcmIdentityFile -Path $list } | Should -Throw '*no identity*'
        { Read-PcmIdentityFile -Path (Join-Path $script:linkFolder 'missing.txt') } | Should -Throw '*file not found*'
    }
    It 'opens the window with -Gui alone, and refuses several ways to name the objects, or -CsvPath, -SearchBase, -Group and -MaxObjects for Collect' {
        $command = Get-Command (Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')
        $gui = $command.ParameterSets | Where-Object Name -eq 'Gui'
        @($gui.Parameters | Where-Object IsMandatory | ForEach-Object Name) | Should -Be @('Gui')
        @($gui.Parameters | ForEach-Object Name) | Should -Not -Contain 'Action'
        $source = [IO.File]::ReadAllText((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'))
        $source | Should -Match 'Choose one way to say which objects'
        $source | Should -Match '-CsvPath, -SearchBase, -Group and -MaxObjects are for Check, Convert and Recover'
        @($command.Parameters['CsvPath'].Aliases) | Should -Contain 'IdentityPath'
    }
}
