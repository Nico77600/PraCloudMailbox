<#
.SYNOPSIS
    PRA Cloud Mailbox - Collect module: reads Exchange on-premises before the disaster.

.DESCRIPTION
    Runs in Windows PowerShell 5.1 with the Exchange cmdlets (Exchange Management Shell on this server,
    or PowerShell remoting to an Exchange server). Reads, never writes:

        - the mailboxes in scope (users, shared, optionally rooms and equipment): identities, GUIDs,
          addresses (X500 included), custom attributes;
        - the shared mailbox permissions: FullAccess (with AutoMapping), SendAs, SendOnBehalf. Trustees are
          resolved to their UPN / SMTP address; groups are expanded to their members (recursive);
        - optionally the mail contacts and the distribution groups (members, owners, delivery settings).

    The entry script stores the result in the SQLite database (PRA2.Store.psm1). After the disaster this
    snapshot is the only source of these values: Exchange and Active Directory are gone.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Exchange cmdlets used by Collect (imported from the remote session in Remote mode).
$script:ExchangeCommands = @('Get-Mailbox','Get-MailboxStatistics','Get-MailboxPermission','Get-ADPermission','Get-Recipient','Get-User',
    'Get-Group','Get-MailContact','Get-DistributionGroup','Get-DistributionGroupMember','Get-DynamicDistributionGroup','Get-ExchangeServer','Set-ADServerSettings')
# Trustees never reproduced: system and Exchange/AD administration accounts (on top of Collect.ExcludeTrustees).
$script:BuiltInTrusteePatterns = @('^NT AUTHORITY\\', '^AUTORITE NT\\', '^S-1-5-', '^Everyone$', '^Tout le monde$',
    '\\(Organization Management|Exchange Servers|Exchange Trusted Subsystem|Exchange Windows Permissions|Exchange Domain Servers|Exchange Enterprise Servers|Delegated Setup|Domain Admins|Enterprise Admins|Admins du domaine|Administrateurs de l.entreprise|Public Folder Management|Managed Availability Servers|Discovery Management|Exchange Install Domain Servers)$')
$script:SystemMailboxPatterns = @('^HealthMailbox', '^SystemMailbox\{', '^Microsoft Exchange', '^DiscoverySearchMailbox', '^FederatedEmail\.', '^Migration\.', '^SM_')
$script:KindByType = @{ UserMailbox = 'User'; SharedMailbox = 'Shared'; RoomMailbox = 'Room'; EquipmentMailbox = 'Equipment' }

#region Connection ---------------------------------------------------------------------------------

function Get-Pra2DcParameter {
    <# -DomainController splat when Exchange.DomainController is set (same DC for every read). #>
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = [string]$Context.Config.Exchange.DomainController
    if ($dc) { return @{ DomainController = $dc } }
    return @{}
}

function Connect-Pra2Exchange {
    <#
    .SYNOPSIS
        Makes the Exchange cmdlets available: local snap-in (Exchange Management Shell) or remote session.
    .OUTPUTS
        Text describing the connection (also stored in $Context.Server).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $exchange = $Context.Config.Exchange
    if ($exchange.ConnectionMode -eq 'Remote') {
        $uri = 'http://{0}/PowerShell/' -f $exchange.Server
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -ErrorAction Stop
        $Context['ExchangeSession'] = $session
        $module = Import-PSSession -Session $session -CommandName $script:ExchangeCommands -DisableNameChecking -AllowClobber -ErrorAction Stop
        Import-Module $module -Global -DisableNameChecking -ErrorAction Stop
        $Context.Server = [string]$exchange.Server
        $how = "remote session to $($exchange.Server) (Kerberos)"
    } else {
        if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
            try { Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop }
            catch { throw "Exchange cmdlets not available on this computer ($($_.Exception.Message)). Run Collect on an Exchange server or with the Exchange management tools, or set Exchange.ConnectionMode = 'Remote'." }
        }
        $Context.Server = $env:COMPUTERNAME
        $how = "Exchange Management Shell on $env:COMPUTERNAME"
    }
    if (Get-Command Set-ADServerSettings -ErrorAction SilentlyContinue) { Set-ADServerSettings -ViewEntireForest $true -ErrorAction SilentlyContinue }
    $version = ''
    $shortName = ($Context.Server -split '\.')[0]
    try { $server = Get-ExchangeServer -ErrorAction Stop | Where-Object { $_.Name -eq $shortName } | Select-Object -First 1; if ($server) { $version = [string]$server.AdminDisplayVersion } } catch { $version = '' }
    return ($how + $(if ($version) { " · $version" } else { '' }))
}

function Disconnect-Pra2Exchange {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($Context.ContainsKey('ExchangeSession') -and $Context['ExchangeSession']) {
        Remove-PSSession $Context['ExchangeSession'] -ErrorAction SilentlyContinue
        $Context['ExchangeSession'] = $null
    }
}
#endregion

#region Mailboxes ----------------------------------------------------------------------------------

function Get-Pra2RecipientType {
    <# Exchange RecipientTypeDetails of the mailbox kinds in scope, after the -Scope filter of the command line. #>
    param([Parameter(Mandatory)][hashtable]$Scope, [ValidateSet('All','UsersOnly','SharedOnly')][string]$Filter = 'All')
    $types = @()
    if ($Scope.IncludeUsers -and $Filter -ne 'SharedOnly') { $types += 'UserMailbox' }
    if ($Filter -ne 'UsersOnly') {
        if ($Scope.IncludeShared) { $types += 'SharedMailbox' }
        if ($Scope.IncludeRoom) { $types += 'RoomMailbox' }
        if ($Scope.IncludeEquipment) { $types += 'EquipmentMailbox' }
    }
    return $types
}

function Test-Pra2SystemMailbox {
    param([AllowEmptyString()][string]$Name)
    foreach ($pattern in $script:SystemMailboxPatterns) { if ($Name -match $pattern) { return $true } }
    return $false
}

function Get-Pra2GroupMemberDn {
    <#
    .SYNOPSIS
        Distinguished names of the members of a group, recursively (nested groups expanded, loops ignored).
        Returns user/contact members only; the nested groups are walked, not returned.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Identity)
    $dc = Get-Pra2DcParameter $Context
    $seen = @{}; $result = [System.Collections.Generic.List[string]]::new()
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($Identity)
    while ($queue.Count) {
        $current = $queue.Dequeue()
        if ($seen.ContainsKey($current)) { continue }
        $seen[$current] = $true
        $group = Get-Group -Identity $current -ErrorAction Stop @dc
        foreach ($member in @($group.Members)) {
            $dn = if ($member -is [string]) { $member } elseif ($member.PSObject.Properties['DistinguishedName']) { [string]$member.DistinguishedName } else { [string]$member }
            if (-not $dn) { continue }
            $isGroup = $false
            try { $null = Get-Group -Identity $dn -ErrorAction Stop @dc; $isGroup = $true } catch { $isGroup = $false }
            if ($isGroup) { $queue.Enqueue($dn) } elseif (-not $result.Contains($dn)) { [void]$result.Add($dn) }
        }
    }
    return $result.ToArray()
}

function Get-Pra2ScopeMailbox {
    <#
    .SYNOPSIS
        The mailboxes to collect: -Identity, or the configured scope (Auto, OU, Group, Csv), sorted by sAMAccountName.
    .OUTPUTS
        Exchange mailbox objects. Excluded objects are reported in $Context.Excluded.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [string]$Identity, [ValidateSet('All','UsersOnly','SharedOnly')][string]$Filter = 'All')
    $scope = $Context.Config.Scope
    $types = @(Get-Pra2RecipientType -Scope $scope -Filter $Filter)
    if (-not $types.Count) { throw "No mailbox type left after -Scope $Filter and the configuration (Scope.Include*)." }
    $dc = Get-Pra2DcParameter $Context
    $mailboxes = @()
    if ($Identity) {
        $mailbox = Get-Mailbox -Identity $Identity -ErrorAction Stop @dc
        if ([string]$mailbox.RecipientTypeDetails -notin $types) { throw ("{0} is a {1}: not a mailbox type in scope ({2})." -f $Identity, $mailbox.RecipientTypeDetails, ($types -join ', ')) }
        $mailboxes = @($mailbox)
    } else {
        switch ($scope.Mode) {
            'Auto' {
                $ou = @{}; if ($scope.SearchBase) { $ou = @{ OrganizationalUnit = $scope.SearchBase } }
                $mailboxes = @(Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $types -ErrorAction Stop @ou @dc)
            }
            'OU' { $mailboxes = @(Get-Mailbox -OrganizationalUnit $scope.SearchBase -ResultSize Unlimited -RecipientTypeDetails $types -ErrorAction Stop @dc) }
            'Group' {
                foreach ($dn in (Get-Pra2GroupMemberDn -Context $Context -Identity $scope.GroupDN)) {
                    $mailbox = Get-Mailbox -Identity $dn -ErrorAction SilentlyContinue @dc
                    if ($mailbox -and [string]$mailbox.RecipientTypeDetails -in $types) { $mailboxes += $mailbox }
                }
            }
            'Csv' {
                $entries = @(Import-Csv -LiteralPath $scope.CsvPath)
                if ($entries.Count -and -not $entries[0].PSObject.Properties['Identity']) { throw "Scope.CsvPath: column Identity missing ($($scope.CsvPath))." }
                foreach ($entry in $entries) {
                    if (-not [string]$entry.Identity) { continue }
                    $mailbox = Get-Mailbox -Identity ([string]$entry.Identity).Trim() -ErrorAction Stop @dc
                    if ([string]$mailbox.RecipientTypeDetails -in $types) { $mailboxes += $mailbox }
                    else { [void]$Context.Excluded.Add(('{0}: {1} not in scope' -f $entry.Identity, $mailbox.RecipientTypeDetails)) }
                }
            }
        }
    }
    $excludedNames = @($scope.ExcludeSamAccountNames | ForEach-Object { [string]$_ })
    $kept = @()
    $seen = @{}
    foreach ($mailbox in $mailboxes) {
        $key = [string]$mailbox.Guid
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        if (Test-Pra2SystemMailbox ([string]$mailbox.Name)) { [void]$Context.Excluded.Add(('{0}: system mailbox' -f $mailbox.Name)); continue }
        if ([string]$mailbox.SamAccountName -in $excludedNames) { [void]$Context.Excluded.Add(('{0}: Scope.ExcludeSamAccountNames' -f $mailbox.SamAccountName)); continue }
        $kept += $mailbox
    }
    return @($kept | Sort-Object { [string]$_.SamAccountName })
}

function ConvertTo-Pra2ImmutableId {
    <# Base64 of the objectGUID: the onPremisesImmutableId of the Entra ID object when the source anchor is the objectGUID or ms-DS-ConsistencyGuid. #>
    param([Parameter(Mandatory)][object]$Guid)
    return [Convert]::ToBase64String(([guid][string]$Guid).ToByteArray())
}

function ConvertTo-Pra2SizeBytes {
    <# '1.2 GB (1,288,490,188 bytes)' or ByteQuantifiedSize -> bytes; $null when unknown. #>
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value.PSObject.Methods['ToBytes']) { return [long]$Value.ToBytes() }
    $text = [string]$Value
    if ($text -match '\(([\d,\. \u00A0\u202F]+) bytes\)') { return [long]($Matches[1] -replace '[^\d]', '') }
    return $null
}

function ConvertTo-Pra2MailboxRecord {
    <#
    .SYNOPSIS
        Exchange mailbox (+ statistics) -> row of the mailbox table.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Mailbox, [AllowNull()][object]$Statistics)
    $attributes = [ordered]@{}
    for ($i = 1; $i -le 15; $i++) {
        $name = "CustomAttribute$i"
        $value = if ($Mailbox.PSObject.Properties[$name]) { [string]$Mailbox.$name } else { '' }
        if ($value) { $attributes[$name] = $value }
    }
    $addresses = @($Mailbox.EmailAddresses | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $onBehalf = @($Mailbox.GrantSendOnBehalfTo | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $type = [string]$Mailbox.RecipientTypeDetails
    $kind = $script:KindByType[$type]
    if (-not $kind) { throw "Unsupported mailbox type $type for $($Mailbox.Identity)." }
    $archive = if ($Mailbox.PSObject.Properties['ArchiveGuid'] -and [string]$Mailbox.ArchiveGuid -and [string]$Mailbox.ArchiveGuid -ne [guid]::Empty.ToString()) { [string]$Mailbox.ArchiveGuid } else { $null }
    return [ordered]@{
        object_guid = [string]$Mailbox.Guid; kind = $kind; recipient_type_details = $type
        sam_account_name = [string]$Mailbox.SamAccountName; user_principal_name = [string]$Mailbox.UserPrincipalName
        name = [string]$Mailbox.Name; display_name = [string]$Mailbox.DisplayName; alias = [string]$Mailbox.Alias
        distinguished_name = [string]$Mailbox.DistinguishedName; organizational_unit = [string]$Mailbox.OrganizationalUnit
        primary_smtp_address = [string]$Mailbox.PrimarySmtpAddress; windows_email_address = [string]$Mailbox.WindowsEmailAddress
        legacy_exchange_dn = [string]$Mailbox.LegacyExchangeDN; exchange_guid = [string]$Mailbox.ExchangeGuid; archive_guid = $archive
        archive_status = $(if ($Mailbox.PSObject.Properties['ArchiveStatus']) { [string]$Mailbox.ArchiveStatus } else { $null })
        hidden_from_address_lists = [bool]$Mailbox.HiddenFromAddressListsEnabled; immutable_id = (ConvertTo-Pra2ImmutableId $Mailbox.Guid)
        database = [string]$Mailbox.Database; email_addresses_json = (ConvertTo-Json -InputObject @($addresses) -Compress)
        custom_attributes_json = ($attributes | ConvertTo-Json -Compress); grant_send_on_behalf_json = (ConvertTo-Json -InputObject @($onBehalf) -Compress)
        item_count = $(if ($Statistics) { [long]$Statistics.ItemCount } else { $null })
        total_item_size_bytes = $(if ($Statistics) { ConvertTo-Pra2SizeBytes $Statistics.TotalItemSize } else { $null })
    }
}
#endregion

#region Shared mailbox permissions -------------------------------------------------------------------

function Test-Pra2ExcludedTrustee {
    <# $true for system / administration trustees and the trustees of Collect.ExcludeTrustees (wildcards allowed). #>
    param([Parameter(Mandatory)][hashtable]$Context, [AllowEmptyString()][string]$Trustee)
    if (-not $Trustee) { return $true }
    foreach ($pattern in $script:BuiltInTrusteePatterns) { if ($Trustee -match $pattern) { return $true } }
    $short = ($Trustee -split '\\')[-1]
    foreach ($exclude in @($Context.Config.Collect.ExcludeTrustees)) {
        $text = [string]$exclude
        if ($text -and ($Trustee -like $text -or $short -like $text)) { return $true }
    }
    return $false
}

function Resolve-Pra2Trustee {
    <#
    .SYNOPSIS
        Trustee as written by Exchange (DOMAIN\sam, name or DN) -> GUID, kind (User, Group, Contact), UPN, SMTP, DN.
        Results are cached for the run in $Context.TrusteeCache.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Trustee)
    if (-not $Context.ContainsKey('TrusteeCache')) { $Context['TrusteeCache'] = @{} }
    $cache = $Context['TrusteeCache']
    if ($cache.ContainsKey($Trustee)) { return $cache[$Trustee] }
    $dc = Get-Pra2DcParameter $Context
    $result = [ordered]@{ Resolved = $false; Guid = $null; Kind = 'Unknown'; Upn = $null; Smtp = $null; Name = $null; Dn = $null; Note = '' }
    $recipient = $null
    try { $recipient = Get-Recipient -Identity $Trustee -ErrorAction Stop @dc } catch { $recipient = $null }
    if ($recipient) {
        $type = [string]$recipient.RecipientType
        $result.Kind = if ($type -match 'Group') { 'Group' } elseif ($type -match 'Contact') { 'Contact' } else { 'User' }
        $result.Guid = [string]$recipient.Guid; $result.Smtp = [string]$recipient.PrimarySmtpAddress
        $result.Name = [string]$recipient.Name; $result.Dn = [string]$recipient.DistinguishedName; $result.Resolved = $true
        if ($result.Kind -eq 'User') {
            try { $result.Upn = [string](Get-User -Identity $result.Dn -ErrorAction Stop @dc).UserPrincipalName } catch { $result.Note = 'UPN not read' }
        }
    } else {
        try {
            $user = Get-User -Identity $Trustee -ErrorAction Stop @dc
            $result.Kind = 'User'; $result.Guid = [string]$user.Guid; $result.Upn = [string]$user.UserPrincipalName
            $result.Name = [string]$user.Name; $result.Dn = [string]$user.DistinguishedName; $result.Resolved = $true; $result.Note = 'not mail-enabled'
        } catch {
            try {
                $group = Get-Group -Identity $Trustee -ErrorAction Stop @dc
                $result.Kind = 'Group'; $result.Guid = [string]$group.Guid; $result.Name = [string]$group.Name
                $result.Dn = [string]$group.DistinguishedName; $result.Resolved = $true; $result.Note = 'group not mail-enabled'
            } catch { $result.Note = 'not found (deleted account or orphan SID)' }
        }
    }
    $cache[$Trustee] = [pscustomobject]$result
    return $cache[$Trustee]
}

function Get-Pra2DelegateListLink {
    <#
    .SYNOPSIS
        msExchDelegateListLink of a mailbox (the FullAccess trustees with AutoMapping), read in AD with ADSI.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$DistinguishedName)
    $dc = [string]$Context.Config.Exchange.DomainController
    $path = if ($dc) { "LDAP://$dc/$DistinguishedName" } else { "LDAP://$DistinguishedName" }
    $entry = New-Object System.DirectoryServices.DirectoryEntry($path)
    try { return @($entry.Properties['msExchDelegateListLink'] | ForEach-Object { [string]$_ }) }
    finally { $entry.Dispose() }
}

function Get-Pra2SharedPermission {
    <#
    .SYNOPSIS
        FullAccess, SendAs and SendOnBehalf of one mailbox, trustees resolved, groups expanded.
    .OUTPUTS
        @{ Permissions = rows of the permission table; Members = rows of the group_member table; Warnings = texts }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Mailbox)
    $dc = Get-Pra2DcParameter $Context
    $identity = [string]$Mailbox.DistinguishedName
    $mailboxGuid = [string]$Mailbox.Guid
    $warnings = [System.Collections.Generic.List[string]]::new()
    $entries = [System.Collections.Generic.List[object]]::new()
    foreach ($ace in @(Get-MailboxPermission -Identity $identity -ErrorAction Stop @dc)) {
        if ($ace.IsInherited -or $ace.Deny) { continue }
        if (-not (@($ace.AccessRights | ForEach-Object { [string]$_ }) -match 'FullAccess')) { continue }
        [void]$entries.Add(@{ Right = 'FullAccess'; Trustee = [string]$ace.User })
    }
    foreach ($ace in @(Get-ADPermission -Identity $identity -ErrorAction Stop @dc)) {
        if ($ace.IsInherited -or $ace.Deny) { continue }
        if (-not (@($ace.ExtendedRights | ForEach-Object { [string]$_ }) -match 'Send-As')) { continue }
        [void]$entries.Add(@{ Right = 'SendAs'; Trustee = [string]$ace.User })
    }
    foreach ($delegate in @($Mailbox.GrantSendOnBehalfTo)) {
        $text = if ($delegate -is [string]) { $delegate } elseif ($delegate.PSObject.Properties['DistinguishedName']) { [string]$delegate.DistinguishedName } else { [string]$delegate }
        if ($text) { [void]$entries.Add(@{ Right = 'SendOnBehalf'; Trustee = $text }) }
    }
    $autoMapping = @{}
    if (@($entries | Where-Object { $_.Right -eq 'FullAccess' }).Count) {
        try { foreach ($dn in (Get-Pra2DelegateListLink -Context $Context -DistinguishedName $identity)) { $autoMapping[$dn.ToLowerInvariant()] = $true } }
        catch { [void]$warnings.Add(('{0}: AutoMapping not read ({1})' -f $Mailbox.PrimarySmtpAddress, $_.Exception.Message)) }
    }
    $permissions = [System.Collections.Generic.List[object]]::new()
    $members = [System.Collections.Generic.List[object]]::new()
    $expanded = @{}
    $seenPair = @{}
    foreach ($entry in $entries) {
        $trustee = $entry.Trustee
        if (Test-Pra2ExcludedTrustee -Context $Context -Trustee $trustee) { continue }
        $resolved = Resolve-Pra2Trustee -Context $Context -Trustee $trustee
        if ($resolved.Guid -and $resolved.Guid -eq $mailboxGuid) { continue }
        $pair = '{0}|{1}' -f $entry.Right, $(if ($resolved.Guid) { $resolved.Guid } else { $trustee.ToLowerInvariant() })
        if ($seenPair.ContainsKey($pair)) { continue }
        $seenPair[$pair] = $true
        if (-not $resolved.Resolved) { [void]$warnings.Add(('{0}: {1} trustee {2} {3}' -f $Mailbox.PrimarySmtpAddress, $entry.Right, $trustee, $resolved.Note)) }
        $mapping = $null
        if ($entry.Right -eq 'FullAccess' -and $resolved.Dn) { $mapping = [bool]$autoMapping.ContainsKey(([string]$resolved.Dn).ToLowerInvariant()) }
        [void]$permissions.Add([ordered]@{
            mailbox_guid = $mailboxGuid; access_right = $entry.Right; trustee = $trustee; trustee_guid = $resolved.Guid; trustee_kind = $resolved.Kind
            trustee_upn = $resolved.Upn; trustee_smtp = $resolved.Smtp; trustee_name = $resolved.Name; trustee_dn = $resolved.Dn
            auto_mapping = $mapping; resolved = [bool]$resolved.Resolved; note = $resolved.Note
        })
        if ($resolved.Kind -eq 'Group' -and $Context.Config.Collect.ExpandGroupTrustees -and $resolved.Dn -and -not $expanded.ContainsKey($resolved.Guid)) {
            $expanded[$resolved.Guid] = $true
            try {
                foreach ($memberDn in (Get-Pra2GroupMemberDn -Context $Context -Identity $resolved.Dn)) {
                    $member = Resolve-Pra2Trustee -Context $Context -Trustee $memberDn
                    [void]$members.Add([ordered]@{
                        group_guid = $resolved.Guid; member_guid = $member.Guid; member_kind = $member.Kind; member_upn = $member.Upn
                        member_smtp = $member.Smtp; member_name = $member.Name; via_group_guid = $resolved.Guid; depth = 1
                    })
                }
            } catch { [void]$warnings.Add(('{0}: members of group {1} not read ({2})' -f $Mailbox.PrimarySmtpAddress, $resolved.Name, $_.Exception.Message)) }
        }
    }
    return @{ Permissions = $permissions.ToArray(); Members = $members.ToArray(); Warnings = $warnings.ToArray() }
}
#endregion

#region Contacts and distribution groups -------------------------------------------------------------

function Get-Pra2CustomAttributeJson {
    param([Parameter(Mandatory)][object]$Object)
    $attributes = [ordered]@{}
    for ($i = 1; $i -le 15; $i++) {
        $name = "CustomAttribute$i"
        if ($Object.PSObject.Properties[$name] -and [string]$Object.$name) { $attributes[$name] = [string]$Object.$name }
    }
    return ($attributes | ConvertTo-Json -Compress)
}

function Get-Pra2Contact {
    <# Mail contacts (optionally under Collect.ContactsSearchBase) -> rows of the contact table. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = Get-Pra2DcParameter $Context
    $ou = @{}; if ($Context.Config.Collect.ContactsSearchBase) { $ou = @{ OrganizationalUnit = $Context.Config.Collect.ContactsSearchBase } }
    $rows = foreach ($contact in @(Get-MailContact -ResultSize Unlimited -ErrorAction Stop @ou @dc)) {
        [ordered]@{
            object_guid = [string]$contact.Guid; name = [string]$contact.Name; display_name = [string]$contact.DisplayName; alias = [string]$contact.Alias
            external_email_address = [string]$contact.ExternalEmailAddress; primary_smtp_address = [string]$contact.PrimarySmtpAddress
            legacy_exchange_dn = [string]$contact.LegacyExchangeDN; hidden_from_address_lists = [bool]$contact.HiddenFromAddressListsEnabled
            organizational_unit = [string]$contact.OrganizationalUnit
            email_addresses_json = (ConvertTo-Json -InputObject @($contact.EmailAddresses | ForEach-Object { [string]$_ }) -Compress)
            custom_attributes_json = (Get-Pra2CustomAttributeJson $contact)
        }
    }
    return @($rows)
}

function Get-Pra2DistributionGroup {
    <#
    .SYNOPSIS
        Distribution and mail-enabled security groups (and dynamic groups when enabled), with their members.
    .OUTPUTS
        @{ Groups = rows of distribution_group; Members = rows of dl_member }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = Get-Pra2DcParameter $Context
    $collect = $Context.Config.Collect
    $ou = @{}; if ($collect.DistributionGroupsSearchBase) { $ou = @{ OrganizationalUnit = $collect.DistributionGroupsSearchBase } }
    $groups = [System.Collections.Generic.List[object]]::new()
    $members = [System.Collections.Generic.List[object]]::new()
    $list = $null
    $settingNames = @('RequireSenderAuthenticationEnabled','MemberJoinRestriction','MemberDepartRestriction','ModerationEnabled','SendModerationNotifications',
        'AcceptMessagesOnlyFromSendersOrMembers','RejectMessagesFromSendersOrMembers','ModeratedBy','BypassModerationFromSendersOrMembers',
        'GrantSendOnBehalfTo','ReportToManagerEnabled','ReportToOriginatorEnabled','SendOofMessageToOriginatorEnabled','MaxSendSize','MaxReceiveSize','Notes')
    foreach ($group in @(Get-DistributionGroup -ResultSize Unlimited -ErrorAction Stop @ou @dc)) {
        $settings = [ordered]@{}
        foreach ($name in $settingNames) {
            if (-not $group.PSObject.Properties[$name]) { continue }
            $value = $group.$name
            $settings[$name] = if ($value -is [System.Collections.IEnumerable] -and $value -isnot [string]) { @($value | ForEach-Object { [string]$_ }) } else { [string]$value }
        }
        $list = @(Get-DistributionGroupMember -Identity ([string]$group.DistinguishedName) -ResultSize Unlimited -ErrorAction Stop @dc)
        foreach ($member in $list) {
            [void]$members.Add([ordered]@{
                group_guid = [string]$group.Guid; member_guid = [string]$member.Guid; member_kind = [string]$member.RecipientType
                member_name = [string]$member.Name; member_smtp = [string]$member.PrimarySmtpAddress; member_upn = $null
            })
        }
        [void]$groups.Add([ordered]@{
            object_guid = [string]$group.Guid; kind = $(if ([string]$group.RecipientTypeDetails -match 'Security') { 'Security' } else { 'Distribution' })
            name = [string]$group.Name; display_name = [string]$group.DisplayName; alias = [string]$group.Alias
            primary_smtp_address = [string]$group.PrimarySmtpAddress; legacy_exchange_dn = [string]$group.LegacyExchangeDN
            hidden_from_address_lists = [bool]$group.HiddenFromAddressListsEnabled; organizational_unit = [string]$group.OrganizationalUnit
            email_addresses_json = (ConvertTo-Json -InputObject @($group.EmailAddresses | ForEach-Object { [string]$_ }) -Compress)
            managed_by_json = (ConvertTo-Json -InputObject @($group.ManagedBy | ForEach-Object { [string]$_ }) -Compress)
            settings_json = ($settings | ConvertTo-Json -Compress -Depth 4); recipient_filter = $null; member_count = $list.Count
        })
    }
    if ($collect.DynamicDistributionGroups) {
        foreach ($group in @(Get-DynamicDistributionGroup -ResultSize Unlimited -ErrorAction Stop @ou @dc)) {
            [void]$groups.Add([ordered]@{
                object_guid = [string]$group.Guid; kind = 'Dynamic'; name = [string]$group.Name; display_name = [string]$group.DisplayName
                alias = [string]$group.Alias; primary_smtp_address = [string]$group.PrimarySmtpAddress; legacy_exchange_dn = [string]$group.LegacyExchangeDN
                hidden_from_address_lists = [bool]$group.HiddenFromAddressListsEnabled; organizational_unit = [string]$group.OrganizationalUnit
                email_addresses_json = (ConvertTo-Json -InputObject @($group.EmailAddresses | ForEach-Object { [string]$_ }) -Compress)
                managed_by_json = (ConvertTo-Json -InputObject @($group.ManagedBy | ForEach-Object { [string]$_ }) -Compress)
                settings_json = '{}'; recipient_filter = [string]$group.RecipientFilter; member_count = $null
            })
        }
    }
    return @{ Groups = $groups.ToArray(); Members = $members.ToArray() }
}
#endregion

Export-ModuleMember -Function Connect-Pra2Exchange, Disconnect-Pra2Exchange, Get-Pra2RecipientType, Get-Pra2ScopeMailbox, Get-Pra2GroupMemberDn,
    ConvertTo-Pra2ImmutableId, ConvertTo-Pra2SizeBytes, ConvertTo-Pra2MailboxRecord, Test-Pra2ExcludedTrustee, Resolve-Pra2Trustee,
    Get-Pra2DelegateListLink, Get-Pra2SharedPermission, Get-Pra2Contact, Get-Pra2DistributionGroup
