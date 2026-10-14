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
    Version : 1.3.0
#>
#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Exchange cmdlets used by Collect (imported from the remote session in Remote mode).
$script:ExchangeCommands = @('Get-Mailbox','Get-MailboxStatistics','Get-MailboxPermission','Get-ADPermission','Get-Recipient','Get-User',
    'Get-Group','Get-MailContact','Get-DistributionGroup','Get-DistributionGroupMember','Get-DynamicDistributionGroup','Get-ExchangeServer','Set-ADServerSettings',
    'Get-CalendarProcessing','Get-MailboxCalendarConfiguration')
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
    param([Parameter(Mandatory)][object]$Mailbox, [AllowNull()][object]$Statistics, [AllowNull()][hashtable]$ResourceSettings)
    $attributes = [ordered]@{}
    for ($i = 1; $i -le 15; $i++) {
        $name = "CustomAttribute$i"
        $value = if ($Mailbox.PSObject.Properties[$name]) { [string]$Mailbox.$name } else { '' }
        if ($value) { $attributes[$name] = $value }
    }
    $addresses = @($Mailbox.EmailAddresses | ForEach-Object { [string]$_ } | Where-Object { $_ })
    $onBehalf = @(if (-not $ResourceSettings) { $Mailbox.GrantSendOnBehalfTo | ForEach-Object { [string]$_ } | Where-Object { $_ } })
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
        resource_settings_json = $(if ($ResourceSettings) { ConvertTo-Json -InputObject $ResourceSettings -Depth 6 -Compress } else { $null })
    }
}

function Get-Pra2ResourceSettings {
    <# A portable booking snapshot: recipient references are SMTP addresses, not AD DNs. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Mailbox)
    if ([string]$Mailbox.RecipientTypeDetails -notin @('RoomMailbox','EquipmentMailbox')) { throw 'Resource settings require a room or equipment mailbox.' }
    $identity = [string]$Mailbox.DistinguishedName
    $dc = Get-Pra2DcParameter $Context
    $processing = Get-CalendarProcessing -Identity $identity @dc -ErrorAction Stop
    $work = Get-MailboxCalendarConfiguration -Identity $identity @dc -ErrorAction Stop
    if (-not $processing -or -not $work) { throw "Resource settings of $($Mailbox.PrimarySmtpAddress) could not be read: no calendar configuration returned." }
    $user = Get-User -Identity $identity @dc -ErrorAction Stop
    if (-not $user) { throw "Resource descriptive properties of $($Mailbox.PrimarySmtpAddress) could not be read." }
    $directory = [ordered]@{}
    foreach ($name in @('Office','StreetAddress','City','StateOrProvince','PostalCode','CountryOrRegion','Phone','Notes')) {
        if (-not $user.PSObject.Properties[$name]) { throw "Resource descriptive property $name is not exposed on $($Mailbox.PrimarySmtpAddress); no incomplete snapshot is written." }
        $directory[$name] = [string]$user.$name
    }
    foreach ($property in @('ResourceCapacity','ResourceCustom')) {
        if (-not $Mailbox.PSObject.Properties[$property]) { throw "Resource settings: $property missing on $($Mailbox.PrimarySmtpAddress)." }
    }
    $delivery = Get-Pra2DeliverySettings -Context $Context -Recipient $Mailbox
    $calendar = [ordered]@{}
    $required = @('AutomateProcessing','AllBookInPolicy','AllRequestInPolicy','AllRequestOutOfPolicy',
        'ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy','ScheduleOnlyDuringWorkHours')
    $scalars = @('AutomateProcessing','AllowConflicts','BookingWindowInDays','MaximumDurationInMinutes','AllowRecurringMeetings',
        'EnforceSchedulingHorizon','ScheduleOnlyDuringWorkHours','ConflictPercentageAllowed','MaximumConflictInstances',
        'ForwardRequestsToDelegates','DeleteAttachments','DeleteComments','RemovePrivateProperty','DeleteSubject',
        'AddOrganizerToSubject','DeleteNonCalendarItems','TentativePendingApproval','EnableResponseDetails','OrganizerInfo',
        'AllBookInPolicy','AllRequestInPolicy','AllRequestOutOfPolicy','AddAdditionalResponse','AdditionalResponse',
        'ProcessExternalMeetingMessages','RemoveOldMeetingMessages')
    foreach ($name in $required) {
        if (-not $processing.PSObject.Properties[$name]) { throw "Resource settings: $name missing on $($Mailbox.PrimarySmtpAddress); booking policies cannot be safely reproduced." }
    }
    foreach ($name in $scalars) {
        if ($processing.PSObject.Properties[$name]) {
            $value = $processing.$name
            $calendar[$name] = if ($name -in @('AutomateProcessing','AdditionalResponse')) { [string]$value } else { $value }
        }
    }
    foreach ($name in @('ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy')) {
        $recipients = [System.Collections.Generic.List[string]]::new()
        foreach ($reference in @($processing.$name)) {
            if (-not $reference) { continue }
            $trustee = Resolve-Pra2Trustee -Context $Context -Trustee ([string]$reference)
            if (-not $trustee.Resolved -or -not $trustee.Smtp) {
                throw "Resource settings: $name trustee '$reference' on $($Mailbox.PrimarySmtpAddress) has no resolved SMTP address. Correct it before Collect."
            }
            if (-not $recipients.Contains([string]$trustee.Smtp)) { $recipients.Add([string]$trustee.Smtp) }
        }
        $calendar[$name] = $recipients.ToArray()
    }
    $hours = [ordered]@{}
    foreach ($name in @('WorkingHoursStartTime','WorkingHoursEndTime','WorkingHoursDays','WorkingHoursTimeZone')) {
        # WorkDays is the service property; WorkingHoursDays remains the portable snapshot key.
        $readName = if ($name -eq 'WorkingHoursDays') { 'WorkDays' } else { $name }
        if (-not $work.PSObject.Properties[$readName]) { throw "Resource settings: $readName missing on $($Mailbox.PrimarySmtpAddress)." }
        $value = $work.$readName
        if ($null -eq $value) { throw "Resource settings: $readName unreadable on $($Mailbox.PrimarySmtpAddress)." }
        switch ($name) {
            'WorkingHoursDays' {
                $days = [string[]]@($value | ForEach-Object { ([string]$_ -split ',') | ForEach-Object { $_.Trim() } })
                if (-not $days.Count -or @($days | Where-Object { $_ -notin @('Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday','Weekdays','WeekendDays','AllDays') }).Count) {
                    throw "Resource settings: WorkDays invalid on $($Mailbox.PrimarySmtpAddress); no incomplete working-hours snapshot is written."
                }
                $hours[$name] = $days
            }
            'WorkingHoursTimeZone' {
                $hours[$name] = if ($value.PSObject.Properties['Id']) { [string]$value.Id } else { [string]$value }
                if (-not $hours[$name]) { throw "Resource settings: empty working-hours time zone on $($Mailbox.PrimarySmtpAddress)." }
            }
            default { $hours[$name] = ([timespan]::Parse([string]$value, [Globalization.CultureInfo]::InvariantCulture)).ToString('c') }
        }
    }
    $roomLists = @(if ([string]$Mailbox.RecipientTypeDetails -eq 'RoomMailbox') { Get-Pra2ResourceRoomLists -Context $Context -Mailbox $Mailbox })
    return @{ SchemaVersion = 2; Capacity = [int]$Mailbox.ResourceCapacity
        ResourceCustom = [string[]]@($Mailbox.ResourceCustom | ForEach-Object { [string]$_ }); Calendar = $calendar; CalendarConfiguration = $hours
        Directory = $directory; RoomLists = $roomLists; Delivery = $delivery }
}

function Get-Pra2DeliverySettings {
    param([hashtable]$Context, [object]$Recipient)
    $delivery = [ordered]@{}
    foreach ($name in @('RequireSenderAuthenticationEnabled','ModerationEnabled','SendModerationNotifications')) {
        if (-not $Recipient.PSObject.Properties[$name]) { throw "Resource delivery restriction $name is not exposed on $($Recipient.PrimarySmtpAddress)." }
        $delivery[$name] = if ($name -eq 'SendModerationNotifications') { [string]$Recipient.$name } else { $Recipient.$name }
        if ($name -ne 'SendModerationNotifications' -and $delivery[$name] -isnot [bool]) { throw "Resource delivery restriction $name is not a boolean on $($Recipient.PrimarySmtpAddress)." }
    }
    foreach ($name in @('AcceptMessagesOnlyFromSendersOrMembers','RejectMessagesFromSendersOrMembers','ModeratedBy','BypassModerationFromSendersOrMembers')) {
        if (-not $Recipient.PSObject.Properties[$name]) { throw "Resource delivery restriction $name is not exposed on $($Recipient.PrimarySmtpAddress)." }
        $values = @(foreach ($reference in @($Recipient.$name)) {
                if ($null -eq $reference) { continue }
                $trustee = Resolve-Pra2Trustee -Context $Context -Trustee ([string]$reference)
                if (-not $trustee.Resolved -or -not $trustee.Smtp) { throw "Resource delivery restriction $name has an unresolved recipient '$reference' on $($Recipient.PrimarySmtpAddress)." }
                [string]$trustee.Smtp
            })
        $delivery[$name] = @($values | Select-Object -Unique)
    }
    if ($delivery.SendModerationNotifications -notin @('Never','Internal','Always') -or ($delivery.ModerationEnabled -and -not $delivery.ModeratedBy.Count)) {
        throw "Resource delivery restriction moderation is incomplete on $($Recipient.PrimarySmtpAddress)."
    }
    return $delivery
}

function Get-Pra2ResourceRoomLists {
    param([hashtable]$Context, [object]$Mailbox)
    if (-not $Context.ContainsKey('ResourceRoomListInventory')) {
        $dc = Get-Pra2DcParameter $Context
        $inventory = @(foreach ($group in @(Get-DistributionGroup -Filter "RecipientTypeDetails -eq 'RoomList'" -ResultSize Unlimited @dc -ErrorAction Stop)) {
                foreach ($name in @('Guid','Name','DisplayName','Alias','PrimarySmtpAddress','HiddenFromAddressListsEnabled',
                        'EmailAddresses','LegacyExchangeDN','MemberJoinRestriction','MemberDepartRestriction')) {
                    if (-not $group.PSObject.Properties[$name]) { throw "Room list metadata $name is not exposed; no incomplete resource snapshot is written." }
                }
                if ([string]$group.PrimarySmtpAddress -notmatch '^[^@\s]+@[^@\s]+$' -or -not [string]$group.Guid -or
                    -not [string]$group.Name -or -not [string]$group.DisplayName -or -not [string]$group.Alias -or
                    $group.HiddenFromAddressListsEnabled -isnot [bool]) { throw 'Room list identity/SMTP/visibility is missing; no incomplete resource snapshot is written.' }
                if ([string]$group.MemberJoinRestriction -notin @('Closed','Open','ApprovalRequired') -or
                    [string]$group.MemberDepartRestriction -notin @('Closed','Open')) { throw "Room list $($group.PrimarySmtpAddress) has unreadable membership controls." }
                $members = @(Get-DistributionGroupMember -Identity ([string]$group.Guid) -ResultSize Unlimited @dc -ErrorAction Stop)
                foreach ($member in $members) {
                    if (-not $member.PSObject.Properties['Guid'] -or -not [string]$member.Guid) { throw "Room list $($group.PrimarySmtpAddress) has an unreadable member identity." }
                }
                [pscustomobject]@{ MemberGuids = @($members | ForEach-Object { [string]$_.Guid })
                    Metadata = [ordered]@{ SchemaVersion = 1; SourceGuid = [string]$group.Guid; Name = [string]$group.Name; DisplayName = [string]$group.DisplayName
                        Alias = [string]$group.Alias; PrimarySmtpAddress = [string]$group.PrimarySmtpAddress
                        HiddenFromAddressListsEnabled = [bool]$group.HiddenFromAddressListsEnabled
                        EmailAddresses = @($group.EmailAddresses | ForEach-Object { [string]$_ }); LegacyExchangeDN = [string]$group.LegacyExchangeDN
                        MemberJoinRestriction = [string]$group.MemberJoinRestriction; MemberDepartRestriction = [string]$group.MemberDepartRestriction
                        Delivery = (Get-Pra2DeliverySettings -Context $Context -Recipient $group) } }
            })
        $Context['ResourceRoomListInventory'] = $inventory
    }
    foreach ($entry in $Context.ResourceRoomListInventory) {
        if ([string]$Mailbox.Guid -in $entry.MemberGuids) { $entry.Metadata }
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
    Get-Pra2DelegateListLink, Get-Pra2SharedPermission, Get-Pra2Contact, Get-Pra2DistributionGroup, Get-Pra2ResourceSettings

# SIG # Begin signature block
# MIIghQYJKoZIhvcNAQcCoIIgdjCCIHICAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUudzc/dR61kEm
# ka08IBCTb6XTILQiOIYutoel3eVJYKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA9dmNvG7Wj
# Pz3l9AqGZVCNe1h8xh2sddm4/mgSjbsaBjANBgkqhkiG9w0BAQEFAASCAgCxyMyr
# 5qL0IANKmJBHzR8SLDE04uxTqnStWGOhtrHfL8psXKfVYRacDUNwcIZJEiM+00av
# Ehf+ije2uW8xha/efKP2YGI/vA9Pj8OBRWqAM/MwowvoNsGnWUR8GKAelvvcgmR0
# 2YI+im0HsYXE+Lf/699MWh5jHm5+iQ7uFGG8k7bmmM5xSlNObTE79hiKnggSkhFC
# fU5VDbIHNmO/f7obbjF4DKUSW/oW2Tjxpiiw8miApdiLlp5x8qECNEqMdPt6BW/X
# BGzzJEkyS65+7Rs+Cvbv/Obb2G4+GxhZ6Coflnx0E0s+J1ZkHWx5YG3sUR70qbsH
# rC9NEz6mf9VFZR4ictxPXjhuZq1DS1ySqlO55CpZ3pzbVCYa57Yxw4EHdmSsJejX
# fmBmqyZUBqMROTsn284gHVaQreAEIKqVJr7xjqxqRIsLNbQ/JCmtGKphI0jcPMwF
# kY48qE35l9YKdAfiblFlQayt3r/oDfZr6Vc/hc3S3Lvbb8QxLJVJarclNoK1VA9y
# 2gWZ5oksVdb0pBeEGm5p8Bu/HrhVzVLHYcRcem5gaTs6xVpTSiY/0dDQEIOtkDnC
# M89OzbNTUR9CtkatRKwDX2mSh8Mx8nE4do9I1pluXxI9jkCFoZzaXNdrve4/F9Ai
# vjfUK5eCrLZHxwuXGUtTc0PVUpJre37dc7HqUKGCDykwgg8lBgorBgEEAYI3AwMB
# MYIPFTCCDxEGCSqGSIb3DQEHAqCCDwIwgg7+AgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIG1k6TQOJl3Rtw6foXGc4AWojkGfcdTNAkYjp9Xj4CF4Aggks1c43qwG
# YhgPMjAyNjEwMTQxNTIzNDZaMAMCAQECEEO0tu9Q6EirurxhCK9kviygggwAMIIE
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
# HAYJKoZIhvcNAQkFMQ8XDTI2MTAxNDE1MjM0NlowKAYJKoZIhvcNAQk0MRswGTAL
# BglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIGclbQ7nM5XL
# 2/N3wnnc3Ib2wupqZgNuRqRhhkrgDjVZMIHJBgsqhkiG9w0BCRACLzGBuTCBtjCB
# szCBsAQgVCr5oWqNci5mEUl4iumUwYqaruWmXLNEolSa+Wx5x4swgYswd6R1MHMx
# CzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjER
# MA8GA1UECgwIU1NMIENvcnAxLzAtBgNVBAMMJlNTTC5jb20gVGltZXN0YW1waW5n
# IElzc3VpbmcgUlNBIENBIFIxAhAfaxZi0i4bbF3xwMGgYA44MAoGCCqGSM49BAMC
# BEcwRQIhAOP9wBYer4kRfLhfGhcDIfo0clcldGXX5hpz9v1+hTDbAiASyqn+jRWD
# RQqHx3/hjpHlzx1KxAD526MSugcNFSWnZw==
# SIG # End signature block
