<#
.SYNOPSIS
    PRA Cloud Mailbox - Collect module: reads Exchange on-premises before the disaster.

.DESCRIPTION
    Runs in PowerShell 7.5 or later, from any domain-joined Windows server that reaches an Exchange server with Kerberos: the Exchange
    cmdlets come from a remote session to Exchange.Server (Import-PSSession), so every object Collect reads is a deserialized copy
    (properties kept, methods lost, sizes as text such as '1.2 GB (1,288,490,188 bytes)', references to directory objects as ADObjectId
    copies). Reads, never writes:

        - the mailboxes in scope (users, shared, optionally rooms and equipment): identities, GUIDs,
          addresses (X500 included), custom attributes;
        - the shared mailbox permissions: FullAccess (with AutoMapping), SendAs, SendOnBehalf. Trustees are
          resolved to their UPN / SMTP address; groups are expanded to their members (recursive);
        - optionally the mail contacts and the distribution groups (members, owners, delivery settings).

    The entry script stores the result in the SQLite database (PRA.Store.psm1). After the disaster this
    snapshot is the only source of these values: Exchange and Active Directory are gone.

    The region "The daily task" serves -Action Schedule: the settings of the daily Collect task, the checks of its account (a domain user that Exchange accepts, and on a
    domain controller the right to log on as a batch job) and its registration in the Windows Task Scheduler.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
#>
#Requires -Version 7.5
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# >>> PRA package integrity bootstrap: identical in every runtime PowerShell file (checked by the tests).
# It authenticates module\Test-PackageIntegrity.ps1 before running it; there is no unsigned or bypass path.
& {
    param([Parameter(Mandatory)][string]$PackageRoot)
    $verifier = [IO.Path]::GetFullPath([IO.Path]::Combine($PackageRoot, 'module', 'Test-PackageIntegrity.ps1'))
    $cursor = $verifier
    while ($cursor) {
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if (-not [string]::Equals([string]$signature.Status, 'Valid', [StringComparison]::Ordinal) -or $null -eq $signature.SignerCertificate -or
        -not [string]::Equals([string]$signature.SignerCertificate.Thumbprint, '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6', [StringComparison]::Ordinal) -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
    }
    & $verifier -Root $PackageRoot
} ([IO.Path]::GetDirectoryName($PSScriptRoot))
# <<< PRA package integrity bootstrap
# Exchange cmdlets used by Collect (imported from the remote session).
$script:ExchangeCommands = @('Get-Mailbox','Get-MailboxStatistics','Get-MailboxPermission','Get-ADPermission','Get-Recipient','Get-User',
    'Get-Group','Get-MailContact','Get-DistributionGroup','Get-DistributionGroupMember','Get-DynamicDistributionGroup','Get-ExchangeServer','Set-ADServerSettings',
    'Get-CalendarProcessing','Get-MailboxCalendarConfiguration')
# Trustees never reproduced: system and Exchange/AD administration accounts (on top of Collect.ExcludeTrustees).
$script:BuiltInTrusteePatterns = @('^NT AUTHORITY\\', '^AUTORITE NT\\', '^S-1-5-', '^Everyone$', '^Tout le monde$',
    '\\(Organization Management|Exchange Servers|Exchange Trusted Subsystem|Exchange Windows Permissions|Exchange Domain Servers|Exchange Enterprise Servers|Delegated Setup|Domain Admins|Enterprise Admins|Admins du domaine|Administrateurs de l.entreprise|Public Folder Management|Managed Availability Servers|Discovery Management|Exchange Install Domain Servers)$')
$script:KindByType = @{ UserMailbox = 'User'; SharedMailbox = 'Shared'; RoomMailbox = 'Room'; EquipmentMailbox = 'Equipment' }

#region Connection ---------------------------------------------------------------------------------

function Get-PcmDcParameter {
    <# -DomainController splat when Exchange.DomainController is set (same DC for every read). #>
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = [string]$Context.Config.Exchange.DomainController
    if ($dc) { return @{ DomainController = $dc } }
    return @{}
}

function Connect-PcmExchange {
    <#
    .SYNOPSIS
        Opens the remote session to Exchange.Server (Kerberos) and imports the Exchange cmdlets that Collect uses.
    .OUTPUTS
        Text describing the connection (also stored in $Context.Server).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $server = ([string]$Context.Config.Exchange.Server).Trim()
    if (-not $server) { throw 'Exchange.Server is empty. Collect reads Exchange through a remote session: set Exchange.Server to the host name of an Exchange server (for example EX01.contoso.com) in the configuration file.' }
    $uri = 'http://{0}/PowerShell/' -f $server
    try {
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -ErrorAction Stop
        $Context['ExchangeSession'] = $session
        $module = Import-PSSession -Session $session -CommandName $script:ExchangeCommands -DisableNameChecking -AllowClobber -ErrorAction Stop
        Import-Module $module -Global -DisableNameChecking -ErrorAction Stop
    } catch {
        Disconnect-PcmExchange -Context $Context
        throw ("Cannot open the remote session to Exchange on {0}: {1} Collect needs a domain-joined computer, an account allowed to use Exchange remote PowerShell, and an Exchange server that answers on {2} with Kerberos." -f $server, $_.Exception.Message.Trim(), $uri)
    }
    $Context.Server = $server
    try {
        if (-not (Get-Command Set-ADServerSettings -ErrorAction SilentlyContinue)) { throw 'the command Set-ADServerSettings is not available to this account' }
        Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
    } catch {
        Disconnect-PcmExchange -Context $Context
        throw ("Collect cannot read the whole forest on {0}: {1}. Without it Exchange would answer for its own domain only, and the mailboxes of the other domains would be missing from the snapshot. Check that the account is in the Exchange role group View-Only Organization Management, wait a few minutes for the change to replicate, then run Collect again. Nothing was written." -f $server, $_.Exception.Message.Trim().TrimEnd('.'))
    }
    $version = ''
    $shortName = ($server -split '\.')[0]
    try { $found = Get-ExchangeServer -ErrorAction Stop | Where-Object { $_.Name -eq $shortName } | Select-Object -First 1; if ($found) { $version = [string]$found.AdminDisplayVersion } } catch { $version = '' }
    return ("remote session to $server (Kerberos)" + $(if ($version) { " $([char]0x00B7) $version" } else { '' }))
}

function Disconnect-PcmExchange {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    if ($Context.ContainsKey('ExchangeSession') -and $Context['ExchangeSession']) {
        Remove-PSSession $Context['ExchangeSession'] -ErrorAction SilentlyContinue
        $Context['ExchangeSession'] = $null
    }
}
#endregion

#region Objects from the remote session ------------------------------------------------------------

function ConvertTo-PcmDistinguishedName {
    <#
    .SYNOPSIS
        The distinguished name of a reference to a directory object as an Exchange cmdlet returns it: a string, an ADObjectId, or the
        deserialized copy of one that a remote session returns (its properties are kept, its methods are lost). When no distinguished
        name is there, the text of the reference is returned.
    #>
    [CmdletBinding()]
    param([AllowNull()][object]$Reference)
    if ($null -eq $Reference) { return '' }
    if ($Reference -is [string]) { return $Reference.Trim() }
    $property = $Reference.PSObject.Properties['DistinguishedName']
    if ($property -and $null -ne $property.Value -and ([string]$property.Value).Trim()) { return ([string]$property.Value).Trim() }
    return ([string]$Reference).Trim()
}

function ConvertTo-PcmPlainValue {
    <#
    .SYNOPSIS
        A value of an Exchange object as plain data: text, number, boolean, or a list of those. An enumeration, a size, an address or any
        other object of Exchange becomes its text, so that the JSON of the snapshot is the same for a live object and for its deserialized copy
        (the deserialized copy of an enumeration is a number that prints as its name: it must become the name). A date is written as UTC text
        (round-trip format), a duration and a GUID as text.
    #>
    [CmdletBinding()]
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [string] -or $Value -is [bool]) { return $Value }
    if ($Value -is [Collections.IEnumerable]) { return @($Value | ForEach-Object { ConvertTo-PcmPlainValue $_ }) }
    if ($Value -is [enum] -or $Value -is [guid] -or [string]$Value.PSObject.TypeNames[0] -like 'Deserialized.*') { return [string]$Value }
    if ($Value -is [timespan]) { return $Value.ToString('c', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [ValueType]) { return $Value }
    return [string]$Value
}
#endregion

#region Mailboxes ----------------------------------------------------------------------------------

function Get-PcmRecipientType {
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

function Get-PcmGroupMemberDn {
    <#
    .SYNOPSIS
        Distinguished names of the members of a group, recursively (nested groups expanded, loops ignored).
        Returns user/contact members only; the nested groups are walked, not returned.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Identity)
    $dc = Get-PcmDcParameter $Context
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $listed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $result = [System.Collections.Generic.List[string]]::new()
    $queue = [System.Collections.Generic.Queue[string]]::new()
    $queue.Enqueue($Identity)
    while ($queue.Count) {
        $current = $queue.Dequeue()
        if (-not $seen.Add($current)) { continue }
        $group = Get-Group -Identity $current -ErrorAction Stop @dc
        foreach ($member in @($group.Members)) {
            $dn = ConvertTo-PcmDistinguishedName $member
            if (-not $dn) { continue }
            $isGroup = $false
            try { $null = Get-Group -Identity $dn -ErrorAction Stop @dc; $isGroup = $true }
            catch {
                if (-not (Test-PcmNotFoundError -ErrorRecord $_)) { throw ('the member {0} of group {1} could not be read to tell whether it is a group: {2}' -f $dn, $current, $_.Exception.Message.Trim()) }
                $isGroup = $false
            }
            if ($isGroup) { $queue.Enqueue($dn) } elseif ($listed.Add($dn)) { [void]$result.Add($dn) }
        }
    }
    return $result.ToArray()
}

function Test-PcmNotFoundError {
    <# Is this error the answer "there is no such object" (Exchange: ManagementObjectNotFoundException), as opposed to a read that failed (session, network, server busy)? #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$ErrorRecord)
    $text = '{0} {1}' -f $ErrorRecord.FullyQualifiedErrorId, $ErrorRecord.Exception.Message
    return $text -match 'ManagementObjectNotFound|couldn.t be found|could not be found|no such object|\b(group|object|mailbox|recipient|user)\b[^.]*\b(not found|does not exist)'
}

function Get-PcmLdapPath {
    <# The LDAP path of a distinguished name for ADSI: the "/" of a name is escaped (ADSI reads an unescaped "/" as the end of the name), the domain controller is named when one is set. #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][string]$DistinguishedName, [AllowEmptyString()][string]$DomainController = '')
    $escaped = $DistinguishedName.Replace('/', '\/')
    if ($DomainController) { return "LDAP://$DomainController/$escaped" }
    return "LDAP://$escaped"
}

function Get-PcmDirectoryObject {
    <#
    .SYNOPSIS
        One Active Directory object read with ADSI, whatever its class (user, computer, contact, group): class, names, GUID. $null when it cannot be read.
    .PARAMETER Strict
        $null only when the object does not exist; any other failure of the read (the domain controller does not answer) is thrown, so that it is not taken for a missing object.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$DistinguishedName, [switch]$Strict)
    $path = Get-PcmLdapPath -DistinguishedName $DistinguishedName -DomainController ([string]$Context.Config.Exchange.DomainController)
    $entry = $null
    try {
        $entry = New-Object System.DirectoryServices.DirectoryEntry($path)
        $read = { param($name) $values = @($entry.Properties[$name]); if ($values.Count -and $null -ne $values[0]) { [string]$values[0] } else { '' } }
        $classes = @($entry.Properties['objectClass'] | ForEach-Object { [string]$_ })
        $bytes = @($entry.Properties['objectGUID'])
        $class = if ($classes -contains 'computer') { 'Computer' } elseif ($classes -contains 'group') { 'Group' } elseif ($classes -contains 'contact') { 'Contact' }
            elseif ($classes -contains 'user' -or $classes -contains 'inetOrgPerson') { 'User' } elseif ($classes.Count) { [string]$classes[-1] } else { 'Unknown' }
        return [pscustomobject]@{
            Class = $class; Name = (& $read 'name'); Sam = (& $read 'sAMAccountName'); Upn = (& $read 'userPrincipalName'); Smtp = (& $read 'mail')
            DisplayName = (& $read 'displayName'); Dn = $DistinguishedName; Guid = $(if ($bytes.Count -and $bytes[0] -is [byte[]]) { ([guid]$bytes[0]).ToString() } else { '' })
        }
    } catch {
        if ($Strict -and -not (Test-PcmNotFoundError -ErrorRecord $_)) { throw }
        return $null
    }
    finally { if ($entry) { try { $entry.PSBase.Dispose() } catch { $null = $_ } } }
}

function ConvertFrom-PcmCanonicalName {
    <#
    .SYNOPSIS
        The distinguished name of a directory object named by its canonical name (contoso.com/Paris/Users/adele), as a remote Exchange
        session writes a reference to an object (ADObjectId). The name is translated by Active Directory (NameTranslate), so containers
        and organizational units are told apart. '' when Active Directory cannot translate it.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$CanonicalName)
    $dc = [string]$Context.Config.Exchange.DomainController
    try {
        $translate = New-Object -ComObject NameTranslate
        $type = $translate.GetType()
        if ($dc) { [void]$type.InvokeMember('Init', 'InvokeMethod', $null, $translate, @(2, $dc)) }
        else { [void]$type.InvokeMember('Init', 'InvokeMethod', $null, $translate, @(3, '')) }
        [void]$type.InvokeMember('Set', 'InvokeMethod', $null, $translate, @(2, $CanonicalName))
        return [string]$type.InvokeMember('Get', 'InvokeMethod', $null, $translate, @(1))
    } catch { return '' }
}

function Resolve-PcmMemberDn {
    <#
    .SYNOPSIS
        The distinguished name of a group member as Exchange returns it: a distinguished name is kept, a canonical name (the form of a remote
        session) is translated. '' when it cannot be translated.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [AllowNull()][object]$Member)
    $text = ConvertTo-PcmDistinguishedName $Member
    if (-not $text) { return '' }
    if ($text -match '^(?i)(CN|OU|DC|O|L)=') { return $text }
    if ($text -match '/') { return (ConvertFrom-PcmCanonicalName -Context $Context -CanonicalName $text) }
    return $text
}

function Get-PcmMemberLabel {
    <# The name of a group member for a message: its name, else its distinguished name. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$DistinguishedName)
    $object = Get-PcmDirectoryObject -Context $Context -DistinguishedName $DistinguishedName
    if ($object -and $object.Name) { return [string]$object.Name }
    return $DistinguishedName
}

function Get-PcmNoMailboxReason {
    <# Why a member of the scope group is not collected when it has no mailbox: a computer, a contact, or a user without mailbox. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$DistinguishedName)
    $object = Get-PcmDirectoryObject -Context $Context -DistinguishedName $DistinguishedName
    switch ($(if ($object) { [string]$object.Class } else { '' })) {
        'Computer' { return 'a computer, not a mailbox' }
        'Contact' { return 'a contact, not a mailbox' }
        'User' { return 'a user without an Exchange mailbox' }
        default { return 'not a mailbox' }
    }
}

function Get-PcmScopeMailbox {
    <#
    .SYNOPSIS
        The mailboxes to collect: -Identity, or the configured scope (Auto, OU, Group, Csv), sorted by sAMAccountName.
    .OUTPUTS
        Exchange mailbox objects. Excluded objects are reported in $Context.Excluded.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [string]$Identity, [ValidateSet('All','UsersOnly','SharedOnly')][string]$Filter = 'All')
    $scope = $Context.Config.Scope
    $types = @(Get-PcmRecipientType -Scope $scope -Filter $Filter)
    if (-not $types.Count) { throw "No mailbox type left after -Scope $Filter and the configuration (Scope.Include*)." }
    $dc = Get-PcmDcParameter $Context
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
                foreach ($dn in (Get-PcmGroupMemberDn -Context $Context -Identity $scope.GroupDN)) {
                    $mailbox = $null
                    try { $mailbox = Get-Mailbox -Identity $dn -ErrorAction Stop @dc }
                    catch {
                        if (-not (Test-PcmNotFoundError -ErrorRecord $_)) { throw ("Collect stopped before writing the snapshot: the member {0} of the group {1} could not be read ({2}). Run Collect again; a member that cannot be read is never taken for a member without mailbox." -f $dn, $scope.GroupDN, $_.Exception.Message.Trim()) }
                    }
                    if ($mailbox -and [string]$mailbox.RecipientTypeDetails -in $types) { $mailboxes += $mailbox; continue }
                    if ($mailbox) { [void]$Context.Excluded.Add(('{0}: {1} not in scope' -f $mailbox.Name, $mailbox.RecipientTypeDetails)); continue }
                    [void]$Context.Excluded.Add(('{0}: {1}' -f (Get-PcmMemberLabel -Context $Context -DistinguishedName $dn), (Get-PcmNoMailboxReason -Context $Context -DistinguishedName $dn)))
                }
            }
            'Csv' {
                $csvText = Read-PcmCsvText -Path $scope.CsvPath
                if ([string]::IsNullOrWhiteSpace($csvText)) { throw "Scope.CsvPath: the file is empty ($($scope.CsvPath))." }
                $csvDelimiter = Get-PcmCsvDelimiter (($csvText -split '\r?\n', 2)[0])
                if (-not $csvDelimiter) { throw "Scope.CsvPath: column Identity missing, or the columns are not separated by a comma or a semicolon ($($scope.CsvPath))." }
                $entries = @(ConvertFrom-Csv -InputObject $csvText -Delimiter $csvDelimiter -ErrorAction Stop)
                foreach ($entry in $entries) {
                    $identityText = ([string]$entry.Identity).Trim()
                    if (-not $identityText) { continue }
                    $mailbox = Get-Mailbox -Identity $identityText -ErrorAction Stop @dc
                    if ([string]$mailbox.RecipientTypeDetails -in $types) { $mailboxes += $mailbox }
                    else { [void]$Context.Excluded.Add(('{0}: {1} not in scope' -f $identityText, $mailbox.RecipientTypeDetails)) }
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
        if (Test-PcmSystemMailboxName ([string]$mailbox.Name)) { [void]$Context.Excluded.Add(('{0}: system mailbox' -f $mailbox.Name)); continue }
        if ([string]$mailbox.SamAccountName -in $excludedNames) { [void]$Context.Excluded.Add(('{0}: Scope.ExcludeSamAccountNames' -f $mailbox.SamAccountName)); continue }
        $kept += $mailbox
    }
    return @($kept | Sort-Object { [string]$_.SamAccountName })
}

function ConvertTo-PcmImmutableId {
    <# Base64 of the objectGUID: the onPremisesImmutableId of the Entra ID object when the source anchor is the objectGUID or ms-DS-ConsistencyGuid. #>
    param([Parameter(Mandatory)][object]$Guid)
    return [Convert]::ToBase64String(([guid][string]$Guid).ToByteArray())
}

function ConvertTo-PcmSizeBytes {
    <#
    .SYNOPSIS
        A size as Exchange returns it -> bytes; $null when it is unknown or unlimited. A live ByteQuantifiedSize has ToBytes(); the
        deserialized copy of a remote session has lost its methods and keeps only its text, '1.2 GB (1,288,490,188 bytes)'. A number is taken as bytes.
    #>
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [uint32] -or $Value -is [uint64] -or $Value -is [double] -or $Value -is [decimal]) { return [long]$Value }
    if ($Value.PSObject.Methods['ToBytes']) { return [long]$Value.ToBytes() }
    $text = [string]$Value
    if ($text -match '\(([\d,\.\s\u00A0\u202F]+)\s*bytes\)') { return [long]($Matches[1] -replace '[^\d]', '') }
    return $null
}

function ConvertTo-PcmMailboxRecord {
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
        hidden_from_address_lists = [bool]$Mailbox.HiddenFromAddressListsEnabled; immutable_id = (ConvertTo-PcmImmutableId $Mailbox.Guid)
        database = [string]$Mailbox.Database; email_addresses_json = (ConvertTo-Json -InputObject @($addresses) -Compress)
        custom_attributes_json = ($attributes | ConvertTo-Json -Compress); grant_send_on_behalf_json = (ConvertTo-Json -InputObject @($onBehalf) -Compress)
        item_count = $(if ($Statistics) { [long]$Statistics.ItemCount } else { $null })
        total_item_size_bytes = $(if ($Statistics) { ConvertTo-PcmSizeBytes $Statistics.TotalItemSize } else { $null })
        resource_settings_json = $(if ($ResourceSettings) { ConvertTo-Json -InputObject $ResourceSettings -Depth 6 -Compress } else { $null })
    }
}

function Get-PcmResourceSettings {
    <# A portable booking snapshot: recipient references are SMTP addresses, not AD DNs. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Mailbox)
    if ([string]$Mailbox.RecipientTypeDetails -notin @('RoomMailbox','EquipmentMailbox')) { throw 'Resource settings require a room or equipment mailbox.' }
    $identity = [string]$Mailbox.DistinguishedName
    $dc = Get-PcmDcParameter $Context
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
    $delivery = Get-PcmDeliverySettings -Context $Context -Recipient $Mailbox
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
            $calendar[$name] = ConvertTo-PcmPlainValue $processing.$name
        }
    }
    foreach ($name in @('ResourceDelegates','BookInPolicy','RequestInPolicy','RequestOutOfPolicy')) {
        $recipients = [System.Collections.Generic.List[string]]::new()
        $recipientSeen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($reference in @($processing.$name)) {
            if (-not $reference) { continue }
            $trustee = Resolve-PcmTrustee -Context $Context -Trustee ([string]$reference)
            if (-not $trustee.Resolved -or -not $trustee.Smtp) {
                throw "Resource settings: $name trustee '$reference' on $($Mailbox.PrimarySmtpAddress) has no resolved SMTP address. Correct it before Collect."
            }
            if ($recipientSeen.Add([string]$trustee.Smtp)) { $recipients.Add([string]$trustee.Smtp) }
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
    $roomLists = @(if ([string]$Mailbox.RecipientTypeDetails -eq 'RoomMailbox') { Get-PcmResourceRoomLists -Context $Context -Mailbox $Mailbox })
    return @{ SchemaVersion = 2; Capacity = [int]$Mailbox.ResourceCapacity
        ResourceCustom = [string[]]@($Mailbox.ResourceCustom | ForEach-Object { [string]$_ }); Calendar = $calendar; CalendarConfiguration = $hours
        Directory = $directory; RoomLists = $roomLists; Delivery = $delivery }
}

function Get-PcmDeliverySettings {
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
                $trustee = Resolve-PcmTrustee -Context $Context -Trustee ([string]$reference)
                if (-not $trustee.Resolved -or -not $trustee.Smtp) { throw "Resource delivery restriction $name has an unresolved recipient '$reference' on $($Recipient.PrimarySmtpAddress)." }
                [string]$trustee.Smtp
            })
        $delivery[$name] = @($values | Get-PcmDistinctText)
    }
    if ($delivery.SendModerationNotifications -notin @('Never','Internal','Always') -or ($delivery.ModerationEnabled -and -not $delivery.ModeratedBy.Count)) {
        throw "Resource delivery restriction moderation is incomplete on $($Recipient.PrimarySmtpAddress)."
    }
    return $delivery
}

function Get-PcmResourceRoomLists {
    param([hashtable]$Context, [object]$Mailbox)
    if (-not $Context.ContainsKey('ResourceRoomListInventory')) {
        $dc = Get-PcmDcParameter $Context
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
                        Delivery = (Get-PcmDeliverySettings -Context $Context -Recipient $group) } }
            })
        $Context['ResourceRoomListInventory'] = $inventory
    }
    foreach ($entry in $Context.ResourceRoomListInventory) {
        if ([string]$Mailbox.Guid -in $entry.MemberGuids) { $entry.Metadata }
    }
}
#endregion

#region Target groups --------------------------------------------------------------------------------

function Find-PcmExchangeGroup {
    <#
    .SYNOPSIS
        A group named by its name, sAMAccountName, distinguished name, GUID or SMTP address. $null when there is none; an error when the value names several groups.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Value)
    $dc = Get-PcmDcParameter $Context
    $group = @()
    try { $group = @(Get-Group -Identity $Value -ErrorAction Stop @dc) } catch { $group = @() }
    if (-not $group.Count) {
        $recipient = @()
        try { $recipient = @(Get-Recipient -Identity $Value -ErrorAction Stop @dc) } catch { $recipient = @() }
        $recipient = @($recipient | Where-Object { [string]$_.RecipientType -match 'Group' })
        if ($recipient.Count -eq 1) { try { $group = @(Get-Group -Identity ([string]$recipient[0].DistinguishedName) -ErrorAction Stop @dc) } catch { $group = @() } }
        elseif ($recipient.Count -gt 1) { throw ('{0} names {1} groups: use the distinguished name' -f $Value, $recipient.Count) }
    }
    if ($group.Count -gt 1) { throw ('{0} names {1} groups: use the distinguished name' -f $Value, $group.Count) }
    if ($group.Count -eq 1) { return $group[0] }
    return $null
}

function Get-PcmTargetGroup {
    <#
    .SYNOPSIS
        The groups to capture as run targets (Collect.TargetGroups and Scope.GroupDN of the Group mode) with ALL their members, nested groups
        walked, loops ignored: user, computer, contact and group members are stored with their class so that a run can skip what is not a mailbox
        and say why.
    .OUTPUTS
        @{ Groups = rows of target_group; Members = rows of target_member; Warnings = texts (a group that is not found does not stop Collect) }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = Get-PcmDcParameter $Context
    $config = $Context.Config
    $wanted = [System.Collections.Generic.List[string]]::new()
    foreach ($value in @(@(Get-PcmValue $config.Collect 'TargetGroups' @()) | ForEach-Object { $_ })) { if ([string]$value) { $wanted.Add([string]$value) } }
    if ((Test-PcmSameText (Get-PcmValue $config.Scope 'Mode' '') 'Group') -and (Get-PcmValue $config.Scope 'GroupDN' '')) { $wanted.Add([string](Get-PcmValue $config.Scope 'GroupDN' '')) }
    $groups = [System.Collections.Generic.List[object]]::new()
    $members = [System.Collections.Generic.List[object]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()
    $done = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($value in $wanted) {
        $found = $null
        try { $found = Find-PcmExchangeGroup -Context $Context -Value $value }
        catch { [void]$warnings.Add(('Target group {0} not captured: {1}.' -f $value, $_.Exception.Message)); continue }
        if (-not $found) { [void]$warnings.Add(('Target group {0} not found in Active Directory: it is not captured (a run will read it from Microsoft Entra ID).' -f $value)); continue }
        $topGuid = ([string]$found.Guid).ToLowerInvariant()
        if (-not $done.Add($topGuid)) { continue }
        $smtp = ''
        try { $smtp = [string](Get-Recipient -Identity ([string]$found.DistinguishedName) -ErrorAction Stop @dc).PrimarySmtpAddress } catch { $smtp = '' }
        $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $listed = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $queue = [System.Collections.Generic.Queue[object]]::new()
        $queue.Enqueue([pscustomobject]@{ Source = $found; Guid = $topGuid; Depth = 1 })
        [void]$seen.Add($topGuid)
        $rows = [System.Collections.Generic.List[object]]::new()
        $unreadable = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $unreadableList = [System.Collections.Generic.List[string]]::new()
        $count = 0
        try {
            while ($queue.Count) {
                $current = $queue.Dequeue()
                $source = $current.Source
                if (-not $source.PSObject.Properties['Members']) { $source = Get-Group -Identity ([string]$source.DistinguishedName) -ErrorAction Stop @dc }
                foreach ($member in @($source.Members)) {
                    if (-not (ConvertTo-PcmDistinguishedName $member)) { continue }
                    $dn = Resolve-PcmMemberDn -Context $Context -Member $member
                    if (-not $dn) { throw ('the member {0} cannot be found in Active Directory' -f (ConvertTo-PcmDistinguishedName $member)) }
                    $object = Get-PcmDirectoryObject -Context $Context -DistinguishedName $dn -Strict
                    $class = if ($object) { [string]$object.Class } else { 'Unknown' }
                    $guid = if ($object) { ([string]$object.Guid).ToLowerInvariant() } else { '' }
                    if (-not $object -and $unreadable.Add($dn)) { $unreadableList.Add($dn) }
                    if ($class -eq 'Group' -and $guid) {
                        if ($seen.Add($guid)) {
                            $queue.Enqueue([pscustomobject]@{ Source = [pscustomobject]@{ DistinguishedName = $dn }; Guid = $guid; Depth = ($current.Depth + 1) })
                            [void]$rows.Add([ordered]@{ group_guid = $topGuid; member_guid = $guid; member_class = 'Group'; member_name = [string]$object.Name; member_upn = $null
                                member_smtp = $(if ($object.Smtp) { [string]$object.Smtp } else { $null }); member_dn = $dn; via_group_guid = $current.Guid; depth = $current.Depth })
                        }
                        continue
                    }
                    $key = if ($guid) { $guid } else { $dn }
                    if (-not $listed.Add($key)) { continue }
                    $count++
                    [void]$rows.Add([ordered]@{ group_guid = $topGuid; member_guid = $(if ($guid) { $guid } else { $null }); member_class = $class
                        member_name = $(if ($object) { [string]$object.Name } else { $dn }); member_upn = $(if ($object -and $object.Upn) { [string]$object.Upn } else { $null })
                        member_smtp = $(if ($object -and $object.Smtp) { [string]$object.Smtp } else { $null }); member_dn = $dn; via_group_guid = $current.Guid; depth = $current.Depth })
                }
            }
        } catch {
            [void]$warnings.Add(('Target group {0}: its members could not all be read ({1}); the group is not captured.' -f $value, $_.Exception.Message))
            continue
        }
        if ($unreadableList.Count) {
            [void]$warnings.Add(('Target group {0}: {1} member(s) do not exist in Active Directory any more and are stored as Unknown, which a run skips ({2}{3}).' -f $value, $unreadableList.Count, (($unreadableList | Select-Object -First 3) -join '; '), $(if ($unreadableList.Count -gt 3) { '; ...' } else { '' })))
        }
        foreach ($row in $rows) { [void]$members.Add($row) }
        [void]$groups.Add([ordered]@{ group_guid = $topGuid; name = [string]$found.Name; sam_account_name = [string]$found.SamAccountName
            distinguished_name = [string]$found.DistinguishedName; primary_smtp_address = $(if ($smtp) { $smtp } else { $null }); display_name = [string]$found.DisplayName; member_count = $count })
    }
    return @{ Groups = $groups.ToArray(); Members = $members.ToArray(); Warnings = $warnings.ToArray() }
}
#endregion

#region Shared mailbox permissions -------------------------------------------------------------------

function Test-PcmExcludedTrustee {
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

function Resolve-PcmTrustee {
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
    $dc = Get-PcmDcParameter $Context
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

function Get-PcmDelegateListLink {
    <#
    .SYNOPSIS
        msExchDelegateListLink of a mailbox (the FullAccess trustees with AutoMapping), read in AD with ADSI.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$DistinguishedName)
    $path = Get-PcmLdapPath -DistinguishedName $DistinguishedName -DomainController ([string]$Context.Config.Exchange.DomainController)
    $entry = New-Object System.DirectoryServices.DirectoryEntry($path)
    try { return @($entry.Properties['msExchDelegateListLink'] | ForEach-Object { [string]$_ }) }
    finally { $entry.Dispose() }
}

function Get-PcmSharedPermission {
    <#
    .SYNOPSIS
        FullAccess, SendAs and SendOnBehalf of one mailbox, trustees resolved, groups expanded.
    .OUTPUTS
        @{ Permissions = rows of the permission table; Members = rows of the group_member table; Warnings = texts }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$Mailbox)
    $dc = Get-PcmDcParameter $Context
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
        $text = ConvertTo-PcmDistinguishedName $delegate
        if ($text) { [void]$entries.Add(@{ Right = 'SendOnBehalf'; Trustee = $text }) }
    }
    $autoMapping = @{}
    if (@($entries | Where-Object { $_.Right -eq 'FullAccess' }).Count) {
        try { foreach ($dn in (Get-PcmDelegateListLink -Context $Context -DistinguishedName $identity)) { $autoMapping[$dn.ToLowerInvariant()] = $true } }
        catch { [void]$warnings.Add(('{0}: AutoMapping not read ({1})' -f $Mailbox.PrimarySmtpAddress, $_.Exception.Message)) }
    }
    $permissions = [System.Collections.Generic.List[object]]::new()
    $members = [System.Collections.Generic.List[object]]::new()
    $expanded = @{}
    $seenPair = @{}
    foreach ($entry in $entries) {
        $trustee = $entry.Trustee
        if (Test-PcmExcludedTrustee -Context $Context -Trustee $trustee) { continue }
        $resolved = Resolve-PcmTrustee -Context $Context -Trustee $trustee
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
                foreach ($memberDn in (Get-PcmGroupMemberDn -Context $Context -Identity $resolved.Dn)) {
                    $member = Resolve-PcmTrustee -Context $Context -Trustee $memberDn
                    if (-not $member.Resolved) { [void]$warnings.Add(('{0}: member {1} of group {2} could not be read ({3}); it is stored as Unknown' -f $Mailbox.PrimarySmtpAddress, $memberDn, $resolved.Name, $member.Note)) }
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

function Get-PcmCustomAttributeJson {
    param([Parameter(Mandatory)][object]$Object)
    $attributes = [ordered]@{}
    for ($i = 1; $i -le 15; $i++) {
        $name = "CustomAttribute$i"
        if ($Object.PSObject.Properties[$name] -and [string]$Object.$name) { $attributes[$name] = [string]$Object.$name }
    }
    return ($attributes | ConvertTo-Json -Compress)
}

function Get-PcmContact {
    <# Mail contacts (optionally under Collect.ContactsSearchBase) -> rows of the contact table. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = Get-PcmDcParameter $Context
    $ou = @{}; if ($Context.Config.Collect.ContactsSearchBase) { $ou = @{ OrganizationalUnit = $Context.Config.Collect.ContactsSearchBase } }
    $rows = foreach ($contact in @(Get-MailContact -ResultSize Unlimited -ErrorAction Stop @ou @dc)) {
        [ordered]@{
            object_guid = [string]$contact.Guid; name = [string]$contact.Name; display_name = [string]$contact.DisplayName; alias = [string]$contact.Alias
            external_email_address = [string]$contact.ExternalEmailAddress; primary_smtp_address = [string]$contact.PrimarySmtpAddress
            legacy_exchange_dn = [string]$contact.LegacyExchangeDN; hidden_from_address_lists = [bool]$contact.HiddenFromAddressListsEnabled
            organizational_unit = [string]$contact.OrganizationalUnit
            email_addresses_json = (ConvertTo-Json -InputObject @($contact.EmailAddresses | ForEach-Object { [string]$_ }) -Compress)
            custom_attributes_json = (Get-PcmCustomAttributeJson $contact)
        }
    }
    return @($rows)
}

function Get-PcmDistributionGroup {
    <#
    .SYNOPSIS
        Distribution and mail-enabled security groups (and dynamic groups when enabled), with their members.
    .OUTPUTS
        @{ Groups = rows of distribution_group; Members = rows of dl_member }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $dc = Get-PcmDcParameter $Context
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

#region The daily task ---------------------------------------------------------------------------------
# What -Action Schedule needs: the settings of the task, the checks of the account, the registration. The steps and the messages of the run are in the entry script.

$script:SchedulePrefixes = @('S-1-5-80-', 'S-1-5-82-', 'S-1-5-83-', 'S-1-5-90-', 'S-1-5-96-')

function Test-PcmElevated {
    <# True when this PowerShell runs with administrator rights (a scheduled task cannot be registered without them). #>
    [CmdletBinding()]
    param()
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    return [Security.Principal.WindowsPrincipal]::new($identity).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Test-PcmInteractiveHost {
    <# False when this PowerShell was started with -NonInteractive: it cannot ask anything. #>
    [CmdletBinding()]
    param()
    foreach ($argument in [Environment]::GetCommandLineArgs()) { if ($argument -match '^[-/]noni') { return $false } }
    return $true
}

function Test-PcmScheduleWindow {
    <# Elevated and Interactive: what this PowerShell window can do for the registration. #>
    [CmdletBinding()]
    param()
    return [pscustomobject]@{ Elevated = [bool](Test-PcmElevated); Interactive = [bool](Test-PcmInteractiveHost) }
}

function Hide-PcmPassword {
    <# The text without the password of the credential: a message that could repeat it never leaves this module as it is. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Text, [AllowNull()][pscredential]$Credential)
    if (-not $Text -or $null -eq $Credential) { return $Text }
    $password = $Credential.GetNetworkCredential().Password
    if (-not $password) { return $Text }
    return $Text.Replace($password, '********')
}

function Get-PcmScheduleSetting {
    <#
    .SYNOPSIS
        What the daily task is: its name, its folder, its time and its command (the PowerShell that runs now, this script, Collect in Apply mode).
    .PARAMETER At
        The time typed on the command line (24 hours, HH:mm); empty = Collect.Schedule.At of the configuration.
    .PARAMETER ConfigPath
        The configuration file the task must use when it is not the default one (empty = the default file of the package).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Config, [AllowEmptyString()][string]$At = '', [Parameter(Mandatory)][string]$ScriptPath, [AllowEmptyString()][string]$ConfigPath = '')
    $time = if ($At.Trim()) { ConvertTo-PcmScheduleTime $At } else { ConvertTo-PcmScheduleTime ([string]$Config.Collect.Schedule.At) }
    if (-not $time) { throw ("-At '{0}' is not a time of day. Give it on 24 hours, for example -At 02:00." -f $At) }
    $pwsh = Join-Path $PSHOME 'pwsh.exe'
    $argument = '-NoProfile -NonInteractive -File "{0}" -Action Collect -Mode Apply -Force' -f $ScriptPath
    if ($ConfigPath) { $argument += ' -ConfigPath "{0}"' -f $ConfigPath }
    $parts = $time.Split(':')
    return [pscustomobject]@{
        TaskName = [string]$Config.Collect.Schedule.TaskName; TaskPath = '\'; At = $time; Start = [datetime]::Today.AddHours([int]$parts[0]).AddMinutes([int]$parts[1])
        Execute = $pwsh; Argument = $argument; WorkingDirectory = [IO.Path]::GetDirectoryName($ScriptPath)
    }
}

function Get-PcmDirectoryAccount {
    <#
    .SYNOPSIS
        What Windows and Active Directory say about an account: its SID, its canonical name, its classes in the directory and the SIDs of its groups. Nothing is changed.
    .DESCRIPTION
        Found, Sid, Name (DOMAIN\name), Source (WellKnown, Local or Directory), Classes (the objectClass values), Groups and GroupsKnown (the SIDs of the token groups), Problem.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    $result = [pscustomobject]@{ Found = $false; Sid = ''; Name = ''; Source = 'Directory'; Classes = @(); Groups = @(); GroupsKnown = $false; Problem = '' }
    try {
        $sid = if ($Name -match '^S-1-[0-9-]+$') { [Security.Principal.SecurityIdentifier]::new($Name) } else { [Security.Principal.NTAccount]::new($Name).Translate([Security.Principal.SecurityIdentifier]) }
        $account = [string]$sid.Translate([Security.Principal.NTAccount]).Value
    } catch { $result.Problem = $_.Exception.Message; return $result }
    $result.Found = $true; $result.Sid = $sid.Value; $result.Name = $account
    if ($sid.Value -notmatch '^S-1-5-21-') { $result.Source = 'WellKnown'; return $result }
    if (Test-PcmSameText ($account.Split('\')[0]) $env:COMPUTERNAME) { $result.Source = 'Local'; return $result }
    try {
        $roots = [System.Collections.Generic.List[string]]::new()
        $roots.Add('')
        try { $rootDse = [adsi]'LDAP://RootDSE'; $roots.Add('GC://' + [string]$rootDse.Properties['rootDomainNamingContext'][0]) } catch { $null = $_ }
        $found = $null
        foreach ($root in $roots) {
            $searcher = if ($root) { [System.DirectoryServices.DirectorySearcher]::new([System.DirectoryServices.DirectoryEntry]::new($root)) } else { [System.DirectoryServices.DirectorySearcher]::new() }
            $searcher.Filter = '(objectSid={0})' -f $sid.Value
            [void]$searcher.PropertiesToLoad.Add('objectClass')
            $found = $searcher.FindOne()
            if ($found) { break }
        }
        if (-not $found) { $result.Found = $false; $result.Problem = 'The account is not in Active Directory.'; return $result }
        $result.Classes = @($found.Properties['objectclass'] | ForEach-Object { [string]$_ })
        try {
            $entry = $found.GetDirectoryEntry()
            $entry.RefreshCache(@('tokenGroups'))
            $result.Groups = @($entry.Properties['tokenGroups'] | ForEach-Object { [Security.Principal.SecurityIdentifier]::new([byte[]]$_, 0).Value })
            $result.GroupsKnown = $true
        } catch { $result.GroupsKnown = $false }
    } catch { $result.Found = $false; $result.Problem = $_.Exception.Message }
    return $result
}

function Get-PcmScheduleAccount {
    <#
    .SYNOPSIS
        Tells whether an account can run the daily Collect: it must be a domain user. SYSTEM, LOCAL SERVICE, NETWORK SERVICE, service identities, local accounts, computer accounts and
        managed service accounts are refused, with the reason in plain words (Exchange remote PowerShell refuses them).
    .OUTPUTS
        Ok, Typed, Name (DOMAIN\name), Sid, Groups, GroupsKnown, Message (the reason when it is not Ok).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)
    $typed = $Name.Trim()
    $result = [pscustomobject]@{ Ok = $false; Typed = $typed; Name = $typed; Sid = ''; Groups = @(); GroupsKnown = $false; Message = '' }
    $account = Get-PcmDirectoryAccount -Name $typed
    if (-not $account.Found) {
        $result.Message = "The account '{0}' was not found in the domain. Give it as DOMAIN\name or name@domain, for a domain user." -f $typed
        return $result
    }
    $result.Name = [string]$account.Name; $result.Sid = [string]$account.Sid; $result.Groups = @($account.Groups); $result.GroupsKnown = [bool]$account.GroupsKnown
    $shown = [string]$account.Name
    $sid = [string]$account.Sid
    $service = ($sid -in @('S-1-5-18', 'S-1-5-19', 'S-1-5-20')) -or (@($script:SchedulePrefixes | Where-Object { $sid.StartsWith($_, [StringComparison]::OrdinalIgnoreCase) }).Count -gt 0)
    $advice = 'Use a standard domain user account that is a member of View-Only Organization Management.'
    if ($service) { $result.Message = '{0} is a built-in identity of this computer, not a domain user. Exchange refuses it, so it cannot run Collect. {1}' -f $shown, $advice }
    elseif (Test-PcmSameText $account.Source 'Local') { $result.Message = '{0} is a local account of this computer. Exchange needs a domain user, so it cannot run Collect. {1}' -f $shown, $advice }
    elseif (@($account.Classes) -contains 'msDS-GroupManagedServiceAccount') { $result.Message = '{0} is a group managed service account (gMSA). Exchange remote PowerShell refuses it, so it cannot run Collect. {1}' -f $shown, $advice }
    elseif (@($account.Classes) -contains 'msDS-ManagedServiceAccount') { $result.Message = '{0} is a managed service account. Exchange remote PowerShell refuses it, so it cannot run Collect. {1}' -f $shown, $advice }
    elseif (@($account.Classes) -contains 'computer') { $result.Message = '{0} is a computer account. Exchange remote PowerShell refuses it, so it cannot run Collect. {1}' -f $shown, $advice }
    elseif (@($account.Classes) -contains 'user') { $result.Ok = $true }
    else { $result.Message = '{0} is not a user account of the domain. {1}' -f $shown, $advice }
    return $result
}

function Test-PcmExchangeAccess {
    <#
    .SYNOPSIS
        Opens the remote session to Exchange.Server with the credential (Kerberos), reads one recipient and closes the session: the proof that the account may run Collect.
    .OUTPUTS
        Ok, Server, Message (what it proved, or the reason in plain words).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Server, [Parameter(Mandatory)][pscredential]$Credential)
    $uri = 'http://{0}/PowerShell/' -f $Server
    $session = $null
    try {
        $session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $uri -Authentication Kerberos -Credential $Credential -ErrorAction Stop
        $null = Invoke-Command -Session $session -ScriptBlock { Get-Recipient -ResultSize 1 -ErrorAction Stop } -ErrorAction Stop
        return [pscustomobject]@{ Ok = $true; Server = $Server; Message = ('Exchange on {0} accepted the account (Kerberos) and let it read recipients.' -f $Server) }
    } catch {
        $text = (Hide-PcmPassword -Text ([string]$_.Exception.Message) -Credential $Credential).Trim()
        $message = if ($text -match "isn.t assigned to any management roles") {
            'Exchange knows the account, but it is not in any management role. Add it to the Exchange role group View-Only Organization Management, wait a few minutes for the change to replicate, then run this again.'
        } elseif ($text -match "couldn.t be found") {
            'Exchange does not know this account as a user. Use a domain user account, not a computer account or SYSTEM. If the account is new, or was just added to View-Only Organization Management, wait a few minutes for the replication and run this again.'
        } elseif ($text -match 'is not recognized as|is not recognised as|not allowed to run') {
            'Exchange let the account in, but it may not read recipients. Add it to the Exchange role group View-Only Organization Management, wait a few minutes for the change to replicate, then run this again.'
        } elseif ($text -match 'user name or password is incorrect|logon failure|0x8009030c|0x80090311|0x8009030e|Access is denied') {
            'Exchange refused the sign-in of this account. Check the password, and that the account is a domain user that is not disabled or locked, then run this again.'
        } else {
            'Cannot open the remote session to Exchange on {0}: {1} Check that this computer is domain-joined, that Exchange.Server answers on {2} with Kerberos, and that the account is in View-Only Organization Management.' -f $Server, $text, $uri
        }
        return [pscustomobject]@{ Ok = $false; Server = $Server; Message = $message }
    } finally {
        if ($session) { Remove-PSSession -Session $session -ErrorAction SilentlyContinue }
    }
}

function Test-PcmDomainController {
    <# True when this computer is a domain controller (DomainRole 4 or 5). #>
    [CmdletBinding()]
    param()
    try { return ([int](Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop).DomainRole) -in @(4, 5) } catch { return $false }
}

function Export-PcmUserRight {
    <# Writes the user rights of this computer to a text file (secedit); returns the exit code of secedit. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $null = & secedit.exe /export /cfg $Path /areas USER_RIGHTS /quiet
    return [int]$LASTEXITCODE
}

function Test-PcmBatchLogonRight {
    <#
    .SYNOPSIS
        Best effort: does the account have the right "Log on as a batch job" on this domain controller? State: Granted, Missing, Denied or Unknown (Message says why).
    .DESCRIPTION
        The rights are exported to a temporary file in $Folder and the file is deleted at once. The account counts through its own SID and through its groups.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Account, [Parameter(Mandatory)][string]$Folder)
    $unknown = { param($why) [pscustomobject]@{ State = 'Unknown'; Message = ('The right "Log on as a batch job" could not be read ({0}). Check that {1} has it, in the Default Domain Controllers Policy, or the task will not start.' -f $why, $Account.Name) } }
    $file = Join-Path $Folder ('PRA_rights_{0}.inf' -f [guid]::NewGuid().ToString('N'))
    try {
        $code = Export-PcmUserRight -Path $file
        if ($code -ne 0 -or -not (Test-Path -LiteralPath $file -PathType Leaf)) { return (& $unknown 'secedit did not write the rights') }
        $lines = @(Get-Content -LiteralPath $file -ErrorAction Stop)
    } catch { return (& $unknown $_.Exception.Message) }
    finally { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue }
    $read = {
        param($right)
        $line = @($lines | Where-Object { $_ -match ('^\s*{0}\s*=' -f $right) }) | Select-Object -First 1
        if (-not $line) { return $null }
        return @(($line -replace '^[^=]*=', '').Split(',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    $mine = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($sid in @($Account.Sid, 'S-1-1-0', 'S-1-5-11') + @($Account.Groups)) { if ($sid) { [void]$mine.Add([string]$sid) } }
    $listed = {
        param($entries)
        foreach ($entry in @($entries)) {
            $text = ([string]$entry).TrimStart('*')
            if ($mine.Contains($text) -or (Test-PcmSameText $text $Account.Name)) { return $true }
            if (-not ([string]$entry).StartsWith('*', [StringComparison]::OrdinalIgnoreCase)) {
                try { if ($mine.Contains([Security.Principal.NTAccount]::new($text).Translate([Security.Principal.SecurityIdentifier]).Value)) { return $true } } catch { $null = $_ }
            }
        }
        return $false
    }
    $denied = & $read 'SeDenyBatchLogonRight'
    if ($null -ne $denied -and (& $listed $denied)) {
        return [pscustomobject]@{ State = 'Denied'; Message = ('{0} is denied the right "Log on as a batch job" on this domain controller (Deny log on as a batch job). Remove it from that policy, or the task will not start.' -f $Account.Name) }
    }
    $granted = & $read 'SeBatchLogonRight'
    if ($null -eq $granted) {
        if (@($lines | Where-Object { $_ -match '^\s*\[Privilege Rights\]' }).Count) { $granted = @() }
        else { return (& $unknown 'the file holds no list of rights') }
    }
    if (& $listed $granted) { return [pscustomobject]@{ State = 'Granted'; Message = ('{0} has the right "Log on as a batch job" on this domain controller.' -f $Account.Name) } }
    if (-not $Account.GroupsKnown) { return (& $unknown 'the groups of the account could not be read') }
    return [pscustomobject]@{ State = 'Missing'; Message = ('{0} does not have the right "Log on as a batch job" on this domain controller. Add it in the Default Domain Controllers Policy (Computer Configuration, Windows Settings, Security Settings, Local Policies, User Rights Assignment, Log on as a batch job), run gpupdate /force on the domain controller, then run this again.' -f $Account.Name) }
}

function Get-PcmScheduledCollectTask {
    <# The task of this name in the root folder (the letter case of the name is ignored), or $null: its account, its time and its command. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Setting)
    $found = @(Get-ScheduledTask -TaskPath $Setting.TaskPath -ErrorAction SilentlyContinue | Where-Object { Test-PcmSameText $_.TaskName $Setting.TaskName }) | Select-Object -First 1
    if (-not $found) { return $null }
    $action = @($found.Actions) | Select-Object -First 1
    $trigger = @($found.Triggers) | Select-Object -First 1
    $at = ''
    if ($trigger -and $trigger.StartBoundary) {
        $start = [datetime]::MinValue
        if ([datetime]::TryParse([string]$trigger.StartBoundary, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$start)) { $at = $start.ToString('HH:mm', [Globalization.CultureInfo]::InvariantCulture) }
    }
    $command = if ($action) { ('{0} {1}' -f $action.Execute, $action.Arguments).Trim() } else { '' }
    return [pscustomobject]@{ Name = [string]$found.TaskName; Account = [string]$found.Principal.UserId; At = $at; Command = $command }
}

function Register-PcmCollectTask {
    <#
    .SYNOPSIS
        Registers (or replaces) the daily task: this PowerShell, this script with Collect in Apply mode, once a day at the time of the setting, under the account, with its password
        stored by Windows (logon type Password, standard rights), working folder = the package folder, one run at a time, 4 hours at most, started late when the computer was off.
    .DESCRIPTION
        The password is read from the credential for this one call and is never written anywhere: a message that could repeat it is cleaned first.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Setting, [Parameter(Mandatory)][string]$UserName, [Parameter(Mandatory)][pscredential]$Credential)
    $action = New-ScheduledTaskAction -Execute $Setting.Execute -Argument $Setting.Argument -WorkingDirectory $Setting.WorkingDirectory
    $trigger = New-ScheduledTaskTrigger -Daily -At $Setting.Start
    # PowerShell 7 writes the start of this trigger in UTC ("...Z", synchronised across time zones): the task would move by an hour at each change of the clock. The local time, without "Z", keeps the same hour on the wall.
    $trigger.StartBoundary = $Setting.Start.ToString('yyyy-MM-ddTHH:mm:ss', [Globalization.CultureInfo]::InvariantCulture)
    $settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit (New-TimeSpan -Hours 4)
    $password = $Credential.GetNetworkCredential().Password
    try {
        $null = Register-ScheduledTask -TaskName $Setting.TaskName -TaskPath $Setting.TaskPath -Action $action -Trigger $trigger -Settings $settings -User $UserName -Password $password -RunLevel Limited `
            -Description 'Daily PRA Cloud Mailbox Collect: writes the snapshot of the mailboxes from Exchange.' -Force -ErrorAction Stop
    } catch {
        $text = (Hide-PcmPassword -Text ([string]$_.Exception.Message) -Credential $Credential).Trim()
        $why = if ($text -match 'Access is denied|0x80070005') { 'This window is not elevated: open PowerShell 7 with "Run as administrator" and run the same command again.' }
        elseif ($text -match 'user name or password is incorrect|0x8007052e') { 'Windows refused the password of {0}.' -f $UserName }
        elseif ($text -match 'No mapping between account names|0x80070534') { 'Windows does not know the account {0}.' -f $UserName }
        elseif ($text -match 'not been granted the requested logon type|0x80070569') { 'The account {0} does not have the right "Log on as a batch job" here.' -f $UserName }
        else { $text }
        throw ("The task '{0}' could not be registered: {1}" -f $Setting.TaskName, $why)
    } finally { $password = $null }
}
#endregion

Export-ModuleMember -Function Connect-PcmExchange, Disconnect-PcmExchange, Get-PcmRecipientType, Get-PcmScopeMailbox, Get-PcmGroupMemberDn,
    ConvertTo-PcmImmutableId, ConvertTo-PcmSizeBytes, ConvertTo-PcmMailboxRecord, Test-PcmExcludedTrustee, Resolve-PcmTrustee,
    Get-PcmDelegateListLink, Get-PcmSharedPermission, Get-PcmContact, Get-PcmDistributionGroup, Get-PcmResourceSettings, Get-PcmTargetGroup, Find-PcmExchangeGroup, Get-PcmDirectoryObject, Get-PcmLdapPath, Test-PcmNotFoundError,
    Test-PcmElevated, Test-PcmInteractiveHost, Test-PcmScheduleWindow, Hide-PcmPassword, Get-PcmScheduleSetting, Get-PcmDirectoryAccount, Get-PcmScheduleAccount, Test-PcmExchangeAccess,
    Test-PcmDomainController, Export-PcmUserRight, Test-PcmBatchLogonRight, Get-PcmScheduledCollectTask, Register-PcmCollectTask

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC52bOlZM83tRdc
# YkfKqM6XvS1Vi0LD8lU4/AwlX0d9PqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAHNvAqzIsa
# ewA0CwgnVgwCjSA5kb79xexD22T1MzzYNzANBgkqhkiG9w0BAQEFAASCAgAQ5uHw
# WadOCBnquznrQ9PAZO2aPyyCwx83/PVJ3pHec69KLhXOvoVVL5KyFW0tFPc9W09q
# WbKwk2SmtxOMNhKBNCN/nOoTIErvp7CM6Z2P4mPB2+/sjeFE3QM3dW65zdFwxnyO
# +3+mXHzvH+4rSbGG58zJsdte7V7wGgjrbwfVhzGGz2j9EAzs2GMYHMsiT7OzSKFp
# pqb4qmTpe2DQN0cupVO5M3gkI2SBE3ys404Pbts0h0VS2Jn393aM+C22BIC4PfeP
# XHRJWijQxpm2VV1X4OvLFdFuWqRgLBFsp2I+Tp6OCfyGWye9CycIrjQnbBL1Ano/
# TSn2Rxh8tSZxX5rB7z2tJPrGLrg5s7gBce2nNu0eatDUGDH5VSHPNAVDi6MH1KrL
# mfi5yWrHOV+ejNL/CjyMZr+aYzU6hHkD6FWj/G51UFPukNqtA8fvuKKH8Fg5FxPJ
# QP2S4Qk8BtvlwnwMeffQVCfPWxU7n8Moui0hu4d9WqNj5s+++5WGC6Z+SAV1qsEl
# warZgXaHXmfMEucOl4XUgn5fYf9FKGSIICDOIN41BNSN0MWcOKMgNB4kLAbMNWke
# Zr4kh5OBlDInlJG663WG1ftiIVFrojH5KPcHLQ3GnG714j/axrjjYuqnKITllB0V
# 1+4qkN/INkYAXmC6qkWRepa7yZ6PKjYePt+N+aGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIPQP2cVkDsSQ/rzAchRChvIKMSLeGY+Zw0MSj8dSJ1f/AggbQcV28zD4
# whgPMjAyNjExMTQwODE2NTBaMAMCAQECEG0UfQv9MjF5ezymo7KJmu2gghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNDA4MTY1MFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEII+/
# DJGB52ZElAvvT9mFEWM+AFG/CjGRKxuunXCoyKMpMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiB022c4AFNKCjK8771ezNWpXjJzulc/6Wo7GVEtwzl4/AIhAIwY
# Uy5P/Mq+NqQeqlKZxk2UJkXuFHnLuGxwk/4d1VQu
# SIG # End signature block
