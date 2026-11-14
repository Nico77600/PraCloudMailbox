<#
.SYNOPSIS
    Fixtures of the objects that a remote Exchange session returns (dot-sourced by the test files).
.DESCRIPTION
    Collect reads Exchange through Import-PSSession: every object it gets is a deserialized copy. A deserialized object keeps its properties but loses its methods
    (a size is the text '1.2 GB (1,288,490,188 bytes)' with no ToBytes()), an enumeration becomes a number that still prints as its name, a collection becomes an
    ArrayList, a directory reference keeps its properties and its text. These fixtures build objects of private types that mimic the Exchange ones, then send them
    through the serializer of PowerShell remoting (PSSerializer, depth 8), which is the very conversion of a remote session: what comes out is a genuine Deserialized.* object.
    ConvertTo-DeserializedFixture does the round trip; the New-Fixture* functions build the live objects.
    $global:PcmRemoteStubs is the text of the commands of a test runner that has no Exchange: a remote session that opens and the round trip, to be pasted into a runner script.
.NOTES
    Author : Nicolas Fabert
    Version: 2.1.0
#>
#Requires -Version 7.5

if (-not ('PcmFixture.ByteQuantifiedSize' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.ObjectModel;
using System.Globalization;
namespace PcmFixture {
    public class ByteQuantifiedSize {
        private long bytes;
        public ByteQuantifiedSize(long value) { bytes = value; }
        public long ToBytes() { return bytes; }
        public override string ToString() {
            string unit = "B"; double scaled = bytes;
            if (bytes >= 1073741824L) { unit = "GB"; scaled = bytes / 1073741824.0; }
            else if (bytes >= 1048576L) { unit = "MB"; scaled = bytes / 1048576.0; }
            else if (bytes >= 1024L) { unit = "KB"; scaled = bytes / 1024.0; }
            return scaled.ToString(unit == "B" ? "0" : "0.0", CultureInfo.InvariantCulture) + " " + unit + " (" + bytes.ToString("N0", CultureInfo.InvariantCulture) + " bytes)";
        }
    }
    public class UnlimitedSize {
        public bool IsUnlimited { get; set; }
        public override string ToString() { return "Unlimited"; }
    }
    public class AdObjectId {
        public string DistinguishedName { get; set; }
        public Guid ObjectGuid { get; set; }
        public string Name { get; set; }
        public string Canonical { get; set; }
        public override string ToString() { return Canonical; }
    }
    public class ProxyAddress {
        public string Prefix { get; set; }
        public string AddressString { get; set; }
        public bool IsPrimaryAddress { get; set; }
        public override string ToString() { return Prefix + ":" + AddressString; }
    }
    public class NameOnlyReference {
        public string Name { get; set; }
        public string Text { get; set; }
        public override string ToString() { return Text; }
    }
    public class PrincipalId {
        public string Value { get; set; }
        public override string ToString() { return Value; }
    }
    public class MultiValue : Collection<object> { }
    public enum MailboxRights { None = 0, ReadPermission = 1, FullAccess = 2, DeleteItem = 4 }
    public enum RecipientTypeDetails { UserMailbox = 1, SharedMailbox = 2, RoomMailbox = 4, EquipmentMailbox = 8, RoomList = 16 }
    public enum RecipientType { UserMailbox = 1, MailUniversalSecurityGroup = 2, MailUniversalDistributionGroup = 3, MailContact = 4, User = 5 }
    public enum AutomateProcessing { None = 0, AutoUpdate = 1, AutoAccept = 2 }
    public enum ModerationNotifications { Never = 0, Internal = 1, Always = 2 }
    public enum JoinRestriction { Open = 0, Closed = 1, ApprovalRequired = 2 }
    public enum DepartRestriction { Open = 0, Closed = 1 }
    [Flags]
    public enum WorkDays { None = 0, Sunday = 1, Monday = 2, Tuesday = 4, Wednesday = 8, Thursday = 16, Friday = 32, Saturday = 64, Weekdays = 62, WeekendDays = 65, AllDays = 127 }
    public class TimeZoneInfoLike {
        public string Id { get; set; }
        public string DisplayName { get; set; }
    }
    public class TimeZoneValue {
        public TimeZoneInfoLike TimeZone { get; set; }
        public override string ToString() { return TimeZone.Id; }
    }
}
'@
}

function global:ConvertTo-DeserializedFixture {
    <# The conversion of a remote session: serialize and deserialize with PSSerializer. A collection comes back as the stream of its elements, like the output of a command. #>
    param([AllowNull()][Parameter(Mandatory)]$Value)
    $xml = [Management.Automation.PSSerializer]::Serialize($Value, 8)
    $copy = [Management.Automation.PSSerializer]::Deserialize($xml)
    if ($copy -is [Collections.IList]) { foreach ($element in $copy) { $element } } else { $copy }
}

function global:New-FixtureSize([long]$Bytes) { [PcmFixture.ByteQuantifiedSize]::new($Bytes) }

function global:New-FixtureAdId {
    param([Parameter(Mandatory)][string]$DistinguishedName, [string]$Canonical, [string]$Guid = ([guid]::NewGuid().ToString()))
    $leaf = ($DistinguishedName -split ',')[0] -replace '^[A-Za-z]+=', ''
    $id = [PcmFixture.AdObjectId]::new()
    $id.DistinguishedName = $DistinguishedName; $id.ObjectGuid = [guid]$Guid; $id.Name = $leaf
    $id.Canonical = $(if ($Canonical) { $Canonical } else { $leaf })
    return $id
}

function global:New-FixtureProxy([string]$Prefix, [string]$Address) {
    $proxy = [PcmFixture.ProxyAddress]::new()
    $proxy.Prefix = $Prefix; $proxy.AddressString = $Address; $proxy.IsPrimaryAddress = ($Prefix -ceq 'SMTP')
    return $proxy
}

function global:New-FixturePrincipal([string]$Text) {
    $principal = [PcmFixture.PrincipalId]::new()
    $principal.Value = $Text
    return $principal
}

function global:New-FixtureMulti {
    <# A multi-valued property (MultiValuedProperty): a collection of objects. #>
    param([object[]]$Item = @())
    $list = [PcmFixture.MultiValue]::new()
    foreach ($entry in $Item) { $list.Add($entry) }
    return , $list
}

function global:New-FixtureTimeZone([string]$Id, [string]$DisplayName = '') {
    $zone = [PcmFixture.TimeZoneInfoLike]::new()
    $zone.Id = $Id; $zone.DisplayName = $DisplayName
    $value = [PcmFixture.TimeZoneValue]::new()
    $value.TimeZone = $zone
    return $value
}

function global:Set-FixtureTypeName {
    <# The type name of an Exchange object on a fixture, so that the deserialized copy is a Deserialized.<that type> like the one of a real remote session. #>
    param([Parameter(Mandatory)]$Object, [Parameter(Mandatory)][string]$TypeName)
    $Object.PSObject.TypeNames.Insert(0, $TypeName)
    return $Object
}

function global:New-FixtureMailbox {
    <# A mailbox as Get-Mailbox returns it, with the typed values of Exchange (live). #>
    param([Parameter(Mandatory)][string]$Sam, [ValidateSet('UserMailbox', 'SharedMailbox', 'RoomMailbox', 'EquipmentMailbox')][string]$Type = 'UserMailbox', [string]$Guid = ([guid]::NewGuid().ToString()),
        [string]$Ou = 'OU=PRA,DC=contoso,DC=com', [string]$OuCanonical = 'contoso.com/PRA', [object[]]$OnBehalf = @())
    $addresses = New-FixtureMulti @((New-FixtureProxy 'SMTP' "$Sam@contoso.com"), (New-FixtureProxy 'smtp' "$Sam@contoso.mail.onmicrosoft.com"), (New-FixtureProxy 'X500' "/o=Old/cn=$Sam"))
    $mailbox = [pscustomobject]@{
        Guid = [guid]$Guid; Identity = (New-FixtureAdId "CN=$Sam,$Ou" "$OuCanonical/$Sam" $Guid); Name = $Sam; DisplayName = "Display $Sam"; Alias = $Sam; SamAccountName = $Sam
        UserPrincipalName = "$Sam@contoso.com"; DistinguishedName = "CN=$Sam,$Ou"; OrganizationalUnit = (New-FixtureAdId $Ou $OuCanonical)
        PrimarySmtpAddress = "$Sam@contoso.com"; WindowsEmailAddress = "$Sam@contoso.com"
        LegacyExchangeDN = "/o=Contoso/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=$Sam"
        ExchangeGuid = [guid]::NewGuid(); ArchiveGuid = [guid]::Empty; ArchiveStatus = 'None'
        RecipientTypeDetails = [PcmFixture.RecipientTypeDetails]$Type; HiddenFromAddressListsEnabled = ($Type -ne 'UserMailbox'); Database = (New-FixtureAdId 'CN=DB01,CN=Databases,CN=Exchange Administrative Group,CN=Administrative Groups,CN=Contoso,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=contoso,DC=com' 'DB01')
        EmailAddresses = $addresses; GrantSendOnBehalfTo = (New-FixtureMulti $OnBehalf); CustomAttribute1 = ''; CustomAttribute2 = 'Paris'
    }
    return (Set-FixtureTypeName $mailbox 'Microsoft.Exchange.Data.Directory.Management.Mailbox')
}

$global:PcmRemoteStubs = @'
function global:ConvertTo-PcmDeserialized { param($Value) $x = [Management.Automation.PSSerializer]::Serialize($Value, 8); [Management.Automation.PSSerializer]::Deserialize($x) }
function global:New-PSSession { [CmdletBinding()] param($ConfigurationName, $ConnectionUri, $Authentication) [IO.File]::AppendAllText($global:PcmSessionLog, ('{0}|{1}|{2}' -f $ConfigurationName, $ConnectionUri, $Authentication) + "`r`n"); [pscustomobject]@{ Id = 1 } }
function global:Import-PSSession { [CmdletBinding()] param($Session, $CommandName, [switch]$DisableNameChecking, [switch]$AllowClobber) New-Module -Name PcmFakeExchange -ScriptBlock { } }
function global:Remove-PSSession { [CmdletBinding()] param($Session) [IO.File]::AppendAllText($global:PcmSessionLog, 'removed' + "`r`n") }
'@
