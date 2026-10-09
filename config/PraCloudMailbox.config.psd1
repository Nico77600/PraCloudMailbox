#
#  PRA Cloud Mailbox - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.2.0
#
#  This file is read by Invoke-PraCloudMailbox.ps1. It is a PowerShell data
#  file: text between quotes, $true / $false, numbers, and @( ) for lists.
#  Lines starting with # are comments. An unknown setting is refused
#  (spelling mistakes are detected).
#
#  Relative paths (.\data, .\logs ...) are relative to the tool folder.
#  The same file is used on the Exchange side (Collect, Windows PowerShell 5.1)
#  and on the cloud admin server (Check, PowerShell 7): copy it with the database.
#
@{
    # Label of this environment, used in the backup file names (letters, digits, - _ .).
    Environment      = 'PROD'

    # ---------------------------------------------------------------------
    # Exchange on-premises (Collect only, before the disaster).
    #   Local  : Exchange Management Shell on this server (Exchange server or management tools)
    #   Remote : PowerShell remoting (Kerberos) to Server
    # DomainController: the DC used for every read of a run ('' = chosen by Exchange).
    # ---------------------------------------------------------------------
    Exchange = @{
        ConnectionMode   = 'Local'
        Server           = ''
        DomainController = ''
    }

    # ---------------------------------------------------------------------
    # Which mailboxes are collected (Collect without -Identity).
    #   Auto  : every mailbox of the selected types (under SearchBase when set)
    #   OU    : every mailbox under SearchBase (SearchBase required)
    #   Group : members of GroupDN (recursive)
    #   Csv   : the objects listed in CsvPath (column Identity = UPN, SMTP, sAMAccountName, DN or GUID)
    # System mailboxes (HealthMailbox*, SystemMailbox*, DiscoverySearchMailbox* ...) are always excluded.
    # ---------------------------------------------------------------------
    Scope = @{
        Mode                   = 'Auto'         # a pilot: 'OU' with SearchBase, 'Group' or 'Csv'
        SearchBase             = ''             # e.g. 'OU=Users,OU=Paris,DC=contoso,DC=com'
        GroupDN                = ''
        CsvPath                = ''             # e.g. '.\config\Targets.csv'
        IncludeUsers           = $true
        IncludeShared          = $true
        IncludeRoom            = $false        # rooms and equipment are collected when true; Convert still refuses
        IncludeEquipment       = $false        # them unless ConvertRooms is also true (booking settings are lost)
        ExcludeSamAccountNames = @()
        ConvertRooms           = $false        # $true: Convert/Recover a room or equipment mailbox like a shared one
    }

    # ---------------------------------------------------------------------
    # What Collect reads on top of the mailboxes.
    # ---------------------------------------------------------------------
    Collect = @{
        SharedPermissions            = $true    # FullAccess (+ AutoMapping), SendAs, SendOnBehalf of shared/room/equipment mailboxes
        ExpandGroupTrustees          = $true    # groups used as trustees: members stored (recursive), granted one by one at Convert
        ExcludeTrustees              = @('Administrator')   # sAMAccountName, DOMAIN\sam or wildcard; system accounts are always excluded
        MailboxStatistics            = $false   # item count and size (slower; useful to size Kiosk mailboxes, 2 GB)
        Contacts                     = $false   # mail contacts (optional)
        ContactsSearchBase           = ''
        DistributionGroups           = $false   # distribution and mail-enabled security groups, with members and settings (optional)
        DistributionGroupsSearchBase = ''
        DynamicDistributionGroups    = $false
    }

    # ---------------------------------------------------------------------
    # SQLite database of the snapshots. Collect writes it; the cloud actions only read it.
    # Copy it to the cloud admin server after each Collect (or replicate BackupFolder).
    # ---------------------------------------------------------------------
    Store = @{
        Path               = '.\data\PraCloudMailbox.db'
        KeepSnapshots      = 10                 # older complete snapshots are deleted
        BackupFolder       = '.\data\backup'    # a consistent copy after each Collect -Mode Apply
        MaxSnapshotAgeDays = 7                  # Check warns when the last snapshot is older
        JournalPath        = '.\data\PraCloudMailbox-journal.db'   # cloud side: batches of Convert / Recover (never overwritten by a Collect)
    }

    # ---------------------------------------------------------------------
    # Microsoft 365 tenant: certificate sign-in of the app registration (cloud actions).
    # ---------------------------------------------------------------------
    Cloud = @{
        TenantId              = ''             # directory (tenant) ID
        Organization          = ''             # e.g. 'contoso.onmicrosoft.com'
        AppId                 = ''             # application (client) ID of the app registration
        CertificateThumbprint = ''             # certificate of the app, in Cert:\CurrentUser\My or Cert:\LocalMachine\My
        DefaultUsageLocation  = ''             # two letters (e.g. 'FR'), used when a user has no usageLocation (a licence needs one)
    }

    # How long the cloud actions wait (Exchange Online provisioning, Entra Connect, holds) and how often they look.
    Polling = @{
        IntervalSeconds       = 20
        MailboxTimeoutMinutes = 30
        SyncTimeoutMinutes    = 30
        HoldTimeoutMinutes    = 30
    }

    # ---------------------------------------------------------------------
    # Exchange plan given at Convert.
    #   Users.Mode = Group  : the users are added to GroupId (object ID). A group synchronised from AD is
    #                         switched to cloud management first (test T15); a cloud group works as is.
    #   Users.Mode = Kiosk  : the Exchange Kiosk plan of the licence the user already has (e.g. Teams
    #                         Enterprise): no extra unit, 2 GB mailbox, no Outlook desktop.
    #   Users.Mode = Direct : SkuPartNumber assigned directly to each user.
    #   Shared.SkuPartNumber: temporary licence of the shared mailboxes, given back once each one is converted.
    #   Shared.Parallel     : shared mailboxes converted together, in waves (one temporary unit each, at most the
    #                         free units of the SKU at the start of each wave; 1 = one after another).
    # ---------------------------------------------------------------------
    Licensing = @{
        Users  = @{ Mode = 'Group'; GroupId = ''; SkuPartNumber = '' }     # GroupId: object ID of the licence group
        Shared = @{ SkuPartNumber = ''; Parallel = 100 }     # e.g. 'SPE_E3' or 'EXCHANGESTANDARD'
    }

    # ---------------------------------------------------------------------
    # Retention (Recover).
    #   TagAttribute/TagValue: written on the shared mailboxes at Convert so that the retention policy
    #                          (adaptive scope IsInactiveMailbox + tag) keeps them once inactive.
    #   HoldPolicy           : eDiscovery case hold policy used at Recover (users: protect the
    #                          ComponentShared; shared: hold before the identity is deleted).
    #   HoldPolicyLimit      : mailboxes per case hold policy (1000 = Microsoft Purview limit). When the
    #                          policy is full, Recover creates <HoldPolicy>-02, -03... in the same case.
    # ---------------------------------------------------------------------
    Retention = @{
        TagAttribute    = 'CustomAttribute1'
        TagValue        = 'Converted'
        HoldPolicy      = ''                # name of the case hold policy, e.g. 'PRA-Recover-Hold'
        HoldPolicyLimit = 1000
    }

    # ---------------------------------------------------------------------
    # Entra Connect rebuilt after the disaster (Recover: delta cycles; shared mailboxes: scheduler paused).
    #   Mode = Remoting : Invoke-Command to Server with the current account (ADSync cmdlets)
    #   Mode = Script   : ScriptPath -Operation Pause|Resume|Delta (any way to reach the server; exit 0 = done)
    #   Mode = Manual   : the operator does it when the tool asks
    # ---------------------------------------------------------------------
    EntraConnect = @{
        Mode       = 'Manual'
        Server     = ''                     # e.g. 'aadc01.contoso.com' (Remoting)
        ScriptPath = ''                     # e.g. '.\config\EntraConnect.sample.ps1' (Script)
        MinVersion = '2.5.76.0'
    }

    Logging = @{
        Folder = '.\logs'                    # one log + one transcript per run
    }
    Report = @{
        Enabled = $true                      # CSV + HTML report of each run
        Folder  = '.\reports'
    }
}
