#
#  PRA Cloud Mailbox - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 2.0.0
#
#  This file is read by Invoke-PraCloudMailbox.ps1. It is a PowerShell data
#  file: text between quotes, $true / $false, numbers, and @( ) for lists.
#  Lines starting with # are comments. An unknown setting is refused
#  (spelling mistakes are detected).
#
#  Relative paths (.\data, .\logs ...) are relative to the tool folder.
#  The same file is used by Collect (before the disaster) and by the cloud
#  actions (Check, Convert, Recover): copy it with the database.
#
@{
    # Label of this environment, used in the backup file names (letters, digits, - _ .).
    Environment      = 'PROD'

    # ---------------------------------------------------------------------
    # Exchange on-premises (Collect only, before the disaster).
    # Collect reads Exchange through a remote session (http://<Server>/PowerShell/, Kerberos): run it from any domain-joined
    # Windows server with PowerShell 7.5, typically the cloud admin server.
    #   Server           : host name of an Exchange server (required to run Collect), e.g. 'EX01.contoso.com'
    #   DomainController : the DC used for every read of a run ('' = chosen by Exchange)
    # ---------------------------------------------------------------------
    Exchange = @{
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
        IncludeRoom            = $false        # rooms and equipment are collected when true
        IncludeEquipment       = $false        # Convert still requires ConvertRooms
        ExcludeSamAccountNames = @()
        ConvertRooms           = $false        # resource handling is optional, disabled by default
        ResourceMode           = 'SynchronizedShared' # legacy synced path; 'CloudResource' for native resources EXCLUDED from sync
    }

    # ---------------------------------------------------------------------
    # What Collect reads on top of the mailboxes.
    # ---------------------------------------------------------------------
    Collect = @{
        SharedPermissions            = $true    # FullAccess (+ AutoMapping), SendAs, SendOnBehalf of shared/legacy synced resources; never native CloudResource
        ExpandGroupTrustees          = $true    # groups used as trustees: members stored (recursive), granted one by one at Convert
        ExcludeTrustees              = @('Administrator')   # sAMAccountName, DOMAIN\sam or wildcard; system accounts are always excluded
        MailboxStatistics            = $false   # item count and size (slower; useful to size Kiosk mailboxes, 2 GB)
        Contacts                     = $false   # mail contacts (optional)
        ContactsSearchBase           = ''
        DistributionGroups           = $false   # distribution and mail-enabled security groups, with members and settings (optional)
        DistributionGroupsSearchBase = ''
        DynamicDistributionGroups    = $false
        # Groups to keep with ALL their members (nested groups included, computers and contacts too), so that a run after the disaster can
        # take a wave with -Group <name> when Active Directory is gone. A name, a sAMAccountName, a distinguished name or an SMTP address.
        # The group of Scope.GroupDN (Scope.Mode = 'Group') is always kept too. A group that is not found is a warning, not an error.
        TargetGroups                 = @()      # e.g. @('PRA-Wave1', 'PRA-Wave2')
        # The daily task registered by -Action Schedule (Windows Task Scheduler, root folder).
        Schedule = @{
            TaskName = 'PraCloudMailbox Collect'    # the name of the task (not case sensitive); a task of this name is replaced
            At       = '02:00'                      # the time of day, on 24 hours; -At on the command line takes precedence
        }
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
    # Microsoft 365 tenant (cloud actions).
    #   SignIn = Certificate : the app registration signs in with its certificate (nobody has to be present; the default).
    #   SignIn = Interactive : an administrator signs in with a Microsoft 365 account, in a window (not in a scheduled task):
    #                          Microsoft Graph with delegated permissions (an administrator of the tenant consents once),
    #                          Exchange Online and Security & Compliance with the same account. TenantId is required;
    #                          AppId and CertificateThumbprint are not used.
    #   UserPrincipalName    : the administrator to propose at the sign-in (optional, any letter case).
    # ---------------------------------------------------------------------
    Cloud = @{
        SignIn                = 'Certificate'  # 'Certificate' or 'Interactive'
        UserPrincipalName     = ''             # e.g. 'admin@contoso.com' (Interactive: proposed at the sign-in)
        TenantId              = ''             # directory (tenant) ID
        Organization          = ''             # e.g. 'contoso.onmicrosoft.com'
        AppId                 = ''             # application (client) ID of the app registration
        CertificateThumbprint = ''             # certificate of the app, in Cert:\CurrentUser\My or Cert:\LocalMachine\My
        DefaultUsageLocation  = ''             # two letters (e.g. 'FR'), used when a user has no usageLocation (a licence needs one)
    }

    # Exchange Online module of the cloud admin server. The run stops before any change when it is older than this
    # (Microsoft.Graph.Authentication 2.41.1 or later is also required).
    Exo = @{
        MinModuleVersion = '3.10.1'
    }

    # How long the cloud actions wait (Exchange Online provisioning, Entra Connect, holds) and how often they look.
    # Check -Batch waits the same way: MailboxTimeoutMinutes for -Expect Provisioned, SyncTimeoutMinutes for Deprovisioned, HoldTimeoutMinutes for Retained.
    # CallTimeoutMinutes (1 to 60) is how long the tool waits for ONE answer of Microsoft 365 (every read, every change, every sign-in; an interactive sign-in window
    # gets at least 15 minutes). When Microsoft 365 does not answer in time, the run stops with a plain message (exit code 1) and the next steps give the command to
    # run again; a Convert or Recover batch is continued from its journal. Keep the default unless a very large read legitimately takes longer.
    Polling = @{
        IntervalSeconds       = 20
        MailboxTimeoutMinutes = 30
        SyncTimeoutMinutes    = 30
        HoldTimeoutMinutes    = 30
        CallTimeoutMinutes    = 5
    }

    # ---------------------------------------------------------------------
    # Exchange plan given at Convert.
    #   Users.Mode = Group  : the users are added to GroupId (object ID). A group synchronised from AD is
    #                         switched to cloud management first; a cloud group works as is.
    #   Users.Mode = Kiosk  : the Exchange Kiosk plan of the licence the user already has (e.g. Teams
    #                         Enterprise): no extra unit, 2 GB mailbox, no Outlook desktop.
    #   Users.Mode = Direct : SkuPartNumber assigned directly to each user.
    #   Shared.SkuPartNumber: temporary licence of the shared mailboxes, given back once each one is converted.
    #   Shared.Parallel     : shared mailboxes converted together, in waves (one temporary unit each, at most the
    #                         free units of the SKU at the start of each wave; 1 = one after another).
    #   Shared.KeepFree     : units of Shared.SkuPartNumber that Convert never uses (0 = none kept). Waves are at most
    #                         the free units minus this number. When the users get their Exchange plan from the same
    #                         licence (group or Direct), they must leave these units free too.
    # ---------------------------------------------------------------------
    Licensing = @{
        Users  = @{ Mode = 'Group'; GroupId = ''; SkuPartNumber = '' }     # GroupId: object ID of the licence group
        Shared = @{ SkuPartNumber = ''; Parallel = 100; KeepFree = 0 }     # e.g. 'SPE_E3' or 'EXCHANGESTANDARD'
    }

    # Recover.WaveSize: shared mailboxes rolled back together (1 to 1000). 0 = the same number as
    # Licensing.Shared.Parallel. Recover uses no licence unit: this number is independent from the licence pool.
    Recover = @{ WaveSize = 0 }

    # Native rooms/equipment EXCLUDED from Entra sync (Scope.ResourceMode = 'CloudResource').
    # Collect captures their booking settings; Convert creates actual room/equipment mailboxes.
    # Standard resources up to 50 GB need no licence or hold. LicensingGroupId is ONLY for old protected journals.
    # Recover routes via a MailContact to alias@OnPremisesRoutingDomain; verify that address and connector
    # reach the rebuilt resource BEFORE explicitly setting OnPremisesReady = $true.
    Resources = @{
        LicensingGroupId       = ''           # backwards compatibility only; new native resources never use this group
        OnPremisesRoutingDomain = ''          # e.g. 'resources.contoso.com'; must not loop back to the cloud resource
        OnPremisesReady        = $false       # operator assertion: native resources and hybrid mail routing restored
        AllowUnprotectedDeletion = $false     # explicit data-loss consent, reset in packages: no hold/inactive copy or cloud content preservation
        Parallel               = 100
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
