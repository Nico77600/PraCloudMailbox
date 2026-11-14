<#
.SYNOPSIS
    PRA Cloud Mailbox - common module: console, log, configuration, audit and reports.

.DESCRIPTION
    Helper functions used by Invoke-PraCloudMailbox.ps1 and by the other modules. Taken from PRA Remote
    Mailbox 2.0.0 (scenario 1) so that both tools look and behave the same. The module never reads
    Exchange and never connects to Microsoft 365. It runs in PowerShell 7.5 or later. It is organised in regions, in the order of an execution:

        1. Console theme       icons, frames and colours (console colours, no ANSI sequences)
        2. Console and log     Write-PcmBanner, Write-PcmStep, Write-PcmItem, Write-PcmLog, Write-PcmSummary;
                               for the window: Write-PcmEvent, Test-PcmStopRequest, Request-PcmOperator
        3. Configuration       Import-PcmConfiguration (reads and checks the .psd1 file)
        4. Audit               Initialize-PcmAudit (log file + PowerShell transcript)
        5. Results             Get-PcmOutcome, Write-PcmReport (CSV + HTML), Complete-PcmRun

    The run state is one hashtable (the "context") created by the entry script and passed to every
    function. Its main keys are described in docs\PraCloudMailbox-Guide.md, chapter "Architecture".

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    History : see CHANGELOG.md
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
$script:ManifestPath = Join-Path ([IO.Path]::GetDirectoryName($PSScriptRoot)) 'PraCloudMailbox.psd1'
$script:ToolVersion = $null
$script:PcmNativePath = $null
# Minimum versions of the cloud modules (Exo.MinModuleVersion in the configuration can only raise the Exchange one).
$script:MinGraphVersion = '2.41.1'
$script:MinExoVersion = '3.10.1'
$script:TranscriptOwner = $null
# The window runs the tool in a child process and follows it through this file of events (one JSON object per
# line); it asks for a stop with the file PRA_STOP_FILE (see Write-PcmEvent, Test-PcmStopRequest).
$script:EventFile = [string]$env:PRA_EVENT_FILE
$script:QuestionCount = 0
# Columns of the CSV report, in this order (also the fields copied from each result row).
$script:RowFields = @('Identity','Kind','PrimarySmtpAddress','ObjectGuid','ExchangeGuid','Action','Entra','ExchangeOnline',
    'TeamsStorage','Licence','Holds','Permissions','FinalStatus','Detail','Warnings')
# Final statuses a row can have. 'AlreadyDone' is counted as a success.
$script:FinalStatuses = @('Planned','AlreadyDone','Success','Error','Pending','Skipped')

#region 1. Console theme ----------------------------------------------------------------------
# The classic console (conhost) is the reference: colours are console
# colours (no ANSI sequences, so the transcript and the log stay clean text) and the default icons
# are characters present in the Consolas and Lucida Console fonts. Emoji are used only in Windows
# Terminal and VS Code. Force a style with the environment variable PRA_ICONS = Emoji|Symbols|Ascii.

function Get-PcmIconSet {
    <# Icons of one console style. 'Symbols' uses only characters of the classic console fonts. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji','Symbols','Ascii')][string]$Style)
    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F6DF; Ok = & $u 0x2705; Warn = & $u 0x1F7E1; Fail = & $u 0x274C; Info = & $u 0x1F539
                Skip = & $u 0x23E9; Directory = & $u 0x1F3E2; Plan = & $u 0x1F50E; Backup = & $u 0x1F4BE; Write = & $u 0x270F
                Sync = & $u 0x1F504; Cloud = & $u 0x2601; Key = & $u 0x1F510; Batch = & $u 0x1F4E6; Config = & $u 0x1F4DD
                Report = & $u 0x1F4CA; Folder = & $u 0x1F4C1; Clock = & $u 0x23F3; Mail = & $u 0x1F4E8; Target = & $u 0x1F3AF
                Log = & $u 0x1F4C4; Done = & $u 0x1F389; People = & $u 0x1F465; Next = & $u 0x1F449; Mode = & $u 0x1F9ED
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Directory = & $u 0x2302; Plan = & $u 0x25BA; Backup = & $u 0x25A0; Write = & $u 0x00B1
                Sync = & $u 0x2194; Cloud = & $u 0x263C; Key = & $u 0x00A7; Batch = & $u 0x25D9; Config = & $u 0x2261
                Report = & $u 0x2261; Folder = & $u 0x2302; Clock = & $u 0x25CB; Mail = '@'; Target = & $u 0x25D9
                Log = & $u 0x00B6; Done = & $u 0x221A; People = & $u 0x2192; Next = & $u 0x2192; Mode = & $u 0x25BA
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Directory = '#'; Plan = '?'; Backup = '='
                Write = '~'; Sync = '<>'; Cloud = '@'; Key = '$'; Batch = '#'; Config = '='; Report = '='; Folder = '>'
                Clock = ':'; Mail = '@'; Target = 'o'; Log = '='; Done = '*'; People = '&'; Next = '>'; Mode = '>'
            }
        }
    }
}

function Get-PcmFrameSet {
    <# Frame characters: rounded corners (Consolas), square corners for Lucida Console, ASCII otherwise. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji','Symbols','Ascii')][string]$Style)
    if ($Style -eq 'Ascii') { return @{ TopLeft='+'; TopRight='+'; BottomLeft='+'; BottomRight='+'; Horizontal='-'; Vertical='|' } }
    return @{ TopLeft=[string][char]0x256D; TopRight=[string][char]0x256E; BottomLeft=[string][char]0x2570
        BottomRight=[string][char]0x256F; Horizontal=[string][char]0x2500; Vertical=[string][char]0x2502 }
}

$script:IconStyle = if ($env:PRA_ICONS -in @('Emoji','Symbols','Ascii')) { $env:PRA_ICONS }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }
$script:Icons = Get-PcmIconSet $script:IconStyle
$script:Frame = Get-PcmFrameSet $script:IconStyle
# Emoji are two columns wide, symbols one: pad symbols so that the text stays aligned.
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
# Console colours of the theme. Accent = frames and step numbers.
$script:Theme = @{ Accent='Magenta'; AccentBack='DarkMagenta'; Dim='DarkGray'; Ok='Green'; Warn='Yellow'; Fail='Red'; Text='Gray'; Strong='White'; Info='Cyan' }

function Get-PcmIcon {
    param([Parameter(Mandatory)][string]$Name)
    $icon = $script:Icons[$Name]
    if (-not $icon) { $icon = $script:Icons['Info'] }
    return $icon + $script:IconPad
}

function Write-PcmHost {
    <#
    .SYNOPSIS
        Writes one console line made of coloured segments: @( @('text','Color'), @('text'), ... ).
        Console output only (the transcript records it); never the log file.
    #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Single console writer of the tool.')]
    param([Parameter(Mandatory)][object[]]$Segments, [string]$Color)
    # ONE Write-Host per line and one colour per line: the transcript writes every Write-Host -NoNewline piece
    # on its own line, which would make the audit transcript unreadable.
    # The colour is -Color, or the first colour given in the segments.
    $text = (@($Segments | ForEach-Object { [string]($_ | Select-Object -First 1) })) -join ''
    if (-not $Color) {
        foreach ($segment in $Segments) { $parts = @($segment); if ($parts.Count -ge 2 -and $parts[1]) { $Color = [string]$parts[1]; break } }
    }
    # PRA_ANSI=1: the colour as an ANSI sequence (console captures for the documentation).
    if ($env:PRA_ANSI -eq '1' -and $Color) {
        $codes = @{ Black=30; DarkBlue=34; DarkGreen=32; DarkCyan=36; DarkRed=31; DarkMagenta=35; DarkYellow=33; Gray=37; DarkGray=90; Blue=94; Green=92; Cyan=96; Red=91; Magenta=95; Yellow=93; White=97 }
        Write-Host ("{0}[{1}m{2}{0}[0m" -f [char]27, $codes[$Color], $text)
        return
    }
    if ($Color -and -not [Console]::IsOutputRedirected) { Write-Host $text -ForegroundColor $Color }
    else { Write-Host $text }
}
#endregion

#region 2. Console and log ----------------------------------------------------------------------

function Get-PcmValue {
    <#
    .SYNOPSIS
        Reads a key or a property without mixing up "missing", $null, $false, 0 and empty collections.
    .PARAMETER Object
        Dictionary or object, possibly $null.
    .PARAMETER Name
        Key or property name.
    .PARAMETER Default
        Value returned only when the object or the member is missing.
    #>
    [CmdletBinding()]
    param([AllowNull()][object]$Object, [Parameter(Mandatory)][string]$Name, [AllowNull()][object]$Default = $null)
    if ($null -eq $Object) { return ,$Default }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($entry in $Object.GetEnumerator()) { if ($entry.Key -eq $Name) { return ,($entry.Value) } }
    } else {
        $property = $Object.PSObject.Properties[$Name]
        if ($null -ne $property) { return ,($property.Value) }
    }
    return ,$Default
}

function Get-PcmToolVersion {
    <# The tool version: the ModuleVersion of PraCloudMailbox.psd1 is the one place that holds it. #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if (-not $script:ToolVersion) { $script:ToolVersion = [string](Import-PowerShellDataFile -LiteralPath $script:ManifestPath).ModuleVersion }
    return $script:ToolVersion
}

function Test-PcmSameText {
    <# Two values are the same text: compared by character code, letter case ignored (never by culture rules). #>
    [CmdletBinding()]
    param([AllowNull()][object]$Left, [AllowNull()][object]$Right)
    if ($null -eq $Left -or $null -eq $Right) { return ($null -eq $Left -and $null -eq $Right) }
    return [string]::Equals([string]$Left, [string]$Right, [StringComparison]::OrdinalIgnoreCase)
}

function Get-PcmDistinctText {
    <# Distinct texts of the pipeline, letter case ignored: the first spelling is kept, the order is kept. #>
    [CmdletBinding()]
    param([Parameter(ValueFromPipeline, Mandatory)][AllowNull()][AllowEmptyString()][object]$InputObject)
    begin { $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
    process { if ($null -ne $InputObject -and $seen.Add([string]$InputObject)) { [string]$InputObject } }
}

function Resolve-PcmOfficialValue {
    <# The official spelling of a value of a closed list (letter case ignored, compared by character code), or $null when the value is not in the list. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value, [Parameter(Mandatory)][string[]]$Allowed)
    if ($null -eq $Value) { return $null }
    foreach ($candidate in $Allowed) { if ([string]::Equals($candidate, $Value, [StringComparison]::OrdinalIgnoreCase)) { return $candidate } }
    return $null
}

$script:SystemMailboxPatterns = @('^HealthMailbox', '^SystemMailbox\{', '^Microsoft Exchange', '^DiscoverySearchMailbox', '^FederatedEmail\.', '^Migration\.', '^SM_')

function Test-PcmSystemMailboxName {
    <# The name of an Exchange system mailbox (health, arbitration, discovery, migration...): never collected, never converted. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Name)
    if (-not $Name) { return $false }
    foreach ($pattern in $script:SystemMailboxPatterns) { if ($Name -match $pattern) { return $true } }
    return $false
}

function ConvertTo-PcmDnKey {
    <# A distinguished name without the spaces around its commas and equal signs (an escaped comma stays), ready to be compared with the letter case ignored. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if (-not $Value) { return '' }
    $parts = [System.Collections.Generic.List[string]]::new()
    $current = New-Object Text.StringBuilder
    $escaped = $false
    foreach ($character in $Value.Trim().ToCharArray()) {
        if ($escaped) { [void]$current.Append($character); $escaped = $false; continue }
        if ($character -eq '\') { [void]$current.Append($character); $escaped = $true; continue }
        if ($character -eq ',') { $parts.Add($current.ToString()); $current.Clear() | Out-Null; continue }
        [void]$current.Append($character)
    }
    $parts.Add($current.ToString())
    $clean = foreach ($part in $parts) {
        $pair = $part.Trim() -split '\s*=\s*', 2
        if ($pair.Count -eq 2) { '{0}={1}' -f $pair[0].Trim(), $pair[1].Trim() } else { $part.Trim() }
    }
    return ($clean -join ',')
}

function ConvertFrom-PcmDnToCanonical {
    <# CN=Ann,OU=Sales,OU=Paris,DC=contoso,DC=com -> contoso.com/Paris/Sales/Ann (the canonical name Exchange shows for the organizational unit). #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$DistinguishedName)
    $key = ConvertTo-PcmDnKey $DistinguishedName
    if (-not $key) { return '' }
    $domain = [System.Collections.Generic.List[string]]::new(); $path = [System.Collections.Generic.List[string]]::new()
    foreach ($part in ([regex]::Split($key, '(?<!\\),'))) {
        $pair = $part -split '=', 2
        if ($pair.Count -ne 2) { return '' }
        $value = $pair[1] -replace '\\(.)', '$1'
        if ($pair[0] -match '^(?i:DC)$') { $domain.Add($value) } else { $path.Insert(0, $value) }
    }
    return (@(($domain -join '.')) + @($path) | Where-Object { $_ }) -join '/'
}

function Test-PcmInSearchBase {
    <#
    .SYNOPSIS
        Is an object under an organizational unit, sub-OUs included? The unit is given as a distinguished name (OU=Paris,DC=contoso,DC=com) or as a
        canonical name (contoso.com/Paris); the object is known by its distinguished name and/or by the canonical name of its organizational unit.
        The letter case never matters.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$SearchBase, [AllowNull()][AllowEmptyString()][string]$DistinguishedName, [AllowNull()][AllowEmptyString()][string]$OrganizationalUnit)
    $base = ([string]$SearchBase).Trim()
    if (-not $base) { return $false }
    $isDn = $base -match '^\s*[A-Za-z][A-Za-z0-9-]*\s*='
    $objectDn = ConvertTo-PcmDnKey $DistinguishedName
    $objectOu = ([string]$OrganizationalUnit).Trim().TrimEnd('/')
    if ($isDn) {
        $baseKey = ConvertTo-PcmDnKey $base
        if ($objectDn) { return $objectDn.EndsWith(',' + $baseKey, [StringComparison]::OrdinalIgnoreCase) }
        $baseCanonical = ConvertFrom-PcmDnToCanonical $baseKey
        if (-not $objectOu -or -not $baseCanonical) { return $false }
        return (Test-PcmSameText $objectOu $baseCanonical) -or $objectOu.StartsWith($baseCanonical + '/', [StringComparison]::OrdinalIgnoreCase)
    }
    $canonical = $base.TrimEnd('/')
    if (-not $objectOu -and $objectDn) {
        $full = ConvertFrom-PcmDnToCanonical $objectDn
        $objectOu = if ($full.Contains('/')) { $full.Substring(0, $full.LastIndexOf('/')) } else { '' }
    }
    if (-not $objectOu) { return $false }
    return (Test-PcmSameText $objectOu $canonical) -or $objectOu.StartsWith($canonical + '/', [StringComparison]::OrdinalIgnoreCase)
}

function ConvertTo-PcmGuidKey {
    <# A GUID in any notation and letter case becomes the lower-case hyphenated form; any other text is returned as it is (compared with the letter case ignored). #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)
    $guid = [guid]::Empty
    if ($Value -and [guid]::TryParse($Value.Trim(), [ref]$guid)) { return $guid.ToString() }
    return [string]$Value
}

function New-PcmRecordIndex {
    <# The mailboxes of a snapshot by GUID, UPN, SMTP address and immutable ID, to find the object that a group member stands for. The immutable ID is base64: its letter case counts. #>
    [CmdletBinding()]
    param([AllowEmptyCollection()][object[]]$Records = @())
    $index = [pscustomobject]@{
        ByGuid = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        ByUpn = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        BySmtp = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
        ByImmutable = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::Ordinal)
    }
    foreach ($record in $Records) {
        $guid = ConvertTo-PcmGuidKey ([string]$record.object_guid)
        if ($guid) { $index.ByGuid[$guid] = $record }
        if ([string]$record.user_principal_name) { $index.ByUpn[[string]$record.user_principal_name] = $record }
        if ([string]$record.primary_smtp_address) { $index.BySmtp[[string]$record.primary_smtp_address] = $record }
        if ($record.PSObject.Properties['immutable_id'] -and [string]$record.immutable_id) { $index.ByImmutable[[string]$record.immutable_id] = $record }
    }
    return $index
}

function Find-PcmRecord {
    <# The snapshot mailbox that a group member stands for: by GUID, then immutable ID, then UPN, then SMTP address. $null when there is none. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$Index, [string]$Guid, [string]$ImmutableId, [string]$Upn, [string]$Smtp)
    $found = $null
    if ($Guid -and $Index.ByGuid.TryGetValue((ConvertTo-PcmGuidKey $Guid), [ref]$found)) { return $found }
    if ($ImmutableId -and $Index.ByImmutable.TryGetValue($ImmutableId, [ref]$found)) { return $found }
    if ($Upn -and $Index.ByUpn.TryGetValue($Upn, [ref]$found)) { return $found }
    if ($Smtp -and $Index.BySmtp.TryGetValue($Smtp, [ref]$found)) { return $found }
    return $null
}

function Get-PcmCloudModuleReport {
    <#
    .SYNOPSIS
        What the cloud actions need from this computer, in plain words: Microsoft.Graph.Authentication at the minimum version or later, and
        ExchangeOnlineManagement at the minimum version or later. Reads the installed modules only; nothing is imported or connected.
    .DESCRIPTION
        If Microsoft Graph modules are already loaded in this PowerShell session, they must all be in one version, at the minimum or later: another version
        cannot be loaded next to it. A missing or too old module gives a message with the versions found, the minimum and the exact Install-Module command
        (one module per command).
    .PARAMETER MinExo
        Minimum ExchangeOnlineManagement version (Exo.MinModuleVersion of the configuration; it cannot go under the built-in minimum).
    .PARAMETER Available
        The installed modules (Get-Module -ListAvailable); read here when not given.
    .PARAMETER Loaded
        The Microsoft Graph modules already loaded in this session (Get-Module); read here when not given.
    .OUTPUTS
        Ok, Messages (one per problem), Graph (Ok, Version, Message) and Exo (Ok, Version, Message).
    #>
    [CmdletBinding()]
    param([string]$MinGraph = $script:MinGraphVersion, [string]$MinExo = $script:MinExoVersion, [object[]]$Available, [object[]]$Loaded)
    if (-not $PSBoundParameters.ContainsKey('Available')) {
        $Available = @(Get-Module -ListAvailable -Name 'Microsoft.Graph.Authentication', 'ExchangeOnlineManagement' -SkipEditionCheck -ErrorAction SilentlyContinue)
    }
    if (-not $PSBoundParameters.ContainsKey('Loaded')) { $Loaded = @(Get-Module -Name 'Microsoft.Graph.*' -ErrorAction SilentlyContinue) }
    $minGraphVersion = [version]$MinGraph
    $minExoVersion = [version]$(if ([version]$MinExo -lt [version]$script:MinExoVersion) { $script:MinExoVersion } else { $MinExo })
    $versionsOf = {
        param([object[]]$Modules, [string]$Name)
        $found = [Collections.Generic.List[version]]::new()
        foreach ($module in @($Modules)) {
            if ($null -eq $module -or -not [string]::Equals([string]$module.Name, $Name, [StringComparison]::OrdinalIgnoreCase)) { continue }
            $version = [version]$module.Version
            if (-not $found.Contains($version)) { $found.Add($version) }
        }
        $found.Sort(); $found.Reverse()
        return , $found
    }
    $texts = { param($List) if ($List.Count) { ($List | ForEach-Object { $_.ToString() }) -join ', ' } else { 'not installed' } }
    $scopeHint = 'to install for every user, use -Scope AllUsers in a window opened as administrator'
    $graphCommand = 'Install-Module Microsoft.Graph.Authentication -MinimumVersion {0} -Scope CurrentUser' -f $MinGraph
    $auth = & $versionsOf $Available 'Microsoft.Graph.Authentication'
    $recent = [Collections.Generic.List[version]]::new()
    foreach ($version in $auth) { if ($version -ge $minGraphVersion) { $recent.Add($version) } }
    $loadedVersions = [Collections.Generic.List[version]]::new()
    foreach ($module in @($Loaded)) {
        if ($null -eq $module -or -not ([string]$module.Name).StartsWith('Microsoft.Graph', [StringComparison]::OrdinalIgnoreCase)) { continue }
        if (-not $loadedVersions.Contains([version]$module.Version)) { $loadedVersions.Add([version]$module.Version) }
    }
    $graph = [ordered]@{ Ok = $false; Version = ''; Message = '' }
    if ($loadedVersions.Count -gt 1) {
        $graph.Message = 'Microsoft Graph: several versions are loaded in this PowerShell session ({0}). Close this window and start the tool in a new PowerShell 7 window.' -f (& $texts $loadedVersions)
    }
    elseif ($loadedVersions.Count -eq 1 -and $loadedVersions[0] -lt $minGraphVersion) {
        $graph.Message = 'Microsoft Graph {0} is already loaded in this PowerShell session, but the tool needs version {1} or later. Close this window, update the module with: {2} ({3}), then start the tool in a new PowerShell 7 window.' -f $loadedVersions[0], $MinGraph, $graphCommand, $scopeHint
    }
    elseif ($loadedVersions.Count -eq 1) { $graph.Ok = $true; $graph.Version = $loadedVersions[0].ToString() }
    elseif ($recent.Count) { $graph.Ok = $true; $graph.Version = $recent[0].ToString() }
    else {
        $graph.Message = 'Microsoft.Graph.Authentication: {0}; the tool needs version {1} or later. Run: {2} ({3}).' -f $(if ($auth.Count) { 'found ' + (& $texts $auth) } else { 'not installed' }), $MinGraph, $graphCommand, $scopeHint
    }
    $exoVersions = & $versionsOf $Available 'ExchangeOnlineManagement'
    $exo = [ordered]@{ Ok = $false; Version = ''; Message = '' }
    if ($exoVersions.Count -and $exoVersions[0] -ge $minExoVersion) { $exo.Ok = $true; $exo.Version = $exoVersions[0].ToString() }
    else {
        $exo.Message = 'ExchangeOnlineManagement: {0}; the tool needs version {1} or later. Run: Install-Module ExchangeOnlineManagement -MinimumVersion {1} -Scope CurrentUser ({2}).' -f $(if ($exoVersions.Count) { 'found ' + (& $texts $exoVersions) } else { 'not installed' }), $minExoVersion, $scopeHint
    }
    $messages = @(@($graph.Message, $exo.Message) | Where-Object { $_ })
    return [pscustomobject]@{ Ok = ($graph.Ok -and $exo.Ok); Messages = $messages; Graph = [pscustomobject]$graph; Exo = [pscustomobject]$exo }
}

function Format-PcmDuration {
    <#
    .SYNOPSIS
        '4.2 s', '3 min 05 s', '1 h 02 min'.
    #>
    param([Parameter(Mandatory)][double]$Seconds)
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalHours -ge 1) { return '{0} h {1:00} min' -f [int][Math]::Floor($t.TotalHours), $t.Minutes }
    if ($t.TotalMinutes -ge 1) { return '{0} min {1:00} s' -f $t.Minutes, $t.Seconds }
    return ('{0:0.0} s' -f $t.TotalSeconds).Replace(',', '.')
}

function Add-PcmIssue {
    <#
    .SYNOPSIS
        Records a run error (reports, exit code) without any output. Used by Write-PcmLog -Level Error.
    #>
    [CmdletBinding()]
    param([hashtable]$Context, [string]$Message, [string]$Source = 'Execution')
    $phase = [string](Get-PcmValue $Context 'CurrentPhase' '')
    if (-not $phase) { $phase = [string](Get-PcmValue $Context 'Phase' '') }
    [void]$Context.Issues.Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('o'); Phase = $phase; Source = $Source
        Operation = [string](Get-PcmValue $Context 'CurrentOperation' '')
        Identity = [string](Get-PcmValue $Context 'CurrentIdentity' ''); Message = $Message
    })
    $Context.ExitCode = 1
}

function ConvertTo-PcmPlainMessage {
    <#
    .SYNOPSIS
        The text of an error without its stack: the lines "at Namespace.Type.Method(...)", "at <ScriptBlock>, file: line N" and "--- End of ... stack trace ---" are removed.
        The card, the console and the report show this text; the log keeps the whole detail.
    #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    $frame = '^\s*(at\s+(<[^>\r\n]+>|[^\s(,]+\([^\r\n]*|[^\s,]+,\s+[^\r\n]*?\bline\s+\d+[^\r\n]*)|---\s*End of [^\r\n]*---)\s*$'
    $kept = @([regex]::Split($Text, '\r?\n') | Where-Object { $_ -notmatch $frame })
    return (($kept -join [Environment]::NewLine).TrimEnd())
}

function Add-PcmWarning {
    <#
    .SYNOPSIS
        Records a run warning (summary card, HTML report) without any output and without changing the
        exit code. Used by Write-PcmLog -Level Warning and Write-PcmItem -Status Warn.
    #>
    [CmdletBinding()]
    param([hashtable]$Context, [string]$Message)
    if (-not $Context.ContainsKey('WarningList')) { $Context['WarningList'] = New-Object 'Collections.Generic.List[object]' }
    [void]$Context['WarningList'].Add([pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'); Phase = [string](Get-PcmValue $Context 'CurrentPhase' ''); Message = $Message
    })
}

function Write-PcmEvent {
    <#
    .SYNOPSIS
        One event of the run for the window (PRA_EVENT_FILE): one JSON object per line. Nothing when the run does
        not belong to the window; never stops the run.
    .DESCRIPTION
        Kinds: start (banner), plan (the steps of the run, announced before the first one), step (a step starts, or ends when it has a
        state), count (how many objects the steps that work object by object will handle), target (which objects the run takes),
        progress (done/total of the current phase with a rough 'time left', or the attempts of a wait), item (console line), ask
        (question to the operator), summary (final card), result (the object returned by Complete-PcmRun). Every event has 'time'
        and 'kind'.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Kind, [hashtable]$Data = @{})
    if (-not $script:EventFile) { return }
    $record = [ordered]@{ time = (Get-Date).ToString('o'); kind = $Kind }
    foreach ($key in $Data.Keys) { $record[$key] = $Data[$key] }
    $line = ($record | ConvertTo-Json -Compress -Depth 6) + "`n"
    # The window reads the file while it is written (shared read): one more try if both meet.
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try { [IO.File]::AppendAllText($script:EventFile, $line, (New-Object Text.UTF8Encoding($false))); return }
        catch { Start-Sleep -Milliseconds 50 }
    }
}

function Test-PcmStopRequest {
    <# The window asked for a stop (the file PRA_STOP_FILE exists): the run starts no new object. #>
    [CmdletBinding()]
    param()
    return [bool]($env:PRA_STOP_FILE -and (Test-Path -LiteralPath $env:PRA_STOP_FILE))
}

function Request-PcmOperator {
    <#
    .SYNOPSIS
        Asks the operator a yes/no question: in the window when the run belongs to it (event 'ask', answered in the
        file named by the event), otherwise in the console ($Caller.ShouldContinue). A stop request answers no.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Question, [string]$Title = 'PRA Cloud Mailbox', [object]$Caller)
    if ($script:EventFile) {
        $script:QuestionCount++
        $answer = '{0}.answer-{1}' -f $script:EventFile, $script:QuestionCount
        Write-PcmEvent 'ask' @{ id = $script:QuestionCount; title = $Title; text = $Question; answer = $answer }
        while (-not (Test-Path -LiteralPath $answer)) {
            if (Test-PcmStopRequest) { return $false }
            Start-Sleep -Milliseconds 500
        }
        return ([string](Get-Content -LiteralPath $answer -Raw) -match '^\s*yes')
    }
    if (-not $Caller) { throw "$Question - the operator cannot be asked in this run." }
    return [bool]$Caller.ShouldContinue($Question, $Title)
}

function Write-PcmLogFile {
    <# Appends one line to the log file. A log that cannot be written stops the run (audit). #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Last-resort console message when the log is broken.')]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Level, [AllowEmptyString()][string]$Message)
    if (-not $Context.LogFile) { return }
    $line = '{0} [{1,-6}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'), $Level, ($Message -replace '[\r\n]+', ' | ')
    try { [IO.File]::AppendAllText($Context.LogFile, $line + "`r`n", (New-Object Text.UTF8Encoding($false))) }
    catch {
        $failure = 'The log file cannot be written: {0}' -f $_.Exception.Message
        Add-PcmIssue $Context $failure 'Log'
        Write-Host ('  [ERROR] ' + $failure) -ForegroundColor Red
        throw
    }
}

function Write-PcmLog {
    <#
    .SYNOPSIS
        Writes a message to the log file and, depending on the level, to the console.
    .PARAMETER Level
        Info      console item (bullet) + log
        Success   console item (tick, green) + log
        Warning   console item (yellow) + log; counted in the warnings
        Error     console item (red) + log; recorded as a run error (exit code 1)
        Sub       indented console detail line + log
        Detail    log only (full values, technical traces)
        Debug     log only; also on the console with -Verbose
        Step      log only (step headers are drawn by Write-PcmStep)
    .PARAMETER ForegroundColor
        Console colour of the text (presentation only: it never changes the level or the counters).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][AllowEmptyString()][string]$Message,
        [ValidateSet('Info','Sub','Success','Warning','Error','Debug','Step','Detail')][string]$Level = 'Info',
        [ConsoleColor]$ForegroundColor)
    $logLevel = @{ Info='INFO'; Sub='INFO'; Success='OK'; Warning='WARN'; Error='ERROR'; Debug='DEBUG'; Step='STEP'; Detail='DETAIL' }[$Level]
    if ($Level -eq 'Warning') { $Context.Warnings = [int](Get-PcmValue $Context 'Warnings' 0) + 1; Add-PcmWarning $Context $Message }
    if ($Level -eq 'Error') { Add-PcmIssue $Context $Message }
    Write-PcmLogFile -Context $Context -Level $logLevel -Message $Message
    $text = (ConvertTo-PcmPlainMessage $Message) -replace '[\r\n]+', ' | '
    switch ($Level) {
        'Debug' {
            $verbose = [bool](Get-PcmValue $Context 'VerboseEnabled' ($VerbosePreference -eq 'Continue'))
            Write-Verbose $text -Verbose:$verbose
        }
        'Detail' { }
        'Step' { }
        'Sub' {
            $color = if ($PSBoundParameters.ContainsKey('ForegroundColor')) { [string]$ForegroundColor } else { $script:Theme.Dim }
            Write-PcmHost @(,@(('         ' + $text), $color))
        }
        default {
            $status = @{ Info='Info'; Success='Ok'; Warning='Warn'; Error='Fail' }[$Level]
            $iconColor = @{ Info=$script:Theme.Info; Ok=$script:Theme.Ok; Warn=$script:Theme.Warn; Fail=$script:Theme.Fail }[$status]
            $textColor = if ($PSBoundParameters.ContainsKey('ForegroundColor')) { [string]$ForegroundColor }
                elseif ($status -in @('Warn','Fail')) { $iconColor } else { '' }
            Write-PcmHost @('      ', (Get-PcmIcon $status), $text) -Color $textColor
        }
    }
    if ($Level -in @('Info','Sub','Success','Warning','Error')) {
        $objectName = [string](Get-PcmValue $Context 'CurrentIdentity' '')
        Write-PcmEvent 'item' @{ status = @{ Info='Info'; Sub='Sub'; Success='Ok'; Warning='Warn'; Error='Fail' }[$Level]; text = $text; identity = $objectName; name = $objectName; stepId = [string](Get-PcmValue $Context 'CurrentStepId' '') }
    }
}

function Write-PcmItem {
    <#
    .SYNOPSIS
        One indented result line with a status icon, also written to the log.
    .PARAMETER Status
        Ok, Warn, Fail, Info or Skip. Fail does NOT record a run error: use Write-PcmLog -Level Error for that.
    .PARAMETER Icon
        Optional icon name instead of the status icon (Directory, Backup, Cloud, Mail, ...).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [ValidateSet('Ok','Warn','Fail','Info','Skip')][string]$Status = 'Info',
        [Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Icon)
    $color = @{ Ok=$script:Theme.Ok; Warn=$script:Theme.Warn; Fail=$script:Theme.Fail; Info=$script:Theme.Info; Skip=$script:Theme.Dim }[$Status]
    $textColor = if ($Status -in @('Warn','Fail','Skip')) { $color } else { '' }
    $symbol = if ($Icon) { Get-PcmIcon $Icon } else { Get-PcmIcon $Status }
    Write-PcmHost @('      ', $symbol, $Text) -Color $textColor
    if ($Status -eq 'Warn') { $Context.Warnings = [int](Get-PcmValue $Context 'Warnings' 0) + 1; Add-PcmWarning $Context $Text }
    Write-PcmLogFile -Context $Context -Level (@{ Ok='OK'; Warn='WARN'; Fail='FAIL'; Info='INFO'; Skip='SKIP' }[$Status]) -Message $Text
    $objectName = [string](Get-PcmValue $Context 'CurrentIdentity' '')
    Write-PcmEvent 'item' @{ status = $Status; text = $Text; identity = $objectName; name = $objectName; stepId = [string](Get-PcmValue $Context 'CurrentStepId' '') }
}

function Write-PcmProgress {
    <#
    .SYNOPSIS
        Progress of a long loop of objects: one console/log line (Write-PcmItem) plus a 'progress' event (window:
        a second progress bar and a line under the step one, done/total of the current phase only).
    .PARAMETER Phase
        Short label shown with the counts, e.g. 'Users', 'Shared mailboxes'.
    .PARAMETER Done
        Objects already finished (before the one about to start): the caller decides what counts as done.
    .PARAMETER ElapsedSeconds
        Time spent on this phase so far (a stopwatch started by the caller at its first object): used for a rough
        'time left at this pace' once at least one object is done. Omitted or 0 = no estimate.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Phase, [Parameter(Mandatory)][int]$Done,
        [Parameter(Mandatory)][int]$Total, [double]$ElapsedSeconds = 0)
    if ($Total -le 0) { return }
    $Done = [Math]::Max(0, [Math]::Min($Done, $Total))
    $percent = [int][Math]::Round(100.0 * $Done / $Total)
    $eta = ''
    if ($Done -gt 0 -and $Done -lt $Total -and $ElapsedSeconds -gt 0) { $eta = 'about {0} left at this pace' -f (Format-PcmDuration ($ElapsedSeconds * ($Total - $Done) / $Done)) }
    Write-PcmItem -Context $Context -Status Info -Icon Batch -Text ('{0}: {1} of {2} ({3}%){4}' -f $Phase, $Done, $Total, $percent, $(if ($eta) { '; ' + $eta } else { '' }))
    Write-PcmEvent 'progress' @{ phase = $Phase; done = $Done; total = $Total; percent = $percent; eta = $eta }
}

function Get-PcmRecoverWaveSize {
    <# Shared mailboxes rolled back together by Recover: Recover.WaveSize, or Licensing.Shared.Parallel when it is absent, empty or 0. #>
    param([Parameter(Mandatory)][object]$Config)
    $size = [int](Get-PcmValue (Get-PcmValue $Config 'Recover' @{}) 'WaveSize' 0)
    if ($size -ge 1) { return $size }
    return [Math]::Max(1, [int](Get-PcmValue (Get-PcmValue (Get-PcmValue $Config 'Licensing' @{}) 'Shared' @{}) 'Parallel' 100))
}

function Get-PcmDurationEstimate {
    <#
    .SYNOPSIS
        Rough order of magnitude for a Convert or Recover, from the number of waves and typical times: about 2 minutes of waiting per user wave (synchronisation, 23 s to 9 min seen), about 3 minutes per
        shared-mailbox pipeline round (mailbox provisioning and permissions, 24 s to 14 min seen). A real tenant can
        be much slower some evenings: this plans a maintenance window, it is not a guarantee.
    .PARAMETER UserWaves
        Number of waves of up to UserWaveSize users (0 when there is none).
    .PARAMETER SharedRounds
        Number of pipeline rounds of up to Licensing.Shared.Parallel shared mailboxes for Convert, or Recover.WaveSize for Recover (0 when there is none).
    #>
    param([Parameter(Mandatory)][int]$UserWaves, [Parameter(Mandatory)][int]$SharedRounds, [double]$UserWaveSeconds = 120, [double]$SharedRoundSeconds = 180)
    $seconds = ($UserWaves * $UserWaveSeconds) + ($SharedRounds * $SharedRoundSeconds)
    return [pscustomobject]@{ Seconds = $seconds; Text = $(if ($seconds -gt 0) { Format-PcmDuration $seconds } else { '' }) }
}

function Write-PcmBanner {
    <#
    .SYNOPSIS
        Title card at the start of an execution:

          ╭────────────────────────────────────────────────────────────────────────────╮
          │  ♦  PRA Cloud Mailbox                            v1.1.0 · Nicolas Fabert   │
          │     Exchange disaster recovery · scenario 2: on-premises lost → Exchange O │
          ╰────────────────────────────────────────────────────────────────────────────╯
               ►  Action     Convert · Preview (nothing is changed)
    .PARAMETER Details
        Ordered list of rows: key = label, value = @(IconName, Text) or plain text.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Title, [string]$Subtitle,
        [System.Collections.Specialized.OrderedDictionary]$Details)
    $F = $script:Frame; $T = $script:Theme; $width = 76
    $right = "v$($Context.Version) $([char]0x00B7) Nicolas Fabert"
    $left = "  $($script:Icons.Logo)  $Title"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { $script:Icons.Logo.Length }
    $visible = $left.Length - $script:Icons.Logo.Length + $iconWidth
    $gap = [Math]::Max(1, $width - $visible - $right.Length - 2)
    Write-Host ''
    Write-PcmHost @(,@(('  ' + $F.TopLeft + ($F.Horizontal * $width) + $F.TopRight), $T.Accent))
    Write-PcmHost @('  ', $F.Vertical, $left, (' ' * $gap), $right, '  ', $F.Vertical) -Color $T.Accent
    if ($Subtitle) {
        $sub = ('     ' + $Subtitle)
        if ($sub.Length -gt $width) { $sub = $sub.Substring(0, $width) }
        Write-PcmHost @('  ', $F.Vertical, $sub.PadRight($width), $F.Vertical) -Color $T.Accent
    }
    Write-PcmHost @(,@(('  ' + $F.BottomLeft + ($F.Horizontal * $width) + $F.BottomRight), $T.Accent))
    Write-PcmLogFile -Context $Context -Level 'STEP' -Message ("=== $Title v$($Context.Version) - run $($Context.RunId) ===")
    $eventDetails = [ordered]@{}
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon = '   '; $text = $value
            if ($value -is [array]) { $icon = Get-PcmIcon $value[0]; $text = $value[1] }
            Write-PcmHost @('     ', $icon, ('{0,-11}' -f $key), (' ' + $text))
            Write-PcmLogFile -Context $Context -Level 'INFO' -Message ('{0}: {1}' -f $key, $text)
            $eventDetails[$key] = [string]$text
        }
    }
    Write-PcmEvent 'start' @{ title = $Title; action = [string](Get-PcmValue $Context 'Action' ''); mode = [string](Get-PcmValue $Context 'Mode' ''); version = [string]$Context.Version
        runId = [string]$Context.RunId; logFile = [string](Get-PcmValue $Context 'LogFile' ''); details = $eventDetails }
}

function Write-PcmStep {
    <#
    .SYNOPSIS
        Step header with a coloured number pill and an icon, e.g.

           3/7  ►  Reading the targets and planning the changes
    .DESCRIPTION
        The step number comes from $Context.StepIndex (incremented here) and the total from
        $Context.StepTotal (set by the entry script for the selected action and phase).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Title, [string]$Icon = 'Plan')
    Write-PcmStepState -Context $Context
    $Context.StepIndex = [int](Get-PcmValue $Context 'StepIndex' 0) + 1
    $Context.CurrentPhase = $Title
    $stepId = Get-PcmPlanStepId -Context $Context -Title $Title
    $Context.CurrentStepId = $stepId; $Context.OpenStepId = $stepId
    $Context.StepWarningBase = [int](Get-PcmValue $Context 'Warnings' 0); $Context.StepIssueBase = $(if ($Context.ContainsKey('Issues')) { $Context.Issues.Count } else { 0 })
    $total = [int](Get-PcmValue $Context 'StepTotal' 0)
    $number = if ($total -gt 0) { '{0}/{1}' -f $Context.StepIndex, $total } else { [string]$Context.StepIndex }
    Write-Host ''
    Write-PcmHost @('  ', ('{0,5}' -f $number), '  ', (Get-PcmIcon $Icon), $Title) -Color $script:Theme.Accent
    Write-PcmLogFile -Context $Context -Level 'STEP' -Message ("[$number] $Title")
    $stepEvent = @{ index = $Context.StepIndex; total = $total; title = $Title }
    if ($stepId) { $stepEvent.id = $stepId }
    Write-PcmEvent 'step' $stepEvent
}

function Write-PcmPlan {
    <#
    .SYNOPSIS
        Announces the steps of the run before the first one (event 'plan'): the window shows them all from the start. The titles are the ones that
        Write-PcmStep is called with; each step gets the id s1, s2... in this order, carried by its 'step' events and by the 'item' events of the step.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string[]]$Titles)
    $steps = @(for ($i = 0; $i -lt $Titles.Count; $i++) { @{ id = ('s{0}' -f ($i + 1)); label = $Titles[$i]; count = 0 } })
    $Context.Plan = $steps
    $Context.StepTotal = $Titles.Count
    Write-PcmEvent 'plan' @{ steps = $steps }
}

function Get-PcmPlanStepId {
    <# The id that the plan gave to a step (empty when the step is not in the plan). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Title)
    if ($Context.ContainsKey('Plan')) {
        foreach ($step in $Context.Plan) { if ([string]::Equals([string]$step.label, $Title, [StringComparison]::OrdinalIgnoreCase)) { return [string]$step.id } }
    }
    return ''
}

function Write-PcmStepState {
    <#
    .SYNOPSIS
        Ends the step that is open (event 'step' with its id and a state: ok, warn, fail or skipped). Without -State: fail when the step recorded a run error,
        warn when it recorded a warning, else ok. Nothing when no step of the plan is open.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [ValidateSet('ok', 'warn', 'fail', 'skipped')][string]$State)
    $id = [string](Get-PcmValue $Context 'OpenStepId' '')
    if (-not $id) { return }
    $Context.OpenStepId = ''
    if (-not $State) {
        $State = 'ok'
        if ([int](Get-PcmValue $Context 'Warnings' 0) -gt [int](Get-PcmValue $Context 'StepWarningBase' 0)) { $State = 'warn' }
        if ($Context.ContainsKey('Issues') -and $Context.Issues.Count -gt [int](Get-PcmValue $Context 'StepIssueBase' 0)) { $State = 'fail' }
    }
    Write-PcmEvent 'step' @{ id = $id; state = $State }
}

function Write-PcmStepCount {
    <# How many objects the steps with these titles will handle (event 'count'): the window shows "0 / N objects". #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string[]]$Titles, [Parameter(Mandatory)][int]$Count)
    $ids = @($Titles | ForEach-Object { Get-PcmPlanStepId -Context $Context -Title $_ } | Where-Object { $_ })
    if ($Count -gt 0 -and $ids.Count) { Write-PcmEvent 'count' @{ count = $Count; stepIds = $ids } }
}

function Write-PcmSummary {
    <#
    .SYNOPSIS
        Final summary card:

          ╭─ √  Convert finished ──────────────────────────────────────────────────────╮
               √  Result     2 done · 0 pending · 0 errors
               ■  Snapshot   12 · 2026-10-07 15:40 (data\PraCloudMailbox.db)
          ╰────────────────────────────────────────────────────────────────────────────╯
    .PARAMETER Values
        Ordered list: key = label, value = @(IconName, Text) or plain text. A text can be an array
        (several lines under the same label).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Values, [ValidateSet('Ok','Warn','Fail')][string]$Status = 'Ok')
    $F = $script:Frame; $T = $script:Theme; $width = 76
    $color = @{ Ok=$T.Ok; Warn=$T.Warn; Fail=$T.Fail }[$Status]
    $icon = $script:Icons[@{ Ok='Done'; Warn='Warn'; Fail='Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { $icon.Length }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-PcmHost @('  ', ($F.TopLeft + $F.Horizontal), $head, (($F.Horizontal * $rest) + $F.TopRight)) -Color $color
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon = '   '; $text = $value
        if ($value -is [array] -and $value.Count -eq 2 -and $script:Icons.ContainsKey([string]$value[0])) { $rowIcon = Get-PcmIcon $value[0]; $text = $value[1] }
        $first = $true
        foreach ($line in @($text)) {
            $label = if ($first) { '{0,-11}' -f $key } else { ' ' * 11 }
            $iconText = if ($first) { $rowIcon } else { '   ' }
            Write-PcmHost @('     ', $iconText, $label, (' ' + $line))
            Write-PcmLogFile -Context $Context -Level 'INFO' -Message ('Summary - {0}: {1}' -f $key, $line)
            $first = $false
        }
    }
    Write-PcmHost @(,@(('  ' + $F.BottomLeft + ($F.Horizontal * $width) + $F.BottomRight), $color))
    $eventValues = [ordered]@{}
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $text = if ($value -is [array] -and $value.Count -eq 2 -and $script:Icons.ContainsKey([string]$value[0])) { $value[1] } else { $value }
        $eventValues[$key] = @(@($text) | ForEach-Object { [string]$_ })
    }
    Write-PcmEvent 'summary' @{ title = $Title; status = $Status; values = $eventValues }
}
#endregion

#region 3. Configuration ----------------------------------------------------------------------

function Resolve-PcmPath {
    <#
    .SYNOPSIS
        Full path of a file or folder; a relative path is relative to the tool folder.
    .DESCRIPTION
        Invalid characters (often a line break pasted with the path) are reported with their
        position and Unicode code instead of being removed silently.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Path, [string]$Name = 'Path', [Parameter(Mandatory)][string]$Root)
    $invalid = [IO.Path]::GetInvalidPathChars(); $count = 0; $positions = @()
    for ($i = 0; $i -lt $Path.Length; $i++) {
        if ($Path[$i] -in $invalid) {
            $count++
            if ($positions.Count -lt 8) { $positions += ('{0}:U+{1:X4}' -f $i, [int][char]$Path[$i]) }
        }
    }
    if ($count) {
        throw ('Invalid path in {0}: length={1}, invalid characters={2}, positions (0-based UTF-16)={3}, others={4}. Retype the value (a pasted line break is the usual cause).' -f $Name, $Path.Length, $count, ($positions -join ', '), ($count - $positions.Count))
    }
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath((Join-Path $Root $Path))
}

function ConvertTo-PcmScheduleTime {
    <# The time of day of the daily task, on 24 hours (2:00, 02:00, 23:30): the text as HH:mm, or $null when it is not a time of day. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Text)
    if ($null -eq $Text) { return $null }
    $parsed = [regex]::Match($Text.Trim(), '^(?<hour>[0-9]{1,2}):(?<minute>[0-9]{2})$')
    if (-not $parsed.Success) { return $null }
    $hour = [int]$parsed.Groups['hour'].Value; $minute = [int]$parsed.Groups['minute'].Value
    if ($hour -gt 23 -or $minute -gt 59) { return $null }
    return ('{0:00}:{1:00}' -f $hour, $minute)
}

function Merge-PcmConfiguration {
    <# Adds the missing keys from the defaults and checks the type of every known key. #>
    param([hashtable]$Values, [hashtable]$Defaults, [string]$Prefix = '')
    foreach ($key in $Defaults.Keys) {
        if (-not $Values.ContainsKey($key)) { $Values[$key] = $Defaults[$key]; continue }
        $value = $Values[$key]; $default = $Defaults[$key]
        if ($default -is [hashtable]) {
            if ($value -isnot [hashtable]) { throw "Configuration: $Prefix$key must be a section @{ ... }." }
            Merge-PcmConfiguration $value $default "$Prefix$key."
        }
        elseif ($default -is [bool] -and $value -isnot [bool]) { throw "Configuration: $Prefix$key must be `$true or `$false." }
        elseif ($default -is [string] -and $value -isnot [string]) { throw "Configuration: $Prefix$key must be text between quotes." }
        elseif ($default -is [int] -and $value -isnot [int]) { throw "Configuration: $Prefix$key must be a whole number." }
        elseif ($default -is [array] -and $value -isnot [array]) { throw "Configuration: $Prefix$key must be a list @( ... )." }
    }
    foreach ($key in @($Values.Keys)) {
        if (-not $Defaults.ContainsKey($key)) { throw "Configuration: unknown setting $Prefix$key (check the spelling, see the administrator guide, chapter 6)." }
    }
}

function Get-PcmDefaultConfiguration {
    <# Default values of every setting (the reference list of the valid keys). #>
    return @{
        Environment = 'PROD'
        Exchange = @{ Server=''; DomainController='' }
        Scope = @{ Mode='OU'; SearchBase=''; GroupDN=''; CsvPath=''; IncludeUsers=$true; IncludeShared=$true; IncludeRoom=$false; IncludeEquipment=$false; ExcludeSamAccountNames=@(); ConvertRooms=$false; ResourceMode='SynchronizedShared' }
        Collect = @{
            SharedPermissions=$true; ExpandGroupTrustees=$true; ExcludeTrustees=@(); MailboxStatistics=$false
            Contacts=$false; ContactsSearchBase=''; DistributionGroups=$false; DistributionGroupsSearchBase=''; DynamicDistributionGroups=$false
            TargetGroups=@(); MaxDropPercent=50
            Schedule=@{ TaskName='PraCloudMailbox Collect'; At='02:00' }
        }
        Store = @{ Path='.\data\PraCloudMailbox.db'; KeepSnapshots=10; BackupFolder='.\data\backup'; MaxSnapshotAgeDays=7; JournalPath='.\data\PraCloudMailbox-journal.db' }
        Cloud = @{ TenantId=''; Organization=''; AppId=''; CertificateThumbprint=''; DefaultUsageLocation=''; SignIn='Certificate'; UserPrincipalName='' }
        Exo = @{ MinModuleVersion='3.10.1' }
        Polling = @{ IntervalSeconds=20; MailboxTimeoutMinutes=30; SyncTimeoutMinutes=30; HoldTimeoutMinutes=30; CallTimeoutMinutes=5 }
        Licensing = @{
            Users = @{ Mode='Group'; GroupId=''; SkuPartNumber='' }
            Shared = @{ SkuPartNumber=''; Parallel=100; KeepFree=0 }
        }
        Recover = @{ WaveSize=0 }
        Retention = @{ TagAttribute='CustomAttribute1'; TagValue='Converted'; HoldPolicy=''; HoldPolicyLimit=1000 }
        Resources = @{ LicensingGroupId=''; OnPremisesRoutingDomain=''; OnPremisesReady=$false; AllowUnprotectedDeletion=$false; Parallel=100 }
        EntraConnect = @{ Mode='Remoting'; Server=''; ScriptPath=''; CredentialPath=''; MinVersion='2.5.76.0' }
        Logging = @{ Folder='.\logs' }
        Report = @{ Enabled=$true; Folder='.\reports' }
    }
}

function Add-PcmNativeImport {
    <# One kernel32 import, emitted with Reflection.Emit: DllImport with SetLastError = true (the Win32 error is kept for Marshal.GetLastWin32Error). #>
    param([Parameter(Mandatory)][Reflection.Emit.TypeBuilder]$Type, [Parameter(Mandatory)][string]$Name, [string]$EntryPoint = '', [Parameter(Mandatory)][Type]$Return, [Parameter(Mandatory)][Type[]]$Parameters)
    if (-not $EntryPoint) { $EntryPoint = $Name }
    $method = $Type.DefineMethod($Name, 'Public,Static,PinvokeImpl,HideBySig', $Return, $Parameters)
    $method.SetImplementationFlags([Reflection.MethodImplAttributes]::PreserveSig)
    $import = [Runtime.InteropServices.DllImportAttribute]
    $attribute = New-Object Reflection.Emit.CustomAttributeBuilder($import.GetConstructor([Type[]]@([string])), [object[]]@('kernel32.dll'),
        [Reflection.FieldInfo[]]@($import.GetField('EntryPoint'), $import.GetField('CharSet'), $import.GetField('SetLastError')),
        [object[]]@($EntryPoint, [Runtime.InteropServices.CharSet]::Unicode, $true))
    $method.SetCustomAttribute($attribute)
}

function Get-PcmWin32ErrorText {
    <# '<code>: <system message>' for a Win32 error code (Marshal.GetLastWin32Error after a native call). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$Code)
    return ('{0}: {1}' -f $Code, (New-Object ComponentModel.Win32Exception($Code)).Message)
}

function Get-PcmNativePathApi {
    <# Two kernel32 calls (open a handle, read its final path) bound in memory with Reflection.Emit (static AssemblyBuilder.DefineDynamicAssembly, one-argument DefineDynamicModule and CreateType: the same calls in every PowerShell edition): no compiler, temporary file, DLL or cache. #>
    if ($script:PcmNativePath) { return $script:PcmNativePath }
    $assembly = [Reflection.Emit.AssemblyBuilder]::DefineDynamicAssembly((New-Object Reflection.AssemblyName 'PcmNativePath'), [Reflection.Emit.AssemblyBuilderAccess]::Run)
    $type = $assembly.DefineDynamicModule('PcmNativePath').DefineType('Api', 'Public,Class,Sealed,Abstract')
    Add-PcmNativeImport -Type $type -Name 'CreateFileW' -Return ([Microsoft.Win32.SafeHandles.SafeFileHandle]) -Parameters ([Type[]]@([string], [uint32], [uint32], [IntPtr], [uint32], [uint32], [IntPtr]))
    Add-PcmNativeImport -Type $type -Name 'GetFinalPathNameByHandleW' -Return ([uint32]) -Parameters ([Type[]]@([Microsoft.Win32.SafeHandles.SafeFileHandle], [Text.StringBuilder], [uint32], [uint32]))
    $script:PcmNativePath = $type.CreateType()
    return $script:PcmNativePath
}

function Get-PcmPhysicalPath {
    <#
    .SYNOPSIS
        Physical location of a path as the filesystem resolves it (junctions, symbolic links, SUBST and mapped drives,
        8.3 names, \\?\ prefixes), even when its last folders do not exist yet.
    .DESCRIPTION
        The deepest existing ancestor is opened and its final path read from the handle; the missing folders are
        appended unchanged. Anything that cannot be resolved throws: a path is never assumed to be elsewhere.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    if ($full -match '^\\\\[?.]\\') {
        if (@($full.Split('\') | Select-Object -Skip 3 | Where-Object { [string]::Equals($_, '.', [StringComparison]::Ordinal) -or [string]::Equals($_, '..', [StringComparison]::Ordinal) }).Count) { throw "Path $Path uses a device namespace with relative segments; it cannot be resolved safely." }
    }
    $existing = $full.TrimEnd('\')
    if ($existing -match '^[A-Za-z]:$' -or $existing -match '^\\\\[?.]\\[A-Za-z]:$') { $existing += '\' }
    $missing = New-Object Collections.Generic.List[string]
    while (-not ([IO.Directory]::Exists($existing) -or [IO.File]::Exists($existing))) {
        $parent = [IO.Path]::GetDirectoryName($existing)
        if (-not $parent) { throw "Path $Path cannot be resolved: neither it nor any parent exists, so its physical location is unknown." }
        $missing.Insert(0, [IO.Path]::GetFileName($existing))
        $existing = $parent
    }
    $native = if ($existing -match '^\\\\[?.]\\') { $existing }
              elseif ($existing.StartsWith('\\', [StringComparison]::Ordinal)) { '\\?\UNC\' + $existing.Substring(2) }
              else { '\\?\' + $existing }
    $api = Get-PcmNativePathApi
    $handle = $api::CreateFileW($native, [uint32]0, [uint32]7, [IntPtr]::Zero, [uint32]3, [uint32]0x02000000, [IntPtr]::Zero)
    try {
        if ($handle.IsInvalid) { throw "Path $Path cannot be resolved: the filesystem refused to open $existing." }
        $buffer = New-Object Text.StringBuilder 1024
        $length = $api::GetFinalPathNameByHandleW($handle, $buffer, [uint32]$buffer.Capacity, [uint32]0)
        if ($length -gt $buffer.Capacity) {
            $buffer = New-Object Text.StringBuilder ([int]$length + 1)
            $length = $api::GetFinalPathNameByHandleW($handle, $buffer, [uint32]$buffer.Capacity, [uint32]0)
        }
        if ($length -eq 0 -or $length -gt $buffer.Capacity) { throw "Path $Path cannot be resolved: the final path of $existing is unavailable." }
        $physical = $buffer.ToString()
    }
    finally { if ($handle) { $handle.Dispose() } }
    if ($physical.StartsWith('\\?\UNC\', [StringComparison]::OrdinalIgnoreCase)) { $physical = '\\' + $physical.Substring(8) }
    elseif ($physical -match '^\\\\\?\\[A-Za-z]:\\') { $physical = $physical.Substring(4) }
    elseif ($physical.StartsWith('\\?\', [StringComparison]::Ordinal)) { throw "Path $Path cannot be resolved to a drive or share path ($physical); refusing the ambiguous namespace." }
    foreach ($name in $missing) { $physical = $physical.TrimEnd('\') + '\' + $name }
    return $physical
}

function Assert-PcmOutputFolder {
    <#
    .SYNOPSIS
        Refuses an output folder inside the authenticated package other than its own canonical folder.
    .DESCRIPTION
        The package integrity check treats every file of the package as stable except the generated root
        folders data, logs and reports (exact case). A backup, log or report folder anywhere else inside
        the package would make the next start fail, so it is refused here, before any file is created. The
        test uses the physical location of the package and of the folder (Get-PcmPhysicalPath), so a \\?\ prefix,
        an 8.3 name, a junction, SUBST or a legacy alias such as 'Documents and Settings' cannot hide the package;
        a path that cannot be resolved is refused. The returned path is the one to use: the physical path when
        it is inside the package, otherwise the configured folder unchanged.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Canonical, [Parameter(Mandatory)][string]$Root)
    $rootFull = [IO.Path]::GetFullPath($Root).TrimEnd('\')
    $full = [IO.Path]::GetFullPath($Path)
    if ($full.Length -gt [IO.Path]::GetPathRoot($full).Length) { $full = $full.TrimEnd('\') }
    $lexical = [string]::Equals($full, $rootFull, [StringComparison]::OrdinalIgnoreCase) -or $full.StartsWith($rootFull + '\', [StringComparison]::OrdinalIgnoreCase)
    if ($lexical) {
        $relative = $full.Substring($rootFull.Length).TrimStart('\')
        $first = if ($relative) { $relative.Split('\')[0] } else { '' }
        if (-not [string]::Equals($first, $Canonical, [StringComparison]::Ordinal)) {
            throw "Configuration: $Name ($full) is inside the signed package and is not the $Canonical folder. Use '.\$Canonical' (a sub-folder of it is accepted) or a folder outside the package; any other folder inside the package would break package integrity."
        }
        $current = $rootFull
        foreach ($segment in $relative.Split('\')) {
            $current = $current + '\' + $segment
            if (-not (Test-Path -LiteralPath $current)) { break }
            if ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Configuration: $Name goes through a reparse point inside the package ($current); the integrity check refuses it."
            }
        }
    }
    try {
        $physicalRoot = (Get-PcmPhysicalPath $rootFull).TrimEnd('\')
        $physical = (Get-PcmPhysicalPath $full).TrimEnd('\')
    }
    catch { throw "Configuration: $Name ($full) cannot be checked against the signed package, so it is refused: $($_.Exception.Message)" }
    $inside = [string]::Equals($physical, $physicalRoot, [StringComparison]::OrdinalIgnoreCase) -or $physical.StartsWith($physicalRoot + '\', [StringComparison]::OrdinalIgnoreCase)
    if (-not $inside) {
        if ($lexical) { throw "Configuration: $Name ($full) is inside the signed package but resolves outside it ($physical); the location is ambiguous and is refused." }
        return $full
    }
    $relative = $physical.Substring($physicalRoot.Length).TrimStart('\')
    $first = if ($relative) { $relative.Split('\')[0] } else { '' }
    if (-not [string]::Equals($first, $Canonical, [StringComparison]::Ordinal)) {
        throw "Configuration: $Name ($full) resolves to $physical, inside the signed package, and is not the $Canonical folder. Use '.\$Canonical' (a sub-folder of it is accepted) or a folder outside the package; any other folder inside the package would break package integrity."
    }
    $current = $physicalRoot
    foreach ($segment in $relative.Split('\')) {
        $current = $current + '\' + $segment
        if (-not (Test-Path -LiteralPath $current)) { break }
        if ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Configuration: $Name goes through a reparse point inside the package ($current); the integrity check refuses it."
        }
    }
    return $physical
}

function Import-PcmConfiguration {
    <#
    .SYNOPSIS
        Reads the configuration file, adds the default values, checks every value and resolves the paths.
    .PARAMETER Path
        The .psd1 configuration file.
    .PARAMETER Root
        Tool folder: relative paths in the file are relative to it.
    .OUTPUTS
        Hashtable. Paths (Store.Path, Store.BackupFolder, Logging.Folder, Report.Folder, Scope.CsvPath)
        are returned as full paths.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Root)
    $file = Resolve-PcmPath $Path 'ConfigPath' $Root
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw "Configuration file not found: $file" }
    $config = Import-PowerShellDataFile -LiteralPath $file -ErrorAction Stop
    if ($config.ContainsKey('Recover') -and $config.Recover -is [hashtable] -and $config.Recover.ContainsKey('WaveSize') -and ($null -eq $config.Recover.WaveSize -or ($config.Recover.WaveSize -is [string] -and -not $config.Recover.WaveSize.Trim()))) { $config.Recover.WaveSize = 0 }
    if ($config.ContainsKey('Exchange') -and $config.Exchange -is [hashtable] -and $config.Exchange.ContainsKey('ConnectionMode')) {
        $mode = ([string]$config.Exchange.ConnectionMode).Trim()
        if (Test-PcmSameText $mode 'Local') { throw "Configuration: Exchange.ConnectionMode = 'Local' (the Exchange Management Shell of this server) is not supported any more. Collect now always reads Exchange through a remote session: remove ConnectionMode and set Exchange.Server to the host name of an Exchange server." }
        if (-not (Test-PcmSameText $mode 'Remote')) { throw "Configuration: Exchange.ConnectionMode is no longer a setting (Collect always reads Exchange through a remote session): remove it and set Exchange.Server." }
        $config.Exchange.Remove('ConnectionMode')
    }
    Merge-PcmConfiguration $config (Get-PcmDefaultConfiguration)

    if ($config.Environment -notmatch '^[A-Za-z0-9_.-]{1,64}$') { throw 'Configuration: Environment must be 1 to 64 letters, digits, dot, dash or underscore (it is used in file names).' }
    $exchange = $config.Exchange
    foreach ($name in @('Server','DomainController')) {
        if ($exchange[$name] -and $exchange[$name] -notmatch '^[A-Za-z0-9.-]+$') { throw "Configuration: Exchange.$name must be a host name." }
    }
    $scope = $config.Scope
    $official = Resolve-PcmOfficialValue ([string]$scope.Mode) @('Auto','OU','Group','Csv')
    if (-not $official) { throw 'Configuration: Scope.Mode must be Auto, OU, Group or Csv.' }
    $scope.Mode = $official
    if ($scope.Mode -eq 'OU' -and -not $scope.SearchBase) { throw 'Configuration: Scope.Mode = OU requires Scope.SearchBase (a search of the whole organisation needs Mode = Auto).' }
    if ($scope.Mode -eq 'Group' -and -not $scope.GroupDN) { throw 'Configuration: Scope.Mode = Group requires Scope.GroupDN.' }
    if ($scope.Mode -eq 'Csv' -and -not $scope.CsvPath) { throw 'Configuration: Scope.Mode = Csv requires Scope.CsvPath.' }
    if (-not ($scope.IncludeUsers -or $scope.IncludeShared -or $scope.IncludeRoom -or $scope.IncludeEquipment)) { throw 'Configuration: Scope must include at least one mailbox type (IncludeUsers, IncludeShared, IncludeRoom, IncludeEquipment).' }
    $official = Resolve-PcmOfficialValue ([string]$scope.ResourceMode) @('SynchronizedShared','CloudResource')
    if (-not $official) { throw 'Configuration: Scope.ResourceMode must be SynchronizedShared (legacy synced resources) or CloudResource (unsynchronised native resources).' }
    $scope.ResourceMode = $official
    $resources = $config.Resources
    $groupId = [guid]::Empty
    if ($resources.LicensingGroupId -and -not [guid]::TryParse($resources.LicensingGroupId, [ref]$groupId)) { throw 'Configuration: Resources.LicensingGroupId must be the object ID (GUID) of the cloud licence group.' }
    if ($resources.LicensingGroupId) { $resources.LicensingGroupId = $groupId.ToString() }
    if ($resources.OnPremisesRoutingDomain -and ($resources.OnPremisesRoutingDomain.Length -gt 253 -or $resources.OnPremisesRoutingDomain -notmatch '^(?=.{1,253}$)([A-Za-z0-9](?:[A-Za-z0-9-]{0,61}[A-Za-z0-9])?\.)+[A-Za-z]{2,63}$')) { throw 'Configuration: Resources.OnPremisesRoutingDomain must be a DNS domain (not an email address or URL).' }
    if ($resources.Parallel -isnot [int] -or $resources.Parallel -lt 1 -or $resources.Parallel -gt 1000) { throw 'Configuration: Resources.Parallel must be a whole number between 1 and 1000.' }
    $store = $config.Store
    if ($store.KeepSnapshots -lt 1 -or $store.KeepSnapshots -gt 1000) { throw 'Configuration: Store.KeepSnapshots must be between 1 and 1000.' }
    if ($store.MaxSnapshotAgeDays -lt 1 -or $store.MaxSnapshotAgeDays -gt 365) { throw 'Configuration: Store.MaxSnapshotAgeDays must be between 1 and 365.' }
    if (-not $store.Path) { throw 'Configuration: Store.Path is empty (SQLite database file).' }
    if ($config.Collect.MaxDropPercent -lt 0 -or $config.Collect.MaxDropPercent -gt 100) { throw 'Configuration: Collect.MaxDropPercent must be a whole number between 0 and 100 (a Collect that finds more than this percentage fewer mailboxes than the last complete snapshot, or none, is refused; 100 accepts any size).' }
    $targetGroups = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in @($config.Collect.TargetGroups)) {
        $text = ([string]$entry).Trim()
        if (-not $text) { continue }
        if ($text -match '[\x00-\x1f]') { throw 'Configuration: Collect.TargetGroups holds a value with a control character (a pasted line break is the usual cause).' }
        if (-not (@($targetGroups | Where-Object { Test-PcmSameText $_ $text }).Count)) { $targetGroups.Add($text) }
    }
    $config.Collect.TargetGroups = $targetGroups.ToArray()
    $schedule = $config.Collect.Schedule
    $schedule.TaskName = ([string]$schedule.TaskName).Trim()
    if (-not $schedule.TaskName -or $schedule.TaskName.Length -gt 200 -or $schedule.TaskName -match '[\\/:*?"<>|\x00-\x1f]') { throw 'Configuration: Collect.Schedule.TaskName must be a task name of 1 to 200 characters, without \ / : * ? " < > |.' }
    $time = ConvertTo-PcmScheduleTime ([string]$schedule.At)
    if (-not $time) { throw 'Configuration: Collect.Schedule.At must be a time of day on 24 hours, for example 02:00.' }
    $schedule.At = $time
    $cloud = $config.Cloud
    $official = Resolve-PcmOfficialValue ([string]$cloud.SignIn) @('Certificate','Interactive')
    if (-not $official) { throw 'Configuration: Cloud.SignIn must be Certificate (the app registration signs in with its certificate) or Interactive (an administrator signs in with a Microsoft 365 account).' }
    $cloud.SignIn = $official
    $cloud.UserPrincipalName = ([string]$cloud.UserPrincipalName).Trim()
    if ($cloud.UserPrincipalName -and $cloud.UserPrincipalName -notmatch '^[^@\s]+@[^@\s]+$') { throw 'Configuration: Cloud.UserPrincipalName must be the sign-in name of the administrator (admin@contoso.com).' }
    if ($cloud.SignIn -eq 'Interactive' -and -not $cloud.TenantId) { throw 'Configuration: Cloud.SignIn = Interactive requires Cloud.TenantId (the sign-in must go to the right tenant).' }
    if ($cloud.SignIn -eq 'Certificate' -and [bool]$cloud.AppId -ne [bool]$cloud.CertificateThumbprint) { throw 'Configuration: Cloud.AppId and Cloud.CertificateThumbprint go together (certificate sign-in of the app registration).' }
    if ($cloud.TenantId -and $cloud.TenantId -notmatch '^[0-9a-fA-F-]{36}$') { throw 'Configuration: Cloud.TenantId must be the tenant GUID.' }
    if ($cloud.Organization -and $cloud.Organization -notmatch '^[A-Za-z0-9.-]+\.[A-Za-z]{2,}$') { throw 'Configuration: Cloud.Organization must be the initial domain (tenant.onmicrosoft.com).' }
    if ($cloud.TenantId) { $cloud.TenantId = ([string]$cloud.TenantId).ToLowerInvariant() }
    $users = $config.Licensing.Users
    $official = Resolve-PcmOfficialValue ([string]$users.Mode) @('Group','Kiosk','Direct')
    if (-not $official) { throw 'Configuration: Licensing.Users.Mode must be Group, Kiosk or Direct.' }
    $users.Mode = $official
    if ($users.Mode -eq 'Group' -and $users.GroupId -and $users.GroupId -notmatch '^[0-9a-fA-F-]{36}$') { throw 'Configuration: Licensing.Users.GroupId must be the object ID (GUID) of the licence group.' }
    if ($users.GroupId) { $users.GroupId = ([string]$users.GroupId).ToLowerInvariant() }
    if ($users.Mode -eq 'Direct' -and -not $users.SkuPartNumber) { throw 'Configuration: Licensing.Users.Mode = Direct requires Licensing.Users.SkuPartNumber.' }
    $parallel = $config.Licensing.Shared.Parallel
    if ($parallel -isnot [int] -or $parallel -lt 1 -or $parallel -gt 1000) { throw 'Configuration: Licensing.Shared.Parallel must be a whole number between 1 and 1000 (temporary licences used at the same time by the shared mailboxes).' }
    $keepFree = $config.Licensing.Shared.KeepFree
    if ($keepFree -isnot [int] -or $keepFree -lt 0) { throw 'Configuration: Licensing.Shared.KeepFree must be a whole number, 0 or more (units of the shared mailbox licence that Convert never uses).' }
    $waveSize = $config.Recover.WaveSize
    if ($waveSize -isnot [int] -or $waveSize -lt 0 -or $waveSize -gt 1000) { throw 'Configuration: Recover.WaveSize must be a whole number between 1 and 1000 (shared mailboxes rolled back together), or 0 to use Licensing.Shared.Parallel.' }
    $retention = $config.Retention
    if ($retention.TagAttribute -notmatch '^CustomAttribute([1-9]|1[0-5])$') { throw 'Configuration: Retention.TagAttribute must be CustomAttribute1 to CustomAttribute15.' }
    $retention.TagAttribute = 'CustomAttribute' + $Matches[1]
    if (-not $retention.TagValue) { throw 'Configuration: Retention.TagValue is empty.' }
    $limit = $retention.HoldPolicyLimit
    if ($limit -isnot [int] -or $limit -lt 1 -or $limit -gt 1000) { throw 'Configuration: Retention.HoldPolicyLimit must be a whole number between 1 and 1000 (mailboxes per case hold policy; 1000 is the Microsoft Purview limit).' }
    try { $null = [version]$config.EntraConnect.MinVersion } catch { throw 'Configuration: EntraConnect.MinVersion must be a version number (2.5.76.0).' }
    if ($config.EntraConnect.Server -and $config.EntraConnect.Server -notmatch '^[A-Za-z0-9.-]+$') { throw 'Configuration: EntraConnect.Server must be a host name.' }
    $official = Resolve-PcmOfficialValue ([string]$config.EntraConnect.Mode) @('Remoting','Script','Manual')
    if (-not $official) { throw 'Configuration: EntraConnect.Mode must be Remoting, Script or Manual.' }
    $config.EntraConnect.Mode = $official
    if ($config.EntraConnect.Mode -eq 'Remoting' -and -not $config.EntraConnect.Server) { throw 'Configuration: EntraConnect.Mode = Remoting requires EntraConnect.Server.' }
    if ($config.EntraConnect.Mode -eq 'Script' -and -not $config.EntraConnect.ScriptPath) { throw 'Configuration: EntraConnect.Mode = Script requires EntraConnect.ScriptPath.' }
    if ($config.Cloud.DefaultUsageLocation -and $config.Cloud.DefaultUsageLocation -notmatch '^[A-Za-z]{2}$') { throw 'Configuration: Cloud.DefaultUsageLocation must be a two-letter country code (FR).' }
    if ($config.Cloud.DefaultUsageLocation) { $config.Cloud.DefaultUsageLocation = ([string]$config.Cloud.DefaultUsageLocation).ToUpperInvariant() }
    try { $null = [version]$config.Exo.MinModuleVersion } catch { throw 'Configuration: Exo.MinModuleVersion must be a version number (3.10.1).' }
    $polling = $config.Polling
    if ($polling.IntervalSeconds -lt 5 -or $polling.IntervalSeconds -gt 300) { throw 'Configuration: Polling.IntervalSeconds must be between 5 and 300.' }
    foreach ($name in @('MailboxTimeoutMinutes','SyncTimeoutMinutes','HoldTimeoutMinutes')) { if ($polling[$name] -lt 1 -or $polling[$name] -gt 720) { throw "Configuration: Polling.$name must be between 1 and 720." } }
    if ($polling.CallTimeoutMinutes -isnot [int] -or $polling.CallTimeoutMinutes -lt 1 -or $polling.CallTimeoutMinutes -gt 60) { throw 'Configuration: Polling.CallTimeoutMinutes must be a whole number between 1 and 60 (the minutes the tool waits for one answer of Microsoft 365 before it stops the run).' }

    $config.Store.Path = Resolve-PcmPath $config.Store.Path 'Store.Path' $Root
    $config.Store.BackupFolder = Resolve-PcmPath $config.Store.BackupFolder 'Store.BackupFolder' $Root
    $config.Store.JournalPath = Resolve-PcmPath $config.Store.JournalPath 'Store.JournalPath' $Root
    if ($config.EntraConnect.ScriptPath) { $config.EntraConnect.ScriptPath = Resolve-PcmPath $config.EntraConnect.ScriptPath 'EntraConnect.ScriptPath' $Root }
    if ($config.EntraConnect.CredentialPath) { $config.EntraConnect.CredentialPath = Resolve-PcmPath $config.EntraConnect.CredentialPath 'EntraConnect.CredentialPath' $Root }
    $config.Logging.Folder = Resolve-PcmPath $config.Logging.Folder 'Logging.Folder' $Root
    $config.Report.Folder = Resolve-PcmPath $config.Report.Folder 'Report.Folder' $Root
    $config.Store.Path = Assert-PcmOutputFolder $config.Store.Path 'Store.Path' 'data' $Root
    $config.Store.BackupFolder = Assert-PcmOutputFolder $config.Store.BackupFolder 'Store.BackupFolder' 'data' $Root
    $config.Store.JournalPath = Assert-PcmOutputFolder $config.Store.JournalPath 'Store.JournalPath' 'data' $Root
    $config.Logging.Folder = Assert-PcmOutputFolder $config.Logging.Folder 'Logging.Folder' 'logs' $Root
    $config.Report.Folder = Assert-PcmOutputFolder $config.Report.Folder 'Report.Folder' 'reports' $Root
    if ($config.Scope.CsvPath) { $config.Scope.CsvPath = Resolve-PcmPath $config.Scope.CsvPath 'Scope.CsvPath' $Root }
    $config['_Path'] = $file
    return $config
}
#endregion

#region 4. Audit ------------------------------------------------------------------------------

function Get-PcmArtifactStem {
    <# Unique, file-name safe prefix of the log, transcript and report files of this run. #>
    [CmdletBinding()]
    param([hashtable]$Context)
    $id = ([string]$Context.RunId -replace '[^a-zA-Z0-9_.-]', '_')
    if ($id.Length -gt 80) { $id = $id.Substring(0, 80) }
    return ('PRA_{0}_{1}' -f $id, [guid]::NewGuid().ToString('N').Substring(0, 8))
}

function Wait-PcmTranscriptFile {
    <#
    .SYNOPSIS
        Waits until the transcript file really exists and is not empty.
    .DESCRIPTION
        Start-Transcript can return before the file is created. The first
        console output starts it; this function waits for it (up to 5 s) instead of creating a
        replacement file. A missing or empty transcript stops the run (no audit, no processing).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$LiteralPath, [ValidateRange(1,60000)][int]$TimeoutMilliseconds = 5000)
    $clock = [Diagnostics.Stopwatch]::StartNew()
    do {
        if (Test-Path -LiteralPath $LiteralPath -PathType Leaf -ErrorAction Stop) {
            if ((Get-Item -LiteralPath $LiteralPath -ErrorAction Stop).Length -gt 0) { return }
        }
        if ($clock.ElapsedMilliseconds -ge $TimeoutMilliseconds) { break }
        Start-Sleep -Milliseconds 50
    } while ($true)
    throw "The transcript file is missing or empty after ${TimeoutMilliseconds} ms: $LiteralPath. No audit, no processing."
}

function Initialize-PcmFolder {
    <#
    .SYNOPSIS
        Makes sure a log or report folder exists. One failure, one message: the folder and the reason.
    .DESCRIPTION
        Not subject to -WhatIf / -Confirm: the audit trail of a Preview must be written too.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (Test-Path -LiteralPath $Path -PathType Container) { return }
    try {
        if (Test-Path -LiteralPath $Path) { throw 'A file exists where the folder is expected.' }
        [void][IO.Directory]::CreateDirectory($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path))
    }
    catch {
        $reason = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        throw ('The folder {0} cannot be created: {1}' -f $Path, $reason)
    }
}

function Initialize-PcmAudit {
    <#
    .SYNOPSIS
        Creates the log file of the run and starts the PowerShell transcript. Any failure stops the run.
    .DESCRIPTION
        Both files are created in $Context.LogFolder with a unique name; an existing file is never
        overwritten and a transcript started by someone else is left untouched. When reports are
        enabled, the report folder is also checked (a probe file is created and deleted).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $Context.LogFile = ''
    $Context['AuditAttempted'] = $true
    try {
        if ($null -ne $script:TranscriptOwner) { throw 'A transcript of this module is already running.' }
        $stem = Get-PcmArtifactStem $Context
        $Context['_PraArtifactStem'] = $stem
        $Context.CurrentOperation = 'Create-Log'
        Initialize-PcmFolder -Path $Context.LogFolder
        $logPath = Join-Path $Context.LogFolder ($stem + '.log')
        $null = New-Item -ItemType File -Path $logPath -WhatIf:$false -Confirm:$false -ErrorAction Stop
        $Context.LogFile = $logPath
        $transcript = Join-Path $Context.LogFolder ($stem + '.transcript.txt')
        $Context.TranscriptPath = $transcript
        $Context.CurrentOperation = 'Start-Transcript'; $Context['_PraTranscriptCreated'] = $false
        # Never Stop-Transcript first and never -Append: a transcript started by someone else stays intact.
        Start-Transcript -Path $transcript -NoClobber -WhatIf:$false -Confirm:$false -ErrorAction Stop | Out-Null
        $Context.TranscriptStarted = $true; $script:TranscriptOwner = $Context
        # The first host output makes PowerShell create the file.
        Write-PcmLog -Context $Context -Message ("Audit: host=$($Host.Name) PowerShell=$($PSVersionTable.PSVersion) transcript=$transcript") -Level Debug
        Write-Host ''
        $Context.CurrentOperation = 'Wait-TranscriptFile'
        Wait-PcmTranscriptFile -LiteralPath $transcript
        $Context['_PraTranscriptCreated'] = $true
        $Context.CurrentOperation = 'Check-ReportFolder'
        if (-not $Context.NoReport) {
            Initialize-PcmFolder -Path $Context.ReportFolder
            $probe = Join-Path $Context.ReportFolder ($stem + '.probe'); $created = $false
            try {
                $null = New-Item -ItemType File -Path $probe -WhatIf:$false -Confirm:$false -ErrorAction Stop; $created = $true
                Set-Content -LiteralPath $probe -Value 'PRA audit probe' -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop
            } finally { if ($created) { Remove-Item -LiteralPath $probe -WhatIf:$false -Confirm:$false -ErrorAction Stop } }
        }
        Write-PcmLog $Context ('Audit started: run {0}, version {1}, account {2}, computer {3}' -f $Context.RunId, $Context.Version, [Security.Principal.WindowsIdentity]::GetCurrent().Name, $env:COMPUTERNAME) -Level Detail
    } catch {
        Write-PcmLog $Context ('Audit could not start: {0}' -f $_.Exception.Message) -Level Error
        throw
    }
}
#endregion

#region 5. Results ----------------------------------------------------------------------------

function Get-PcmOutcome {
    <#
    .SYNOPSIS
        Result rows (one per object + one per run error not attached to an object), counters and exit code.
    .DESCRIPTION
        Exit code: 0 = done, 1 = failed (any error), 2 = done but at least one object is Pending
        (a next step is needed, e.g. the cloud phase or Finalize).
    #>
    [CmdletBinding()]
    param([hashtable]$Context)
    $rows = [System.Collections.Generic.List[object]]::new()
    $errorIdentities = @{}
    # Errors of the run itself (a stop, a failed read), apart from the objects found in error or not in the expected state.
    $runErrors = 0
    foreach ($row in $Context.Rows.ToArray()) {
        $copy = [ordered]@{}
        foreach ($field in $script:RowFields) { $copy[$field] = Get-PcmValue $row $field }
        if ($copy.FinalStatus -notin $script:FinalStatuses) {
            $copy.Detail = 'Final status missing or invalid: {0}. {1}' -f $copy.FinalStatus, $copy.Detail; $copy.FinalStatus = 'Error'
        }
        if ($copy.FinalStatus -eq 'Error') {
            foreach ($field in @('ObjectGuid','Identity','PrimarySmtpAddress')) { if ($copy[$field]) { $errorIdentities[[string]$copy[$field]] = $true } }
        }
        [void]$rows.Add([pscustomobject]$copy)
    }
    foreach ($issue in $Context.Issues.ToArray()) {
        $identity = [string](Get-PcmValue $issue 'Identity' '')
        $source = [string](Get-PcmValue $issue 'Source' 'Execution')
        # The error of an object is its row: the cloud error and the stop that follow it are the same failure, not two more.
        if ($source -in @('Execution', 'Cloud') -and $identity -and $errorIdentities.ContainsKey($identity)) { continue }
        $failure = [ordered]@{}
        foreach ($field in $script:RowFields) { $failure[$field] = $null }
        $failure.Action = $Context.Action; $failure.Identity = $identity
        $failure.FinalStatus = 'Error'
        $failure.Detail = '[Run/{0}{1}] {2}' -f $source, $(if (Get-PcmValue $issue 'Operation' '') { '/' + (Get-PcmValue $issue 'Operation' '') } else { '' }), (Get-PcmValue $issue 'Message' ([string]$issue))
        [void]$rows.Add([pscustomobject]$failure); $runErrors++
    }
    $counts = @{ Success = 0; Error = 0; Pending = 0; Skipped = 0; Planned = 0; RunError = $runErrors }
    foreach ($row in $rows) {
        $key = [string]$row.FinalStatus
        if ($key -eq 'AlreadyDone') { $key = 'Success' }
        $counts[$key]++
    }
    # The members of a selected group that are not taken have no row: they are skipped objects of the run.
    $counts.Skipped += @(@(Get-PcmValue $Context 'SkippedMembers' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ }).Count
    $requested = [string](Get-PcmValue $Context 'ResultStatus' '')
    $code = 0; $status = 'Success'
    # An operator who answered no to the confirmation cancelled the run: nothing was written and no error is counted (exit code 1, the work was not done).
    if ([bool](Get-PcmValue $Context 'Cancelled' $false) -and $counts.Error -eq 0 -and $Context.Issues.Count -eq 0 -and -not $Context.CloudConnectFailed) { $code = 1; $status = 'Cancelled' }
    elseif ($counts.Error -gt 0 -or $Context.Issues.Count -gt 0 -or [int]$Context.ExitCode -eq 1 -or $Context.CloudConnectFailed -or $requested -eq 'Failed') {
        $code = 1; $status = 'Failed'
    } elseif ($counts.Pending -gt 0) { $code = 2; $status = 'Pending' }
    elseif ($Context.Mode -ne 'Apply' -and $counts.Planned -gt 0) { $status = 'Planned' }
    return [pscustomobject]@{ Rows = $rows.ToArray(); Counts = $counts; Status = $status; ExitCode = $code }
}

function Get-PcmVerdict {
    <# The verdict line of a run: the last line of the log and of the transcript. #>
    param($Outcome)
    if (Test-PcmSameText ([string]$Outcome.Status) 'Cancelled') { return 'RESULT: CANCELLED' }
    if ($Outcome.ExitCode -eq 1) { return 'RESULT: FAIL' }
    if ($Outcome.ExitCode -eq 2) { return 'RESULT: PENDING' }
    return 'RESULT: PASS'
}

function ConvertTo-PcmHtmlText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return '' }
    return [Net.WebUtility]::HtmlEncode([string]$Value)
}

$script:ReportChunkRows = 20000
$script:ReportChunkChars = 4000000
$script:JsonEscape = New-Object Text.RegularExpressions.Regex('["\\\u0000-\u001f\u007f-\u009f  <>&]')
$script:JsonEscapeEvaluator = [Text.RegularExpressions.MatchEvaluator]{
    param($m)
    # On the character CODE, never on the string: PowerShell string comparison is culture-aware and, with the ICU collation of .NET on PowerShell 7,
    # ignores control characters, which would write \b for U+0001.
    $code = [int][char]$m.Value
    switch ($code) {
        34 { '\"' } 92 { '\\' } 10 { '\n' } 13 { '\r' } 9 { '\t' } 8 { '\b' } 12 { '\f' }
        default { '\u' + $code.ToString('x4') }
    }
}

function ConvertTo-PcmJsonString {
    <# JSON string literal of a text. Quotes, backslashes, control characters, the HTML-significant characters < > & and the line separators are escaped: safe in a script element. #>
    [CmdletBinding()]
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return 'null' }
    $Text = [string]$Value
    if ($script:JsonEscape.IsMatch($Text)) { return '"' + $script:JsonEscape.Replace($Text, $script:JsonEscapeEvaluator) + '"' }
    return '"' + $Text + '"'
}

function Get-PcmReportScalar {
    <# Keeps strings, booleans and finite numbers as they are and converts any other value to its text. #>
    [CmdletBinding()]
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or $Value -is [string] -or $Value -is [bool]) { return $Value }
    if ($Value -is [int] -or $Value -is [long]) { return $Value }
    if (($Value -is [double] -or $Value -is [single]) -and -not ([double]::IsNaN([double]$Value) -or [double]::IsInfinity([double]$Value))) { return $Value }
    if ($Value -is [decimal]) { return $Value }
    return [string]$Value
}

function ConvertTo-PcmJsonScalar {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [string]) { return (ConvertTo-PcmJsonString $Value) }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [double] -or $Value -is [single]) { return ([double]$Value).ToString('R', [Globalization.CultureInfo]::InvariantCulture) }
    return ([IFormattable]$Value).ToString($null, [Globalization.CultureInfo]::InvariantCulture)
}

function Write-PcmReportChunks {
    <#
    .SYNOPSIS
        Writes items as gzip+base64 chunks of dictionary-indexed columns (script elements of type application/x-pra-chunk). Each chunk holds at most ChunkRows items and
        about ChunkChars characters of the long text fields, so memory stays bounded whatever the number of items. Dictionary index 0 is null. Returns the number of chunks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][IO.TextWriter]$Writer, [Parameter(Mandatory)][string]$Kind,
        [AllowEmptyCollection()][object[]]$Items, [Parameter(Mandatory)][string[]]$Fields, [string[]]$SizeFields = @(),
        [int]$ChunkRows = $script:ReportChunkRows, [int]$ChunkChars = $script:ReportChunkChars
    )
    $dictionaries = @(for ($f = 0; $f -lt $Fields.Count; $f++) { , (New-Object 'Collections.Generic.Dictionary[object,int]') })
    $utf8 = New-Object Text.UTF8Encoding($false)
    $jsonEscape = $script:JsonEscape; $jsonEvaluator = $script:JsonEscapeEvaluator
    $position = 0; $chunks = 0
    while ($position -lt $Items.Count) {
        $end = $position; $chars = 0
        while ($end -lt $Items.Count -and ($end - $position) -lt $ChunkRows) {
            $item = $Items[$end]
            foreach ($sizeField in $SizeFields) { $text = $item.$sizeField; if ($text -is [string]) { $chars += $text.Length } }
            if ($chars -gt $ChunkChars -and $end -gt $position) { break }
            $end++
        }
        $count = $end - $position
        $slice = New-Object object[] $count
        [Array]::Copy($Items, $position, $slice, 0, $count)
        $json = New-Object Text.StringBuilder
        $indexes = New-Object Text.StringBuilder
        [void]$json.Append('{"n":').Append($count).Append(',"d":[')
        for ($f = 0; $f -lt $Fields.Count; $f++) {
            $name = $Fields[$f]; $dictionary = $dictionaries[$f]
            $column = New-Object int[] $count
            $added = New-Object 'Collections.Generic.List[string]'
            if ($chunks -eq 0) { $added.Add('null') }
            $k = 0
            for ($i = 0; $i -lt $count; $i++) {
                $value = $slice[$i].$name
                if ($null -eq $value) { continue }
                if ($value -isnot [string] -and $value -isnot [bool]) { $value = Get-PcmReportScalar $value }
                if ($dictionary.TryGetValue($value, [ref]$k)) { $column[$i] = $k }
                else {
                    $k = $dictionary.Count + 1
                    $dictionary.Add($value, $k)
                    if ($value -is [string]) {
                        if ($jsonEscape.IsMatch($value)) { $added.Add('"' + $jsonEscape.Replace($value, $jsonEvaluator) + '"') } else { $added.Add('"' + $value + '"') }
                    } elseif ($value -is [bool]) { if ($value) { $added.Add('true') } else { $added.Add('false') } }
                    else { $added.Add((ConvertTo-PcmJsonScalar $value)) }
                    $column[$i] = $k
                }
            }
            if ($f) { [void]$json.Append(','); [void]$indexes.Append(',') }
            [void]$json.Append('[').Append([string]::Join(',', $added)).Append(']')
            [void]$indexes.Append('[').Append([string]::Join(',', $column)).Append(']')
        }
        [void]$json.Append('],"c":[').Append($indexes.ToString()).Append(']}')
        $memory = New-Object IO.MemoryStream
        $gzip = New-Object IO.Compression.GZipStream($memory, [IO.Compression.CompressionMode]::Compress, $true)
        try { $bytes = $utf8.GetBytes($json.ToString()); $gzip.Write($bytes, 0, $bytes.Length) } finally { $gzip.Dispose() }
        $Writer.Write('<script type="application/x-pra-chunk" data-kind="' + $Kind + '" data-seq="' + $chunks + '">')
        $Writer.Write([Convert]::ToBase64String($memory.ToArray()))
        $Writer.Write("</script>`r`n")
        $memory.Dispose()
        $chunks++; $position = $end
    }
    return $chunks
}

function Write-PcmReportHtml {
    <#
    .SYNOPSIS
        Streams the HTML report (templates\Report.template.html) to a writer. The header, the tiles and the next steps are HTML-encoded markup; the object rows (the fields
        of the CSV), the run errors and the warnings are written by Write-PcmReportChunks and read by the script of the page. Used by Write-PcmReport (file) and
        Get-PcmReportHtml (string).
    #>
    [CmdletBinding()]
    param([hashtable]$Context, $Outcome, [Parameter(Mandatory)][IO.TextWriter]$Writer,
        [int]$ChunkRows = $script:ReportChunkRows, [int]$ChunkChars = $script:ReportChunkChars)
    $templatePath = Join-Path $Context.Root 'templates\Report.template.html'
    $template = if (Test-Path -LiteralPath $templatePath -PathType Leaf) { [IO.File]::ReadAllText($templatePath, [Text.Encoding]::UTF8) }
        else { '<!doctype html><html><head><meta charset="utf-8"><title>{{TITLE}}</title></head><body><h1>{{TITLE}}</h1><p>{{SUBTITLE}}</p>{{TILES}}{{NEXT}}<p>{{DATA_NOTICE}} Template missing: only the CSV file holds the rows.</p></body></html>' }
    $c = $Outcome.Counts
    $warningList = @(@(Get-PcmValue $Context 'WarningList' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ })
    $issueList = @($Context.Issues.ToArray())
    $statusClass = @{ Success='ok'; Planned='info'; Pending='warn'; Cancelled='warn'; Failed='fail' }[$Outcome.Status]
    $tiles = New-Object Text.StringBuilder
    foreach ($tile in @(
            @('Result', $Outcome.Status, $statusClass), @('Objects', $Outcome.Rows.Count, ''), @('Done', $c.Success, 'ok'),
            @('Pending', $c.Pending, $(if ($c.Pending) { 'warn' } else { '' })), @('Errors', $c.Error, $(if ($c.Error) { 'fail' } else { '' })),
            @('Warnings', $warningList.Count, $(if ($warningList.Count) { 'warn' } else { '' })),
            @('Planned', $c.Planned, $(if ($c.Planned) { 'info' } else { '' })), @('Skipped', $c.Skipped, ''))) {
        [void]$tiles.AppendFormat('<div class="tile {2}"><div class="tile-label">{0}</div><div class="tile-value">{1}</div></div>', (ConvertTo-PcmHtmlText $tile[0]), (ConvertTo-PcmHtmlText $tile[1]), $tile[2])
    }
    $modeText = if ($Context.Action -eq 'Check') { 'Read-only' } elseif ($Context.Mode -eq 'Apply') { 'Apply' } else { 'Preview - nothing changed' }
    $skipped = @(@(Get-PcmValue $Context 'SkippedMembers' @()) | ForEach-Object { $_ } | ForEach-Object { [string]$_ })
    $overrides = @(@(Get-PcmValue $Context 'ConfigOverrides' @()) | ForEach-Object { $_ } | ForEach-Object { [string]$_ })
    $meta = [ordered]@{
        'Action' = ('{0} · {1}' -f $Context.Action, $modeText)
        'Environment' = [string](Get-PcmValue $Context.Config 'Environment' '')
        'Objects' = [string](Get-PcmValue $Context 'ObjectsText' '')
        'Skipped members' = $(if ($skipped.Count) { '{0} member(s) not taken: {1}{2}' -f $skipped.Count, (($skipped | Select-Object -First 50) -join '; '), $(if ($skipped.Count -gt 50) { '; and {0} more, see the log' -f ($skipped.Count - 50) } else { '' }) } else { '' })
        'Batch' = [string](Get-PcmValue $Context 'BatchId' '')
        'Snapshot' = [string](Get-PcmValue $Context 'SnapshotLabel' '')
        'Snapshot check' = [string](Get-PcmValue $Context 'SnapshotCheck' '')
        'Exchange server' = [string](Get-PcmValue $Context 'Server' '')
        'Tenant' = [string](Get-PcmValue (Get-PcmValue $Context.Config 'Cloud' @{}) 'Organization' '')
        'Computer / account' = ('{0} · {1}\{2}' -f $env:COMPUTERNAME, $env:USERDOMAIN, $env:USERNAME)
        'Started' = ([datetime]$Context.StartTime).ToString('yyyy-MM-dd HH:mm:ss')
        'Duration' = (Format-PcmDuration ((Get-Date) - [datetime]$Context.StartTime).TotalSeconds)
        'Run ID' = [string]$Context.RunId
        'Run settings changed' = ($overrides -join '; ')
    }
    $metaHtml = New-Object Text.StringBuilder
    foreach ($key in $meta.Keys) { if ($meta[$key]) { [void]$metaHtml.AppendFormat('<div><span>{0}</span>{1}</div>', (ConvertTo-PcmHtmlText $key), (ConvertTo-PcmHtmlText $meta[$key])) } }
    $next = @(@(Get-PcmValue $Context 'NextSteps' @()) | ForEach-Object { $_ } | Where-Object { $_ })
    $nextHtml = ''
    if ($next.Count) { $nextHtml = '<section class="card next"><h2>Next step</h2>' + (($next | ForEach-Object { '<pre>' + (ConvertTo-PcmHtmlText $_) + '</pre>' }) -join '') + '</section>' }
    $notice = '{0} object row(s), {1} run error(s) and {2} warning(s) are stored in this file.' -f $Outcome.Rows.Count, $issueList.Count, $warningList.Count
    $values = @{
        '{{TITLE}}' = ConvertTo-PcmHtmlText ('PRA Cloud Mailbox - {0}' -f $Context.Action)
        '{{SUBTITLE}}' = ConvertTo-PcmHtmlText ('{0} · {1}' -f $modeText, ([datetime]$Context.StartTime).ToString('yyyy-MM-dd HH:mm'))
        '{{VERSION}}' = ConvertTo-PcmHtmlText $Context.Version
        '{{META}}' = $metaHtml.ToString(); '{{TILES}}' = $tiles.ToString(); '{{NEXT}}' = $nextHtml
        '{{DATA_NOTICE}}' = ConvertTo-PcmHtmlText $notice
        '{{ROWS_TABLE}}' = ''; '{{ISSUES}}' = ''; '{{WARNINGS}}' = ''
    }
    foreach ($key in $values.Keys) { $template = $template.Replace($key, $values[$key]) }
    $marker = '{{DATA}}'
    $at = $template.LastIndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($at -lt 0) { $Writer.Write($template); return }
    $Writer.Write($template.Substring(0, $at))

    $issueItems = foreach ($issue in $issueList) {
        [pscustomobject]@{ Timestamp = [string](Get-PcmValue $issue 'Timestamp' ''); Phase = [string](Get-PcmValue $issue 'Phase' ''); Operation = [string](Get-PcmValue $issue 'Operation' '')
            Identity = [string](Get-PcmValue $issue 'Identity' ''); Source = [string](Get-PcmValue $issue 'Source' ''); Message = [string](Get-PcmValue $issue 'Message' ([string]$issue)) }
    }
    $warningItems = foreach ($warning in $warningList) {
        [pscustomobject]@{ Timestamp = [string](Get-PcmValue $warning 'Timestamp' ''); Phase = [string](Get-PcmValue $warning 'Phase' ''); Message = [string](Get-PcmValue $warning 'Message' ([string]$warning)) }
    }
    $issueFields = @('Timestamp','Phase','Operation','Identity','Source','Message'); $warningFields = @('Timestamp','Phase','Message')
    $rowItems = [object[]]@($Outcome.Rows)
    $rowChunks = Write-PcmReportChunks -Writer $Writer -Kind 'rows' -Items $rowItems -Fields $script:RowFields -SizeFields @('Detail','Permissions','Warnings') -ChunkRows $ChunkRows -ChunkChars $ChunkChars
    $issueChunks = Write-PcmReportChunks -Writer $Writer -Kind 'issues' -Items ([object[]]@($issueItems)) -Fields $issueFields -SizeFields @('Message') -ChunkRows $ChunkRows -ChunkChars $ChunkChars
    $warningChunks = Write-PcmReportChunks -Writer $Writer -Kind 'warnings' -Items ([object[]]@($warningItems)) -Fields $warningFields -SizeFields @('Message') -ChunkRows $ChunkRows -ChunkChars $ChunkChars
    $statuses = ($script:FinalStatuses | ForEach-Object { ConvertTo-PcmJsonString $_ }) -join ','
    $fieldsJson = ($script:RowFields | ForEach-Object { ConvertTo-PcmJsonString $_ }) -join ','
    $json = '{{"format":1,"statuses":[{0}],"rows":{{"count":{1},"fields":[{2}]}},"issues":{{"count":{3},"fields":[{4}],"titles":["Time","Step","Operation","Object","Source","Message"]}},"warnings":{{"count":{5},"fields":[{6}],"titles":["Time","Step","Message"]}},"chunks":{{"rows":{7},"issues":{8},"warnings":{9}}}}}' -f
        $statuses, $rowItems.Count, $fieldsJson, $issueList.Count, (($issueFields | ForEach-Object { ConvertTo-PcmJsonString $_ }) -join ','), $warningList.Count, (($warningFields | ForEach-Object { ConvertTo-PcmJsonString $_ }) -join ','), $rowChunks, $issueChunks, $warningChunks
    $Writer.Write('<script type="application/json" id="pra-meta">' + $json + "</script>`r`n")
    $Writer.Write($template.Substring($at + $marker.Length))
}

function Get-PcmReportHtml {
    <#
    .SYNOPSIS
        Returns the HTML report as one string (small reports, tests). It uses the same writer as the report file, so there is a single serializer; production code
        streams to the file instead (Write-PcmReport).
    #>
    [CmdletBinding()]
    param([hashtable]$Context, $Outcome)
    $writer = New-Object IO.StringWriter
    try { Write-PcmReportHtml -Context $Context -Outcome $Outcome -Writer $writer; return $writer.ToString() } finally { $writer.Dispose() }
}

function Read-PcmReportData {
    <# Reads the rows, run errors and warnings stored in an HTML report, as the page script does (tests, support). Returns Rows, Issues and Warnings as lists of objects. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Html)
    $meta = [regex]::Match($Html, '<script type="application/json" id="pra-meta">(.*?)</script>', [Text.RegularExpressions.RegexOptions]::Singleline)
    if (-not $meta.Success) { throw 'The report holds no data header.' }
    $header = $meta.Groups[1].Value | ConvertFrom-Json
    $result = @{ rows = [System.Collections.Generic.List[object]]::new(); issues = [System.Collections.Generic.List[object]]::new(); warnings = [System.Collections.Generic.List[object]]::new() }
    $dictionaries = @{}
    foreach ($chunk in [regex]::Matches($Html, '<script type="application/x-pra-chunk" data-kind="(\w+)" data-seq="(\d+)">([^<]*)</script>')) {
        $kind = $chunk.Groups[1].Value
        $bytes = [Convert]::FromBase64String($chunk.Groups[3].Value)
        $reader = New-Object IO.StreamReader((New-Object IO.Compression.GZipStream((New-Object IO.MemoryStream(, $bytes)), [IO.Compression.CompressionMode]::Decompress)), [Text.Encoding]::UTF8)
        try { $data = $reader.ReadToEnd() | ConvertFrom-Json } finally { $reader.Dispose() }
        $fields = @($header.$kind.fields)
        if (-not $dictionaries.ContainsKey($kind)) { $dictionaries[$kind] = @(for ($f = 0; $f -lt $fields.Count; $f++) { , (New-Object 'Collections.Generic.List[object]') }) }
        for ($f = 0; $f -lt $fields.Count; $f++) { foreach ($value in @($data.d[$f])) { $dictionaries[$kind][$f].Add($value) } }
        for ($i = 0; $i -lt $data.n; $i++) {
            $row = [ordered]@{}
            for ($f = 0; $f -lt $fields.Count; $f++) { $row[$fields[$f]] = $dictionaries[$kind][$f][[int]$data.c[$f][$i]] }
            $result[$kind].Add([pscustomobject]$row)
        }
    }
    return [pscustomobject]@{ Rows = $result.rows.ToArray(); Issues = $result.issues.ToArray(); Warnings = $result.warnings.ToArray(); Header = $header }
}

function Write-PcmReport {
    <#
    .SYNOPSIS
        Writes the CSV and HTML reports (also when there are zero rows). A failed format is recorded as
        a run error; the other format is still written, then refreshed to include that error.
    #>
    [CmdletBinding()]
    param([hashtable]$Context, [hashtable]$Reports)
    for ($attempt = 0; $attempt -lt 2; $attempt++) {
        $issueCount = $Context.Issues.Count
        foreach ($format in @('Csv','Html')) {
            if (-not $Reports[$format] -or $Reports[($format + 'Failed')]) { continue }
            try {
                $outcome = Get-PcmOutcome $Context
                if ($format -eq 'Csv') {
                    if ($outcome.Rows.Count) { $outcome.Rows | Export-Csv -LiteralPath $Reports.Csv -NoTypeInformation -Delimiter ';' -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop }
                    else { Set-Content -LiteralPath $Reports.Csv -Value ('"' + ($script:RowFields -join '";"') + '"') -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop }
                } else {
                    $stream = [IO.File]::Open($Reports.Html, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
                    $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($true)), 65536)
                    try { Write-PcmReportHtml -Context $Context -Outcome $outcome -Writer $writer } finally { $writer.Dispose() }
                }
                if (-not (Test-Path -LiteralPath $Reports[$format] -PathType Leaf) -or (Get-Item -LiteralPath $Reports[$format]).Length -eq 0) { throw 'Report missing or empty after writing.' }
            } catch { $Reports[($format + 'Failed')] = $true; Add-PcmIssue $Context ('{0} report: {1}' -f $format, $_.Exception.Message) 'Report' }
        }
        if ($Context.Issues.Count -eq $issueCount) { break }
    }
}

function Complete-PcmRun {
    <#
    .SYNOPSIS
        Ends every run: reports, summary card, transcript, verdict. Returns one result object (no exit).
    .DESCRIPTION
        Always called from the finally block of the entry script, also after an error. The verdict
        'RESULT: PASS' or 'RESULT: FAIL' is the last line of the log and of the transcript; it is
        added after the transcript is stopped so that a failure to stop it is still reported.
    .OUTPUTS
        PSCustomObject: RunId, Status (Success, Planned, Pending, Failed), ExitCode, counters,
        BatchId, NextSteps, file paths and Issues.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context)
    $reports = @{ Csv = ''; Html = ''; CsvFailed = $false; HtmlFailed = $false }
    $paths = @{ LogFile = ''; TranscriptPath = ''; BackupFiles = @(); StateFiles = @() }
    $safeLog = {
        param($message, $level = 'Detail')
        try { Write-PcmLog $Context $message -Level $level }
        catch { $console = $Context.Clone(); $console.LogFile = ''; Write-PcmLog $console $message -Level $level }
    }
    $ownsTranscript = [object]::ReferenceEquals($script:TranscriptOwner, $Context)
    if (([int]$Context.ExitCode -ne 0 -or $Context.CloudConnectFailed -or (Get-PcmValue $Context 'ResultStatus' '') -eq 'Failed') -and $Context.Issues.Count -eq 0 -and -not [bool](Get-PcmValue $Context 'Cancelled' $false)) {
        Add-PcmIssue $Context 'Failure reported by the run (exit code / cloud connection).' 'Run'
    }
    # No stack line reaches the card, the console or the report: the log holds the whole detail.
    foreach ($entry in @($Context.Issues.ToArray()) + @(@(Get-PcmValue $Context 'WarningList' @()) | ForEach-Object { $_ })) {
        if ($entry -is [pscustomobject] -and $entry.PSObject.Properties['Message']) { $plain = ConvertTo-PcmPlainMessage ([string]$entry.Message); if ($plain -ne [string]$entry.Message) { $entry.Message = $plain } }
    }
    foreach ($key in @('LogFile','TranscriptPath','BackupFiles','StateFiles')) {
        $candidates = Get-PcmValue $Context $key
        if ($key -in @('LogFile','TranscriptPath')) { $candidates = [string]$candidates }
        if ($key -eq 'TranscriptPath' -and -not (Get-PcmValue $Context '_PraTranscriptCreated' $false)) { $candidates = '' }
        foreach ($path in $candidates) {
            # An empty log or transcript path means the audit never started: that cause is already an issue of the run.
            if (-not $path -and $key -in @('LogFile','TranscriptPath') -and @($Context.Issues.ToArray() | Where-Object { (Test-PcmSameText ([string](Get-PcmValue $_ 'Source' '')) 'Audit') -or ([string](Get-PcmValue $_ 'Message' '')).StartsWith('Audit ', [StringComparison]::OrdinalIgnoreCase) }).Count) { continue }
            try {
                if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'file not found' }
                $file = Get-Item -LiteralPath $path -ErrorAction Stop
                if ($file.Length -eq 0 -and $key -ne 'StateFiles') { throw 'empty file' }
                if ($key -in @('BackupFiles','StateFiles')) { $paths[$key] += $file.FullName } else { $paths[$key] = $file.FullName }
            } catch { Add-PcmIssue $Context ('Audit file {0} unavailable: {1}' -f $key, $_.Exception.Message) 'Audit' }
        }
    }
    if (-not $Context.NoReport) {
        try {
            Initialize-PcmFolder -Path $Context.ReportFolder
            $stem = [string](Get-PcmValue $Context '_PraArtifactStem' '')
            if (-not $stem) { $stem = Get-PcmArtifactStem $Context }
            foreach ($format in @('Csv','Html')) {
                try {
                    $path = Join-Path $Context.ReportFolder ($stem + '.' + $format.ToLowerInvariant())
                    $null = New-Item -ItemType File -Path $path -WhatIf:$false -Confirm:$false -ErrorAction Stop
                    $reports[$format] = $path
                } catch { $reports[($format + 'Failed')] = $true; Add-PcmIssue $Context ('{0} report cannot be created: {1}' -f $format, $_.Exception.Message) 'Report' }
            }
        } catch { Add-PcmIssue $Context ('Report folder: {0}' -f $_.Exception.Message) 'Report' }
        Write-PcmReport $Context $reports
    }

    # Summary card.
    $outcome = Get-PcmOutcome $Context
    $c = $outcome.Counts
    $isApply = $Context.Mode -eq 'Apply'
    # A Preview (not a Check, which has its own words) changes nothing: its objects are "ready" or "planned", never "done".
    $isCheckRun = $Context.Action -eq 'Check'
    $isPreviewRun = (-not $isApply) -and (-not $isCheckRun)
    # A Check that ran to the end and found objects not in the expected state is finished; it "stopped" only when a run error prevented it.
    $checkRanToEnd = $isCheckRun -and $outcome.Status -eq 'Failed' -and [bool](Get-PcmValue $Context 'CheckFinished' $false) -and $c.RunError -eq 0 -and $c.Error -gt 0
    $title = switch ($outcome.Status) {
        'Failed' { if ($checkRanToEnd) { 'Check finished - {0} object(s) not in the expected state' -f $c.Error } else { '{0} stopped - see the errors below' -f $Context.Action } }
        'Cancelled' { '{0} cancelled - nothing was written' -f $Context.Action }
        'Pending' { '{0} done - next step required' -f $Context.Action }
        default { if ($isPreviewRun) { '{0} preview finished - nothing was changed' -f $Context.Action } else { '{0} finished' -f $Context.Action } }
    }
    $status = @{ Failed='Fail'; Pending='Warn'; Cancelled='Warn'; Planned='Ok'; Success='Ok' }[$outcome.Status]
    $dot = [char]0x00B7
    # A Check counts "not as expected" only for objects that were read and compared; a run error is an error.
    $runErrorText = if ($c.RunError) { '{0} error(s) {1} ' -f $c.RunError, $dot } else { '' }
    $objects = if ($isCheckRun) { '{0} as expected {3} {1} not as expected {3} {4}{2} skipped' -f $c.Success, ($c.Error - $c.RunError), $c.Skipped, $dot, $runErrorText }
        elseif ($isPreviewRun) { '{0} ready {5} {1} planned {5} {4}{2} error(s) {5} {3} skipped' -f $c.Success, $c.Planned, $c.Error, $c.Skipped, $(if ($c.Pending) { '{0} pending {1} ' -f $c.Pending, $dot } else { '' }), $dot }
        elseif ($isApply -or $c.Success -or $c.Pending) { '{0} done {4} {1} pending {4} {2} error(s) {4} {3} skipped' -f $c.Success, $c.Pending, $c.Error, $c.Skipped, $dot }
        else { '{0} planned {3} {1} error(s) {3} {2} skipped' -f $c.Planned, $c.Error, $c.Skipped, $dot }
    $values = [ordered]@{ 'Result' = @('Target', $objects) }
    # The rows that an action adds to its own card, after the result (label = @(IconName, Text)).
    $actionRows = Get-PcmValue $Context 'SummaryRows' $null
    if ($actionRows -is [Collections.IDictionary]) { foreach ($key in $actionRows.Keys) { $values[[string]$key] = $actionRows[$key] } }
    $batch = [string](Get-PcmValue $Context 'BatchId' '')
    if ($batch) { $values['Batch'] = @('Batch', $batch) }
    $snapshot = [string](Get-PcmValue $Context 'SnapshotLabel' '')
    if ($snapshot) { $values['Snapshot'] = @('Batch', $snapshot) }
    $snapshotCheck = [string](Get-PcmValue $Context 'SnapshotCheck' '')
    if ($snapshotCheck) { $values['Integrity'] = @('Key', $snapshotCheck) }
    $objectsText = [string](Get-PcmValue $Context 'ObjectsText' '')
    if ($objectsText) { $values['Objects'] = @('Target', $objectsText) }
    $overrides = @(@(Get-PcmValue $Context 'ConfigOverrides' @()) | ForEach-Object { $_ } | Where-Object { $_ })
    if ($overrides.Count) {
        $overrideLines = @($overrides | Select-Object -First 4)
        if ($overrides.Count -gt 4) { $overrideLines += ('... and {0} more, see the report' -f ($overrides.Count - 4)) }
        $values['Settings'] = @('Key', (@('{0} setting(s) of this run differ from the configuration file:' -f $overrides.Count) + $overrideLines))
    }
    # The errors of the objects come first (at most 5, then how many more); the line that says where the run stopped is the last one.
    $issueTexts = @($Context.Issues.ToArray() | ForEach-Object { [string](Get-PcmValue $_ 'Message' ([string]$_)) })
    $stopAt = -1
    for ($i = $issueTexts.Count - 1; $i -ge 0; $i--) { if ($issueTexts[$i].StartsWith("Stopped at step '", [StringComparison]::OrdinalIgnoreCase)) { $stopAt = $i; break } }
    $objectTexts = @(for ($i = 0; $i -lt $issueTexts.Count; $i++) { if ($i -ne $stopAt) { $issueTexts[$i] } })
    $issueLines = @($objectTexts | Select-Object -First 5)
    if ($objectTexts.Count -gt 5) { $issueLines += ('... and {0} more, see the report' -f ($objectTexts.Count - 5)) }
    if ($stopAt -ge 0) { $issueLines += $issueTexts[$stopAt] }
    $issueLines = @($issueLines | ForEach-Object { if ($_.Length -gt 240) { $_.Substring(0, 237) + '...' } else { $_ } })
    if ($issueLines.Count) {
        # The same count as the Result line: an object that failed and the stop that follows its error are one error, not two.
        if ($Context.Issues.Count -gt 2) { $issueLines = @(('{0} error(s) in total, see the log and the report' -f $(if ($isCheckRun) { $c.RunError } else { $c.Error }))) + $issueLines }
        $values['Error'] = @('Fail', $issueLines)
    }
    $warningList = @(@(Get-PcmValue $Context 'WarningList' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ })
    if ($warningList.Count) {
        $warningLines = @($warningList | Select-Object -Last 2 | ForEach-Object {
                $message = [string](Get-PcmValue $_ 'Message' ([string]$_))
                if ($message.Length -gt 240) { $message = $message.Substring(0, 237) + '...' }
                $message })
        if ($warningList.Count -gt 2) { $warningLines = @(('{0} warning(s) in total, see the log and the report' -f $warningList.Count)) + $warningLines }
        $values['Warning'] = @('Warn', $warningLines)
        # Warnings never change the exit code; they only colour the card of a successful run.
        if ($status -eq 'Ok') { $status = 'Warn'; $title += (' {0} {1} warning(s)' -f [char]0x00B7, $warningList.Count) }
    }
    if ($paths.BackupFiles.Count) { $values['Backup'] = @('Backup', @($paths.BackupFiles)) }
    if ($reports.Html -and -not $reports.HtmlFailed) { $values['Report'] = @('Report', $reports.Html) }
    if ($paths.LogFile) { $values['Log'] = @('Log', $paths.LogFile) }
    $values['Duration'] = @('Clock', (Format-PcmDuration ((Get-Date) - [datetime]$Context.StartTime).TotalSeconds))
    $next = @(@(Get-PcmValue $Context 'NextSteps' @()) | ForEach-Object { $_ } | Where-Object { $_ })
    # A failed run always ends with a next step: the one the run set on purpose, else the one proposed by the entry script (what was kept, what to run again).
    $showNext = ($outcome.ExitCode -ne 1 -or [bool](Get-PcmValue $Context 'NoChangeMade' $false) -or [bool](Get-PcmValue $Context 'NextOnFailure' $false))
    if ($outcome.ExitCode -eq 1 -and (-not $showNext -or -not $next.Count)) {
        $fallback = @(@(Get-PcmValue $Context 'FallbackNext' @()) | ForEach-Object { $_ } | Where-Object { $_ })
        if ($fallback.Count) { $next = $fallback; $Context.NextSteps = $fallback; $showNext = $true }
    }
    if ($showNext -and $next.Count) { $values['Next'] = @('Next', $next) }
    try { Write-PcmSummary -Context $Context -Title $title -Values $values -Status $status }
    catch { & $safeLog ('Summary unavailable: ' + $_.Exception.Message) }
    foreach ($key in @('BackupFiles','StateFiles')) { foreach ($path in $paths[$key]) { & $safeLog ('{0}: {1}' -f $key, $path) } }
    & $safeLog ('Status {0}; total {1}; done {2}; errors {3}; pending {4}; skipped {5}; planned {6}.' -f $outcome.Status, $outcome.Rows.Count, $c.Success, $c.Error, $c.Pending, $c.Skipped, $c.Planned)

    # Stop our own transcript (never someone else's), then write the verdict.
    if ($ownsTranscript) {
        try { Stop-Transcript -ErrorAction Stop | Out-Null }
        catch { Add-PcmIssue $Context ('Transcript could not be stopped: {0}' -f $_.Exception.Message) 'Audit' }
        finally { $Context.TranscriptStarted = $false; $script:TranscriptOwner = $null }
    }
    Write-PcmReport $Context $reports
    $outcome = Get-PcmOutcome $Context
    $verdict = Get-PcmVerdict $outcome
    try { if ($ownsTranscript -and $paths.TranscriptPath) { Add-Content -LiteralPath $paths.TranscriptPath -Value $verdict -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop } }
    catch { Add-PcmIssue $Context ('Transcript verdict: {0}' -f $_.Exception.Message) 'Audit'; Write-PcmReport $Context $reports }
    $outcome = Get-PcmOutcome $Context; $beforeVerdict = $Context.Issues.Count
    $verdict = Get-PcmVerdict $outcome
    try { Write-PcmLogFile -Context $Context -Level 'INFO' -Message $verdict } catch { $null = $_ }
    Write-PcmHost @(,@(('  ' + $verdict), $script:Theme.Dim))
    Write-Host ''
    if ($Context.Issues.Count -ne $beforeVerdict) {
        try { if ($ownsTranscript -and $paths.TranscriptPath) { Add-Content -LiteralPath $paths.TranscriptPath -Value 'RESULT: FAIL' -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop } }
        catch { Add-PcmIssue $Context ('Transcript verdict: {0}' -f $_.Exception.Message) 'Audit' }
        Write-PcmReport $Context $reports
    }
    $outcome = Get-PcmOutcome $Context
    $Context.ExitCode = $outcome.ExitCode; $Context.ResultStatus = $outcome.Status
    $result = [pscustomobject][ordered]@{
        RunId = $Context.RunId; Version = $Context.Version; Action = $Context.Action; Mode = $Context.Mode; Phase = $Context.Phase
        Status = $outcome.Status; ExitCode = $outcome.ExitCode; SuccessCount = $outcome.Counts.Success; ErrorCount = $outcome.Counts.Error
        PendingCount = $outcome.Counts.Pending; SkippedCount = $outcome.Counts.Skipped; PlannedCount = $outcome.Counts.Planned
        TotalCount = $outcome.Rows.Count; Duration = ((Get-Date) - [datetime]$Context.StartTime)
        BatchId = [string](Get-PcmValue $Context 'BatchId' ''); NextSteps = [string[]]@(@(Get-PcmValue $Context 'NextSteps' @()) | ForEach-Object { $_ } | Where-Object { $_ })
        Server = $Context.Server; BackupFiles = [string[]]$paths.BackupFiles; StateFiles = [string[]]$paths.StateFiles
        LogFile = $paths.LogFile; TranscriptPath = $paths.TranscriptPath
        CsvReport = $(if (-not $reports.CsvFailed) { $reports.Csv } else { '' })
        HtmlReport = $(if (-not $reports.HtmlFailed) { $reports.Html } else { '' }); GlobalReport = [string](Get-PcmValue $Context 'GlobalReportPath' ''); ConfigOverrides = [string[]]@(@(Get-PcmValue $Context 'ConfigOverrides' @()) | ForEach-Object { $_ } | Where-Object { $_ })
        Issues = [object[]]$Context.Issues.ToArray()
        WarningCount = @(@(Get-PcmValue $Context 'WarningList' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ }).Count
        Warnings = [object[]]@(@(Get-PcmValue $Context 'WarningList' @()) | ForEach-Object { $_ } | Where-Object { $null -ne $_ })
    }
    Write-PcmStepState -Context $Context
    Write-PcmEvent 'result' @{ status = $result.Status; exitCode = $result.ExitCode; action = [string]$result.Action; mode = [string]$result.Mode
        success = $result.SuccessCount; error = $result.ErrorCount; pending = $result.PendingCount; skipped = $result.SkippedCount; planned = $result.PlannedCount
        total = $result.TotalCount; seconds = [int]$result.Duration.TotalSeconds; batchId = $result.BatchId; nextSteps = @($result.NextSteps)
        logFile = [string]$result.LogFile; csvReport = [string]$result.CsvReport; htmlReport = [string]$result.HtmlReport; globalReport = [string]$result.GlobalReport; configOverrides = @($result.ConfigOverrides); warnings = $result.WarningCount
        issues = @($result.Issues | ForEach-Object { [string](Get-PcmValue $_ 'Message' ([string]$_)) }) }
    return $result
}
#endregion

#region 6. Global report of a batch ---------------------------------------------------------------

# One page per Convert batch, written at the end of every run that touches it (Convert, a resume of Convert, a Recover wave, Check -Batch): the batch end to end,
# every run, every object with its journey. More objects than this and the journey is written only for the objects that are not simply done (the CSV lists every object).
$script:GlobalDetailedObjects = 500
$script:GlobalEventsPerObject = 200

$script:PraGlobalCss = @'
:root{color-scheme:light dark;--bg:#f7f4ef;--surface:#fff;--soft:#f5f5f5;--border:#dedede;--text:#242424;--muted:#5c5c5c;--accent:#b11f4b;--accent-2:#9a1a41;--accent-fg:#fff;--ok:#16a34a;--warn:#b45309;--fail:#dc2626;--info:#0078d4}
@media (prefers-color-scheme:dark){:root{--bg:#3d3b3a;--surface:#292929;--soft:#2e2e2e;--border:#474747;--text:#dedede;--muted:#a0a0a0;--accent:#fd8ea1;--accent-2:#fb7b91;--accent-fg:#1a1a1a;--ok:#4ade80;--warn:#fbbf24;--fail:#f87171;--info:#4da6ff}}
*{box-sizing:border-box}body{margin:0;padding:28px 32px;background:var(--bg);color:var(--text);font:14px "Segoe UI",Aptos,Calibri,sans-serif}
header{border:1px solid var(--border);border-top:4px solid var(--accent);border-radius:16px;padding:24px;margin-bottom:20px;background:var(--surface)}
.origin{color:var(--accent);font-size:11px;letter-spacing:.1em;text-transform:uppercase;font-weight:700}h1{font-size:28px;font-weight:650;margin:8px 0 6px}h2{font-size:16px;font-weight:600;margin:0 0 12px}.subtitle{color:var(--muted);margin:0}
.tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:14px;margin:0 0 20px}.tile{padding:14px 16px;border:1px solid var(--border);border-radius:16px;background:var(--surface)}
.tile-label{font-size:12px;font-weight:600;color:var(--muted)}.tile-value{font-size:26px;font-weight:650;margin-top:6px}.tile.ok .tile-value{color:var(--ok)}.tile.warn .tile-value{color:var(--warn)}.tile.fail .tile-value{color:var(--fail)}
.card{background:var(--surface);border:1px solid var(--border);border-radius:16px;padding:20px;margin-bottom:20px;overflow-x:auto}.card.warn{border-left:4px solid var(--warn)}.card.fail{border-left:4px solid var(--fail)}
table{border-collapse:collapse;width:100%;font-size:13px}th{text-align:left;font-size:11px;letter-spacing:.05em;text-transform:uppercase;color:var(--accent);padding:6px 10px;border-bottom:1px solid var(--border)}td{padding:6px 10px;border-bottom:1px solid var(--border);vertical-align:top;word-break:break-word}
.pill{display:inline-block;padding:1px 10px;border-radius:999px;font-weight:600;font-size:12px;background:var(--soft)}.pill.ok{color:var(--ok)}.pill.warn{color:var(--warn)}.pill.fail{color:var(--fail)}.pill.muted{color:var(--muted)}
details{border-top:1px solid var(--border);padding:8px 0}summary{cursor:pointer;font-weight:600}.muted{color:var(--muted)}.mono{font-family:Consolas,"Cascadia Mono",monospace;font-size:12px}ul.plain{margin:6px 0 0;padding-left:18px}
@media (max-width:700px){body{padding:14px 12px}}
'@

function Get-PcmStatusClass {
    <# Style class of a status or verdict text (display only). #>
    param([string]$Status)
    foreach ($word in @('Converted', 'Rolled back', 'Done', 'Ok', 'Success', 'Met')) { if (Test-PcmSameText $Status $word) { return 'ok' } }
    foreach ($word in @('In progress', 'Rollback in progress', 'Running', 'Planned', 'Pending', 'Not yet')) { if (Test-PcmSameText $Status $word) { return 'warn' } }
    foreach ($word in @('Failed', 'Rollback failed', 'Fail', 'Error')) { if (Test-PcmSameText $Status $word) { return 'fail' } }
    return 'muted'
}

function Format-PcmUtcText {
    param([AllowNull()][object]$Value)
    $text = [string]$Value
    if (-not $text) { return '' }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse($text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) { return ($parsed.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss') + ' UTC') }
    return $text
}

function Get-PcmGlobalModel {
    <#
    .SYNOPSIS
        The batch end to end, from the journal: every batch (the Convert batch, then the Recover batches that rolled it back), every run, and for each object its
        stage in each batch, its journey (the events of the journal in time order) and a verdict.
    .PARAMETER History
        Output of Get-PcmBatchHistory.
    .PARAMETER FollowUp
        Rows of the follow-up of Check -Batch of this run (Identity, Expect, Met, Detail), shown as one more section.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$History, [AllowEmptyCollection()][object[]]$FollowUp = @(), [string]$Version = '', [string]$Environment = '')
    $batchById = @{}
    foreach ($batch in @($History.Batches)) { $batchById[[string]$batch.id] = $batch }
    $objects = [System.Collections.Generic.Dictionary[string, object]]::new([StringComparer]::OrdinalIgnoreCase)
    $get = {
        param([string]$Guid)
        if (-not $objects.ContainsKey($Guid)) { $objects[$Guid] = [pscustomobject]@{ ObjectGuid = $Guid; Identity = ''; Kind = ''; Stages = [System.Collections.Generic.List[object]]::new(); Events = [System.Collections.Generic.List[object]]::new(); Verdict = ''; Note = '' } }
        return $objects[$Guid]
    }
    foreach ($item in @($History.Items)) {
        $entry = & $get ([string]$item.object_guid)
        if (-not $entry.Identity -and [string]$item.identity) { $entry.Identity = [string]$item.identity }
        if (-not $entry.Kind -and [string]$item.kind) { $entry.Kind = [string]$item.kind }
        $batch = $batchById[[string]$item.batch_id]
        $entry.Stages.Add([pscustomobject]@{ BatchId = [string]$item.batch_id; Action = [string]$batch.action; Status = [string]$item.status; Step = [string]$item.step; Message = [string]$item.message; Created = [string]$batch.created_utc })
    }
    $batchEvents = [System.Collections.Generic.List[object]]::new()
    foreach ($journalEvent in @($History.Events)) {
        $batch = $batchById[[string]$journalEvent.batch_id]
        $row = [pscustomobject]@{ Time = [string]$journalEvent.time_utc; BatchId = [string]$journalEvent.batch_id; Action = [string]$batch.action; Step = [string]$journalEvent.step; Outcome = [string]$journalEvent.outcome; Detail = [string]$journalEvent.detail }
        if ([string]$journalEvent.object_guid) { $entry = & $get ([string]$journalEvent.object_guid); $entry.Events.Add($row) } else { $batchEvents.Add($row) }
    }
    $counts = [ordered]@{ Converted = 0; RolledBack = 0; InProgress = 0; Failed = 0; Skipped = 0 }
    foreach ($entry in $objects.Values) {
        $recover = @($entry.Stages | Where-Object { Test-PcmSameText $_.Action 'Recover' })
        $convert = @($entry.Stages | Where-Object { Test-PcmSameText $_.Action 'Convert' })
        $lastRecover = $recover | Select-Object -Last 1
        $lastConvert = $convert | Select-Object -Last 1
        $verdict = 'In progress'
        if ($recover | Where-Object { Test-PcmSameText $_.Status 'Done' }) { $verdict = 'Rolled back' }
        elseif ($lastRecover -and (Test-PcmSameText $lastRecover.Status 'Failed')) { $verdict = 'Rollback failed' }
        elseif ($lastRecover) { $verdict = 'Rollback in progress' }
        elseif ($lastConvert -and (Test-PcmSameText $lastConvert.Status 'Done')) { $verdict = 'Converted' }
        elseif ($lastConvert -and (Test-PcmSameText $lastConvert.Status 'Failed')) { $verdict = 'Failed' }
        elseif ($lastConvert -and (Test-PcmSameText $lastConvert.Status 'Skipped')) { $verdict = 'Skipped' }
        $entry.Verdict = $verdict
        $failure = @($entry.Stages | Where-Object { [string]$_.Message } | Select-Object -Last 1)
        if ($failure.Count -and @('Failed', 'Rollback failed') -contains $verdict) { $entry.Note = [string]$failure[0].Message }
        switch ($verdict) { 'Converted' { $counts.Converted++ } 'Rolled back' { $counts.RolledBack++ } 'Failed' { $counts.Failed++ } 'Rollback failed' { $counts.Failed++ } 'Skipped' { $counts.Skipped++ } default { $counts.InProgress++ } }
    }
    $ordered = @($objects.Values | Sort-Object { [string]$_.Identity }, { [string]$_.ObjectGuid })
    return [pscustomobject]@{ Version = $Version; Environment = $Environment; Batches = @($History.Batches); Runs = @($History.Runs); BatchEvents = $batchEvents.ToArray()
        Objects = $ordered; Counts = [pscustomobject]$counts; FollowUp = @($FollowUp) }
}

function Write-PcmGlobalReportHtml {
    <# The global report: one page, no script, no external resource. Every value is HTML-encoded. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][IO.TextWriter]$Writer)
    $h = { param($Value) ConvertTo-PcmHtmlText $Value }
    $dot = [char]0x00B7
    $counts = $Model.Counts
    $total = @($Model.Objects).Count
    $convertBatch = @($Model.Batches | Where-Object { Test-PcmSameText $_.action 'Convert' } | Select-Object -First 1)
    $batchId = if ($convertBatch.Count) { [string]$convertBatch[0].id } else { '' }
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('<!doctype html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>').Append((& $h ('PRA Cloud Mailbox - batch ' + $batchId + ' - global report'))).Append('</title><style>').Append($script:PraGlobalCss).Append('</style></head><body>')
    [void]$sb.Append('<header><div class="origin">PRA Cloud Mailbox ').Append($dot).Append(' global report</div><h1>').Append((& $h ('Batch {0} from end to end' -f $batchId))).Append('</h1><p class="subtitle">')
    [void]$sb.Append((& $h ('Environment {0} {1} tool v{2} {1} one page for the whole life of the batch: Convert, its resumes, Recover and the follow-up checks' -f $Model.Environment, $dot, $Model.Version))).Append('</p></header>')
    [void]$sb.Append('<div class="tiles">')
    foreach ($tile in @(@('Objects', $total, ''), @('Converted', $counts.Converted, 'ok'), @('Rolled back', $counts.RolledBack, 'ok'), @('In progress', $counts.InProgress, $(if ($counts.InProgress) { 'warn' } else { '' })),
            @('Failed', $counts.Failed, $(if ($counts.Failed) { 'fail' } else { '' })), @('Skipped', $counts.Skipped, ''))) {
        [void]$sb.Append('<div class="tile ').Append($tile[2]).Append('"><div class="tile-label">').Append((& $h $tile[0])).Append('</div><div class="tile-value">').Append((& $h $tile[1])).Append('</div></div>')
    }
    [void]$sb.Append('</div>')
    # Batches.
    [void]$sb.Append('<section class="card"><h2>Batches</h2><table><thead><tr><th>Batch</th><th>Action</th><th>Status</th><th>Created</th><th>Updated</th><th>Snapshot</th><th>Operator</th><th>Computer</th><th>Tool</th></tr></thead><tbody>')
    foreach ($batch in @($Model.Batches)) {
        $status = [string]$batch.status
        $pill = switch ($status) { 'Complete' { 'ok' } 'Failed' { 'fail' } default { 'warn' } }
        [void]$sb.Append('<tr><td class="mono">').Append((& $h $batch.id)).Append('</td><td>').Append((& $h $batch.action)).Append('</td><td><span class="pill ').Append($pill).Append('">').Append((& $h $status)).Append('</span></td><td>')
        [void]$sb.Append((& $h (Format-PcmUtcText $batch.created_utc))).Append('</td><td>').Append((& $h (Format-PcmUtcText $batch.updated_utc))).Append('</td><td>').Append((& $h $batch.snapshot_id)).Append('</td><td>').Append((& $h $batch.account)).Append('</td><td>')
        [void]$sb.Append((& $h $batch.computer)).Append('</td><td>').Append((& $h $batch.tool_version)).Append('</td></tr>')
    }
    [void]$sb.Append('</tbody></table></section>')
    # Runs.
    [void]$sb.Append('<section class="card"><h2>Runs</h2>')
    if (-not @($Model.Runs).Count) { [void]$sb.Append('<p class="muted">No run was recorded: the batch was written by an earlier version of the tool.</p>') }
    else {
        [void]$sb.Append('<table><thead><tr><th>Run</th><th>Batch</th><th>Action</th><th>Started</th><th>Ended</th><th>Result</th><th>Done</th><th>Pending</th><th>Errors</th><th>Objects chosen by</th></tr></thead><tbody>')
        $number = 0
        foreach ($run in @($Model.Runs)) {
            $number++
            $pill = switch ([int]$run.exit_code) { 0 { 'ok' } 2 { 'warn' } default { 'fail' } }
            [void]$sb.Append('<tr><td>').Append($number).Append('</td><td class="mono">').Append((& $h $run.batch_id)).Append('</td><td>').Append((& $h $run.action)).Append('</td><td>').Append((& $h (Format-PcmUtcText $run.started_utc))).Append('</td><td>')
            [void]$sb.Append((& $h (Format-PcmUtcText $run.ended_utc))).Append('</td><td><span class="pill ').Append($pill).Append('">').Append((& $h $run.status)).Append('</span></td><td>').Append((& $h $run.done_count)).Append('</td><td>').Append((& $h $run.pending_count)).Append('</td><td>')
            [void]$sb.Append((& $h $run.error_count)).Append('</td><td>').Append((& $h $run.source_text)).Append('</td></tr>')
        }
        [void]$sb.Append('</tbody></table>')
    }
    [void]$sb.Append('</section>')
    # Settings that differed from the configuration file.
    [void]$sb.Append('<section class="card"><h2>Run settings changed from the configuration file</h2>')
    $any = $false
    foreach ($run in @($Model.Runs)) {
        $items = @(([string]$run.overrides_text) -split ' \| ' | Where-Object { $_ })
        if (-not $items.Count) { continue }
        $any = $true
        [void]$sb.Append('<p><strong>').Append((& $h ('{0} {1} {2}' -f $run.action, $dot, (Format-PcmUtcText $run.started_utc)))).Append('</strong></p><ul class="plain">')
        foreach ($item in $items) { [void]$sb.Append('<li class="mono">').Append((& $h $item)).Append('</li>') }
        [void]$sb.Append('</ul>')
    }
    if (-not $any) { [void]$sb.Append('<p class="muted">None: every run used the settings of the configuration file.</p>') }
    [void]$sb.Append('</section>')
    # Follow-up.
    if (@($Model.FollowUp).Count) {
        [void]$sb.Append('<section class="card"><h2>Follow-up check of this run</h2><table><thead><tr><th>Object</th><th>Expected</th><th>Result</th><th>Detail</th></tr></thead><tbody>')
        foreach ($row in @($Model.FollowUp)) {
            [void]$sb.Append('<tr><td>').Append((& $h $row.Identity)).Append('</td><td>').Append((& $h $row.Expect)).Append('</td><td><span class="pill ').Append($(if ($row.Met) { 'ok' } else { 'warn' })).Append('">').Append($(if ($row.Met) { 'Met' } else { 'Not yet' })).Append('</span></td><td>').Append((& $h $row.Detail)).Append('</td></tr>')
        }
        [void]$sb.Append('</tbody></table></section>')
    }
    # Objects.
    [void]$sb.Append('<section class="card"><h2>Objects, end to end</h2>')
    $detailed = $total -le $script:GlobalDetailedObjects
    if (-not $detailed) { [void]$sb.Append('<p class="muted">').Append((& $h ('More than {0} objects: the journey is written for the objects that are not simply converted or rolled back; the CSV file lists every object.' -f $script:GlobalDetailedObjects))).Append('</p>') }
    [void]$sb.Append('<table><thead><tr><th>Object</th><th>Type</th><th>Verdict</th><th>Stages</th></tr></thead><tbody>')
    foreach ($entry in @($Model.Objects)) {
        $stages = (@($entry.Stages | ForEach-Object { '{0} {1}: {2}{3}' -f $_.Action, $_.BatchId, $_.Status, $(if ($_.Step) { ' (' + $_.Step + ')' } else { '' }) }) -join '; ')
        [void]$sb.Append('<tr><td>').Append((& $h $(if ($entry.Identity) { $entry.Identity } else { $entry.ObjectGuid }))).Append('</td><td>').Append((& $h $entry.Kind)).Append('</td><td><span class="pill ').Append((Get-PcmStatusClass $entry.Verdict)).Append('">').Append((& $h $entry.Verdict)).Append('</span></td><td>').Append((& $h $stages)).Append('</td></tr>')
        $simple = @('Converted', 'Rolled back') -contains $entry.Verdict
        if (($detailed -or -not $simple) -and ($entry.Events.Count -or $entry.Note)) {
            [void]$sb.Append('<tr><td colspan="4"><details><summary>Journey of ').Append((& $h $(if ($entry.Identity) { $entry.Identity } else { $entry.ObjectGuid }))).Append('</summary>')
            if ($entry.Note) { [void]$sb.Append('<p class="muted">').Append((& $h $entry.Note)).Append('</p>') }
            [void]$sb.Append('<table><thead><tr><th>Time</th><th>Batch</th><th>Step</th><th>Outcome</th><th>Detail</th></tr></thead><tbody>')
            $shown = @($entry.Events | Select-Object -First $script:GlobalEventsPerObject)
            foreach ($step in $shown) {
                [void]$sb.Append('<tr><td>').Append((& $h (Format-PcmUtcText $step.Time))).Append('</td><td>').Append((& $h ('{0} {1}' -f $step.Action, $step.BatchId))).Append('</td><td>').Append((& $h $step.Step)).Append('</td><td><span class="pill ').Append((Get-PcmStatusClass $step.Outcome)).Append('">').Append((& $h $step.Outcome)).Append('</span></td><td>').Append((& $h $step.Detail)).Append('</td></tr>')
            }
            if ($entry.Events.Count -gt $shown.Count) { [void]$sb.Append('<tr><td colspan="5" class="muted">').Append((& $h ('... and {0} more event(s): the journal holds them all.' -f ($entry.Events.Count - $shown.Count)))).Append('</td></tr>') }
            [void]$sb.Append('</tbody></table></details></td></tr>')
        }
    }
    if (-not $total) { [void]$sb.Append('<tr><td colspan="4" class="muted">No object.</td></tr>') }
    [void]$sb.Append('</tbody></table></section>')
    [void]$sb.Append('<p class="muted">Generated by PRA Cloud Mailbox v').Append((& $h $Model.Version)).Append('. This report holds names, statuses and steps only: no secret, no token.</p></body></html>')
    $Writer.Write($sb.ToString())
}

function Write-PcmGlobalReportCsv {
    <# One line per object: verdict, the stages and the journey as text (a text that starts like a formula gets an apostrophe). Semicolon, UTF-8 with BOM. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$Path)
    $guard = { param($text) $value = [string]$text; if ($value -match '^[=+\-@\t\r]') { "'" + $value } else { $value } }
    $rows = foreach ($entry in @($Model.Objects)) {
        [pscustomobject][ordered]@{
            ObjectGuid = (& $guard $entry.ObjectGuid); Identity = (& $guard $entry.Identity); Kind = (& $guard $entry.Kind); Verdict = $entry.Verdict
            Stages = (& $guard (@($entry.Stages | ForEach-Object { '{0} {1}: {2}{3}' -f $_.Action, $_.BatchId, $_.Status, $(if ($_.Step) { ' (' + $_.Step + ')' } else { '' }) }) -join ' | '))
            Note = (& $guard $entry.Note)
            Journey = (& $guard (@($entry.Events | Select-Object -First $script:GlobalEventsPerObject | ForEach-Object { '{0} {1}/{2} {3}: {4}' -f (Format-PcmUtcText $_.Time), $_.Action, $_.BatchId, $_.Step, $_.Detail }) -join ' | '))
        }
    }
    $fields = @('ObjectGuid', 'Identity', 'Kind', 'Verdict', 'Stages', 'Note', 'Journey')
    if (@($rows).Count) { $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Delimiter ';' -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop }
    else { Set-Content -LiteralPath $Path -Value ('"' + ($fields -join '";"') + '"') -Encoding UTF8 -WhatIf:$false -Confirm:$false -ErrorAction Stop }
}

function Write-PcmGlobalReport {
    <#
    .SYNOPSIS
        Writes the global HTML and CSV report of a batch into the report folder (the same two files are replaced at each run). Returns Html and Csv (empty when a file could
        not be written: a warning says why; the run itself is never stopped by it).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$Context, [Parameter(Mandatory)][object]$History, [string]$ConvertBatchId, [AllowEmptyCollection()][object[]]$FollowUp = @())
    $result = @{ Html = ''; Csv = '' }
    try {
        $model = Get-PcmGlobalModel -History $History -FollowUp $FollowUp -Version ([string]$Context.Version) -Environment ([string](Get-PcmValue $Context.Config 'Environment' ''))
        $null = New-Item -ItemType Directory -Path $Context.ReportFolder -Force -WhatIf:$false -Confirm:$false -ErrorAction Stop
        $stem = 'PraCloudMailbox_{0}_batch{1}_global' -f [string](Get-PcmValue $Context.Config 'Environment' 'PROD'), $ConvertBatchId
        $html = Join-Path $Context.ReportFolder ($stem + '.html'); $csv = Join-Path $Context.ReportFolder ($stem + '.csv')
        $stream = [IO.File]::Open($html, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::None)
        $writer = New-Object IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($true)), 65536)
        try { Write-PcmGlobalReportHtml -Model $model -Writer $writer } finally { $writer.Dispose() }
        $result.Html = $html
        Write-PcmGlobalReportCsv -Model $model -Path $csv
        $result.Csv = $csv
        $Context['GlobalModel'] = $model
    }
    catch { Write-PcmLog -Context $Context -Message ('The global report could not be written: {0}' -f $_.Exception.Message) -Level Warning }
    return $result
}

function Get-PcmConfigurationFlat {
    <# Dotted key -> text, for every leaf of a configuration (lists are joined with ', '); a key that starts with _ is not a setting. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][Collections.IDictionary]$Configuration, [string]$Prefix = '')
    foreach ($key in @($Configuration.Keys | ForEach-Object { [string]$_ } | Sort-Object)) {
        if ($key.StartsWith('_', [StringComparison]::OrdinalIgnoreCase)) { continue }
        $value = $Configuration[$key]
        if ($value -is [Collections.IDictionary]) { Get-PcmConfigurationFlat -Configuration $value -Prefix ($Prefix + $key + '.') }
        elseif ($value -is [Collections.IEnumerable] -and $value -isnot [string]) { [pscustomobject]@{ Key = $Prefix + $key; Value = (@($value | ForEach-Object { [string]$_ }) -join ', ') } }
        else { [pscustomobject]@{ Key = $Prefix + $key; Value = [string]$value } }
    }
}

function Get-PcmConfigurationDifference {
    <# What a run changes compared with the configuration file: one object (Key, Base, Effective, Text = 'Key: base -> effective') per setting whose value differs (the letter case is not a difference). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][Collections.IDictionary]$Base, [Parameter(Mandatory)][Collections.IDictionary]$Effective)
    $baseFlat = @{}
    foreach ($row in @(Get-PcmConfigurationFlat -Configuration $Base)) { $baseFlat[$row.Key] = $row.Value }
    foreach ($row in @(Get-PcmConfigurationFlat -Configuration $Effective)) {
        $before = if ($baseFlat.ContainsKey($row.Key)) { [string]$baseFlat[$row.Key] } else { '' }
        if (-not (Test-PcmSameText $before $row.Value)) {
            $beforeText = if ($before) { $before } else { '(empty)' }; $afterText = if ($row.Value) { $row.Value } else { '(empty)' }
            [pscustomobject][ordered]@{ Key = $row.Key; Base = $before; Effective = $row.Value; Text = ('{0}: {1} -> {2}' -f $row.Key, $beforeText, $afterText) }
        }
    }
}
#endregion

Export-ModuleMember -Function Get-PcmValue, Format-PcmDuration, Add-PcmIssue, Write-PcmLog, Write-PcmItem, Write-PcmProgress, Get-PcmDurationEstimate, Get-PcmRecoverWaveSize,
    Write-PcmBanner, Write-PcmStep, Write-PcmPlan, Write-PcmStepState, Write-PcmStepCount, Write-PcmSummary, Write-PcmHost, Resolve-PcmPath, Import-PcmConfiguration, Initialize-PcmAudit, Get-PcmOutcome, Complete-PcmRun,
    Write-PcmEvent, Test-PcmStopRequest, Request-PcmOperator, Get-PcmToolVersion, Test-PcmSameText, Get-PcmDistinctText, Resolve-PcmOfficialValue, Get-PcmCloudModuleReport,
    Get-PcmReportHtml, Write-PcmReportHtml, Read-PcmReportData, ConvertTo-PcmHtmlText, Get-PcmGlobalModel, Write-PcmGlobalReport, Write-PcmGlobalReportHtml, Write-PcmGlobalReportCsv, Format-PcmUtcText,
    Get-PcmConfigurationDifference, Get-PcmConfigurationFlat, ConvertTo-PcmScheduleTime, ConvertTo-PcmPlainMessage,
    Test-PcmSystemMailboxName, ConvertTo-PcmDnKey, ConvertFrom-PcmDnToCanonical, Test-PcmInSearchBase, ConvertTo-PcmGuidKey, New-PcmRecordIndex, Find-PcmRecord

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVudO+QVXokAV2
# BWdUEkN5VmmfRbPT8xo2g/xP7MHa+qCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBWo1xO/vPj
# 1m24JxFCkWpW47+lEIqPJ3SPkwDMDv0GEDANBgkqhkiG9w0BAQEFAASCAgBgmmYL
# 2bZOqKtVMoavVvRMgdT5PnifzhFtLFIHxgT0gVnuzKPc87Dei1lm09U2zV4xkQF0
# rXa8RNg8kCts5O0oP/UpP9EKMd0CDq5kGFR4LVd4w9cZnH/rLBJz+/JUE6JBsR6O
# yB7Po35lWd7lY0Lnf4DGqncstTqMVqbx9fVkRnL/1XoZDEq0KNr6vhUQl61ddnj4
# kIiAt2m+2VtsoQCqo5lBdoTUH1A+NbRmT7GtnRBMOdEivVYBQvBTHM8UeSJhAsnl
# c5nDyzwe7f4AFqN6N2WEraUtlUcxdSYxRqGVLw0j486OPwolkv9SgrCDviw9LtAM
# H3kF5p9a+D8OAgk0kwNPZV4gX4J2Gzo5riPCguuJqA44ZzoV7OYNRo16vMKm3ZPk
# V5Zyre+sEMTCWDQVXJtwjV18D7MxUQMJnD39T5DbNfE4QMEK6x7qU167jojhq+3B
# FJt8d22YYzaB4w+7eSbrKa9/fDD9leNNHzUFsEbejT8EW721lNO81Fbx2P8qqKiQ
# bLe5y7HA5jJauyrc+k4cAjRctvdoSDQP+8aQ3K28QUASKCpAcFhXmNWJcheoLymt
# Ktp7oozP9xbXVbGC/+YxCqePK+XNqbr7S2nYQfrV3DC8+tHJj7j6MyNWYM0qxzcm
# oaIw+rATZ8Pc1sAwXtpLfnz3a1DfOsGXSnsriqGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEILsqcWEHpZF4XHGx9OarYt5vC82xh/DAH+2zgKibU6DzAgh23FXquilG
# uRgPMjAyNjExMTQwODE2MzhaMAMCAQECEC82YpK5W985lsGu9gQUtyegghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNDA4MTYzOFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEILEW
# 4Rlle3Y2V+offF+5KRQ12c5tqZWjNFZjSG27zIivMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiBH5QOFQStLUvnXUrxzrLo0r8AS+jQwV3gXii6Ct8EK0gIhAL2X
# H0I3KEKn/ICqlaLLsIRTboXH8DPwEG9fX9823lrF
# SIG # End signature block
