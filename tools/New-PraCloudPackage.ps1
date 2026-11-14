#Requires -Version 7.5
<#
.SYNOPSIS
    Copies the files needed to run PRA Cloud Mailbox into a separate folder, ready to be zipped, and optionally writes the release ZIP and its .sha256 file.

.DESCRIPTION
    The runtime lives in the package\ folder of the repository; the package keeps the same paths as the folder that is extracted from the
    release ZIP (Invoke-PraCloudMailbox.ps1 at its root, module\, config\, templates\, lib\, docs\). The distribution contains the files of the
    signed inventory of package\PackageIntegrity.psd1, that file itself, and CHANGELOG.md (copied from the repository root):
        Invoke-PraCloudMailbox.ps1, PraCloudMailbox.psd1, PraCloudMailbox.psm1, PackageIntegrity.psd1, README.md, LICENSE, THIRD-PARTY-NOTICES.md,
        CHANGELOG.md, module\ (modules, console, verifier and XAML), templates\, config\ (configuration and samples), lib\sqlite\,
        docs\ (user guide and administrator guide, Markdown and self-contained HTML, with images)
    It never copies data\, logs\, reports\, tests\ or tools\: the snapshots and the journal hold names, GUIDs and permissions.

    The package source and the finished package must pass the owner-signed integrity check (module\Test-PackageIntegrity.ps1 and
    PackageIntegrity.psd1: exact SHA256 of every stable file). A source that is unsigned or changed is refused before anything is written: there is
    no unsigned mode. The guides (HTML) are not rebuilt here: they are part of the signed inventory (tools\Build-Documentation.ps1 runs before signing).

    The configuration file is copied with the environment values emptied (Exchange server and domain controller, scope, tenant, app registration,
    licences, case hold policy, Entra Connect): the administrator fills them in (administrator guide, chapter 6). The script then checks that none of
    these values appears in the package. The configuration is data outside the integrity manifest.

.PARAMETER Destination
    Package folder. Default: package\PraCloudMailbox-<version>, next to the repository folder.

.PARAMETER Zip
    Also writes <Destination>.zip, with the package folder at its root (the zip of a release), and <Destination>.zip.sha256 (the line
    '<sha256>  <file name>', the same format as the transfer packages of the PRA tools).

.PARAMETER Force
    Replace the destination folder (and zip) if they already exist. A folder that contains data\, logs\ or reports\ (a package that has been run)
    is never replaced.

.EXAMPLE
    .\tools\New-PraCloudPackage.ps1 -Zip
    Creates ..\package\PraCloudMailbox-2.1.0, ..\package\PraCloudMailbox-2.1.0.zip and ..\package\PraCloudMailbox-2.1.0.zip.sha256.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Zip,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

function Reset-PraPackageConsent {
    <# Data-loss/routing consents are operator assertions: the packaged configuration never carries them as true. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    foreach ($key in 'OnPremisesReady', 'AllowUnprotectedDeletion') {
        # Line start, '{' or ';' only (never a comment); any value other than the literal $false is reset, also inline.
        $pattern = "((?:^|[;{])[ \t]*)($key)([ \t]*=[ \t]*)(?!\`$false(?![\w]))[^\s;}#]+"
        $Text = [regex]::Replace($Text, $pattern, '$1$2$3$$false', [Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [Text.RegularExpressions.RegexOptions]::Multiline)
    }
    return $Text
}

function Assert-AuthenticatedPackage([string]$Path, [switch]$Distribution) {
    $verifier = Join-Path $Path 'module\Test-PackageIntegrity.ps1'
    $cursor = [IO.Path]::GetFullPath($verifier)
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
    if ($Distribution) { & $verifier -Root $Path -Distribution } else { & $verifier -Root $Path }
}

# Ordinal comparisons throughout: on .NET with ICU the culture comparison (-eq, -in, Select-Object -Unique) ignores zero-weight characters.
function Test-OrdinalIn([string]$Value, [string[]]$Set, [switch]$IgnoreCase) {
    $comparison = if ($IgnoreCase) { [StringComparison]::OrdinalIgnoreCase } else { [StringComparison]::Ordinal }
    foreach ($item in $Set) { if ([string]::Equals($Value, $item, $comparison)) { return $true } }
    return $false
}
function New-OrdinalSet([string[]]$Items) { $set = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::Ordinal); foreach ($item in $Items) { [void]$set.Add($item) }; , $set }

$root = Split-Path $PSScriptRoot -Parent
$packageRoot = Join-Path $root 'package'
# The version has one source: the ModuleVersion of the module manifest (read as data, nothing is executed).
$version = [string](Import-PowerShellDataFile -LiteralPath (Join-Path $packageRoot 'PraCloudMailbox.psd1')).ModuleVersion
if ($version -notmatch '^\d+\.\d+\.\d+$') { throw 'Version not found in package\PraCloudMailbox.psd1 (ModuleVersion).' }

# The source is authenticated first: an unsigned or changed source never produces a package.
Assert-AuthenticatedPackage $packageRoot

if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "package\PraCloudMailbox-$version" }
if (-not [IO.Path]::IsPathRooted($Destination)) { $Destination = Join-Path (Get-Location).Path $Destination }
$Destination = [IO.Path]::GetFullPath($Destination).TrimEnd('\')
$zipPath = "$Destination.zip"
$hashPath = "$zipPath.sha256"

$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the tool folder: $Destination"
}
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-PraCloudMailbox.ps1'))) { throw "The destination is not a PRA Cloud Mailbox package, it is not replaced: $Destination" }
    foreach ($used in 'data', 'logs', 'reports') {
        if (Test-Path -LiteralPath (Join-Path $Destination $used)) { throw "The destination contains a $used folder (a package that has been run), it is not replaced: $Destination" }
    }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}
if ($Zip) {
    foreach ($existing in $zipPath, $hashPath) {
        if (-not (Test-Path -LiteralPath $existing)) { continue }
        if (-not $Force) { throw "The file already exists: $existing. Use -Force to replace it." }
        Remove-Item -LiteralPath $existing -Force
    }
}

# ---- Files needed at run time (paths relative to package\, the same as in the extracted ZIP) ----------------
$manifest = Import-PowerShellDataFile -LiteralPath (Join-Path $packageRoot 'PackageIntegrity.psd1')
$files = New-Object 'Collections.Generic.List[string]'
foreach ($entry in $manifest.Files) { $files.Add([string]$entry.Path) }
$files.Add('PackageIntegrity.psd1')
if ((New-OrdinalSet $files.ToArray()).Count -ne $files.Count) { throw 'Package integrity: duplicate distribution file path.' }

foreach ($f in $files) {
    # CHANGELOG.md is the only file copied from the repository root; it is hashed by the manifest.
    $source = if ([string]::Equals($f, 'CHANGELOG.md', [StringComparison]::Ordinal)) { Join-Path $root $f } else { Join-Path $packageRoot $f }
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Configuration with the environment values emptied ---------------------------------------------------
$configRelative = 'config\PraCloudMailbox.config.psd1'
$config = [IO.File]::ReadAllText((Join-Path $packageRoot $configRelative), [Text.Encoding]::UTF8)
$emptied = New-Object 'Collections.Generic.List[string]'
$keys = 'DomainController', 'Server', 'SearchBase', 'GroupDN', 'CsvPath', 'ContactsSearchBase', 'DistributionGroupsSearchBase', 'TenantId', 'Organization',
    'AppId', 'CertificateThumbprint', 'UserPrincipalName', 'DefaultUsageLocation', 'GroupId', 'SkuPartNumber', 'HoldPolicy', 'ScriptPath',
    'LicensingGroupId', 'OnPremisesRoutingDomain'
foreach ($key in $keys) {
    # Every occurrence, also inside an inline hashtable (Users = @{ Mode = 'Group'; GroupId = '...' }).
    $pattern = "(?<![\w.])($key\s*=\s*)'([^']*)'"
    $found = [regex]::Matches($config, $pattern)
    if (-not $found.Count) { throw "The key $key is not in $configRelative." }
    foreach ($match in $found) { if ($match.Groups[2].Value) { $emptied.Add($match.Groups[2].Value) } }
    $config = [regex]::Replace($config, $pattern, '$1''''')
}
$config = [regex]::Replace($config, "(?m)^(\s*Environment\s*=\s*)'[^']*'", '$1''PROD''')
# Entra Connect without its server or script: the operator mode, valid until the administrator chooses.
$config = [regex]::Replace($config, "(?m)^(\s*Mode\s*=\s*)'(Script|Remoting)'", '$1''Manual''')
$config = Reset-PraPackageConsent -Text $config
$configTarget = Join-Path $Destination $configRelative
[IO.File]::WriteAllText($configTarget, $config, [Text.UTF8Encoding]::new($true))

# ---- Checks -------------------------------------------------------------------------------------------------
$problems = New-Object 'Collections.Generic.List[string]'
foreach ($name in 'data', 'logs', 'reports', 'tests', 'tools') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { Test-OrdinalIn $_.Extension @('.db', '.bak', '.sha256', '.clixml', '.jsonl', '.log') -IgnoreCase } | ForEach-Object { $problems.Add("File not allowed in the package: $($_.Name)") }
$textFiles = Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { Test-OrdinalIn $_.Extension @('.ps1', '.psm1', '.psd1', '.csv', '.md', '.html', '.txt', '.xaml') -IgnoreCase }
foreach ($value in (New-OrdinalSet @($emptied | Where-Object { $_.Length -ge 4 }))) {
    foreach ($file in $textFiles) {
        if ([IO.File]::ReadAllText($file.FullName).IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $problems.Add("An environment value of the configuration ($value) appears in $($file.FullName.Substring($Destination.Length + 1)).")
        }
    }
}
$check = Import-PowerShellDataFile -LiteralPath $configTarget
if ($check.Cloud.TenantId -or $check.Cloud.AppId -or $check.Cloud.CertificateThumbprint -or $check.Exchange.Server -or $check.Licensing.Users.GroupId -or $check.Retention.HoldPolicy) { $problems.Add('Tenant or Exchange values are still in the packaged configuration.') }
if ($check.Resources.OnPremisesReady -ne $false -or $check.Resources.AllowUnprotectedDeletion -ne $false) { $problems.Add('Resources.OnPremisesReady and Resources.AllowUnprotectedDeletion must both be $false in the packaged configuration.') }
$allowed = @($files) + @($configRelative)
$delivered = @(Get-ChildItem -LiteralPath $Destination -Recurse -File | ForEach-Object { $_.FullName.Substring($Destination.Length + 1) })
foreach ($extra in @($delivered | Where-Object { -not (Test-OrdinalIn $_ $allowed) })) { $problems.Add("Unexpected file in the package: $extra") }
foreach ($missing in @($allowed | Where-Object { -not (Test-OrdinalIn $_ $delivered) })) { $problems.Add("Expected file missing from the package: $missing") }
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }
try { Assert-AuthenticatedPackage $Destination -Distribution }
catch { Remove-Item -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue; throw }

if ($Zip) {
    Compress-Archive -LiteralPath $Destination -DestinationPath $zipPath -CompressionLevel Optimal
    $hash = (Get-FileHash -LiteralPath $zipPath -Algorithm SHA256).Hash.ToLowerInvariant()
    [IO.File]::WriteAllText($hashPath, ('{0}  {1}' -f $hash, [IO.Path]::GetFileName($zipPath)) + "`n", [Text.UTF8Encoding]::new($false))
}

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  PRA Cloud Mailbox $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
if ($Zip) {
    Write-Host ("  Zip      : {0} ({1:N1} MB)" -f $zipPath, ((Get-Item -LiteralPath $zipPath).Length / 1MB))
    Write-Host "  SHA-256  : $hashPath"
}
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host '  Integrity: owner signatures and SHA256 inventory verified'
Write-Host "  Config   : environment values emptied ($((New-OrdinalSet $emptied.ToArray()).Count)) - fill in the configuration (administrator guide, chapter 6)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,12:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
