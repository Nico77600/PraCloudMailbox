#Requires -Version 7.4
<#
.SYNOPSIS
    Copies the files needed to run PRA Cloud Mailbox into a separate folder, ready to be zipped.

.DESCRIPTION
    The package contains only what Invoke-PraCloudMailbox.ps1 needs at run time, plus the HTML guides:
        Invoke-PraCloudMailbox.ps1, module\*.psm1, lib\sqlite\, templates\, config\ (configuration and samples),
        docs\PraCloudMailbox-UserGuide.html, docs\PraCloudMailbox-Guide.html,
        README.md, CHANGELOG.md, LICENSE, THIRD-PARTY-NOTICES.md
    The HTML guides are rebuilt first from their Markdown sources (tools\Build-Documentation.ps1): they are
    self-contained (images inline), so the Markdown sources and the images are not copied.
    It never copies data\ (snapshots and journal: names, GUIDs, permissions), logs\, reports\, tests\ or tools\,
    nor a saved credential (config\*.cred.xml).

    The configuration is copied with the environment values emptied (Exchange server and domain controller,
    scope, tenant, app registration, licences, case hold policy, Entra Connect): the administrator fills them
    in (developer guide, chapter 6). The script then checks that none of these values appears in the package.

.PARAMETER Destination
    Package folder. Default: package\PraCloudMailbox-<version>, next to the tool folder.

.PARAMETER Zip
    Also writes <Destination>.zip, with the package folder at its root (the zip of a release).

.PARAMETER Force
    Replace the destination folder (and zip) if they already exist. A folder that contains data\, logs\ or
    reports\ (a package that has been run) is never replaced.

.EXAMPLE
    .\tools\New-PraCloudPackage.ps1 -Zip
    Creates ..\package\PraCloudMailbox-1.0.0 and ..\package\PraCloudMailbox-1.0.0.zip.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0  (from PRA Remote Mailbox 2.0.1 and Web Services Client for Exchange 1.0.0)
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Zip,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$entry = Join-Path $root 'Invoke-PraCloudMailbox.ps1'
$version = ([regex]::Match([IO.File]::ReadAllText($entry), "(?m)^\`$toolVersion = '([0-9.]+)'")).Groups[1].Value
if (-not $version) { throw 'Version not found in Invoke-PraCloudMailbox.ps1.' }
if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "package\PraCloudMailbox-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$zipPath = "$Destination.zip"

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
if ($Zip -and (Test-Path -LiteralPath $zipPath)) {
    if (-not $Force) { throw "The zip already exists: $zipPath. Use -Force to replace it." }
    Remove-Item -LiteralPath $zipPath -Force
}

# ---- HTML guides, rebuilt from the Markdown sources ----------------------------------------------------
& (Join-Path $PSScriptRoot 'Build-Documentation.ps1') | Out-Null

# ---- Files needed at run time ----------------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-PraCloudMailbox.ps1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', 'templates\Report.template.html',
    'config\Targets.sample.csv', 'config\EntraConnect.sample.ps1', 'docs\PraCloudMailbox-UserGuide.html', 'docs\PraCloudMailbox-Guide.html') { $files.Add($f) }
foreach ($m in 'PRA2.Common.psm1', 'PRA2.Store.psm1', 'PRA2.Collect.psm1', 'PRA2.Cloud.psm1') { $files.Add("module\$m") }
Get-ChildItem -LiteralPath (Join-Path $root 'lib\sqlite') -Recurse -File | ForEach-Object { $files.Add($_.FullName.Substring($rootPrefix.Length)) }

foreach ($f in $files) {
    $source = Join-Path $root $f
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Configuration with the environment values emptied -----------------------------------------------------
$configRelative = 'config\PraCloudMailbox.config.psd1'
$config = [IO.File]::ReadAllText((Join-Path $root $configRelative), [Text.Encoding]::UTF8)
$emptied = [Collections.Generic.List[string]]::new()
$keys = 'DomainController', 'Server', 'SearchBase', 'GroupDN', 'CsvPath', 'ContactsSearchBase', 'DistributionGroupsSearchBase', 'TenantId', 'Organization',
    'AppId', 'CertificateThumbprint', 'DefaultUsageLocation', 'GroupId', 'SkuPartNumber', 'HoldPolicy', 'ScriptPath'
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
$configTarget = Join-Path $Destination $configRelative
[IO.File]::WriteAllText($configTarget, $config, [Text.UTF8Encoding]::new($true))
$files.Add($configRelative)

# ---- Checks -------------------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'data', 'logs', 'reports', 'tests', 'tools') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { $_.Extension -in '.db', '.bak', '.xml', '.log' } | ForEach-Object { $problems.Add("File not allowed in the package: $($_.Name)") }
$textFiles = Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1', '.csv', '.md', '.html', '.txt' }
foreach ($value in ($emptied | Where-Object { $_.Length -ge 4 } | Select-Object -Unique)) {
    foreach ($file in $textFiles) {
        if ([IO.File]::ReadAllText($file.FullName).IndexOf($value, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            $problems.Add("An environment value of the configuration ($value) appears in $($file.FullName.Substring($Destination.Length + 1)).")
        }
    }
}
$check = Import-PowerShellDataFile -LiteralPath $configTarget
if ($check.Cloud.TenantId -or $check.Cloud.AppId -or $check.Cloud.CertificateThumbprint -or $check.Licensing.Users.GroupId -or $check.Retention.HoldPolicy) { $problems.Add('Tenant values are still in the packaged configuration.') }
foreach ($file in Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.ps1', '*.psm1', '*.psd1') {
    $errors = $null
    $null = [System.Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors)
    if ($errors) { $problems.Add("Parse error in $($file.Name): $($errors[0].Message)") }
}
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

if ($Zip) { Compress-Archive -LiteralPath $Destination -DestinationPath $zipPath -CompressionLevel Optimal }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  PRA Cloud Mailbox $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
if ($Zip) { Write-Host ("  Zip      : {0} ({1:N1} MB)" -f $zipPath, ((Get-Item -LiteralPath $zipPath).Length / 1MB)) }
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Config   : environment values emptied ($(@($emptied | Select-Object -Unique).Count)) - fill in the configuration (developer guide, chapter 6)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,12:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
