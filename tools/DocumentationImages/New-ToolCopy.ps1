<#
.SYNOPSIS
    A temporary copy of the package for the console pictures and tests: the real modules and the window, the synthetic engine instead of the real script (repository tool, never shipped).
.DESCRIPTION
    -RealEngine keeps the real entry script (the pictures of the consoles and of the reports). The copy goes through the integrity fixture of tests\IntegrityFixture.ps1 (only the signature check is substituted), so the real modules and the real window start from it.
    The pictures of docs\images that do not exist yet are empty placeholders and the obsolete ones are left out, so that the copy has exactly the inventory of the integrity check.
    Dot-source this file, then call New-DocumentationToolCopy.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
function New-DocumentationToolCopy {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Repository, [Parameter(Mandatory)][string]$Destination, [switch]$RealEngine)
    . (Join-Path $Repository 'tests\IntegrityFixture.ps1')
    $source = Join-Path $Repository 'package'
    $null = New-Item -ItemType Directory -Path $Destination -Force
    foreach ($item in Get-ChildItem -LiteralPath $source -Force) {
        if ($item.Name -in @('data', 'logs', 'reports')) { continue }
        Copy-Item -LiteralPath $item.FullName -Destination $Destination -Recurse -Force
    }
    $images = Join-Path $Destination 'docs\images'
    if (Test-Path -LiteralPath $images) { Remove-Item -LiteralPath $images -Recurse -Force }
    $null = New-Item -ItemType Directory -Path $images -Force
    foreach ($relative in @(Get-PcmExpectedInventory | Where-Object { $_ -like 'docs\images\*' })) {
        $real = Join-Path $source $relative
        $target = Join-Path $Destination $relative
        if (Test-Path -LiteralPath $real -PathType Leaf) { Copy-Item -LiteralPath $real -Destination $target -Force }
        else { [IO.File]::WriteAllBytes($target, [byte[]](0)) }
    }
    if ($RealEngine) { Initialize-PcmIntegrityFixture -Root $Destination -Changelog (Join-Path $Repository 'CHANGELOG.md') }
    else {
        Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'StubEngine.ps1') -Destination (Join-Path $Destination 'Invoke-PraCloudMailbox.ps1') -Force
        Initialize-PcmIntegrityFixture -Root $Destination -Changelog (Join-Path $Repository 'CHANGELOG.md') -Replaced @('Invoke-PraCloudMailbox.ps1')
    }
    return (Resolve-Path -LiteralPath $Destination).Path
}
