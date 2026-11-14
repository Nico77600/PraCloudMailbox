#Requires -Version 7.5
<#
.SYNOPSIS
    Writes package\PackageIntegrity.psd1, the SHA256 inventory that the owner signs (repository tool, not in the package).

.DESCRIPTION
    The inventory is the list of files of module\Test-PackageIntegrity.ps1 (the files it requires, plus CHANGELOG.md, which the package builder
    copies from the repository root). One entry per file: its path and its SHA256.

    Without -Final it writes a DRAFT: the SHA256 of the bytes of the working tree, signed or not, under a header that says so. A draft keeps the
    manifest in step with the files while the code changes; the package refuses to run from it (the signatures are missing), the test suites use
    their own fixture.

    With -Final it is the last step of a release, after the owner has signed every PowerShell file and the verifier: it refuses to write unless
    each PowerShell file of the inventory carries a valid, timestamped signature of the owner, and writes the manifest without the DRAFT
    header. The owner then signs PackageIntegrity.psd1 itself.

.PARAMETER Final
    Writes the manifest of a release (every PowerShell file must already be signed).

.EXAMPLE
    .\tools\Update-PackageIntegrity.ps1
    Rewrites the DRAFT manifest from the current files.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[CmdletBinding()]
param([switch]$Final)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$package = Join-Path $root 'package'
$verifier = Join-Path $package 'module\Test-PackageIntegrity.ps1'
$signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'

$ast = [Management.Automation.Language.Parser]::ParseFile($verifier, [ref]$null, [ref]$null)
function Get-VerifierList([string]$Name) {
    $assignment = $ast.Find({ param($node) $node -is [Management.Automation.Language.AssignmentStatementAst] -and $node.Left.Extent.Text -eq ('$' + $Name) }, $true)
    if (-not $assignment) { throw "The verifier has no list named $Name." }
    @($assignment.Right.FindAll({ param($node) $node -is [Management.Automation.Language.StringConstantExpressionAst] }, $true) | ForEach-Object { [string]$_.Value })
}
$inventory = @(Get-VerifierList 'required') + @(Get-VerifierList 'distributionOnly')
if (-not $inventory.Count) { throw 'The inventory of the verifier is empty.' }
$version = [string](Import-PowerShellDataFile -LiteralPath (Join-Path $package 'PraCloudMailbox.psd1')).ModuleVersion

$lines = foreach ($relative in $inventory) {
    $path = if ($relative -eq 'CHANGELOG.md') { Join-Path $root $relative } else { Join-Path $package $relative }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Missing file of the inventory: $relative" }
    if ($Final -and $relative -match '\.ps(m1|d1|1)$') {
        $signature = Get-AuthenticodeSignature -LiteralPath $path
        if ([string]$signature.Status -ne 'Valid' -or [string]$signature.SignerCertificate.Thumbprint -ne $signer -or $null -eq $signature.TimeStamperCertificate) {
            throw "Not signed by the owner (valid, timestamped): $relative. Sign every PowerShell file first."
        }
    }
    "        @{ Path = '$relative'; Sha256 = '$((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)' }"
}
$header = @("# PRA Cloud Mailbox $version - authenticated package integrity manifest")
if ($Final) { $header += '# Generated after the final stable-file signatures and timestamps.' }
else {
    $header += '# Generated only after the final stable-file signatures and timestamps.'
    $header += "# DRAFT FOR $version (unsigned): the SHA256 below are those of the bytes of the working tree when it was written, before any signature. The owner signs the PowerShell"
    $header += "# files, then runs tools\Update-PackageIntegrity.ps1 -Final and signs this file."
}
$text = ($header -join "`r`n") + "`r`n@{`r`n    SchemaVersion = 1`r`n    Algorithm = 'SHA256'`r`n    SignerThumbprint = '$signer'`r`n    Files = @(`r`n" + ($lines -join "`r`n") + "`r`n    )`r`n}`r`n"
$target = Join-Path $package 'PackageIntegrity.psd1'
[IO.File]::WriteAllText($target, $text, [Text.UTF8Encoding]::new($true))
Write-Host ('PackageIntegrity.psd1 written: {0} files ({1}).' -f $inventory.Count, $(if ($Final) { 'final' } else { 'draft' }))
