<#
.SYNOPSIS
    Controlled integrity fixtures for the offline PRA Cloud Mailbox suites (dot-sourced by the test files).
.DESCRIPTION
    The runtime fails closed on any missing, invalid or untrusted signature, and on any change after signing: the package
    in the working tree is only usable once the owner has signed the final bytes and regenerated PackageIntegrity.psd1. These helpers build a
    COPY of the package in which exactly one thing is substituted: every call of
    Microsoft.PowerShell.Security\Get-AuthenticodeSignature (one per file that carries the bootstrap, and one in the verifier) returns a
    Valid, timestamped, owner-thumbprint result. Everything else is the real code and the real bytes: the bootstrap, the verifier, the
    exact inventory, the SHA256 comparison, the path/reparse guards and the file scan. A manifest is generated from the copy's own bytes.
    This proves the integrity LOGIC. It does NOT prove Authenticode trust: the native signature state is asserted
    separately in PackageIntegrity.Tests.ps1, on the unmodified package, with no substitution. The two are never
    claimed to be equivalent.
.NOTES
    Author : Nicolas Fabert
    Version: 2.1.0
#>
#Requires -Version 7.5

$global:PcmIntegrity = @{
    Signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
    Token = 'Microsoft.PowerShell.Security\Get-AuthenticodeSignature'
    Generated = @('data', 'logs', 'reports')
}

function global:Get-PcmExpectedInventory {
    # Independent from the verifier: the tests compare both, so neither can drift silently.
    $images = 'cli-check-dark', 'cli-check-light', 'cli-convert-dark', 'cli-convert-group-dark',
        'cli-convert-group-light', 'cli-convert-light', 'cli-follow-dark', 'cli-follow-light',
        'cli-recover-dark', 'cli-recover-light', 'console-collect-next-dark', 'console-collect-next-light',
        'console-collect-run-done-dark', 'console-collect-run-done-light', 'console-collect-setup-dark', 'console-collect-setup-light',
        'console-collect-task-dark', 'console-collect-task-light', 'console-collect-verify-preview-dark', 'console-collect-verify-preview-light',
        'console-next-check-blocked-dark', 'console-next-check-blocked-light', 'console-next-dark', 'console-next-failed-dark',
        'console-next-failed-light', 'console-next-light', 'console-run-check-blocked-dark', 'console-run-check-blocked-light',
        'console-run-dark', 'console-run-done-dark', 'console-run-done-light', 'console-run-failed-dark',
        'console-run-failed-light', 'console-run-light', 'console-setup-dark', 'console-setup-light',
        'console-setup-settings-dark', 'console-setup-settings-light', 'console-setup-signin-dark', 'console-setup-signin-light',
        'console-task-check-dark', 'console-task-check-light', 'console-task-continue-dark', 'console-task-continue-light',
        'console-task-convert-csv-dark', 'console-task-convert-csv-light', 'console-task-convert-group-dark', 'console-task-convert-group-light',
        'console-task-convert-one-dark', 'console-task-convert-one-light', 'console-task-convert-ou-dark', 'console-task-convert-ou-light',
        'console-task-rollback-dark', 'console-task-rollback-light', 'console-task-rollback-wave-dark', 'console-task-rollback-wave-light',
        'console-verify-before-dark', 'console-verify-before-light', 'console-verify-confirmed-dark', 'console-verify-confirmed-light',
        'console-verify-preview-dark', 'console-verify-preview-light', 'readme-banner-dark', 'readme-banner-light',
        'readme-flows-dark', 'readme-flows-light', 'readme-how-it-works-dark', 'readme-how-it-works-light',
        'readme-permissions-dark', 'readme-permissions-light', 'readme-principles-dark', 'readme-principles-light',
        'readme-resources-dark', 'readme-resources-light', 'readme-states-dark', 'readme-states-light',
        'report-check-dark', 'report-check-light', 'report-convert-dark', 'report-convert-light',
        'report-global-dark', 'report-global-light'
    @(
        'Invoke-PraCloudMailbox.ps1', 'PraCloudMailbox.psd1', 'PraCloudMailbox.psm1'
        'module\PRA.Common.psm1', 'module\PRA.Store.psm1', 'module\PRA.Collect.psm1', 'module\PRA.Cloud.psm1', 'module\PRA.Gui.psm1'
        'module\PRA.Gui.Data.ps1', 'module\PRA.Gui.Detect.ps1', 'module\PRA.Gui.Process.ps1', 'module\PRA.Gui.Run.ps1', 'module\PRA.Gui.RunConfig.ps1', 'module\PRA.Gui.View.ps1'
        'module\Test-PackageIntegrity.ps1', 'module\PRA.Gui.xaml'
        'templates\Report.template.html', 'templates\GlobalReport.template.html'
        'config\Targets.sample.csv', 'config\EntraConnect.sample.ps1'
        'lib\sqlite\README.txt', 'lib\sqlite\netstandard2.0\System.Data.SQLite.dll', 'lib\sqlite\netstandard2.0\x64\SQLite.Interop.dll'
        'README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md'
        'docs\PraCloudMailbox-Guide.html', 'docs\PraCloudMailbox-Guide.md', 'docs\PraCloudMailbox-UserGuide.html', 'docs\PraCloudMailbox-UserGuide.md'
    ) + @($images | ForEach-Object { 'docs\images\' + $_ + '.png' }) + @('CHANGELOG.md')
}

function global:Get-PcmSignedPowerShellFiles {
    @(Get-PcmExpectedInventory | Where-Object { $_ -match '\.ps(m1|d1|1)$' })
}

function global:Get-PcmBootstrapFiles {
    # The files that carry the bootstrap (it authenticates the verifier): the entry script, the root module and the modules. The parts of the
    # window (PRA.Gui.*.ps1) are dot-sourced by PRA.Gui.psm1 and need no bootstrap of their own: the verifier covers them.
    @('Invoke-PraCloudMailbox.ps1', 'PraCloudMailbox.psm1', 'module\PRA.Common.psm1', 'module\PRA.Store.psm1', 'module\PRA.Collect.psm1', 'module\PRA.Cloud.psm1', 'module\PRA.Gui.psm1')
}

function global:Read-PcmFixtureText([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    $bom = $bytes.Length -ge 3 -and $bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191
    $encoding = New-Object Text.UTF8Encoding($false, $true)
    [pscustomobject]@{ Text = $(if ($bom) { $encoding.GetString($bytes, 3, $bytes.Length - 3) } else { $encoding.GetString($bytes) }); Bom = $bom }
}

function global:Write-PcmFixtureText([string]$Path, [string]$Text, [bool]$Bom) {
    $body = (New-Object Text.UTF8Encoding($false)).GetBytes($Text)
    if ($Bom) { $body = [byte[]](239, 187, 191) + $body }
    [IO.File]::WriteAllBytes($Path, $body)
}

function global:Get-PcmSignatureSubstitute {
    '& { [CmdletBinding()] param([string]$LiteralPath) [pscustomobject]@{ Status = ''Valid''; SignerCertificate = [pscustomobject]@{ Thumbprint = ''' +
        $global:PcmIntegrity.Signer + ''' }; TimeStamperCertificate = [pscustomobject]@{ Subject = ''TEST-FIXTURE-SUBSTITUTE'' } } }'
}

function global:Write-PcmFixtureManifest {
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$Changelog)
    $lines = foreach ($relative in Get-PcmExpectedInventory) {
        $path = if ($relative -eq 'CHANGELOG.md') { $Changelog } else { Join-Path $Root $relative }
        "        @{ Path = '$relative'; Sha256 = '$((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)' }"
    }
    $text = "# TEST FIXTURE manifest generated from the copied bytes; the signature check is substituted. Never a release artifact.`r`n" +
        "@{`r`n    SchemaVersion = 1`r`n    Algorithm = 'SHA256'`r`n    SignerThumbprint = '$($global:PcmIntegrity.Signer)'`r`n    Files = @(`r`n" +
        ($lines -join "`r`n") + "`r`n    )`r`n}`r`n"
    Write-PcmFixtureText (Join-Path $Root 'PackageIntegrity.psd1') $text $true
}

function global:Initialize-PcmIntegrityFixture {
    <# Turns a copy of the package into a fixture: substitutes the signature check only and writes a manifest from the copy's bytes. #>
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$Changelog, [switch]$IncludeChangelog, [switch]$KeepNativeSignatureCheck, [string[]]$Replaced = @())
    foreach ($folder in $global:PcmIntegrity.Generated) { Remove-Item -LiteralPath (Join-Path $Root $folder) -Recurse -Force -ErrorAction SilentlyContinue }
    if (-not $KeepNativeSignatureCheck) {
        foreach ($relative in (Get-PcmBootstrapFiles) + @('module\Test-PackageIntegrity.ps1')) {
            if ($relative -in $Replaced) { continue }
            $path = Join-Path $Root $relative
            $file = Read-PcmFixtureText $path
            $count = ([regex]::Matches($file.Text, [regex]::Escape($global:PcmIntegrity.Token))).Count
            if ($count -ne 1) { throw "Integrity fixture: expected exactly one signature call in $relative, found $count." }
            Write-PcmFixtureText $path $file.Text.Replace($global:PcmIntegrity.Token, (Get-PcmSignatureSubstitute)) $file.Bom
        }
    }
    Write-PcmFixtureManifest -Root $Root -Changelog $Changelog
    if ($IncludeChangelog) { Copy-Item -LiteralPath $Changelog -Destination (Join-Path $Root 'CHANGELOG.md') }
}

function global:New-PcmIntegrityFixture {
    <# Copies the package, substitutes the signature check only, writes a manifest from the copy's bytes. #>
    param([Parameter(Mandatory)][string]$SourcePackage, [Parameter(Mandatory)][string]$Destination,
        [Parameter(Mandatory)][string]$Changelog, [switch]$IncludeChangelog, [switch]$KeepNativeSignatureCheck)
    $null = New-Item -ItemType Directory -Path (Split-Path $Destination -Parent) -Force
    Copy-Item -LiteralPath $SourcePackage -Destination $Destination -Recurse
    Initialize-PcmIntegrityFixture -Root $Destination -Changelog $Changelog -IncludeChangelog:$IncludeChangelog -KeepNativeSignatureCheck:$KeepNativeSignatureCheck
}

function global:Remove-PcmTestPackage {
    <# Removes the temporary folder of a test package (the SQLite library that a process has loaded stays until that process ends). #>
    param([string]$Package)
    if (-not $Package) { return }
    $parent = Split-Path $Package -Parent
    if ((Split-Path $parent -Leaf) -like 'PCM-*') { Remove-Item -LiteralPath $parent -Recurse -Force -ErrorAction SilentlyContinue }
}

function global:New-PcmTestPackage {
    <# The package of the repository, as a verified copy under TEMP, for the suites that import the modules or run the entry script. #>
    param([Parameter(Mandatory)][string]$Repository, [string]$Parent = [IO.Path]::GetTempPath())
    # A test file loads its modules from its own copy of the package: the modules of the copy of a previous file must not stay loaded under the same names.
    Get-Module | Where-Object { $_.Name -like 'PRA.*' -or $_.Name -eq 'PraCloudMailbox' } | Remove-Module -Force -ErrorAction SilentlyContinue
    $destination = Join-Path $Parent ('PCM-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '\package')
    New-PcmIntegrityFixture -SourcePackage (Join-Path $Repository 'package') -Destination $destination -Changelog (Join-Path $Repository 'CHANGELOG.md')
    $destination
}

function global:Set-PcmFixtureHash {
    <# Re-hashes one inventory file in the fixture manifest after a deliberate fixture edit. #>
    param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$Relative)
    $manifest = Join-Path $Root 'PackageIntegrity.psd1'
    $file = Read-PcmFixtureText $manifest
    $hash = (Get-FileHash -LiteralPath (Join-Path $Root $Relative) -Algorithm SHA256).Hash
    $pattern = "(?m)(Path = '" + [regex]::Escape($Relative) + "'; Sha256 = ')[0-9A-F]{64}(')"
    if (-not [regex]::IsMatch($file.Text, $pattern)) { throw "Integrity fixture: $Relative is not in the manifest." }
    Write-PcmFixtureText $manifest ([regex]::Replace($file.Text, $pattern, ('${1}' + $hash + '${2}'))) $file.Bom
}
