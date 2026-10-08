<#
.SYNOPSIS
    Runs the PRA Cloud Mailbox test gate offline and keeps the evidence.
.DESCRIPTION
    Pester 5+ (imported explicitly) and PSScriptAnalyzer (errors fail the gate). Run it in both editions:
    Collect runs in Windows PowerShell 5.1, the cloud actions in PowerShell 7.
.PARAMETER EvidenceDirectory
    Evidence folder; default tests\evidence\gate\<edition>-<timestamp>.
.EXAMPLE
    powershell.exe -NoProfile -File .\tests\Invoke-TestGate.ps1
.EXAMPLE
    pwsh -NoProfile -File .\tests\Invoke-TestGate.ps1
.NOTES
    Author : Nicolas Fabert
    Version: 0.1.0
#>
#Requires -Version 5.1
[CmdletBinding()]
param([string]$EvidenceDirectory)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (-not $EvidenceDirectory) { $EvidenceDirectory = Join-Path $root ('tests\evidence\gate\{0}-{1}' -f $PSVersionTable.PSEdition, (Get-Date -Format 'yyyyMMdd-HHmmss')) }
$null = New-Item -ItemType Directory -Path $EvidenceDirectory -Force
$summary = [ordered]@{ Edition = $PSVersionTable.PSEdition; PowerShell = [string]$PSVersionTable.PSVersion; Started = (Get-Date).ToString('o')
    Pester = ''; Analyzer = ''; Total = 0; Passed = 0; Failed = 0; Skipped = 0; AnalyzerErrors = 0; AnalyzerWarnings = 0; Result = 'FAIL' }
$exitCode = 1
try {
    $pester = Get-Module -ListAvailable Pester | Where-Object { $_.Version -ge [version]'5.0' } | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $pester) {
        $candidate = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules\Pester'
        $pester = Get-ChildItem $candidate -Directory -ErrorAction SilentlyContinue | Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1 |
            ForEach-Object { [pscustomobject]@{ Path = (Join-Path $_.FullName 'Pester.psd1'); Version = [version]$_.Name } }
    }
    if (-not $pester) { throw 'Pester 5 or later is required (Install-Module Pester -Scope CurrentUser -Force -SkipPublisherCheck).' }
    Import-Module $pester.Path -Force
    $summary.Pester = [string](Get-Module Pester).Version
    $configuration = New-PesterConfiguration
    $configuration.Run.Path = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.Tests.ps1' | Sort-Object Name | ForEach-Object { $_.FullName })
    $configuration.Run.PassThru = $true
    $configuration.Output.Verbosity = 'Normal'
    $configuration.TestResult.Enabled = $true
    $configuration.TestResult.OutputPath = Join-Path $EvidenceDirectory 'pester.xml'
    $result = Invoke-Pester -Configuration $configuration
    $summary.Total = $result.TotalCount; $summary.Passed = $result.PassedCount; $summary.Failed = $result.FailedCount; $summary.Skipped = $result.SkippedCount

    $analyzer = Get-Module -ListAvailable PSScriptAnalyzer | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $analyzer) {
        $candidate = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'PowerShell\Modules\PSScriptAnalyzer'
        $analyzer = Get-ChildItem $candidate -Directory -ErrorAction SilentlyContinue | Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1 |
            ForEach-Object { [pscustomobject]@{ Path = (Join-Path $_.FullName 'PSScriptAnalyzer.psd1'); Version = [version]$_.Name } }
    }
    $findings = @()
    if ($analyzer) {
        Import-Module $analyzer.Path -Force
        $summary.Analyzer = [string](Get-Module PSScriptAnalyzer).Version
        $files = @(Get-ChildItem $root -Recurse -File -Include *.ps1, *.psm1 | Where-Object { $_.FullName -notmatch '\\tests\\' })
        $excluded = @('PSAvoidUsingWriteHost', 'PSUseSingularNouns', 'PSUseShouldProcessForStateChangingFunctions', 'PSAvoidUsingPositionalParameters')
        # PSScriptAnalyzer 1.25 in Windows PowerShell 5.1 now and then throws a NullReferenceException on a file it analyses
        # without trouble the next time: one more try before the gate fails.
        $findings = @(foreach ($file in $files) {
                try { Invoke-ScriptAnalyzer -Path $file.FullName -ExcludeRule $excluded -ErrorAction Stop }
                catch { Write-Warning ("Analyzer, second try on {0}: {1}" -f $file.Name, $_.Exception.Message); Invoke-ScriptAnalyzer -Path $file.FullName -ExcludeRule $excluded -ErrorAction Stop }
            })
        $findings | Select-Object ScriptName, Line, Severity, RuleName, Message | Export-Csv (Join-Path $EvidenceDirectory 'analyzer.csv') -NoTypeInformation -Encoding UTF8
        $summary.AnalyzerErrors = @($findings | Where-Object Severity -eq 'Error').Count
        $summary.AnalyzerWarnings = @($findings | Where-Object Severity -eq 'Warning').Count
        foreach ($finding in @($findings | Where-Object { $_.Severity -in @('Error', 'Warning') })) { Write-Host ('  analyzer {0} {1}:{2} {3} {4}' -f $finding.Severity, $finding.ScriptName, $finding.Line, $finding.RuleName, $finding.Message) }
    } else { Write-Warning 'PSScriptAnalyzer not found: static analysis skipped.' }
    if ($result.FailedCount -eq 0 -and $result.FailedBlocksCount -eq 0 -and $result.FailedContainersCount -eq 0 -and $summary.AnalyzerErrors -eq 0 -and $result.TotalCount -gt 0) {
        $summary.Result = 'PASS'; $exitCode = 0
    }
} catch {
    $summary.Error = $_.Exception.Message
    Write-Host ('Gate error: {0}' -f $_.Exception.Message)
} finally {
    $summary.Finished = (Get-Date).ToString('o')
    [IO.File]::WriteAllText((Join-Path $EvidenceDirectory 'summary.json'), ($summary | ConvertTo-Json), (New-Object Text.UTF8Encoding($true)))
    Write-Host ('RESULT: {0} ({1} {2}: {3} passed, {4} failed, {5} skipped, analyzer {6} error(s) {7} warning(s)) - {8}' -f $summary.Result, $summary.Edition, $summary.PowerShell,
        $summary.Passed, $summary.Failed, $summary.Skipped, $summary.AnalyzerErrors, $summary.AnalyzerWarnings, $EvidenceDirectory)
}
exit $exitCode
