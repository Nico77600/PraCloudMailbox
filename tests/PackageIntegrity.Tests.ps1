<#
.SYNOPSIS
    Package integrity tests for PRA Cloud Mailbox (offline; no Exchange, Graph, signing key or network).
.DESCRIPTION
    Three kinds of evidence, never mixed up:
      * Static: bootstrap identity and order, exact inventory, git attributes, no bypass, no private data.
      * Controlled fixture (IntegrityFixture.ps1): a COPY of the package where only the Authenticode call is
        substituted. It exercises the real verifier, bootstrap, inventory, hash, path and reparse logic. It does not
        prove signature trust.
      * Native: the unmodified package with the real Get-AuthenticodeSignature. While the package is unsigned it must
        fail closed before any import or operation; once the owner has signed every file it must verify. The
        assertion depends on the observed signature state, so it is never skipped and never weakened.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $script:Source = Join-Path $script:Repo 'package'
    $script:Work = Join-Path ([IO.Path]::GetTempPath()) ('PI-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $null = New-Item -ItemType Directory -Path $script:Work -Force
    $script:Pwsh = Join-Path $PSHOME 'pwsh.exe'
    $script:Ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:Changelog = Join-Path $script:Repo 'CHANGELOG.md'
    $script:Fixture = Join-Path $script:Work 'fixture'
    New-PcmIntegrityFixture -SourcePackage $script:Source -Destination $script:Fixture -Changelog $script:Changelog -IncludeChangelog

    function script:Invoke-Pwsh {
        param([Parameter(Mandatory)][string[]]$Arguments)
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try { $output = & $script:Pwsh -NoLogo -NoProfile -NonInteractive @Arguments 2>&1 | ForEach-Object { [string]$_ } }
        finally { $ErrorActionPreference = $previous }
        [pscustomobject]@{ Code = $LASTEXITCODE; Text = ($output -join "`n") }
    }
    function script:Copy-Case {
        $case = Join-Path $script:Work ([guid]::NewGuid().ToString('N').Substring(0, 8))
        Copy-Item -LiteralPath $script:Fixture -Destination $case -Recurse
        $case
    }
    function script:Invoke-Verifier([string]$Root, [switch]$Distribution) {
        $arguments = @('-File', (Join-Path $Root 'module\Test-PackageIntegrity.ps1'), '-Root', $Root)
        if ($Distribution) { $arguments += '-Distribution' }
        Invoke-Pwsh -Arguments $arguments
    }
    function script:Edit-Manifest([string]$Root, [scriptblock]$Edit) {
        $path = Join-Path $Root 'PackageIntegrity.psd1'
        $file = Read-PcmFixtureText $path
        Write-PcmFixtureText $path (& $Edit $file.Text) $true
    }
    function script:Add-Byte([string]$Path) {
        $stream = [IO.File]::Open($Path, 'Append', 'Write')
        try { $stream.WriteByte(32) } finally { $stream.Dispose() }
    }
    function script:Get-NativeState {
        $names = @(Get-PcmSignedPowerShellFiles) + @('PackageIntegrity.psd1')
        $statuses = @($names | ForEach-Object { [string](Get-AuthenticodeSignature -LiteralPath (Join-Path $script:Source $_)).Status })
        [pscustomobject]@{ AllValid = (@($statuses | Where-Object { $_ -ne 'Valid' }).Count -eq 0); Statuses = $statuses }
    }
}

AfterAll {
    if ($script:Work -and [IO.Directory]::Exists($script:Work)) {
        Get-ChildItem -LiteralPath $script:Work -Recurse -Force -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint } | ForEach-Object { [IO.Directory]::Delete($_.FullName) }
        Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue
    }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Static package contract' {
    It 'keeps the verifier inventory identical to the independently listed exact inventory' {
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Source 'module\Test-PackageIntegrity.ps1'), [ref]$null, [ref]$null)
        $assignment = $ast.Find({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$required' }, $true)
        $required = @($assignment.Right.FindAll({ param($n) $n -is [Management.Automation.Language.StringConstantExpressionAst] }, $true) | ForEach-Object Value)
        $expected = @(Get-PcmExpectedInventory | Where-Object { $_ -ne 'CHANGELOG.md' })
        (($required | Sort-Object) -join '|') | Should -BeExactly (($expected | Sort-Object) -join '|')
        $text = [IO.File]::ReadAllText((Join-Path $script:Source 'module\Test-PackageIntegrity.ps1'))
        $text | Should -Match "\`$distributionOnly = @\('CHANGELOG\.md'\)"
    }
    It 'has every inventory file in the source package and no unlisted stable file' {
        foreach ($relative in Get-PcmExpectedInventory | Where-Object { $_ -ne 'CHANGELOG.md' }) { Test-Path -LiteralPath (Join-Path $script:Source $relative) -PathType Leaf | Should -BeTrue -Because $relative }
        Test-Path -LiteralPath $script:Changelog -PathType Leaf | Should -BeTrue
        $inventory = @(Get-PcmExpectedInventory) + @('PackageIntegrity.psd1', 'config\PraCloudMailbox.config.psd1')
        $unlisted = @(Get-ChildItem -LiteralPath $script:Source -Recurse -File -Force | ForEach-Object { $_.FullName.Substring($script:Source.Length + 1) } |
                Where-Object { $_ -notin $inventory -and $_ -notmatch '^(data|logs|reports)\\' })
        $unlisted | Should -BeNullOrEmpty
    }
    It 'covers every executable PowerShell file of the package, plus the inventory manifest as a signed file' {
        $executables = @(Get-ChildItem -LiteralPath $script:Source -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1' } |
                ForEach-Object { $_.FullName.Substring($script:Source.Length + 1) } | Where-Object { $_ -ne 'config\PraCloudMailbox.config.psd1' })
        $signed = @(Get-PcmSignedPowerShellFiles) + @('PackageIntegrity.psd1')
        (($executables | Sort-Object) -join '|') | Should -BeExactly (($signed | Sort-Object) -join '|')
    }
    It 'runs an identical bootstrap, before any import or function, in the entry, the root module and the five modules' {
        $files = @(Get-PcmBootstrapFiles)
        $files.Count | Should -Be 7
        $blocks = @()
        foreach ($relative in $files) {
            $text = [IO.File]::ReadAllText((Join-Path $script:Source $relative)).Replace("`r`n", "`n")
            $match = [regex]::Match($text, '(?s)# >>> PRA package integrity bootstrap.*?# <<< PRA package integrity bootstrap')
            $match.Success | Should -BeTrue -Because $relative
            ([regex]::Matches($text, '# >>> PRA package integrity bootstrap')).Count | Should -Be 1 -Because $relative
            $blocks += ($match.Value -replace '(?m)^\} .*$', '} <root>')
            $before = $text.Substring(0, $match.Index)
            $before | Should -Not -Match '(?m)^\s*(Import-Module|function )' -Because $relative
        }
        @($blocks | Select-Object -Unique).Count | Should -Be 1
    }
    It 'has no bootstrap in the parts of the window (the verifier covers them), which the module loads after its own bootstrap' {
        foreach ($relative in Get-PcmSignedPowerShellFiles | Where-Object { $_ -like 'module\PRA.Gui.*.ps1' }) {
            [IO.File]::ReadAllText((Join-Path $script:Source $relative)) | Should -Not -Match 'integrity bootstrap' -Because $relative
        }
        $gui = [IO.File]::ReadAllText((Join-Path $script:Source 'module\PRA.Gui.psm1'))
        $gui.IndexOf('# <<< PRA package integrity bootstrap') | Should -BeLessThan $gui.IndexOf("PRA.Gui.Data.ps1')")
    }
    It 'has exactly eight native signature call sites in the package and none unqualified' {
        $sites = 0
        foreach ($file in Get-ChildItem -LiteralPath $script:Source -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' }) {
            $text = [IO.File]::ReadAllText($file.FullName)
            $sites += ([regex]::Matches($text, [regex]::Escape($global:PcmIntegrity.Token))).Count
            ([regex]::Matches($text, '(?<!Security\\)Get-AuthenticodeSignature')).Count | Should -Be 0 -Because $file.Name
        }
        $sites | Should -Be 8
    }
    It 'offers no unsigned mode, switch or environment bypass' {
        foreach ($file in Get-ChildItem -LiteralPath $script:Source -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' }) {
            [IO.File]::ReadAllText($file.FullName) | Should -Not -Match '(?i)AllowUnsigned|SkipIntegrity|NoIntegrity|DisableIntegrity|PRA_(SKIP|ALLOW|BYPASS)\w*|env:\w*(UNSIGNED|INTEGRITY)' -Because $file.Name
        }
    }
    It 'pins the fixed owner thumbprint in the verifier, the bootstrap and the builder' {
        foreach ($path in @((Join-Path $script:Source 'module\Test-PackageIntegrity.ps1'), (Join-Path $script:Source 'Invoke-PraCloudMailbox.ps1'), (Join-Path $script:Repo 'tools\New-PraCloudPackage.ps1'))) {
            [IO.File]::ReadAllText($path) | Should -Match ([regex]::Escape($global:PcmIntegrity.Signer))
        }
    }
    It 'keeps the signed and hashed paths byte-exact through .gitattributes (-text, cr-at-eol on signed files)' {
        $signed = @(Get-PcmSignedPowerShellFiles | ForEach-Object { 'package/' + $_.Replace('\', '/') }) + @('package/PackageIntegrity.psd1')
        $hashed = @(Get-PcmExpectedInventory | Where-Object { $_ -notmatch '\.ps(m1|d1|1)$' } | ForEach-Object { if ($_ -eq 'CHANGELOG.md') { $_ } else { 'package/' + $_.Replace('\', '/') } })
        Push-Location $script:Repo
        try {
            foreach ($path in $signed + $hashed) {
                $text = (& git check-attr text -- $path) -join ' '
                $text | Should -Match ': text: unset$' -Because $path
            }
            foreach ($path in $signed) {
                ((& git check-attr whitespace -- $path) -join ' ') | Should -Match 'cr-at-eol' -Because $path
            }
        }
        finally { Pop-Location }
    }
    It 'distributes no data, configuration secrets, tests or tools (package source folders and values)' {
        foreach ($forbidden in 'tests', 'tools', 'data', 'logs', 'reports', '.github', '.git') { Test-Path -LiteralPath (Join-Path $script:Source $forbidden) | Should -BeFalse -Because $forbidden }
        @(Get-ChildItem -LiteralPath (Join-Path $script:Source 'config') -File | ForEach-Object Name | Sort-Object) -join '|' |
            Should -BeExactly 'EntraConnect.sample.ps1|PraCloudMailbox.config.psd1|Targets.sample.csv'
    }
    It 'ships the SQLite library of PowerShell 7 only' {
        @(Get-ChildItem -LiteralPath (Join-Path $script:Source 'lib\sqlite') -Directory | ForEach-Object Name) | Should -Be @('netstandard2.0')
    }
}

Describe 'Module manifest - PraCloudMailbox.psd1 and its root module' {
    BeforeAll {
        $script:ManifestPath = Join-Path $script:Source 'PraCloudMailbox.psd1'
        $script:ManifestData = Import-PowerShellDataFile -LiteralPath $script:ManifestPath
        function script:Get-ModuleExports([string]$Relative) {
            $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Source $Relative), [ref]$null, [ref]$null)
            $export = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Export-ModuleMember' }, $true)
            @($export.CommandElements | Select-Object -Skip 1 | ForEach-Object { $_.FindAll({ param($n) $n -is [Management.Automation.Language.StringConstantExpressionAst] }, $true) } |
                    ForEach-Object Value | Where-Object { $_ -cne 'Function' })
        }
        function script:Get-CommandNames([string]$Relative) {
            $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Source $Relative), [ref]$null, [ref]$null)
            @($ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() } | Where-Object { $_ })
        }
    }
    It 'passes Test-ModuleManifest, requires PowerShell 7.5 and is Core only' {
        $info = Test-ModuleManifest -Path $script:ManifestPath -ErrorAction Stop
        $info.Name | Should -Be 'PraCloudMailbox'
        $info.Author | Should -Be 'Nicolas Fabert'
        $info.PowerShellVersion | Should -Be ([version]'7.5')
        @($info.CompatiblePSEditions) | Should -Be @('Core')
        $info.RootModule | Should -Be 'PraCloudMailbox.psm1'
    }
    It 'has one version: the manifest, Get-PcmToolVersion and the top entry of the changelog' {
        $version = [string]$script:ManifestData.ModuleVersion
        $version | Should -Match '^\d+\.\d+\.\d+$'
        $result = Invoke-Pwsh -Arguments @('-Command', "Import-Module -Name '$(Join-Path $script:Fixture 'PraCloudMailbox.psd1')' -Force; Get-PcmToolVersion")
        $result.Text.Trim() | Should -BeExactly $version
        $top = [regex]::Match([IO.File]::ReadAllText($script:Changelog), '(?m)^## (\d+\.\d+\.\d+)').Groups[1].Value
        $top | Should -BeExactly $version
        foreach ($relative in 'Invoke-PraCloudMailbox.ps1', 'module\PRA.Common.psm1', 'module\PRA.Store.psm1', 'module\PRA.Collect.psm1', 'module\PRA.Cloud.psm1', 'module\PRA.Gui.psm1', 'PraCloudMailbox.psm1') {
            [IO.File]::ReadAllText((Join-Path $script:Source $relative)) | Should -Match ("(?m)^\s*Version\s*:\s*{0}" -f [regex]::Escape($version)) -Because $relative
        }
    }
    It 'exports an explicit list: no wildcard, no duplicate, exactly the functions that PRA.Common and PRA.Store export' {
        $list = @($script:ManifestData.FunctionsToExport)
        $list.Count | Should -BeGreaterThan 0
        foreach ($name in $list) { $name | Should -Not -Match '[*?]' }
        @($list | Select-Object -Unique).Count | Should -Be $list.Count
        $defined = @(Get-ModuleExports 'module\PRA.Common.psm1') + @(Get-ModuleExports 'module\PRA.Store.psm1')
        (($list | Sort-Object) -join ',') | Should -BeExactly (($defined | Sort-Object) -join ',')
    }
    It 'exports every function that the entry script, the window and the other modules call across a module boundary' {
        $common = @(Get-ModuleExports 'module\PRA.Common.psm1')
        $store = @(Get-ModuleExports 'module\PRA.Store.psm1')
        $callers = [ordered]@{
            'Invoke-PraCloudMailbox.ps1' = $common + $store; 'module\PRA.Gui.psm1' = $common + $store; 'module\PRA.Collect.psm1' = $common + $store
            'module\PRA.Cloud.psm1' = $common + $store; 'module\PRA.Store.psm1' = $common
        }
        $used = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
        foreach ($relative in $callers.Keys) {
            foreach ($name in Get-CommandNames $relative) { if ($callers[$relative] -ccontains $name) { [void]$used.Add($name) } }
        }
        $listed = @($script:ManifestData.FunctionsToExport)
        $missing = @($used | Where-Object { $listed -cnotcontains $_ } | Sort-Object)
        ($missing -join ',') | Should -BeExactly '' -Because 'called across a module boundary but not exported'
    }
    It 'is imported by the entry script, which no longer imports Common or Store by path' {
        $text = [IO.File]::ReadAllText((Join-Path $script:Source 'Invoke-PraCloudMailbox.ps1'))
        $text | Should -Match "PraCloudMailbox\.psd1'\) -Force"
        $text | Should -Not -Match 'PRA\.(Common|Store)\.psm1'
    }
    It 'keeps Collect, the cloud actions and the window on demand: the root module loads none of them' {
        $root = [IO.File]::ReadAllText((Join-Path $script:Source 'PraCloudMailbox.psm1'))
        $root | Should -Not -Match '(?m)^\s*Import-Module .*PRA\.(Collect|Cloud|Gui)\.psm1'
        ([regex]::Matches($root, '(?m)^\s*Import-Module ')).Count | Should -Be 2
        $script:ManifestData.ContainsKey('NestedModules') | Should -BeFalse
        $entry = [IO.File]::ReadAllText((Join-Path $script:Source 'Invoke-PraCloudMailbox.ps1'))
        ([regex]::Matches($entry, 'Import-Module \(Join-Path \$PSScriptRoot ''module\\PRA\.Gui\.psm1''\)')).Count | Should -Be 1
        ([regex]::Matches($entry, 'Import-Module \(Join-Path \$PSScriptRoot ''module\\PRA\.Collect\.psm1''\)')).Count | Should -Be 1
        ([regex]::Matches($entry, 'Import-Module \(Join-Path \$PSScriptRoot ''module\\PRA\.Cloud\.psm1''\)')).Count | Should -Be 1
    }
    It 'loads the exact list of the manifest, and only PRA.Common and PRA.Store, in a fresh process' {
        $manifest = Join-Path $script:Fixture 'PraCloudMailbox.psd1'
        $command = @(
            'try {'
            "Import-Module -Name '$manifest' -Force -ErrorAction Stop"
            '''EXPORTED='' + (@((Get-Module PraCloudMailbox).ExportedFunctions.Keys) -join '','')'
            '''PARTS='' + (@(Get-Module -All | Where-Object { $_.Name -like ''PRA.*'' } | ForEach-Object Name | Sort-Object) -join '','')'
            '} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }'
        ) -join "`n"
        $result = Invoke-Pwsh -Arguments @('-Command', $command)
        $result.Code | Should -Be 0 -Because $result.Text
        $exported = [regex]::Match($result.Text, '(?m)^EXPORTED=(.*)$').Groups[1].Value.Trim().Split(',')
        $exported.Count | Should -Be @($script:ManifestData.FunctionsToExport).Count
        foreach ($name in $script:ManifestData.FunctionsToExport) { $exported -ccontains $name | Should -BeTrue -Because $name }
        [regex]::Match($result.Text, '(?m)^PARTS=(.*)$').Groups[1].Value.Trim() | Should -BeExactly 'PRA.Common,PRA.Store'
    }
    It 'exports only approved verbs and prints no warning when the manifest and every module are imported (the console start shows none)' {
        $paths = @("'$(Join-Path $script:Fixture 'PraCloudMailbox.psd1')'") + @(Get-ChildItem -LiteralPath (Join-Path $script:Fixture 'module') -Filter '*.psm1' | Sort-Object Name | ForEach-Object { "'$($_.FullName)'" })
        $command = @(
            'try {'
            '$warned = New-Object ''Collections.Generic.List[string]'''
            "foreach (`$path in @($($paths -join ','))) {"
            '$module = Import-Module -Name $path -Force -PassThru -ErrorAction Stop -WarningVariable captured'
            'foreach ($item in @($captured)) { $warned.Add([string]$item) }'
            'foreach ($name in $module.ExportedCommands.Keys) { ''COMMAND='' + $module.Name + '':'' + $name }'
            '}'
            '''WARNINGS='' + $warned.Count'
            '} catch { [Console]::Error.WriteLine($_.Exception.Message); exit 1 }'
        ) -join "`n"
        $result = Invoke-Pwsh -Arguments @('-Command', $command)
        $result.Code | Should -Be 0 -Because $result.Text
        [regex]::Match($result.Text, '(?m)^WARNINGS=(.*)$').Groups[1].Value.Trim() | Should -BeExactly '0' -Because $result.Text
        $result.Text | Should -Not -Match '(?m)^WARNING:' -Because 'a warning text would be printed at the console start'
        $verbs = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([string[]]@(Get-Verb | ForEach-Object Verb)), ([StringComparer]::Ordinal)
        $commands = @([regex]::Matches($result.Text, '(?m)^COMMAND=(?<Module>[^:]+):(?<Name>\S+)\s*$') | ForEach-Object { [pscustomobject]@{ Module = $_.Groups['Module'].Value; Name = $_.Groups['Name'].Value } })
        $commands.Count | Should -BeGreaterThan 100
        foreach ($module in 'PraCloudMailbox', 'PRA.Common', 'PRA.Store', 'PRA.Collect', 'PRA.Cloud', 'PRA.Gui') {
            @($commands | Where-Object { $_.Module -ceq $module }).Count | Should -BeGreaterThan 0 -Because $module
        }
        foreach ($command in $commands) {
            $verbs.Contains($command.Name.Split('-')[0]) | Should -BeTrue -Because "$($command.Name) (exported by $($command.Module)) must use a verb from Get-Verb"
        }
    }
    It 'is refused by Windows PowerShell 5.1 with the minimum version in the message' -Skip:(-not (Test-Path (Join-Path $env:SystemRoot System32\WindowsPowerShell\v1.0\powershell.exe))) {
        $command = "try { Import-Module -Name '$(Join-Path $script:Fixture 'PraCloudMailbox.psd1')' -Force -ErrorAction Stop; 'LOADED' } catch { [Console]::Error.WriteLine(`$_.Exception.Message); exit 1 }"
        $output = & $script:Ps51 -NoLogo -NoProfile -NonInteractive -Command $command 2>&1 | ForEach-Object { [string]$_ }
        $LASTEXITCODE | Should -Not -Be 0
        ($output -join "`n") | Should -Match '7\.5'
        ($output -join "`n") | Should -Not -Match 'LOADED|Package integrity'
    }
}

Describe 'Controlled fixture - real verifier, bootstrap, hash and path logic (signature call substituted)' {
    It 'accepts the fixture package, with and without the distribution-only changelog' {
        (Invoke-Verifier $script:Fixture -Distribution).Code | Should -Be 0
        $case = Copy-Case
        Remove-Item -LiteralPath (Join-Path $case 'CHANGELOG.md')
        (Invoke-Verifier $case).Code | Should -Be 0
        $result = Invoke-Verifier $case -Distribution
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity:'
    }
    It 'rejects a changed stable file <Name>' -ForEach @(
        @{ Name = 'entry'; Path = 'Invoke-PraCloudMailbox.ps1' }, @{ Name = 'Common module'; Path = 'module\PRA.Common.psm1' },
        @{ Name = 'Store module'; Path = 'module\PRA.Store.psm1' }, @{ Name = 'Collect module'; Path = 'module\PRA.Collect.psm1' },
        @{ Name = 'Cloud module'; Path = 'module\PRA.Cloud.psm1' }, @{ Name = 'window part'; Path = 'module\PRA.Gui.Run.ps1' },
        @{ Name = 'module manifest'; Path = 'PraCloudMailbox.psd1' }, @{ Name = 'root module'; Path = 'PraCloudMailbox.psm1' },
        @{ Name = 'XAML'; Path = 'module\PRA.Gui.xaml' }, @{ Name = 'report template'; Path = 'templates\Report.template.html' },
        @{ Name = 'sample CSV'; Path = 'config\Targets.sample.csv' }, @{ Name = 'Entra Connect sample'; Path = 'config\EntraConnect.sample.ps1' },
        @{ Name = 'SQLite library'; Path = 'lib\sqlite\netstandard2.0\System.Data.SQLite.dll' }, @{ Name = 'README'; Path = 'README.md' }, @{ Name = 'licence'; Path = 'LICENSE' },
        @{ Name = 'notices'; Path = 'THIRD-PARTY-NOTICES.md' }, @{ Name = 'guide'; Path = 'docs\PraCloudMailbox-Guide.html' }, @{ Name = 'user guide'; Path = 'docs\PraCloudMailbox-UserGuide.html' },
        @{ Name = 'image'; Path = 'docs\images\console-setup-light.png' }, @{ Name = 'changelog'; Path = 'CHANGELOG.md' }, @{ Name = 'verifier'; Path = 'module\Test-PackageIntegrity.ps1' }
    ) {
        $case = Copy-Case
        Add-Byte (Join-Path $case $Path)
        $result = Invoke-Verifier $case -Distribution
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity:'
    }
    It 'rejects a missing stable file <Path>' -ForEach @(@{ Path = 'module\PRA.Gui.xaml' }, @{ Path = 'docs\images\console-setup-light.png' }, @{ Path = 'LICENSE' }, @{ Path = 'lib\sqlite\netstandard2.0\x64\SQLite.Interop.dll' }) {
        $case = Copy-Case
        Remove-Item -LiteralPath (Join-Path $case $Path)
        $result = Invoke-Verifier $case -Distribution
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity:'
    }
    It 'rejects an unexpected stable file at <Path>' -ForEach @(
        @{ Path = 'extra.txt' }, @{ Path = 'module\Extra.ps1' }, @{ Path = 'docs\notes.md' }, @{ Path = 'templates\Other.html' },
        @{ Path = 'config\Extra.psd1' }, @{ Path = 'config\nested\data.csv' }, @{ Path = 'docs\data\x.txt' }, @{ Path = 'lib\sqlite\net46\System.Data.SQLite.dll' }
    ) {
        $case = Copy-Case
        $null = New-Item -ItemType Directory -Path (Split-Path (Join-Path $case $Path) -Parent) -Force
        [IO.File]::WriteAllText((Join-Path $case $Path), 'x')
        $result = Invoke-Verifier $case -Distribution
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'unexpected stable file'
    }
    It 'permits editable configuration, the credential file, operational CSV lists and generated folders (data, logs, reports)' {
        $case = Copy-Case
        [IO.File]::WriteAllText((Join-Path $case 'config\PraCloudMailbox.config.psd1'), "@{ Anything = 'edited' }")
        [IO.File]::WriteAllText((Join-Path $case 'config\EntraConnect.cred.xml'), '<Objs/>')
        [IO.File]::WriteAllText((Join-Path $case 'config\Operational-List.csv'), 'a;b')
        foreach ($folder in 'data', 'logs', 'reports') {
            $null = New-Item -ItemType Directory -Path (Join-Path $case "$folder\sub") -Force
            [IO.File]::WriteAllText((Join-Path $case "$folder\sub\data.bin"), 'generated')
        }
        $result = Invoke-Verifier $case -Distribution
        $result.Code | Should -Be 0 -Because $result.Text
    }
    It 'rejects a manifest <Name>' -ForEach @(
        @{ Name = 'with a wrong thumbprint'; Edit = { param($t) $t.Replace('0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6', '1111111111111111111111111111111111111111') } }
        @{ Name = 'with another algorithm'; Edit = { param($t) $t.Replace("'SHA256'", "'SHA1'") } }
        @{ Name = 'with an extra key'; Edit = { param($t) $t.Replace('    Algorithm', "    Extra = 1`r`n    Algorithm") } }
        @{ Name = 'with a duplicated path'; Edit = { param($t) $t.Replace("Path = 'LICENSE'", "Path = 'README.md'") } }
        @{ Name = 'with a traversal path'; Edit = { param($t) $t.Replace("Path = 'LICENSE'", "Path = '..\LICENSE'") } }
        @{ Name = 'with a lowercase hash'; Edit = { param($t) [regex]::Replace($t, "Sha256 = '([0-9A-F]{64})'", { param($m) "Sha256 = '" + $m.Groups[1].Value.ToLowerInvariant() + "'" }, 'None') } }
        @{ Name = 'with a missing entry'; Edit = { param($t) [regex]::Replace($t, "(?m)^\s*@\{ Path = 'LICENSE'.*\r?\n", '') } }
        @{ Name = 'with a schema version 2'; Edit = { param($t) $t.Replace('SchemaVersion = 1', 'SchemaVersion = 2') } }
        @{ Name = 'empty (the unsigned placeholder shape)'; Edit = { param($t) [regex]::Replace($t, '(?s)Files = @\(.*\r?\n    \)', 'Files = @()') } }
    ) {
        $case = Copy-Case
        Edit-Manifest $case $Edit
        $result = Invoke-Verifier $case -Distribution
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity:|Import-PowerShellDataFile|parse|Hash'
    }
    It 'rejects a reparse point in the stable tree (<Path>)' -ForEach @(@{ Path = 'docs\images' }, @{ Path = 'templates' }, @{ Path = 'module' }, @{ Path = 'lib' }) {
        $case = Copy-Case
        $moved = Join-Path $script:Work ('moved-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        Move-Item -LiteralPath (Join-Path $case $Path) -Destination $moved
        $null = New-Item -ItemType Junction -Path (Join-Path $case $Path) -Target $moved
        $verifier = if ($Path -eq 'module') { Join-Path $moved 'Test-PackageIntegrity.ps1' } else { Join-Path $case 'module\Test-PackageIntegrity.ps1' }
        $result = Invoke-Pwsh -Arguments @('-File', $verifier, '-Root', $case, '-Distribution')
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity:'
    }
    It 'fails the entry (CLI and window) before any import, configuration or operation when a stable file changed' {
        $case = Copy-Case
        Add-Byte (Join-Path $case 'LICENSE')
        foreach ($arguments in @(@('-Action', 'Check', '-Mode', 'Preview'), @('-Gui'))) {
            $result = Invoke-Pwsh -Arguments (@('-File', (Join-Path $case 'Invoke-PraCloudMailbox.ps1')) + $arguments)
            $result.Code | Should -Not -Be 0
            $result.Text | Should -Match 'Package integrity: SHA256 mismatch: LICENSE'
            $result.Text | Should -Not -Match 'Configuration file|Snapshot'
        }
        foreach ($folder in 'data', 'logs', 'reports') { Test-Path -LiteralPath (Join-Path $case $folder) | Should -BeFalse -Because $folder }
    }
    It 'does not let a module bypass the check (<Relative>)' -ForEach @(
        @{ Relative = 'module\PRA.Common.psm1' }, @{ Relative = 'module\PRA.Store.psm1' }, @{ Relative = 'module\PRA.Collect.psm1' },
        @{ Relative = 'module\PRA.Cloud.psm1' }, @{ Relative = 'module\PRA.Gui.psm1' }, @{ Relative = 'PraCloudMailbox.psm1' }, @{ Relative = 'PraCloudMailbox.psd1' }
    ) {
        $case = Copy-Case
        Add-Byte (Join-Path $case 'LICENSE')
        $target = Join-Path $case $Relative
        $result = Invoke-Pwsh -Arguments @('-Command', ("try { Import-Module -Name '$target' -Force -ErrorAction Stop } catch { [Console]::Error.WriteLine(`$_.Exception.Message); exit 1 }"))
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity: SHA256 mismatch: LICENSE'
    }
    It 'is refused by Windows PowerShell 5.1 at #Requires, before the integrity bootstrap and any other code (<Relative>)' -Skip:(-not (Test-Path (Join-Path $env:SystemRoot System32\WindowsPowerShell\v1.0\powershell.exe))) -ForEach @(
        @{ Relative = 'Invoke-PraCloudMailbox.ps1' }, @{ Relative = 'PraCloudMailbox.psm1' }, @{ Relative = 'module\PRA.Common.psm1' }, @{ Relative = 'module\PRA.Store.psm1' },
        @{ Relative = 'module\PRA.Collect.psm1' }, @{ Relative = 'module\PRA.Cloud.psm1' }, @{ Relative = 'module\PRA.Gui.psm1' }, @{ Relative = 'module\Test-PackageIntegrity.ps1' }
    ) {
        $target = Join-Path $script:Fixture $Relative
        $text = [IO.File]::ReadAllText($target)
        $text | Should -Match '(?m)^#Requires -Version 7\.5\r?$'
        $text | Should -Not -Match '(?m)^#Requires -PSEdition'
        $previous = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $command = if ($Relative -like '*.psm1') { "Import-Module -Name '$target' -Force -ErrorAction Stop" } else { "& '$target'" }
            $output = & $script:Ps51 -NoLogo -NoProfile -NonInteractive -Command $command 2>&1 | ForEach-Object { [string]$_ }
            $code = $LASTEXITCODE
        }
        finally { $ErrorActionPreference = $previous }
        $code | Should -Not -Be 0
        ($output -join "`n") | Should -Match '7\.5'
        ($output -join "`n") | Should -Not -Match 'Package integrity'
    }
    It 'imports the modules of an intact fixture package (the bootstrap passes, then the module loads)' {
        $result = Invoke-Pwsh -Arguments @('-Command', "try { Import-Module -Name '$(Join-Path $script:Fixture 'module\PRA.Common.psm1')' -Force -ErrorAction Stop; 'LOADED' } catch { [Console]::Error.WriteLine(`$_.Exception.Message); exit 1 }")
        $result.Text | Should -Match 'LOADED'
        $result.Code | Should -Be 0
    }
}

Describe 'Package builder and manifest tool - controlled fixture, same paths and no private data' {
    BeforeAll {
        # A repository-shaped copy: package\ (fixture, no changelog), CHANGELOG.md and the real builder with its own signature call substituted.
        $script:BuilderRepo = Join-Path $script:Work 'builder-repo'
        New-PcmIntegrityFixture -SourcePackage $script:Source -Destination (Join-Path $script:BuilderRepo 'package') -Changelog $script:Changelog
        Copy-Item -LiteralPath $script:Changelog -Destination $script:BuilderRepo
        $null = New-Item -ItemType Directory -Path (Join-Path $script:BuilderRepo 'tools') -Force
        $builder = Join-Path $script:BuilderRepo 'tools\New-PraCloudPackage.ps1'
        Copy-Item -LiteralPath (Join-Path $script:Repo 'tools\New-PraCloudPackage.ps1') -Destination $builder
        $file = Read-PcmFixtureText $builder
        ([regex]::Matches($file.Text, [regex]::Escape($global:PcmIntegrity.Token))).Count | Should -Be 1
        Write-PcmFixtureText $builder $file.Text.Replace($global:PcmIntegrity.Token, (Get-PcmSignatureSubstitute)) $file.Bom
        $script:Builder = $builder
        Copy-Item -LiteralPath (Join-Path $script:Repo 'tools\Update-PackageIntegrity.ps1') -Destination (Join-Path $script:BuilderRepo 'tools\Update-PackageIntegrity.ps1')
    }
    It 'builds exactly the extracted-runtime layout: inventory, manifest and sanitized config, nothing else' {
        $destination = Join-Path $script:Work 'built'
        $result = Invoke-Pwsh -Arguments @('-File', $script:Builder, '-Destination', $destination)
        $result.Code | Should -Be 0 -Because $result.Text
        $built = @(Get-ChildItem -LiteralPath $destination -Recurse -File | ForEach-Object { $_.FullName.Substring($destination.Length + 1) } | Sort-Object)
        $expected = @(Get-PcmExpectedInventory) + @('PackageIntegrity.psd1', 'config\PraCloudMailbox.config.psd1')
        ($built -join '|') | Should -BeExactly (($expected | Sort-Object) -join '|')
        Test-Path -LiteralPath (Join-Path $destination 'Invoke-PraCloudMailbox.ps1') -PathType Leaf | Should -BeTrue
        foreach ($forbidden in 'tests', 'tools', 'data', 'logs', 'reports', 'package') { Test-Path -LiteralPath (Join-Path $destination $forbidden) | Should -BeFalse }
        (Invoke-Verifier $destination -Distribution).Code | Should -Be 0
        $packaged = Import-PowerShellDataFile -LiteralPath (Join-Path $destination 'config\PraCloudMailbox.config.psd1')
        $packaged.Exchange.Server | Should -BeNullOrEmpty
        $packaged.Cloud.TenantId | Should -BeNullOrEmpty
        $packaged.Resources.OnPremisesReady | Should -BeFalse
        $packaged.Resources.AllowUnprotectedDeletion | Should -BeFalse
    }
    It 'writes the release zip and its .sha256 as one line: the hash, two spaces, the file name' {
        $destination = Join-Path $script:Work 'built-zip\PraCloudMailbox-9.9.9'
        $result = Invoke-Pwsh -Arguments @('-File', $script:Builder, '-Destination', $destination, '-Zip')
        $result.Code | Should -Be 0 -Because $result.Text
        $zip = "$destination.zip"
        Test-Path -LiteralPath $zip -PathType Leaf | Should -BeTrue
        $line = [IO.File]::ReadAllText("$zip.sha256")
        $line | Should -Match '^[0-9a-f]{64}  PraCloudMailbox-9\.9\.9\.zip\n$'
        $line.Substring(0, 64) | Should -BeExactly (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $names = [IO.Compression.ZipFile]::OpenRead($zip)
        try { @($names.Entries | ForEach-Object FullName | Where-Object { $_ -like 'PraCloudMailbox-9.9.9/Invoke-PraCloudMailbox.ps1' }).Count | Should -Be 1 }
        finally { $names.Dispose() }
    }
    It 'refuses a changed source before writing anything' {
        $changed = Join-Path $script:Work 'builder-changed'
        Copy-Item -LiteralPath $script:BuilderRepo -Destination $changed -Recurse
        Add-Byte (Join-Path $changed 'package\LICENSE')
        $destination = Join-Path $script:Work 'built-refused'
        $result = Invoke-Pwsh -Arguments @('-File', (Join-Path $changed 'tools\New-PraCloudPackage.ps1'), '-Destination', $destination)
        $result.Code | Should -Not -Be 0
        $result.Text | Should -Match 'Package integrity:'
        Test-Path -LiteralPath $destination | Should -BeFalse
    }
    It 'writes a DRAFT manifest of the same files and hashes as the fixture manifest, and refuses -Final while a file is unsigned' {
        $copy = Join-Path $script:Work 'manifest-repo'
        Copy-Item -LiteralPath $script:Source -Destination (Join-Path $copy 'package') -Recurse
        $null = New-Item -ItemType Directory -Path (Join-Path $copy 'tools') -Force
        Copy-Item -LiteralPath $script:Changelog -Destination $copy
        Copy-Item -LiteralPath (Join-Path $script:Repo 'tools\Update-PackageIntegrity.ps1') -Destination (Join-Path $copy 'tools\Update-PackageIntegrity.ps1')
        $result = Invoke-Pwsh -Arguments @('-File', (Join-Path $copy 'tools\Update-PackageIntegrity.ps1'))
        $result.Code | Should -Be 0 -Because $result.Text
        $text = [IO.File]::ReadAllText((Join-Path $copy 'package\PackageIntegrity.psd1'))
        $text | Should -Match '# DRAFT FOR \d+\.\d+\.\d+ \(unsigned\)'
        $data = Import-PowerShellDataFile -LiteralPath (Join-Path $copy 'package\PackageIntegrity.psd1')
        $data.SchemaVersion | Should -Be 1
        $data.SignerThumbprint | Should -BeExactly $global:PcmIntegrity.Signer
        (@($data.Files | ForEach-Object { $_.Path } | Sort-Object) -join '|') | Should -BeExactly ((@(Get-PcmExpectedInventory) | Sort-Object) -join '|')
        foreach ($entry in $data.Files) {
            $path = if ($entry.Path -eq 'CHANGELOG.md') { $script:Changelog } else { Join-Path $script:Source $entry.Path }
            $entry.Sha256 | Should -BeExactly (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -Because $entry.Path
        }
        $final = Invoke-Pwsh -Arguments @('-File', (Join-Path $copy 'tools\Update-PackageIntegrity.ps1'), '-Final')
        if ((Get-NativeState).AllValid) { $final.Code | Should -Be 0 -Because $final.Text }
        else {
            $final.Code | Should -Not -Be 0
            $final.Text | Should -Match 'Not signed by the owner'
        }
    }
}

Describe 'Native state of the real package - no substitution (honest, state-dependent)' {
    It 'fails closed before any import or operation while unsigned, and verifies once all the files are signed' {
        $state = Get-NativeState
        $entry = Join-Path $script:Source 'Invoke-PraCloudMailbox.ps1'
        $verifier = Join-Path $script:Source 'module\Test-PackageIntegrity.ps1'
        $module = Invoke-Pwsh -Arguments @('-Command', "try { Import-Module -Name '$(Join-Path $script:Source 'module\PRA.Common.psm1')' -Force -ErrorAction Stop; 'LOADED' } catch { [Console]::Error.WriteLine(`$_.Exception.Message); exit 1 }")
        $direct = Invoke-Pwsh -Arguments @('-File', $verifier, '-Root', $script:Source)
        if ($state.AllValid) {
            $direct.Code | Should -Be 0 -Because $direct.Text
            $module.Code | Should -Be 0 -Because $module.Text
            $module.Text | Should -Match 'LOADED'
            # An intact signed entry could contact Exchange or open the window. Test its native rejection on a tampered copy instead.
            $case = Join-Path $script:Work 'native-signed-tampered'
            Copy-Item -LiteralPath $script:Source -Destination $case -Recurse
            Add-Byte (Join-Path $case 'LICENSE')
            foreach ($arguments in @(@('-Action', 'Check', '-Mode', 'Preview'), @('-Gui'))) {
                $rejected = Invoke-Pwsh -Arguments (@('-File', (Join-Path $case 'Invoke-PraCloudMailbox.ps1')) + $arguments)
                $rejected.Code | Should -Not -Be 0
                $rejected.Text | Should -Match 'Package integrity: SHA256 mismatch: LICENSE'
                $rejected.Text | Should -Not -Match 'Configuration file'
            }
            foreach ($folder in 'data', 'logs', 'reports') { Test-Path -LiteralPath (Join-Path $case $folder) | Should -BeFalse -Because $folder }
        }
        else {
            $cli = Invoke-Pwsh -Arguments @('-File', $entry, '-Action', 'Check', '-Mode', 'Preview')
            $gui = Invoke-Pwsh -Arguments @('-File', $entry, '-Gui')
            foreach ($result in $cli, $gui, $module, $direct) {
                $result.Code | Should -Not -Be 0
                $result.Text | Should -Match 'Package integrity'
            }
            $module.Text | Should -Not -Match 'LOADED'
            foreach ($folder in 'data', 'logs', 'reports') { Test-Path -LiteralPath (Join-Path $script:Source $folder) | Should -BeFalse -Because $folder }
        }
    }
    It 'refuses to build a package from an unsigned source and writes nothing' {
        $state = Get-NativeState
        $destination = Join-Path $script:Work 'native-built'
        $result = Invoke-Pwsh -Arguments @('-File', (Join-Path $script:Repo 'tools\New-PraCloudPackage.ps1'), '-Destination', $destination)
        if ($state.AllValid) { $result.Code | Should -Be 0 -Because $result.Text }
        else {
            $result.Code | Should -Not -Be 0
            $result.Text | Should -Match 'Package integrity'
            Test-Path -LiteralPath $destination | Should -BeFalse
        }
    }
}
