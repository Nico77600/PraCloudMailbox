<#
.SYNOPSIS
    PRA Cloud Mailbox and PRA Remote Mailbox in the same PowerShell session: no command of one replaces a command of the other.
.DESCRIPTION
    Every command of this tool has the noun prefix Pcm; PRA Remote Mailbox uses Pra. The static test compares the commands that this tool defines and exports with
    the frozen list of the commands of PRA Remote Mailbox 3.1.3 (tests\fixtures\PraRemoteMailbox-3.1.3-commands.json). The live test imports both manifests in one
    session and checks that every command still comes from its own module; it runs when a package folder of PRA Remote Mailbox is given in the environment
    variable PRA_REMOTE_MAILBOX_PACKAGE: the folder that holds PraRemoteMailbox.psd1 and passes its own integrity check (a signed release, or the copy that the
    integrity fixture of its test suite makes).
.NOTES
    Author  : Nicolas Fabert
    Version : 2.0.0
#>
#Requires -Version 7.5

BeforeAll {
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Source = Join-Path $script:Repo 'package'
    $script:Remote = Get-Content -LiteralPath (Join-Path $PSScriptRoot 'fixtures\PraRemoteMailbox-3.1.3-commands.json') -Raw | ConvertFrom-Json
    $script:RemoteDefined = New-Object 'Collections.Generic.HashSet[string]' ([string[]]@($script:Remote.Defined), [StringComparer]::OrdinalIgnoreCase)
    $script:RemoteExported = New-Object 'Collections.Generic.HashSet[string]' ([string[]]@($script:Remote.Exported), [StringComparer]::OrdinalIgnoreCase)
    $script:Defined = @{}
    # The verifier runs as a script in a scope of its own: its helper functions never reach the session (PRA Remote Mailbox has the same ones).
    foreach ($file in Get-ChildItem -LiteralPath $script:Source -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' -and $_.Name -ne 'Test-PackageIntegrity.ps1' }) {
        $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
        foreach ($function in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $script:Defined[$function.Name] = $file.Name }
    }
    $script:Exported = New-Object 'Collections.Generic.List[string]'
    foreach ($relative in 'module\PRA.Common.psm1', 'module\PRA.Store.psm1', 'module\PRA.Collect.psm1', 'module\PRA.Cloud.psm1', 'module\PRA.Gui.psm1') {
        $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Source $relative), [ref]$null, [ref]$null)
        $export = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -ceq 'Export-ModuleMember' }, $true)
        foreach ($name in @($export.CommandElements | Select-Object -Skip 1 | ForEach-Object { $_.FindAll({ param($n) $n -is [Management.Automation.Language.StringConstantExpressionAst] }, $true) } | ForEach-Object Value | Where-Object { $_ -cne 'Function' })) { $script:Exported.Add($name) }
    }
}

Describe 'No command of PRA Cloud Mailbox can replace a command of PRA Remote Mailbox' {
    It 'defines no function with a name that PRA Remote Mailbox defines' {
        $collisions = @($script:Defined.Keys | Where-Object { $script:RemoteDefined.Contains($_) } | Sort-Object)
        ($collisions -join ', ') | Should -BeExactly ''
    }
    It 'exports no command with a name that PRA Remote Mailbox exports or defines' {
        $collisions = @($script:Exported | Where-Object { $script:RemoteExported.Contains($_) -or $script:RemoteDefined.Contains($_) } | Sort-Object)
        ($collisions -join ', ') | Should -BeExactly ''
    }
    It 'gives every function of the modules the noun prefix Pcm, so that a future function of PRA Remote Mailbox cannot collide either' {
        $other = @($script:Defined.Keys | Where-Object { $_ -notmatch '^[A-Z][a-z]+-Pcm[A-Z0-9]' } | Sort-Object)
        ($other -join ', ') | Should -BeExactly ''
    }
    It 'keeps the module files of the same name apart: they are loaded by path and never removed by name' {
        foreach ($file in Get-ChildItem -LiteralPath $script:Source -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1' }) {
            [IO.File]::ReadAllText($file.FullName) | Should -Not -Match 'Remove-Module\s+(-Name\s+)?[''"]?PRA\.' -Because $file.Name
        }
    }
}

Describe 'Both tools in one session' {
    BeforeAll {
        $script:RemotePackage = [Environment]::GetEnvironmentVariable('PRA_REMOTE_MAILBOX_PACKAGE')
        $script:RemoteAvailable = [bool]$script:RemotePackage -and (Test-Path -LiteralPath (Join-Path $script:RemotePackage 'PraRemoteMailbox.psd1'))
        $script:Package = New-PcmTestPackage -Repository $script:Repo
    }
    AfterAll { Remove-PcmTestPackage $script:Package }
    It 'imports the manifests of both tools, in either order, and every command still comes from its own module' -Skip:(-not [Environment]::GetEnvironmentVariable('PRA_REMOTE_MAILBOX_PACKAGE')) {
        foreach ($order in 'remote-first', 'cloud-first') {
            $script = @"
`$ErrorActionPreference = 'Stop'
`$ours = '$($script:Package -replace "'", "''")\PraCloudMailbox.psd1'
`$theirs = '$($script:RemotePackage -replace "'", "''")\PraRemoteMailbox.psd1'
`$paths = if ('$order' -eq 'remote-first') { @(`$theirs, `$ours) } else { @(`$ours, `$theirs) }
foreach (`$path in `$paths) { Import-Module -Name `$path -Force -ErrorAction Stop }
foreach (`$path in `$paths) { Import-Module -Name `$path -Force -ErrorAction Stop }
`$wrong = @()
foreach (`$name in (Get-Module PraCloudMailbox).ExportedFunctions.Keys) { if ((Get-Command `$name).Source -ne 'PraCloudMailbox') { `$wrong += `$name } }
foreach (`$name in (Get-Module PraRemoteMailbox).ExportedFunctions.Keys) { if ((Get-Command `$name).Source -ne 'PraRemoteMailbox') { `$wrong += `$name } }
'WRONG=' + (`$wrong -join ',')
'OURS=' + @((Get-Module PraCloudMailbox).ExportedFunctions.Keys).Count
'THEIRS=' + @((Get-Module PraRemoteMailbox).ExportedFunctions.Keys).Count
"@
            $file = Join-Path ([IO.Path]::GetTempPath()) ('coexist-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.ps1')
            [IO.File]::WriteAllText($file, $script, [Text.UTF8Encoding]::new($true))
            try { $output = (& (Join-Path $PSHOME 'pwsh.exe') -NoLogo -NoProfile -NonInteractive -File $file 2>&1 | Out-String) } finally { Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue }
            $output | Should -Match '(?m)^WRONG=\s*$' -Because $output
            [int][regex]::Match($output, '(?m)^OURS=(\d+)').Groups[1].Value | Should -BeGreaterThan 50 -Because $output
            [int][regex]::Match($output, '(?m)^THEIRS=(\d+)').Groups[1].Value | Should -BeGreaterThan 50 -Because $output
        }
    }
}

Describe 'The time limit of the calls leaves the session as it found it' {
    BeforeAll {
        $script:LimitPackage = New-PcmTestPackage -Repository $script:Repo
        foreach ($name in 'PRA.Common', 'PRA.Store', 'PRA.Cloud') { Import-Module (Join-Path $script:LimitPackage "module\$name.psm1") -Force -Global }
        # Stand-ins for the Exchange Online command of this tool (no prefix) and for the same command as PRA Remote Mailbox loads it (prefix Pra).
        function global:Get-Mailbox { [CmdletBinding()] param([string]$Identity) "plain:$Identity" }
        function global:Get-PraMailbox { [CmdletBinding()] param([string]$Identity) "pra:$Identity" }
    }
    AfterAll {
        foreach ($name in 'Get-Mailbox', 'Get-PraMailbox') { Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue }
        Remove-Item -Path 'Alias:Get-Mailbox' -ErrorAction SilentlyContinue
        Remove-PcmTestPackage $script:LimitPackage
    }
    BeforeEach { $script:LimitContext = @{ Config = @{ Polling = @{ CallTimeoutMinutes = 5 } }; CurrentPhase = 'Objects'; Action = 'Check' } }
    AfterEach { Stop-PcmCallGuard; Remove-Item -Path 'Alias:Get-Mailbox' -ErrorAction SilentlyContinue }

    It 'never touches the commands with the prefix of PRA Remote Mailbox: they answer as before during our run' {
        $names = & (Get-Module PRA.Cloud) { Get-PcmGuardedCommandName }
        @($names | Where-Object { $_ -match '^[A-Za-z]+-Pra' }).Count | Should -Be 0
        $names | Should -Contain 'Get-Mailbox'
        Start-PcmCallGuard -Context $script:LimitContext
        Install-PcmCallGuard -Name @($names | Where-Object { $_ -in 'Get-Mailbox', 'Get-PraMailbox' })
        (Get-Command Get-Mailbox).CommandType | Should -Be 'Alias'
        (Get-Command Get-PraMailbox).CommandType | Should -Be 'Function'
        Get-Command 'PcmGuard_Get-PraMailbox' -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
        Get-PraMailbox 'x' | Should -Be 'pra:x'
        Get-Mailbox 'x' | Should -Be 'plain:x'
    }
    It 'leaves no alias and no proxy command behind once the run is over, and the commands answer directly again' {
        Start-PcmCallGuard -Context $script:LimitContext
        Install-PcmCallGuard -Name 'Get-Mailbox'
        Stop-PcmCallGuard
        (Get-Command Get-Mailbox).CommandType | Should -Be 'Function'
        @(Get-Alias -Name 'Get-Mailbox' -ErrorAction SilentlyContinue).Count | Should -Be 0
        @(Get-Command -Name 'PcmGuard_*' -ErrorAction SilentlyContinue).Count | Should -Be 0
        @(Get-Alias -ErrorAction SilentlyContinue | Where-Object { [string]$_.Definition -like 'PcmGuard_*' }).Count | Should -Be 0
        Get-Mailbox 'y' | Should -Be 'plain:y'
        Get-PraMailbox 'y' | Should -Be 'pra:y'
    }
    It 'puts back an alias of the same name that another tool had made, and removes only its own' {
        Set-Alias -Name Get-Mailbox -Value Get-PraMailbox -Scope Global -Force
        Start-PcmCallGuard -Context $script:LimitContext
        Install-PcmCallGuard -Name 'Get-Mailbox'
        (Get-Alias -Name Get-Mailbox).Definition | Should -Be 'PcmGuard_Get-Mailbox'
        Stop-PcmCallGuard
        (Get-Alias -Name Get-Mailbox).Definition | Should -Be 'Get-PraMailbox'
        Get-Mailbox 'z' | Should -Be 'pra:z'
        @(Get-Command -Name 'PcmGuard_*' -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
    It 'does not remove an alias that another tool made after ours' {
        Start-PcmCallGuard -Context $script:LimitContext
        Install-PcmCallGuard -Name 'Get-Mailbox'
        Set-Alias -Name Get-Mailbox -Value Get-PraMailbox -Scope Global -Force
        Stop-PcmCallGuard
        (Get-Alias -Name Get-Mailbox).Definition | Should -Be 'Get-PraMailbox'
    }
    It 'starts clean again after a run that ended: two runs one after the other leave nothing' {
        foreach ($time in 1, 2) {
            Start-PcmCallGuard -Context $script:LimitContext
            Install-PcmCallGuard -Name 'Get-Mailbox'
            Get-Mailbox 'a' | Should -Be 'plain:a'
            Stop-PcmCallGuard
            @(Get-Alias -Name 'Get-Mailbox' -ErrorAction SilentlyContinue).Count | Should -Be 0
            @(Get-Command -Name 'PcmGuard_*' -ErrorAction SilentlyContinue).Count | Should -Be 0
        }
    }
}
