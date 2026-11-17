<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the guided console (Pester 5+). No Exchange and no cloud: the state is a synthetic configuration, snapshot and journal
    (tools\DocumentationImages\ConsoleFixture.ps1), the engine is the synthetic engine (tools\DocumentationImages\StubEngine.ps1), the probes are mocked.
.DESCRIPTION
    The pure functions are tested in this process; the real window is built and driven in child processes (pwsh -STA), one per scenario, which print their
    observations as one JSON line.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    . (Join-Path $script:Repo 'tools\DocumentationImages\New-ToolCopy.ps1')
    . (Join-Path $script:Repo 'tools\DocumentationImages\ConsoleFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    Import-Module (Join-Path $script:Root 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA.Store.psm1') -Force
    $script:Gui = Import-Module (Join-Path $script:Root 'module\PRA.Gui.psm1') -Force -PassThru
    # The SQLite library stays loaded, and so locked, until the process ends, and it comes from the first folder that opens a store: it is the one of the temporary package
    # (removed without error at the end), never the copy under TestDrive, which Pester must be able to delete.
    & (Get-Module PRA.Store) { Import-PcmSqlite -Root $args[0] } $script:Root
    # One synthetic tool folder for the pure-function tests (a snapshot, a journal, a group captured by Collect).
    $script:Tool = New-DocumentationToolCopy -Repository $script:Repo -Destination (Join-Path $TestDrive 'tool')
    $script:Fixture = New-ConsoleFixture -ToolRoot $script:Tool
    $script:State = Read-PcmGuiState -Root $script:Tool -ConfigPath $script:Fixture.Config

    # The prelude of the child processes: the tool copy, the fixture, the real window with mocked probes, and the helpers to drive it.
    $script:Prelude = @'
param([string]$Repo, [string]$Work, [string]$Kind = 'Cloud')
$ErrorActionPreference = 'Stop'
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
. (Join-Path $Repo 'tools\DocumentationImages\New-ToolCopy.ps1'); . (Join-Path $Repo 'tools\DocumentationImages\ConsoleFixture.ps1')
$root = New-DocumentationToolCopy -Repository $Repo -Destination $Work
Import-Module (Join-Path $root 'module\PRA.Common.psm1') -Force; Import-Module (Join-Path $root 'module\PRA.Store.psm1') -Force
$gm = Import-Module (Join-Path $root 'module\PRA.Gui.psm1') -Force -PassThru
$fixture = New-ConsoleFixture -ToolRoot $root
function Pump($ms = 80) { $f = New-Object Windows.Threading.DispatcherFrame; $t = New-Object Windows.Threading.DispatcherTimer; $t.Interval = [timespan]::FromMilliseconds($ms); $t.Tag = $f; $t.Add_Tick({ param($s) $s.Tag.Continue = $false; $s.Stop() }); $t.Start(); [Windows.Threading.Dispatcher]::PushFrame($f) }
function Click($n) { $g.Controls[$n].RaiseEvent((New-Object Windows.RoutedEventArgs ([Windows.Controls.Button]::ClickEvent))); Pump 30 }
function Check($n) { $g.Controls[$n].IsChecked = $true; Pump 30 }
function Enter($n, $text) { $g.Controls[$n].Text = $text; Pump 30 }
function WaitRun { $d = [datetime]::UtcNow.AddSeconds(90); while ($g.Run -and [datetime]::UtcNow -lt $d) { Pump 60 }; if ($g.Run) { throw 'the run did not end' }; Pump 100 }
function WaitFor([scriptblock]$Condition, [int]$Seconds = 45) { $d = [datetime]::UtcNow.AddSeconds($Seconds); while (-not (& $Condition) -and [datetime]::UtcNow -lt $d) { Pump 80 }; if (-not (& $Condition)) { throw 'condition not reached' } }
function Hold($flag) { [IO.File]::WriteAllText((Join-Path $root "hold.$flag"), '1') }
function Release($flag) { Remove-Item (Join-Path $root "hold.$flag") -ErrorAction SilentlyContinue }
function Scenario($holdAfter) { $s = [IO.File]::ReadAllText((Join-Path $root 'scenario.json')) | ConvertFrom-Json; $s.holdAfter = $holdAfter; [IO.File]::WriteAllText((Join-Path $root 'scenario.json'), ($s | ConvertTo-Json -Depth 6)) }
function Ask { $s = [IO.File]::ReadAllText((Join-Path $root 'scenario.json')) | ConvertFrom-Json; $s | Add-Member -NotePropertyName ask -NotePropertyValue $true -Force; [IO.File]::WriteAllText((Join-Path $root 'scenario.json'), ($s | ConvertTo-Json -Depth 6)) }
function Args_ { [IO.File]::ReadAllText((Join-Path $root 'last-args.json')) | ConvertFrom-Json }
function Mod([scriptblock]$Script) { & $gm $Script }
$cloudLights = @([ordered]@{ Name = 'PowerShell 7.5'; Status = 'Ok'; Detail = 'ok' }, [ordered]@{ Name = 'Cloud modules'; Status = 'Ok'; Detail = 'ok' }, [ordered]@{ Name = 'Sign-in: certificate'; Status = 'Ok'; Detail = 'ok' })
$cloudProbe = { param($r) [pscustomobject]@{ CanCollect = $true; CanCloudEngine = $true; CanCloudBase = $true; ExchangeServer = 'ex01.contoso.com'; Net = ''; Role = 'None'; Lights = $cloudLights } }
$collectProbe = { param($r) [pscustomobject]@{ CanCollect = $true; CanCloudEngine = $true; CanCloudBase = $false; ExchangeServer = 'ex01.contoso.com'; Net = ''; Role = 'None'; Lights = @(
            [ordered]@{ Name = 'PowerShell 7.5'; Status = 'Ok'; Detail = 'ok' }, [ordered]@{ Name = 'Cloud modules'; Status = 'Fail'; Detail = 'missing' }) } }
$exchangeAnswers = { param($server, $seconds) [pscustomobject]@{ IsCompleted = $true; Result = [pscustomobject]@{ Ok = $true; Why = '' } } }
$exchangeSilent = { param($server, $seconds) [pscustomobject]@{ IsCompleted = $true; Result = [pscustomobject]@{ Ok = $false; Why = 'The server did not answer.' } } }
$questions = New-Object Collections.Generic.List[string]
$g = New-PcmGuiWindow -Root $root -ConfigPath $fixture.Config -Version 'test' -Theme Light -PrerequisiteProbe $(if ($Kind -eq 'Collect') { $collectProbe } else { $cloudProbe }) -ExchangeProbe $(if ($Kind -eq 'ExchangeDown') { $exchangeSilent } else { $exchangeAnswers }) -Hooks @{ NoConsole = $true; Question = { param($t, $x) $questions.Add($x); $true }; PickOu = { param($units) @($units)[0].Name } }
$g.Form.WindowStartupLocation = 'Manual'; $g.Form.Left = -3000; $g.Form.Top = 0; $g.Form.ShowActivated = $false
$g.Form.Show(); Pump 200
$c = $g.Controls
$configHash = (Get-FileHash -LiteralPath $fixture.Config -Algorithm SHA256).Hash
function Done($result) { Write-Output ('RESULT:' + ($result | ConvertTo-Json -Depth 8 -Compress)); $g.Form.Dispatcher.InvokeShutdown() 2>$null }
'@
    function script:Invoke-GuiChild {
        <# Runs a scenario in its own pwsh -STA process (the prelude first) and returns the observations it printed. #>
        param([Parameter(Mandatory)][string]$Body, [ValidateSet('Cloud', 'Collect', 'ExchangeDown')][string]$Kind = 'Cloud')
        $file = Join-Path $TestDrive ('child-{0}.ps1' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
        [IO.File]::WriteAllText($file, ($script:Prelude + "`r`n" + $Body), (New-Object Text.UTF8Encoding($true)))
        $work = Join-Path $TestDrive ('work-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $text = (& pwsh -NoProfile -STA -File $file -Repo $script:Repo -Work $work -Kind $Kind 2>&1 | Out-String)
        $line = @($text -split "`r?`n" | Where-Object { $_.StartsWith('RESULT:') }) | Select-Object -Last 1
        if (-not $line) { throw "The child process gave no result:`n$text" }
        return ($line.Substring(7) | ConvertFrom-Json)
    }
}

AfterAll {
    Remove-PcmTestPackage $script:Root
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Console: the request and the command line of a run' {
    It 'puts the official spelling of every value, whatever the letter case typed' {
        $request = New-PcmGuiRequest -Action 'convert' -Mode 'apply' -Scope 'usersonly' -Expect ''
        $request.Action | Should -BeExactly 'Convert'
        $request.Mode | Should -BeExactly 'Apply'
        $request.Scope | Should -BeExactly 'UsersOnly'
        (New-PcmGuiRequest -Action 'check' -Batch 'ABCD1234' -Expect 'DEPROVISIONED').Expect | Should -BeExactly 'Deprovisioned'
    }
    It 'refuses what the engine would refuse, in plain words' {
        { New-PcmGuiRequest -Action Convert -Identity 'a@contoso.com' -Group 'PRA-Wave1' } | Should -Throw '*one way to pick the objects*'
        { New-PcmGuiRequest -Action Convert -CsvPath 'a.csv' -SearchBase 'contoso.com/Paris' } | Should -Throw '*one way to pick the objects*'
        { New-PcmGuiRequest -Action Recover } | Should -Throw '*Pick the batch*'
        { New-PcmGuiRequest -Action Check -Mode Apply } | Should -Throw '*only reads*'
        { New-PcmGuiRequest -Action Convert -Expect Provisioned } | Should -Throw '*follows a batch*'
        { New-PcmGuiRequest -Action Collect -Group 'x' } | Should -Throw '*Collect reads Exchange*'
        { New-PcmGuiRequest -Action Convert -Identity "a`r`nb" } | Should -Throw '*control character*'
        { New-PcmGuiRequest -Action Convert -MaxObjects -1 } | Should -Throw '*zero*'
        { New-PcmGuiRequest -Action Convert -Group 'g' -ResolveOnly $true } | Should -Throw '*only for a check of a group*'
        { New-PcmGuiRequest -Action Check -Batch '1a2b3c4d' -Group 'g' -ResolveOnly $true } | Should -Throw '*only for a check of a group*'
        { New-PcmGuiRequest -Action Convert -Scope Everything } | Should -Throw
    }
    It 'maps every source and filter to the parameters of the engine' -TestCases @(
        @{ Name = 'one mailbox'; Request = @{ Action = 'Convert'; Identity = 'adele.vance@contoso.com' }; Expected = '-Identity adele.vance@contoso.com' }
        @{ Name = 'a group'; Request = @{ Action = 'Convert'; Group = 'PRA-Wave1' }; Expected = '-Group PRA-Wave1' }
        @{ Name = 'an OU'; Request = @{ Action = 'Convert'; SearchBase = 'contoso.com/Paris' }; Expected = '-SearchBase contoso.com/Paris' }
        @{ Name = 'a CSV file (wave)'; Request = @{ Action = 'Recover'; Batch = 'abcd1234'; CsvPath = 'C:\PRA\logs\gui\wave-1.csv' }; Expected = '-Batch abcd1234 -CsvPath C:\PRA\logs\gui\wave-1.csv' }
        @{ Name = 'a scope and a limit'; Request = @{ Action = 'Convert'; Scope = 'SharedOnly'; MaxObjects = 200 }; Expected = '-Scope SharedOnly -MaxObjects 200' }
        @{ Name = 'a snapshot'; Request = @{ Action = 'Check'; Snapshot = 40 }; Expected = '-Snapshot 40' }
        @{ Name = 'a group lookup'; Request = @{ Action = 'Check'; Group = 'PRA-Wave1'; ResolveOnly = $true }; Expected = '-Group PRA-Wave1 -ResolveOnly' }
        @{ Name = 'a follow-up'; Request = @{ Action = 'Check'; Batch = '1a2b3c4d'; Expect = 'retained'; Once = $true }; Expected = '-Batch 1a2b3c4d -Expect Retained -Once' }) {
        param($Name, $Request, $Expected)
        $command = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\config\PraCloudMailbox.config.psd1' -Request (New-PcmGuiRequest @Request) -Engine 'pwsh.exe'
        ($command.Tokens -join ' ') | Should -BeLike "*$Expected*"
        $command.Display | Should -BeLike "*$Expected*"
    }
    It 'never uses -NonInteractive (the interactive sign-in needs its window), runs the script of the tool folder and adds -Force only for an Apply' {
        $preview = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\config\PraCloudMailbox.config.psd1' -Request (New-PcmGuiRequest -Action Convert -Group 'g') -Engine 'pwsh.exe'
        $apply = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\config\PraCloudMailbox.config.psd1' -Request (New-PcmGuiRequest -Action Convert -Mode Apply -Group 'g') -Engine 'pwsh.exe'
        $preview.Arguments | Should -Not -Match 'NonInteractive'
        $preview.Arguments | Should -BeLike '-NoProfile -STA -File "C:\PRA\Invoke-PraCloudMailbox.ps1" -Action Convert -Mode Preview -Group g -ConfigPath*'
        $preview.Tokens | Should -Not -Contain '-Force'
        $apply.Tokens | Should -Contain '-Force'
        $apply.Display | Should -BeLike '*-Mode Apply*-Force'
        $preview.NeedsConsoleWindow | Should -BeTrue
    }
    It 'runs Collect like the cloud actions, in PowerShell 7 with a console window of its own, and a check without -Mode' {
        $collect = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\c.psd1' -Request (New-PcmGuiRequest -Action Collect -Mode Apply) -Engine 'pwsh.exe'
        $collect.NeedsConsoleWindow | Should -BeTrue
        $collect.Arguments | Should -Not -Match 'NonInteractive'
        $collect.Display | Should -BeLike '.\Invoke-PraCloudMailbox.ps1 -Action Collect -Mode Apply*-Force'
        (Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\c.psd1' -Request (New-PcmGuiRequest -Action Check) -Engine 'pwsh.exe').Tokens | Should -Not -Contain '-Mode'
    }
    It 'schedules the daily Collect like any action: the official spelling, no objects, -Mode Apply -Force and never -NonInteractive' {
        $request = New-PcmGuiRequest -Action 'sCHEdule' -Mode 'apply'
        $request.Action | Should -BeExactly 'Schedule'
        $command = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\config\PraCloudMailbox.config.psd1' -Request $request -Engine 'pwsh.exe'
        $command.Display | Should -BeExactly '.\Invoke-PraCloudMailbox.ps1 -Action Schedule -Mode Apply -Force'
        $command.Arguments | Should -BeLike '-NoProfile -STA -File "C:\PRA\Invoke-PraCloudMailbox.ps1" -Action Schedule -Mode Apply -ConfigPath*-Force'
        $command.Arguments | Should -Not -Match 'NonInteractive'
        $command.NeedsConsoleWindow | Should -BeTrue
        { New-PcmGuiRequest -Action Schedule -Group 'g' } | Should -Throw '*takes no mailbox, batch, OU, group, CSV file or snapshot*'
        { New-PcmGuiRequest -Action Schedule -Once $true } | Should -Throw '*takes no check option*'
        (& $script:Gui { Get-PcmGuiPlanTitle -Action Schedule }) | Should -Be @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task')
    }
    It 'passes the generated configuration and names the original one for the next steps and the report' {
        $command = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\logs\gui\run-1\effective.config.psd1' -BaseConfigPath 'D:\PRA\lab.config.psd1' -Request (New-PcmGuiRequest -Action Check) -Engine 'pwsh.exe'
        $command.Tokens | Should -Contain '-BaseConfigPath'
        $command.Arguments | Should -BeLike '*-ConfigPath "C:\PRA\logs\gui\run-1\effective.config.psd1" -BaseConfigPath "D:\PRA\lab.config.psd1"'
        $command.Display | Should -BeLike '*-ConfigPath D:\PRA\lab.config.psd1'
        $command.Display | Should -Not -BeLike '*effective*'
    }
    It 'quotes a value with spaces or quotes for the process and shows it as an operator would type it' {
        $command = Get-PcmGuiCommand -Root 'C:\Program Files\PRA' -ConfigPath 'C:\Program Files\PRA\config\c.psd1' -Request (New-PcmGuiRequest -Action Convert -SearchBase 'contoso.com/Paris "North"') -Engine 'pwsh.exe'
        $command.Arguments | Should -BeLike '*-File "C:\Program Files\PRA\Invoke-PraCloudMailbox.ps1"*-SearchBase "contoso.com/Paris \"North\""*'
        ConvertTo-PcmGuiNativeArgument 'C:\dir\' | Should -BeExactly '"C:\dir\\"'
        { ConvertTo-PcmGuiNativeArgument ('a' + [char]0 + 'b') } | Should -Throw
    }
    Context 'The command that is shown and copied holds no expansion' {
        BeforeAll {
            function script:Get-ShownParse {
                param([string]$Command)
                $tokens = $null; $errors = $null
                $ast = [System.Management.Automation.Language.Parser]::ParseInput($Command, [ref]$tokens, [ref]$errors)
                $expanding = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.SubExpressionAst] -or $node -is [System.Management.Automation.Language.VariableExpressionAst] -or
                            $node -is [System.Management.Automation.Language.ExpandableStringExpressionAst] -or $node -is [System.Management.Automation.Language.ArrayExpressionAst] -or
                            $node -is [System.Management.Automation.Language.ParenExpressionAst] -or $node -is [System.Management.Automation.Language.ScriptBlockExpressionAst] }, $true))
                [pscustomobject]@{ Ast = $ast; Errors = @($errors); Expanding = $expanding; Commands = @($ast.FindAll({ param($node) $node -is [System.Management.Automation.Language.CommandAst] }, $true)) }
            }
        }
        It 'shows a plain value bare, and any other value as a single-quoted literal that reads back as the same text (<Value>)' -Tag 'ShownCommand' -TestCases @(
            @{ Value = 'alice@contoso.com'; Bare = $true }, @{ Value = 'D:\PRA\lab.config.psd1'; Bare = $true }, @{ Value = 'contoso.com/Paris'; Bare = $true }
            @{ Value = 'Sales $(Get-Date) team'; Bare = $false }, @{ Value = '$env:USERNAME'; Bare = $false }, @{ Value = 'ends with a backtick`'; Bare = $false }, @{ Value = "it's"; Bare = $false }
            @{ Value = 'a "b" ${c} @(1) ; & calc'; Bare = $false }, @{ Value = "smart $([char]0x2019) and $([char]0x2018) quotes"; Bare = $false }, @{ Value = ''; Bare = $false }
            @{ Value = '@splat'; Bare = $false }, @{ Value = '-Force'; Bare = $false }, @{ Value = 'a b'; Bare = $false }, @{ Value = '$1'; Bare = $false }
            @{ Value = "abc`n"; Bare = $false }, @{ Value = "abc`r`n"; Bare = $false }, @{ Value = "abc`n--"; Bare = $false }) {
            param($Value, $Bare)
            $shown = & $script:Gui { param($v) ConvertTo-PcmGuiShownArgument $v } $Value
            $parse = Get-ShownParse "Get-X -Group $shown"
            $parse.Errors.Count | Should -Be 0 -Because $shown
            $parse.Expanding.Count | Should -Be 0 -Because $shown
            $element = $parse.Commands[0].CommandElements[2]
            $element | Should -BeOfType ([System.Management.Automation.Language.StringConstantExpressionAst])
            $element.Value | Should -BeExactly $Value
            if ($Bare) { $shown | Should -BeExactly $Value } else { $shown | Should -BeLike "'*'" }
        }
        It 'reads the events of the engine with a date-like text and the time of an event as text' -Tag 'ShownCommand' {
            $path = Join-Path $TestDrive 'events.jsonl'
            [IO.File]::WriteAllText($path, ('{"kind":"item","status":"Ok","time":"2026-11-16T10:20:30Z","text":"2025-01-31T00:00:00Z","identity":"a@contoso.com"}' + "`n"), [Text.UTF8Encoding]::new($false))
            $read = Read-PcmGuiEvents -Path $path
            @($read.Events).Count | Should -Be 1
            $read.Events[0].text | Should -BeOfType [string]
            $read.Events[0].text | Should -BeExactly '2025-01-31T00:00:00Z'
            $read.Events[0].time | Should -BeOfType [string]
            $read.Events[0].time | Should -BeExactly '2026-11-16T10:20:30Z'
        }
        It 'shows every value of the command of a run as a literal: Continue, Copy, the next step and the Started line all take this text' -Tag 'ShownCommand' {
            $request = New-PcmGuiRequest -Action Convert -Mode Apply -Group 'Sales $(Get-Date) team' -Scope SharedOnly
            $command = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\config\PraCloudMailbox.config.psd1' -Request $request -Engine 'pwsh.exe'
            $command.Display | Should -BeExactly ".\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Scope SharedOnly -Group 'Sales `$(Get-Date) team' -Force"
            $parse = Get-ShownParse $command.Display
            $parse.Errors.Count | Should -Be 0
            $parse.Expanding.Count | Should -Be 0
            $command.Tokens | Should -Contain 'Sales $(Get-Date) team'
            $continue = Get-PcmGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\c.psd1' -Request (New-PcmGuiRequest -Action Recover -Batch 'c0c0c0c0' -CsvPath 'C:\wave $env:USERNAME`.csv') -Engine 'pwsh.exe'
            $continue.Display | Should -Match ([regex]::Escape("-CsvPath 'C:\wave `$env:USERNAME``.csv'"))
            (Get-ShownParse $continue.Display).Expanding.Count | Should -Be 0
            $next = & $script:Gui { $old = $script:Gui; $script:Gui = [pscustomobject]@{ Root = 'C:\PRA'; ConfigPath = 'C:\PRA\config\PraCloudMailbox.config.psd1' }; try { Get-PcmGuiNextCommand -Result ([pscustomobject]@{ nextSteps = @() }) -Action Recover -BatchId 'c0c0c0c0' -Mode Preview } finally { $script:Gui = $old } }
            $next | Should -BeExactly '.\Invoke-PraCloudMailbox.ps1 -Action Recover -Mode Preview -Batch c0c0c0c0'
        }
    }
}

Describe 'Console: the typed confirmation' {
    It 'accepts the name of the action in any letter case, spaces around it ignored (<Typed>)' -TestCases @(@{ Typed = 'convert' }, @{ Typed = ' Convert ' }, @{ Typed = 'CONVERT' }, @{ Typed = "`tConVert`r`n" }) {
        param($Typed)
        Test-PcmGuiConfirmation -Typed $Typed -Word 'CONVERT' | Should -BeTrue
    }
    It 'refuses another word and any invisible character (<Name>)' -TestCases @(
        @{ Name = 'other word'; Typed = 'recover' }, @{ Name = 'empty'; Typed = '' }, @{ Name = 'zero-width space'; Typed = ('conv' + [char]0x200B + 'ert') }
        @{ Name = 'zero-width at the end'; Typed = ('Convert' + [char]0x200B) }, @{ Name = 'control character'; Typed = ('Convert' + [char]1) }, @{ Name = 'soft hyphen'; Typed = ('Conv' + [char]0xAD + 'ert') }) {
        param($Name, $Typed)
        Test-PcmGuiConfirmation -Typed $Typed -Word 'CONVERT' | Should -BeFalse
    }
}

Describe 'Console: the events of the engine and the waves' {
    It 'reads the events incrementally, complete lines only, and keeps a line that is not JSON as text' {
        $path = Join-Path $TestDrive 'events.jsonl'
        [IO.File]::WriteAllText($path, '{"kind":"step","index":1}' + "`n" + 'not json' + "`n" + '{"kind":"item","te')
        $first = Read-PcmGuiEvents -Path $path
        @($first.Events).Count | Should -Be 2
        $first.Events[0].kind | Should -Be 'step'
        $first.Events[1].text | Should -Be 'not json'
        $first.HasPartial | Should -BeTrue
        [IO.File]::AppendAllText($path, 'xt":"done"}' + "`n")
        $second = Read-PcmGuiEvents -Path $path -Position $first.Position
        @($second.Events).Count | Should -Be 1
        $second.Events[0].text | Should -Be 'done'
        @((Read-PcmGuiEvents -Path $path -Position $second.Position).Events).Count | Should -Be 0
        @((Read-PcmGuiEvents -Path (Join-Path $TestDrive 'missing.jsonl')).Events).Count | Should -Be 0
    }
    It 'writes the objects ticked for a wave as the CSV given to -CsvPath: Identity (the object ID) and Name' {
        $path = Write-PcmGuiWave -Folder (Join-Path $TestDrive 'waves') -Name 'abcd1234' -Rows @([pscustomobject]@{ Guid = 'AAA'; Name = 'Adele "AV" Vance' }, [pscustomobject]@{ Guid = 'bbb'; Name = 'Alex' })
        $rows = @(Import-Csv -LiteralPath $path)
        $rows.Count | Should -Be 2
        $rows[0].Identity | Should -Be 'AAA'
        $rows[0].Name | Should -Be "Adele 'AV' Vance"
        (Get-Content -LiteralPath $path -TotalCount 1) | Should -Match 'Identity,Name'
    }
    It 'keys a selection by its objects, whatever their order and letter case' {
        $one = Get-PcmGuiSelectionKey -Prefix 'abcd1234' -Guids @('B', 'a')
        $one | Should -Be (Get-PcmGuiSelectionKey -Prefix 'abcd1234' -Guids @('A', 'b'))
        $one | Should -Not -Be (Get-PcmGuiSelectionKey -Prefix 'abcd1234' -Guids @('a', 'c'))
        $one | Should -Match '^abcd1234\|2\|[0-9A-F]{16}$'
    }
}

Describe 'Console: what this computer can do' {
    BeforeAll {
        $script:Modules = @([pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = [version]'2.41.1' }, [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = [version]'3.10.1' })
        $script:Snapshot = [pscustomobject]@{ Path = 'D:\PRA\snap.db'; Exists = $true; Id = 42; Taken = '2026-11-11 22:00'; AgeHours = 14.0; MaxAgeDays = 2; Error = '' }
        $script:Journal = [pscustomobject]@{ Path = 'D:\PRA\journal.db'; Exists = $true; Batches = 1; Error = '' }
        $script:CertificateOk = [pscustomobject]@{ Certificate = [pscustomobject]@{ Subject = 'CN=PRA-Cloud-App'; NotAfter = [datetime]::Now.AddDays(230); HasPrivateKey = $true }; Store = 'CurrentUser' }
        $script:Pwsh = [pscustomobject]@{ Ok = $true; Path = 'C:\Program Files\PowerShell\7\pwsh.exe'; Version = '7.5.1'; Problem = '' }
    }
    It 'shows the lights of a computer that can run every action, and settles the role from the Microsoft 365 and Exchange tests' {
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.5.1'; Modules = $script:Modules; Snapshot = $script:Snapshot; Journal = $script:Journal; Certificate = $script:CertificateOk; Thumbprint = 'AA'; AppId = 'id'
            ExchangeServer = 'EX01.contoso.com'; Pwsh = $script:Pwsh; Task = [pscustomobject]@{ Found = $true; Text = 'Scheduled task "PRA Collect" found.' } }
        @($detect.Lights | ForEach-Object { $_.Name }) | Should -Be @('PowerShell 7.5', 'Cloud modules', 'Sign-in: certificate', 'Snapshot', 'Journal', 'Daily Collect')
        @($detect.Lights | Where-Object { $_.Status -eq 'Fail' }).Count | Should -Be 0
        $detect.CanCloudBase | Should -BeTrue
        $detect.CanCollect | Should -BeTrue
        $detect.ExchangeServer | Should -Be 'EX01.contoso.com'
        Get-PcmGuiRole -Detect $detect -Net 'Ok' -Exchange 'Ok' | Should -Be 'Both'
        Get-PcmGuiRole -Detect $detect -Net 'Ok' -Exchange '' | Should -Be 'Both'
        Get-PcmGuiRole -Detect $detect -Net 'Ok' -Exchange 'Fail' | Should -Be 'Cloud'
        Get-PcmGuiRole -Detect $detect -Net 'Fail' -Exchange 'Ok' | Should -Be 'Collect'
        Get-PcmGuiRole -Detect $detect -Net 'Checking' -Exchange 'Fail' | Should -Be 'Checking'
        Get-PcmGuiRole -Detect $detect -Net 'Fail' -Exchange 'Fail' | Should -Be 'None'
    }
    It 'does not let this computer collect without Exchange.Server, and keeps the cloud actions when the cloud lights are green' {
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.5.1'; Modules = $script:Modules; Snapshot = $script:Snapshot; Journal = $script:Journal; Certificate = $script:CertificateOk; Thumbprint = 'AA'; AppId = 'id'
            ExchangeServer = '  '; Pwsh = $script:Pwsh }
        $detect.CanCollect | Should -BeFalse
        Get-PcmGuiRole -Detect $detect -Net 'Ok' | Should -Be 'Cloud'
    }
    It 'needs PowerShell 7.5 for every action' {
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.4.6'; Modules = $script:Modules; Snapshot = $script:Snapshot; ExchangeServer = 'EX01.contoso.com'; Pwsh = $script:Pwsh }
        ($detect.Lights | Where-Object { $_.Name -eq 'PowerShell 7.5' }).Status | Should -Be 'Fail'
        $detect.CanCollect | Should -BeFalse
        $detect.CanCloudBase | Should -BeFalse
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Desktop'; Version = '5.1.17763.1'; Modules = $script:Modules; Snapshot = $script:Snapshot; ExchangeServer = 'EX01.contoso.com'; Pwsh = $script:Pwsh }
        ($detect.Lights | Where-Object { $_.Name -eq 'PowerShell 7.5' }).Detail | Should -BeLike '*every action needs PowerShell 7.5 or later*'
        $detect.CanCollect | Should -BeFalse
    }
    It 'writes the Exchange light: not set, checking, answering, or not reachable (amber: only Collect uses it)' {
        (Get-PcmGuiExchangeLight -State '' -Server '').Status | Should -Be 'NA'
        (Get-PcmGuiExchangeLight -State '' -Server '').Detail | Should -BeLike '*Exchange.Server is not set*'
        (Get-PcmGuiExchangeLight -State 'Checking' -Server 'EX01.contoso.com').Status | Should -Be 'Info'
        $ok = Get-PcmGuiExchangeLight -State 'Ok' -Server 'EX01.contoso.com'
        $ok.Status | Should -Be 'Ok'
        $ok.Detail | Should -BeLike '*EX01.contoso.com*Kerberos*'
        $fail = Get-PcmGuiExchangeLight -State 'Fail' -Answer ([pscustomobject]@{ Host = 'EX01.contoso.com'; Ok = $false; Why = 'The server did not answer.' }) -Server 'EX01.contoso.com'
        $fail.Status | Should -Be 'Warn'
        $fail.Detail | Should -BeLike '*cannot be opened*Collect cannot run here*The server did not answer.*'
    }
    It 'tests Exchange with a task-like probe that never blocks the window (answer, silence, refusal)' {
        $answered = [pscustomobject]@{ IsCompleted = $true; Result = [pscustomobject]@{ Ok = $true; Why = '' } }
        $silent = [pscustomobject]@{ IsCompleted = $false; Result = $null }
        $refused = [pscustomobject]@{ IsCompleted = $true; Result = [pscustomobject]@{ Ok = $false; Why = 'Access is denied.' } }
        $answer = { param($target, $seconds) ConvertTo-PcmGuiExchangeAnswer -Target $target -Seconds $seconds }
        (& $script:Gui $answer ([pscustomobject]@{ Host = 'ex'; Task = $answered; Failure = '' }) 15).Ok | Should -BeTrue
        $none = & $script:Gui $answer ([pscustomobject]@{ Host = 'ex'; Task = $silent; Failure = '' }) 15
        $none.Ok | Should -BeFalse
        $none.Why | Should -BeLike 'no answer within 15 seconds'
        (& $script:Gui $answer ([pscustomobject]@{ Host = 'ex'; Task = $refused; Failure = '' }) 15).Why | Should -Be 'Access is denied.'
        (& $script:Gui $answer ([pscustomobject]@{ Host = 'ex'; Task = $null; Failure = 'no probe' }) 15).Why | Should -Be 'no probe'
    }
    It 'explains an old module, an expired certificate and a certificate without a private key in plain words' {
        $old = @([pscustomobject]@{ Name = 'Microsoft.Graph.Authentication'; Version = [version]'2.10.0' }, [pscustomobject]@{ Name = 'ExchangeOnlineManagement'; Version = [version]'3.10.1' })
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.5.1'; Modules = $old; Pwsh = $script:Pwsh; Thumbprint = 'AA'; AppId = 'id'
            Certificate = [pscustomobject]@{ Certificate = [pscustomobject]@{ Subject = 'CN=x'; NotAfter = [datetime]::Now.AddDays(-3); HasPrivateKey = $true }; Store = 'CurrentUser' } }
        ($detect.Lights | Where-Object { $_.Name -eq 'Cloud modules' }).Detail | Should -BeLike '*Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.41.1*'
        ($detect.Lights | Where-Object { $_.Name -eq 'Sign-in: certificate' }).Detail | Should -BeLike '*expired*'
        $detect.CanCloudBase | Should -BeFalse
        $light = { param($SignIn, $Thumbprint, $AppId, $Tenant, $Find) Get-PcmGuiSignInLight -SignIn $SignIn -Thumbprint $Thumbprint -AppId $AppId -Tenant $Tenant -Find $Find }
        (& $script:Gui $light 'Certificate' 'AA' 'id' '' ([pscustomobject]@{ Certificate = [pscustomobject]@{ Subject = 'CN=x'; NotAfter = [datetime]::Now.AddDays(90); HasPrivateKey = $false }; Store = 'CurrentUser' })).Detail | Should -BeLike '*no private key*'
        (& $script:Gui $light 'interactive' '' '' '' $null).Status | Should -Be 'Fail'
        (& $script:Gui $light 'INTERACTIVE' '' '' '00000000-0000-0000-0000-000000000000' $null).Ok | Should -BeTrue
    }
    It 'warns when the snapshot is older than the maximum age and fails when there is none' {
        $old = [pscustomobject]@{ Path = 'D:\PRA\snap.db'; Exists = $true; Id = 3; Taken = '2026-10-01 22:00'; AgeHours = 24.0 * 40; MaxAgeDays = 7; Error = '' }
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.5.1'; Modules = $script:Modules; Snapshot = $old; Pwsh = $script:Pwsh }
        ($detect.Lights | Where-Object { $_.Name -eq 'Snapshot' }).Status | Should -Be 'Warn'
        $none = [pscustomobject]@{ Path = 'D:\PRA\snap.db'; Exists = $false; Id = 0; Taken = ''; AgeHours = 0.0; MaxAgeDays = 7; Error = '' }
        $detect = Get-PcmGuiPrerequisite -Request @{ Edition = 'Core'; Version = '7.5.1'; Modules = $script:Modules; Snapshot = $none; Pwsh = $script:Pwsh }
        ($detect.Lights | Where-Object { $_.Name -eq 'Snapshot' }).Status | Should -Be 'Fail'
    }
    It 'says why Microsoft 365 cannot be reached: a blocked or refused connection, a name that is not found, a silent network' {
        $socket = { param($code) [Net.Sockets.SocketException]::new($code) }
        $wrap = { param($inner) [Net.Http.HttpRequestException]::new('An error occurred while sending the request.', $inner) }
        Get-PcmGuiNetworkWhy (& $wrap (& $socket 10013)) | Should -BeLike 'blocked on this computer*'
        Get-PcmGuiNetworkWhy (& $wrap (& $socket 10061)) | Should -BeLike 'refused by the network*'
        Get-PcmGuiNetworkWhy (& $wrap (& $socket 11001)) | Should -BeLike 'the name cannot be resolved*'
        Get-PcmGuiNetworkWhy (& $wrap (& $socket 10060)) | Should -BeLike 'no answer from the network*'
        Get-PcmGuiNetworkWhy ([InvalidOperationException]::new('boom')) | Should -Be 'boom'
    }
    It 'reads a hook that the window was not given as no hook, whatever the letter case, without a strict-mode error' {
        $answer = & $script:Gui {
            $saved = $script:Gui
            try {
                $script:Gui = @{ Hooks = @{ Question = { $true } } }
                [pscustomobject]@{ Given = [bool](Get-PcmGuiHook 'QUESTION'); Absent = ($null -eq (Get-PcmGuiHook 'NoConsole')) }
            } finally { $script:Gui = $saved }
        }
        $answer.Given | Should -BeTrue
        $answer.Absent | Should -BeTrue
    }
    It 'decides Ok, Partial or Fail from the answers, and writes the light' {
        $good = [pscustomobject]@{ Host = 'login.microsoftonline.com'; Ok = $true; Why = '' }
        $bad = [pscustomobject]@{ Host = 'outlook.office365.com'; Ok = $false; Why = 'no answer within 3 seconds' }
        Get-PcmGuiNetworkVerdict @($good, $good) | Should -Be 'Ok'
        Get-PcmGuiNetworkVerdict @($good, $bad) | Should -Be 'Partial'
        Get-PcmGuiNetworkVerdict @($bad, $bad) | Should -Be 'Fail'
        (Get-PcmGuiNetworkLight -State 'Partial' -Answers @($good, $bad)).Detail | Should -BeLike '*outlook.office365.com does not answer*no answer within 3 seconds*'
        (Get-PcmGuiNetworkLight -State 'Ok' -Answers @($good, $good)).Status | Should -Be 'Ok'
        (Get-PcmGuiNetworkLight -State 'Checking').Status | Should -Be 'Info'
        (Get-PcmGuiNetworkLight -State 'Checking' -Answers $null).Status | Should -Be 'Info'
        (Get-PcmGuiNetworkLight -State 'Fail' -Answers $null).Status | Should -Be 'Fail'
    }
    It 'tests the two Microsoft 365 hosts with tasks that never block the window (answers, timeouts and failures)' {
        $done = [Threading.Tasks.Task]::FromResult([object]::new())
        $stuck = [Threading.Tasks.TaskCompletionSource[object]]::new().Task
        $failed = [Threading.Tasks.Task]::FromException([Exception]::new('connection refused'))
        $answers = @(& $script:Gui { param($a, $b, $c) ConvertTo-PcmGuiNetworkAnswer -Targets @([pscustomobject]@{ Host = 'a'; Task = $a; Failure = '' }, [pscustomobject]@{ Host = 'b'; Task = $b; Failure = '' }, [pscustomobject]@{ Host = 'c'; Task = $null; Failure = 'no probe' }) -Seconds 3 } $done $stuck $failed)
        $answers[0].Ok | Should -BeTrue
        $answers[1].Ok | Should -BeFalse
        $answers[1].Why | Should -BeLike 'no answer within 3 seconds'
        $answers[2].Why | Should -Be 'no probe'
    }
}

Describe 'Console: settings for one run only' {
    It 'writes the changes to a generated copy, never to the configuration file, and lists what differs' {
        $before = (Get-FileHash -LiteralPath $script:Fixture.Config).Hash
        $made = New-PcmGuiRunConfig -Root $script:Tool -ConfigPath $script:Fixture.Config -Overrides @{ Environment = 'LAB'; Cloud = @{ TenantId = '11111111-2222-3333-4444-555555555555'; SignIn = 'Interactive'; UserPrincipalName = 'admin@contoso.com' } }
        (Get-FileHash -LiteralPath $script:Fixture.Config).Hash | Should -Be $before
        $made.Path | Should -BeLike '*logs\gui\run-*\effective.config.psd1'
        $made.Base | Should -Be $script:Fixture.Config
        @($made.Keys) | Should -Be @('Cloud.SignIn', 'Cloud.TenantId', 'Cloud.UserPrincipalName', 'Environment')
        $loaded = Import-PcmConfiguration -Path $made.Path -Root $script:Tool
        $loaded.Environment | Should -Be 'LAB'
        $loaded.Cloud.SignIn | Should -Be 'Interactive'
        $loaded.Cloud.TenantId | Should -Be '11111111-2222-3333-4444-555555555555'
        $loaded.Cloud.Organization | Should -Be 'contoso.onmicrosoft.com'
        $made.Hash | Should -Be (Get-FileHash -LiteralPath $made.Path).Hash.ToLowerInvariant()
    }
    It 'writes an empty wave size as 0, which the engine reads as "the same as the temporary licences", and keeps the licences kept free' {
        $answer = & $script:Gui {
            param($root, $config)
            $zero = [int](ConvertTo-PcmGuiWholeNumber '')
            $made = New-PcmGuiRunConfig -Root $root -ConfigPath $config -Overrides @{ Licensing = @{ Shared = @{ Parallel = 40; KeepFree = 3 } }; Recover = @{ WaveSize = $zero } }
            $loaded = Import-PcmConfiguration -Path $made.Path -Root $root
            [pscustomobject]@{ Zero = $zero; Size = (Get-PcmRecoverWaveSize $loaded); Keep = $loaded.Licensing.Shared.KeepFree; Keys = @($made.Keys) }
        } $script:Tool $script:Fixture.Config
        $answer.Zero | Should -Be 0
        $answer.Size | Should -Be 40
        $answer.Keep | Should -Be 3
        @($answer.Keys) | Should -Be @('Licensing.Shared.KeepFree', 'Licensing.Shared.Parallel', 'Recover.WaveSize')
    }
    It 'gives the configuration file itself when nothing was changed, and refuses a change the engine would refuse' {
        $same = New-PcmGuiRunConfig -Root $script:Tool -ConfigPath $script:Fixture.Config
        $same.Path | Should -Be $script:Fixture.Config
        $same.Base | Should -BeNullOrEmpty
        { New-PcmGuiRunConfig -Root $script:Tool -ConfigPath $script:Fixture.Config -Overrides @{ Cloud = @{ SignIn = 'Magic' } } } | Should -Throw '*not valid*SignIn*'
        { New-PcmGuiRunConfig -Root $script:Tool -ConfigPath (Join-Path $TestDrive 'missing.psd1') -Overrides @{ Environment = 'X' } } | Should -Throw '*not found*'
    }
    It 'writes text, numbers, booleans, lists and sections that PowerShell reads back unchanged' {
        $text = ConvertTo-PcmGuiDataLiteral @{ A = "it's"; B = 3; C = $true; D = @('x', 'y'); E = @{ F = '' }; G = @() }
        $back = & ([scriptblock]::Create($text))
        $back.A | Should -Be "it's"
        $back.B | Should -Be 3
        $back.C | Should -BeTrue
        $back.D | Should -Be @('x', 'y')
        $back.E.F | Should -Be ''
        @($back.G).Count | Should -Be 0
    }
}

Describe 'Console: the snapshot, the journal and the lookups' {
    It 'reads the configuration, the snapshot, the objects and the journal without changing anything' {
        $script:State.ConfigError | Should -BeNullOrEmpty
        $script:State.Snapshot.id | Should -Be 42
        @($script:State.Objects).Count | Should -Be 24
        @($script:State.Batches).Count | Should -Be 4
        (Get-PcmGuiSnapshotInfo -State $script:State).MaxAgeDays | Should -Be 2
        [math]::Round((Get-PcmGuiSnapshotInfo -State $script:State).AgeHours) | Should -BeGreaterOrEqual 13
    }
    It 'reports an unreadable file in the state instead of throwing' {
        $broken = Join-Path $TestDrive 'broken.psd1'
        [IO.File]::WriteAllText($broken, '@{ Environment = ''PROD'' ; Scope = @{ Mode = ''Magic'' } }')
        (Read-PcmGuiState -Root $script:Tool -ConfigPath $broken).ConfigError | Should -BeLike '*Scope.Mode*'
    }
    It 'finds a mailbox by any identity, in any letter case (<Text>)' -TestCases @(
        @{ Text = 'ADELE.VANCE@CONTOSO.COM' }, @{ Text = 'adelevance' }, @{ Text = 'Adele Vance' }, @{ Text = 'A0000000-0000-0000-0000-000000000001' }, @{ Text = '{a0000000-0000-0000-0000-000000000001}' }) {
        param($Text)
        @(Find-PcmGuiMailbox -State $script:State -Text $Text).Count | Should -Be 1
    }
    It 'finds nothing for an unknown identity or an empty text' {
        @(Find-PcmGuiMailbox -State $script:State -Text 'nobody@contoso.com').Count | Should -Be 0
        @(Find-PcmGuiMailbox -State $script:State -Text '  ').Count | Should -Be 0
    }
    It 'lists the OUs of the snapshot with the mailboxes under each, sub-OUs included' {
        $units = @(Get-PcmGuiOuList -State $script:State)
        ($units | Where-Object { $_.Name -eq 'contoso.com/Paris' }).Count | Should -Be 12
        ($units | Where-Object { $_.Name -eq 'contoso.com/Paris/Users' }).Count | Should -Be 12
        ($units | Where-Object { $_.Name -eq 'contoso.com/Lyon' }).Count | Should -Be 8
        @($units | Where-Object { $_.Name -eq 'contoso.com' }).Count | Should -Be 0
    }
    It 'matches an OU given as a path or as a distinguished name, in any letter case, sub-OUs included' -TestCases @(
        @{ Text = 'CONTOSO.COM/paris'; Count = 12 }, @{ Text = 'contoso.com/Paris/Users'; Count = 12 }, @{ Text = 'OU=Users,OU=Paris,DC=contoso,DC=com'; Count = 12 }
        @{ Text = 'ou=lyon,dc=CONTOSO,dc=com'; Count = 8 }, @{ Text = 'contoso.com/Shared'; Count = 4 }, @{ Text = 'contoso.com/Nowhere'; Count = 0 }) {
        param($Text, $Count)
        @(Get-PcmGuiOuMatch -State $script:State -Text $Text).Count | Should -Be $Count
    }
    It 'finds the group Collect captured by name, account name, DN or ID, and counts its members and the ones that are not mailboxes' -TestCases @(
        @{ Text = 'pra-wave1' }, @{ Text = 'PRA-WAVE1' }, @{ Text = 'cn=pra-wave1,ou=groups,dc=contoso,dc=com' }, @{ Text = 'B0000000-0000-0000-0000-000000000001' }) {
        param($Text)
        $found = Get-PcmGuiGroupInSnapshot -State $script:State -Value $Text
        $found.HasTables | Should -BeTrue
        $found.Count | Should -Be 1
        $found.Name | Should -Be 'PRA-Wave1'
        $found.Members | Should -Be 26
        $found.Mailboxes | Should -Be 24
        $found.Skipped | Should -Be 2
    }
    It 'says a group is not in the snapshot (it will be looked for in Microsoft Entra), also for a snapshot written before groups were captured' {
        (Get-PcmGuiGroupInSnapshot -State $script:State -Value 'Nobody').Count | Should -Be 0
        $old = New-DocumentationToolCopy -Repository $script:Repo -Destination (Join-Path $TestDrive 'old-tool')
        $fixture = New-ConsoleFixture -ToolRoot $old
        $connection = Open-PcmStore -Path (Join-Path $old 'data\PraCloudMailbox.db') -Root $old
        try { $null = Invoke-PcmSql $connection 'DROP TABLE target_member'; $null = Invoke-PcmSql $connection 'DROP TABLE target_group' } finally { Close-PcmStore $connection }
        $state = Read-PcmGuiState -Root $old -ConfigPath $fixture.Config
        $answer = Get-PcmGuiGroupInSnapshot -State $state -Value 'PRA-Wave1'
        $answer.HasTables | Should -BeFalse
        $answer.Count | Should -Be 0
        $answer.Error | Should -BeNullOrEmpty
    }
    It 'sums up each Convert batch: converted, left, waiting for Recover (the case of the object IDs ignored)' {
        $summary = @(Get-PcmGuiBatchSummary -State $script:State)
        ($summary | Where-Object { $_.Id -eq '1a2b3c4d' }).Waiting | Should -Be 24
        $partial = $summary | Where-Object { $_.Id -eq '5e6f7a8b' }
        $partial.Left | Should -Be 8
        $partial.Converted | Should -Be 12
        $rolled = $summary | Where-Object { $_.Id -eq '9c8d7e6f' }
        $rolled.Waiting | Should -Be 2
        $rolled.Recover | Should -BeLike 'Partial*'
        $rolled.Back.Contains('A0000000-0000-0000-0000-000000000001') | Should -BeTrue
    }
    It 'builds the batch lists of the Task page: batches to continue, to roll back, rollbacks to continue, batches to follow' {
        $continue = @(Get-PcmGuiBatchEntry -State $script:State -Kind ContinueConvert)
        ($continue | Where-Object { $_.BatchId -eq '5e6f7a8b' }).Selectable | Should -BeTrue
        ($continue | Where-Object { $_.BatchId -eq '1a2b3c4d' }).Selectable | Should -BeFalse
        ($continue | Where-Object { $_.BatchId -eq '1a2b3c4d' }).State | Should -Be 'Complete'
        $back = @(Get-PcmGuiBatchEntry -State $script:State -Kind Rollback)
        ($back | Where-Object { $_.BatchId -eq '1a2b3c4d' }).State | Should -Be 'Can be rolled back'
        ($back | Where-Object { $_.BatchId -eq '9c8d7e6f' }).State | Should -Be 'Partly rolled back'
        $interrupted = @(Get-PcmGuiBatchEntry -State $script:State -Kind ContinueRollback)
        $interrupted.Count | Should -Be 0
        @(Get-PcmGuiBatchEntry -State $script:State -Kind Follow).Count | Should -Be 4
    }
    It 'lists for a wave the mailboxes of a batch that are still in the cloud' {
        @(Get-PcmGuiWaveRow -State $script:State -BatchId '1a2b3c4d').Count | Should -Be 24
        @(Get-PcmGuiWaveRow -State $script:State -BatchId '9C8D7E6F').Count | Should -Be 2
        (@(Get-PcmGuiWaveRow -State $script:State -BatchId '1a2b3c4d')[0]).Name | Should -Be 'Adele Vance'
        @(Get-PcmGuiWaveRow -State $script:State -BatchId 'ffffffff').Count | Should -Be 0
    }
}

Describe 'Console: the child process of a run' {
    It 'runs a cloud action in its own console with the three standard streams redirected, and an end of file on its input' {
        $pwsh = Get-PcmGuiEngine
        $pwsh.Ok | Should -BeTrue
        $command = [pscustomobject]@{ FilePath = $pwsh.Path; NeedsConsoleWindow = $true
            Arguments = '-NoProfile -Command "[Console]::Out.WriteLine(\"out-line\"); [Console]::Error.WriteLine(\"err-line\"); $in = [Console]::In.ReadLine(); if ($null -eq $in) { exit 7 } else { exit 8 }"' }
        $child = Start-PcmGuiChildProcess -Command $command -Root $script:Tool -EventPath (Join-Path $TestDrive 'child.events.jsonl') -StopPath (Join-Path $TestDrive 'child.stop')
        $null = $child.Process.WaitForExit(30000)
        $child.Process.ExitCode | Should -Be 7
        $child.Output.Result | Should -BeLike '*out-line*'
        $child.Errors.Result | Should -BeLike '*err-line*'
        $child.Process.Dispose()
    }
    It 'gives the child the file of its events and the file that asks it to stop' {
        $pwsh = Get-PcmGuiEngine
        $command = [pscustomobject]@{ FilePath = $pwsh.Path; NeedsConsoleWindow = $false; Arguments = '-NoProfile -Command "[Console]::Out.WriteLine($env:PRA_EVENT_FILE + \"|\" + $env:PRA_STOP_FILE)"' }
        $child = Start-PcmGuiChildProcess -Command $command -Root $script:Tool -EventPath (Join-Path $TestDrive 'x.events.jsonl') -StopPath (Join-Path $TestDrive 'x.stop')
        $null = $child.Process.WaitForExit(30000)
        $child.Output.Result.Trim() | Should -Be ((Join-Path $TestDrive 'x.events.jsonl') + '|' + (Join-Path $TestDrive 'x.stop'))
        $child.Process.Dispose()
    }
}

Describe 'Console: the real window, Convert of a group (synthetic engine, no live connection)' {
    BeforeAll {
        $script:Convert = Invoke-GuiChild -Body @'
$o = [ordered]@{}
$o.firstPage = $g.Page; $o.firstNext = $c.BtnNext.Content; $o.nextEnabled = $c.BtnNext.IsEnabled; $o.role = $c.RoleName.Text
$o.page3Reachable = [bool](Mod { Test-PcmGuiPageReachable 3 }); $o.page5Reachable = [bool](Mod { Test-PcmGuiPageReachable 5 })
$o.lights = $g.LightRows.Count
Click 'BtnNext'
$o.page = $g.Page
$o.tiles = @{ Collect = $c.T_Collect.IsEnabled; Convert = $c.T_Convert.IsEnabled; Recover = $c.T_Recover.IsEnabled; Check = $c.T_Check.IsEnabled }
$o.collectPill = $c.NT_Collect.Text; $o.cloudPill = $c.NT_Convert.Text
$o.readyBeforeTask = [bool](Mod { Test-PcmGuiTaskReady })
Check 'T_Convert'; Check 'S_Source_Group'
$o.readyNoGroup = [bool](Mod { Test-PcmGuiTaskReady })
Enter 'F_Group' 'pra-wave1'; Click 'BtnCheckGroup'
WaitFor { -not $g.GroupLookup }
$o.groupInfo = $c.GroupInfo.Text
Check 'S_Scope_UsersOnly'
$c.SettingsCard.Visibility = 'Visible'
Enter 'F_Tenant' '11111111-2222-3333-4444-555555555555'
$o.changed = @(Mod { Get-PcmGuiChangedSetting })
$o.overrideKeys = @(Mod { (Get-PcmGuiOverride).Keys })
$o.runLine = $c.RunLineSub.Text
Check 'S_Scope_All'
Click 'BtnNext'
$o.verifyPage = $g.Page; $o.confirmEnabledBefore = $c.ConfirmBox.IsEnabled; $o.runEnabledBefore = $c.BtnNext.IsEnabled; $o.summaryBefore = $c.SumTitle.Text
$o.applyRefused = ''
try { Mod { Start-PcmGuiRun -Mode Apply } } catch { $o.applyRefused = $_.Exception.Message }
Click 'BtnPreview'; WaitRun
$a = Args_
$o.previewArgs = @{ Action = $a.Action; Mode = $a.Mode; Group = $a.Group; Scope = $a.Scope; Force = $a.Force; BaseConfigPath = $a.BaseConfigPath; ConfigPath = $a.ConfigPath }
$o.effectiveHasTenant = (Get-Content -LiteralPath (Join-Path $root 'effective-copy.psd1') -Raw) -match '11111111-2222-3333-4444-555555555555'
$o.previewed = $g.Previewed; $o.rows = $g.VerRows.Count; $o.firstRow = $g.VerRows[0].Status; $o.firstName = $g.VerRows[0].Upn
$o.skippedRows = @($g.VerRows | Where-Object { $_.Status -eq 'Skip' }).Count; $o.sharedLast = $g.VerRows[$g.VerRows.Count - 1].Type
$o.summary = $c.SumTitle.Text; $o.target = $c.TargetText.Text; $o.chips = @($c.Chips.Children | ForEach-Object { $_.Child.Text })
$o.confirmEnabledAfter = $c.ConfirmBox.IsEnabled; $o.runEnabledAfter = $c.BtnNext.IsEnabled
Enter 'ConfirmBox' ('conv' + [char]0x200B + 'ert'); $o.zeroWidthEnabled = $c.BtnNext.IsEnabled
Enter 'ConfirmBox' 'recover'; $o.otherWordEnabled = $c.BtnNext.IsEnabled
Enter 'ConfirmBox' ' Convert '; $o.typedEnabled = $c.BtnNext.IsEnabled
Check 'S_Scope_UsersOnly'
$o.afterChangePending = [bool]$g.Pending; $o.afterChangeRows = $g.VerRows.Count; $o.afterChangeBox = $c.ConfirmBox.Text; $o.afterChangeEnabled = $c.BtnNext.IsEnabled
Check 'S_Scope_All'
Click 'BtnPreview'; WaitRun
Enter 'ConfirmBox' 'CONVERT'
$o.finalEnabled = $c.BtnNext.IsEnabled
Ask
Click 'BtnNext'
$o.runningPage = $g.Page; $o.stopVisible = $c.BtnStop.Visibility.ToString()
WaitRun
$a = Args_
$o.applyArgs = @{ Action = $a.Action; Mode = $a.Mode; Group = $a.Group; Scope = $a.Scope; Force = $a.Force }
$o.steps = @($g.Model.Steps | ForEach-Object { $_.Title })
$o.beforeFirstStep = ($g.Model.Raw['s1'] -join "`n"); $o.lastStepRaw = $(if ($g.Model.Raw.ContainsKey('s7')) { $g.Model.Raw['s7'] -join "`n" } else { '' })
$o.stepStatus = @($g.Model.Steps | ForEach-Object { $_.Status })
$o.user6 = $g.Model.Objects['s6'].Count
$o.questions = @($questions); $o.answer = (Get-Content -LiteralPath (Join-Path $root 'asked-answer.txt') -Raw).Trim()
$o.lastValid = $g.LastRun.Valid; $o.lastLevel = $g.LastRun.Level
$o.runPageAfter = $g.Page; $o.whatNext = $c.BtnNext.Content
Click 'BtnNext'
$o.nextPage = $g.Page; $o.pill = $c.NextPillText.Text; $o.title = $c.NextTitle.Text; $o.batch = $c.BatchIdText.Text; $o.command = $c.NextCmd.Text; $o.result = $c.ResultText.Text
$o.timeline = @($g.TimelineRows | ForEach-Object { '{0}|{1}' -f $_.Title, $_.Status })
$o.stepsList = @($g.NextRows | ForEach-Object { $_.Text })
$o.backEnabled = $c.BtnBack.IsEnabled; $o.nav3 = [bool](Mod { Test-PcmGuiPageReachable 3 })
$o.configUnchanged = ((Get-FileHash -LiteralPath $fixture.Config -Algorithm SHA256).Hash -eq $configHash)
Click 'BtnNext'
$o.newTaskPage = $g.Page; $o.newTaskAction = $g.Action
Done $o
'@
    }
    It 'opens on Setup, with Next steps and Verify out of reach, the lights shown and the role found' {
        $script:Convert.firstPage | Should -Be 1
        $script:Convert.firstNext | Should -BeLike 'Choose a task*'
        $script:Convert.nextEnabled | Should -BeTrue
        $script:Convert.role | Should -Be 'Everything'
        $script:Convert.page3Reachable | Should -BeFalse
        $script:Convert.page5Reachable | Should -BeFalse
        $script:Convert.lights | Should -Be 4
    }
    It 'offers every task when this computer can run them, and asks for a group before going on' {
        $script:Convert.page | Should -Be 2
        $script:Convert.tiles.Collect | Should -BeTrue
        $script:Convert.tiles.Convert | Should -BeTrue
        $script:Convert.readyBeforeTask | Should -BeFalse
        $script:Convert.readyNoGroup | Should -BeFalse
    }
    It 'finds the group in the snapshot, whatever the letter case, and says how many members are skipped' {
        $script:Convert.groupInfo | Should -BeLike 'Found in the snapshot (captured by Collect): PRA-Wave1 * 26 members including nested groups, 2 skipped (the Preview says why)'
    }
    It 'marks what was changed for this run only and keeps it out of the configuration file' {
        @($script:Convert.changed) | Should -Be @('Scope', 'Tenant ID')
        @($script:Convert.overrideKeys) | Should -Be @('Cloud')
        $script:Convert.runLine | Should -BeLike '2 settings changed for this run only: Scope, Tenant ID.'
        $script:Convert.configUnchanged | Should -BeTrue
    }
    It 'asks for a preview first: no run, no confirmation, and a refusal in plain words' {
        $script:Convert.verifyPage | Should -Be 3
        $script:Convert.confirmEnabledBefore | Should -BeFalse
        $script:Convert.runEnabledBefore | Should -BeFalse
        $script:Convert.summaryBefore | Should -Be 'Preview needed before running'
        $script:Convert.applyRefused | Should -BeLike '*preview and the typed confirmation*'
    }
    It 'previews with the source, the scope and the copy of the configuration, never with -Force' {
        $script:Convert.previewArgs.Action | Should -Be 'Convert'
        $script:Convert.previewArgs.Mode | Should -Be 'Preview'
        $script:Convert.previewArgs.Group | Should -Be 'pra-wave1'
        $script:Convert.previewArgs.Scope | Should -BeNullOrEmpty
        $script:Convert.previewArgs.Force | Should -BeFalse
        $script:Convert.previewArgs.BaseConfigPath | Should -BeLike '*PraCloudMailbox.config.psd1'
        $script:Convert.previewArgs.ConfigPath | Should -BeLike '*effective.config.psd1'
        $script:Convert.effectiveHasTenant | Should -BeTrue
    }
    It 'shows what needs attention first, then the users, then the shared mailboxes, with the counts of the engine' {
        $script:Convert.previewed | Should -BeTrue
        $script:Convert.rows | Should -Be 26
        $script:Convert.firstRow | Should -Be 'Warn'
        $script:Convert.firstName | Should -Be 'alex.wilber@contoso.com'
        $script:Convert.skippedRows | Should -Be 2
        $script:Convert.sharedLast | Should -Be 'Shared'
        $script:Convert.summary | Should -Be '24 mailboxes will be converted · nothing blocking'
        $script:Convert.target | Should -BeLike 'Group pra-wave1 (captured by Collect, snapshot 42)*'
        @($script:Convert.chips) | Should -Contain 'Users 20'
        @($script:Convert.chips) | Should -Contain 'Shared 4'
        @($script:Convert.chips) | Should -Contain 'Skipped 2'
        @($script:Convert.chips) | Should -Contain '0 blocking'
        @($script:Convert.chips) | Should -Contain '1 warning'
    }
    It 'unlocks Run only with the word of the action, in any letter case, never with an invisible character' {
        $script:Convert.confirmEnabledAfter | Should -BeTrue
        $script:Convert.runEnabledAfter | Should -BeFalse
        $script:Convert.zeroWidthEnabled | Should -BeFalse
        $script:Convert.otherWordEnabled | Should -BeFalse
        $script:Convert.typedEnabled | Should -BeTrue
    }
    It 'voids the preview and the confirmation when the task or a setting changes' {
        $script:Convert.afterChangePending | Should -BeFalse
        $script:Convert.afterChangeRows | Should -Be 0
        $script:Convert.afterChangeBox | Should -Be ''
        $script:Convert.afterChangeEnabled | Should -BeFalse
        $script:Convert.finalEnabled | Should -BeTrue
    }
    It 'runs with -Force once confirmed, and follows the steps of the engine with their objects' {
        $script:Convert.runningPage | Should -Be 4
        $script:Convert.stopVisible | Should -Be 'Visible'
        $script:Convert.applyArgs.Mode | Should -Be 'Apply'
        $script:Convert.applyArgs.Force | Should -BeTrue
        $script:Convert.applyArgs.Group | Should -Be 'pra-wave1'
        $script:Convert.applyArgs.Scope | Should -BeNullOrEmpty
        @($script:Convert.steps).Count | Should -Be 7
        $script:Convert.steps[5] | Should -Be '6. Convert the users'
        @($script:Convert.stepStatus | Select-Object -Unique) | Should -Not -Contain 'Running'
        $script:Convert.beforeFirstStep | Should -BeLike '*Cloud modules: Microsoft.Graph.Authentication*'
        $script:Convert.lastStepRaw | Should -Not -BeLike '*Cloud modules*'
        $script:Convert.user6 | Should -Be 20
        $script:Convert.lastValid | Should -BeTrue
        $script:Convert.runPageAfter | Should -Be 4
        $script:Convert.whatNext | Should -BeLike 'What next?*'
    }
    It 'asks the question of the run in the window and answers it in the file the run named' {
        @($script:Convert.questions).Count | Should -Be 1
        $script:Convert.questions[0] | Should -BeLike 'Start a delta synchronisation*'
        $script:Convert.answer | Should -Be 'yes'
    }
    It 'ends on the Next steps page with the batch ID, the command to roll it back, the timeline and no way back to Verify' {
        $script:Convert.nextPage | Should -Be 5
        $script:Convert.pill | Should -BeLike 'Done*warning'
        $script:Convert.title | Should -Be 'The conversion is complete'
        $script:Convert.batch | Should -Be '1a2b3c4d'
        $script:Convert.command | Should -BeLike '.\Invoke-PraCloudMailbox.ps1 -Action Recover -Batch 1a2b3c4d'
        $script:Convert.result | Should -BeLike 'Converted 24 (20 users, 4 shared)*1 warning*'
        @($script:Convert.timeline) | Should -Be @('Collect|Ok', 'Convert|Warn', 'Recover|Skip')
        @($script:Convert.stepsList).Count | Should -BeGreaterOrEqual 4
        $script:Convert.stepsList[0] | Should -BeLike 'Open the report and read the 1 warning*'
        $script:Convert.nav3 | Should -BeFalse
        $script:Convert.newTaskPage | Should -Be 2
        $script:Convert.newTaskAction | Should -BeNullOrEmpty
    }
}

Describe 'Console: the real window, a wave of a rollback, the stop file and the closing of the window' {
    BeforeAll {
        $script:Wave = Invoke-GuiChild -Body @'
$o = [ordered]@{}
Click 'BtnNext'; Check 'T_Recover'
$o.batches = @($g.BatchRows | ForEach-Object { '{0}|{1}|{2}' -f $_.BatchId, $_.State, $_.Selectable })
$o.readyNoBatch = [bool](Mod { Test-PcmGuiTaskReady })
$c.BatchList.SelectedItem = $g.BatchRows[0]; Pump 50
$o.envLock = $c.EnvLock.Visibility.ToString(); $o.fEnvEnabled = $c.F_Env.IsEnabled
$lock = [string](Mod { $script:Gui.EnvLock })
Mod { $script:Gui.EnvLock = 'dr-eu'; Set-PcmGuiField 'Env' 'dr-eu' }
$o.overrideLocked = [string](Mod { (Get-PcmGuiOverride).Environment })
$o.changedLocked = @(Mod { Get-PcmGuiChangedSetting })
Mod ([scriptblock]::Create("Set-PcmGuiField 'Env' '$lock'; `$script:Gui.EnvLock = '$lock'"))
$o.waveRows = $g.WaveRows.Count; $o.readyWhole = [bool](Mod { Test-PcmGuiTaskReady })
Check 'S_Wave_Wave'
$o.readyNoTick = [bool](Mod { Test-PcmGuiTaskReady })
foreach ($row in @($g.WaveRows | Select-Object -First 3)) { $row.Checked = $true }
Mod { Invoke-PcmGuiChange }
$o.count = $c.WaveCount.Text; $o.readyTicked = [bool](Mod { Test-PcmGuiTaskReady }); $o.happen = $c.HappenText.Text
Click 'BtnNext'
Click 'BtnPreview'; WaitRun
$a = Args_
$o.previewArgs = @{ Action = $a.Action; Mode = $a.Mode; Batch = $a.Batch; Csv = [bool]$a.CsvPath }
$o.waveCsv = @(Import-Csv -LiteralPath (Join-Path $root 'wave-copy.csv') | ForEach-Object { $_.Identity })
$o.waveFile = (Get-Content -LiteralPath (Join-Path $root 'wave-copy.csv') -TotalCount 1)
$o.confirmText = $c.ConfirmText.Text; $o.summary = $c.SumTitle.Text
Enter 'ConfirmBox' 'recover'
Scenario 2; Hold 'users'
Click 'BtnNext'
WaitFor { $g.Model.Objects.ContainsKey('s5') -and $g.Model.Objects['s5'].Count -ge 2 }
$o.stopEnabledBefore = $c.BtnStop.IsEnabled
Click 'BtnStop'
$o.stopFile = Test-Path -LiteralPath $g.Run.StopPath
$o.stopEnabledAfter = $c.BtnStop.IsEnabled; $o.pillStopping = $c.RunPillText.Text
$g.Form.Close(); Pump 100
$o.stillOpen = $g.Form.IsVisible; $o.closingMessage = ($g.Lines | Select-Object -Last 1).Text
$o.backEnabledRunning = $c.BtnBack.IsEnabled
Release 'users'
WaitRun
$o.level = $g.LastRun.Level; $o.stopped = $g.LastRun.Stopped; $o.exit = $g.LastRun.ExitCode
Click 'BtnNext'
$o.pill = $c.NextPillText.Text; $o.title = $c.NextTitle.Text; $o.continueVisible = $c.ContinueCmdBox.Visibility.ToString(); $o.continueCmd = $c.ContinueCmd.Text
$o.stepsList = @($g.NextRows | ForEach-Object { $_.Text })
Done $o
'@
    }
    It 'lists the Convert batches that can be rolled back, with their state, and asks for a batch first' {
        @($script:Wave.batches) | Should -Contain '1a2b3c4d|Can be rolled back|True'
        @($script:Wave.batches) | Should -Contain '9c8d7e6f|Partly rolled back|True'
        $script:Wave.readyNoBatch | Should -BeFalse
    }
    It 'locks the environment to the one of the batch and lists the mailboxes in the cloud for a wave' {
        $script:Wave.envLock | Should -Be 'Visible'
        $script:Wave.fEnvEnabled | Should -BeFalse
        $script:Wave.overrideLocked | Should -BeExactly 'dr-eu'
        $script:Wave.changedLocked | Should -Contain 'Environment'
        $script:Wave.waveRows | Should -Be 24
        $script:Wave.readyWhole | Should -BeTrue
        $script:Wave.readyNoTick | Should -BeFalse
        $script:Wave.count | Should -Be '3 ticked of 24 in the cloud'
        $script:Wave.readyTicked | Should -BeTrue
        $script:Wave.happen | Should -BeLike '*puts 3 ticked mailboxes back on-premises*'
    }
    It 'gives the ticked objects to the engine as a CSV file with -CsvPath, on the batch' {
        $script:Wave.previewArgs.Action | Should -Be 'Recover'
        $script:Wave.previewArgs.Batch | Should -Be '1a2b3c4d'
        $script:Wave.previewArgs.Csv | Should -BeTrue
        @($script:Wave.waveCsv) | Should -Be @('a0000000-0000-0000-0000-000000000001', 'a0000000-0000-0000-0000-000000000002', 'a0000000-0000-0000-0000-000000000003')
        $script:Wave.waveFile | Should -Match 'Identity,Name'
        $script:Wave.confirmText | Should -BeLike 'Type RECOVER*'
        $script:Wave.summary | Should -BeLike '3 mailboxes will be rolled back*'
    }
    It 'asks the run to stop after the current object with the stop file, and refuses to close the window while the run goes on' {
        $script:Wave.stopEnabledBefore | Should -BeTrue
        $script:Wave.stopFile | Should -BeTrue
        $script:Wave.stopEnabledAfter | Should -BeFalse
        $script:Wave.pillStopping | Should -Be 'Stopping'
        $script:Wave.stillOpen | Should -BeTrue
        $script:Wave.closingMessage | Should -BeLike 'This window cannot be closed while a run is going*'
        $script:Wave.backEnabledRunning | Should -BeFalse
    }
    It 'ends a stopped run normally, says it stopped as asked and gives the command to continue the batch' {
        $script:Wave.exit | Should -Be 2
        $script:Wave.level | Should -Be 'Pending'
        $script:Wave.stopped | Should -BeTrue
        $script:Wave.pill | Should -Be 'Stopped as asked'
        $script:Wave.continueVisible | Should -Be 'Visible'
        $script:Wave.continueCmd | Should -BeLike '*-Action Recover*-Batch 1a2b3c4d*'
    }
}

Describe 'Console: the real window on a computer that can only collect, Collect in PowerShell 7 with a console of its own' {
    BeforeAll {
        $script:Collect = Invoke-GuiChild -Kind Collect -Body @'
$o = [ordered]@{}
$o.role = $c.RoleName.Text
$o.settingsStore = $c.Row_Store.Visibility.ToString(); $o.settingsCloud = $c.CloudFields.Visibility.ToString()
Click 'BtnNext'
$o.tiles = @{ Collect = $c.T_Collect.IsEnabled; Convert = $c.T_Convert.IsEnabled; Recover = $c.T_Recover.IsEnabled; Check = $c.T_Check.IsEnabled }
$o.cloudPill = $c.NT_Check.Text
Check 'T_Collect'; Check 'C_Scope_UsersOnly'
$o.scopeField = Mod { Get-PcmGuiField 'Scope' }; $o.collectView = $c.CollectPanel.Visibility.ToString(); $o.sourceBox = $c.SourceBox.Visibility.ToString()
Click 'BtnNext'
$o.confirmText = $c.ConfirmText.Text
Click 'BtnPreview'; WaitRun
$a = Args_
$o.previewMode = $a.Mode; $o.previewScope = $a.Scope
Enter 'ConfirmBox' 'collect'
$o.runLabel = $c.BtnNext.Content; $o.runEnabled = $c.BtnNext.IsEnabled
Click 'BtnNext'; WaitRun; Click 'BtnNext'
$a = Args_
$o.applyMode = $a.Mode; $o.applyForce = $a.Force; $o.applyScope = $a.Scope; $o.applyBatch = [string]$a.Batch
$o.steps = @($g.Model.Steps | ForEach-Object { $_.Title })
$o.title = $c.NextTitle.Text; $o.batchBox = $c.BatchIdBox.Visibility.ToString()
Done $o
'@
    }
    It 'finds the role Collect and shows every setting: the snapshot file and the Microsoft 365 connection' {
        $script:Collect.role | Should -Be 'Collect'
        $script:Collect.settingsStore | Should -Be 'Visible'
        $script:Collect.settingsCloud | Should -Be 'Visible'
    }
    It 'greys the three cloud tiles with the reason, and offers Collect with its own scope' {
        $script:Collect.tiles.Collect | Should -BeTrue
        $script:Collect.tiles.Convert | Should -BeFalse
        $script:Collect.tiles.Recover | Should -BeFalse
        $script:Collect.tiles.Check | Should -BeFalse
        $script:Collect.cloudPill | Should -Be 'Needs the cloud modules and a sign-in'
        $script:Collect.scopeField | Should -Be 'UsersOnly'
        $script:Collect.collectView | Should -Be 'Visible'
        $script:Collect.sourceBox | Should -Be 'Collapsed'
    }
    It 'previews then writes the snapshot after the word COLLECT' {
        $script:Collect.confirmText | Should -BeLike 'Type COLLECT*'
        $script:Collect.previewMode | Should -Be 'Preview'
        $script:Collect.previewScope | Should -Be 'UsersOnly'
        $script:Collect.runLabel | Should -Be 'Collect the snapshot'
        $script:Collect.runEnabled | Should -BeTrue
        $script:Collect.applyMode | Should -Be 'Apply'
        $script:Collect.applyForce | Should -BeTrue
        $script:Collect.applyBatch | Should -BeNullOrEmpty
        @($script:Collect.steps).Count | Should -Be 6
        $script:Collect.title | Should -Be 'The snapshot is written'
        $script:Collect.batchBox | Should -Be 'Collapsed'
    }
}

Describe 'Console: the button that schedules the daily Collect' {
    BeforeAll {
        $script:Scheduled = Invoke-GuiChild -Kind Collect -Body @'
$o = [ordered]@{}
$global:probes = 0; $inner = $g.Probe; $g.Probe = { param($r) $global:probes++; & $inner $r }.GetNewClosure()
Click 'BtnNext'
$o.hiddenOnOtherTask = $c.BtnSchedule.IsVisible
Check 'T_Collect'
$o.visible = $c.BtnSchedule.IsVisible; $o.enabled = $c.BtnSchedule.IsEnabled; $o.label = [string]$c.BtnSchedule.Content
$before = $global:probes
Click 'BtnSchedule'; WaitRun
$a = Args_
$o.action = $a.Action; $o.mode = $a.Mode; $o.force = $a.Force; $o.batch = [string]$a.Batch; $o.identity = [string]$a.Identity
$o.question = @($questions)[-1]
$o.page = $g.Page
$o.runTitle = $c.RunTitle.Text
$o.steps = @($g.Model.Steps | ForEach-Object { $_.EngineTitle })
$o.refreshed = ($global:probes - $before)
$o.requestAction = $g.LastRun.Request.Action
$o.valid = $g.LastRun.Valid
Click 'BtnNext'
$o.page2 = $g.Page; $o.title = $c.NextTitle.Text; $o.lead = $c.NextLead.Text
$o.next = @($g.NextRows | ForEach-Object { $_.Text })
$o.nextCmd = $c.NextCmd.Text
Done $o
'@
    }
    It 'shows the button on the Collect page only, and enabled when this computer can collect' {
        $script:Scheduled.hiddenOnOtherTask | Should -BeFalse
        $script:Scheduled.visible | Should -BeTrue
        $script:Scheduled.enabled | Should -BeTrue
        $script:Scheduled.label | Should -Be 'Schedule the daily Collect'
    }
    It 'asks one question, then runs the action Schedule in Apply mode like the others, with no objects and no preview' {
        $script:Scheduled.question | Should -BeLike 'Schedule the daily Collect on this computer?*asks for the account that runs it*View-Only Organization Management*started as administrator*'
        $script:Scheduled.action | Should -Be 'Schedule'
        $script:Scheduled.mode | Should -Be 'Apply'
        $script:Scheduled.force | Should -BeTrue
        $script:Scheduled.batch | Should -BeNullOrEmpty
        $script:Scheduled.identity | Should -BeNullOrEmpty
        $script:Scheduled.runTitle | Should -Be 'Schedule the daily Collect'
        @($script:Scheduled.steps) | Should -Be @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task')
    }
    It 'looks at the lights again when the run is over, and says what to do next' {
        $script:Scheduled.refreshed | Should -BeGreaterOrEqual 1
        $script:Scheduled.requestAction | Should -Be 'Schedule'
        $script:Scheduled.valid | Should -BeTrue
        $script:Scheduled.page2 | Should -Be 5
        $script:Scheduled.title | Should -Be 'The daily Collect is scheduled'
        ($script:Scheduled.next -join ' ') | Should -BeLike '*Start-ScheduledTask -TaskName ''PraCloudMailbox Collect''*Daily Collect*'
        $script:Scheduled.nextCmd | Should -BeLike '*Invoke-PraCloudMailbox.ps1 -Action Check*'
    }
}

Describe 'Console: the real window when Exchange cannot be reached from this computer' {
    BeforeAll {
        $script:Down = Invoke-GuiChild -Kind ExchangeDown -Body @'
$o = [ordered]@{}
$o.role = $c.RoleName.Text; $o.roleText = $c.RoleText.Text
$o.exchangeLight = @($g.Detect.Lights | Where-Object { $_.Name -eq 'Exchange reachable' } | ForEach-Object { '{0}|{1}' -f $_.Status, $_.Detail })
Click 'BtnNext'
$o.tiles = @{ Collect = $c.T_Collect.IsEnabled; Convert = $c.T_Convert.IsEnabled; Check = $c.T_Check.IsEnabled }
$o.collectPill = $c.NT_Collect.Text
Check 'T_Collect'
$o.blocker = Mod { Get-PcmGuiBlocker }
Done $o
'@
        $script:Up = Invoke-GuiChild -Kind Cloud -Body @'
$o = [ordered]@{}
$o.role = $c.RoleName.Text
$o.exchangeLight = @($g.Detect.Lights | Where-Object { $_.Name -eq 'Exchange reachable' } | ForEach-Object { '{0}|{1}' -f $_.Status, $_.Detail })
Click 'BtnNext'
$o.tiles = @{ Collect = $c.T_Collect.IsEnabled; Convert = $c.T_Convert.IsEnabled; Check = $c.T_Check.IsEnabled }
Done $o
'@
    }
    It 'lights Exchange reachable amber, keeps the cloud actions and greys Collect with the reason' {
        $script:Down.role | Should -Be 'Cloud actions'
        $script:Down.roleText | Should -BeLike '*Collect needs Exchange.Server to be set and reachable*'
        @($script:Down.exchangeLight)[0] | Should -BeLike 'Warn|A remote session to ex01.contoso.com cannot be opened*The server did not answer.*'
        $script:Down.tiles.Collect | Should -BeFalse
        $script:Down.tiles.Convert | Should -BeTrue
        $script:Down.tiles.Check | Should -BeTrue
        $script:Down.collectPill | Should -Be 'Exchange not reachable'
        $script:Down.blocker | Should -BeLike 'This computer cannot run Collect: A remote session to ex01.contoso.com cannot be opened*'
    }
    It 'lights Exchange reachable green and offers every action when the remote session opens' {
        $script:Up.role | Should -Be 'Everything'
        @($script:Up.exchangeLight)[0] | Should -BeLike 'Ok|Exchange answers on ex01.contoso.com*'
        $script:Up.tiles.Collect | Should -BeTrue
        $script:Up.tiles.Convert | Should -BeTrue
    }
}

Describe 'Console: the real window, a follow-up check, the other sources, a failure and the theme' {
    BeforeAll {
        $script:Other = Invoke-GuiChild -Body @'
$o = [ordered]@{}
Click 'BtnNext'; Check 'T_Check'
$o.modes = @($c.Mode_New.Content, $c.Mode_Continue.Content)
Check 'Mode_Continue'
$o.followBatches = $g.BatchRows.Count
$c.BatchList.SelectedItem = ($g.BatchRows | Where-Object { $_.BatchId -eq '1a2b3c4d' } | Select-Object -First 1); Pump 50
$o.expectVisible = $c.CheckPanel.Visibility.ToString()
Check 'S_Expect_Deprovisioned'; $c.F_Once.IsChecked = $true; Pump 50
$o.happen = $c.HappenText.Text
Click 'BtnNext'
$o.verifySummary = $c.SumTitle.Text; $o.previewVisible = $c.BtnPreview.Visibility.ToString(); $o.runLabel = $c.BtnNext.Content; $o.runEnabled = $c.BtnNext.IsEnabled
Click 'BtnNext'
WaitRun
Click 'BtnNext'
$a = Args_
$o.checkArgs = @{ Action = $a.Action; Batch = $a.Batch; Expect = $a.Expect; Once = $a.Once; Force = $a.Force }
$o.checkNext = $c.NextTitle.Text
Click 'BtnNext'
# The other ways to pick the objects of a new conversion.
Check 'T_Convert'; Check 'S_Source_Single'
Enter 'F_Identity' 'ADELE.VANCE@CONTOSO.COM'; Mod { Update-PcmGuiView }
$o.single = @{ Info = $c.IdentityInfo.Text; Ready = [bool](Mod { Test-PcmGuiTaskReady }); Identity = (Mod { (Get-PcmGuiWindowRequest).Identity }) }
Enter 'F_Identity' 'nobody@contoso.com'; Mod { Update-PcmGuiView }
$o.unknown = @{ Info = $c.IdentityInfo.Text; Ready = [bool](Mod { Test-PcmGuiTaskReady }) }
Check 'S_Source_Ou'
Click 'BtnPickOu'
$o.ou = @{ Picked = $c.F_Ou.Text; Info = $c.OuInfo.Text; SearchBase = (Mod { (Get-PcmGuiWindowRequest).SearchBase }) }
Enter 'F_Ou' 'OU=Lyon,DC=contoso,DC=com'; Mod { Update-PcmGuiView }
$o.ouDn = $c.OuInfo.Text
Check 'S_Source_Csv'
Enter 'F_Csv' (Join-Path $root 'missing.csv'); Mod { Update-PcmGuiView }
$o.csvMissing = @{ Info = $c.CsvInfo.Text; Ready = [bool](Mod { Test-PcmGuiTaskReady }) }
$list = Join-Path $root 'list.csv'; [IO.File]::WriteAllText($list, "identity;note`r`nadele.vance@contoso.com;a`r`nalex.wilber@contoso.com;b`r`n")
Enter 'F_Csv' $list
$o.csv = @{ Info = $c.CsvInfo.Text; Ready = [bool](Mod { Test-PcmGuiTaskReady }); CsvPath = (Mod { (Get-PcmGuiWindowRequest).CsvPath }) }
Check 'S_Source_All'
$o.all = @{ Text = $c.Row_All.Text; Ready = [bool](Mod { Test-PcmGuiTaskReady }) }
# A run that fails: the engine cannot sign in.
[IO.File]::WriteAllText((Join-Path $root 'fail.flag'), '1')
Check 'S_Source_Single'; Enter 'F_Identity' 'adele.vance@contoso.com'
Click 'BtnNext'; Click 'BtnPreview'; WaitRun
$o.failedPreview = @{ Previewed = $g.Previewed; Error = $g.PreviewError; Summary = $c.SumTitle.Text; Confirm = $c.ConfirmBox.IsEnabled }
Mod { Show-PcmGuiPage 2 }
Check 'T_Check'; Check 'S_Source_All'
Mod { Show-PcmGuiPage 3 }
Click 'BtnNext'; WaitRun; Click 'BtnNext'
$o.failed = @{ Valid = $g.LastRun.Valid; Level = $g.LastRun.Level; Pill = $c.NextPillText.Text; Title = $c.NextTitle.Text; Lead = $c.NextLead.Text; Steps = @($g.NextRows | ForEach-Object { $_.Text }); Status = @($g.Model.Steps | ForEach-Object { $_.Status }) }
# Interactive sign-in and a chosen snapshot, for this run only.
Remove-Item (Join-Path $root 'fail.flag') -ErrorAction SilentlyContinue
Mod { Show-PcmGuiPage 2 }
Check 'T_Check'; Check 'S_Source_All'
$c.SettingsCard.Visibility = 'Visible'
Check 'S_Signin_Interactive'; Enter 'F_Upn' 'admin@contoso.com'; Check 'S_Snap_Specific'; Enter 'F_SnapId' '41'
$o.interactive = @{ Override = (Mod { (Get-PcmGuiOverride) | ConvertTo-Json -Compress -Depth 4 }); RowAppId = $c.Row_AppId.Visibility.ToString(); RowUpn = $c.Row_Upn.Visibility.ToString(); Changed = @(Mod { Get-PcmGuiChangedSetting }); Number = (Mod { Get-PcmGuiSnapshotNumber }) }
Enter 'F_SnapId' 'abc'
$o.badSnapshot = $c.SettingsError.Text; $o.badNext = $c.BtnNext.IsEnabled
Enter 'F_SnapId' '41'
Mod { Show-PcmGuiPage 3 }
Click 'BtnNext'; WaitRun
$a = Args_
$o.interactiveArgs = @{ Snapshot = $a.Snapshot; Base = [bool]$a.BaseConfigPath }
$o.effective = (Get-Content -LiteralPath (Join-Path $root 'effective-copy.psd1') -Raw)
$o.configUnchanged = ((Get-FileHash -LiteralPath $fixture.Config -Algorithm SHA256).Hash -eq $configHash)
# The theme.
$light = $g.Form.TryFindResource('PraBrandText').Color.ToString()
Check 'Th_Dark'; $dark = $g.Form.TryFindResource('PraBrandText').Color.ToString(); Check 'Th_Light'
$o.theme = @{ Light = $light; Dark = $dark }
Done $o
'@
    }
    It 'follows a batch with the expectation and the one-pass choice, without a preview' {
        @($script:Other.modes) | Should -Be @('Check the snapshot', 'Follow a batch')
        $script:Other.followBatches | Should -Be 4
        $script:Other.expectVisible | Should -Be 'Visible'
        $script:Other.happen | Should -BeLike 'This computer follows batch 1a2b3c4d in Microsoft 365 and looks once. It changes nothing.'
        $script:Other.previewVisible | Should -Be 'Collapsed'
        $script:Other.runLabel | Should -Be 'Run the check'
        $script:Other.runEnabled | Should -BeTrue
        $script:Other.checkArgs.Action | Should -Be 'Check'
        $script:Other.checkArgs.Batch | Should -Be '1a2b3c4d'
        $script:Other.checkArgs.Expect | Should -Be 'Deprovisioned'
        $script:Other.checkArgs.Once | Should -BeTrue
        $script:Other.checkArgs.Force | Should -BeFalse
        $script:Other.checkNext | Should -Be 'The check is finished'
    }
    It 'picks one mailbox from the snapshot in any letter case, and refuses one that is not there' {
        $script:Other.single.Info | Should -BeLike 'Adele Vance * user mailbox'
        $script:Other.single.Ready | Should -BeTrue
        $script:Other.single.Identity | Should -Be 'ADELE.VANCE@CONTOSO.COM'
        $script:Other.unknown.Info | Should -BeLike 'Not found in snapshot #42*'
        $script:Other.unknown.Ready | Should -BeFalse
    }
    It 'picks an OU from the snapshot, by path or by distinguished name, and a CSV file with its identities' {
        $script:Other.ou.Picked | Should -Be 'contoso.com/Lyon'
        $script:Other.ou.SearchBase | Should -Be $script:Other.ou.Picked
        $script:Other.ou.Info | Should -Be '8 mailboxes in snapshot #42, sub-OUs included.'
        $script:Other.ouDn | Should -BeLike '8 mailboxes in snapshot #42*'
        $script:Other.csvMissing.Ready | Should -BeFalse
        $script:Other.csvMissing.Info | Should -BeLike 'This file cannot be read.'
        $script:Other.csv.Info | Should -Be '2 identities'
        $script:Other.csv.Ready | Should -BeTrue
        $script:Other.csv.CsvPath | Should -BeLike '*list.csv'
        $script:Other.all.Text | Should -BeLike 'Every mailbox of snapshot #42 (24 mailboxes)*'
        $script:Other.all.Ready | Should -BeTrue
    }
    It 'tells a failed run plainly: the first reason, the failed step and what to do' {
        $script:Other.failedPreview.Previewed | Should -BeFalse
        $script:Other.failedPreview.Summary | Should -Be 'The preview did not complete'
        $script:Other.failedPreview.Error | Should -BeLike '*not found in Cert:*'
        $script:Other.failedPreview.Confirm | Should -BeFalse
        $script:Other.failed.Valid | Should -BeFalse
        $script:Other.failed.Level | Should -Be 'Failed'
        $script:Other.failed.Pill | Should -Be 'Failed'
        $script:Other.failed.Title | Should -Be 'The run stopped'
        $script:Other.failed.Lead | Should -BeLike '*not found in Cert:*'
        @($script:Other.failed.Status) | Should -Contain 'Fail'
        $script:Other.failed.Steps[0] | Should -BeLike 'Go back to the Run page*'
    }
    It 'runs with the interactive sign-in and a chosen snapshot for this run only, and checks what was typed' {
        $script:Other.interactive.Override | Should -BeLike '*"SignIn":"Interactive"*'
        $script:Other.interactive.Override | Should -BeLike '*"UserPrincipalName":"admin@contoso.com"*'
        $script:Other.interactive.Override | Should -Not -BeLike '*AppId*'
        $script:Other.interactive.RowAppId | Should -Be 'Collapsed'
        $script:Other.interactive.RowUpn | Should -Be 'Visible'
        @($script:Other.interactive.Changed) | Should -Be @('Snapshot', 'Sign-in', 'Administrator')
        $script:Other.interactive.Number | Should -Be 41
        $script:Other.badSnapshot | Should -BeLike 'Snapshot: type the number*'
        $script:Other.badNext | Should -BeFalse
        $script:Other.interactiveArgs.Snapshot | Should -Be 41
        $script:Other.interactiveArgs.Base | Should -BeTrue
        $script:Other.effective | Should -Match "SignIn = 'Interactive'"
        $script:Other.configUnchanged | Should -BeTrue
    }
    It 'switches between the light and the dark palette without rebuilding the window' {
        $script:Other.theme.Light | Should -Be '#FFB11F4B'
        $script:Other.theme.Dark | Should -Be '#FFFD8EA1'
    }
}

Describe 'Console: the numeric settings of the waves (temporary licences, licences kept free, objects per rollback wave)' {
    BeforeAll {
        $script:Numbers = Invoke-GuiChild -Body @'
$o = [ordered]@{}
$o.defaults = @{ Parallel = $c.F_Parallel.Text; KeepFree = $c.F_KeepFree.Text; WaveSize = $c.F_WaveSize.Text }
$o.rowsBefore = @{ Parallel = $c.Row_Parallel.Visibility.ToString(); KeepFree = $c.Row_KeepFree.Visibility.ToString(); WaveSize = $c.Row_WaveSize.Visibility.ToString() }
$o.problemOf = @{}
foreach ($pair in @(@('Parallel', ' 40 '), @('Parallel', '0'), @('Parallel', '1000'), @('Parallel', '1001'), @('Parallel', ''), @('Parallel', '4 0'), @('Parallel', '4.5'), @('Parallel', '-3'), @('KeepFree', '100000'), @('KeepFree', '100001'), @('KeepFree', 'ten'), @('KeepFree', ' 0 '))) {
    Enter ('F_' + $pair[0]) $pair[1]
    $o.problemOf[('{0}|{1}' -f $pair[0], $pair[1])] = $c.SettingsError.Text
    Mod { Reset-PcmGuiField 'Parallel'; Reset-PcmGuiField 'KeepFree' }
}
Enter 'F_Parallel' ' 40 '; Enter 'F_KeepFree' ' 10 '
$o.modified = @{ Parallel = $c.M_Parallel.Visibility.ToString(); KeepFree = $c.M_KeepFree.Visibility.ToString() }
$o.changed = @(Mod { Get-PcmGuiChangedSetting })
$o.override = Mod { (Get-PcmGuiOverride | ConvertTo-Json -Compress -Depth 5) }
$o.overrideIsInt = (Mod { (Get-PcmGuiOverride).Licensing.Shared.Parallel -is [int] })
Enter 'F_Parallel' '100'
$o.sameNumberNotModified = @{ Marker = $c.M_Parallel.Visibility.ToString(); Changed = @(Mod { Get-PcmGuiChangedSetting }) }
Enter 'F_Parallel' ' 40 '
Click 'BtnNext'; Check 'T_Convert'; Check 'S_Source_All'
$o.convertText = $c.HappenText.Text
$o.waveRowConvert = $c.Row_WaveSize.Visibility.ToString()
Click 'BtnNext'
$o.used = $c.UsedText.Text
Click 'BtnPreview'; WaitRun
$o.effective = (Get-Content -LiteralPath (Join-Path $root 'effective-copy.psd1') -Raw)
$o.configUnchanged = ((Get-FileHash -LiteralPath $fixture.Config -Algorithm SHA256).Hash -eq $configHash)
$o.original = [IO.File]::ReadAllText($fixture.Config)
Mod { Show-PcmGuiPage 2 }
Check 'T_Recover'
$o.waveRowRecover = $c.Row_WaveSize.Visibility.ToString()
$o.recoverDefaultText = $c.HappenText.Text
$o.waveProblems = @{}
foreach ($text in @('0', '1001', 'x', ' 25 ', '')) { Enter 'F_WaveSize' $text; $o.waveProblems[$(if ($text) { $text } else { 'empty' })] = $c.SettingsError.Text }
Enter 'F_WaveSize' ' 25 '
$o.recoverText = $c.HappenText.Text
$o.recoverOverride = Mod { (Get-PcmGuiOverride | ConvertTo-Json -Compress -Depth 5) }
$o.changedRecover = @(Mod { Get-PcmGuiChangedSetting })
Click 'BtnResetAll'
$o.afterReset = @{ Parallel = $c.F_Parallel.Text; KeepFree = $c.F_KeepFree.Text; WaveSize = $c.F_WaveSize.Text }
Done $o
'@
    }
    It 'shows the values of the configuration, and the wave size only for a rollback' {
        $script:Numbers.defaults.Parallel | Should -Be '100'
        $script:Numbers.defaults.KeepFree | Should -Be '0'
        $script:Numbers.defaults.WaveSize | Should -Be ''
        $script:Numbers.rowsBefore.Parallel | Should -Be 'Visible'
        $script:Numbers.rowsBefore.KeepFree | Should -Be 'Visible'
        $script:Numbers.rowsBefore.WaveSize | Should -Be 'Collapsed'
        $script:Numbers.waveRowConvert | Should -Be 'Collapsed'
        $script:Numbers.waveRowRecover | Should -Be 'Visible'
    }
    It 'accepts whole numbers in their range, spaces around them ignored, and refuses the rest in plain words' {
        $problems = $script:Numbers.problemOf
        foreach ($ok in 'Parallel| 40 ', 'Parallel|1000', 'KeepFree|100000', 'KeepFree| 0 ') { $problems.$ok | Should -BeNullOrEmpty -Because $ok }
        foreach ($bad in 'Parallel|0', 'Parallel|1001', 'Parallel|', 'Parallel|4 0', 'Parallel|4.5', 'Parallel|-3') { $problems.$bad | Should -Be 'Temporary licences: a whole number from 1 to 1000.' -Because $bad }
        foreach ($bad in 'KeepFree|100001', 'KeepFree|ten') { $problems.$bad | Should -Be 'Licences kept free: a whole number from 0 to 100000.' -Because $bad }
    }
    It 'marks what changed and writes it, as numbers, to the settings of this run only' {
        $script:Numbers.modified.Parallel | Should -Be 'Visible'
        $script:Numbers.modified.KeepFree | Should -Be 'Visible'
        @($script:Numbers.changed) | Should -Be @('Temporary licences', 'Licences kept free')
        $script:Numbers.override | Should -Match '^\{"Licensing":\{"Shared":\{.*\}\}\}$'
        $script:Numbers.override | Should -Match '"Parallel":40'
        $script:Numbers.override | Should -Match '"KeepFree":10'
        $script:Numbers.overrideIsInt | Should -BeTrue
        $script:Numbers.sameNumberNotModified.Marker | Should -Be 'Collapsed'
        @($script:Numbers.sameNumberNotModified.Changed) | Should -Be @('Licences kept free')
    }
    It 'never writes them to the configuration file, only to the generated copy' {
        $script:Numbers.configUnchanged | Should -BeTrue
        $script:Numbers.original | Should -Not -Match 'KeepFree'
        $script:Numbers.original | Should -Not -Match 'Parallel'
        $script:Numbers.effective | Should -Match 'Parallel = 40'
        $script:Numbers.effective | Should -Match 'KeepFree = 10'
        $script:Numbers.effective | Should -Match "SkuPartNumber = 'EXCHANGESTANDARD'"
    }
    It 'says the waves in the card of what will happen: temporary licences, the plan and the licences kept free' {
        $script:Numbers.convertText | Should -BeLike '*Shared mailboxes in waves of up to 40 temporary licences (EXCHANGESTANDARD), 10 kept free.'
        $script:Numbers.used | Should -BeLike '*Shared mailbox waves: 40 temporary licences, 10 kept free*'
    }
    It 'validates the wave size of a rollback (empty = the temporary licences) and says it in the card' {
        $script:Numbers.waveProblems.'0' | Should -Be 'Wave size: a whole number from 1 to 1000.'
        $script:Numbers.waveProblems.'1001' | Should -Be 'Wave size: a whole number from 1 to 1000.'
        $script:Numbers.waveProblems.x | Should -Be 'Wave size: a whole number from 1 to 1000.'
        $script:Numbers.waveProblems.' 25 ' | Should -BeNullOrEmpty
        $script:Numbers.waveProblems.empty | Should -BeNullOrEmpty
        $script:Numbers.recoverDefaultText | Should -BeLike '*Rollback in waves of up to 40 objects.'
        $script:Numbers.recoverText | Should -BeLike '*Rollback in waves of up to 25 objects.'
        $script:Numbers.recoverOverride | Should -BeLike '*"Recover":{"WaveSize":25}*'
        @($script:Numbers.changedRecover) | Should -Contain 'Wave size'
    }
    It 'puts every number back to the value of the file with Reset everything' {
        $script:Numbers.afterReset.Parallel | Should -Be '100'
        $script:Numbers.afterReset.KeepFree | Should -Be '0'
        $script:Numbers.afterReset.WaveSize | Should -Be ''
    }
}

Describe 'Console: the group check asks the engine' {
    BeforeAll {
        $script:Lookup = Invoke-GuiChild -Body @'
$o = [ordered]@{}
Click 'BtnNext'
Check 'T_Convert'; Check 'S_Source_Group'
Enter 'F_Group' 'pra-WAVE1'; Click 'BtnCheckGroup'
$o.runningText = $c.GroupInfo.Text; $o.buttonWhileRunning = $c.BtnCheckGroup.IsEnabled; $o.lookupWhileRunning = [bool]$g.GroupLookup
$o.startRefused = ''
try { Mod { Start-PcmGuiRun -Mode Preview } } catch { $o.startRefused = $_.Exception.Message }
WaitFor { -not $g.GroupLookup }
$a = Args_
$o.args = @{ Action = $a.Action; Group = $a.Group; ResolveOnly = $a.ResolveOnly; Mode = $a.Mode; Force = $a.Force; Edition = $a.Edition }
$o.found = $c.GroupInfo.Text; $o.buttonAfter = $c.BtnCheckGroup.IsEnabled; $o.glyph = $c.GroupGlyph.Visibility.ToString()
Enter 'F_Group' 'Nobody'; Click 'BtnCheckGroup'; WaitFor { -not $g.GroupLookup }
$o.missing = $c.GroupInfo.Text; $o.missingGlyph = $c.GroupGlyph.Visibility.ToString()
$g.NetResult = [pscustomobject]@{ State = 'Fail'; Answers = @() }
Enter 'F_Group' 'PRA-WAVE1'; Click 'BtnCheckGroup'
$o.fallbackFound = $c.GroupInfo.Text; $o.fallbackLookup = [bool]$g.GroupLookup
Enter 'F_Group' 'Nobody'; Click 'BtnCheckGroup'
$o.fallbackMissing = $c.GroupInfo.Text
Done $o
'@
    }
    It 'runs the engine in a child process (Check -Group, with -ResolveOnly) and shows its answer, the group found in the snapshot' {
        $script:Lookup.args.Action | Should -Be 'Check'
        $script:Lookup.args.Group | Should -Be 'pra-WAVE1'
        $script:Lookup.args.ResolveOnly | Should -BeTrue
        $script:Lookup.args.Force | Should -BeFalse
        $script:Lookup.found | Should -Be 'Found in the snapshot (captured by Collect): PRA-Wave1 · 26 members including nested groups, 2 skipped (the Preview says why)'
        $script:Lookup.glyph | Should -Be 'Visible'
    }
    It 'says that it is looking, locks the button and refuses a run while the engine answers' {
        $script:Lookup.runningText | Should -BeLike 'Looking up the group*'
        $script:Lookup.buttonWhileRunning | Should -BeFalse
        $script:Lookup.lookupWhileRunning | Should -BeTrue
        $script:Lookup.startRefused | Should -BeLike '*group check is still going*'
        $script:Lookup.buttonAfter | Should -BeTrue
    }
    It 'shows the reason of the engine when the group is nowhere' {
        $script:Lookup.missing | Should -Be 'Could not check: Group Nobody is not in snapshot 42 nor in Microsoft Entra. Nothing was changed: check the name.'
        $script:Lookup.missingGlyph | Should -Be 'Collapsed'
    }
    It 'falls back to the snapshot alone, and says so, when Microsoft 365 cannot be reached from this computer' {
        $script:Lookup.fallbackLookup | Should -BeFalse
        $script:Lookup.fallbackFound | Should -BeLike 'Found in the snapshot (captured by Collect): PRA-Wave1 · 26 members including nested groups, 2 skipped (the Preview says why). Only the snapshot was checked: *'
        $script:Lookup.fallbackMissing | Should -BeLike 'Not in the snapshot. Microsoft Entra was not checked: *'
    }
}

Describe 'Console: the engine behind the console (the real script, a synthetic snapshot)' {
    BeforeAll {
        $script:Engine = New-DocumentationToolCopy -Repository $script:Repo -Destination (Join-Path $TestDrive 'engine') -RealEngine
        $script:EngineFixture = New-ConsoleFixture -ToolRoot $script:Engine -NoScenario
        # The engine checks that the cloud modules are installed before it reads anything: empty modules at the minimum versions stand in for them, in a folder of their own.
        $script:CloudModules = Join-Path $TestDrive 'cloud-modules'
        foreach ($module in @{ Name = 'ExchangeOnlineManagement'; Version = '3.10.1' }, @{ Name = 'Microsoft.Graph.Authentication'; Version = '2.41.1' }) {
            $folder = Join-Path $script:CloudModules "$($module.Name)\$($module.Version)"
            $null = New-Item -ItemType Directory -Path $folder -Force
            New-ModuleManifest -Path (Join-Path $folder "$($module.Name).psd1") -ModuleVersion $module.Version
        }
        function script:Invoke-RealEngine {
            param([string[]]$Arguments)
            $events = Join-Path $TestDrive ('events-{0}.jsonl' -f [guid]::NewGuid().ToString('N').Substring(0, 8))
            $modulePath = $env:PSModulePath
            $env:PRA_EVENT_FILE = $events
            $env:PSModulePath = $script:CloudModules + [IO.Path]::PathSeparator + $modulePath
            try { $output = (& pwsh -NoProfile -File (Join-Path $script:Engine 'Invoke-PraCloudMailbox.ps1') @Arguments -ConfigPath $script:EngineFixture.Config 2>&1 | Out-String) } finally { $env:PSModulePath = $modulePath; Remove-Item Env:\PRA_EVENT_FILE -ErrorAction SilentlyContinue }
            return [pscustomobject]@{ Exit = $LASTEXITCODE; Output = $output; Events = @($(if (Test-Path -LiteralPath $events) { Get-Content -LiteralPath $events -Encoding UTF8 | ForEach-Object { $_ | ConvertFrom-Json } })) }
        }
        $script:HashBefore = @(Get-ChildItem -LiteralPath (Join-Path $script:Engine 'data') -Filter '*.db' | Sort-Object Name | ForEach-Object { '{0}={1}' -f $_.Name, (Get-FileHash -LiteralPath $_.FullName).Hash })
        $script:Resolved = Invoke-RealEngine -Arguments @('-Action', 'Check', '-Group', 'pra-WAVE1', '-ResolveOnly')
    }
    It 'looks a group of the snapshot up (any letter case), prints the line of the run and ends with exit code 0' {
        $script:Resolved.Exit | Should -Be 0 -Because $script:Resolved.Output
        $script:Resolved.Output | Should -Match 'Group PRA-Wave1 \(captured by Collect, snapshot 42\): 26 member\(s\) including nested groups, 2 skipped'
    }
    It 'tells the console which group, from where, with how many members and how many skipped (event target)' {
        $target = @($script:Resolved.Events | Where-Object kind -eq 'target')
        $target.Count | Should -Be 1
        $target[0].name | Should -Be 'PRA-Wave1'
        $target[0].source | Should -Be 'Snapshot'
        $target[0].members | Should -Be 26
        $target[0].skipped | Should -Be 2
        $target[0].count | Should -Be 24
        $target[0].text | Should -BeLike 'Group PRA-Wave1 (captured by Collect, snapshot 42): 26 member(s) including nested groups, 2 skipped'
    }
    It 'announces its plan, numbers its steps with the ids of the plan, ends them with a state and ends with a result' {
        $kinds = @($script:Resolved.Events | ForEach-Object { $_.kind })
        $kinds[0] | Should -Be 'start'
        $kinds[1] | Should -Be 'plan'
        $plan = $script:Resolved.Events[1]
        @($plan.steps | ForEach-Object { '{0}={1}' -f $_.id, $_.label }) | Should -Be @('s1=Collect snapshot')
        $step = @($script:Resolved.Events | Where-Object { $_.kind -eq 'step' -and $_.PSObject.Properties['title'] })
        $step.Count | Should -Be 1
        $step[0].id | Should -Be 's1'
        @($script:Resolved.Events | Where-Object { $_.kind -eq 'step' -and $_.PSObject.Properties['state'] } | ForEach-Object { '{0}:{1}' -f $_.id, $_.state }) | Should -Be @('s1:ok')
        $started = [array]::IndexOf(@($script:Resolved.Events | ForEach-Object { $_.kind }), 'step')
        $items = @($script:Resolved.Events | Select-Object -Skip $started | Where-Object kind -eq 'item')
        $items.Count | Should -BeGreaterThan 0
        @($items | Where-Object { $_.stepId -ne 's1' }).Count | Should -Be 0
        $result = @($script:Resolved.Events | Where-Object kind -eq 'result')
        $result.Count | Should -Be 1
        $result[0].exitCode | Should -Be 0
        $result[0].csvReport | Should -BeNullOrEmpty
        $result[0].PSObject.Properties.Name | Should -Contain 'globalReport'
    }
    It 'changes nothing: no report, no database written, no cloud sign-in' {
        @(Get-ChildItem -LiteralPath (Join-Path $script:Engine 'reports') -File -ErrorAction SilentlyContinue).Count | Should -Be 0
        $after = @(Get-ChildItem -LiteralPath (Join-Path $script:Engine 'data') -Filter '*.db' | Sort-Object Name | ForEach-Object { '{0}={1}' -f $_.Name, (Get-FileHash -LiteralPath $_.FullName).Hash })
        $after | Should -Be $script:HashBefore
        $script:Resolved.Output | Should -Not -Match 'Microsoft Graph'
    }
    It 'refuses -ResolveOnly without a group, and for another action, with exit code 1' -TestCases @(
        @{ Arguments = @('-Action', 'Check', '-ResolveOnly') }, @{ Arguments = @('-Action', 'Convert', '-Group', 'PRA-Wave1', '-ResolveOnly') }) {
        param($Arguments)
        $run = Invoke-RealEngine -Arguments $Arguments
        $run.Exit | Should -Be 1
        $run.Output | Should -BeLike '*-ResolveOnly is for Check with -Group*'
    }
    It 'follows a rollback by the Convert batch ID, and says that a rollback batch ID is the rollback of its Convert batch' {
        $journal = Open-PcmStore -Path (Join-Path $script:Engine 'data\PraCloudMailbox-journal.db') -Root $script:Engine -Kind Journal
        try {
            $null = Invoke-PcmSql $journal 'INSERT INTO batch (id, action, convert_batch, status, created_utc, updated_utc, snapshot_id, environment, computer, account, tool_version) VALUES (@id, @a, @c, @s, @t, @t, @n, @e, @m, @u, @v)' @{ id = 'ee11ee11'; a = 'Recover'; c = '1a2b3c4d'; s = 'Partial'; t = [datetime]::UtcNow; n = 42; e = 'PROD'; m = 'SRV-PRA01'; u = 'CONTOSO\pra.operator'; v = '2.0.0' }
        } finally { Close-PcmStore $journal }
        $run = Invoke-RealEngine -Arguments @('-Action', 'Recover', '-Batch', 'EE11EE11')
        $run.Output | Should -Match 'Batch ee11ee11 is a rollback of Convert batch 1a2b3c4d: the rollback goes on with Convert batch 1a2b3c4d'
        $run.Output | Should -Not -Match 'Convert batch EE11EE11 not found'
    }
}

Describe 'Console: the window files' {
    It 'keeps the layout and the code free of the old editions and of -NonInteractive' {
        foreach ($name in 'PRA.Gui.psm1', 'PRA.Gui.View.ps1', 'PRA.Gui.Run.ps1', 'PRA.Gui.Data.ps1', 'PRA.Gui.Detect.ps1', 'PRA.Gui.Process.ps1', 'PRA.Gui.RunConfig.ps1', 'PRA.Gui.xaml') {
            $text = [IO.File]::ReadAllText((Join-Path $script:Root "module\$name"))
            $text | Should -Not -Match '7\.4' -Because $name
            if ($name.EndsWith('.xaml')) { continue }
            $tokens = $null; $null = [Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$null)
            @($tokens | Where-Object { $_.Kind -in 'Parameter', 'StringLiteral', 'StringExpandable', 'Generic' -and $_.Text -match 'NonInteractive' }).Count | Should -Be 0 -Because $name
        }
    }
    It 'parses without error' {
        foreach ($file in Get-ChildItem (Join-Path $script:Root 'module') -Filter 'PRA.Gui*' | Where-Object { $_.Extension -in '.ps1', '.psm1' }) {
            $errors = $null; $null = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors)
            @($errors).Count | Should -Be 0 -Because $file.Name
        }
    }
    It 'has no trace of the classic controls, of Windows PowerShell 5.1 or of a local Exchange mode' {
        foreach ($name in 'PRA.Gui.psm1', 'PRA.Gui.View.ps1', 'PRA.Gui.Run.ps1', 'PRA.Gui.Data.ps1', 'PRA.Gui.Detect.ps1', 'PRA.Gui.Process.ps1', 'PRA.Gui.RunConfig.ps1', 'PRA.Gui.xaml') {
            $text = [IO.File]::ReadAllText((Join-Path $script:Root "module\$name"))
            $text | Should -Not -Match 'classic controls|ConnectionMode|Exchange tools|\.Fluent\b' -Because $name
            $text | Should -Not -Match 'WindowsPowerShell\\v1\.0' -Because $name
        }
    }
}

Describe 'Console: continuing a rollback stays on the wave that was interrupted' {
    BeforeAll {
        function script:New-TestBatch {
            param([string]$Id, [string]$Action, [string]$Convert, [string]$Status, [int]$Hours = 5)
            [pscustomobject]@{ id = $Id; action = $Action; convert_batch = $Convert; status = $Status; created_utc = [datetime]::UtcNow.AddHours(-$Hours).ToString('o'); snapshot_id = 7; environment = 'PROD'; account = 'CONTOSO\operator' }
        }
        function script:New-TestItem {
            param([string]$Batch, [int]$Number, [string]$Status, [string]$Step = 'Mailbox')
            [pscustomobject]@{ batch_id = $Batch; object_guid = ('{0:x8}-0000-0000-0000-{0:x12}' -f $Number); kind = 'User'; identity = ('user{0}@contoso.com' -f $Number); step = $Step; status = $Status; message = '' }
        }
        function script:New-TestState {
            param([object[]]$Batches, [object[]]$Items)
            [pscustomobject]@{ Batches = $Batches; Items = $Items; Objects = @(); Config = $null; JournalError = '' }
        }
        function script:Get-TestConvert { 1..10 | ForEach-Object { New-TestItem 'c0c0c0c0' $_ 'Done' } }
        function script:Invoke-TestCarry {
            param($State, $Request, [string]$BatchId, [string]$Root)
            & $script:Gui { param($s, $r, $b, $folder) $old = $script:Gui; $script:Gui = [pscustomobject]@{ State = $s; Root = $folder }; try { Get-PcmGuiRollbackCarry -Request $r -BatchId $b } finally { $script:Gui = $old } } $State $Request $BatchId $Root
        }
    }
    It 'does not list a wave that finished normally as an interrupted rollback' {
        $state = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9), (New-TestBatch 'a1a1a1a1' 'Recover' 'c0c0c0c0' 'Partial')) (@(Get-TestConvert) + @(1..3 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Done' 'RecoverUser' }))
        @(Get-PcmGuiBatchEntry -State $state -Kind ContinueRollback).Count | Should -Be 0
        $new = @(Get-PcmGuiBatchEntry -State $state -Kind Rollback)
        $new[0].State | Should -Be 'Partly rolled back'
        $new[0].Selectable | Should -BeTrue
    }
    It 'lists an interrupted wave with the objects it still has to roll back, and only those' {
        $items = @(Get-TestConvert) + @(1..3 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Done' 'RecoverUser' }) + @(4..5 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Running' 'RecoverUser' })
        $state = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9), (New-TestBatch 'a1a1a1a1' 'Recover' 'c0c0c0c0' 'Partial')) $items
        $entry = @(Get-PcmGuiBatchEntry -State $state -Kind ContinueRollback)
        $entry.Count | Should -Be 1
        $entry[0].BatchId | Should -Be 'c0c0c0c0'
        $entry[0].Selectable | Should -BeTrue
        $entry[0].Sub | Should -BeLike '*2 of the interrupted wave still in the cloud*'
        @($entry[0].Wave | ForEach-Object { $_.Name }) | Should -Be @('user4@contoso.com', 'user5@contoso.com')
    }
    It 'carries no wave when the interrupted rollback was the whole batch' {
        $items = @(Get-TestConvert) + @(1..4 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Done' 'RecoverUser' }) + @(5..10 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Running' 'RecoverUser' })
        $state = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9), (New-TestBatch 'a1a1a1a1' 'Recover' 'c0c0c0c0' 'Partial')) $items
        $entry = @(Get-PcmGuiBatchEntry -State $state -Kind ContinueRollback)
        $entry.Count | Should -Be 1
        @($entry[0].Wave).Count | Should -Be 0
        $entry[0].Sub | Should -BeLike '*6 still in the cloud*'
    }
    It 'lists a rollback that did not end or failed, in any letter case of the batch ID' -TestCases @(
        @{ Status = 'Running' }, @{ Status = 'Failed' }) {
        param($Status)
        $state = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9), (New-TestBatch 'a1a1a1a1' 'Recover' 'C0C0C0C0' $Status)) (@(Get-TestConvert) + @(New-TestItem 'A1A1A1A1' 1 'Failed' 'RecoverUser'))
        $entry = @(Get-PcmGuiBatchEntry -State $state -Kind ContinueRollback)
        $entry.Count | Should -Be 1
        $entry[0].Sub | Should -BeLike '*still in the cloud*'
    }
    It 'gives the command that continues a wave the CSV file of the objects still to roll back' {
        $items = @(Get-TestConvert) + @(1..3 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Done' 'RecoverUser' }) + @(4..5 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Running' 'RecoverUser' })
        $state = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9), (New-TestBatch 'a1a1a1a1' 'Recover' 'c0c0c0c0' 'Partial')) $items
        $request = New-PcmGuiRequest -Action Recover -Mode Apply -Batch 'c0c0c0c0'
        $carry = Invoke-TestCarry -State $state -Request $request -BatchId 'c0c0c0c0' -Root (Join-Path $TestDrive 'carry-root')
        Test-Path -LiteralPath $carry.CsvPath | Should -BeTrue
        @(Import-Csv -LiteralPath $carry.CsvPath | ForEach-Object { $_.Identity }) | Should -Be @(('{0:x8}-0000-0000-0000-{0:x12}' -f 4), ('{0:x8}-0000-0000-0000-{0:x12}' -f 5))
        $command = (Get-PcmGuiCommand -Root $script:Root -ConfigPath 'c.psd1' -Request (New-PcmGuiRequest -Action Recover -Mode Preview -Batch 'c0c0c0c0' -CsvPath $carry.CsvPath) -Engine 'pwsh.exe').Display
        $command | Should -BeLike '*-Action Recover*-Batch c0c0c0c0 -CsvPath*wave-c0c0c0c0-*.csv*'
    }
    It 'carries nothing when the whole batch is still to roll back, and the selection of the run when the journal has no object of it' {
        $whole = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9), (New-TestBatch 'a1a1a1a1' 'Recover' 'c0c0c0c0' 'Partial')) (@(Get-TestConvert) + @(1..10 | ForEach-Object { New-TestItem 'a1a1a1a1' $_ 'Running' 'RecoverUser' }))
        (Invoke-TestCarry -State $whole -Request (New-PcmGuiRequest -Action Recover -Mode Apply -Batch 'c0c0c0c0') -BatchId 'c0c0c0c0' -Root $TestDrive).Count | Should -Be 0
        $none = New-TestState @((New-TestBatch 'c0c0c0c0' 'Convert' '' 'Complete' 9)) (Get-TestConvert)
        $carry = Invoke-TestCarry -State $none -Request (New-PcmGuiRequest -Action Recover -Mode Apply -Batch 'c0c0c0c0' -CsvPath 'D:\waves\wave1.csv' -Scope UsersOnly) -BatchId 'C0C0C0C0' -Root $TestDrive
        $carry.CsvPath | Should -Be 'D:\waves\wave1.csv'
        $carry.Scope | Should -Be 'UsersOnly'
        $carry.ContainsKey('Batch') | Should -BeFalse
        (Invoke-TestCarry -State $none -Request (New-PcmGuiRequest -Action Recover -Mode Apply -Batch 'c0c0c0c0' -MaxObjects 5) -BatchId 'c0c0c0c0' -Root $TestDrive).Count | Should -Be 0
    }
}

Describe 'Console: the real window continues an interrupted wave of a rollback on the same objects' {
    BeforeAll {
        $script:Continue = Invoke-GuiChild -Body @'
$o = [ordered]@{}
$journal = Open-PcmStore -Path (Join-Path $root 'data\PraCloudMailbox-journal.db') -Root $root -Kind Journal
try {
    $null = Invoke-PcmSql $journal 'INSERT INTO batch (id, action, convert_batch, status, created_utc, updated_utc, snapshot_id, environment, computer, account, tool_version) VALUES (@id, @a, @c, @s, @t, @t, @n, @e, @m, @u, @v)' @{ id = 'cd56ef78'; a = 'Recover'; c = '1a2b3c4d'; s = 'Partial'; t = [datetime]::UtcNow.AddHours(-1); n = $fixture.Snapshot; e = 'PROD'; m = 'SRV-PRA01'; u = 'CONTOSO\pra.operator'; v = '2.0.0' }
    $number = 0
    foreach ($object in @($fixture.Objects | Select-Object -First 5)) {
        $number++
        $null = Invoke-PcmSql $journal 'INSERT INTO batch_item (batch_id, object_guid, kind, identity, step, status, updated_utc) VALUES (@b, @o, @k, @i, @s, @t, @u)' @{ b = 'cd56ef78'; o = $object.guid; k = $object.kind; i = $object.upn; s = 'RecoverUser'; t = $(if ($number -le 3) { 'Done' } else { 'Running' }); u = [datetime]::UtcNow }
    }
} finally { Close-PcmStore $journal }
Mod { Update-PcmGuiState }
Click 'BtnNext'; Check 'T_Recover'; Check 'Mode_Continue'
$o.rows = @($g.BatchRows | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.BatchId, $_.State, $_.Selectable, @($_.Wave).Count })
$c.BatchList.SelectedItem = @($g.BatchRows | Where-Object { $_.BatchId -eq '1a2b3c4d' })[0]; Pump 50
$request = Mod { Get-PcmGuiWindowRequest -Mode Preview }
$o.requestCsv = [string]$request.CsvPath
$o.waveIds = @(Import-Csv -LiteralPath $request.CsvPath | ForEach-Object { $_.Identity })
$o.expected = @($fixture.Objects | Select-Object -Skip 3 -First 2 | ForEach-Object { $_.guid })
$o.preview = [string](Mod { Get-PcmGuiCommandPreview })
Click 'BtnNext'; Click 'BtnPreview'; WaitRun
$a = Args_
$o.engine = @{ Action = $a.Action; Batch = $a.Batch; Csv = [bool]$a.CsvPath }
$o.engineIds = @(Import-Csv -LiteralPath (Join-Path $root 'wave-copy.csv') | ForEach-Object { $_.Identity })
Done $o
'@
    }
    It 'lists the interrupted wave with its own objects' {
        @($script:Continue.rows) | Should -Contain '1a2b3c4d|Can be continued|True|2'
    }
    It 'asks the engine for the batch with the CSV file of the wave, never for the whole batch' {
        $script:Continue.requestCsv | Should -BeLike '*wave-1a2b3c4d-*.csv'
        @($script:Continue.waveIds) | Should -Be @($script:Continue.expected)
        $script:Continue.preview | Should -BeLike '*-Action Recover*-Batch 1a2b3c4d -CsvPath*'
        $script:Continue.engine.Action | Should -Be 'Recover'
        $script:Continue.engine.Batch | Should -Be '1a2b3c4d'
        $script:Continue.engine.Csv | Should -BeTrue
        @($script:Continue.engineIds) | Should -Be @($script:Continue.expected)
    }
}

Describe 'Console: a preview is only good for the files it was made on' {
    BeforeAll {
        $script:KeyTool = New-DocumentationToolCopy -Repository $script:Repo -Destination (Join-Path $TestDrive 'key-tool')
        $script:KeyFixture = New-ConsoleFixture -ToolRoot $script:KeyTool
        $script:KeyState = Read-PcmGuiState -Root $script:KeyTool -ConfigPath $script:KeyFixture.Config
        function script:Get-TestKey {
            param($Request, $Config = $null)
            if (-not $Config) { $Config = [pscustomobject]@{ Path = $script:KeyFixture.Config; Base = ''; Hash = ''; Keys = @() } }
            & $script:Gui { param($state, $root, $r, $c) $old = $script:Gui; $script:Gui = [pscustomobject]@{ State = $state; Root = $root }; try { Get-PcmGuiRunKey -Request $r -Config $c } finally { $script:Gui = $old } } $script:KeyState $script:KeyTool $Request $Config
        }
        function script:Invoke-TestPin {
            param($Request, $Pending = $null)
            & $script:Gui { param($state, $root, $r, $p) $old = $script:Gui; $script:Gui = [pscustomobject]@{ State = $state; Root = $root }; try { Set-PcmGuiSnapshotPin -Request $r -Pending $p } finally { $script:Gui = $old } } $script:KeyState $script:KeyTool $Request $Pending
        }
        $script:KeyCsv = Join-Path $TestDrive 'targets.csv'
        [IO.File]::WriteAllText($script:KeyCsv, "Identity`r`nadele.vance@contoso.com`r`n")
    }
    It 'gives the same key twice for the same files' {
        $request = New-PcmGuiRequest -Action Convert -Mode Preview -CsvPath $script:KeyCsv
        (Get-TestKey $request) | Should -Be (Get-TestKey $request)
        (Get-TestKey (New-PcmGuiRequest -Action Convert -Mode Apply -CsvPath $script:KeyCsv)) | Should -Be (Get-TestKey $request)
    }
    It 'changes the key when the configuration file of the run changes on disk, even with no setting changed in the window' {
        $request = New-PcmGuiRequest -Action Convert -Mode Preview
        $before = Get-TestKey $request
        $original = [IO.File]::ReadAllBytes($script:KeyFixture.Config)
        try {
            [IO.File]::AppendAllText($script:KeyFixture.Config, "`r`n# edited`r`n")
            (Get-TestKey $request) | Should -Not -Be $before
        } finally { [IO.File]::WriteAllBytes($script:KeyFixture.Config, $original) }
        (Get-TestKey $request) | Should -Be $before
    }
    It 'changes the key when the snapshot file changes on disk, for the actions that read it' -TestCases @(
        @{ Action = 'Convert' }, @{ Action = 'Check' }) {
        param($Action)
        $request = New-PcmGuiRequest -Action $Action -Mode Preview
        $db = $script:KeyState.Config.Store.Path
        $before = Get-TestKey $request
        $original = [IO.File]::ReadAllBytes($db)
        try {
            $changed = [byte[]]$original.Clone(); $changed[$changed.Length - 1] = [byte](($changed[$changed.Length - 1] + 1) % 256)
            [IO.File]::WriteAllBytes($db, $changed)
            (Get-TestKey $request) | Should -Not -Be $before
        } finally { [IO.File]::WriteAllBytes($db, $original) }
    }
    It 'does not look at the snapshot file for a Collect, which writes it' {
        $request = New-PcmGuiRequest -Action Collect -Mode Preview
        $db = $script:KeyState.Config.Store.Path
        $before = Get-TestKey $request
        $original = [IO.File]::ReadAllBytes($db)
        try {
            $changed = [byte[]]$original.Clone(); $changed[$changed.Length - 1] = [byte](($changed[$changed.Length - 1] + 1) % 256)
            [IO.File]::WriteAllBytes($db, $changed)
            (Get-TestKey $request) | Should -Be $before
        } finally { [IO.File]::WriteAllBytes($db, $original) }
    }
    It 'changes the key when the CSV file that lists the objects changes' {
        $request = New-PcmGuiRequest -Action Convert -Mode Preview -CsvPath $script:KeyCsv
        $before = Get-TestKey $request
        [IO.File]::WriteAllText($script:KeyCsv, "Identity`r`nadele.vance@contoso.com`r`nalex.wilber@contoso.com`r`n")
        (Get-TestKey $request) | Should -Not -Be $before
    }
    It 'changes the key when the engine or a module changes (checked on a copy)' {
        $request = New-PcmGuiRequest -Action Convert -Mode Preview
        $before = Get-TestKey $request
        $module = Join-Path $script:KeyTool 'module\PRA.Gui.Detect.ps1'
        $original = [IO.File]::ReadAllBytes($module)
        try {
            [IO.File]::AppendAllText($module, "`r`n# edited`r`n")
            (Get-TestKey $request) | Should -Not -Be $before
        } finally { [IO.File]::WriteAllBytes($module, $original) }
    }
    It 'gives the previewed snapshot ID to the run, whatever is written after the preview' {
        $latest = & $script:Gui { param($state, $root) $old = $script:Gui; $script:Gui = [pscustomobject]@{ State = $state; Root = $root }; try { Get-PcmGuiLatestSnapshotId } finally { $script:Gui = $old } } $script:KeyState $script:KeyTool
        $latest | Should -Be 42
        $pinned = Invoke-TestPin (New-PcmGuiRequest -Action Convert -Mode Preview)
        $pinned.Snapshot | Should -Be 42
        (Get-PcmGuiCommand -Root $script:KeyTool -ConfigPath 'c.psd1' -Request $pinned -Engine 'pwsh.exe').Display | Should -BeLike '*-Action Convert -Mode Preview -Snapshot 42*'
        $applied = Invoke-TestPin (New-PcmGuiRequest -Action Convert -Mode Apply) @{ Request = $pinned }
        $applied.Snapshot | Should -Be 42
    }
    It 'leaves alone a request that names a snapshot or a batch, and the actions that do not take -Snapshot' {
        (Invoke-TestPin (New-PcmGuiRequest -Action Convert -Mode Preview -Snapshot 9)).Snapshot | Should -Be 9
        (Invoke-TestPin (New-PcmGuiRequest -Action Convert -Mode Preview -Batch 'abcd1234')).Snapshot | Should -Be 0
        (Invoke-TestPin (New-PcmGuiRequest -Action Recover -Mode Preview -Batch 'abcd1234')).Snapshot | Should -Be 0
        (Invoke-TestPin (New-PcmGuiRequest -Action Collect -Mode Preview)).Snapshot | Should -Be 0
    }
}

Describe 'Console: the real window refuses a run whose preview is out of date' {
    BeforeAll {
        $script:Stale = Invoke-GuiChild -Body @'
$o = [ordered]@{}
Click 'BtnNext'; Check 'T_Convert'; Check 'S_Source_Group'; Enter 'F_Group' 'pra-wave1'
Click 'BtnNext'; Click 'BtnPreview'; WaitRun
$o.previewSnapshot = [string](Args_).Snapshot
$o.logHeader = [string]$c.LogExpander.Header.Text
$o.previewed = $g.Previewed
Enter 'ConfirmBox' 'convert'
Add-Content -LiteralPath $fixture.Config -Value '# edited after the preview'
$o.refusedConfig = ''; try { Mod { Start-PcmGuiRun -Mode Apply } } catch { $o.refusedConfig = $_.Exception.Message }
$o.pendingAfterConfig = [bool]$g.Pending
Click 'BtnPreview'; WaitRun
Enter 'ConfirmBox' 'convert'
$db = Join-Path $root 'data\PraCloudMailbox.db'
$bytes = [IO.File]::ReadAllBytes($db); $original = [byte[]]$bytes.Clone()
$bytes[$bytes.Length - 1] = [byte](($bytes[$bytes.Length - 1] + 1) % 256); [IO.File]::WriteAllBytes($db, $bytes); $null = Write-PcmHashFile -Path $db
$o.refusedSnapshot = ''; try { Mod { Start-PcmGuiRun -Mode Apply } } catch { $o.refusedSnapshot = $_.Exception.Message }
[IO.File]::WriteAllBytes($db, $original); $null = Write-PcmHashFile -Path $db
Click 'BtnPreview'; WaitRun
Enter 'ConfirmBox' 'convert'
$o.runEnabled = $c.BtnNext.IsEnabled
Click 'BtnNext'; WaitRun
$a = Args_
$o.apply = @{ Mode = $a.Mode; Snapshot = [string]$a.Snapshot; Force = $a.Force }
Done $o
'@
    }
    It 'previews and runs on the snapshot the preview read, by its ID' {
        $script:Stale.previewSnapshot | Should -Be '42'
        $script:Stale.apply.Mode | Should -Be 'Apply'
        $script:Stale.apply.Snapshot | Should -Be '42'
        $script:Stale.apply.Force | Should -BeTrue
    }
    It 'refuses the run after the configuration file changed, and asks for a new preview' {
        $script:Stale.previewed | Should -BeTrue
        $script:Stale.refusedConfig | Should -BeLike '*changed since the preview. Preview again*'
        $script:Stale.pendingAfterConfig | Should -BeFalse
    }
    It 'refuses the run after the snapshot file changed' {
        $script:Stale.refusedSnapshot | Should -BeLike '*changed since the preview. Preview again*'
        $script:Stale.runEnabled | Should -BeTrue
    }
    It 'tells where the complete log is, since the activity log keeps the last 2,000 lines' {
        $script:Stale.logHeader | Should -BeLike '*the last 2000 lines*complete log is in*'
    }
}

Describe 'Console: the CSV report of a run is read as the report writes it' {
    BeforeAll {
        $script:ReportFolder = Join-Path $TestDrive 'report'
        $null = New-Item -ItemType Directory -Path $script:ReportFolder
        $script:Hostile = @("=HYPERLINK(""http://x"")", '+33 1 23 45', '-1', '@SUM(A1)', "`tTabbed", "`rReturned")
    }
    It 'removes the apostrophe the report puts before = + - @ TAB and CR, and nothing else' {
        foreach ($value in $script:Hostile) {
            (& $script:Gui { param($v) Remove-PcmGuiCsvGuard $v } ("'" + $value)) | Should -Be $value
        }
        foreach ($value in @('plain', "'quoted", '', "it's", '=not guarded', '''x=1')) {
            (& $script:Gui { param($v) Remove-PcmGuiCsvGuard $v } $value) | Should -Be $value
        }
        (& $script:Gui { Remove-PcmGuiCsvGuard $null }) | Should -BeNullOrEmpty
    }
    It 'reads a report with the guard and the byte order mark, as the engine writes it' {
        $path = Join-Path $script:ReportFolder 'guarded.csv'
        $rows = foreach ($value in $script:Hostile) { [pscustomobject]@{ Identity = 'user@contoso.com'; FinalStatus = 'Success'; Detail = ("'" + $value) } }
        $rows | Export-Csv -LiteralPath $path -Delimiter ';' -Encoding utf8BOM -NoTypeInformation
        [IO.File]::ReadAllBytes($path)[0..2] | Should -Be @(239, 187, 191)
        $read = @(& $script:Gui { param($p) Read-PcmGuiReport $p } $path)
        $read.Count | Should -Be $script:Hostile.Count
        @($read | ForEach-Object { $_.Detail }) | Should -Be $script:Hostile
        $read[0].Identity | Should -Be 'user@contoso.com'
    }
    It 'reads a report of the earlier edition, with no guard and no byte order mark, as it is' {
        $path = Join-Path $script:ReportFolder 'old.csv'
        $text = "Identity;FinalStatus;Detail`r`nuser@contoso.com;Success;'quoted`r`nuser2@contoso.com;Error;=value`r`n"
        [IO.File]::WriteAllText($path, $text, (New-Object Text.UTF8Encoding($false)))
        $read = @(& $script:Gui { param($p) Read-PcmGuiReport $p } $path)
        $read.Count | Should -Be 2
        $read[0].Detail | Should -Be "'quoted"
        $read[1].Detail | Should -Be '=value'
        $read[1].Identity | Should -Be 'user2@contoso.com'
    }
}

Describe 'Console: a big preview does not freeze the window (2,500 objects, no window needed)' {
    BeforeAll {
        $script:Count = 2500
        $script:Big = [pscustomobject]@{
            Objects = @(1..$script:Count | ForEach-Object {
                    [pscustomobject]@{ object_guid = ('{0:x8}-0000-0000-0000-{0:x12}' -f $_); kind = 'User'; user_principal_name = ('user{0}@contoso.com' -f $_); primary_smtp_address = ('user{0}@contoso.com' -f $_)
                        sam_account_name = ('user{0}' -f $_); name = ('User {0}' -f $_); display_name = ('User {0}' -f $_)
                        distinguished_name = ('CN=User {0},OU=Site{1},DC=contoso,DC=com' -f $_, ($_ % 20)); organizational_unit = ('contoso.com/Site{0}' -f ($_ % 20)) }
                })
            Items = @(); Batches = @(); Config = $null
        }
    }
    It 'builds the preview rows of 2,500 objects in seconds, with one lookup by object ID each' {
        $rows = foreach ($object in $script:Big.Objects) { [pscustomobject]@{ FinalStatus = 'Planned'; Warnings = ''; Detail = ''; Kind = 'User'; Identity = $object.user_principal_name; ObjectGuid = $object.object_guid.ToUpperInvariant() } }
        $time = Measure-Command {
            $built = & $script:Gui { param($state, $csv) $old = $script:Gui; $script:Gui = [pscustomobject]@{ State = $state }; try { foreach ($row in $csv) { ConvertTo-PcmGuiObjectRow $row } } finally { $script:Gui = $old } } $script:Big $rows
            $script:Built = @($built)
        }
        $script:Built.Count | Should -Be $script:Count
        $script:Built[1233].Name | Should -Be 'User 1234'
        $time.TotalSeconds | Should -BeLessThan 20
    }
    It 'returns the row itself, with its values, not the pairs of its properties' {
        $row = & $script:Gui {
            $model = @{ Objects = @{}; ObjectIndex = @{} }
            Add-PcmGuiModelObject -Model $model -StepId 's5' -Row (New-PcmGuiRow @{ Name = 'user1@contoso.com'; Result = 'mailbox ready'; Status = 'Ok' })
            Get-PcmGuiModelObject -Model $model -StepId 's5' -Identity 'USER1@contoso.com'
        }
        $row.Result | Should -Be 'mailbox ready'
        $row.Status | Should -Be 'Ok'
    }
    It 'finds the row of an object of a step by its identity without reading the list' {
        $time = Measure-Command {
            $found = & $script:Gui {
                param($count)
                $model = @{ Objects = @{}; ObjectIndex = @{} }
                foreach ($i in 1..$count) { Add-PcmGuiModelObject -Model $model -StepId 's5' -Row (New-PcmGuiRow @{ Name = ('user{0}@contoso.com' -f $i); Status = 'Running' }) }
                $hits = 0
                foreach ($i in 1..$count) { if (Get-PcmGuiModelObject -Model $model -StepId 's5' -Identity ('USER{0}@Contoso.com' -f $i)) { $hits++ } }
                [pscustomobject]@{ Hits = $hits; Listed = $model.Objects['s5'].Count; Unknown = [bool](Get-PcmGuiModelObject -Model $model -StepId 's5' -Identity 'nobody@contoso.com'); NoStep = [bool](Get-PcmGuiModelObject -Model $model -StepId 's9' -Identity 'x') }
            } $script:Count
        }
        $found.Hits | Should -Be $script:Count
        $found.Listed | Should -Be $script:Count
        $found.Unknown | Should -BeFalse
        $found.NoStep | Should -BeFalse
        $time.TotalSeconds | Should -BeLessThan 20
    }
    It 'looks up mailboxes and OUs of a big snapshot quickly, letter case ignored' {
        $time = Measure-Command {
            foreach ($i in 1..100) {
                @(Find-PcmGuiMailbox -State $script:Big -Text ('USER{0}@contoso.com' -f ($i * 20))).Count | Should -Be 1
                @(Get-PcmGuiOuMatch -State $script:Big -Text 'CONTOSO.COM/site3').Count | Should -Be ($script:Count / 20)
            }
        }
        $time.TotalSeconds | Should -BeLessThan 30
    }
    It 'sums up the batches and builds the lists of 2,500 objects in seconds' {
        $batches = @([pscustomobject]@{ id = 'c0c0c0c0'; action = 'Convert'; convert_batch = ''; status = 'Complete'; created_utc = [datetime]::UtcNow.AddHours(-3).ToString('o'); snapshot_id = 7; environment = 'PROD'; account = 'x' })
        $items = @(1..$script:Count | ForEach-Object { [pscustomobject]@{ batch_id = 'c0c0c0c0'; object_guid = ('{0:x8}-0000-0000-0000-{0:x12}' -f $_); kind = 'User'; identity = ('user{0}@contoso.com' -f $_); step = 'Mailbox'; status = 'Done'; message = '' } })
        $state = [pscustomobject]@{ Batches = $batches; Items = $items; Objects = $script:Big.Objects; Config = $null }
        $time = Measure-Command {
            (@(Get-PcmGuiBatchSummary -State $state))[0].Waiting | Should -Be $script:Count
            @(Get-PcmGuiBatchEntry -State $state -Kind Follow).Count | Should -Be 1
            @(Get-PcmGuiWaveRow -State $state -BatchId 'c0c0c0c0').Count | Should -Be $script:Count
        }
        $time.TotalSeconds | Should -BeLessThan 20
    }
    It 'reads the report of 2,500 objects in seconds' {
        $path = Join-Path $TestDrive 'big.csv'
        1..$script:Count | ForEach-Object { [pscustomobject]@{ Identity = ('user{0}@contoso.com' -f $_); FinalStatus = 'Planned'; Detail = ("'=x{0}" -f $_) } } | Export-Csv -LiteralPath $path -Delimiter ';' -Encoding utf8BOM -NoTypeInformation
        $time = Measure-Command { $script:BigRead = @(& $script:Gui { param($p) Read-PcmGuiReport $p } $path) }
        $script:BigRead.Count | Should -Be $script:Count
        $script:BigRead[9].Detail | Should -Be '=x10'
        $time.TotalSeconds | Should -BeLessThan 20
    }
}

Describe 'Console: the snapshot and the scheduled tasks are not read again at every keystroke' {
    BeforeAll {
        $script:CacheTool = New-DocumentationToolCopy -Repository $script:Repo -Destination (Join-Path $TestDrive 'cache-tool')
        $script:CacheFixture = New-ConsoleFixture -ToolRoot $script:CacheTool
        $script:CacheState = Read-PcmGuiState -Root $script:CacheTool -ConfigPath $script:CacheFixture.Config
    }
    It 'hashes the snapshot file once, and again only when its length or its time of last write changes' {
        $first = Get-PcmGuiSnapshotInfo -State $script:CacheState
        $first.Integrity | Should -Be 'Verified'
        $path = $first.Path
        # Same length, same time of last write: the answer is kept, so the file is not read again.
        $time = (Get-Item -LiteralPath $path).LastWriteTimeUtc
        $bytes = [IO.File]::ReadAllBytes($path); $bytes[$bytes.Length - 1] = [byte](($bytes[$bytes.Length - 1] + 1) % 256)
        [IO.File]::WriteAllBytes($path, $bytes); (Get-Item -LiteralPath $path).LastWriteTimeUtc = $time
        (Get-PcmGuiSnapshotInfo -State $script:CacheState).Integrity | Should -Be 'Verified'
        # A new time of last write: the file is hashed again and found changed.
        (Get-Item -LiteralPath $path).LastWriteTimeUtc = $time.AddSeconds(5)
        (Get-PcmGuiSnapshotInfo -State $script:CacheState).Integrity | Should -Be 'Changed'
        # Restored: verified again.
        $bytes[$bytes.Length - 1] = [byte](($bytes[$bytes.Length - 1] + 255) % 256); [IO.File]::WriteAllBytes($path, $bytes)
        (Get-PcmGuiSnapshotInfo -State $script:CacheState).Integrity | Should -Be 'Verified'
    }
    It 'lists the scheduled tasks once a minute at most, and again when asked to refresh' {
        $calls = & $script:Gui {
            $script:GuiTaskCache = $null
            $counter = @{ N = 0 }
            function Find-PcmGuiCollectTask { $counter.N++; [pscustomobject]@{ Found = $false; Text = '' } }
            $null = Get-PcmGuiCollectTask; $null = Get-PcmGuiCollectTask; $null = Get-PcmGuiCollectTask
            $afterThree = $counter.N
            $null = Get-PcmGuiCollectTask -Refresh
            $script:GuiTaskCache = $null
            Remove-Item Function:\Find-PcmGuiCollectTask
            [pscustomobject]@{ AfterThree = $afterThree; AfterRefresh = $counter.N }
        }
        $calls.AfterThree | Should -Be 1
        $calls.AfterRefresh | Should -Be 2
    }
}

Describe 'Console: the lists of the window virtualize, and the activity log says where the whole log is' {
    BeforeAll {
        $script:Xaml = [xml][IO.File]::ReadAllText((Join-Path $script:Root 'module\PRA.Gui.xaml'))
        $script:XamlText = [IO.File]::ReadAllText((Join-Path $script:Root 'module\PRA.Gui.xaml'))
    }
    It 'draws the grid of a wave with a virtualizing panel, not inside a ScrollViewer that disables it' {
        $manager = New-Object Xml.XmlNamespaceManager $script:Xaml.NameTable
        $manager.AddNamespace('p', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
        $manager.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
        $grid = $script:Xaml.SelectSingleNode("//p:ItemsControl[@x:Name='WaveGrid']", $manager)
        $grid | Should -Not -BeNullOrEmpty
        $grid.ParentNode.LocalName | Should -Not -Be 'ScrollViewer'
        $grid.SelectSingleNode('p:ItemsControl.ItemsPanel/p:ItemsPanelTemplate/p:VirtualizingStackPanel', $manager) | Should -Not -BeNullOrEmpty
        $grid.GetAttribute('VirtualizingPanel.VirtualizationMode') | Should -Be 'Recycling'
        $grid.GetAttribute('ScrollViewer.CanContentScroll') | Should -Be 'True'
    }
    It 'never turns the content scroll off, and asks the plain lists to recycle their rows' {
        $script:XamlText | Should -Not -Match 'CanContentScroll\s*=\s*"False"'
        $script:XamlText | Should -Not -Match 'IsVirtualizing\s*=\s*"False"'
        $script:XamlText | Should -Match '(?s)<Style x:Key="PlainList".*?VirtualizingPanel\.VirtualizationMode.*?Recycling.*?</Style>'
    }
}

Describe 'Console: the folders of a run are created private, and a failure is reported' {
    It 'creates a missing folder with protected permissions' {
        $path = Join-Path $TestDrive 'gui-private\logs'
        & $script:Gui { param($p) Initialize-PcmGuiFolder $p } $path
        $path | Should -Exist
        (Get-Acl -LiteralPath $path).AreAccessRulesProtected | Should -BeTrue
    }
    It 'leaves an existing folder as it is' {
        $path = Join-Path $TestDrive 'gui-existing'
        $null = New-Item -ItemType Directory -Path $path
        & $script:Gui { param($p) Initialize-PcmGuiFolder $p } $path
        (Get-Acl -LiteralPath $path).AreAccessRulesProtected | Should -BeFalse
    }
    It 'reports a folder that cannot be created instead of going on' {
        $file = Join-Path $TestDrive 'gui-a-file'
        Set-Content -LiteralPath $file -Value 'x'
        { & $script:Gui { param($p) Initialize-PcmGuiFolder $p } $file } | Should -Throw '*a file exists where the folder is expected*'
        { & $script:Gui { param($p) Initialize-PcmGuiFolder $p } (Join-Path $TestDrive 'gui|bad') } | Should -Throw '*cannot be created*'
    }
}
