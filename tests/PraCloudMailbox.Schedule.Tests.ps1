<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of -Action Schedule, the registration of the daily Collect task (Pester 5+, offline): the checks of the account, the proof from Exchange, the right to log
    on as a batch job, the registration itself, the secrecy of the password and the house format of the command line.
.DESCRIPTION
    Runs the real entry script (without its package integrity bootstrap and its module imports) in this process. Everything that touches Windows, Active Directory or Exchange is
    replaced: the directory lookup, the remote session, the user rights, the elevation and the Task Scheduler. Nothing is registered and nothing leaves this computer. PowerShell 7 only.
.NOTES
    Author  : Nicolas Fabert
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:AnsiBefore = $env:PRA_ANSI; $script:IconsBefore = $env:PRA_ICONS
    $env:PRA_ANSI = '1'; $env:PRA_ICONS = 'Emoji'
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    foreach ($name in 'PRA.Common', 'PRA.Store', 'PRA.Cloud', 'PRA.Collect') { Import-Module (Join-Path $script:Root "module\$name.psm1") -Force -Global }
    $script:Theme = & (Get-Module PRA.Common) { $script:Theme }
    $script:Dot = [string][char]0x00B7
    $script:Password = 'Zx9!Pw-Secret-7731'
    function script:New-TestCredential {
        <# A credential made character by character: the tests hold a fake password in clear text on purpose. #>
        param([Parameter(Mandatory)][string]$Account, [Parameter(Mandatory)][string]$Phrase)
        $secure = [securestring]::new()
        foreach ($character in $Phrase.ToCharArray()) { $secure.AppendChar($character) }
        [pscredential]::new($Account, $secure)
    }
    $script:Cred = New-TestCredential 'CONTOSO\svc-pra-collect' $script:Password

    $text = [IO.File]::ReadAllText((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'))
    $start = $text.IndexOf('# >>> PRA package integrity bootstrap'); $end = $text.IndexOf('# <<< PRA package integrity bootstrap')
    $text = $text.Remove($start, $end - $start)
    $text = [regex]::Replace($text, '(?m)^Import-Module \(Join-Path \$PSScriptRoot ''PraCloudMailbox\.psd1''\).*$', '')
    $text = [regex]::Replace($text, '(?m)^if \(\$Action -in @\(''Collect'', ''Schedule''\)\) \{ Import-Module.*\r?\n.*Import-Module.*$', '')
    $text = $text.Replace('$caller = $PSCmdlet', '$caller = if ($global:ConsoleAnswersNo) { $global:ConsoleCaller } else { $PSCmdlet }')
    $script:Entry = Join-Path $script:Root 'Invoke-Schedule.ps1'
    [IO.File]::WriteAllText($script:Entry, $text, [Text.UTF8Encoding]::new($true))
    $global:ConsoleAnswersNo = $false
    $global:ConsoleCaller = [pscustomobject]@{}
    $global:ConsoleCaller | Add-Member -MemberType ScriptMethod -Name ShouldContinue -Value { param($q, $t) $false }
    $global:ConsoleCaller | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($t, $o) $true }

    function script:Set-ScheduleFake {
        <# The synthetic computer: what the directory knows, what Exchange answers, the rights of the domain controller and the tasks of the Task Scheduler. #>
        param([int]$DomainRole = 3, [string]$Rights = '', [bool]$Elevated = $true, [bool]$Interactive = $true)
        $user = { param($name, $classes, $groups = @('S-1-5-21-1-2-3-513')) [pscustomobject]@{ Found = $true; Sid = 'S-1-5-21-1-2-3-1105'; Name = $name; Source = 'Directory'; Classes = $classes; Groups = $groups; GroupsKnown = $true; Problem = '' } }
        $plain = @('top', 'person', 'organizationalPerson', 'user')
        $wellKnown = { param($sid, $name) [pscustomobject]@{ Found = $true; Sid = $sid; Name = $name; Source = 'WellKnown'; Classes = @(); Groups = @(); GroupsKnown = $false; Problem = '' } }
        $global:ScheduleFake = @{
            Accounts = @{
                'contoso\svc-pra-collect' = (& $user 'CONTOSO\svc-pra-collect' $plain)
                'svc-pra-collect@contoso.com' = (& $user 'CONTOSO\svc-pra-collect' $plain)
                'contoso\srv01$' = (& $user 'CONTOSO\SRV01$' ($plain + 'computer'))
                'contoso\gmsa-pra$' = (& $user 'CONTOSO\gmsa-pra$' ($plain + 'computer' + 'msDS-GroupManagedServiceAccount'))
                'contoso\msa-pra$' = (& $user 'CONTOSO\msa-pra$' ($plain + 'computer' + 'msDS-ManagedServiceAccount'))
                'contoso\staff' = (& $user 'CONTOSO\staff' @('top', 'group'))
                'system' = (& $wellKnown 'S-1-5-18' 'NT AUTHORITY\SYSTEM')
                'nt authority\system' = (& $wellKnown 'S-1-5-18' 'NT AUTHORITY\SYSTEM')
                'nt authority\local service' = (& $wellKnown 'S-1-5-19' 'NT AUTHORITY\LOCAL SERVICE')
                'network service' = (& $wellKnown 'S-1-5-20' 'NT AUTHORITY\NETWORK SERVICE')
                's-1-5-20' = (& $wellKnown 'S-1-5-20' 'NT AUTHORITY\NETWORK SERVICE')
                'nt service\mssqlserver' = (& $wellKnown 'S-1-5-80-3880718306-3832830129-1677859214-2598158968-1052248003' 'NT SERVICE\MSSQLSERVER')
                'ex01\administrator' = [pscustomobject]@{ Found = $true; Sid = 'S-1-5-21-9-9-9-500'; Name = 'EX01\Administrator'; Source = 'Local'; Classes = @(); Groups = @(); GroupsKnown = $false; Problem = '' }
            }
            Elevated = $Elevated; Interactive = $Interactive; DomainRole = $DomainRole; Tasks = @(); Registered = $null; RegisterCalls = 0; Sessions = @(); SessionCalls = 0; Removed = 0
            SessionError = ''; CommandError = ''; RegisterError = ''; Rights = $Rights; Prompts = 0; PromptAnswer = $null
        }
        if (-not $Rights) { $global:ScheduleFake.Rights = "[Unicode]`r`nUnicode=yes`r`n[Privilege Rights]`r`nSeBatchLogonRight = *S-1-5-32-544,*S-1-5-32-551,*S-1-5-32-559`r`n" }
    }

    function script:Install-ScheduleMocks {
        Mock -ModuleName PRA.Collect Test-PcmElevated { [bool]$global:ScheduleFake.Elevated }
        Mock -ModuleName PRA.Collect Test-PcmInteractiveHost { [bool]$global:ScheduleFake.Interactive }
        Mock -ModuleName PRA.Collect Get-PcmDirectoryAccount {
            $key = $Name.Trim().ToLowerInvariant()
            if ($global:ScheduleFake.Accounts.ContainsKey($key)) { return $global:ScheduleFake.Accounts[$key] }
            [pscustomobject]@{ Found = $false; Sid = ''; Name = ''; Source = 'Directory'; Classes = @(); Groups = @(); GroupsKnown = $false; Problem = 'No mapping between account names and security IDs was done.' }
        }
        Mock -ModuleName PRA.Collect New-PSSession {
            $global:ScheduleFake.SessionCalls++
            $global:ScheduleFake.SessionArguments = @{ ConfigurationName = $ConfigurationName; ConnectionUri = [string]$ConnectionUri; Authentication = [string]$Authentication; UserName = $Credential.UserName }
            if ($global:ScheduleFake.SessionError) { throw $global:ScheduleFake.SessionError }
            [pscustomobject]@{ Id = 7 }
        }
        Mock -ModuleName PRA.Collect -RemoveParameterType Session Invoke-Command {
            $global:ScheduleFake.ProofScript = [string]$ScriptBlock
            if ($global:ScheduleFake.CommandError) { throw $global:ScheduleFake.CommandError }
            'ok'
        }
        Mock -ModuleName PRA.Collect -RemoveParameterType Session Remove-PSSession { $global:ScheduleFake.Removed++ }
        Mock -ModuleName PRA.Collect Get-CimInstance { [pscustomobject]@{ DomainRole = $global:ScheduleFake.DomainRole } }
        Mock -ModuleName PRA.Collect Export-PcmUserRight {
            $global:ScheduleFake.RightsFile = $Path
            Set-Content -LiteralPath $Path -Value $global:ScheduleFake.Rights -Encoding Unicode
            0
        }
        Mock -ModuleName PRA.Collect Get-ScheduledTask { $global:ScheduleFake.Tasks }
        Mock -ModuleName PRA.Collect Register-ScheduledTask {
            $global:ScheduleFake.RegisterCalls++
            if ($global:ScheduleFake.RegisterError) { throw $global:ScheduleFake.RegisterError }
            $global:ScheduleFake.Registered = [pscustomobject]@{ TaskName = $TaskName; TaskPath = $TaskPath; Action = $Action; Trigger = $Trigger; Settings = $Settings; User = $User; Password = $Password; RunLevel = [string]$RunLevel; Force = [bool]$Force }
            $global:ScheduleFake.Tasks = @([pscustomobject]@{ TaskName = $TaskName; TaskPath = '\'; Principal = [pscustomobject]@{ UserId = $User }
                    Triggers = @([pscustomobject]@{ StartBoundary = [string]$Trigger.StartBoundary }); Actions = @([pscustomobject]@{ Execute = $Action.Execute; Arguments = $Action.Arguments }) })
        }
        Mock Get-Credential { $global:ScheduleFake.Prompts++; $global:ScheduleFake.PromptAnswer }
    }

    function script:Reset-ScheduleWork {
        <# A work folder with a configuration file (Exchange.Server set), and the synthetic computer. #>
        param([string]$Collect = '', [string]$Server = 'ex01.contoso.com', [int]$DomainRole = 3, [string]$Rights = '', [bool]$Elevated = $true, [bool]$Interactive = $true)
        $script:Work = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory -Path $script:Work, (Join-Path $script:Work 'data')
        $script:ConfigFile = Join-Path $script:Work 'schedule.config.psd1'
        $work = $script:Work
        Set-Content -LiteralPath $script:ConfigFile -Encoding UTF8 -Value @"
@{
    Environment = 'PROD'
    Exchange = @{ Server = '$Server' }
    Scope = @{ Mode = 'OU'; SearchBase = 'OU=Paris,DC=contoso,DC=com' }
    EntraConnect = @{ Mode = 'Manual' }
    $Collect
    Store = @{ Path = '$work\data\PraCloudMailbox.db'; JournalPath = '$work\data\PraCloudMailbox-journal.db'; BackupFolder = '$work\data\backup' }
    Logging = @{ Folder = '$work\logs' }
    Report = @{ Folder = '$work\reports' }
}
"@
        Set-ScheduleFake -DomainRole $DomainRole -Rights $Rights -Elevated $Elevated -Interactive $Interactive
    }

    function script:Invoke-Schedule {
        <# One run of the entry script. Lines: the text and the colour of each console line; Text: all of them; Code: the exit code. #>
        param([hashtable]$Arguments = @{}, [switch]$DefaultConfig)
        if (-not $Arguments.ContainsKey('ConfigPath') -and -not $DefaultConfig) { $Arguments = $Arguments + @{ ConfigPath = $script:ConfigFile } }
        if (-not $Arguments.ContainsKey('Action')) { $Arguments.Action = 'Schedule' }
        $global:LASTEXITCODE = $null
        $records = @(& $script:Entry @Arguments 6>&1 5>&1 4>&1 3>&1 2>&1)
        $code = $global:LASTEXITCODE
        $names = @{ 30 = 'Black'; 31 = 'DarkRed'; 32 = 'DarkGreen'; 33 = 'DarkYellow'; 34 = 'DarkBlue'; 35 = 'DarkMagenta'; 36 = 'DarkCyan'; 37 = 'Gray'; 90 = 'DarkGray'; 91 = 'Red'; 92 = 'Green'; 93 = 'Yellow'; 94 = 'Blue'; 95 = 'Magenta'; 96 = 'Cyan'; 97 = 'White' }
        $lines = [Collections.Generic.List[object]]::new()
        foreach ($record in $records) {
            if ($record -isnot [Management.Automation.InformationRecord]) { continue }
            $data = $record.MessageData
            $raw = if ($data -is [Management.Automation.HostInformationMessage]) { [string]$data.Message } else { [string]$data }
            $match = [regex]::Match($raw, '^\x1b\[(?<code>\d+)m(?<text>.*)\x1b\[0m$', [Text.RegularExpressions.RegexOptions]::Singleline)
            if ($match.Success) { $lines.Add([pscustomobject]@{ Color = $names[[int]$match.Groups['code'].Value]; Text = $match.Groups['text'].Value }) }
            else { $lines.Add([pscustomobject]@{ Color = ''; Text = $raw }) }
        }
        [pscustomobject]@{ Code = $code; Records = $records; Lines = $lines.ToArray(); Text = (($lines | ForEach-Object { $_.Text }) -join "`n") }
    }

    function script:Get-ScheduleCard {
        param([Parameter(Mandatory)][object]$Run, [ValidateSet('Banner', 'Summary')][string]$Kind)
        $lines = @($Run.Lines)
        $top = @(0..($lines.Count - 1) | Where-Object { $lines[$_].Text -match $(if ($Kind -eq 'Banner') { '^  ╭─+╮$' } else { '^  ╭─ .+ ─+╮$' }) }) | Select-Object -First 1
        $bottomAt = @($top..($lines.Count - 1) | Where-Object { $lines[$_].Text -match '^  ╰─+╯$' }) | Select-Object -First 1
        if ($Kind -eq 'Banner') {
            $rows = @(); $at = $bottomAt + 1
            while ($at -lt $lines.Count -and $lines[$at].Text.Trim()) { $rows += $lines[$at]; $at++ }
            return [pscustomobject]@{ Top = $lines[$top]; Bottom = $lines[$bottomAt]; Rows = $rows; Box = @($lines[$top..$bottomAt]) }
        }
        [pscustomobject]@{ Top = $lines[$top]; Bottom = $lines[$bottomAt]; Rows = @($lines[($top + 1)..($bottomAt - 1)]); Box = @($lines[$top..$bottomAt]) }
    }

    function script:Get-ScheduleLabel {
        param([object[]]$Rows)
        foreach ($row in $Rows) {
            $match = [regex]::Match($row.Text, '^ {5}(?<icon>\S+) (?<label>.{11}) (?<text>\S.*)$')
            if ($match.Success) { [pscustomobject]@{ Label = $match.Groups['label'].Value.TrimEnd(); Text = $match.Groups['text'].Value; Color = $row.Color } }
        }
    }

    function script:Get-ScheduleStep {
        param([object]$Run)
        @($Run.Lines | Where-Object { $_.Text -match '^ {2}\s*\d+/\d+  \S+ .+$' } | ForEach-Object {
                $m = [regex]::Match($_.Text, '^ {2}\s*(?<n>\d+)/(?<total>\d+)  (?<icon>\S+) (?<title>.+)$')
                [pscustomobject]@{ Number = [int]$m.Groups['n'].Value; Total = [int]$m.Groups['total'].Value; Title = $m.Groups['title'].Value; Color = $_.Color } })
    }

    function script:Get-ScheduleVerdict { param([object]$Run) @($Run.Lines | Where-Object { $_.Text })[-1] }

    function script:Find-InWorkFiles {
        <# The files that the run left in its work folder (the log, the transcript, the reports) and that hold the text: a password must never be in one. #>
        param([Parameter(Mandatory)][string]$Text)
        @(Get-ChildItem -LiteralPath $script:Work -Recurse -File | ForEach-Object { $name = $_.Name; @(Get-Content -LiteralPath $_.FullName | Where-Object { $_.Contains($Text) } | ForEach-Object { '{0}: {1}' -f $name, $_.Trim() }) })
    }

    function script:Get-LocalStart {
        <# The time of day of a trigger, on this computer's clock: a start without Z is the local time, a start with Z is in UTC. #>
        param([Parameter(Mandatory)][object]$Trigger)
        $start = [datetime]::Parse([string]$Trigger.StartBoundary, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        if ($start.Kind -eq [DateTimeKind]::Utc) { $start = $start.ToLocalTime() }
        $start.ToString('HH:mm')
    }

    $script:AllowedLabels = @('Action', 'Mode', 'Task', 'Config', 'Log', 'Result', 'Account', 'Runs at', 'Command', 'Exchange', 'Batch right', 'Replaces', 'Error', 'Warning', 'Report', 'Duration', 'Next')
}

AfterAll {
    Remove-PcmTestPackage $script:Root
    Remove-Variable -Name ScheduleFake, ConsoleAnswersNo, ConsoleCaller -Scope Global -ErrorAction SilentlyContinue
    if ($null -eq $script:AnsiBefore) { Remove-Item Env:\PRA_ANSI -ErrorAction SilentlyContinue } else { $env:PRA_ANSI = $script:AnsiBefore }
    if ($null -eq $script:IconsBefore) { Remove-Item Env:\PRA_ICONS -ErrorAction SilentlyContinue } else { $env:PRA_ICONS = $script:IconsBefore }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Schedule: the settings' {
    It 'has a task name and a time in the configuration, with defaults, and checks them' {
        $config = Import-PcmConfiguration -Path (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1') -Root $script:Root
        $config.Collect.Schedule.TaskName | Should -Be 'PraCloudMailbox Collect'
        $config.Collect.Schedule.At | Should -Be '02:00'
        $path = Join-Path $TestDrive 'bad.psd1'
        foreach ($case in @(@{ Text = "Collect = @{ Schedule = @{ TaskName = 'a\b' } }"; Message = '*Collect.Schedule.TaskName*' }, @{ Text = "Collect = @{ Schedule = @{ TaskName = '  ' } }"; Message = '*Collect.Schedule.TaskName*' }
                @{ Text = "Collect = @{ Schedule = @{ At = '25:00' } }"; Message = '*Collect.Schedule.At*' }, @{ Text = "Collect = @{ Schedule = @{ At = '2am' } }"; Message = '*Collect.Schedule.At*' }, @{ Text = "Collect = @{ Schedule = @{ Frequency = 'Weekly' } }"; Message = '*unknown setting Collect.Schedule.Frequency*' })) {
            Set-Content -LiteralPath $path -Value ("@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; " + $case.Text + ' }') -Encoding UTF8
            { Import-PcmConfiguration -Path $path -Root $TestDrive } | Should -Throw $case.Message
        }
        Set-Content -LiteralPath $path -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Collect = @{ Schedule = @{ TaskName = ' Nightly Collect '; At = '7:05' } } }" -Encoding UTF8
        $read = Import-PcmConfiguration -Path $path -Root $TestDrive
        $read.Collect.Schedule.TaskName | Should -Be 'Nightly Collect'
        $read.Collect.Schedule.At | Should -Be '07:05'
    }
    It 'reads a time of day on 24 hours and nothing else' -TestCases @(
        @{ Text = '2:00'; Expected = '02:00' }, @{ Text = '02:00'; Expected = '02:00' }, @{ Text = ' 23:59 '; Expected = '23:59' }, @{ Text = '0:00'; Expected = '00:00' }
        @{ Text = '24:00'; Expected = $null }, @{ Text = '12:60'; Expected = $null }, @{ Text = '2am'; Expected = $null }, @{ Text = '2'; Expected = $null }, @{ Text = ''; Expected = $null }) {
        param($Text, $Expected)
        ConvertTo-PcmScheduleTime $Text | Should -Be $Expected
    }
    It 'builds the command of the task: the PowerShell that runs now, the full path of this script, Collect in Apply mode, and the configuration file only when it is not the default' {
        $config = Import-PcmConfiguration -Path (Join-Path $script:Root 'config\PraCloudMailbox.config.psd1') -Root $script:Root
        $setting = Get-PcmScheduleSetting -Config $config -ScriptPath 'C:\PRA\Invoke-PraCloudMailbox.ps1'
        $setting.Execute | Should -Be (Join-Path $PSHOME 'pwsh.exe')
        $setting.Argument | Should -BeExactly '-NoProfile -NonInteractive -File "C:\PRA\Invoke-PraCloudMailbox.ps1" -Action Collect -Mode Apply -Force'
        $setting.WorkingDirectory | Should -Be 'C:\PRA'
        $setting.At | Should -Be '02:00'
        $setting.TaskPath | Should -Be '\'
        (Get-PcmScheduleSetting -Config $config -At '5:30' -ScriptPath 'C:\PRA\Invoke-PraCloudMailbox.ps1').At | Should -Be '05:30'
        (Get-PcmScheduleSetting -Config $config -ScriptPath 'C:\PRA\Invoke-PraCloudMailbox.ps1' -ConfigPath 'D:\cfg\other.psd1').Argument | Should -BeLike '*-Force -ConfigPath "D:\cfg\other.psd1"'
        { Get-PcmScheduleSetting -Config $config -At '25:00' -ScriptPath 'C:\PRA\Invoke-PraCloudMailbox.ps1' } | Should -Throw "*is not a time of day*"
    }
}

Describe 'Schedule: which account may run the task' {
    BeforeEach { Reset-ScheduleWork; Install-ScheduleMocks }
    It 'accepts a domain user, typed with any letter case or as a user principal name' -TestCases @(@{ Typed = 'CONTOSO\svc-pra-collect' }, @{ Typed = 'cOnToSo\SVC-PRA-collect' }, @{ Typed = 'svc-pra-collect@contoso.com' }, @{ Typed = '  contoso\svc-pra-collect ' }) {
        param($Typed)
        $account = Get-PcmScheduleAccount -Name $Typed
        $account.Ok | Should -BeTrue
        $account.Name | Should -BeExactly 'CONTOSO\svc-pra-collect'
    }
    It 'refuses <Why>, with the reason in plain words' -TestCases @(
        @{ Why = 'a computer account'; Typed = 'CONTOSO\SRV01$'; Reason = 'is a computer account. Exchange remote PowerShell refuses it' }
        @{ Why = 'a group managed service account'; Typed = 'contoso\GMSA-PRA$'; Reason = 'is a group managed service account (gMSA)' }
        @{ Why = 'a managed service account'; Typed = 'CONTOSO\msa-pra$'; Reason = 'is a managed service account' }
        @{ Why = 'SYSTEM, in any letter case'; Typed = 'SyStEm'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'SYSTEM, with its domain'; Typed = 'NT AUTHORITY\SYSTEM'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'LOCAL SERVICE'; Typed = 'NT AUTHORITY\Local Service'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'NETWORK SERVICE'; Typed = 'network service'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'the SID of NETWORK SERVICE'; Typed = 'S-1-5-20'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'a service identity'; Typed = 'NT SERVICE\MSSQLSERVER'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'a local account'; Typed = 'EX01\Administrator'; Reason = 'is a local account of this computer' }
        @{ Why = 'a group'; Typed = 'CONTOSO\staff'; Reason = 'is not a user account of the domain' }
        @{ Why = 'an unknown name'; Typed = 'CONTOSO\nobody'; Reason = 'was not found in the domain' }) {
        param($Why, $Typed, $Reason)
        $account = Get-PcmScheduleAccount -Name $Typed
        $account.Ok | Should -BeFalse
        $account.Message | Should -BeLike "*$Reason*"
        if ($Why -ne 'an unknown name') { $account.Message | Should -BeLike '*View-Only Organization Management*' }
    }
}

Describe 'Schedule: the proof from Exchange' {
    BeforeEach { Reset-ScheduleWork; Install-ScheduleMocks }
    It 'opens a Kerberos session to Exchange.Server with the credential, reads a recipient and closes the session' {
        $proof = Test-PcmExchangeAccess -Server 'ex01.contoso.com' -Credential $script:Cred
        $proof.Ok | Should -BeTrue
        $global:ScheduleFake.SessionArguments.ConfigurationName | Should -Be 'Microsoft.Exchange'
        $global:ScheduleFake.SessionArguments.ConnectionUri | Should -Be 'http://ex01.contoso.com/PowerShell/'
        $global:ScheduleFake.SessionArguments.Authentication | Should -Be 'Kerberos'
        $global:ScheduleFake.SessionArguments.UserName | Should -Be 'CONTOSO\svc-pra-collect'
        $global:ScheduleFake.Removed | Should -Be 1
    }
    It 'sends one command only to the Exchange session, which runs in no-language mode (no variable, no assignment, no literal)' {
        $null = Test-PcmExchangeAccess -Server 'ex01.contoso.com' -Credential $script:Cred
        $script = ([string]$global:ScheduleFake.ProofScript).Trim()
        $script | Should -Match '^Get-Recipient\b[^$=;|{}]*$'
    }
    It 'says in plain words <Name>, never repeats the password, and still closes the session' -TestCases @(
        @{ Name = 'that the account has no management role'; Where = 'SessionError'; Raised = 'Connecting to remote server ex01.contoso.com failed: The user "CONTOSO\svc-pra-collect" isn''t assigned to any management roles.'; Expected = '*not in any management role*View-Only Organization Management*wait a few minutes*' }
        @{ Name = 'that Exchange does not know the account'; Where = 'SessionError'; Raised = 'The user "CONTOSO\SRV01$" couldn''t be found.'; Expected = '*does not know this account as a user*not a computer account or SYSTEM*' }
        @{ Name = 'that the session opened but no recipient can be read'; Where = 'CommandError'; Raised = 'The term ''Get-Recipient'' is not recognized as a name of a cmdlet, function, script file, or executable program.'; Expected = '*may not read recipients*View-Only Organization Management*' }
        @{ Name = 'that the sign-in was refused'; Where = 'SessionError'; Raised = 'Logon failure: the user name or password is incorrect (password Zx9!Pw-Secret-7731).'; Expected = '*refused the sign-in*' }
        @{ Name = 'what failed when it is anything else'; Where = 'SessionError'; Raised = 'The WinRM client cannot process the request because Zx9!Pw-Secret-7731 timed out.'; Expected = '*Cannot open the remote session to Exchange on ex01.contoso.com*' }) {
        param($Name, $Where, $Raised, $Expected)
        $global:ScheduleFake[$Where] = $Raised
        $proof = Test-PcmExchangeAccess -Server 'ex01.contoso.com' -Credential $script:Cred
        $proof.Ok | Should -BeFalse
        $proof.Message | Should -BeLike $Expected
        $proof.Message | Should -Not -BeLike "*$script:Password*"
        $global:ScheduleFake.Removed | Should -Be $(if ($Where -eq 'CommandError') { 1 } else { 0 })
    }
}

Describe 'Schedule: the right to log on as a batch job' {
    BeforeEach { Reset-ScheduleWork; Install-ScheduleMocks; $script:Account = Get-PcmScheduleAccount -Name 'CONTOSO\svc-pra-collect'; $script:Folder = Join-Path $script:Work 'logs'; $null = New-Item -ItemType Directory -Path $script:Folder }
    It 'finds the right through the SID of the account or through one of its groups, and deletes the temporary file' -TestCases @(
        @{ Granted = '*S-1-5-32-544,*S-1-5-21-1-2-3-1105' }, @{ Granted = '*S-1-5-32-544,*S-1-5-21-1-2-3-513' }, @{ Granted = '*S-1-5-32-544,*S-1-1-0' }, @{ Granted = '*S-1-5-32-544,CONTOSO\SVC-PRA-COLLECT' }) {
        param($Granted)
        $global:ScheduleFake.Rights = "[Privilege Rights]`r`nSeBatchLogonRight = $Granted`r`n"
        (Test-PcmBatchLogonRight -Account $script:Account -Folder $script:Folder).State | Should -Be 'Granted'
        @(Get-ChildItem -LiteralPath $script:Folder).Count | Should -Be 0
        Test-Path -LiteralPath $global:ScheduleFake.RightsFile | Should -BeFalse
    }
    It 'says when the right is missing, with the way to give it' {
        $right = Test-PcmBatchLogonRight -Account $script:Account -Folder $script:Folder
        $right.State | Should -Be 'Missing'
        $right.Message | Should -BeLike '*Default Domain Controllers Policy*Log on as a batch job*gpupdate /force*'
    }
    It 'says when the right is denied' {
        $global:ScheduleFake.Rights = "[Privilege Rights]`r`nSeBatchLogonRight = *S-1-1-0`r`nSeDenyBatchLogonRight = *S-1-5-21-1-2-3-513`r`n"
        $right = Test-PcmBatchLogonRight -Account $script:Account -Folder $script:Folder
        $right.State | Should -Be 'Denied'
    }
    It 'does not fail when the rights cannot be read: unknown, with a reason' {
        Mock -ModuleName PRA.Collect Export-PcmUserRight { 5 }
        $right = Test-PcmBatchLogonRight -Account $script:Account -Folder $script:Folder
        $right.State | Should -Be 'Unknown'
        $right.Message | Should -BeLike '*could not be read*'
    }
}

Describe 'Schedule: a Preview' {
    BeforeAll {
        Reset-ScheduleWork; Install-ScheduleMocks
        $script:Before = @(Get-ChildItem -LiteralPath $script:Work -Recurse -File | ForEach-Object FullName)
        $script:Preview = Invoke-Schedule @{ Credential = $script:Cred }
    }
    It 'writes the banner in the house format: a framed card, then the rows in a fixed order' {
        $banner = Get-ScheduleCard $script:Preview Banner
        $banner.Box.Count | Should -Be 4
        $banner.Box[0].Text.Length | Should -Be 80
        $banner.Box[1].Text | Should -Match '^  │  \S+  PRA Cloud Mailbox +v\d+\.\d+\.\d+ · Nicolas Fabert  │$'
        @($banner.Box | ForEach-Object { $_.Color } | Select-Object -Unique) | Should -Be @($script:Theme.Accent)
        $labels = @(Get-ScheduleLabel $banner.Rows)
        @($labels | ForEach-Object Label) | Should -Be @('Action', 'Mode', 'Task', 'Config', 'Log')
        $labels[0].Text | Should -BeExactly 'Schedule'
        $labels[1].Text | Should -BeExactly 'Preview (nothing is changed)'
        $labels[2].Text | Should -BeExactly "PraCloudMailbox Collect $script:Dot every day at 02:00"
    }
    It 'numbers the five steps from 1 to the announced total, in the accent colour' {
        $steps = @(Get-ScheduleStep $script:Preview)
        @($steps | ForEach-Object Title) | Should -Be @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task')
        @($steps | ForEach-Object Number) | Should -Be @(1, 2, 3, 4, 5)
        @($steps | ForEach-Object Total | Select-Object -Unique) | Should -Be @(5)
        @($steps | ForEach-Object Color | Select-Object -Unique) | Should -Be @($script:Theme.Accent)
    }
    It 'checks the account and Exchange, skips the step of the batch right on a computer that is not a domain controller, and shows what would be registered' {
        $script:Preview.Text | Should -Match 'Account CONTOSO\\svc-pra-collect: a domain user\.'
        $script:Preview.Text | Should -Match 'Exchange on ex01\.contoso\.com accepted the account \(Kerberos\) and let it read recipients\.'
        $script:Preview.Text | Should -Match '(?m)^ {6}\S+ This computer is not a domain controller: the account needs no extra right here\.$'
        $script:Preview.Text | Should -Match 'Task "PraCloudMailbox Collect" in the root folder, every day at 02:00, as CONTOSO\\svc-pra-collect'
        $script:Preview.Text | Should -Match '(?m)^ {6}\S+ Preview: not registered \(add -Mode Apply\)\.$'
        @($script:Preview.Lines | Where-Object { $_.Text -match 'Preview: not registered' } | ForEach-Object Color) | Should -Be @($script:Theme.Dim)
    }
    It 'ends with the summary card: the title, the labelled rows, the same width as the banner, the next step, PASS and exit code 0' {
        $card = Get-ScheduleCard $script:Preview Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Schedule preview finished - nothing was changed ─'
        $card.Top.Text.Length | Should -Be 80
        $card.Bottom.Text.Length | Should -Be 80
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Ok)
        $labels = @(Get-ScheduleLabel $card.Rows)
        @($labels | ForEach-Object Label) | Should -Be @('Result', 'Task', 'Account', 'Runs at', 'Command', 'Exchange', 'Batch right', 'Report', 'Log', 'Duration', 'Next')
        foreach ($label in $labels) { $script:AllowedLabels | Should -Contain $label.Label }
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "0 ready $script:Dot 1 planned $script:Dot 0 error(s) $script:Dot 0 skipped"
        ($labels | Where-Object Label -eq 'Account').Text | Should -Be 'CONTOSO\svc-pra-collect'
        ($labels | Where-Object Label -eq 'Exchange').Text | Should -Be 'accepted by ex01.contoso.com (Kerberos)'
        ($labels | Where-Object Label -eq 'Batch right').Text | Should -Be 'not needed (not a domain controller)'
        (Get-ScheduleVerdict $script:Preview).Text | Should -Be '  RESULT: PASS'
        (Get-ScheduleVerdict $script:Preview).Color | Should -Be $script:Theme.Dim
        $script:Preview.Code | Should -Be 0
    }
    It 'shows the command of the task on two lines under its label, and proposes the Apply as the next step' {
        $script:Preview.Text | Should -Match ('(?m)^ {{5}}\S+ Command +{0}$' -f [regex]::Escape((Join-Path $PSHOME 'pwsh.exe')))
        $script:Preview.Text | Should -Match ('(?m)^ {{20}}-NoProfile -NonInteractive -File "{0}" -Action Collect -Mode Apply -Force -ConfigPath "{1}"$' -f [regex]::Escape((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')), [regex]::Escape($script:ConfigFile))
        $script:Preview.Text | Should -Match ('(?m)^ {{20}}\.\\Invoke-PraCloudMailbox\.ps1 -Action Schedule -Mode Apply -ConfigPath ''{0}''$' -f [regex]::Escape($script:ConfigFile))
    }
    It 'writes nothing: no task, no file other than the log, the transcript and the report' {
        Should -Invoke -CommandName Register-ScheduledTask -ModuleName PRA.Collect -Times 0 -Exactly
        $global:ScheduleFake.Registered | Should -BeNullOrEmpty
        $global:ScheduleFake.Tasks.Count | Should -Be 0
        $new = @(Get-ChildItem -LiteralPath $script:Work -Recurse -File | Where-Object { $script:Before -notcontains $_.FullName })
        @($new | Where-Object { $_.DirectoryName -notmatch '\\(logs|reports)$' }).Count | Should -Be 0
        @($new | Where-Object { $_.Name -like '*.inf' }).Count | Should -Be 0
    }
    It 'writes everything to the host: no warning, error, verbose or debug record' {
        @($script:Preview.Records | Where-Object { $_ -isnot [Management.Automation.InformationRecord] }).Count | Should -Be 0
    }
}

Describe 'Schedule: a Preview without an account, or in a window that is not elevated' {
    It 'does not test the account or Exchange without -Credential, and says so as information, not as a failure' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $run = Invoke-Schedule @{}
        $run.Code | Should -Be 0
        $run.Text | Should -Match 'Not tested: no -Credential given\. Give -Credential to check the account in the Preview'
        $run.Text | Should -Match 'Not tested: no -Credential given, so no session to Exchange was opened\.'
        Should -Invoke -CommandName New-PSSession -ModuleName PRA.Collect -Times 0 -Exactly
        $labels = @(Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Account').Text | Should -Be 'not given (the Apply asks for it)'
        ($labels | Where-Object Label -eq 'Exchange').Text | Should -Be 'not tested (no -Credential)'
        @($run.Lines | Where-Object { $_.Text -match 'Not tested' } | ForEach-Object Color | Select-Object -Unique) | Should -Be @('')
    }
    It 'warns in a Preview that is not elevated and goes on' {
        Reset-ScheduleWork -Elevated $false; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Credential = $script:Cred }
        $run.Code | Should -Be 0
        $run.Text | Should -Match 'This window is not elevated\. The Preview goes on, but -Mode Apply needs an elevated window'
        (Get-ScheduleCard $run Summary).Top.Text | Should -Match '1 warning\(s\) ─'
    }
}

Describe 'Schedule: an Apply' {
    BeforeAll {
        Reset-ScheduleWork; Install-ScheduleMocks
        # The package's own configuration file is the default one: the task then names no -ConfigPath.
        $script:DefaultConfigFile = Join-Path $script:Root 'config\PraCloudMailbox.config.psd1'
        $script:DefaultConfigBackup = [IO.File]::ReadAllBytes($script:DefaultConfigFile)
        Copy-Item -LiteralPath $script:ConfigFile -Destination $script:DefaultConfigFile -Force
        try { $script:Apply = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true } -DefaultConfig }
        finally { [IO.File]::WriteAllBytes($script:DefaultConfigFile, $script:DefaultConfigBackup) }
        $script:Registered = $global:ScheduleFake.Registered
    }
    It 'registers the task once, with the account and its password, as a standard user, replacing a task of the same name' {
        $global:ScheduleFake.RegisterCalls | Should -Be 1
        $script:Registered.TaskName | Should -BeExactly 'PraCloudMailbox Collect'
        $script:Registered.TaskPath | Should -BeExactly '\'
        $script:Registered.User | Should -BeExactly 'CONTOSO\svc-pra-collect'
        $script:Registered.Password | Should -BeExactly $script:Password
        $script:Registered.RunLevel | Should -Be 'Limited'
        $script:Registered.Force | Should -BeTrue
    }
    It 'runs the PowerShell that is running now, with the exact arguments, from the package folder' {
        $script:Registered.Action.Execute | Should -BeExactly (Join-Path $PSHOME 'pwsh.exe')
        $script:Registered.Action.Arguments | Should -BeExactly ('-NoProfile -NonInteractive -File "{0}" -Action Collect -Mode Apply -Force' -f (Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'))
        $script:Registered.Action.WorkingDirectory | Should -BeExactly $script:Root
    }
    It 'triggers it every day at the time of the configuration' {
        $script:Registered.Trigger.CimClass.CimClassName | Should -Be 'MSFT_TaskDailyTrigger'
        Get-LocalStart $script:Registered.Trigger | Should -Be '02:00'
        [int]$script:Registered.Trigger.DaysInterval | Should -Be 1
    }
    It 'writes the start as a local time, without Z, so that the task keeps the same hour on the wall when the clock changes' {
        [string]$script:Registered.Trigger.StartBoundary | Should -Match '^\d{4}-\d{2}-\d{2}T02:00:00$'
    }
    It 'tells the same in both guides for a task created by hand: the trigger is set to a local time' {
        foreach ($name in 'PraCloudMailbox-Guide.md', 'PraCloudMailbox-UserGuide.md') {
            $text = [IO.File]::ReadAllText((Join-Path $script:Root "docs\$name"))
            $text | Should -Match '(?m)^\$trigger\.StartBoundary = \(Get-Date -Hour 2 -Minute 0 -Second 0\)\.ToString\(''s''\)' -Because $name
        }
    }
    It 'starts it late when the computer was off, never twice at once, and stops it after 4 hours' {
        $script:Registered.Settings.StartWhenAvailable | Should -BeTrue
        [string]$script:Registered.Settings.CimInstanceProperties['MultipleInstances'].Value | Should -Match 'IgnoreNew|2'
        $script:Registered.Settings.ExecutionTimeLimit | Should -Be 'PT4H'
    }
    It 'ends done: one card, the task, the account, the time, the command, the proof, the next steps, PASS and exit code 0' {
        $script:Apply.Code | Should -Be 0
        $card = Get-ScheduleCard $script:Apply Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Schedule finished ─'
        $labels = @(Get-ScheduleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "1 done $script:Dot 0 pending $script:Dot 0 error(s) $script:Dot 0 skipped"
        ($labels | Where-Object Label -eq 'Task').Text | Should -Be 'PraCloudMailbox Collect (root folder)'
        ($labels | Where-Object Label -eq 'Runs at').Text | Should -Be 'every day at 02:00'
        $script:Apply.Text | Should -Match 'Task "PraCloudMailbox Collect" registered: every day at 02:00, as CONTOSO\\svc-pra-collect\.'
        $script:Apply.Text | Should -Match "(?m)^ {20}Start-ScheduledTask -TaskName 'PraCloudMailbox Collect'$"
        $script:Apply.Text | Should -Match '(?m)^ {20}\.\\Invoke-PraCloudMailbox\.ps1 -Action Check$'
        (Get-ScheduleVerdict $script:Apply).Text | Should -Be '  RESULT: PASS'
    }
    It 'gives the task the configuration file that was asked for when it is not the default one' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 0
        $global:ScheduleFake.Registered.Action.Arguments | Should -BeExactly ('-NoProfile -NonInteractive -File "{0}" -Action Collect -Mode Apply -Force -ConfigPath "{1}"' -f (Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'), $script:ConfigFile)
    }
    It 'finds the task again afterwards, so that the console can see it' {
        $task = Get-PcmScheduledCollectTask -Setting (Get-PcmScheduleSetting -Config (Import-PcmConfiguration -Path $script:ConfigFile -Root $script:Root) -ScriptPath 'C:\PRA\Invoke-PraCloudMailbox.ps1')
        $task.Account | Should -Be 'CONTOSO\svc-pra-collect'
        $task.At | Should -Be '02:00'
        $task.Command | Should -BeLike '*-Action Collect -Mode Apply -Force*'
    }
}

Describe 'Schedule: a task of the same name exists' {
    It 'shows in a Preview, in plain words, what it is now and that it will be replaced, whatever the letter case of its name' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake.Tasks = @([pscustomobject]@{ TaskName = 'PRACLOUDMAILBOX COLLECT'; TaskPath = '\'; Principal = [pscustomobject]@{ UserId = 'CONTOSO\old-account' }
                Triggers = @([pscustomobject]@{ StartBoundary = '2026-01-01T01:30:00' }); Actions = @([pscustomobject]@{ Execute = 'pwsh.exe'; Arguments = '-File "D:\old\Invoke-PraCloudMailbox.ps1" -Action Collect' }) })
        $run = Invoke-Schedule @{ Credential = $script:Cred }
        $run.Code | Should -Be 0
        $run.Text | Should -Match 'A task named "PRACLOUDMAILBOX COLLECT" exists and will be replaced\. It is now: account CONTOSO\\old-account, daily at 01:30, command pwsh\.exe -File "D:\\old\\Invoke-PraCloudMailbox\.ps1" -Action Collect'
        (Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows | ForEach-Object Label) | Should -Contain 'Replaces'
        $global:ScheduleFake.RegisterCalls | Should -Be 0
    }
    It 'replaces it in an Apply: one registration with -Force, and the new account and time' {
        Reset-ScheduleWork -Collect "Collect = @{ Schedule = @{ TaskName = 'pracloudmailbox collect'; At = '23:30' } }"; Install-ScheduleMocks
        $global:ScheduleFake.Tasks = @([pscustomobject]@{ TaskName = 'PraCloudMailbox Collect'; TaskPath = '\'; Principal = [pscustomobject]@{ UserId = 'CONTOSO\old-account' }
                Triggers = @([pscustomobject]@{ StartBoundary = '2026-01-01T01:30:00' }); Actions = @([pscustomobject]@{ Execute = 'pwsh.exe'; Arguments = '-File old.ps1' }) })
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 0
        $run.Text | Should -Match 'Task "pracloudmailbox collect" replaced: every day at 23:30, as CONTOSO\\svc-pra-collect\.'
        $global:ScheduleFake.RegisterCalls | Should -Be 1
        $global:ScheduleFake.Registered.Force | Should -BeTrue
        Get-LocalStart $global:ScheduleFake.Registered.Trigger | Should -Be '23:30'
        $global:ScheduleFake.Registered.User | Should -Be 'CONTOSO\svc-pra-collect'
    }
}

Describe 'Schedule: an account that cannot run the task' {
    It 'stops at the account, before Exchange is called and before anything is registered: <Why>' -TestCases @(
        @{ Why = 'a computer account'; Typed = 'CONTOSO\SRV01$'; Reason = 'is a computer account' }
        @{ Why = 'a gMSA'; Typed = 'contoso\gmsa-pra$'; Reason = 'is a group managed service account \(gMSA\)' }
        @{ Why = 'SYSTEM'; Typed = 'system'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'LOCAL SERVICE'; Typed = 'NT AUTHORITY\LOCAL SERVICE'; Reason = 'is a built-in identity of this computer' }
        @{ Why = 'a SID of NETWORK SERVICE'; Typed = 'S-1-5-20'; Reason = 'is a built-in identity of this computer' }) {
        param($Why, $Typed, $Reason)
        Reset-ScheduleWork; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Force = $true; Credential = (New-TestCredential $Typed 'x') }
        $run.Code | Should -Be 1
        $run.Text | Should -Match $Reason
        @(Get-ScheduleStep $run | ForEach-Object Title) | Should -Be @('Configuration', 'Account')
        Should -Invoke -CommandName New-PSSession -ModuleName PRA.Collect -Times 0 -Exactly
        Should -Invoke -CommandName Register-ScheduledTask -ModuleName PRA.Collect -Times 0 -Exactly
        $card = Get-ScheduleCard $run Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Schedule stopped - see the errors below ─'
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Fail)
        (Get-ScheduleVerdict $run).Text | Should -Be '  RESULT: FAIL'
    }
}

Describe 'Schedule: Exchange refuses the account' {
    It 'stops at the proof with the plain message, closes nothing that was not opened, registers nothing and proposes the same command again' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake.SessionError = 'The user "CONTOSO\svc-pra-collect" isn''t assigned to any management roles.'
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Match "Stopped at step 'Exchange access' \(operation Test-ExchangeAccess\): Exchange knows the account, but it is not in any management role\. Add it to the Exchange role group View-Only Organization Management"
        @(Get-ScheduleStep $run | ForEach-Object Title) | Should -Be @('Configuration', 'Account', 'Exchange access')
        Should -Invoke -CommandName Register-ScheduledTask -ModuleName PRA.Collect -Times 0 -Exactly
        $labels = @(Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "0 done $script:Dot 0 pending $script:Dot 1 error(s) $script:Dot 0 skipped"
        ($labels | Where-Object Label -eq 'Exchange').Text | Should -Be 'refused'
        ($labels | Where-Object Label -eq 'Next').Text | Should -Be 'The task was not registered, or it was left as it was. Fix the problem named above, then run the same command again:'
        $run.Text | Should -Match ('(?m)^ {{20}}\.\\Invoke-PraCloudMailbox\.ps1 -Action Schedule -Mode Apply -ConfigPath ''{0}''$' -f [regex]::Escape($script:ConfigFile))
        (Get-ScheduleVerdict $run).Text | Should -Be '  RESULT: FAIL'
    }
    It 'stops when Exchange.Server is empty, in plain words, at the first step' {
        Reset-ScheduleWork -Server ''; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Credential = $script:Cred }
        $run.Code | Should -Be 1
        @(Get-ScheduleStep $run | ForEach-Object Title) | Should -Be @('Configuration')
        $run.Text | Should -Match 'Exchange\.Server is empty\. The daily Collect reads Exchange through a remote session'
    }
}

Describe 'Schedule: the window and the account' {
    It 'stops an Apply in a window that is not elevated, with nothing else done' {
        Reset-ScheduleWork -Elevated $false; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Match 'Registering a scheduled task needs an elevated window\. Close this window, open PowerShell 7 with "Run as administrator"'
        @(Get-ScheduleStep $run | ForEach-Object Title) | Should -Be @('Configuration')
        Should -Invoke -CommandName Register-ScheduledTask -ModuleName PRA.Collect -Times 0 -Exactly
    }
    It 'stops an Apply that has no -Credential and cannot ask (-NonInteractive), in plain words, without asking' {
        Reset-ScheduleWork -Interactive $false; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Force = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Match 'No account was given, and this PowerShell was started with -NonInteractive, so it cannot ask for one\. Give -Credential'
        $global:ScheduleFake.Prompts | Should -Be 0
        Should -Invoke -CommandName Register-ScheduledTask -ModuleName PRA.Collect -Times 0 -Exactly
    }
    It 'asks for the account and the password when none is given, then registers' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake.PromptAnswer = $script:Cred
        $run = Invoke-Schedule @{ Mode = 'Apply'; Force = $true }
        $run.Code | Should -Be 0
        $global:ScheduleFake.Prompts | Should -Be 1
        $run.Text | Should -Match 'Asking for the account that runs the task'
        $global:ScheduleFake.Registered.User | Should -Be 'CONTOSO\svc-pra-collect'
        $global:ScheduleFake.Registered.Password | Should -BeExactly $script:Password
    }
    It 'ends cancelled when nothing is given at the prompt: exit code 1, no error counted, nothing registered' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake.PromptAnswer = $null
        $run = Invoke-Schedule @{ Mode = 'Apply'; Force = $true }
        $run.Code | Should -Be 1
        (Get-ScheduleCard $run Summary).Top.Text | Should -Match '^  ╭─ \S+  Schedule cancelled - nothing was written ─'
        (Get-ScheduleVerdict $run).Text | Should -Be '  RESULT: CANCELLED'
        $run.Text | Should -Not -Match 'Stopped at step'
        $global:ScheduleFake.RegisterCalls | Should -Be 0
    }
    It 'ends cancelled when the operator answers no to the confirmation, and says to answer yes next time' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ConsoleAnswersNo = $true
        try { $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred } } finally { $global:ConsoleAnswersNo = $false }
        $run.Code | Should -Be 1
        $labels = @(Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "0 done $script:Dot 0 pending $script:Dot 0 error(s) $script:Dot 0 skipped"
        ($labels | Where-Object Label -eq 'Next').Text | Should -Be 'Nothing was written. To run it, give the same command again and answer yes to the confirmation:'
        (Get-ScheduleVerdict $run).Text | Should -Be '  RESULT: CANCELLED'
        $global:ScheduleFake.RegisterCalls | Should -Be 0
    }
}

Describe 'Schedule: a domain controller' {
    It 'checks the right to log on as a batch job and goes on when it is granted' {
        Reset-ScheduleWork -DomainRole 5 -Rights "[Privilege Rights]`r`nSeBatchLogonRight = *S-1-5-32-544,*S-1-5-21-1-2-3-513`r`n"; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 0
        @(Get-ScheduleStep $run | ForEach-Object Title) | Should -Be @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task')
        $run.Text | Should -Match 'CONTOSO\\svc-pra-collect has the right "Log on as a batch job" on this domain controller\.'
        (Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows | Where-Object Label -eq 'Batch right').Text | Should -Be 'granted'
        @(Get-ChildItem -LiteralPath (Join-Path $script:Work 'logs') -Filter 'PRA_rights_*').Count | Should -Be 0
    }
    It 'stops when the right is missing, before the registration, and says how to give it' {
        Reset-ScheduleWork -DomainRole 4; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Match 'does not have the right "Log on as a batch job" on this domain controller\. Add it in the Default Domain Controllers Policy'
        Should -Invoke -CommandName Register-ScheduledTask -ModuleName PRA.Collect -Times 0 -Exactly
        (Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows | Where-Object Label -eq 'Batch right').Text | Should -Be 'missing'
    }
    It 'warns, and goes on, when the right cannot be read' {
        Reset-ScheduleWork -DomainRole 4; Install-ScheduleMocks
        Mock -ModuleName PRA.Collect Export-PcmUserRight { 1 }
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 0
        @($run.Lines | Where-Object { $_.Text -match '^ {6}\S+ The right "Log on as a batch job" could not be read' } | ForEach-Object Color) | Should -Be @($script:Theme.Warn)
        (Get-ScheduleCard $run Summary).Top.Text | Should -Match 'Schedule finished .+ 1 warning\(s\) ─'
        (Get-ScheduleLabel (Get-ScheduleCard $run Summary).Rows | Where-Object Label -eq 'Batch right').Text | Should -Be 'not known: check it by hand'
    }
}

Describe 'Schedule: the password never leaves the credential' {
    It 'is in no console line, log, transcript or report of an Apply, and not in the result object' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true; PassThru = $true }
        $run.Code | Should -Be 0
        $run.Text | Should -Not -BeLike "*$script:Password*"
        @(Find-InWorkFiles $script:Password) | Should -BeNullOrEmpty
        @(Get-ChildItem -LiteralPath $script:Work -Recurse -File -Include '*.transcript.txt').Count | Should -Be 1
        ($run.Records | Where-Object { $_ -is [pscustomobject] } | ConvertTo-Json -Depth 4) | Should -Not -BeLike "*$script:Password*"
    }
    It 'is in no console line, log or report of a run that fails, whatever Exchange or the Task Scheduler repeat' -TestCases @(
        @{ Where = 'SessionError'; Text = 'The WinRM client failed for Zx9!Pw-Secret-7731 at ex01.' }
        @{ Where = 'RegisterError'; Text = 'The task image is corrupt or has been tampered with (Zx9!Pw-Secret-7731).' }) {
        param($Where, $Text)
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake[$Where] = $Text
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true; PassThru = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Not -BeLike "*$script:Password*"
        # The host writes the message of an error raised inside a cmdlet to the transcript as it is, and the real cmdlets never repeat the password: the files that the tool writes itself are checked.
        @(Find-InWorkFiles $script:Password | Where-Object { $_ -notlike '*.transcript.txt:*' }) | Should -BeNullOrEmpty
        ($run.Records | Where-Object { $_ -is [pscustomobject] } | ConvertTo-Json -Depth 4) | Should -Not -BeLike "*$script:Password*"
    }
    It 'is in no file of a run that fails with the messages of the real cmdlets, the transcript included' -TestCases @(
        @{ Where = 'SessionError'; Text = 'The WinRM client cannot process the request because the server name cannot be resolved.' }
        @{ Where = 'RegisterError'; Text = 'No mapping between account names and security IDs was done.' }) {
        param($Where, $Text)
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake[$Where] = $Text
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Not -BeLike "*$script:Password*"
        @(Find-InWorkFiles $script:Password) | Should -BeNullOrEmpty
    }
    It 'reports a registration that Windows refuses, in plain words' -TestCases @(
        @{ Raised = 'Access is denied.'; Expected = 'This window is not elevated' }
        @{ Raised = 'The user name or password is incorrect.'; Expected = 'Windows refused the password of CONTOSO\svc-pra-collect.' }
        @{ Raised = 'No mapping between account names and security IDs was done.'; Expected = 'Windows does not know the account CONTOSO\svc-pra-collect.' }) {
        param($Raised, $Expected)
        Reset-ScheduleWork; Install-ScheduleMocks
        $global:ScheduleFake.RegisterError = $Raised
        $run = Invoke-Schedule @{ Mode = 'Apply'; Credential = $script:Cred; Force = $true }
        $run.Code | Should -Be 1
        $run.Text | Should -Match ([regex]::Escape("The task 'PraCloudMailbox Collect' could not be registered: $Expected"))
        $run.Text | Should -Match 'The task was not registered, or it was left as it was\.'
    }
}

Describe 'Schedule: nothing is case sensitive, and the options are checked' {
    It 'takes the action, the mode, the time, the account and the task name in any letter case, and shows the official spelling' {
        Reset-ScheduleWork -Collect "Collect = @{ Schedule = @{ TaskName = 'Nightly COLLECT' } }"; Install-ScheduleMocks
        $cred = New-TestCredential 'cOnToSo\SVC-pra-COLLECT' $script:Password
        $run = Invoke-Schedule @{ Action = 'sChEdUlE'; Mode = 'aPPly'; Credential = $cred; At = '5:07'; Force = $true }
        $run.Code | Should -Be 0
        $banner = @(Get-ScheduleLabel (Get-ScheduleCard $run Banner).Rows)
        ($banner | Where-Object Label -eq 'Action').Text | Should -BeExactly 'Schedule'
        ($banner | Where-Object Label -eq 'Mode').Text | Should -BeExactly 'Apply (the task is registered)'
        ($banner | Where-Object Label -eq 'Task').Text | Should -BeExactly "Nightly COLLECT $script:Dot every day at 05:07"
        $global:ScheduleFake.Registered.TaskName | Should -BeExactly 'Nightly COLLECT'
        $global:ScheduleFake.Registered.User | Should -BeExactly 'CONTOSO\svc-pra-collect'
        Get-LocalStart $global:ScheduleFake.Registered.Trigger | Should -Be '05:07'
    }
    It 'keeps -At in the command that it proposes' {
        Reset-ScheduleWork; Install-ScheduleMocks
        $run = Invoke-Schedule @{ At = '03:15' }
        $run.Code | Should -Be 0
        $run.Text | Should -Match ('(?m)^ {{20}}\.\\Invoke-PraCloudMailbox\.ps1 -Action Schedule -Mode Apply -At 03:15 -ConfigPath ''{0}''$' -f [regex]::Escape($script:ConfigFile))
    }
    It 'refuses a time that is not a time of day, and the options of the other actions, in plain words' -TestCases @(
        @{ Arguments = @{ At = '25:00' }; Message = "-At '25:00' is not a time of day" }
        @{ Arguments = @{ Identity = 'adele@contoso.com' }; Message = '-Identity is not used by Schedule, which only registers the daily Collect task' }
        @{ Arguments = @{ Scope = 'UsersOnly'; MaxObjects = 3 }; Message = '-MaxObjects, -Scope are not used by Schedule' }
        @{ Arguments = @{ Action = 'Collect'; Credential = $null }; Message = '-Credential and -At are for Schedule' }
        @{ Arguments = @{ Action = 'Check'; At = '02:00' }; Message = '-Credential and -At are for Schedule' }) {
        param($Arguments, $Message)
        Reset-ScheduleWork; Install-ScheduleMocks
        if ($Arguments.ContainsKey('Credential')) { $Arguments.Credential = $script:Cred }
        $run = Invoke-Schedule $Arguments
        $run.Code | Should -Be 1
        $run.Text | Should -Match ([regex]::Escape($Message))
        $global:ScheduleFake.RegisterCalls | Should -Be 0
    }
    It 'is in the usage and in the parameters of the entry script' {
        $usage = [IO.File]::ReadAllText((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'))
        $usage | Should -Match 'Give the task: -Action Collect, Check, Convert, Recover or Schedule \(Preview is the default mode\)'
        $command = Get-Command (Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1')
        @($command.Parameters['Action'].Attributes | Where-Object { $_ -is [Management.Automation.ValidateSetAttribute] })[0].ValidValues | Should -Contain 'Schedule'
        $command.Parameters['Credential'].ParameterType | Should -Be ([pscredential])
        $command.Parameters['At'].ParameterType | Should -Be ([string])
    }
}
