<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the console of the command line (Pester 5+, offline): the house format of the banner, the numbered steps, the item lines, the
    errors, the summary card, the next steps, the verdict line and the exit codes.
.DESCRIPTION
    Runs the real entry script (Invoke-PraCloudMailbox.ps1, without its package integrity bootstrap and its module imports, which would drop the replacements below) in this
    process, against the in-memory tenant of tests\FakeTenant.ps1, and reads what it writes to the host: the text and the colour of every line, and the stream of every record.
    Only the connections, the roles of the token, Entra Connect and the waits are replaced, as in the end-to-end tests. PowerShell 7 only.
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
    . (Join-Path $PSScriptRoot 'FakeTenant.ps1')
    $script:Theme = & (Get-Module PRA.Common) { $script:Theme }
    $script:ToolVersion = Get-PcmToolVersion

    # The entry script as a command that runs here: no bootstrap, no module import, and a way to make the user waves small and to answer no to the confirmation.
    $text = [IO.File]::ReadAllText((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'))
    $start = $text.IndexOf('# >>> PRA package integrity bootstrap'); $end = $text.IndexOf('# <<< PRA package integrity bootstrap')
    $text = $text.Remove($start, $end - $start)
    $text = [regex]::Replace($text, '(?m)^Import-Module \(Join-Path \$PSScriptRoot ''PraCloudMailbox\.psd1''\).*$', '')
    $text = [regex]::Replace($text, '(?m)^if \(\$Action -in @\(''Collect'', ''Schedule''\)\) \{ Import-Module.*\r?\n.*Import-Module.*$', '')
    $text = $text.Replace('$context.Approval = New-PcmApprovalCallback $caller $context', '$context.Approval = New-PcmApprovalCallback $caller $context; if ($global:ConsoleWaveSize) { $context.UserWaveSize = [int]$global:ConsoleWaveSize }')
    $text = $text.Replace('$caller = $PSCmdlet', '$caller = if ($global:ConsoleAnswersNo) { $global:ConsoleCaller } else { $PSCmdlet }')
    $script:Entry = Join-Path $script:Root 'Invoke-Console.ps1'
    [IO.File]::WriteAllText($script:Entry, $text, [Text.UTF8Encoding]::new($true))
    $global:ConsoleWaveSize = 0; $global:ConsoleAnswersNo = $false
    $global:ConsoleCaller = [pscustomobject]@{}
    $global:ConsoleCaller | Add-Member -MemberType ScriptMethod -Name ShouldContinue -Value { param($q, $t) $false }
    $global:ConsoleCaller | Add-Member -MemberType ScriptMethod -Name ShouldProcess -Value { param($t, $o) $true }

    function global:Start-Sleep { Step-PraFake }
    function global:Connect-PcmCloud { 'Microsoft Graph and Exchange Online (fake tenant)' }
    function global:Disconnect-PcmCloud { }
    function global:Connect-PcmCompliance { }
    function global:Get-PcmGraphRole { @('User.ReadWrite.All', 'User-OnPremisesSyncBehavior.ReadWrite.All', 'LicenseAssignment.ReadWrite.All', 'Organization.Read.All', 'GroupMember.ReadWrite.All', 'Group-OnPremisesSyncBehavior.ReadWrite.All') }
    function global:Get-PcmCloudModuleReport { param($MinExo) [pscustomobject]@{ Ok = $true; Graph = [pscustomobject]@{ Version = '2.41.1' }; Exo = [pscustomobject]@{ Version = '3.10.1' }; Messages = @() } }
    function global:Invoke-PcmEntraConnect { param([hashtable]$Context, [string]$Operation, [object]$Caller) $null = Invoke-PraFakeEntraConnect -Operation $Operation; 'SyncCycleEnabled=True InProgress=False' }
    function global:Wait-PcmCondition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds }
    & (Get-Module PRA.Cloud) { function script:Wait-PcmCondition { param([scriptblock]$Test, [int]$TimeoutSeconds, [int]$IntervalSeconds) Invoke-PraFakeWait -Test $Test -TimeoutSeconds $TimeoutSeconds -IntervalSeconds $IntervalSeconds } }

    function script:Reset-ConsoleWork {
        <# A new tenant (adele, megan and alex with a mailbox plan to give, and two shared mailboxes), its snapshot and its configuration file. #>
        param([int]$ExtraUsers = 0, [int]$ExtraShared = 0)
        $script:Work = Join-Path $TestDrive ([guid]::NewGuid().ToString('N')); $null = New-Item -ItemType Directory -Path $script:Work, (Join-Path $script:Work 'data')
        $script:ConfigFile = Join-Path $script:Work 'console.config.psd1'
        $work = $script:Work
        Set-Content -LiteralPath $script:ConfigFile -Encoding UTF8 -Value @"
@{
    Environment = 'PROD'
    Scope = @{ Mode = 'OU'; SearchBase = 'OU=Paris,DC=contoso,DC=com' }
    Store = @{ Path = '$work\data\PraCloudMailbox.db'; JournalPath = '$work\data\PraCloudMailbox-journal.db'; BackupFolder = '$work\data\backup' }
    Cloud = @{ TenantId = '33333333-3333-3333-3333-333333333333'; Organization = 'contoso.onmicrosoft.com'; AppId = 'app'; CertificateThumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'; DefaultUsageLocation = 'FR' }
    Polling = @{ IntervalSeconds = 20; MailboxTimeoutMinutes = 30; SyncTimeoutMinutes = 30; HoldTimeoutMinutes = 30 }
    Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '11111111-1111-1111-1111-111111111111' }; Shared = @{ SkuPartNumber = 'E5' } }
    Retention = @{ HoldPolicy = 'PRA-Recover-Hold' }
    EntraConnect = @{ Mode = 'Remoting'; Server = 'aadc01.contoso.com' }
    Logging = @{ Folder = '$work\logs' }
    Report = @{ Folder = '$work\reports' }
}
"@
        New-PraFakeTenant
        $global:PraFake.CasePolicy.Name = 'PRA-Recover-Hold'
        $global:PraFake.Group.Name = 'LIC-Exchange-Online'
        $script:People = [ordered]@{}
        $entries = @(@{ Upn = 'adele.vance@contoso.com'; Teams = $true }, @{ Upn = 'megan.bowen@contoso.com'; Teams = $true }, @{ Upn = 'alex.wilber@contoso.com'; Teams = $false }
            @{ Upn = 'accounting@contoso.com'; Kind = 'Shared' }, @{ Upn = 'helpdesk@contoso.com'; Kind = 'Shared' })
        if ($ExtraUsers) { $entries += @(1..$ExtraUsers | ForEach-Object { @{ Upn = ('user{0}@contoso.com' -f $_); Teams = $true } }) }
        if ($ExtraShared) { $entries += @(1..$ExtraShared | ForEach-Object { @{ Upn = ('shared{0}@contoso.com' -f $_); Kind = 'Shared' } }) }
        foreach ($entry in $entries) {
            $kind = if ($entry.ContainsKey('Kind')) { $entry.Kind } else { 'User' }
            $script:People[$entry.Upn] = Add-PraFakeUser -Upn $entry.Upn -Kind $kind -Component $(if ($entry.ContainsKey('Teams') -and $entry.Teams) { [guid]::NewGuid().ToString() } else { '' }) -UsageLocation $(if ($kind -eq 'Shared') { '' } else { 'FR' })
        }
        $script:People['alex.wilber@contoso.com'].UsageLocation = ''
        Write-ConsoleSnapshot
    }

    function script:Write-ConsoleSnapshot {
        $byUpn = @{}; foreach ($ad in $global:PraFake.Ad) { $byUpn[$ad.Upn] = $ad }
        $db = Join-Path $script:Work 'data\PraCloudMailbox.db'
        $connection = Open-PcmStore -Path $db -Root $script:Root
        try {
            $id = New-PcmSnapshot $connection @{ RunId = [guid]::NewGuid().ToString('N'); Environment = 'PROD'; ToolVersion = $script:ToolVersion; ExchangeServer = 'EX01.contoso.com'; ScopeJson = '{"Mode":"OU"}' }
            $rows = @(foreach ($upn in $script:People.Keys) {
                    $ad = $byUpn[$upn]
                    [ordered]@{ object_guid = $ad.ObjectGuid; kind = $ad.Kind; user_principal_name = $upn; primary_smtp_address = $upn; sam_account_name = ($upn -split '@')[0]
                        exchange_guid = $ad.OnPremGuid; immutable_id = $ad.ImmutableId; recipient_type_details = "$($ad.Kind)Mailbox" }
                })
            $null = Add-PcmRow $connection mailbox $id $rows
            $grant = { param($mailbox, $upn, $right, $automap) [ordered]@{ mailbox_guid = $byUpn[$mailbox].ObjectGuid; access_right = $right; trustee = "CONTOSO\$(($upn -split '@')[0])"; trustee_kind = 'User'
                    trustee_upn = $upn; trustee_guid = $byUpn[$upn].ObjectGuid; trustee_name = ($upn -split '@')[0]; resolved = 1; auto_mapping = $automap } }
            $null = Add-PcmRow $connection permission $id @((& $grant 'accounting@contoso.com' 'adele.vance@contoso.com' 'FullAccess' 1), (& $grant 'helpdesk@contoso.com' 'alex.wilber@contoso.com' 'FullAccess' 1))
            $wave = 'b0000000-0000-0000-0000-000000000001'
            $null = Add-PcmRow $connection target_group $id @([ordered]@{ group_guid = $wave; name = 'PRA-Wave1'; sam_account_name = 'PRA-Wave1'; distinguished_name = 'CN=PRA-Wave1,OU=Groups,DC=contoso,DC=com'; primary_smtp_address = ''; display_name = 'PRA-Wave1'; member_count = 5 })
            $members = @(foreach ($upn in 'adele.vance@contoso.com', 'megan.bowen@contoso.com', 'alex.wilber@contoso.com') {
                    [ordered]@{ group_guid = $wave; member_guid = $byUpn[$upn].ObjectGuid; member_class = 'User'; member_name = ($upn -split '@')[0]; member_upn = $upn; member_smtp = $upn; member_dn = ''; via_group_guid = ''; depth = 1 }
                })
            $members += [ordered]@{ group_guid = $wave; member_guid = 'c0000000-0000-0000-0000-000000000001'; member_class = 'Contact'; member_name = 'Print Room'; member_upn = ''; member_smtp = 'print.room@contoso.com'; member_dn = ''; via_group_guid = ''; depth = 1 }
            $members += [ordered]@{ group_guid = $wave; member_guid = 'c0000000-0000-0000-0000-000000000002'; member_class = 'Computer'; member_name = 'KIOSK-01'; member_upn = ''; member_smtp = ''; member_dn = ''; via_group_guid = ''; depth = 2 }
            $null = Add-PcmRow $connection target_member $id $members
            Set-PcmSnapshotStatus $connection $id Complete @{ mailbox = $rows.Count }
        } finally { Close-PcmStore $connection }
        $null = Write-PcmHashFile -Path $db
    }

    function script:Invoke-Console {
        <# One run of the entry script. Lines: the text and the colour of each console line; Records: everything the run wrote; Rows: the labelled rows of the banner and of the summary card. #>
        param([hashtable]$Arguments = @{}, [switch]$NoConfig)
        if (-not $NoConfig -and -not $Arguments.ContainsKey('ConfigPath')) { $Arguments = $Arguments + @{ ConfigPath = $script:ConfigFile } }
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

    function script:Get-ConsoleCard {
        <# The banner (the box and its rows) or the summary card (its top line, its rows and its bottom line) of a run. #>
        param([Parameter(Mandatory)][object]$Run, [ValidateSet('Banner', 'Summary')][string]$Kind)
        $lines = @($Run.Lines)
        $top = @(0..($lines.Count - 1) | Where-Object { $lines[$_].Text -match $(if ($Kind -eq 'Banner') { '^  ╭─+╮$' } else { '^  ╭─ .+ ─+╮$' }) }) | Select-Object -First 1
        $bottomAt = @($top..($lines.Count - 1) | Where-Object { $lines[$_].Text -match '^  ╰─+╯$' }) | Select-Object -First 1
        if ($Kind -eq 'Banner') {
            # The rows follow the box and end at the first blank line.
            $rows = @(); $at = $bottomAt + 1
            while ($at -lt $lines.Count -and $lines[$at].Text.Trim()) { $rows += $lines[$at]; $at++ }
            return [pscustomobject]@{ Top = $lines[$top]; Bottom = $lines[$bottomAt]; Rows = $rows; Box = @($lines[$top..$bottomAt]) }
        }
        [pscustomobject]@{ Top = $lines[$top]; Bottom = $lines[$bottomAt]; Rows = @($lines[($top + 1)..($bottomAt - 1)]); Box = @($lines[$top..$bottomAt]) }
    }

    function script:Get-ConsoleLabel {
        <# The label of each row (the continuation lines of a row have none): 5 spaces, an icon, one space, the label padded to 11, one space, the text. #>
        param([object[]]$Rows)
        foreach ($row in $Rows) {
            $match = [regex]::Match($row.Text, '^ {5}(?<icon>\S+) (?<label>.{11}) (?<text>\S.*)$')
            if ($match.Success) { [pscustomobject]@{ Label = $match.Groups['label'].Value.TrimEnd(); Text = $match.Groups['text'].Value; Color = $row.Color } }
        }
    }

    function script:Get-ConsoleVerdict {
        <# The verdict line: the last line with text (the run ends with a blank line). #>
        param([object]$Run)
        @($Run.Lines | Where-Object { $_.Text })[-1]
    }

    function script:Get-ConsoleStep {
        param([object]$Run)
        @($Run.Lines | Where-Object { $_.Text -match '^ {2}\s*\d+/\d+  \S+ .+$' } | ForEach-Object {
                $m = [regex]::Match($_.Text, '^ {2}\s*(?<n>\d+)/(?<total>\d+)  (?<icon>\S+) (?<title>.+)$')
                [pscustomobject]@{ Number = [int]$m.Groups['n'].Value; Total = [int]$m.Groups['total'].Value; Title = $m.Groups['title'].Value; Color = $_.Color; Text = $_.Text } })
    }

    $script:AllowedLabels = @('Action', 'Mode', 'Scope', 'Tenant', 'Database', 'Integrity', 'Config', 'Settings', 'Log', 'Result', 'Batch', 'Snapshot', 'Objects', 'Error', 'Warning', 'Backup', 'Report', 'Duration', 'Next')
    $script:Dot = [string][char]0x00B7
}

AfterAll {
    foreach ($name in 'Start-Sleep', 'Connect-PcmCloud', 'Disconnect-PcmCloud', 'Connect-PcmCompliance', 'Get-PcmGraphRole', 'Get-PcmCloudModuleReport', 'Invoke-PcmEntraConnect', 'Wait-PcmCondition') {
        Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue
    }
    foreach ($name in $global:PraFakeCommands) { Remove-Item -Path "Function:$name" -ErrorAction SilentlyContinue }
    Remove-Variable -Name PraFake, PraFakeCommands, ConsoleWaveSize, ConsoleAnswersNo, ConsoleCaller -Scope Global -ErrorAction SilentlyContinue
    if ($null -eq $script:AnsiBefore) { Remove-Item Env:\PRA_ANSI -ErrorAction SilentlyContinue } else { $env:PRA_ANSI = $script:AnsiBefore }
    if ($null -eq $script:IconsBefore) { Remove-Item Env:\PRA_ICONS -ErrorAction SilentlyContinue } else { $env:PRA_ICONS = $script:IconsBefore }
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Console: usage' {
    It 'stops at once without -Action: one error line with the usage, exit code 1, no banner' {
        $run = Invoke-Console -NoConfig
        $run.Code | Should -Be 1
        @($run.Records).Count | Should -Be 1
        $run.Records[0] | Should -BeOfType ([Management.Automation.ErrorRecord])
        $run.Records[0].Exception.Message | Should -Match '^-Action is missing: nothing was done\. Give the task: -Action Collect, Check, Convert, Recover or Schedule \(Preview is the default mode\)\. Example: \.\\Invoke-PraCloudMailbox\.ps1 -Action Check\.'
        $run.Lines.Count | Should -Be 0
    }
}

Describe 'Console: a Check' {
    BeforeAll {
        Reset-ConsoleWork
        $script:Check = Invoke-Console @{ Action = 'check' }
    }
    It 'writes the banner: a framed title card, then the rows in a fixed order, each label padded to 11 characters' {
        $banner = Get-ConsoleCard $script:Check Banner
        $banner.Box.Count | Should -Be 4
        $banner.Box[0].Text.Length | Should -Be 80
        $banner.Box[1].Text | Should -Match '^  │  \S+  PRA Cloud Mailbox +v\d+\.\d+\.\d+ · Nicolas Fabert  │$'
        $banner.Box[2].Text | Should -Match '^  │     Exchange DR · scenario 2 · on-premises lost → Exchange Online +│$'
        @($banner.Box | ForEach-Object { $_.Color } | Select-Object -Unique) | Should -Be @($script:Theme.Accent)
        $labels = @(Get-ConsoleLabel $banner.Rows)
        @($labels | ForEach-Object Label) | Should -Be @('Action', 'Mode', 'Scope', 'Tenant', 'Database', 'Config', 'Log')
        $labels[0].Text | Should -Be 'Check'
        $labels[1].Text | Should -Be 'Read-only'
    }
    It 'numbers the steps from 1 to the announced total, one header line each, in the accent colour' {
        $steps = @(Get-ConsoleStep $script:Check)
        @($steps | ForEach-Object Title) | Should -Be @('Collect snapshot', 'Microsoft 365 connection', 'Tenant prerequisites', 'Objects')
        @($steps | ForEach-Object Number) | Should -Be @(1, 2, 3, 4)
        @($steps | ForEach-Object Total | Select-Object -Unique) | Should -Be @(4)
        @($steps | ForEach-Object Color | Select-Object -Unique) | Should -Be @($script:Theme.Accent)
        foreach ($step in $steps) { $step.Text | Should -Match '^ {2}\s*\d+/4  \S+ ' }
    }
    It 'writes the result lines with 6 spaces, an icon and the text; a problem is coloured, a good line is not' {
        $items = @($script:Check.Lines | Where-Object { $_.Text -match '^ {6}\S+ ' })
        $items.Count | Should -BeGreaterThan 5
        @($items | Where-Object { $_.Text -match 'Snapshot file verified' })[0].Color | Should -Be ''
    }
    It 'ends with the summary card: the title, the rows, the same width as the banner, then the verdict line' {
        $card = Get-ConsoleCard $script:Check Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Check finished '
        $card.Top.Text.Length | Should -Be 80
        $card.Bottom.Text.Length | Should -Be 80
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Ok)
        $labels = @(Get-ConsoleLabel $card.Rows)
        @($labels | ForEach-Object Label) | Should -Be @('Result', 'Snapshot', 'Integrity', 'Report', 'Log', 'Duration', 'Next')
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "5 as expected $script:Dot 0 not as expected $script:Dot 0 skipped"
        (Get-ConsoleVerdict $script:Check).Text | Should -Be '  RESULT: PASS'
        (Get-ConsoleVerdict $script:Check).Color | Should -Be $script:Theme.Dim
        $script:Check.Code | Should -Be 0
    }
    It 'keeps every label of the banner and of the summary within the 11 characters of the label column' {
        foreach ($kind in 'Banner', 'Summary') {
            $labels = @(Get-ConsoleLabel (Get-ConsoleCard $script:Check $kind).Rows)
            foreach ($label in $labels) { $script:AllowedLabels | Should -Contain $label.Label }
        }
    }
    It 'always prints the next step: the exact command as run from the package folder (.\Invoke-PraCloudMailbox.ps1) and the configuration file that is not the default' {
        $next = @($script:Check.Lines | Where-Object { $_.Text -match '^ {20}\.\\Invoke-PraCloudMailbox\.ps1 ' })
        $next.Count | Should -Be 1
        $next[0].Text | Should -Be ('                    .\Invoke-PraCloudMailbox.ps1 -Action Convert -ConfigPath ''{0}''' -f $script:ConfigFile)
    }
    It 'writes everything to the host: no warning, error, verbose or debug record' {
        @($script:Check.Records | Where-Object { $_ -isnot [Management.Automation.InformationRecord] }).Count | Should -Be 0
    }
}

Describe 'Console: a Preview of Convert on a group' {
    BeforeAll {
        Reset-ConsoleWork
        $script:Preview = Invoke-Console @{ Action = 'cOnVeRt'; Group = 'PRA-WAVE1'; Mode = 'preVIEW' }
    }
    It 'shows the values in their official spelling, the group as typed, the target line and the skipped members' {
        $banner = @(Get-ConsoleLabel (Get-ConsoleCard $script:Preview Banner).Rows)
        ($banner | Where-Object Label -eq 'Action').Text | Should -BeExactly 'Convert'
        ($banner | Where-Object Label -eq 'Mode').Text | Should -BeExactly 'Preview (nothing is changed)'
        ($banner | Where-Object Label -eq 'Scope').Text | Should -BeExactly 'group PRA-WAVE1'
        $script:Preview.Text | Should -Match 'Group PRA-Wave1 \(captured by Collect, snapshot \d+\): 5 member\(s\) including nested groups, 2 skipped'
        @($script:Preview.Lines | Where-Object { $_.Text -match 'not taken, a (contact|computer)' }).Count | Should -Be 2
    }
    It 'words the result of a Preview as ready, planned, errors and skipped, and counts the skipped members' {
        $card = Get-ConsoleCard $script:Preview Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Convert preview finished - nothing was changed '
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "0 ready $script:Dot 3 planned $script:Dot 0 error(s) $script:Dot 2 skipped"
        ($labels | Where-Object Label -eq 'Objects').Text | Should -Match '^Group PRA-Wave1 '
    }
    It 'proposes the Apply with the group as typed and every step of the Preview marked as not run' {
        $script:Preview.Text | Should -Match ('\.\\Invoke-PraCloudMailbox\.ps1 -Action Convert -Mode Apply -Group ''PRA-WAVE1'' -ConfigPath ''{0}''' -f [regex]::Escape($script:ConfigFile))
        @($script:Preview.Lines | Where-Object { $_.Text -match '^ {6}\S+ Preview: not run\.$' }).Count | Should -Be 3
        @($script:Preview.Lines | Where-Object { $_.Text -match 'Preview: not run' } | ForEach-Object Color | Select-Object -Unique) | Should -Be @($script:Theme.Dim)
        $script:Preview.Code | Should -Be 0
    }
}

Describe 'Console: a batch end to end' {
    BeforeAll {
        Reset-ConsoleWork -ExtraUsers 2
        $global:ConsoleWaveSize = 2
        $script:Convert = Invoke-Console @{ Action = 'Convert'; Mode = 'APPLY'; Force = $true }
        $global:ConsoleWaveSize = 0
        $script:Batch = [regex]::Match($script:Convert.Text, '(?m)^ {5}\S+ Batch +(?<id>[0-9a-f]{8})$').Groups['id'].Value
    }
    It 'writes the Convert in 7 steps and ends done: one card, exit code 0, PASS' {
        $steps = @(Get-ConsoleStep $script:Convert)
        @($steps | ForEach-Object Number) | Should -Be @(1..7)
        @($steps | ForEach-Object Total | Select-Object -Unique) | Should -Be @(7)
        $card = Get-ConsoleCard $script:Convert Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Convert finished ─'
        (Get-ConsoleLabel $card.Rows | Where-Object Label -eq 'Result').Text | Should -Be "7 done $script:Dot 0 pending $script:Dot 0 error(s) $script:Dot 0 skipped"
        $script:Batch | Should -Match '^[0-9a-f]{8}$'
        $script:Convert.Code | Should -Be 0
        (Get-ConsoleVerdict $script:Convert).Text | Should -Be '  RESULT: PASS'
    }
    It 'writes the progress of the waves as result lines of the same format, with the time left once an object is done' {
        $progress = @($script:Convert.Lines | Where-Object { $_.Text -match '^ {6}\S+ Users: \d+ of \d+ \(\d+%\)' })
        $progress.Count | Should -BeGreaterThan 1
        $progress[0].Text | Should -Match 'Users: 0 of 5 \(0%\)$'
        @($progress | Where-Object { $_.Text -match 'about .+ left at this pace$' }).Count | Should -BeGreaterThan 0
        @($script:Convert.Lines | Where-Object { $_.Text -match '^ {6}\S+ Shared mailboxes: 0 of 2 \(0%\)$' }).Count | Should -Be 1
    }
    It 'prints the Recover command and the global report path as next steps, continuation lines aligned under the first' {
        $card = Get-ConsoleCard $script:Convert Summary
        $next = @($card.Rows | Where-Object { $_.Text -match '^ {20}\S' })
        $next.Count | Should -BeGreaterThan 1
        ($next | ForEach-Object { $_.Text }) -join "`n" | Should -Match ('\.\\Invoke-PraCloudMailbox\.ps1 -Action Recover -Batch {0} -ConfigPath' -f $script:Batch)
        ($next | ForEach-Object { $_.Text }) -join "`n" | Should -Match 'Global report of batch'
    }
    It 'follows the batch with a Check: 4 steps for 4 announced, the expected state and the single look in the next steps' {
        $follow = Invoke-Console @{ Action = 'Check'; Batch = $script:Batch; Expect = 'provisioned'; Once = $true }
        $follow.Code | Should -Be 0
        $steps = @(Get-ConsoleStep $follow)
        @($steps | ForEach-Object Title) | Should -Be @('Batch journal', 'Collect snapshot', 'Microsoft 365 connection', 'Follow-up of the batch')
        @($steps | ForEach-Object Total | Select-Object -Unique) | Should -Be @(4)
        $banner = @(Get-ConsoleLabel (Get-ConsoleCard $follow Banner).Rows)
        ($banner | Where-Object Label -eq 'Scope').Text | Should -Be "batch $script:Batch"
        (Get-ConsoleLabel (Get-ConsoleCard $follow Summary).Rows | Where-Object Label -eq 'Result').Text | Should -Be "7 as expected $script:Dot 0 not as expected $script:Dot 0 skipped"
    }
    It 'writes the Preview of Recover and the Recover of the batch in 7 steps, with the global report as the next step' {
        $preview = Invoke-Console @{ Action = 'Recover'; Batch = $script:Batch }
        $preview.Code | Should -Be 0
        (Get-ConsoleCard $preview Summary).Top.Text | Should -Match 'Recover preview finished - nothing was changed'
        $preview.Text | Should -Match ('Convert batch {0} of \d{{4}}-\d{{2}}-\d{{2}} \d{{2}}:\d{{2}}:\d{{2}} UTC: 7 object\(s\) to roll back' -f $script:Batch)
        $preview.Text | Should -Match ('-Action Recover -Mode Apply -Batch {0} -ConfigPath' -f $script:Batch)
        $global:ConsoleWaveSize = 2
        $recover = Invoke-Console @{ Action = 'Recover'; Mode = 'Apply'; Batch = $script:Batch; Force = $true }
        $global:ConsoleWaveSize = 0
        $recover.Code | Should -Be 0
        @(Get-ConsoleStep $recover | ForEach-Object Number) | Should -Be @(1..7)
        (Get-ConsoleCard $recover Summary).Top.Text | Should -Match 'Recover finished'
        $recover.Text | Should -Match 'Global report of batch'
    }
}

Describe 'Console: a run that fails' {
    BeforeEach { Reset-ConsoleWork -ExtraShared 7; $script:Failed = $null }
    It 'lists the first 5 errors, then how many more, with the same count as the Result line, and ends with the next step, FAIL and exit code 1' {
        $convert = Invoke-Console @{ Action = 'Convert'; Mode = 'Apply'; Force = $true }
        $convert.Code | Should -Be 0
        $batch = [regex]::Match($convert.Text, '(?m)^ {5}\S+ Batch +(?<id>[0-9a-f]{8})$').Groups['id'].Value
        $global:PraFake.Faults.DeleteFails = $true
        $run = Invoke-Console @{ Action = 'Recover'; Mode = 'Apply'; Batch = $batch; Force = $true }
        $run.Code | Should -Be 1
        $card = Get-ConsoleCard $run Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Recover stopped - see the errors below ─'
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Fail)
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "3 done $script:Dot 0 pending $script:Dot 9 error(s) $script:Dot 0 skipped"
        $at = @(0..($card.Rows.Count - 1) | Where-Object { $card.Rows[$_].Text -match '^ {5}\S+ Error ' }) | Select-Object -First 1
        $errorLines = @($card.Rows[$at..($at + 6)] | ForEach-Object { $_.Text })
        $errorLines[0] | Should -Match '^ {5}\S+ Error +9 error\(s\) in total, see the log and the report$'
        @($errorLines[1..5] | Where-Object { $_ -match '^ {20}\S+@contoso\.com: rollback stopped: ' }).Count | Should -Be 5
        $errorLines[6] | Should -Be '                    ... and 4 more, see the report'
        ($labels | Where-Object Label -eq 'Next').Text | Should -Match '^Some objects are not back yet\.'
        $run.Text | Should -Match ('\.\\Invoke-PraCloudMailbox\.ps1 -Action Recover -Mode Apply -Batch {0} -ConfigPath' -f $batch)
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: FAIL'
        @($run.Lines | Where-Object { $_.Text -match '^ {6}\S+ .*rollback stopped' } | ForEach-Object Color | Select-Object -Unique) | Should -Be @($script:Theme.Fail)
    }
    It 'prints the next step also when the run stopped before anything was done' {
        $run = Invoke-Console @{ Action = 'Convert'; Group = 'nogroup' }
        $run.Code | Should -Be 1
        $card = Get-ConsoleCard $run Summary
        $card.Top.Text | Should -Match 'Convert stopped - see the errors below'
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "0 ready $script:Dot 0 planned $script:Dot 1 error(s) $script:Dot 0 skipped"
        ($labels | Where-Object Label -eq 'Error').Text | Should -Match "^Stopped at step 'Collect snapshot' \(operation Resolve-Group\): Group nogroup is not in snapshot"
        $run.Text | Should -Match '\.\\Invoke-PraCloudMailbox\.ps1 -Action Convert -Group ''nogroup'' -ConfigPath'
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: FAIL'
    }
    It 'finishes a Check that found an object not ready as finished, not stopped, and counts it as not as expected' {
        $global:PraFake.Users.Remove(@($global:PraFake.Users.Keys | Where-Object { $global:PraFake.Users[$_].Upn -eq 'adele.vance@contoso.com' })[0])
        $run = Invoke-Console @{ Action = 'Check'; Scope = 'UsersOnly' }
        $run.Code | Should -Be 1
        $card = Get-ConsoleCard $run Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Check finished - 1 object\(s\) not in the expected state ─'
        (Get-ConsoleLabel $card.Rows | Where-Object Label -eq 'Result').Text | Should -Be "2 as expected $script:Dot 1 not as expected $script:Dot 0 skipped"
        @($run.Lines | Where-Object { $_.Text -match '^ {6}\S+ adele\.vance@contoso\.com \(user\): No Entra ID object' } | ForEach-Object Color) | Should -Be @($script:Theme.Fail)
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: FAIL'
    }
    It 'stops a Collect without Exchange.Server in plain words, at the first step' {
        $run = Invoke-Console @{ Action = 'Collect' }
        $run.Code | Should -Be 1
        @(Get-ConsoleStep $run | ForEach-Object Title) | Should -Be @('Exchange on-premises connection')
        $run.Text | Should -Match "Stopped at step 'Exchange on-premises connection' \(operation Connect-Exchange\): Exchange\.Server is empty\."
        (Get-ConsoleCard $run Summary).Top.Text | Should -Match 'Collect stopped - see the errors below'
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: FAIL'
    }
}

Describe 'Console: a run that needs a next step' {
    It 'ends done with exit code 2, a warning card, the objects pending, the resume command and the verdict PENDING' {
        Reset-ConsoleWork -ExtraUsers 3
        $global:PraFake.Faults.NeverMailbox = @($script:People.Keys | Where-Object { $_ -like 'user*' } | ForEach-Object { $script:People[$_].Id })
        $run = Invoke-Console @{ Action = 'Convert'; Mode = 'Apply'; Force = $true }
        $run.Code | Should -Be 2
        $card = Get-ConsoleCard $run Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Convert done - next step required ─'
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Warn)
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "5 done $script:Dot 3 pending $script:Dot 0 error(s) $script:Dot 0 skipped"
        ($labels | Where-Object Label -eq 'Warning').Text | Should -Be '3 warning(s) in total, see the log and the report'
        $run.Text | Should -Match 'Some objects or lists are not finished\. Fix what is named above, then resume them'
        $run.Text | Should -Match '\.\\Invoke-PraCloudMailbox\.ps1 -Action Convert -Mode Apply -Batch [0-9a-f]{8} -ConfigPath'
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: PENDING'
        (Get-ConsoleVerdict $run).Color | Should -Be $script:Theme.Dim
    }
}

Describe 'Console: the operator answers no to the confirmation' {
    It 'ends cancelled with exit code 1, no error counted and nothing about a stop on the console, and says to answer yes next time' {
        Reset-ConsoleWork
        $global:ConsoleAnswersNo = $true
        try { $run = Invoke-Console @{ Action = 'Convert'; Mode = 'Apply' } } finally { $global:ConsoleAnswersNo = $false }
        $run.Code | Should -Be 1
        $card = Get-ConsoleCard $run Summary
        $card.Top.Text | Should -Match '^  ╭─ \S+  Convert cancelled - nothing was written ─'
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Warn)
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Be "0 done $script:Dot 0 pending $script:Dot 0 error(s) $script:Dot 0 skipped"
        $labels.Label | Should -Not -Contain 'Error'
        ($labels | Where-Object Label -eq 'Next').Text | Should -Be 'Nothing was written. To run it, give the same command again and answer yes to the confirmation:'
        $run.Text | Should -Not -Match 'Stopped at step'
        $index = @(0..($run.Lines.Count - 1) | Where-Object { $run.Lines[$_].Text -match '^ {6}\S+ Planned: ' }) | Select-Object -Last 1
        $run.Lines[$index + 1].Text | Should -Be ''
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: CANCELLED'
    }
}

Describe 'Console: the snapshot file and the run-only settings' {
    BeforeEach { Reset-ConsoleWork; $script:Hash = (Get-ChildItem (Join-Path $script:Work 'data') -Filter '*.sha256')[0].FullName }
    It 'stops at the first step when the snapshot changed after its Collect, and says what to do' {
        [IO.File]::WriteAllText($script:Hash, ('{0}  PraCloudMailbox.db' -f ('0' * 64)))
        $run = Invoke-Console @{ Action = 'Convert' }
        $run.Code | Should -Be 1
        $labels = @(Get-ConsoleLabel (Get-ConsoleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Integrity').Text | Should -Be 'failed: the file does not match its .sha256'
        $run.Text | Should -Match 'was changed after its last Collect: its SHA-256 is'
        $run.Text | Should -Match 'run the same command again with -AcceptUnverifiedSnapshot'
    }
    It 'accepts it on request: a banner row, a warning, the integrity row of the card and the switch in the next step' {
        [IO.File]::WriteAllText($script:Hash, ('{0}  PraCloudMailbox.db' -f ('0' * 64)))
        $run = Invoke-Console @{ Action = 'Convert'; AcceptUnverifiedSnapshot = $true }
        $run.Code | Should -Be 0
        $banner = @(Get-ConsoleLabel (Get-ConsoleCard $run Banner).Rows)
        $banner.Label | Should -Contain 'Integrity'
        ($banner | ForEach-Object Label) | Should -Be @('Action', 'Mode', 'Scope', 'Tenant', 'Database', 'Integrity', 'Config', 'Log')
        $card = Get-ConsoleCard $run Summary
        $card.Top.Text | Should -Match 'Convert preview finished - nothing was changed .+ 1 warning\(s\) ─'
        @($card.Top.Color, $card.Bottom.Color | Select-Object -Unique) | Should -Be @($script:Theme.Warn)
        (Get-ConsoleLabel $card.Rows | Where-Object Label -eq 'Integrity').Text | Should -Be 'skipped on request (-AcceptUnverifiedSnapshot); the file does not match its .sha256'
        $run.Text | Should -Match '-Action Convert -Mode Apply -AcceptUnverifiedSnapshot -ConfigPath'
    }
    It 'warns when there is no .sha256 beside the snapshot and goes on' {
        Remove-Item -LiteralPath $script:Hash
        $run = Invoke-Console @{ Action = 'Check' }
        $run.Code | Should -Be 0
        @($run.Lines | Where-Object { $_.Text -match '^ {6}\S+ The snapshot file .+ has no .+\.sha256 beside it' } | ForEach-Object Color) | Should -Be @($script:Theme.Warn)
        (Get-ConsoleLabel (Get-ConsoleCard $run Summary).Rows | Where-Object Label -eq 'Integrity').Text | Should -Be 'not verifiable: no .sha256 beside the file'
    }
    It 'shows what a run-only configuration changes: a banner row and the settings of the card, at most 4 lines' {
        $changed = Join-Path $script:Work 'run.config.psd1'
        (Get-Content -LiteralPath $script:ConfigFile -Raw).Replace("DefaultUsageLocation = 'FR'", "DefaultUsageLocation = 'DE'").Replace('MailboxTimeoutMinutes = 30', 'MailboxTimeoutMinutes = 45') | Set-Content -LiteralPath $changed -Encoding UTF8
        $run = Invoke-Console @{ Action = 'Convert'; ConfigPath = $changed; BaseConfigPath = $script:ConfigFile }
        $run.Code | Should -Be 0
        $banner = @(Get-ConsoleLabel (Get-ConsoleCard $run Banner).Rows)
        ($banner | ForEach-Object Label) | Should -Be @('Action', 'Mode', 'Scope', 'Tenant', 'Database', 'Config', 'Settings', 'Log')
        ($banner | Where-Object Label -eq 'Settings').Text | Should -Match '^2 setting\(s\) differ from console\.config\.psd1: Cloud\.DefaultUsageLocation: FR -> DE; Polling\.MailboxTimeoutMinutes: 30 -> 45$'
        $card = Get-ConsoleCard $run Summary
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Settings').Text | Should -Be '2 setting(s) of this run differ from the configuration file:'
        $run.Text | Should -Match "(?m)^ {20}Cloud\.DefaultUsageLocation: FR -> DE$"
        $run.Text | Should -Match ('-Action Convert -Mode Apply -ConfigPath ''{0}''' -f [regex]::Escape($script:ConfigFile))
    }
}

Describe 'Console: options that cannot go together' {
    BeforeAll { Reset-ConsoleWork }
    It 'proposes one command for each way to select the objects instead of the command that was refused' {
        $run = Invoke-Console @{ Action = 'Convert'; Identity = 'adele.vance@contoso.com'; Group = 'PRA-Wave1' }
        $run.Code | Should -Be 1
        $labels = @(Get-ConsoleLabel (Get-ConsoleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Next').Text | Should -Be 'These options select the objects in different ways and cannot be combined (-Identity, -Group): keep one of them. Run the command with the one you want:'
        $run.Text | Should -Match "(?m)^ {20}Only -Identity \(remove -Group\):$"
        $run.Text | Should -Match ('(?m)^ {{20}}\.\\Invoke-PraCloudMailbox\.ps1 -Action Convert -Identity ''adele\.vance@contoso\.com'' -ConfigPath ''{0}''$' -f [regex]::Escape($script:ConfigFile))
        $run.Text | Should -Match ('(?m)^ {{20}}\.\\Invoke-PraCloudMailbox\.ps1 -Action Convert -Group ''PRA-Wave1'' -ConfigPath ''{0}''$' -f [regex]::Escape($script:ConfigFile))
        $run.Text | Should -Not -Match "-Identity '[^']*' -Group"
        $run.Text | Should -Match "operation Check-Options"
    }
}

Describe 'Console: the window channel and -PassThru add nothing to the console' {
    It 'writes the same lines with the event file of the window and with -PassThru, and returns the result after the verdict line' {
        Reset-ConsoleWork
        $plain = Invoke-Console @{ Action = 'Check' }
        $events = Join-Path $script:Work 'run.events.jsonl'
        $env:PRA_EVENT_FILE = $events
        & (Get-Module PRA.Common) { $script:EventFile = $env:PRA_EVENT_FILE }
        try { $windowed = Invoke-Console @{ Action = 'Check'; PassThru = $true } }
        finally { Remove-Item Env:\PRA_EVENT_FILE -ErrorAction SilentlyContinue; & (Get-Module PRA.Common) { $script:EventFile = $null } }
        $shape = { param($Run) @($Run.Lines | ForEach-Object { '{0}|{1}' -f $_.Color, ([regex]::Replace(([regex]::Replace($_.Text, '[0-9a-f]{8,}', '#')), '\d', '9')) }) }
        (& $shape $windowed) | Should -Be (& $shape $plain)
        $windowed.Records[-1] | Should -BeOfType ([pscustomobject])
        $windowed.Records[-1].ExitCode | Should -Be 0
        @($windowed.Records | Where-Object { $_ -isnot [Management.Automation.InformationRecord] }).Count | Should -Be 1
        (Get-Content -LiteralPath $events -Raw) | Should -Match '"kind":"start"'
    }
}

Describe 'One error for one cause when the log folder cannot be created' {
    BeforeEach {
        Reset-ConsoleWork
        $script:Blocker = Join-Path $script:Work 'blocker'
        Set-Content -LiteralPath $script:Blocker -Value 'a file where a folder is needed'
    }
    It 'counts one error and names the folder when the log folder is under a file' {
        $folder = Join-Path $script:Blocker 'logs'
        $text = [IO.File]::ReadAllText($script:ConfigFile)
        [IO.File]::WriteAllText($script:ConfigFile, [regex]::Replace($text, '(?m)^\s*Logging = .*$', "    Logging = @{ Folder = '$folder' }"))
        $run = Invoke-Console @{ Action = 'Check' }
        $run.Code | Should -Be 1
        $card = Get-ConsoleCard $run Summary
        $labels = @(Get-ConsoleLabel $card.Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Match '(^|\s)1 error\(s\)(\s|$)' -Because $run.Text
        @($labels | Where-Object Label -eq 'Error').Count | Should -Be 1 -Because $run.Text
        ($labels | Where-Object Label -eq 'Error').Text | Should -Match ('^Audit could not start: The folder {0} cannot be created: ' -f [regex]::Escape($folder)) -Because $run.Text
        $run.Text | Should -Not -Match 'Audit file'
        $run.Text | Should -Not -Match 'Audit unavailable'
        (Get-ConsoleVerdict $run).Text | Should -Be '  RESULT: FAIL'
        Test-Path -LiteralPath $folder | Should -BeFalse
    }
    It 'creates the log of a run whose configuration cannot be read and counts one error: <Name>' -ForEach @(
        @{ Name = 'unknown setting'; Mutate = { param($t) $t.Replace("Environment = 'PROD'", "Environment = 'PROD'`n    Typo = 1") } }
        @{ Name = 'not a data file'; Mutate = { 'this is not a configuration' } }
    ) {
        [IO.File]::WriteAllText($script:ConfigFile, (& $Mutate ([IO.File]::ReadAllText($script:ConfigFile))))
        $logs = Join-Path $script:Root 'logs'
        $before = @(Get-ChildItem -LiteralPath $logs -Filter '*.log' -File -ErrorAction SilentlyContinue | ForEach-Object FullName)
        $run = Invoke-Console @{ Action = 'Check' }
        $run.Code | Should -Be 1
        $labels = @(Get-ConsoleLabel (Get-ConsoleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Match '(^|\s)1 error\(s\)(\s|$)' -Because $run.Text
        @($labels | Where-Object Label -eq 'Error').Count | Should -Be 1 -Because $run.Text
        $run.Text | Should -Not -Match 'Audit file'
        $run.Text | Should -Not -Match 'Audit unavailable'
        $run.Text | Should -Not -Match 'Audit could not start'
        $created = @(Get-ChildItem -LiteralPath $logs -Filter '*.log' -File -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notin $before })
        $created.Count | Should -Be 1
        [IO.File]::ReadAllText($created[0].FullName) | Should -Match 'RESULT: FAIL'
    }
    It 'leaves a normal run as it was: no error, one log, exit code 0' {
        $run = Invoke-Console @{ Action = 'Check' }
        $run.Code | Should -Be 0
        $labels = @(Get-ConsoleLabel (Get-ConsoleCard $run Summary).Rows)
        ($labels | Where-Object Label -eq 'Result').Text | Should -Not -Match 'error\(s\)'
        @($labels | Where-Object Label -eq 'Error').Count | Should -Be 0
        $run.Text | Should -Not -Match 'Audit '
        @(Get-ChildItem -LiteralPath (Join-Path $script:Work 'logs') -Filter '*.log' -File).Count | Should -Be 1
    }
    It 'adds exactly one issue for an empty log path when no audit failure was reported, and none when one was' -ForEach @(
        @{ Existing = $true }
        @{ Existing = $false }
    ) {
        $context = @{ Root = $script:Root; Version = 't'; RunId = 'r'; StartTime = Get-Date; Action = 'Check'; Mode = 'Preview'; Phase = ''; CurrentPhase = 'Start'; CurrentOperation = ''; CurrentIdentity = ''; Warnings = 0
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
            LogFolder = $script:Work; ReportFolder = $script:Work; LogFile = ''; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''
            Config = @{}; BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null; FallbackNext = @('fallback') }
        if ($Existing) { Add-PcmIssue $context 'Audit could not start: the log folder cannot be created'; $context.ExitCode = 1 }
        $result = Complete-PcmRun -Context $context 6>$null
        @($result.Issues).Count | Should -Be 1 -Because (@($result.Issues | ForEach-Object { $_.Message }) -join ' / ')
        @($result.Issues | Where-Object { $_.Message -match '^Audit file' }).Count | Should -Be $(if ($Existing) { 0 } else { 1 })
    }
}
Describe 'Console: the readiness line counts the units left for the shared mailboxes' {
    BeforeAll {
        Reset-ConsoleWork -ExtraShared 20
        $script:Readiness = Invoke-Console @{ Action = 'Convert' }
    }
    It 'says how many shared mailboxes hold a temporary unit at once, after the users of the same run took theirs' {
        $match = [regex]::Match($script:Readiness.Text, '(\d+) new unit\(s\) for users, up to (\d+) at a time for the temporary licence of the shared mailboxes')
        $match.Success | Should -BeTrue -Because $script:Readiness.Text
        $users = [int]$match.Groups[1].Value
        $users | Should -BeGreaterThan 0
        [int]$match.Groups[2].Value | Should -Be (10 - $users)
    }
    It 'shows the licence as the tenant spells it, whatever the letter case typed in the configuration' {
        [IO.File]::WriteAllText($script:ConfigFile, ([IO.File]::ReadAllText($script:ConfigFile).Replace("SkuPartNumber = 'E5'", "SkuPartNumber = 'e5'")))
        $run = Invoke-Console @{ Action = 'Convert' }
        $run.Text | Should -Match 'temporary E5, shared'
        $run.Text | Should -Not -MatchExactly 'temporary e5'
    }
}

Describe 'Console: no stack line in an error' {
    It 'shows the text of an error without its stack in the card, the console and the result, and keeps the whole detail in the log' {
        $log = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.log')
        $context = @{ Root = $script:Root; Version = 't'; RunId = 'r'; StartTime = Get-Date; Action = 'Recover'; Mode = 'Apply'; Phase = ''; CurrentPhase = 'Delta'; CurrentOperation = 'x'; CurrentIdentity = ''; Warnings = 0
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); BackupFiles = (New-Object 'Collections.Generic.List[string]'); StateFiles = (New-Object 'Collections.Generic.List[string]')
            LogFolder = $TestDrive; ReportFolder = $TestDrive; LogFile = $log; TranscriptPath = ''; TranscriptStarted = $false; NoReport = $true; Server = ''; CloudConnectFailed = $false; ExitCode = 0; ResultStatus = ''
            Config = @{}; BatchId = ''; SnapshotLabel = ''; NextSteps = @(); Journal = $null; FallbackNext = @('fallback') }
        $raw = "The call failed.`n   at Microsoft.Foo.Bar(String name) in C:\src\Bar.cs:line 42`n   at <ScriptBlock>, <No file>: line 3`n   at Invoke-PcmGraph, C:\PRA\PRA.Cloud.psm1: line 626`n--- End of stack trace ---"
        Write-PcmLog -Context $context -Message $raw -Level Error 6>$null
        $context.Issues.Add([pscustomobject]@{ Message = "Second failure.`n   at <ScriptBlock>, <No file>: line 9"; Source = 'Execution' })
        $text = (Complete-PcmRun -Context $context 6>&1 | Out-String)
        $text | Should -Match 'The call failed\.'
        $text | Should -Match 'Second failure\.'
        $text | Should -Not -Match '(?m)(^|\|)\s*at (<ScriptBlock>|Microsoft\.Foo|Invoke-PcmGraph)'
        $text | Should -Not -Match 'End of stack trace'
        @($context.Issues | Select-Object -First 2 | ForEach-Object { [string]$_.Message }) | Should -Be @('The call failed.', 'Second failure.')
        $logText = [IO.File]::ReadAllText($log)
        $logText | Should -Match 'at Microsoft\.Foo\.Bar'
    }
    It 'strips the frames of .NET and of PowerShell, and nothing else' {
        ConvertTo-PcmPlainMessage "Exception calling it`n   at System.Net.Http.Foo.Bar() in x.cs:line 5`n   at <ScriptBlock>, <No file>: line 1`nwhy it failed" | Should -Be "Exception calling it$([Environment]::NewLine)why it failed"
        ConvertTo-PcmPlainMessage 'We arrived at the end' | Should -Be 'We arrived at the end'
        ConvertTo-PcmPlainMessage '' | Should -Be ''
    }
}
