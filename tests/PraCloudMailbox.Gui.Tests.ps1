<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the window (Pester 5+), in both editions. No Exchange and no cloud: the state is a
    test configuration, snapshot, journal and Check report; the questions and confirmations are answered by hooks.
.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.0
#>
#Requires -Version 5.1

BeforeAll {
    $script:Root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    Import-Module (Join-Path $script:Root 'module\PRA2.Common.psm1') -Force
    Import-Module (Join-Path $script:Root 'module\PRA2.Store.psm1') -Force
    $script:Gui = Import-Module (Join-Path $script:Root 'module\PRA2.Gui.psm1') -Force -PassThru
    $ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Root 'Invoke-PraCloudMailbox.ps1'), [ref]$null, [ref]$null)
    $function = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Read-PraIdentityFile' }, $true)
    . ([scriptblock]::Create($function.Extent.Text))

    $script:Alice = 'a0000000-0000-0000-0000-000000000001'
    $script:Bob = 'a0000000-0000-0000-0000-000000000002'
    $script:Compta = 'a0000000-0000-0000-0000-000000000003'

    function script:Write-TestCsv {
        <# A report CSV as Complete-PraRun writes it (delimiter ;). #>
        param([string]$Path, [object[]]$Rows)
        $Rows | ForEach-Object { [pscustomobject]$_ } | Export-Csv -LiteralPath $Path -NoTypeInformation -Delimiter ';' -Encoding UTF8
    }
    function script:New-TestState {
        <# Configuration, snapshot (alice, bob, compta), journal (a Partial Convert batch: alice done, bob not started) and Check. #>
        param([string]$Folder)
        $data = Join-Path $Folder 'data'; $reports = Join-Path $Folder 'reports'
        $null = New-Item -ItemType Directory -Path $data, $reports, (Join-Path $Folder 'logs') -Force
        $config = Join-Path $Folder 'test.config.psd1'
        $text = @"
@{
    Environment = 'TEST'
    Scope = @{ Mode = 'Auto' }
    Store = @{ Path = '$data\test.db'; JournalPath = '$data\journal.db'; BackupFolder = '$data\backup' }
    Cloud = @{ TenantId = '11111111-2222-3333-4444-555555555555'; Organization = 'contoso.onmicrosoft.com'; AppId = '66666666-7777-8888-9999-000000000000'; CertificateThumbprint = '$('A' * 40)' }
    Licensing = @{ Users = @{ Mode = 'Group'; GroupId = '12121212-1212-1212-1212-121212121212' }; Shared = @{ SkuPartNumber = 'SPE_E5' } }
    Retention = @{ HoldPolicy = 'PRA-Hold' }
    EntraConnect = @{ Mode = 'Manual' }
    Logging = @{ Folder = '$Folder\logs' }
    Report = @{ Enabled = `$true; Folder = '$reports' }
}
"@
        [IO.File]::WriteAllText($config, $text, (New-Object Text.UTF8Encoding($true)))
        $cn = Open-Pra2Store -Path "$data\test.db" -Root $script:Root
        try {
            $sid = New-Pra2Snapshot $cn @{ RunId = 'r1'; Environment = 'TEST'; ToolVersion = 'test'; ExchangeServer = 'EX1'; ScopeJson = '{}' }
            $rows = @(
                [ordered]@{ object_guid = $script:Alice; kind = 'User'; user_principal_name = 'alice@contoso.com'; primary_smtp_address = 'alice@contoso.com'; sam_account_name = 'alice' }
                [ordered]@{ object_guid = $script:Bob; kind = 'User'; user_principal_name = 'bob@contoso.com'; primary_smtp_address = 'bob@contoso.com'; sam_account_name = 'bob' }
                [ordered]@{ object_guid = $script:Compta; kind = 'Shared'; user_principal_name = 'compta@contoso.com'; primary_smtp_address = 'compta@contoso.com'; sam_account_name = 'compta' })
            $null = Add-Pra2Row $cn mailbox $sid $rows
            Set-Pra2SnapshotStatus $cn $sid Complete @{ mailbox = 3; permission = 0 }
        } finally { Close-Pra2Store $cn }
        $j = Open-Pra2Store -Path "$data\journal.db" -Root $script:Root -Kind Journal
        try {
            $batch = New-Pra2Batch $j Convert @{ SnapshotId = $sid; Environment = 'TEST' }
            Set-Pra2BatchItem $j $batch $script:Alice @{ kind = 'User'; identity = 'alice@contoso.com'; step = 'Mailbox'; status = 'Done' }
            Set-Pra2BatchItem $j $batch $script:Bob @{ kind = 'User'; identity = 'bob@contoso.com'; status = 'Planned' }
            Set-Pra2Batch $j $batch Partial
        } finally { Close-Pra2Store $j }
        $check = { param($guid, $upn, $kind, $status, $detail, $warnings) [ordered]@{ Identity = $upn; Kind = $kind; PrimarySmtpAddress = $upn; ObjectGuid = $guid; Action = 'Check'
                ExchangeOnline = 'MailUser'; Licence = ''; Holds = ''; FinalStatus = $status; Detail = $detail; Warnings = $warnings } }
        Write-TestCsv -Path (Join-Path $reports 'PRA2_check_test.csv') -Rows @(
            (& $check $script:Alice 'alice@contoso.com' 'User' 'Success' 'ready' '')
            (& $check $script:Bob 'bob@contoso.com' 'User' 'Error' 'NO_USAGE_LOCATION: no usage location' '')
            (& $check $script:Compta 'compta@contoso.com' 'Shared' 'Success' 'ready' 'HOLD_PRESENT: hold on the object'))
        return [pscustomobject]@{ Config = $config; Batch = $batch; Snapshot = $sid; Reports = $reports; Data = $data }
    }
    function script:Invoke-TestPump {
        <# Lets the dispatcher run the timers of the window for a moment. #>
        $frame = New-Object Windows.Threading.DispatcherFrame
        [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background,
            [Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $frame)
        [Windows.Threading.Dispatcher]::PushFrame($frame)
    }
}

Describe 'Window: data (pure functions)' {
    It 'builds the command line of a run: same script and parameters, -Force only for Apply, -NonInteractive always' {
        $preview = Get-PraGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\config\c.psd1' -Engine 'pwsh.exe' -Request @{ Action = 'Convert'; Mode = 'Preview'; IdentityPath = 'C:\PRA\data\gui\wave.csv' }
        $preview.Arguments | Should -Match '^-NoProfile -NonInteractive -ExecutionPolicy Bypass -File "C:\\PRA\\Invoke-PraCloudMailbox.ps1" -Action Convert -Mode Preview -IdentityPath "C:\\PRA\\data\\gui\\wave.csv" -ConfigPath "C:\\PRA\\config\\c.psd1"$'
        $preview.Display | Should -Be 'pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Preview -IdentityPath "C:\PRA\data\gui\wave.csv"'
        $apply = Get-PraGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\c.psd1' -Engine 'pwsh.exe' -Request @{ Action = 'Recover'; Mode = 'Apply'; Batch = 'abcd1234' }
        $apply.Arguments | Should -Match '-Action Recover -Mode Apply -Batch abcd1234 -ConfigPath "C:\\PRA\\c.psd1" -Force$'
        $check = Get-PraGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\c.psd1' -Engine 'pwsh.exe' -Request @{ Action = 'Check'; Scope = 'UsersOnly' }
        $check.Arguments | Should -Not -Match '-Mode'
        $check.Arguments | Should -Match '-Action Check -Scope UsersOnly'
        (Get-PraGuiCommand -Root 'C:\PRA' -ConfigPath 'C:\PRA\c.psd1' -Engine 'powershell.exe' -Request @{ Action = 'Collect'; Mode = 'Apply' }).Display | Should -BeLike 'powershell.exe -File *-Action Collect -Mode Apply -Force'
    }
    It 'reads the events incrementally, complete lines only, and keeps a line that is not JSON as text' {
        $path = Join-Path $TestDrive 'events.jsonl'
        [IO.File]::WriteAllText($path, '{"kind":"step","index":1}' + "`n" + 'not json' + "`n" + '{"kind":"item","te')
        $first = Read-PraGuiEvents -Path $path
        @($first.Events).Count | Should -Be 2
        $first.Events[0].kind | Should -Be 'step'
        $first.Events[1].text | Should -Be 'not json'
        [IO.File]::AppendAllText($path, 'xt":"done"}' + "`n")
        $second = Read-PraGuiEvents -Path $path -Position $first.Position
        @($second.Events).Count | Should -Be 1
        $second.Events[0].text | Should -Be 'done'
        @((Read-PraGuiEvents -Path $path -Position $second.Position).Events).Count | Should -Be 0
        @((Read-PraGuiEvents -Path (Join-Path $TestDrive 'missing.jsonl')).Events).Count | Should -Be 0
    }
    It 'filters as typed: quotes and wildcards are literal, the clauses are joined with AND' {
        Get-PraGuiRowFilter -Text "  O'Brien*[x]%  " -Clauses @("Kind = 'User'", '', "InCloud = '0'") | Should -Be "Search LIKE '%o''brien[*][[]x[]][%]%' AND Kind = 'User' AND InCloud = '0'"
        Get-PraGuiRowFilter -Text '' -Clauses @() | Should -Be ''
    }
    It 'keys a selection by its objects, whatever their order' {
        $one = Get-PraGuiSelectionKey -Prefix 'Convert|4' -Guids @('b', 'a')
        $one | Should -Be (Get-PraGuiSelectionKey -Prefix 'Convert|4' -Guids @('a', 'b'))
        $one | Should -Not -Be (Get-PraGuiSelectionKey -Prefix 'Convert|4' -Guids @('a', 'c'))
        $one | Should -Not -Be (Get-PraGuiSelectionKey -Prefix 'Convert|5' -Guids @('a', 'b'))
        $one | Should -Match '^Convert\|4\|2\|[0-9A-F]{16}$'
    }
    It 'turns a Check report into Ready, Warning and Not ready rows, and names a tenant finding' {
        $rows = @(
            [pscustomobject]@{ Identity = 'a@c.com'; Kind = 'User'; ObjectGuid = 'g1'; PrimarySmtpAddress = 'a@c.com'; FinalStatus = 'Success'; Warnings = ''; Detail = 'ok'; Licence = ''; ExchangeOnline = ''; Holds = '' }
            [pscustomobject]@{ Identity = 'b@c.com'; Kind = 'User'; ObjectGuid = 'g2'; PrimarySmtpAddress = 'b@c.com'; FinalStatus = 'Success'; Warnings = 'HOLD'; Detail = 'ok'; Licence = ''; ExchangeOnline = ''; Holds = '' }
            [pscustomobject]@{ Identity = 'c@c.com'; Kind = 'Shared'; ObjectGuid = 'g3'; PrimarySmtpAddress = 'c@c.com'; FinalStatus = 'Error'; Warnings = 'x'; Detail = ('y' * 300); Licence = ''; ExchangeOnline = ''; Holds = '' }
            [pscustomobject]@{ Identity = ''; Kind = ''; ObjectGuid = ''; PrimarySmtpAddress = ''; FinalStatus = 'Error'; Warnings = ''; Detail = '[Run/Execution/Tenant-Facts] 1 free unit for 5 users'; Licence = ''; ExchangeOnline = ''; Holds = '' })
        $table = ConvertTo-PraGuiCheckTable -Rows $rows
        @($table.Rows | ForEach-Object { $_['StatusText'] }) | Should -Be @('Ready', 'Warning', 'Not ready', 'Not ready')
        $table.Rows[1]['Issue'] | Should -Be 'HOLD'
        $table.Rows[2]['Issue'].Length | Should -Be 260
        $table.Rows[3]['Identity'] | Should -Be '(tenant)'
        $table.Rows[3]['Issue'] | Should -Be '1 free unit for 5 users'
        (ConvertTo-PraGuiCheckTable -Rows @()).Rows.Count | Should -Be 0
    }
    It 'shows the last Convert batch of each object, and an object rolled back is no longer in the cloud' {
        $objects = @(
            [pscustomobject]@{ object_guid = 'g1'; kind = 'User'; user_principal_name = 'a@c.com'; primary_smtp_address = 'a@c.com'; sam_account_name = 'a' }
            [pscustomobject]@{ object_guid = 'g2'; kind = 'User'; user_principal_name = 'b@c.com'; primary_smtp_address = 'b@c.com'; sam_account_name = 'b' }
            [pscustomobject]@{ object_guid = 'g3'; kind = 'Shared'; user_principal_name = ''; primary_smtp_address = 'c@c.com'; sam_account_name = 'c' })
        $back = New-Object 'Collections.Generic.HashSet[string]'; [void]$back.Add('g2')
        $newer = [pscustomobject]@{ Id = 'new00001'; Back = (New-Object 'Collections.Generic.HashSet[string]'); Items = @([pscustomobject]@{ object_guid = 'g1'; status = 'Planned'; step = '' }) }
        $older = [pscustomobject]@{ Id = 'old00001'; Back = $back; Items = @([pscustomobject]@{ object_guid = 'g1'; status = 'Done'; step = 'Mailbox' }, [pscustomobject]@{ object_guid = 'g2'; status = 'Done'; step = 'Mailbox' }) }
        $table = ConvertTo-PraGuiConvertTable -Objects $objects -Batches @($newer, $older) -Ticked @('G3')
        @($table.Rows | ForEach-Object { $_['Batch'] }) | Should -Be @('new00001 · not started', 'old00001 · back', '')
        @($table.Rows | ForEach-Object { $_['InCloud'] }) | Should -Be @('0', '0', '0')
        @($table.Rows | ForEach-Object { $_['StatusText'] }) | Should -Be @('Not checked', 'Not checked', 'Not checked')
        $table.Rows[2]['Identity'] | Should -Be 'c@c.com'
        Get-PraGuiTicked $table | Should -Be @('g3')
        $done = [pscustomobject]@{ Id = 'b1'; Back = (New-Object 'Collections.Generic.HashSet[string]'); Items = @([pscustomobject]@{ object_guid = 'g1'; status = 'Failed'; step = 'Licence' }) }
        (ConvertTo-PraGuiConvertTable -Objects $objects -Batches @($done)).Rows[0]['InCloud'] | Should -Be '1'
    }
    It 'shows what is waiting for Recover in a batch' {
        $back = New-Object 'Collections.Generic.HashSet[string]'; [void]$back.Add('g2')
        $batch = [pscustomobject]@{ Id = 'b1'; Back = $back; Items = @(
                [pscustomobject]@{ object_guid = 'g1'; kind = 'User'; identity = 'a'; status = 'Done'; step = 'Mailbox' }
                [pscustomobject]@{ object_guid = 'g2'; kind = 'User'; identity = 'b'; status = 'Done'; step = 'Mailbox' }
                [pscustomobject]@{ object_guid = 'g3'; kind = 'Shared'; identity = 'c'; status = 'Planned'; step = '' }
                [pscustomobject]@{ object_guid = 'g4'; kind = 'Shared'; identity = 'd'; status = 'Failed'; step = 'Permissions' }) }
        $table = ConvertTo-PraGuiRecoverTable -Batch $batch
        @($table.Rows | ForEach-Object { $_['StatusText'] }) | Should -Be @('Waiting', 'Back on-premises', 'Nothing to roll back', 'Waiting')
        @($table.Rows | ForEach-Object { $_['Converted'] }) | Should -Be @('Done', 'Done', 'Not started', 'Failed at Permissions')
    }
    It 'writes a wave that -IdentityPath reads back (object GUIDs)' {
        $rows = @([pscustomobject]@{ ObjectGuid = 'g1'; Identity = 'a@c.com' }, [pscustomobject]@{ ObjectGuid = 'g2'; Identity = 'Dupont "Jean"' })
        $path = Write-PraGuiWave -Folder (Join-Path $TestDrive 'gui') -Name 'convert' -Rows $rows
        $path | Should -Match 'wave-convert-\d{8}-\d{6}\.csv$'
        Read-PraIdentityFile -Path $path | Should -Be @('g1', 'g2')
    }
    It 'counts the objects of a preview from its report: planned users, shared, and the objects not ready' {
        $path = Join-Path $TestDrive 'preview.csv'
        Write-TestCsv -Path $path -Rows @(
            [ordered]@{ Identity = 'a'; Kind = 'User'; ObjectGuid = 'g1'; FinalStatus = 'Planned' }
            [ordered]@{ Identity = 'c'; Kind = 'Shared'; ObjectGuid = 'g3'; FinalStatus = 'Planned' }
            [ordered]@{ Identity = 'b'; Kind = 'User'; ObjectGuid = 'g2'; FinalStatus = 'Error' }
            [ordered]@{ Identity = ''; Kind = ''; ObjectGuid = ''; FinalStatus = 'Error' })
        $count = Get-PraGuiPlanCount -Path $path
        $count.Planned | Should -Be 2; $count.Users | Should -Be 1; $count.Shared | Should -Be 1; $count.Blocked | Should -Be 1
        (Get-PraGuiPlanCount -Path '').Planned | Should -Be 0
    }
    It 'shows times in local time and the paths of the tool folder as relative paths' {
        $utc = [datetime]::new(2026, 10, 8, 12, 0, 0, [DateTimeKind]::Utc)
        Format-PraGuiTime $utc.ToString('o') | Should -Be $utc.ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        Format-PraGuiTime '' | Should -Be ''
        & $script:Gui { Format-PraGuiPath -Path 'C:\PRA\data\x.db' -Root 'C:\PRA\' } | Should -Be '.\data\x.db'
        & $script:Gui { Format-PraGuiPath -Path 'D:\other\x.db' -Root 'C:\PRA' } | Should -Be 'D:\other\x.db'
    }
}

Describe 'Window: fresh installation' {
    It 'opens with the shipped configuration and nothing else: no snapshot, no journal, no Check' {
        $folder = Join-Path $TestDrive 'fresh'
        $null = New-Item -ItemType Directory -Path $folder -Force
        $config = Join-Path $folder 'fresh.config.psd1'
        $text = "@{ Environment = 'PROD'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Mode = 'Manual' }; Store = @{ Path = '$folder\data\PraCloudMailbox.db'; JournalPath = '$folder\data\journal.db'; BackupFolder = '$folder\data\backup' }; Logging = @{ Folder = '$folder\logs' }; Report = @{ Folder = '$folder\reports' } }"
        [IO.File]::WriteAllText($config, $text, (New-Object Text.UTF8Encoding($true)))
        $window = New-PraGuiWindow -Root $script:Root -ConfigPath $config -Theme Light
        try {
            $window.Controls.NextTitle.Text | Should -Be 'Fill in the configuration'
            $window.Controls.NextButton.Tag | Should -Be ''
            $window.Controls.TenantText.Text | Should -Be 'tenant not set'
            @($window.Controls.OvBatches.ItemsSource)[0].Value | Should -Be 'no batch: nothing converted'
            $window.Controls.ConvertApply.IsEnabled | Should -BeFalse
            $window.Controls.ConvertHint.Text | Should -BeLike 'No snapshot*'
            $window.Controls.RecoverHint.Text | Should -BeLike 'No Convert batch*'
            $window.Controls.CheckTotal.Text | Should -Be '-'
            Test-Path -LiteralPath (Join-Path $folder 'data') | Should -BeFalse
        } finally { $window.Form.Close() }
    }
    It 'opens when the configuration cannot be read, and says so' {
        $window = New-PraGuiWindow -Root $script:Root -ConfigPath (Join-Path $TestDrive 'missing.config.psd1') -Theme Light
        try {
            $window.Controls.EnvText.Text | Should -Be 'Configuration error'
            $window.Controls.NextTitle.Text | Should -Be 'Fix the configuration'
            $window.Controls.ConvertHint.Text | Should -Be 'Fix the configuration first.'
        } finally { $window.Form.Close() }
    }
}

Describe 'Window: built on a test state' {
    BeforeAll {
        $script:State = New-TestState -Folder (Join-Path $TestDrive 'state')
        $script:Window = New-PraGuiWindow -Root $script:Root -ConfigPath $script:State.Config -Version '1.2.0' -Theme Light
        $form = $script:Window.Form
        $form.WindowStartupLocation = 'Manual'; $form.Left = -6000; $form.Top = -6000; $form.ShowActivated = $false; $form.ShowInTaskbar = $false
        $form.Show()
        & $script:Gui {
            $script:Gui.Hooks.Question = { param($Title, $Text) $global:PraGuiAsked = $Text; $true }
            $script:Gui.Hooks.Notice = { param($Text) $global:PraGuiNotice = $Text }
        }
        $script:C = $script:Window.Controls
    }
    AfterAll {
        & $script:Gui { if ($script:Gui) { $script:Gui.Run = $null; $script:Gui.Timer.Stop() } }
        $script:Window.Form.Close()
        Remove-Variable -Name PraGuiAsked, PraGuiNotice, PraGuiConfirm -Scope Global -ErrorAction SilentlyContinue
    }

    It 'finds every named element and reads the state: environment, next step, resume of the unfinished batch' {
        @(& $script:Gui { $script:ControlNames }).Count | Should -Be $script:C.Count
        $script:C.EnvText.Text | Should -Be 'Environment TEST'
        $script:C.TenantText.Text | Should -Be 'contoso.onmicrosoft.com'
        $script:C.NextTitle.Text | Should -Be ('Convert batch {0} is not finished' -f $script:State.Batch)
        $script:C.NextButton.Tag | Should -Be 'Convert'
        $script:C.ConvertResumePanel.Visibility | Should -Be 'Visible'
        $script:C.CheckTotal.Text | Should -Be '3'
        $script:C.CheckNotReady.Text | Should -Be '1'
        @($script:C.SnapshotGrid.ItemsSource).Count | Should -Be 1
        $script:C.Nav.Items.Count | Should -Be 5
    }
    It 'shows the objects with their Check and their batch, and filters them as typed' {
        $table = & $script:Gui { , $script:Gui.Tables.Convert }
        $table.Rows.Count | Should -Be 3
        $table.Select("InCloud = '1'")[0]['ObjectGuid'] | Should -Be $script:Alice
        $table.Select("ObjectGuid = '$script:Bob'")[0]['StatusText'] | Should -Be 'Not ready'
        $script:C.ConvertFilter.Text = 'COMPTA'
        & $script:Gui { Update-PraGuiFilter -Page Convert }
        $table.DefaultView.Count | Should -Be 1
        $script:C.ConvertFilter.Text = ''
        $script:C.ConvertStateFilter.SelectedItem = 'Not converted'
        $table.DefaultView.Count | Should -Be 2
        $script:C.ConvertStateFilter.SelectedIndex = 0
        $table.DefaultView.Count | Should -Be 3
    }
    It 'ticks only the ready objects not converted, and the ticks choose the wave' {
        & $script:Gui { Set-PraGuiTicks -Page Convert -Shown }
        & $script:Gui { Get-PraGuiTicked $script:Gui.Tables.Convert } | Should -Be @($script:Compta)
        $script:C.ConvertWave.IsChecked | Should -BeTrue
        $selection = & $script:Gui { Get-PraGuiSelection -Page Convert }
        $selection.Key | Should -Be (Get-PraGuiSelectionKey -Prefix ('Convert|{0}' -f $script:State.Snapshot) -Guids @($script:Compta))
        $script:C.ConvertTicked.Text | Should -BeLike '1 ticked*'
    }
    It 'opens Apply only after a preview of the same selection, and closes it when the selection changes' {
        $script:C.ConvertApply.IsEnabled | Should -BeFalse
        $csv = Join-Path $script:State.Reports 'PRA2_preview_test.csv'
        Write-TestCsv -Path $csv -Rows @([ordered]@{ Identity = 'compta@contoso.com'; Kind = 'Shared'; ObjectGuid = $script:Compta; FinalStatus = 'Planned' })
        & $script:Gui {
            param($csv)
            $selection = Get-PraGuiSelection -Page Convert
            Complete-PraGuiAction -Run @{ Request = @{ Action = 'Convert'; Mode = 'Preview'; IdentityPath = 'wave.csv' }; Selection = $selection } -Result ([pscustomobject]@{ planned = 1; csvReport = $csv; exitCode = 0 })
        } $csv
        $script:C.ConvertApply.IsEnabled | Should -BeTrue
        $script:C.ConvertHint.Text | Should -BeLike '*1 planned (0 user(s), 1 shared)*'
        & $script:Gui { Set-PraGuiTicks -Page Convert }
        $script:C.ConvertApply.IsEnabled | Should -BeFalse
        $script:C.ConvertHint.Text | Should -BeLike 'A wave: tick its objects*'
    }
    It 'runs nothing when the typed confirmation is refused' {
        & $script:Gui {
            $script:Gui.Controls.ConvertAll.IsChecked = $true
            $selection = Get-PraGuiSelection -Page Convert
            $script:Gui.Pending.Convert = @{ Key = $selection.Key; Guard = $selection.Key; Resume = $false; Batch = ''; Request = @{ Action = 'Convert'; Mode = 'Preview' }; Selection = $selection
                Planned = 3; Tone = 'Success'; Confirm = 'three objects'; Hint = 'preview ok' }
            $script:Gui.Hooks.Confirm = { param($Word, $Title, $Text, $Command) $global:PraGuiConfirm = "$Word|$Command"; $false }
            Start-PraGuiApply -Page Convert
        }
        $global:PraGuiConfirm | Should -BeLike 'CONVERT|pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Convert -Mode Apply -Force'
        & $script:Gui { [bool]$script:Gui.Run } | Should -BeFalse
        $script:Window.Lines[$script:Window.Lines.Count - 1].Text | Should -Be 'Cancelled: nothing was run.'
    }
    It 'follows the events of a run: steps, counters, questions answered in the named file only' {
        $events = Join-Path $TestDrive 'stub.events.jsonl'
        & $script:Gui {
            param($events)
            $script:Gui.Run = @{ Request = @{ Action = 'Convert'; Mode = 'Apply' }; Files = @{ Events = $events }; Counts = @{ Ok = 0; Warn = 0; Fail = 0 }; SubLines = 0
                Asked = (New-Object 'Collections.Generic.HashSet[int]'); StopAsked = $false; LogFile = ''; Result = $null; Summary = $null }
            foreach ($json in @(
                    '{"kind":"step","index":2,"total":7,"title":"Users"}'
                    '{"kind":"progress","phase":"Users","done":500,"total":3000,"percent":17,"eta":"about 2 h 30 min left at this pace"}'
                    '{"kind":"item","status":"Ok","text":"alice done","identity":"alice@contoso.com"}'
                    '{"kind":"item","status":"Warn","text":"bob pending","identity":"bob@contoso.com"}'
                    ('{{"kind":"ask","id":1,"title":"Entra Connect","text":"Run a delta sync. Done?","answer":"{0}"}}' -f ("$events.answer-1" -replace '\\', '\\'))
                    ('{{"kind":"ask","id":2,"title":"x","text":"elsewhere","answer":"{0}"}}' -f ((Join-Path $TestDrive 'elsewhere.txt') -replace '\\', '\\'))
                    '{"kind":"result","exitCode":0,"planned":0}')) { Invoke-PraGuiEvent -Record ($json | ConvertFrom-Json) }
        } $events
        $script:C.RunStep.Text | Should -Be 'Step 2/7 · Users'
        $script:C.RunProgress.Value | Should -BeGreaterThan 0.14
        $script:C.RunObjects.Text | Should -Be 'Users: 500 of 3000 (17%) · about 2 h 30 min left at this pace'
        $script:C.RunObjects.Visibility | Should -Be ([Windows.Visibility]::Visible)
        $script:C.RunObjectsBar.Value | Should -BeGreaterThan 0.16
        $script:C.RunOk.Text | Should -Be '1'
        $script:C.RunWarn.Text | Should -Be '1'
        $script:C.RunCurrent.Text | Should -Be 'Object: bob@contoso.com'
        Get-Content -LiteralPath "$events.answer-1" -Raw | Should -Be 'yes'
        $global:PraGuiAsked | Should -Be 'elsewhere'
        Test-Path -LiteralPath (Join-Path $TestDrive 'elsewhere.txt') | Should -BeFalse
        (& $script:Gui { $script:Gui.Run.Result }).exitCode | Should -Be 0
    }
    It 'hides the object progress of the previous step when a new step starts' {
        & $script:Gui {
            Invoke-PraGuiEvent -Record (@{ kind = 'progress'; phase = 'Shared mailboxes'; done = 1; total = 5; percent = 20; eta = '' } | ConvertTo-Json | ConvertFrom-Json)
        }
        $script:C.RunObjects.Visibility | Should -Be ([Windows.Visibility]::Visible)
        & $script:Gui {
            Invoke-PraGuiEvent -Record (@{ kind = 'step'; index = 3; total = 7; title = 'Shared mailboxes' } | ConvertTo-Json | ConvertFrom-Json)
        }
        $script:C.RunObjects.Text | Should -BeNullOrEmpty
        $script:C.RunObjects.Visibility | Should -Be ([Windows.Visibility]::Collapsed)
        $script:C.RunObjectsBar.Visibility | Should -Be ([Windows.Visibility]::Collapsed)
    }
    It 'refuses to close while a run is going' {
        $global:PraGuiNotice = ''
        $script:Window.Form.Close()
        $script:Window.Form.IsVisible | Should -BeTrue
        $global:PraGuiNotice | Should -BeLike 'An action is running*'
        & $script:Gui { $script:Gui.Run = $null; Set-PraGuiBusy -Busy $false }
    }
    It 'runs Check in its own PowerShell and shows its result (the sign-in fails: test tenant)' {
        if (-not (Get-PraGuiEngine -Action Check).Ok) { Set-ItResult -Skipped -Because 'PowerShell 7.4 or later is not installed'; return }
        & $script:Gui { Start-PraGuiCheck }
        $script:C.CheckRun.IsEnabled | Should -BeFalse
        $script:C.StopRun.Visibility | Should -Be 'Visible'
        $clock = [Diagnostics.Stopwatch]::StartNew()
        while ((& $script:Gui { [bool]$script:Gui.Run }) -and $clock.Elapsed.TotalSeconds -lt 240) { Invoke-TestPump; Start-Sleep -Milliseconds 100 }
        & $script:Gui { [bool]$script:Gui.Run } | Should -BeFalse
        $script:C.ResultCard.Visibility | Should -Be 'Visible'
        $script:C.ResultTitle.Text | Should -Not -BeNullOrEmpty
        $script:C.CheckRun.IsEnabled | Should -BeTrue
        $events = @(Get-ChildItem -LiteralPath (Join-Path $script:State.Data 'gui') -Filter 'run-*-check.events.jsonl')
        $events.Count | Should -Be 1
        @(Get-Content -LiteralPath $events[0].FullName | ForEach-Object { ($_ | ConvertFrom-Json).kind }) | Should -Contain 'start'
        @($script:Window.Lines | Where-Object { $_.Text -like 'Started: pwsh -File .\Invoke-PraCloudMailbox.ps1 -Action Check*' }).Count | Should -Be 1
    }
}
