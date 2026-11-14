<#
.SYNOPSIS
    PRA Cloud Mailbox - tests of the reports and of the folders they are written to: the global report of a batch (every object and every event, a page without external
    resource, its script run in a headless browser), the CSV files (formula guard, UTF-8 with BOM), the output folder names (letter case) and the private creation of the
    data, log and report folders. Pester 5+, offline.
.NOTES
    Author  : Nicolas Fabert
#>
#Requires -Version 7.5

BeforeDiscovery {
    $script:BrowserAvailable = [bool](@("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") |
            Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1)
}

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Root = New-PcmTestPackage -Repository $script:Repo
    Import-Module (Join-Path $script:Root 'module\PRA.Common.psm1') -Force

    function script:New-BatchHistory {
        # 1,200 users converted in one batch, one of them (u0007) with 250 more events, one rolled back by a Recover batch, and one event of the batch itself.
        $convert = [pscustomobject]@{ id = 'c0nvert1'; action = 'Convert'; convert_batch = ''; status = 'Complete'; created_utc = '2026-11-01T08:00:00.0000000Z'; updated_utc = '2026-11-01T09:00:00.0000000Z'; snapshot_id = 4; account = 'CONTOSO\admin'; computer = 'ADM1'; tool_version = '2.1.0' }
        $recover = [pscustomobject]@{ id = 'r3cover1'; action = 'Recover'; convert_batch = 'c0nvert1'; status = 'Complete'; created_utc = '2026-11-05T08:00:00.0000000Z'; updated_utc = '2026-11-05T09:00:00.0000000Z'; snapshot_id = 4; account = 'CONTOSO\admin'; computer = 'ADM1'; tool_version = '2.1.0' }
        $items = [System.Collections.Generic.List[object]]::new(); $events = [System.Collections.Generic.List[object]]::new()
        foreach ($number in 1..1200) {
            $guid = 'x{0}' -f $number
            $items.Add([pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = $guid; identity = ('u{0:0000}@contoso.com' -f $number); kind = 'User'; status = 'Done'; step = 'MailboxReady'; message = '' })
            $events.Add([pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = $guid; time_utc = '2026-11-01T08:00:00.0000000Z'; step = 'MailboxReady'; outcome = 'Ok'; detail = ('simple object {0} <b>x</b>' -f $number) })
        }
        foreach ($number in 1..250) {
            $events.Add([pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = 'x7'; time_utc = ('2026-11-02T10:{0:00}:{1:00}.0000000Z' -f [math]::Floor($number / 60), ($number % 60)); step = ('Step{0}' -f $number); outcome = 'Ok'; detail = ('event {0}' -f $number) })
        }
        $items.Add([pscustomobject]@{ batch_id = 'r3cover1'; object_guid = 'x1'; identity = 'u0001@contoso.com'; kind = 'User'; status = 'Done'; step = 'Hold'; message = '' })
        $items.Add([pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = 'x1201'; identity = '=cmd@contoso.com'; kind = 'Shared'; status = 'Failed'; step = 'ConvertShared'; message = '+bad note' })
        $events.Add([pscustomobject]@{ batch_id = 'r3cover1'; object_guid = 'x1'; time_utc = '2026-11-05T08:30:00.0000000Z'; step = 'Hold'; outcome = 'Ok'; detail = 'case hold stamped' })
        $events.Add([pscustomobject]@{ batch_id = 'c0nvert1'; object_guid = ''; time_utc = '2026-11-01T08:00:30.0000000Z'; step = 'GroupSoaCloud'; outcome = 'Ok'; detail = 'licence group managed in the cloud' })
        [pscustomobject]@{ Batches = @($convert, $recover); Items = $items.ToArray(); Events = $events.ToArray()
            Runs = @([pscustomobject]@{ run_id = 'r1'; batch_id = 'c0nvert1'; action = 'Convert'; mode = 'Apply'; started_utc = '2026-11-01T08:00:00.0000000Z'; ended_utc = '2026-11-01T09:00:00.0000000Z'; status = 'Success'; exit_code = 0; done_count = 1200; pending_count = 0; error_count = 0; skipped_count = 0
                    source_text = 'every object of the snapshot'; overrides_text = '' }) }
    }
    function script:New-GlobalContext([string]$Name) {
        @{ ReportFolder = (Join-Path $TestDrive $Name); Version = '2.1.0'; Config = @{ Environment = 'TEST' }; Warnings = 0; LogFile = ''; Root = $script:Root }
    }
    function script:Invoke-BrowserDom {
        # The DOM of a page after its script ran, from a headless Edge or Chrome (virtual time, so the page decodes its data before the DOM is written).
        param([Parameter(Mandatory)][string]$Path, [string]$Hash = '')
        $browser = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Google\Chrome\Application\chrome.exe") |
            Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
        $browserProfile = Join-Path ([IO.Path]::GetTempPath()) ('pcm-browser-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        try {
            $start = [Diagnostics.ProcessStartInfo]::new($browser)
            foreach ($argument in @('--headless=new', '--disable-gpu', '--no-first-run', "--user-data-dir=$browserProfile", '--virtual-time-budget=30000', '--dump-dom', (([Uri]$Path).AbsoluteUri + $Hash))) { $start.ArgumentList.Add($argument) }
            $start.RedirectStandardOutput = $true; $start.RedirectStandardError = $true; $start.UseShellExecute = $false
            $process = [Diagnostics.Process]::Start($start)
            $output = $process.StandardOutput.ReadToEndAsync(); $null = $process.StandardError.ReadToEndAsync()
            if (-not $process.WaitForExit(180000)) { $process.Kill($true) }
            return $output.Result
        } finally { Remove-Item -LiteralPath $browserProfile -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

AfterAll {
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Global report: no object and no event is left out' {
    BeforeAll {
        $script:history = New-BatchHistory
        $script:context = New-GlobalContext 'global'
        $script:result = Write-PcmGlobalReport -Context $script:context -History $script:history -ConvertBatchId 'c0nvert1' -FollowUp @([pscustomobject]@{ Identity = 'u0001@contoso.com'; Expect = 'Deprovisioned'; Met = $false; Detail = 'still a user' })
        $script:html = [IO.File]::ReadAllText($script:result.Html)
        $script:data = Read-PcmGlobalReportData -Html $script:html
    }
    It 'writes the page and the two CSV files' {
        $script:result.Html | Should -Exist
        $script:result.Csv | Should -Exist
        $script:result.EventsCsv | Should -Exist
        $script:result.EventsCsv | Should -Match '_global_events\.csv$'
    }
    It 'holds the 1,201 objects and every one of the 1,452 events, with no cap and no notice about a cap' {
        $script:html | Should -Not -MatchExactly 'More than'
        @($script:data.Objects).Count | Should -Be 1201
        @($script:data.Events).Count | Should -Be 1452
        $script:data.Header.objects.count | Should -Be 1201
        $script:data.Header.events.count | Should -Be 1452
    }
    It 'gives the journey of every object, the converted ones past the 500th included, and the count it announces' {
        $position = @{}
        for ($i = 0; $i -lt $script:data.Objects.Count; $i++) { $position[[string]$script:data.Objects[$i].ObjectGuid] = $i }
        $perObject = @{}
        foreach ($row in $script:data.Events) { if ($null -ne $row.Object) { $perObject[[int]$row.Object] = 1 + [int]$perObject[[int]$row.Object] } }
        foreach ($object in $script:data.Objects) { [int]$perObject[$position[[string]$object.ObjectGuid]] | Should -Be ([int]$object.Events) -Because $object.Identity }
        ($script:data.Objects | Where-Object Identity -eq 'u1100@contoso.com').Verdict | Should -Be 'Converted'
        @($script:data.Events | Where-Object { $_.Object -eq $position['x1100'] }).Count | Should -Be 1
        ($script:data.Objects | Where-Object Identity -eq 'u0001@contoso.com').Verdict | Should -Be 'Rolled back'
    }
    It 'keeps all 251 events of an object, in order, complete' {
        $position = [array]::IndexOf(@($script:data.Objects | ForEach-Object { $_.ObjectGuid }), 'x7')
        $journey = @($script:data.Events | Where-Object { $_.Object -eq $position })
        $journey.Count | Should -Be 251
        $journey[0].Step | Should -Be 'MailboxReady'
        $journey[-1].Step | Should -Be 'Step250'
        $journey[-1].Detail | Should -Be 'event 250'
        $journey[-1].Time | Should -Be '2026-11-02 10:04:10 UTC'
    }
    It 'keeps the events of the batch itself, with no object' {
        $batch = @($script:data.Events | Where-Object { $null -eq $_.Object })
        $batch.Count | Should -Be 1
        $batch[0].Step | Should -Be 'GroupSoaCloud'
    }
    It 'writes the objects CSV without a journey column, and the events CSV with one line per event' {
        $objects = @(Import-Csv -LiteralPath $script:result.Csv -Delimiter ';')
        $objects.Count | Should -Be 1201
        $objects[0].PSObject.Properties.Name | Should -Not -Contain 'Journey'
        ($objects | Where-Object Identity -eq 'u0007@contoso.com').EventCount | Should -Be 251
        $events = @(Import-Csv -LiteralPath $script:result.EventsCsv -Delimiter ';')
        $events.Count | Should -Be 1452
        @($events | Where-Object Identity -eq 'u0007@contoso.com').Count | Should -Be 251
        @($events | Where-Object { $_.Identity -eq 'u0007@contoso.com' -and $_.Step -eq 'Step250' -and $_.Detail -eq 'event 250' }).Count | Should -Be 1
        @($events | Where-Object { -not $_.ObjectGuid -and $_.Step -eq 'GroupSoaCloud' }).Count | Should -Be 1
    }
    It 'writes both CSV files as UTF-8 with a byte order mark, and puts an apostrophe before a formula' {
        foreach ($path in @($script:result.Csv, $script:result.EventsCsv)) {
            $bytes = [IO.File]::ReadAllBytes($path)
            ('{0:X2}{1:X2}{2:X2}' -f $bytes[0], $bytes[1], $bytes[2]) | Should -Be 'EFBBBF'
        }
        $objects = @(Import-Csv -LiteralPath $script:result.Csv -Delimiter ';')
        ($objects | Where-Object Verdict -eq 'Failed').Identity | Should -Be "'=cmd@contoso.com"
        ($objects | Where-Object Verdict -eq 'Failed').Note | Should -Be "'+bad note"
        @(Import-Csv -LiteralPath $script:result.EventsCsv -Delimiter ';' | Where-Object { $_.Identity -like "'=*" }).Count | Should -Be 0
    }
    It 'keeps the summary readable without script: the totals, the batches and the runs are plain HTML' {
        $plain = [regex]::Replace($script:html, '(?s)<script.*?</script>', '')
        $plain | Should -Match 'Batch c0nvert1 from end to end'
        $plain | Should -Match '<div class="tile-label">Objects</div><div class="tile-value">1201</div>'
        $plain | Should -Match '<div class="tile-label">Converted</div><div class="tile-value">1199</div>'
        $plain | Should -Match '<div class="tile-label">Events</div><div class="tile-value">1452</div>'
        $plain | Should -Match '<td class="mono">r3cover1</td>'
        $plain | Should -Match 'every object of the snapshot'
        $plain | Should -Match 'PraCloudMailbox_TEST_batchc0nvert1_global\.csv'
        $plain | Should -Match 'PraCloudMailbox_TEST_batchc0nvert1_global_events\.csv'
    }
    It 'has no external resource and no network call' {
        $script:html | Should -Not -Match 'https?://'
        $script:html | Should -Not -Match '(?i)<link\b|@import|\bsrc\s*=|(?<![a-z])url\(|<iframe|<img|<object|<embed'
        $template = [IO.File]::ReadAllText((Join-Path $script:Root 'templates\GlobalReport.template.html'))
        $template | Should -Not -Match '(?i)\bfetch\s*\(|XMLHttpRequest|WebSocket|sendBeacon|\bimport\s*\(|EventSource|navigator\.serviceWorker'
        $template | Should -Not -Match 'https?://'
    }
    It 'reads the same data from a page written in several chunks' {
        $model = Get-PcmGlobalModel -History $script:history -Version '2.1.0' -Environment 'TEST'
        $writer = New-Object IO.StringWriter
        Write-PcmGlobalReportHtml -Model $model -Writer $writer -TemplateFolder (Join-Path $script:Root 'templates') -ChunkRows 400
        $text = $writer.ToString()
        ([regex]::Matches($text, 'data-kind="objects"')).Count | Should -Be 4
        ([regex]::Matches($text, 'data-kind="events"')).Count | Should -Be 4
        $chunked = Read-PcmGlobalReportData -Html $text
        @($chunked.Objects).Count | Should -Be 1201
        @($chunked.Events).Count | Should -Be 1452
        @($chunked.Events | Where-Object Detail -eq 'case hold stamped').Count | Should -Be 1
    }
    It 'writes the page of a batch with no object and no event' {
        $history = [pscustomobject]@{ Batches = @($script:history.Batches[0]); Items = @(); Events = @(); Runs = @() }
        $result = Write-PcmGlobalReport -Context (New-GlobalContext 'global-empty') -History $history -ConvertBatchId 'c0nvert1'
        $data = Read-PcmGlobalReportData -Html ([IO.File]::ReadAllText($result.Html))
        @($data.Objects).Count | Should -Be 0
        @($data.Events).Count | Should -Be 0
        @(Import-Csv -LiteralPath $result.Csv -Delimiter ';').Count | Should -Be 0
        @(Get-Content -LiteralPath $result.EventsCsv)[0] | Should -Match 'ObjectGuid'
    }
    It 'still gives the summary and the CSV files when the template is missing' {
        $model = Get-PcmGlobalModel -History $script:history -Version '2.1.0' -Environment 'TEST'
        $writer = New-Object IO.StringWriter
        Write-PcmGlobalReportHtml -Model $model -Writer $writer -TemplateFolder (Join-Path $TestDrive 'no-templates')
        $writer.ToString() | Should -Match 'Template missing: only the CSV files'
    }
}

Describe 'Global report: the script of the page, in a headless browser' {
    BeforeAll {
        $script:pageHtml = (Write-PcmGlobalReport -Context (New-GlobalContext 'global-page') -History (New-BatchHistory) -ConvertBatchId 'c0nvert1').Html
    }
    It 'decodes every object and every event, shows at most 200 grid rows, and opens the whole journey of an object from its identity (letter case ignored)' -Skip:(-not $script:BrowserAvailable) {
        $dom = Invoke-BrowserDom -Path $script:pageHtml -Hash '#object=U0007@CONTOSO.COM'
        $dom | Should -Match 'data-state="ready"'
        $dom | Should -Match 'data-objects="1201"'
        $dom | Should -Match 'data-events="1452"'
        $dom | Should -Match '<h2 id="journey-title">Journey of u0007@contoso.com</h2>'
        $dom | Should -Match '<dialog id="journey"[^>]* data-events="251"'
        $dom | Should -Match '1 to 100 of 251, page 1 of 3'
        $dom | Should -Match 'Step99'
        $rows = ([regex]::Matches($dom, 'class="grid-row')).Count
        $rows | Should -BeGreaterThan 0
        $rows | Should -BeLessOrEqual 200
        $dom | Should -Match '1,201 object\(s\), no filter'
    }
    It 'says plainly that the data is incomplete, and shows no object, when a chunk is missing' -Skip:(-not $script:BrowserAvailable) {
        $tampered = Join-Path $TestDrive 'tampered.html'
        $text = [IO.File]::ReadAllText($script:pageHtml)
        $first = [regex]::Match($text, '<script type="application/x-pra-chunk" data-kind="events"[^>]*>[^<]*</script>\r?\n')
        $first.Success | Should -BeTrue
        [IO.File]::WriteAllText($tampered, $text.Remove($first.Index, $first.Length), (New-Object Text.UTF8Encoding($true)))
        $dom = Invoke-BrowserDom -Path $tampered
        $dom | Should -Match 'data-state="error"'
        $dom | Should -Match 'The report data is incomplete'
        $dom | Should -Not -Match 'class="grid-row'
    }
}

Describe 'Output folders: the letter case of the name does not matter' {
    BeforeAll {
        function script:Import-FolderConfig([string]$Logs, [string]$Reports, [string]$Data) {
            $path = Join-Path $TestDrive ('folders-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.psd1')
            Set-Content -LiteralPath $path -Encoding UTF8 -Value "@{ Environment = 'X'; Scope = @{ Mode = 'Auto' }; EntraConnect = @{ Server = 'aadc' }; Logging = @{ Folder = '$Logs' }; Report = @{ Folder = '$Reports' }; Store = @{ Path = '$Data\pcm.db'; BackupFolder = '$Data\backup'; JournalPath = '$Data\journal.db' } }"
            Import-PcmConfiguration -Path $path -Root $script:Root
        }
    }
    It 'accepts the folders written as <Logs>, <Reports> and <Data>, and uses the canonical spelling when they do not exist yet' -ForEach @(
        @{ Logs = '.\Logs'; Reports = '.\REPORTS'; Data = '.\Data' }
        @{ Logs = '.\LOGS\gui'; Reports = '.\Reports\sub'; Data = '.\DATA\nested' }
        @{ Logs = '.\logs'; Reports = '.\reports'; Data = '.\data' }
    ) {
        $config = Import-FolderConfig -Logs $Logs -Reports $Reports -Data $Data
        $config.Logging.Folder | Should -MatchExactly '\\logs(\\gui)?$'
        $config.Report.Folder | Should -MatchExactly '\\reports(\\sub)?$'
        $config.Store.Path | Should -MatchExactly '\\data(\\nested)?\\pcm\.db$'
        $config.Store.BackupFolder | Should -MatchExactly '\\data(\\nested)?\\backup$'
    }
    It 'accepts a folder that exists, written in another letter case, and keeps its spelling' {
        $null = New-Item -ItemType Directory -Path (Join-Path $script:Root 'logs') -Force
        try {
            $config = Import-FolderConfig -Logs '.\LOGS\gui' -Reports '.\reports' -Data '.\data'
            $config.Logging.Folder | Should -MatchExactly '\\logs\\gui$'
        } finally { Remove-Item -LiteralPath (Join-Path $script:Root 'logs') -Recurse -Force -ErrorAction SilentlyContinue }
    }
    It 'still refuses a folder of the package that is not the one of the file' {
        { Import-FolderConfig -Logs '.\Reports' -Reports '.\reports' -Data '.\data' } | Should -Throw '*is not the logs folder*'
        { Import-FolderConfig -Logs '.\logs' -Reports '.\elsewhere' -Data '.\data' } | Should -Throw '*is not the reports folder*'
    }
}

Describe 'Private creation of the data, log and report folders' {
    It 'creates a missing folder private: protected permissions, the account, SYSTEM and the Administrators only' {
        $path = Join-Path $TestDrive 'private\logs'
        Initialize-PcmPrivateFolder -Path $path
        $path | Should -Exist
        $acl = Get-Acl -LiteralPath $path
        $acl.AreAccessRulesProtected | Should -BeTrue
        $sids = @($acl.Access | ForEach-Object { $_.IdentityReference.Translate([Security.Principal.SecurityIdentifier]).Value } | Sort-Object -Unique)
        $sids | Should -Contain 'S-1-5-18'
        $sids | Should -Contain 'S-1-5-32-544'
        $sids | Should -Contain ([Security.Principal.WindowsIdentity]::GetCurrent().User.Value)
        $sids.Count | Should -BeLessOrEqual 3
        (Get-Acl -LiteralPath (Split-Path $path -Parent)).AreAccessRulesProtected | Should -BeFalse
    }
    It 'leaves an existing folder exactly as it is' {
        $path = Join-Path $TestDrive 'existing'
        $null = New-Item -ItemType Directory -Path $path
        Initialize-PcmPrivateFolder -Path $path
        (Get-Acl -LiteralPath $path).AreAccessRulesProtected | Should -BeFalse
    }
    It 'says why a folder cannot be created, and never goes on silently' {
        $file = Join-Path $TestDrive 'a-file'
        Set-Content -LiteralPath $file -Value 'x'
        { Initialize-PcmPrivateFolder -Path $file } | Should -Throw '*a file exists where the folder is expected*'
        { Initialize-PcmPrivateFolder -Path (Join-Path $TestDrive 'bad|name') } | Should -Throw '*cannot be created*'
        Join-Path $TestDrive 'bad|name' | Should -Not -Exist
    }
    It 'creates the report folder of the global report private' {
        $context = New-GlobalContext 'reports-private'
        $null = Write-PcmGlobalReport -Context $context -History (New-BatchHistory) -ConvertBatchId 'c0nvert1'
        (Get-Acl -LiteralPath $context.ReportFolder).AreAccessRulesProtected | Should -BeTrue
    }
}

Describe 'CSV report of a run' {
    It 'puts an apostrophe before a text that starts like a formula, and nowhere else' -ForEach @(
        @{ Text = '=1+1'; Expected = "'=1+1" }, @{ Text = '+x'; Expected = "'+x" }, @{ Text = '-x'; Expected = "'-x" }, @{ Text = '@x'; Expected = "'@x" }
        @{ Text = "`tx"; Expected = "'`tx" }, @{ Text = "`rx"; Expected = "'`rx" }, @{ Text = 'alice@contoso.com'; Expected = 'alice@contoso.com' }, @{ Text = ''; Expected = '' }
    ) {
        ConvertTo-PcmCsvCell $Text | Should -Be $Expected
    }
    It 'keeps booleans and numbers of a row as they are' {
        $row = ConvertTo-PcmCsvRow ([pscustomobject]@{ Identity = '=x'; Count = 3; Flag = $true; Empty = $null })
        $row.Identity | Should -Be "'=x"
        $row.Count | Should -Be 3
        $row.Flag | Should -BeTrue
        $row.Empty | Should -BeNullOrEmpty
    }
    It 'writes the CSV report of a run as UTF-8 with a byte order mark, with the formula guard, also with no row' {
        $csv = Join-Path $TestDrive 'run.csv'
        $context = @{ Action = 'Check'; Mode = 'Preview'; ExitCode = 0; CloudConnectFailed = $false; ResultStatus = ''
            Issues = (New-Object 'Collections.Generic.List[object]'); Rows = (New-Object 'Collections.Generic.List[object]'); SkippedMembers = @() }
        [void]$context.Rows.Add([ordered]@{ Identity = '=cmd@contoso.com'; Kind = 'User'; ObjectGuid = 'g1'; Action = 'Check'; FinalStatus = 'Success'; Detail = '@detail' })
        InModuleScope PRA.Common -Parameters @{ Context = $context; Csv = $csv } -ScriptBlock { param($Context, $Csv) Write-PcmReport -Context $Context -Reports @{ Csv = $Csv; Html = ''; CsvFailed = $false; HtmlFailed = $false } }
        $bytes = [IO.File]::ReadAllBytes($csv)
        ('{0:X2}{1:X2}{2:X2}' -f $bytes[0], $bytes[1], $bytes[2]) | Should -Be 'EFBBBF'
        $rows = @(Import-Csv -LiteralPath $csv -Delimiter ';')
        $rows[0].Identity | Should -Be "'=cmd@contoso.com"
        $rows[0].Detail | Should -Be "'@detail"
        $context.Rows.Clear()
        InModuleScope PRA.Common -Parameters @{ Context = $context; Csv = $csv } -ScriptBlock { param($Context, $Csv) Write-PcmReport -Context $Context -Reports @{ Csv = $Csv; Html = ''; CsvFailed = $false; HtmlFailed = $false } }
        $bytes = [IO.File]::ReadAllBytes($csv)
        ('{0:X2}{1:X2}{2:X2}' -f $bytes[0], $bytes[1], $bytes[2]) | Should -Be 'EFBBBF'
        [IO.File]::ReadAllText($csv) | Should -Match 'Identity'
    }
}

Describe 'Text of a CSV file written by a person' {
    It 'reads a UTF-8 file with or without a byte order mark, and a UTF-16 file' {
        $text = "Identity;Name`r`nalice@contoso.com;Zo$([char]0x00E9)"
        foreach ($case in @(@{ Name = 'utf8bom'; Bytes = ([byte[]](239, 187, 191) + [Text.Encoding]::UTF8.GetBytes($text)) }, @{ Name = 'utf8'; Bytes = [Text.Encoding]::UTF8.GetBytes($text) },
                @{ Name = 'utf16'; Bytes = ([byte[]](255, 254) + [Text.Encoding]::Unicode.GetBytes($text)) })) {
            $path = Join-Path $TestDrive ($case.Name + '.csv')
            [IO.File]::WriteAllBytes($path, $case.Bytes)
            Read-PcmCsvText -Path $path | Should -Be $text -Because $case.Name
        }
    }
    It 'reads a file that is not UTF-8 (Excel or Notepad ANSI) with the ANSI code page of the system' -Skip:([Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage -ne 1252) {
        $path = Join-Path $TestDrive 'ansi.csv'
        [IO.File]::WriteAllBytes($path, [Text.Encoding]::GetEncoding(1252).GetBytes("Identity`r`nZo$([char]0x00E9)@contoso.com"))
        Read-PcmCsvText -Path $path | Should -Be "Identity`r`nZo$([char]0x00E9)@contoso.com"
    }
}
