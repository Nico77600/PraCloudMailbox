<#
.SYNOPSIS
    Synthetic engine for the tests and the documentation pictures of the console (repository tool, never shipped, never signed).
.DESCRIPTION
    Stands in for Invoke-PraCloudMailbox.ps1 inside a temporary copy of the tool. It takes the same parameters as the real script and writes the same
    events as the real engine (start, step, item, progress, summary, result) to PRA_EVENT_FILE, with synthetic names only. The scenario file next to it
    (scenario.json: objects, and optionally batch and skipped members) lists what the run works on. A file named hold.<name> next to the script makes
    the wait of that name last until the harness deletes it, so a run can be photographed half way. A file named fail.flag makes the run fail after
    the plan, as a run does when Microsoft 365 refuses the sign-in. Nothing here reads Exchange, Active Directory, Microsoft Graph or Entra Connect.
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', '', Justification = 'The stub takes every parameter of the real script, used or not.')]
param($Action, $Mode, $Identity, $CsvPath, $SearchBase, $Group, $MaxObjects, $Scope, $Snapshot, $Batch, $Expect, [switch]$Once, [switch]$ResolveOnly, [pscredential]$Credential, $At, $ConfigPath, $BaseConfigPath, [switch]$Force, [switch]$PassThru, [switch]$Gui)
$ErrorActionPreference = 'Stop'
$here = $PSScriptRoot
$scenario = [IO.File]::ReadAllText((Join-Path $here 'scenario.json')) | ConvertFrom-Json
function Ev($h) { $h['time'] = (Get-Date).ToString('o'); [IO.File]::AppendAllText($env:PRA_EVENT_FILE, (($h | ConvertTo-Json -Depth 6 -Compress) + "`n"), (New-Object Text.UTF8Encoding($false))) }
function Hold($name) { $flag = Join-Path $here ('hold.' + $name); $limit = [datetime]::UtcNow.AddSeconds(90); while ((Test-Path $flag) -and [datetime]::UtcNow -lt $limit) { Start-Sleep -Milliseconds 200 } }
function Stopped { return [bool]($env:PRA_STOP_FILE -and (Test-Path $env:PRA_STOP_FILE)) }
$apply = ($Mode -eq 'Apply')
# What the console asked, for the tests (the file of the last run, and the wave file it passed).
[IO.File]::WriteAllText((Join-Path $here 'last-args.json'), ([ordered]@{ Action = $Action; Mode = $Mode; Identity = $Identity; CsvPath = $CsvPath; SearchBase = $SearchBase; Group = $Group; MaxObjects = $MaxObjects; Scope = $Scope
            Snapshot = $Snapshot; Batch = $Batch; Expect = $Expect; Once = [bool]$Once; ResolveOnly = [bool]$ResolveOnly; ConfigPath = $ConfigPath; BaseConfigPath = $BaseConfigPath; Force = [bool]$Force; Edition = $PSVersionTable.PSEdition } | ConvertTo-Json))
if ($CsvPath -and (Test-Path -LiteralPath $CsvPath)) { Copy-Item -LiteralPath $CsvPath -Destination (Join-Path $here 'wave-copy.csv') -Force }
if ($BaseConfigPath -and (Test-Path -LiteralPath $ConfigPath)) { Copy-Item -LiteralPath $ConfigPath -Destination (Join-Path $here 'effective-copy.psd1') -Force }
$objs = @($scenario.objects)
if ($Identity) { $objs = @($objs | Where-Object { $_.upn -eq $Identity -or $_.guid -eq $Identity } | Select-Object -First 1) }
if ($CsvPath -and (Test-Path -LiteralPath $CsvPath)) {
    $wanted = @(Import-Csv -LiteralPath $CsvPath | ForEach-Object { [string]$_.Identity })
    $objs = @($objs | Where-Object { $wanted -contains $_.guid -or $wanted -contains $_.upn })
}
if ($Action -eq 'Schedule') { $objs = @() }
if ($Scope -eq 'UsersOnly') { $objs = @($objs | Where-Object { -not $_.shared }) } elseif ($Scope -eq 'SharedOnly') { $objs = @($objs | Where-Object { $_.shared }) }
if ($MaxObjects -and [int]$MaxObjects -gt 0) { $objs = @($objs | Select-Object -First ([int]$MaxObjects)) }
$batchId = if ($Batch) { [string]$Batch } elseif ($scenario.batchId) { [string]$scenario.batchId } else { '1a2b3c4d' }
$steps = @(switch ($Action) {
    'Collect' { @('Exchange on-premises connection', 'Mailboxes in scope', 'Shared mailbox permissions', 'Contacts and distribution groups', 'Target groups', 'Saving the snapshot') }
    'Schedule' { @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task') }
    'Check' { if ($ResolveOnly) { @('Collect snapshot') } elseif ($Batch) { @('Batch journal', 'Collect snapshot', 'Microsoft 365 connection', 'Follow-up of the batch') } else { @('Collect snapshot', 'Microsoft 365 connection', 'Tenant prerequisites', 'Objects') } }
    'Convert' { @('Collect snapshot', 'Microsoft 365 connection', 'Readiness', 'Plan and confirmation', 'Users: source of authority and Exchange plan', 'Users: cloud mailboxes', 'Shared mailboxes') }
    'Recover' { @('Collect snapshot', 'Microsoft 365 connection', 'Prerequisites', 'Plan and confirmation', 'Users: back on-premises', 'Shared mailboxes', 'Licence group') }
})
$total = $steps.Count
$started = Get-Date
$reports = Join-Path $here 'reports'; $logs = Join-Path $here 'logs'
$null = New-Item -ItemType Directory -Path $reports, $logs -Force
$stem = '{0}-{1}-{2}' -f $Action, $Mode, [guid]::NewGuid().ToString('N').Substring(0, 6)
$logFile = Join-Path $logs ($stem + '.log'); [IO.File]::WriteAllText($logFile, 'run log')
$csvFile = Join-Path $reports ($stem + '.csv'); $htmlFile = Join-Path $reports ($stem + '.html'); [IO.File]::WriteAllText($htmlFile, '<html><body>report</body></html>')
Ev @{ kind = 'start'; title = 'PRA Cloud Mailbox'; action = $Action; mode = $Mode; version = '2.1.1'; logFile = $logFile; details = [ordered]@{ Action = ('{0} · {1}' -f $Action, $Mode); Batch = $(if ($Action -in 'Convert', 'Recover' -and $apply -and -not (Test-Path (Join-Path $here 'fail.flag'))) { $batchId } else { '' }) } }
# The group a lookup finds: the one of the scenario (PRA-Wave1, captured by Collect in snapshot 42), in any letter case.
$groupName = 'PRA-Wave1'
if ($ResolveOnly -and $Group -ine $groupName) {
    $message = ('Group {0} is not in snapshot 42 nor in Microsoft Entra. Nothing was changed: check the name.' -f $Group)
    Ev @{ kind = 'plan'; steps = @(@{ id = 's1'; label = 'Collect snapshot'; count = 0 }) }
    Ev @{ kind = 'step'; index = 1; total = 1; title = 'Collect snapshot'; id = 's1' }
    Ev @{ kind = 'item'; status = 'Fail'; text = $message; identity = ''; name = ''; stepId = 's1' }
    Ev @{ kind = 'step'; id = 's1'; state = 'fail' }
    Ev @{ kind = 'result'; status = 'Failed'; exitCode = 1; action = $Action; mode = $Mode; success = 0; error = 0; pending = 0; skipped = 0; planned = 0; total = 0; seconds = 0; batchId = ''; nextSteps = @(); logFile = $logFile; csvReport = ''; htmlReport = ''; warnings = 0; issues = @($message) }
    exit 1
}
Ev @{ kind = 'plan'; steps = @(for ($p = 0; $p -lt $steps.Count; $p++) { @{ id = ('s{0}' -f ($p + 1)); label = $steps[$p]; count = 0 } }) }
if ($Action -notin 'Collect', 'Schedule') { Ev @{ kind = 'item'; status = 'Ok'; text = 'Cloud modules: Microsoft.Graph.Authentication 2.41.1, ExchangeOnlineManagement 3.10.1'; identity = ''; name = ''; stepId = '' } }
$rows = New-Object Collections.Generic.List[object]
$warnings = 0; $errors = 0; $done = 0; $skippedCount = 0; $left = 0
function Get-DoneText($o, $title) {
    # The wording of the real engine for an object that is finished.
    switch ($Action) {
        'Convert' { if ($o.shared) { return ('{0}: shared mailbox {1}, tagged, 2/2 permission(s) granted' -f $o.upn, $o.guid) } else { return ('{0}: source of authority in the cloud, Exchange plan assigned' -f $o.upn) } }
        'Recover' { if ($o.shared) { return ('{0}: recreated from AD ({1}); the cloud mailbox {2} is inactive under the case hold' -f $o.upn, [guid]::NewGuid().ToString(), $o.guid) } else { return ('{0}: back on-premises (MailUser, on-premises GUID); the cloud mailbox stays as ComponentShared under the case hold' -f $o.upn) } }
        default { return ('{0}: ready' -f $o.upn) }
    }
}
for ($i = 0; $i -lt $total; $i++) {
    $title = $steps[$i]
    Ev @{ kind = 'step'; index = $i + 1; total = $total; title = $title; id = ('s{0}' -f ($i + 1)) }
    switch ($title) {
        'Collect snapshot' {
            if ($Group) {
                $members = $objs.Count + @($scenario.skipped).Count; $skipped = @($scenario.skipped).Count
                Ev @{ kind = 'target'; text = ('Group {0} (captured by Collect, snapshot 42): {1} member(s) including nested groups, {2} skipped' -f $groupName, $members, $skipped); name = $groupName; source = 'Snapshot'; members = $members; nested = 0; skipped = $skipped; count = $objs.Count }
            }
            Ev @{ kind = 'item'; status = 'Ok'; text = ('{0} object(s) selected' -f $objs.Count); identity = ''; name = '' }
        }
        'Objects' {
            $n = 0
            foreach ($o in $objs) {
                $n++; Ev @{ kind = 'progress'; phase = 'Objects'; done = $n - 1; total = $objs.Count; percent = [int](100 * ($n - 1) / [math]::Max(1, $objs.Count)); eta = '' }
                if ($o.blocked) { Ev @{ kind = 'item'; status = 'Fail'; text = [string]$o.blocked; identity = [string]$o.upn }; $errors++ }
                elseif ($o.warning) { Ev @{ kind = 'item'; status = 'Warn'; text = [string]$o.warning; identity = [string]$o.upn }; $warnings++ }
                else { Ev @{ kind = 'item'; status = 'Ok'; text = 'Ready'; identity = [string]$o.upn } }
            }
        }
        { $_ -in 'Users: cloud mailboxes', 'Users: back on-premises', 'Shared mailboxes' } {
            $mine = @($objs | Where-Object { if ($title -eq 'Shared mailboxes') { $_.shared } else { -not $_.shared } })
            if (-not $apply) { Ev @{ kind = 'item'; status = 'Skip'; text = 'Preview: not run.'; identity = '' }; break }
            $n = 0
            foreach ($o in $mine) {
                if (Stopped) { $left++; continue }
                $n++
                Ev @{ kind = 'progress'; phase = $(if ($title -eq 'Shared mailboxes') { 'Shared mailboxes' } else { 'Users' }); done = $n - 1; total = $mine.Count; percent = [int](100 * ($n - 1) / [math]::Max(1, $mine.Count)); eta = '' }
                Ev @{ kind = 'item'; status = 'Info'; text = $(if ($Action -eq 'Recover') { 'Waiting for the on-premises mailboxes of the users' } else { 'Waiting for the cloud mailboxes of the users' }); identity = [string]$o.upn }
                if ($n -eq $scenario.holdAfter) { Hold 'users' }
                if ($o.warning) { Ev @{ kind = 'item'; status = 'Warn'; text = ('{0}: {1}' -f $o.upn, $o.warning); identity = [string]$o.upn }; $warnings++ }
                else { Ev @{ kind = 'item'; status = 'Ok'; text = (Get-DoneText $o $title); identity = [string]$o.upn } }
                $done++
            }
        }
        { $_ -eq 'Shared mailboxes' -and $scenario.ask -and $apply -and $Action -eq 'Convert' } {
            # A question of the run (Entra Connect in Manual mode): asked in the window, answered in the file the event names.
            $answer = $env:PRA_EVENT_FILE + '.answer-1'
            Ev @{ kind = 'ask'; id = 1; title = 'Entra Connect'; text = 'Start a delta synchronisation on the Entra Connect server, then answer Yes.'; answer = $answer }
            $limit = [datetime]::UtcNow.AddSeconds(30)
            while (-not (Test-Path -LiteralPath $answer) -and [datetime]::UtcNow -lt $limit) { Start-Sleep -Milliseconds 100 }
            [IO.File]::WriteAllText((Join-Path $here 'asked-answer.txt'), $(if (Test-Path -LiteralPath $answer) { [IO.File]::ReadAllText($answer) } else { 'none' }))
            Ev @{ kind = 'item'; status = 'Ok'; text = 'Synchronisation confirmed'; identity = '' }
        }
        { $_ -in 'Users: source of authority and Exchange plan', 'Plan and confirmation' } {
            if (-not $apply -and $_ -ne 'Plan and confirmation') { Ev @{ kind = 'item'; status = 'Skip'; text = 'Preview: not run.'; identity = '' }; break }
            Ev @{ kind = 'item'; status = 'Ok'; text = 'Done'; identity = '' }
        }
        default {
            if ($apply -and $Action -eq 'Collect' -and $title -eq 'Saving the snapshot') { $done = $objs.Count }
            if (-not $apply -and $title -in 'Saving the snapshot', 'Licence group') { Ev @{ kind = 'item'; status = 'Skip'; text = 'Preview: not written.'; identity = '' }; break }
            Ev @{ kind = 'item'; status = 'Ok'; text = 'Done'; identity = '' }
        }
    }
    if ((Test-Path (Join-Path $here 'fail.flag')) -and $i -eq 1) {
        Ev @{ kind = 'item'; status = 'Fail'; text = 'Certificate 0A1B2C3D4E5F60718293A4B5C6D7E8F901234567 not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My.'; identity = '' }
        $errors++
        break
    }
    Ev @{ kind = 'step'; id = ('s{0}' -f ($i + 1)); state = $(if ($errors -gt 0 -and (Test-Path (Join-Path $here 'fail.flag')) -and $i -eq 1) { 'fail' } else { 'ok' }) }
    Start-Sleep -Milliseconds 40
}
if ($ResolveOnly) {
    Ev @{ kind = 'result'; status = 'Success'; exitCode = 0; action = $Action; mode = $Mode; success = 0; error = 0; pending = 0; skipped = 0; planned = 0; total = 0; seconds = 0; batchId = ''; nextSteps = @(); logFile = $logFile; csvReport = ''; htmlReport = ''; warnings = 0; issues = @() }
    exit 0
}
foreach ($o in $objs) {
    $status = if ($o.blocked) { 'Error' } elseif ($apply -and $Action -ne 'Check' -and -not (Stopped -and $left)) { 'Success' } elseif ($apply) { 'Success' } else { 'Planned' }
    $rows.Add([pscustomobject]@{ Identity = $o.upn; Kind = $(if ($o.shared) { 'Shared' } else { 'User' }); PrimarySmtpAddress = $o.upn; ObjectGuid = $o.guid; ExchangeGuid = ''; Action = $Action; Entra = ''; ExchangeOnline = ''; TeamsStorage = ''; Licence = ''; Holds = ''
            Permissions = ''; FinalStatus = $status; Detail = $(if ($o.blocked) { $o.blocked } else { '' }); Warnings = $(if ($o.warning) { $o.warning } else { '' }) })
}
foreach ($s in @($scenario.skipped)) { $rows.Add([pscustomobject]@{ Identity = $s.upn; Kind = $s.kind; PrimarySmtpAddress = ''; ObjectGuid = ''; ExchangeGuid = ''; Action = $Action; Entra = ''; ExchangeOnline = ''; TeamsStorage = ''; Licence = ''; Holds = ''; Permissions = ''; FinalStatus = 'Skipped'; Detail = 'not a mailbox'; Warnings = '' }); $skippedCount++ }
$rows | Export-Csv -LiteralPath $csvFile -Delimiter ';' -NoTypeInformation -Encoding UTF8
$failed = $errors -gt 0 -and (Test-Path (Join-Path $here 'fail.flag'))
$exit = if ($failed) { 1 } elseif ($left -gt 0) { 2 } elseif ($errors -gt 0 -and -not $apply) { 1 } else { 0 }
$status = switch ($exit) { 0 { if ($apply) { 'Success' } else { 'Planned' } } 1 { 'Failed' } 2 { 'Pending' } }
$next = @()
$script = '.\Invoke-PraCloudMailbox.ps1'
$again = ('{0} -Action {1}{2}{3}' -f $script, $Action, $(if ($apply) { ' -Mode Apply' } else { '' }), $(if ($Batch) { " -Batch $Batch" } else { '' }))
if ($failed) {
    if ($Action -eq 'Check') { $next = @('Check only reads: nothing was changed. Fix the problem named above, then run the same command again:', $again) }
    elseif (-not $apply) { $next = @('This was a Preview: nothing was changed. Fix the problem named above, then run the same command again:', $again) }
    else { $next = @('The run stopped before it wrote its journal: nothing was changed in the tenant. Fix the problem named above, then run the same command again:', $again) }
}
elseif ($left -gt 0) {
    $next = if ($Action -eq 'Convert') { @('Some objects or lists are not finished. Fix what is named above, then resume them (nothing is converted twice):') } else { @('Some objects are not back yet. Fix what is named above, then resume them (nothing is rolled back twice):') }
    $next += ('{0} -Action {1} -Mode Apply -Batch {2}' -f $script, $Action, $batchId)
}
elseif ($Action -eq 'Convert' -and $apply) { $next = @('When AD, Exchange and Entra Connect are rebuilt, roll this batch back (Preview first, then -Mode Apply):', ('{0} -Action Recover -Batch {1}' -f $script, $batchId)) }
elseif ($Action -eq 'Convert' -or $Action -eq 'Recover') { $next = @('To make the changes, run the same command in Apply mode (the confirmation is asked once):', ('{0} -Action {1} -Mode Apply{2}' -f $script, $Action, $(if ($Batch) { " -Batch $Batch" } else { '' }))) }
elseif ($Action -eq 'Collect' -and $apply) { $next = @(('Copy {0} to the cloud admin server (or let the backup folder be replicated), then, in the folder of the tool there:' -f (Join-Path $here 'data\PraCloudMailbox.db')), '.\Invoke-PraCloudMailbox.ps1 -Action Check') }
elseif ($Action -eq 'Schedule' -and $apply) { $next = @('Run the task once now, in this elevated window:', "Start-ScheduledTask -TaskName 'PraCloudMailbox Collect'", 'When it has finished, check the snapshot, then look at the light Daily Collect on the Setup page of the console:', '.\Invoke-PraCloudMailbox.ps1 -Action Check') }
elseif ($Action -eq 'Check') { $next = @('Disaster: run the Preview of Convert first (nothing is changed), then the same command with -Mode Apply:', ('{0} -Action Convert' -f $script)) }
Ev @{ kind = 'summary'; title = ('{0} finished' -f $Action); status = $(if ($failed) { 'Fail' } else { 'Ok' }); values = [ordered]@{ Result = @('done') } }
Ev @{ kind = 'result'; status = $status; exitCode = $exit; action = $Action; mode = $Mode; success = $(if ($apply) { $done } else { 0 }); error = $errors; pending = $left; skipped = $skippedCount; planned = $(if ($apply) { 0 } else { $objs.Count - $errors })
    total = $rows.Count; seconds = [int]((Get-Date) - $started).TotalSeconds; batchId = $(if ($apply -and -not $failed -and $Action -in 'Convert', 'Recover') { $batchId } else { '' }); nextSteps = $next; logFile = $logFile; csvReport = $csvFile; htmlReport = $htmlFile
    warnings = $warnings; issues = @(if ($failed) { 'Certificate 0A1B2C3D4E5F60718293A4B5C6D7E8F901234567 not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My.' }) }
exit $exit
