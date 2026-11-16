# What each page of the console shows, from its state: the Setup page (lights, role, run settings), the Task page (tiles, objects, batches, waves), the Verify page
# (preview, typed confirmation), the buttons and the navigation. Dot-sourced by PRA.Gui.psm1.

function Get-PcmGuiObjectsText {
    <# The objects of the current task in a few words: 'Group PRA-Wave1', 'One mailbox adele.vance@contoso.com', 'Batch 1a2b3c4d (whole batch)'. #>
    $g = $script:Gui
    if ($g.Action -eq 'Collect') { return 'Every mailbox of the scope, read from Exchange' }
    if (Get-PcmGuiBatchKind) {
        if (-not $g.SelectedBatch) { return 'No batch picked' }
        $id = [string]$g.SelectedBatch.BatchId
        if ($g.Action -eq 'Recover' -and $g.Mode -eq 'New') {
            if ((Get-PcmGuiField 'Wave') -eq 'Wave') { return 'Batch {0} · a wave of {1}' -f $id, (ConvertTo-PcmGuiPlural @(Get-PcmGuiTicked).Count 'object' 'objects') }
            return 'Batch {0} · the whole batch' -f $id
        }
        return 'Batch ' + $id
    }
    switch (Get-PcmGuiSourceKind) {
        'Single' { return 'One mailbox ' + (Get-PcmGuiField 'Identity') }
        'Ou' { return 'OU ' + (Get-PcmGuiField 'Ou') }
        'Group' { return 'Group ' + (Get-PcmGuiField 'Group') }
        'Csv' { return 'CSV file ' + (Get-PcmGuiField 'Csv') }
        default { return 'Everything in the snapshot' }
    }
}

function Get-PcmGuiSnapshotText {
    <# 'snapshot #42 (latest)', 'snapshot #40 (chosen)'. #>
    $g = $script:Gui
    if (-not $g.State.Snapshot) { return 'no snapshot' }
    return 'snapshot #{0} ({1})' -f $g.State.Snapshot.id, $(if (Test-PcmGuiOrdinal (Get-PcmGuiField 'Snap') 'Specific') { 'chosen' } else { 'latest' })
}

function Update-PcmGuiSettingsView {
    $g = $script:Gui; $c = $g.Controls
    $d = $g.Detect
    foreach ($key in $script:GuiSettingKeys) {
        if ($key -eq 'SnapId') { continue }
        $modified = (Test-PcmGuiModified $key) -or ($key -eq 'Snap' -and (Test-PcmGuiModified 'SnapId'))
        $used = Test-PcmGuiSettingUsed $key
        if ($c.ContainsKey('M_' + $key)) { $c['M_' + $key].Visibility = Get-PcmGuiVisibility ($modified -and $used) }
        if ($c.ContainsKey('Row_' + $key) -and $key -notin @('Env', 'Scope', 'ExpConfig')) { $c['Row_' + $key].Visibility = Get-PcmGuiVisibility $used }
    }
    $changed = @(Get-PcmGuiChangedSetting)
    $c.BtnResetAll.Visibility = Get-PcmGuiVisibility ($changed.Count -gt 0)
    $c.F_SnapId.Visibility = Get-PcmGuiVisibility (Test-PcmGuiSettingUsed 'SnapId')
    $latest = @($g.State.Snapshots | Where-Object { $_.status -eq 'Complete' }) | Select-Object -First 1
    $c.SnapLatestText.Text = if ($latest) { 'Latest: #{0}, taken {1}.' -f $latest.id, (Format-PcmGuiTime ([string]$latest.finished_utc)) } else { 'No complete snapshot found.' }
    $c.F_Env.IsEnabled = -not $g.EnvLock
    $c.EnvLock.Visibility = Get-PcmGuiVisibility ([bool]$g.EnvLock)
    if ($g.EnvLock) {
        $text = 'Locked to {0}: the chosen batch was created there, and a batch only continues in its own environment.' -f $g.EnvLock
        if ($g.State.Config -and -not (Test-PcmGuiOrdinal $g.EnvLock $g.Defaults.Env)) { $text += ' The configuration file says {0}.' -f $g.Defaults.Env }
        $c.EnvLockText.Text = $text
    }
    $environment = Get-PcmGuiField 'Env'
    $c.HdrEnv.Text = $(if ($environment) { $environment } else { 'not set' })
    $organization = Get-PcmGuiField 'Org'
    $snapshot = if ($g.State.Snapshot) { 'snapshot #' + $g.State.Snapshot.id } else { 'no snapshot yet' }
    $c.HdrEnvSub.Text = 'Tenant {0} · {1}' -f $(if ($organization) { $organization } else { 'not set' }), $snapshot
    $role = if ($d) { [string]$d.Role } else { 'None' }
    switch ($role) {
        'Both' { $c.RoleName.Text = 'Everything'; Set-PcmGuiPill $c.RolePill $c.RoleName 'Ok'; $c.RoleText.Text = 'This computer can run Collect, Check, Convert and Recover.' }
        'Cloud' { $c.RoleName.Text = 'Cloud actions'; Set-PcmGuiPill $c.RolePill $c.RoleName 'Ok'; $c.RoleText.Text = 'This computer can run Check, Convert and Recover. Collect needs Exchange.Server to be set and reachable: see the Exchange light below.' }
        'Collect' { $c.RoleName.Text = 'Collect'; Set-PcmGuiPill $c.RolePill $c.RoleName 'Ok'; $c.RoleText.Text = 'This computer can run Collect (a snapshot of the mailboxes). Check, Convert and Recover need the cloud modules, a sign-in and a way to Microsoft 365: see the lights below.' }
        'Checking' { $c.RoleName.Text = 'Checking...'; Set-PcmGuiPill $c.RolePill $c.RoleName 'Neutral'; $c.RoleText.Text = 'Checking that Microsoft 365 can be reached from this computer. This takes a few seconds.' }
        default {
            $c.RoleName.Text = 'Nothing yet'; Set-PcmGuiPill $c.RolePill $c.RoleName 'Fail'
            $c.RoleText.Text = $(if (Test-PcmGuiNetworkDown) { 'The cloud actions cannot run here until Microsoft 365 can be reached: see the red light below.' } else { 'Nothing can run here until the red lights below are fixed.' })
        }
    }
    $c.HdrHost.Text = [string]$env:COMPUTERNAME
    $c.HdrHostSub.Text = 'PowerShell {0} · signed in as {1}\{2}' -f $PSVersionTable.PSVersion, $env:USERDOMAIN, $env:USERNAME
    $scope = Get-PcmGuiField 'Scope'
    $parts = [Collections.Generic.List[string]]::new()
    $parts.Add($(if ($environment) { $environment } else { 'no environment' }))
    $parts.Add((Get-PcmGuiSnapshotText))
    $parts.Add($(if (Test-PcmGuiOrdinal (Get-PcmGuiField 'Signin') 'Interactive') { 'interactive sign-in' } else { 'certificate sign-in' }))
    if (-not (Test-PcmGuiOrdinal $scope 'All')) { $parts.Add('scope ' + $script:GuiScopeText[$scope].ToLowerInvariant()) }
    $c.RunLine.Text = 'This run: ' + ($parts -join ' · ')
    $c.RunLineSub.Text = if ($changed.Count) { '{0} changed for this run only: {1}.' -f (ConvertTo-PcmGuiPlural $changed.Count 'setting' 'settings'), ($changed -join ', ') } else { 'Everything comes from the configuration file.' }
    $c.SettingsError.Text = Get-PcmGuiSettingsProblem
    $c.ConfigError.Text = $(if ($g.State.ConfigError) { 'Configuration INVALID: ' + $g.State.ConfigError } else { '' })
}

function Get-PcmGuiCollectPill {
    <# The words on the Collect tile when this computer cannot run it: the first thing that is missing. #>
    $d = $script:Gui.Detect
    if ($d -and -not $d.CanCloudEngine) { return 'Needs PowerShell 7.5' }
    if ($d -and -not $d.ExchangeServer) { return 'Exchange.Server is not set' }
    return 'Exchange not reachable'
}

function Get-PcmGuiCloudPill {
    <# The words on the cloud tiles when this computer cannot run them: the first thing that is missing. #>
    $d = $script:Gui.Detect
    if ($d -and -not $d.CanCloudEngine) { return 'Needs PowerShell 7.5' }
    if (Test-PcmGuiNetworkDown) { return 'Microsoft 365 not reachable' }
    return 'Needs the cloud modules and a sign-in'
}

function Update-PcmGuiTileState {
    $g = $script:Gui; $c = $g.Controls
    $cloudOk = [bool]($g.Detect -and $g.Detect.CanCloudBase) -and -not (Test-PcmGuiNetworkDown)
    $states = @{ Collect = (Test-PcmGuiCanCollect); Convert = $cloudOk; Recover = $cloudOk; Check = $cloudOk }
    foreach ($name in $states.Keys) {
        $enabled = [bool]$states[$name]
        $c['T_' + $name].IsEnabled = $enabled
        $c['N_' + $name].Visibility = Get-PcmGuiVisibility (-not $enabled)
        $c['NT_' + $name].Text = if ($name -eq 'Collect') { Get-PcmGuiCollectPill } else { Get-PcmGuiCloudPill }
    }
    $g.Updating++
    try { foreach ($name in $states.Keys) { $c['T_' + $name].IsChecked = (Test-PcmGuiOrdinal $name $g.Action) } } finally { $g.Updating-- }
}

function Set-PcmGuiGroupInfo {
    param([string]$Text, [string]$Level = 'Info')
    $c = $script:Gui.Controls
    $c.GroupInfo.Text = $Text
    Set-PcmGuiLevelColor $c.GroupInfo $Level
    $c.GroupGlyph.Visibility = Get-PcmGuiVisibility ($Level -eq 'Ok' -and [bool]$Text)
    Set-PcmGuiLevelColor $c.GroupGlyph $Level
}

function Show-PcmGuiGroupResult {
    <# What the lookup found: the group and its members, or the plain reason. #>
    param([string]$Group, $Result)
    $error_ = [string](Get-PcmGuiValue $Result 'Error' '')
    if ($error_) { Set-PcmGuiGroupInfo ('Could not check: ' + $error_) 'Warn'; return }
    $count = [int](Get-PcmGuiValue $Result 'Count' 0)
    $source = [string](Get-PcmGuiValue $Result 'Source' 'Snapshot')
    if ($count -gt 1) {
        Set-PcmGuiGroupInfo ("{0} groups of the snapshot match '{1}': {2}. Type the full name or the distinguished name." -f $count, $Group, (@(Get-PcmGuiValue $Result 'Candidates' @()) -join ', ')) 'Warn'
        return
    }
    $fallback = [string](Get-PcmGuiValue $Result 'Fallback' '')
    if ($count -lt 1) {
        $old = if (-not [bool](Get-PcmGuiValue $Result 'HasTables' $true)) { ' (this snapshot was taken before Collect captured groups)' } else { '' }
        $next = if ($fallback) { "Microsoft Entra was not checked: $fallback" } else { "The tool looks for a Microsoft Entra group named '$Group' when the run starts." }
        Set-PcmGuiGroupInfo ("Not in the snapshot{0}. {1}" -f $old, $next) 'Info'
        return
    }
    $members = [int](Get-PcmGuiValue $Result 'Members' 0); $skipped = [int](Get-PcmGuiValue $Result 'Skipped' 0)
    $where = if (Test-PcmGuiOrdinal $source 'Entra') { 'Microsoft Entra group' } else { 'Found in the snapshot (captured by Collect)' }
    $text = '{0}: {1} · {2} including nested groups' -f $where, (Get-PcmGuiValue $Result 'Name' $Group), (ConvertTo-PcmGuiPlural $members 'member' 'members')
    if ($skipped) { $text += ', {0} skipped (the Preview says why)' -f $skipped }
    if ($fallback) { $text += ". Only the snapshot was checked: $fallback" }
    Set-PcmGuiGroupInfo $text 'Ok'
}

function Get-PcmGuiGroupLookupBlocker {
    <# Why the engine cannot look the group up from this computer (empty when it can): no snapshot, no PowerShell 7.5, no possible sign-in to Microsoft 365. #>
    $g = $script:Gui
    if (-not $g.State.Snapshot) { return 'there is no complete snapshot on this computer yet.' }
    $engine = Get-PcmGuiEngine
    if (-not $engine.Ok) { return $engine.Problem }
    if (-not (Test-PcmGuiCanCloud)) {
        $net = @($g.Detect.Lights | Where-Object { (Test-PcmGuiOrdinal $_.Name $script:GuiNetLight) -and (Test-PcmGuiOrdinal $_.Status 'Fail') })
        if ($net.Count) { return [string]$net[0].Detail }
        return 'this computer cannot sign in to Microsoft 365 yet (see Setup).'
    }
    $problem = Get-PcmGuiSettingsProblem
    if (-not $problem) { $problem = Get-PcmGuiCloudProblem }
    if (-not $problem) { $problem = Get-PcmGuiSnapshotBlocker }
    return $problem
}

function Start-PcmGuiGroupLookup {
    <#
        Looks the group up, read-only, with the engine itself (Check -Group <name> -ResolveOnly in a child process, like a run): in the snapshot where Collect captured it, else in Microsoft
        Entra ID, with the members a run would skip. When no sign-in to Microsoft 365 is possible from this computer, the snapshot alone answers and the text says so.
        A test or a picture passes -GroupProbe.
    #>
    $g = $script:Gui
    $group = Get-PcmGuiField 'Group'
    if (-not $group) { Set-PcmGuiGroupInfo ''; return }
    if ($g.GroupProbe) {
        try { Show-PcmGuiGroupResult $group (& $g.GroupProbe $group) } catch { Show-PcmGuiGroupResult $group ([pscustomobject]@{ Error = $_.Exception.Message }) }
        return
    }
    if ($g.GroupLookup -or $g.Run) { return }
    $reason = Get-PcmGuiGroupLookupBlocker
    if (-not $reason) {
        try {
            $request = New-PcmGuiRequest -Action Check -Group $group -Scope (Get-PcmGuiField 'Scope') -Snapshot (Get-PcmGuiSnapshotNumber) -ResolveOnly $true
            $config = New-PcmGuiRunConfig -Root $g.Root -ConfigPath $g.ConfigPath -Overrides (Get-PcmGuiOverride)
            $engine = Get-PcmGuiEngine
            $command = Get-PcmGuiCommand -Root $g.Root -ConfigPath $config.Path -Request $request -Engine $engine.Path -BaseConfigPath $config.Base
            if (Get-PcmGuiHook 'NoConsole') { $command.NeedsConsoleWindow = $false }
            $folder = Join-Path (Join-Path $g.Root 'logs') 'gui'
            Initialize-PcmGuiFolder (Join-Path $g.Root 'logs'); Initialize-PcmGuiFolder $folder
            $base = Join-Path $folder ([guid]::NewGuid().ToString('N'))
            $child = Start-PcmGuiChildProcess -Command $command -Root $g.Root -EventPath ($base + '.events.jsonl')
            $g.GroupLookup = @{ Group = $group; Process = $child.Process; Output = $child.Output; Errors = $child.Errors; EventPath = ($base + '.events.jsonl'); Position = [long]0
                Target = $null; Result = $null; Failure = ''; Clock = [Diagnostics.Stopwatch]::StartNew() }
            $c = $g.Controls
            $c.BtnCheckGroup.IsEnabled = $false
            Set-PcmGuiGroupInfo ('Looking up the group in the snapshot, then in Microsoft Entra if it is not there. A sign-in window can open.') 'Info'
            $g.Timer.Start()
            return
        } catch { $reason = $_.Exception.Message }
    }
    $answer = Get-PcmGuiGroupInSnapshot -State $g.State -Value $group
    $answer | Add-Member -NotePropertyName Fallback -NotePropertyValue $reason -Force
    Show-PcmGuiGroupResult $group $answer
}

function Complete-PcmGuiGroupLookup {
    <# Called by the window timer: reads the events of the group lookup and shows its answer when the engine has ended. #>
    $g = $script:Gui; $job = $g.GroupLookup
    if (-not $job) { return }
    $exited = $job.Process.HasExited
    try {
        $read = Read-PcmGuiEvents -Path $job.EventPath -Position $job.Position
        $job.Position = $read.Position
        foreach ($record in $read.Events) {
            switch ([string](Get-PcmGuiValue $record 'kind' '')) {
                'target' { $job.Target = $record }
                'result' { $job.Result = $record }
                'item' { if ((Test-PcmGuiOrdinal ([string](Get-PcmGuiValue $record 'status' '')) 'Fail') -and -not $job.Failure) { $job.Failure = [string](Get-PcmGuiValue $record 'text' '') } }
            }
        }
    } catch { Write-Verbose $_.Exception.Message }
    if (-not $exited -and $job.Clock.Elapsed.TotalMinutes -ge 10) {
        try { $job.Process.Kill() } catch { Write-Verbose $_.Exception.Message }
        $job.Failure = 'The group check took too long and was stopped.'
        return
    }
    if (-not ($exited -and $job.Output.IsCompleted -and $job.Errors.IsCompleted)) { return }
    $stderr = ''
    try { $null = $job.Output.Result; $stderr = [string]$job.Errors.Result } catch { Write-Verbose $_.Exception.Message }
    $code = -1
    try { $code = $job.Process.ExitCode } catch { Write-Verbose $_.Exception.Message }
    try { $job.Process.Dispose() } catch { Write-Verbose $_.Exception.Message }
    $g.GroupLookup = $null
    $g.Controls.BtnCheckGroup.IsEnabled = $true
    Sync-PcmGuiTimer
    if (-not (Test-PcmGuiOrdinal (Get-PcmGuiField 'Group') $job.Group)) { return }
    $target = $job.Target
    if ($code -eq 0 -and $target) {
        Show-PcmGuiGroupResult $job.Group ([pscustomobject]@{ HasTables = $true; Count = 1; Name = [string](Get-PcmGuiValue $target 'name' $job.Group); Source = [string](Get-PcmGuiValue $target 'source' 'Snapshot')
                Members = [int](Get-PcmGuiValue $target 'members' 0); Skipped = [int](Get-PcmGuiValue $target 'skipped' 0); Error = '' })
        return
    }
    $why = ''
    if ($job.Result) { foreach ($item in @(Get-PcmGuiValue $job.Result 'issues' @())) { if ([string]$item) { $why = [string]$item; break } } }
    if (-not $why) { $why = $job.Failure }
    if (-not $why -and $stderr.Trim()) { $why = ($stderr.Trim() -split "`r?`n")[0].Trim() }
    if (-not $why) { $why = 'The engine ended without an answer (exit code {0}).' -f $code }
    Show-PcmGuiGroupResult $job.Group ([pscustomobject]@{ Error = $why })
}

function Get-PcmGuiCsvRowCount {
    <# The number of identities of a CSV file or list (header excluded), or -1 when it cannot be read. #>
    param([string]$Path)
    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf) -or (Get-Item -LiteralPath $Path).Length -gt 8MB) { return -1 }
        $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 | Where-Object { $_.Trim() })
        if (-not $lines.Count) { return 0 }
        $first = ($lines[0] -split '[;,]')[0].Trim().Trim('"')
        if (Test-PcmGuiOrdinal $first 'Identity') { return $lines.Count - 1 }
        return $lines.Count
    } catch { return -1 }
}

function Get-PcmGuiWaveText {
    <# The waves of this run in a sentence: shared mailboxes by temporary licences (Convert), objects per wave (Recover). #>
    $g = $script:Gui
    $parallel = ConvertTo-PcmGuiWholeNumber (Get-PcmGuiField 'Parallel'); $keep = ConvertTo-PcmGuiWholeNumber (Get-PcmGuiField 'KeepFree')
    $sku = if ($g.State.Config) { [string](Get-PcmGuiValue $g.State.Config.Licensing.Shared 'SkuPartNumber' '') } else { '' }
    switch ($g.Action) {
        'Convert' {
            if ($null -eq $parallel) { return '' }
            return ' Shared mailboxes in waves of up to {0} temporary licences{1}{2}.' -f $parallel, $(if ($sku) { " ($sku)" } else { '' }), $(if ($keep) { ", $keep kept free" } else { '' })
        }
        'Recover' {
            $size = ConvertTo-PcmGuiWholeNumber (Get-PcmGuiField 'WaveSize')
            if ($null -eq $size) { $size = $parallel }
            if ($null -eq $size) { return '' }
            return ' Rollback in waves of up to {0} objects.' -f $size
        }
        default { return '' }
    }
}

function Update-PcmGuiTaskView {
    $g = $script:Gui; $c = $g.Controls
    $action = $g.Action
    Update-PcmGuiTileState
    $c.OnWhatEmpty.Visibility = Get-PcmGuiVisibility (-not $action)
    $c.OnWhatTitle.Text = if ($action) { 'On what? · ' + $script:GuiActionTitle[$action] } else { 'On what?' }
    $c.CollectPanel.Visibility = Get-PcmGuiVisibility ($action -eq 'Collect')
    $c.ModeBox.Visibility = Get-PcmGuiVisibility ($action -in @('Convert', 'Recover', 'Check'))
    $kind = Get-PcmGuiSourceKind
    $c.SourceBox.Visibility = Get-PcmGuiVisibility ([bool]$action -and $action -in @('Convert', 'Check') -and $g.Mode -eq 'New')
    $batchKind = Get-PcmGuiBatchKind
    $c.BatchBox.Visibility = Get-PcmGuiVisibility ([bool]$batchKind)
    $c.WaveBox.Visibility = Get-PcmGuiVisibility ($action -eq 'Recover' -and $g.Mode -eq 'New' -and [bool]$g.SelectedBatch)
    $c.CheckPanel.Visibility = Get-PcmGuiVisibility ($action -eq 'Check' -and $g.Mode -eq 'Continue' -and [bool]$g.SelectedBatch)
    switch ($action) {
        'Convert' { $c.Mode_New.Content = 'Start a new batch'; $c.Mode_Continue.Content = 'Continue a batch' }
        'Recover' { $c.Mode_New.Content = 'Start a rollback'; $c.Mode_Continue.Content = 'Continue an interrupted rollback' }
        'Check' { $c.Mode_New.Content = 'Check the snapshot'; $c.Mode_Continue.Content = 'Follow a batch' }
    }
    $g.Updating++
    try { $c.Mode_New.IsChecked = ($g.Mode -eq 'New'); $c.Mode_Continue.IsChecked = ($g.Mode -eq 'Continue') } finally { $g.Updating-- }
    $snapshotId = if ($g.State.Snapshot) { $g.State.Snapshot.id } else { '?' }
    if ($c.SourceBox.Visibility -eq 'Visible') {
        $c.Row_Identity.Visibility = Get-PcmGuiVisibility ($kind -eq 'Single'); $c.Row_All.Visibility = Get-PcmGuiVisibility ($kind -eq 'All')
        $c.Row_Ou.Visibility = Get-PcmGuiVisibility ($kind -eq 'Ou'); $c.Row_Group.Visibility = Get-PcmGuiVisibility ($kind -eq 'Group'); $c.Row_Csv.Visibility = Get-PcmGuiVisibility ($kind -eq 'Csv')
        $c.Row_All.Text = 'Every mailbox of snapshot #{0} ({1}), as limited by the scope of this run ({2}).' -f $snapshotId, (ConvertTo-PcmGuiPlural @($g.State.Objects).Count 'mailbox' 'mailboxes'), $script:GuiScopeText[(Get-PcmGuiField 'Scope')].ToLowerInvariant()
        $identity = Get-PcmGuiField 'Identity'
        if ($kind -eq 'Single' -and $identity) {
            $found = @(Find-PcmGuiMailbox -State $g.State -Text $identity)
            if ($found.Count -eq 1) { $c.IdentityInfo.Text = '{0} · {1} · {2}' -f $(if ($found[0].display_name) { $found[0].display_name } else { $found[0].name }), $found[0].primary_smtp_address, $(if ($found[0].kind -eq 'User') { 'user mailbox' } else { ([string]$found[0].kind).ToLowerInvariant() + ' mailbox' }); Set-PcmGuiLevelColor $c.IdentityInfo 'Ok' }
            elseif ($found.Count -gt 1) { $c.IdentityInfo.Text = "$($found.Count) mailboxes match: type the e-mail address."; Set-PcmGuiLevelColor $c.IdentityInfo 'Warn' }
            else { $c.IdentityInfo.Text = "Not found in snapshot #${snapshotId}. Check the spelling."; Set-PcmGuiLevelColor $c.IdentityInfo 'Fail' }
        } else { $c.IdentityInfo.Text = '' }
        $ou = Get-PcmGuiField 'Ou'
        if ($kind -eq 'Ou' -and $ou) {
            $matches_ = @(Get-PcmGuiOuMatch -State $g.State -Text $ou)
            $c.OuInfo.Text = if ($matches_.Count) { '{0} in snapshot #{1}, sub-OUs included.' -f (ConvertTo-PcmGuiPlural $matches_.Count 'mailbox' 'mailboxes'), $snapshotId } else { "No mailbox of snapshot #${snapshotId} is in this OU." }
            Set-PcmGuiLevelColor $c.OuInfo $(if ($matches_.Count) { 'Ok' } else { 'Fail' })
        } else { $c.OuInfo.Text = '' }
        $csv = Get-PcmGuiField 'Csv'
        if ($kind -eq 'Csv' -and $csv) {
            $rows = Get-PcmGuiCsvRowCount $csv
            $c.CsvInfo.Text = if ($rows -ge 0) { ConvertTo-PcmGuiPlural $rows 'identity' 'identities' } else { 'This file cannot be read.' }
            Set-PcmGuiLevelColor $c.CsvInfo $(if ($rows -ge 0) { 'Ok' } else { 'Fail' })
        } else { $c.CsvInfo.Text = '' }
    }
    $c.BatchHelp.Text = switch ($batchKind) {
        'ContinueConvert' { 'Pick the batch to continue. Only the objects not converted yet are processed.' }
        'Rollback' { 'Pick the conversion batch to roll back. Only the mailboxes still in the cloud are listed below.' }
        'ContinueRollback' { 'Pick the rollback to continue. Only the mailboxes still in the cloud are rolled back.' }
        'Follow' { 'Pick the batch to follow in Microsoft 365. Nothing is changed.' }
        default { '' }
    }
    if ($action -eq 'Collect') {
        $c.F_StoreView.Text = if ($g.State.Config) { if (Get-PcmGuiField 'Store') { Get-PcmGuiField 'Store' } else { [string]$g.State.Config.Store.Path } } else { '' }
        $c.CollectNote.Text = if ($g.State.Snapshot) { 'The last snapshot is #{0} ({1}). A scheduled task can do this by itself every day.' -f $g.State.Snapshot.id, (Format-PcmGuiTime ([string]$g.State.Snapshot.finished_utc)) } else { 'There is no snapshot yet.' }
        $c.BtnSchedule.IsEnabled = (Test-PcmGuiTaskReady) -and -not $g.Run
        $g.Updating++
        try { foreach ($choice in $script:GuiChoices.Scope) { $c['C_Scope_' + $choice].IsChecked = (Test-PcmGuiOrdinal $choice (Get-PcmGuiField 'Scope')) } } finally { $g.Updating-- }
    }
    if ($c.WaveBox.Visibility -eq 'Visible') {
        $wave = (Get-PcmGuiField 'Wave') -eq 'Wave'
        $c.WavePanel.Visibility = Get-PcmGuiVisibility $wave; $c.WaveTools.Visibility = Get-PcmGuiVisibility $wave
        Update-PcmGuiWaveCount
    }
    if ($c.CheckPanel.Visibility -eq 'Visible') { $c.ExpectBox.Visibility = 'Visible' }
    if ($action) {
        $happen = switch ($action) {
            'Collect' { @('Collect', 'This computer reads Exchange and writes a new snapshot file. Nothing is sent to the cloud and nothing is changed in Exchange. It takes a few minutes.') }
            'Check' {
                if ($g.Mode -eq 'Continue') { @('Cloud', $(if ($g.SelectedBatch) { 'This computer follows batch {0} in Microsoft 365 and {1}. It changes nothing.' -f $g.SelectedBatch.BatchId, $(if ((Get-PcmGuiField 'Once')) { 'looks once' } else { 'waits until every object is as expected' }) } else { 'Pick the batch to follow: this computer reads Microsoft 365 for each of its objects and changes nothing.' })) }
                else { @('Cloud', ('This computer reads Microsoft 365 for each object of snapshot #{0} and tells you what is ready. Nothing is changed.' -f $snapshotId)) }
            }
            'Convert' {
                if ($g.Mode -eq 'Continue') { @('Cloud', $(if ($g.SelectedBatch) { 'This computer continues batch {0} with the same snapshot: only the objects not converted yet are processed.' -f $g.SelectedBatch.BatchId } else { 'Pick the batch to continue: this computer takes the same snapshot again and processes only the objects not converted yet.' })) }
                else { @('Cloud', ('This computer reads snapshot #{0} ({1}), then converts the users and the shared mailboxes in Exchange Online. Nothing on-premises is read or written: Active Directory is not needed.' -f $snapshotId, $(if ($g.State.Snapshot) { Format-PcmGuiTime ([string]$g.State.Snapshot.finished_utc) } else { 'no date' }))) }
            }
            'Recover' {
                $what = if ($g.Mode -eq 'Continue') { 'the mailboxes of the interrupted rollback that are still in the cloud' } elseif ((Get-PcmGuiField 'Wave') -eq 'Wave') { ConvertTo-PcmGuiPlural @(Get-PcmGuiTicked).Count 'ticked mailbox' 'ticked mailboxes' } else { 'every mailbox of the batch that is still in the cloud' }
                @('Cloud', ('Use this only when Active Directory, Exchange and Entra Connect are rebuilt. This computer puts {0} back on-premises. Their cloud data stays under the hold. Preview first.' -f $what))
            }
        }
        $c.HappenPillText.Text = $happen[0]; Set-PcmGuiPill $c.HappenPill $c.HappenPillText 'Brand'
        $c.HappenText.Text = $happen[1] + (Get-PcmGuiWaveText)
        $blocker = Get-PcmGuiBlocker
        $c.BlockBox.Visibility = Get-PcmGuiVisibility ([bool]$blocker)
        $c.BlockText.Text = $blocker
    } else {
        $c.HappenPillText.Text = 'Task'; Set-PcmGuiPill $c.HappenPill $c.HappenPillText 'Neutral'
        $c.HappenText.Text = 'Pick a task above to see what this computer will do.'
        $c.BlockBox.Visibility = 'Collapsed'
    }
}

function Test-PcmGuiConfirmed {
    $g = $script:Gui
    $word = $script:GuiActionWord[$g.Action]
    if (-not $word) { return $true }
    return Test-PcmGuiConfirmation -Typed ([string]$g.Controls.ConfirmBox.Text) -Word $word
}

function Add-PcmGuiChip {
    param([string]$Text, [string]$Level)
    $border = New-Object Windows.Controls.Border
    $border.CornerRadius = [Windows.CornerRadius]::new(10); $border.Padding = [Windows.Thickness]::new(10, 2, 10, 2); $border.Margin = [Windows.Thickness]::new(0, 0, 8, 4)
    $label = New-Object Windows.Controls.TextBlock
    $label.Text = $Text; $label.FontSize = 12; $label.FontWeight = [Windows.FontWeights]::SemiBold; $border.Child = $label
    Set-PcmGuiPill $border $label $Level
    [void]$script:Gui.Controls.Chips.Children.Add($border)
}

function Get-PcmGuiVerb {
    switch ($script:Gui.Action) { 'Convert' { return 'converted' } 'Recover' { return 'rolled back' } 'Collect' { return 'collected' } default { return 'checked' } }
}

function Update-PcmGuiVerifyView {
    $g = $script:Gui; $c = $g.Controls
    $action = $g.Action
    $word = $script:GuiActionWord[$action]
    $preview = $g.Run -and (Test-PcmGuiOrdinal $g.Run.Purpose 'Preview')
    $c.ConfirmCard.Visibility = Get-PcmGuiVisibility ([bool]$word)
    $c.ConfirmBox.IsEnabled = ($null -ne $g.Pending) -and -not $g.Run
    $c.ConfirmText.Text = if ($word) { 'Type {0} to confirm (any letter case). This avoids a run started by a stray click.' -f $word } else { '' }
    $c.ConfirmTick.Visibility = $(if ($g.Pending -and $word -and (Test-PcmGuiConfirmed)) { 'Visible' } else { 'Hidden' })
    $scopeText = $script:GuiScopeText[(Get-PcmGuiField 'Scope')]
    $signin = if ($action -eq 'Collect') { '' } else { "`nSign-in: " + $(if (Test-PcmGuiOrdinal (Get-PcmGuiField 'Signin') 'Interactive') { 'Interactive' } else { 'Certificate' }) }
    $snapshotLine = if ($action -eq 'Collect') { '' } else { "`nSnapshot: " + $(if ($g.State.Snapshot) { '#{0} ({1})' -f $g.State.Snapshot.id, (Format-PcmGuiTime ([string]$g.State.Snapshot.finished_utc)) } else { 'none' }) }
    $waves = switch ($action) { 'Convert' { "`nShared mailbox waves: $(Get-PcmGuiField 'Parallel') temporary licences, $(Get-PcmGuiField 'KeepFree') kept free" } 'Recover' { "`nRollback waves: $(if (Get-PcmGuiField 'WaveSize') { Get-PcmGuiField 'WaveSize' } else { Get-PcmGuiField 'Parallel' }) objects" } default { '' } }
    $c.UsedText.Text = "Environment: $(Get-PcmGuiField 'Env')$snapshotLine`nObjects: $(Get-PcmGuiObjectsText)`nScope: $scopeText$signin$waves"
    $changed = @(Get-PcmGuiChangedSetting)
    if ($changed.Count) { $c.UsedText.Text += "`n`n" + (ConvertTo-PcmGuiPlural $changed.Count 'setting differs' 'settings differ') + ' from the configuration file, for this run only.' }
    $plan = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
    if (-not $g.Model -or -not $g.Run) {
        $titles = @(Get-PcmGuiPlanTitle -Action $action -Follow:($action -eq 'Check' -and $g.Mode -eq 'Continue'))
        for ($i = 0; $i -lt $titles.Count; $i++) { $plan.Add((New-PcmGuiRow @{ N = [string]($i + 1); Title = (Get-PcmGuiStepLabel $titles[$i]); What = (Get-PcmGuiStepInfo $titles[$i]) })) }
        $c.PlanList.ItemsSource = $plan
    }
    $c.Chips.Children.Clear()
    $c.TargetText.Visibility = Get-PcmGuiVisibility ([bool]$c.TargetText.Text -and $g.Previewed)
    $showList = $g.Previewed -and $g.VerRows.Count -gt 0
    $c.VerList.Visibility = Get-PcmGuiVisibility $showList
    $c.VerEmpty.Visibility = Get-PcmGuiVisibility (-not $showList)
    if ($preview) {
        $c.SumTitle.Text = 'Preview in progress...'; $c.SumText.Text = 'The engine reads the snapshot and Microsoft 365. The preview only reads: nothing is modified.'
        $c.SumGlyph.Text = [string][char]0xE895; Set-PcmGuiPill $c.SumIcon $c.SumGlyph 'Brand'
        $c.VerBar.Visibility = 'Visible'; $c.VerEmptyText.Text = if ($g.Model -and $g.Model.Current) { [string]$g.Model.Current } else { 'Starting...' }
    }
    elseif ($g.PreviewError) {
        $c.SumTitle.Text = 'The preview did not complete'; $c.SumText.Text = $g.PreviewError
        $c.SumGlyph.Text = [string][char]0xEA39; Set-PcmGuiPill $c.SumIcon $c.SumGlyph 'Fail'
        $c.VerBar.Visibility = 'Collapsed'; $c.VerEmptyText.Text = 'Nothing was changed. Read the message above, fix the cause, then preview again.'
    }
    elseif ($action -eq 'Check') {
        $c.SumTitle.Text = 'The check reads Microsoft 365. Nothing is changed.'; $c.SumText.Text = 'There is no preview for a check: it never writes. Click the button below to run it.'
        $c.SumGlyph.Text = [string][char]0xE721; Set-PcmGuiPill $c.SumIcon $c.SumGlyph 'Neutral'
        $c.VerBar.Visibility = 'Collapsed'; $c.VerEmptyText.Text = 'Run the check from the button at the bottom right.'
    }
    elseif (-not $g.Previewed) {
        $c.SumTitle.Text = 'Preview needed before running'
        $c.SumText.Text = '{0}. Click Preview: it only reads, nothing is changed. Run unlocks after the preview and the confirmation word.' -f (Get-PcmGuiObjectsText)
        $c.SumGlyph.Text = [string][char]0xE721; Set-PcmGuiPill $c.SumIcon $c.SumGlyph 'Neutral'
        $c.VerBar.Visibility = 'Collapsed'; $c.VerEmptyText.Text = 'Click Preview to see exactly what would change. The preview only reads: nothing is modified.'
    }
    else {
        $result = $g.PreviewResult
        $rows = @($g.VerRows)
        $shared = @($rows | Where-Object { Test-PcmGuiOrdinal $_.Type 'Shared' }).Count
        $users = @($rows | Where-Object { Test-PcmGuiOrdinal $_.Type 'User' }).Count
        $skipped = if ($result -and $null -ne $result.PSObject.Properties['skipped']) { [int]$result.skipped } else { @($rows | Where-Object { Test-PcmGuiOrdinal $_.Status 'Skip' }).Count }
        $warn = @($rows | Where-Object { Test-PcmGuiOrdinal $_.Status 'Warn' }).Count
        if ($result -and $null -ne $result.PSObject.Properties['warnings']) { $warn = [Math]::Max($warn, [int]$result.warnings) }
        $blocking = @($rows | Where-Object { Test-PcmGuiOrdinal $_.Status 'Fail' }).Count
        $planned = if ($result -and $null -ne $result.PSObject.Properties['planned'] -and [int]$result.planned) { [int]$result.planned } else { @($rows | Where-Object { Test-PcmGuiOrdinalIn $_.Status @('Ok', 'Warn') }).Count }
        $c.VerEmptyText.Text = 'The preview found no object to process.'; $c.VerBar.Visibility = 'Collapsed'
        if ($action -eq 'Collect') { $c.SumTitle.Text = 'A new snapshot will be written · nothing blocking' }
        else { $c.SumTitle.Text = '{0} will be {1} · {2}' -f (ConvertTo-PcmGuiPlural $planned 'mailbox' 'mailboxes'), (Get-PcmGuiVerb), $(if ($blocking) { "$blocking blocking" } else { 'nothing blocking' }) }
        $c.SumText.Text = if ($blocking) { 'Objects marked in red are not ready: the run leaves them as they are. Read them below.' } elseif ($warn) { 'Nothing in the way. ' + (ConvertTo-PcmGuiPlural $warn 'warning' 'warnings') + ' to read below: you can still go on.' } else { 'Nothing in the way.' }
        if ($word) { $c.SumText.Text += " Type $word below (capitals or not), then run." }
        $c.SumGlyph.Text = $(if ($blocking) { [string][char]0xEA39 } else { [string][char]0xE73E })
        Set-PcmGuiPill $c.SumIcon $c.SumGlyph $(if ($blocking) { 'Fail' } else { 'Ok' })
        if ($action -ne 'Collect') {
            Add-PcmGuiChip "Users $users" 'Neutral'
            if ($shared) { Add-PcmGuiChip "Shared $shared" 'Neutral' }
            if ($skipped) { Add-PcmGuiChip "Skipped $skipped" 'Neutral' }
            Add-PcmGuiChip "$blocking blocking" $(if ($blocking) { 'Fail' } else { 'Ok' })
            if ($warn) { Add-PcmGuiChip (ConvertTo-PcmGuiPlural $warn 'warning' 'warnings') 'Warn' }
        }
    }
    $c.GridTitle.Text = if ($g.VerRows.Count) { 'Objects ({0})' -f $g.VerRows.Count } else { 'Objects' }
}

function Get-PcmGuiRunLabel {
    switch ($script:Gui.Action) {
        'Convert' { return 'Convert the mailboxes' }
        'Recover' { return 'Roll back the mailboxes' }
        'Collect' { return 'Collect the snapshot' }
        'Check' { return 'Run the check' }
        default { return 'Run' }
    }
}

function Get-PcmGuiModeText {
    $g = $script:Gui
    switch ($g.Action) {
        'Collect' { return '' }
        'Check' { if ($g.Mode -eq 'Continue') { return 'follow a batch' } return 'read only' }
        'Recover' { if ($g.Mode -eq 'Continue') { return 'continue a rollback' } return 'rollback' }
        default { if ($g.Mode -eq 'Continue') { return 'continue a batch' } return 'new batch' }
    }
}

function Update-PcmGuiButtons {
    $g = $script:Gui; $c = $g.Controls
    $arrow = [string][char]0x2192
    $busy = [bool]$g.Run
    $c.BtnBack.Visibility = Get-PcmGuiVisibility ($g.Page -gt 1)
    $c.BtnBack.IsEnabled = (-not $busy) -and -not ((Test-PcmGuiApplied) -and $g.Page -le 4)
    $c.BtnPreview.Visibility = Get-PcmGuiVisibility ($g.Page -eq 3 -and $g.Action -ne 'Check')
    $stoppable = $busy -and $g.Run.Request.Mode -eq 'Apply' -and $g.Run.Request.Action -in @('Convert', 'Recover')
    $c.BtnStop.Visibility = Get-PcmGuiVisibility ($g.Page -eq 4 -and $stoppable)
    $c.BtnStop.IsEnabled = $busy -and -not $g.Run.StopAsked
    $next = $c.BtnNext
    $next.IsEnabled = $true
    switch ($g.Page) {
        1 { $next.Content = "Choose a task  $arrow"; $next.IsEnabled = (-not $g.State.ConfigError) -and -not (Get-PcmGuiSettingsProblem) -and -not $busy -and -not $g.NetJob }
        2 { $next.Content = "Continue to verification  $arrow"; $next.IsEnabled = (Test-PcmGuiTaskReady) -and -not $busy }
        3 {
            $next.Content = Get-PcmGuiRunLabel
            $ready = (Test-PcmGuiTaskReady) -and -not $busy
            $next.IsEnabled = if ($g.Action -eq 'Check') { $ready } else { $ready -and ($null -ne $g.Pending) -and (Test-PcmGuiConfirmed) }
            $c.BtnPreview.Content = $(if ($g.Previewed) { 'Preview again' } else { 'Preview' })
            $c.BtnPreview.IsEnabled = $ready
            if ($g.Previewed) { $c.BtnPreview.ClearValue([Windows.Controls.Button]::StyleProperty) } else { $c.BtnPreview.Style = $g.Form.FindResource('AccentBtn') }
        }
        4 {
            if ($g.Run) { $next.Content = 'Running...'; $next.IsEnabled = $false }
            elseif ($g.LastRun) { $next.Content = "What next?  $arrow" }
            else { $next.Content = 'Nothing running'; $next.IsEnabled = $false }
        }
        5 { $next.Content = "Start a new task  $arrow" }
    }
    $c.Footer.Text = switch ($g.Page) {
        1 { if ($g.NetJob) { 'Checking that Microsoft 365 can be reached from this computer (a few seconds)...' } else { 'Check the lights and the run settings. Changes here are for this run only.' } }
        2 { 'Pick what you want to do. The console works out the rest.' }
        3 { if ($g.Action -eq 'Check') { 'A check only reads. Nothing is written.' } else { 'The preview comes first. Run unlocks after the preview and the confirmation word.' } }
        4 { if ($g.Run) { 'Do not close this window. Every change is journaled: if it stops, continue the batch.' } else { 'The run is over. Click What next to see what to do now.' } }
        5 { if ($g.LastRun -and $g.LastRun.Request.Action -eq 'Convert') { 'Keep the batch ID and the journal file: they are what lets you roll this batch back later.' } else { 'Everything the run did is in the report and the log.' } }
    }
    $c.ContextText.Text = if ($g.Action) {
        $mode = Get-PcmGuiModeText
        '{0}{1} · {2}' -f $script:GuiActionTitle[$g.Action], $(if ($mode) { ' · ' + $mode } else { '' }), (Get-PcmGuiField 'Env')
    } else { 'No task chosen yet · ' + (Get-PcmGuiField 'Env') }
}

function Update-PcmGuiView {
    <# One refresh of the whole window from its state; cheap (no file or network access). #>
    $g = $script:Gui
    if (-not $g -or $g.Updating -gt 0) { return }
    $g.Updating++
    try {
        Update-PcmGuiSettingsView
        switch ($g.Page) { 2 { Update-PcmGuiTaskView } 3 { Update-PcmGuiVerifyView } }
        Update-PcmGuiButtons
    } finally { $g.Updating-- }
}

function Reset-PcmGuiPreview {
    <# Any change of the task or of the settings voids the preview and the confirmation. #>
    $g = $script:Gui
    $g.Pending = $null; $g.Previewed = $false; $g.PreviewError = ''; $g.PreviewResult = $null
    $g.VerRows.Clear()
    $g.Controls.TargetText.Text = ''
    $g.Updating++
    try { $g.Controls.ConfirmBox.Clear() } finally { $g.Updating-- }
}

function Invoke-PcmGuiChange {
    <# Called by every control that changes the task or the run settings. #>
    $g = $script:Gui
    if (-not $g -or $g.Updating -gt 0 -or $g.Run) { return }
    Reset-PcmGuiPreview
    if ($g.Page -le 2) { Invoke-PcmGuiDetection }
    Update-PcmGuiView
}

function Show-PcmGuiPage {
    param([int]$Number)
    $g = $script:Gui; $c = $g.Controls
    $g.Page = $Number
    foreach ($index in 1..5) { $c["Page$index"].Visibility = Get-PcmGuiVisibility ($index -eq $Number) }
    $g.Updating++
    try { $c["Nav$Number"].IsChecked = $true } finally { $g.Updating-- }
    if ($Number -eq 2) { Update-PcmGuiBatchList }
    Update-PcmGuiView
}

function Test-PcmGuiApplied {
    <# True once an Apply has run (successfully or not) and the operator has not started a new task: Verify cannot be reached again. #>
    $g = $script:Gui
    return ($null -ne $g.LastRun) -and (Test-PcmGuiOrdinal $g.LastRun.Purpose 'Run') -and (Test-PcmGuiOrdinal $g.LastRun.Request.Mode 'Apply')
}

function Test-PcmGuiPageReachable {
    param([int]$Number)
    $g = $script:Gui
    if ($g.Run) { return $Number -eq 4 }
    if (Test-PcmGuiApplied) { return $Number -ge 4 -and $null -ne $g.LastRun }
    switch ($Number) {
        1 { return $true }
        2 { return (-not $g.State.ConfigError) -and -not $g.NetJob }
        3 { return [bool](Test-PcmGuiTaskReady) }
        4 { return $null -ne $g.Model }
        5 { return $null -ne $g.LastRun }
        default { return $false }
    }
}

# SIG # Begin signature block
# MIImKwYJKoZIhvcNAQcCoIImHDCCJhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCUt44ovwlLfa03
# 0FFcWttTuvJRuHMl0e4FsHweqpl/kKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYEjCCGA4CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCUxsh/r9lS
# qkTd4vVB9KFJBuBgrIb0lyRbvFqqWkhY5DANBgkqhkiG9w0BAQEFAASCAgCIo4EI
# RRmjYQGUSCooCI06lVWKCzStxVuqOKwXkRopp2yso3BM072k0AVT7QMiyh6eQG4z
# wyAjwMFCWK0lkjda8+RudYZ+BiKblVmlUyAb3hyBl/Z3ZP7xYrT2hXT36yC9PcvV
# daeSqh/GN7a4d7q71SDt4eVvqxsmtSu0WcbINDvEm7dfCG8gDZIgQDL2zsiad6e7
# s+qTMhhxBBhUqUB+DocoLa4C0Cme2jQvXGFk1QHlzcW6RJzPpHWB6aFD9c7iVI+M
# fh7aQhHih/TWKgzdE5A7m4fQXhQ+8Tz7PuaU8itYcaxFAdqnpK6gJTQczPdyuRZ3
# 3zWZCR2C8ECOqbVMYdBu7Vq4AnLZJYStPbr/J62jfIDDSjNq0gP1OJYrYAXjLYPO
# atorck/zJ71mc0uNaIPQH+9MeJhiFwoO573UQUNdez9/VPgsm3Y3P8HDzMXYE4Sk
# E/BkrXYdNWRTPRe+1auShaCVK3iOdY4lbWg/woBPXVydadbS+ZduNon0Ydt+ai0d
# hCU/a4Gd7p+Q9I553kBConz1zuempDix4mdW1tJkvGqB2U6BMt7lDO9Ll/86uJId
# xV3qI2GFEQSOV//8nbRe43fz1QtNxeM/JTUfKFz3H81qPgBRNxmSbsvKP4VImfBu
# QSKoXsp/HSSNPqzVKKP0L1VBcWakiUKAJjnEm6GCFM8wghTLBgorBgEEAYI3AwMB
# MYIUuzCCFLcGCSqGSIb3DQEHAqCCFKgwghSkAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIHwAj5m+Z9+SaODCXiwFK0voZelkeKP2pldBLxmO87w1AghM8HHczHH4
# PxgPMjAyNjExMTUyMDA2MzdaMAMCAQECEFVMy6NNBweSyot/9SOzk6KgghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlcwggJTAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTIwMDYzN1owKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIFA4
# RXNYBAz4qUsQYzQyb5FapAvewTonrZMTsunLO8MGMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIESDBGAiEAuPa0TYVuZ/bIxmCBPjW2bFXzFo/XZLECG+PBr8OwjWACIQCa
# u2qH3YdrIhLmL9pwtnISwDgHLpsWh1HYCiuYY93Rew==
# SIG # End signature block
