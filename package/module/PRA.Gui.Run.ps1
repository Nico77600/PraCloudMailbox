# The Run page and the Next steps page: the steps of the engine with real progress (from its event stream), the objects of each step, the object card, the child process, the
# stop file, the result of the run and what to do next. Dot-sourced by PRA.Gui.psm1.

#region Steps in plain words
# The engine announces each step by its title; the console shows it in plain words. A title that is not in this list is shown as the engine wrote it.
$script:GuiStepLabel = New-Object 'Collections.Generic.Dictionary[string,string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
$script:GuiStepText = New-Object 'Collections.Generic.Dictionary[string,string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
$script:GuiStepDefinition = @(
    ,@('Collect snapshot', 'Read the snapshot', 'Opens the snapshot and picks the objects to work on.')
    ,@('Batch journal', 'Read the batch', 'Reads the batch in the journal: its objects and what each run did.')
    ,@('Follow-up of the batch', 'Follow the batch', 'Checks in Microsoft 365 that each object of the batch is as expected, and waits for it. Nothing is changed.')
    ,@('Exchange on-premises connection', 'Connect to Exchange', 'Opens the connection to Exchange (read only).')
    ,@('Mailboxes in scope', 'Read the mailboxes', 'Reads every mailbox of the scope: addresses, type, organizational unit.')
    ,@('Shared mailbox permissions', 'Read the permissions of the shared mailboxes', 'Reads who can open, send as and send on behalf of each shared mailbox.')
    ,@('Contacts and distribution groups', 'Read contacts and distribution groups', 'Reads the contacts and the distribution groups that the configuration asks for.')
    ,@('Target groups', 'Read the target groups', 'Reads the groups that Convert and Recover can later pick their objects from, with their members.')
    ,@('Saving the snapshot', 'Save the snapshot', 'Writes the snapshot to the database and keeps a consistent copy.')
    ,@('Microsoft 365 connection', 'Connect to Microsoft 365', 'Signs in to Microsoft Graph and Exchange Online and checks the roles of the sign-in.')
    ,@('Tenant prerequisites', 'Check the tenant', 'Checks the licences, the groups and the settings of the tenant that the run needs.')
    ,@('Prerequisites', 'Check the prerequisites', 'Checks that the tenant is ready for this run.')
    ,@('Objects', 'Check each object', 'Checks each object in Microsoft 365: licence, address, cloud mailbox, holds.')
    ,@('Readiness', 'Check that every object is ready', 'Checks each object in Microsoft 365: licence, address, cloud mailbox, holds.')
    ,@('Plan and confirmation', 'Plan the changes', 'Works out what will change, object by object.')
    ,@('Users: source of authority and Exchange plan', 'Give the users a cloud identity and an Exchange Online plan', 'Moves the source of authority of each user to the cloud and gives the Exchange Online plan.')
    ,@('Users: cloud mailboxes', 'Convert the users', 'Waits for each cloud mailbox. One object at a time: you can stop after the current one.')
    ,@('Shared mailboxes', 'Convert the shared mailboxes', 'Converts the shared mailboxes and puts back their permissions.')
    ,@('Users: back on-premises', 'Put the users back on-premises', 'Gives each user back to Active Directory and their on-premises mailbox; the cloud data stays under the hold.')
    ,@('Licence group', 'Update the licence group', 'Takes the rolled-back users out of the licence group.')
    ,@('Configuration', 'Check the configuration and the window', 'Checks Exchange.Server and that this window is elevated: a task can only be registered by an administrator.')
    ,@('Account', 'Check the account', 'Checks that the account is a domain user. SYSTEM, computer accounts and managed service accounts cannot run Collect.')
    ,@('Exchange access', 'Prove the access to Exchange', 'Opens a remote session to Exchange with the account and reads one recipient, then closes it.')
    ,@('Batch logon right', 'Check the right to log on as a batch job', 'On a domain controller only: checks that the account may log on as a batch job.')
    ,@('Scheduled task', 'Register the task', 'Registers the daily Collect in the Windows Task Scheduler. A task of the same name is replaced. The password is not written anywhere.')
)
foreach ($definition in $script:GuiStepDefinition) { $script:GuiStepLabel[$definition[0]] = $definition[1]; $script:GuiStepText[$definition[0]] = $definition[2] }

function Get-PcmGuiStepLabel {
    param([string]$Title)
    if ($script:GuiStepLabel.ContainsKey($Title)) { return $script:GuiStepLabel[$Title] }
    return $Title
}

function Get-PcmGuiStepInfo {
    param([string]$Title)
    if ($script:GuiStepText.ContainsKey($Title)) { return $script:GuiStepText[$Title] }
    return ''
}

function Get-PcmGuiPlanTitle {
    <# The steps the engine will announce for a run (engine titles, in its order); the run itself replaces them with the steps of its plan event. #>
    param([string]$Action, [switch]$Follow, [switch]$ResolveOnly)
    switch ($Action) {
        'Collect' { return @('Exchange on-premises connection', 'Mailboxes in scope', 'Shared mailbox permissions', 'Contacts and distribution groups', 'Target groups', 'Saving the snapshot') }
        'Schedule' { return @('Configuration', 'Account', 'Exchange access', 'Batch logon right', 'Scheduled task') }
        'Check' {
            if ($ResolveOnly) { return @('Collect snapshot') }
            if ($Follow) { return @('Batch journal', 'Collect snapshot', 'Microsoft 365 connection', 'Follow-up of the batch') }
            return @('Collect snapshot', 'Microsoft 365 connection', 'Tenant prerequisites', 'Objects')
        }
        'Convert' { return @('Collect snapshot', 'Microsoft 365 connection', 'Readiness', 'Plan and confirmation', 'Users: source of authority and Exchange plan', 'Users: cloud mailboxes', 'Shared mailboxes') }
        'Recover' { return @('Collect snapshot', 'Microsoft 365 connection', 'Prerequisites', 'Plan and confirmation', 'Users: back on-premises', 'Shared mailboxes', 'Licence group') }
        default { return @() }
    }
}
#endregion

#region Run model: steps, objects, messages
$script:GuiLogLines = 2000

function Set-PcmGuiLogHeader {
    <# The activity log only shows the last lines of a run: its header says where the complete log is. #>
    param([string]$Path)
    $g = $script:Gui
    if (-not $g -or -not $g.Controls -or -not $g.Controls.ContainsKey('LogExpander')) { return }
    if (-not $Path) { $g.Controls.LogExpander.Header = 'Activity log'; return }
    $header = New-Object Windows.Controls.TextBlock
    $header.Text = 'Activity log: the last {0} lines. The complete log is in {1}' -f $script:GuiLogLines, $Path
    $header.TextWrapping = 'Wrap'
    $g.Controls.LogExpander.Header = $header
}

function Add-PcmGuiLine {
    param([string]$Text, [string]$Status = 'Info')
    $g = $script:Gui
    if (-not $g -or -not $g.Form) { return }
    $key = switch ($Status) { 'Fail' { 'PraCritical' } 'Warn' { 'PraCaution' } 'Ok' { 'PraSuccess' } default { 'TextFillColorSecondaryBrush' } }
    $g.Lines.Add([pscustomobject]@{ Text = $Text; Brush = $g.Form.TryFindResource($key) })
    while ($g.Lines.Count -gt $script:GuiLogLines) { $g.Lines.RemoveAt(0) }
    if ($g.Controls.LogExpander.IsExpanded -and $g.Lines.Count) { $g.Controls.ActivityLog.ScrollIntoView($g.Lines[$g.Lines.Count - 1]) }
}

function Get-PcmGuiModelObject {
    <# The row of an object in the list of a step (the name is the identity the engine gave, letter case ignored), or $null. One dictionary per step: the lookup does not read the list. #>
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$StepId, [Parameter(Mandatory)][AllowEmptyString()][string]$Identity)
    if (-not $Model.ObjectIndex.ContainsKey($StepId)) { return $null }
    $row = $null
    if ($Model.ObjectIndex[$StepId].TryGetValue($Identity, [ref]$row)) { return $row }
    return $null
}

function Add-PcmGuiModelObject {
    <# Adds the row of an object to the list of a step and to the dictionary that finds it again. #>
    param([Parameter(Mandatory)]$Model, [Parameter(Mandatory)][string]$StepId, [Parameter(Mandatory)]$Row)
    if (-not $Model.Objects.ContainsKey($StepId)) { $Model.Objects[$StepId] = New-Object 'Collections.ObjectModel.ObservableCollection[object]' }
    if (-not $Model.ObjectIndex.ContainsKey($StepId)) { $Model.ObjectIndex[$StepId] = New-Object 'Collections.Generic.Dictionary[string,object]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase) }
    $Model.Objects[$StepId].Add($Row)
    $Model.ObjectIndex[$StepId][[string]$Row.Name] = $Row
}

function ConvertTo-PcmGuiStatus {
    <# An engine state or item status as the status of a row (it drives the glyph and the colours). #>
    param([string]$Value)
    $map = @{ running = 'Running'; ok = 'Ok'; success = 'Ok'; warn = 'Warn'; fail = 'Fail'; skipped = 'Skip'; skip = 'Skip'; info = 'Info'; sub = 'Info'; pending = 'Pending'; wait = 'Wait' }
    $key = $Value.ToLowerInvariant()
    if ($map.ContainsKey($key)) { return [string]$map[$key] }
    return 'Info'
}

function ConvertTo-PcmGuiUtc {
    param($Value)
    $parsed = [datetime]::MinValue
    if ($Value -and [datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) { return $parsed.ToUniversalTime() }
    return [datetime]::UtcNow
}

function New-PcmGuiStepRow {
    param([string]$Id, [int]$Index, [string]$EngineTitle, [int]$Expected = 0)
    $label = Get-PcmGuiStepLabel $EngineTitle
    return New-PcmGuiRow @{
        Id = $Id; Index = $Index; EngineTitle = $EngineTitle; Label = $label; Title = ('{0}. {1}' -f $Index, $label); What = (Get-PcmGuiStepInfo $EngineTitle); Status = 'Pending'
        Counter = $(if ($Expected -gt 0) { ConvertTo-PcmGuiPlural $Expected 'object' 'objects' } else { '' }); Progress = [double]0; Note = ''; NoteBase = ''; NextAt = $null
        Elapsed = ''; Started = $null; Ended = $null; Expected = $Expected; Warnings = 0; WaitWhat = ''; WaitAttempt = 0; WaitMax = 0; WaitLimit = 0; SkipText = ''
    }
}

function Initialize-PcmGuiRunModel {
    <# The steps are known before the run starts (the engine confirms them with its own events, which replace this first list). #>
    param($Request)
    $g = $script:Gui; $c = $g.Controls
    $steps = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
    $model = @{ Steps = $steps; ById = @{}; Objects = @{}; ObjectIndex = @{}; Messages = @{}; Raw = @{}; Current = ''; CurrentObject = ''; Follow = $true; Dirty = $true; Selected = $null; Info = @{}; BatchId = ''; Names = @{} }
    $titles = @(Get-PcmGuiPlanTitle -Action $Request.Action -Follow:($Request.Action -eq 'Check' -and [bool]$Request.Batch))
    for ($i = 0; $i -lt $titles.Count; $i++) {
        $row = New-PcmGuiStepRow ('s' + ($i + 1)) ($i + 1) $titles[$i]
        $steps.Add($row); $model.ById[$row.Id] = $row
    }
    $g.Model = $model
    $g.Updating++
    try { $c.StepList.ItemsSource = $steps } finally { $g.Updating-- }
    $c.StepsTitle.Text = if ($steps.Count) { 'Steps ({0}, all known before the start)' -f $steps.Count } else { 'Steps' }
    $c.DetObjList.ItemsSource = $null; $c.DetMsgs.ItemsSource = $null
    $c.DetTitle.Text = 'Waiting for the first step...'; $c.DetWhat.Text = ''; $c.DetObjHeader.Text = ''; $c.DetPillText.Text = ''; $c.DetRaw.Text = ''
    $c.DetNoMsg.Visibility = 'Visible'; $c.WaitBox.Visibility = 'Collapsed'
    $c.ObjCard.Visibility = 'Collapsed'; $c.ObjEmpty.Visibility = 'Visible'
    $c.RunBar.Value = 0; $c.RunBarText.Text = ''
}

function Get-PcmGuiModelStep {
    <# The row of a step: by id, by index, or created on the fly when the engine announced more steps than planned. #>
    param([string]$Id, [int]$Index = 0, [string]$Title = '')
    $m = $script:Gui.Model
    if ($Id -and $m.ById.ContainsKey($Id)) { return , $m.ById[$Id] }
    if ($Index -gt 0 -and $m.ById.ContainsKey('s' + $Index)) { return , $m.ById['s' + $Index] }
    if (-not $Id) { $Id = if ($Index -gt 0) { 's' + $Index } else { 's' + ($m.Steps.Count + 1) } }
    if ($Index -le 0) { $Index = $m.Steps.Count + 1 }
    $row = New-PcmGuiStepRow $Id $Index $(if ($Title) { $Title } else { 'Step ' + $Index })
    $m.Steps.Add($row); $m.ById[$Id] = $row
    return , $row
}

function Set-PcmGuiPlanFromEvent {
    <# The engine announces its steps: they replace the planned ones. #>
    param($Record)
    $g = $script:Gui; $m = $g.Model
    $g.Updating++
    try { $m.Steps.Clear() } finally { $g.Updating-- }
    $m.ById.Clear(); $m.Selected = $null; $m.Follow = $true
    $index = 0
    foreach ($step in @(Get-PcmGuiValue $Record 'steps' @())) {
        $index++
        $id = [string](Get-PcmGuiValue $step 'id' ('s' + $index))
        $title = [string](Get-PcmGuiValue $step 'label' (Get-PcmGuiValue $step 'title' ('Step ' + $index)))
        $row = New-PcmGuiStepRow $id $index $title ([int](Get-PcmGuiValue $step 'count' 0))
        $m.Steps.Add($row); $m.ById[$id] = $row
    }
    $g.Controls.StepsTitle.Text = 'Steps ({0}, all known before the start)' -f $m.Steps.Count
    $m.Dirty = $true
}

function Add-PcmGuiStepRaw {
    <# Technical lines (item detail, engine console lines) are kept per step, for the "Technical message" box. #>
    param([string]$Text, [string]$StepId = '')
    $g = $script:Gui
    if (-not $g -or -not $g.Model) { return }
    $m = $g.Model
    $id = if ($StepId) { $StepId } else { [string]$m.Current }
    if (-not $id) { return }
    if (-not $m.Raw.ContainsKey($id)) { $m.Raw[$id] = New-Object 'Collections.Generic.List[string]' }
    $m.Raw[$id].Add($Text)
    while ($m.Raw[$id].Count -gt 400) { $m.Raw[$id].RemoveAt(0) }
    $m.Dirty = $true
}

function Close-PcmGuiOpenStep {
    <# A step that was running ends when the next one starts. #>
    param([string]$Except = '')
    $m = $script:Gui.Model
    foreach ($other in $m.Steps) {
        if ((Test-PcmGuiOrdinalIn $other.Status @('Running', 'Wait')) -and $other.Id -ne $Except) {
            $final = if ([int]$other.Warnings -gt 0 -and $m.Messages.ContainsKey([string]$other.Id) -and @($m.Messages[[string]$other.Id] | Where-Object { Test-PcmGuiOrdinal $_.Status 'Fail' }).Count) { 'Fail' } elseif ([int]$other.Warnings -gt 0) { 'Warn' } else { 'Ok' }
            $count = if ($m.Objects.ContainsKey([string]$other.Id)) { $m.Objects[[string]$other.Id].Count } else { 0 }
            $ended = [datetime]::UtcNow
            $began = if ($other.Started) { [datetime]$other.Started } else { $ended }
            $data = @{ Status = $final; Ended = $ended; Progress = [double]1; Note = ''; NoteBase = ''; NextAt = $null; Elapsed = (Format-PcmGuiClock (($ended - $began).TotalSeconds)) }
            if ($count) { $data.Counter = $(if ([int]$other.Expected -gt 0) { '{0} / {1} objects' -f $count, $other.Expected } else { ConvertTo-PcmGuiPlural $count 'object' 'objects' }) + $(if ($final -ne 'Ok') { ' · ' + (ConvertTo-PcmGuiPlural ([int]$other.Warnings) 'warning' 'warnings') } else { '' }) }
            Set-PcmGuiRow $other $data
        }
    }
}

function Invoke-PcmGuiStepEvent {
    <# A step starts (the previous one ends). A record with a state (running, ok, warn, fail, skipped) and an id works on that step. #>
    param($Record)
    $g = $script:Gui; $m = $g.Model
    $index = [int](Get-PcmGuiValue $Record 'index' 0)
    $title = [string](Get-PcmGuiValue $Record 'title' (Get-PcmGuiValue $Record 'label' ''))
    $total = [int](Get-PcmGuiValue $Record 'total' 0)
    # The engine announces more steps than the plan: the plan is stretched; fewer: the pending ones that will never come are dropped at the end.
    if ($total -gt 0 -and $m.Steps.Count -lt $total) { for ($i = $m.Steps.Count + 1; $i -le $total; $i++) { [void](Get-PcmGuiModelStep '' $i ('Step ' + $i)) } }
    $row = Get-PcmGuiModelStep ([string](Get-PcmGuiValue $Record 'id' '')) $index $title
    if ($title -and -not (Test-PcmGuiOrdinal $row.EngineTitle $title)) {
        $label = Get-PcmGuiStepLabel $title
        Set-PcmGuiRow $row @{ EngineTitle = $title; Label = $label; Title = ('{0}. {1}' -f $row.Index, $label); What = (Get-PcmGuiStepInfo $title) }
        $g.Controls.StepList.Items.Refresh()
    }
    $state = [string](Get-PcmGuiValue $Record 'state' '')
    if (-not $state) { Close-PcmGuiOpenStep -Except ([string]$row.Id); $state = 'running' }
    $status = ConvertTo-PcmGuiStatus $state
    $started = ConvertTo-PcmGuiUtc (Get-PcmGuiValue $Record 'startedUtc' ([datetime]::UtcNow.ToString('o')))
    switch ($status) {
        'Running' {
            Set-PcmGuiRow $row @{ Status = 'Running'; Started = $started; Ended = $null; Elapsed = '00:00' }
            $m.Current = [string]$row.Id
            if ($m.Follow) { $g.Updating++; try { $g.Controls.StepList.SelectedItem = $row } finally { $g.Updating-- }; $m.Selected = $row }
        }
        'Skip' { Set-PcmGuiRow $row @{ Status = 'Skip'; Progress = [double]0; Note = ''; NoteBase = ''; NextAt = $null } }
        default {
            $ended = ConvertTo-PcmGuiUtc (Get-PcmGuiValue $Record 'endedUtc' ([datetime]::UtcNow.ToString('o')))
            $begun = if ($row.Started) { [datetime]$row.Started } else { $started }
            Set-PcmGuiRow $row @{ Status = $status; Ended = $ended; Progress = [double]1; Note = ''; NoteBase = ''; NextAt = $null; Elapsed = (Format-PcmGuiClock (($ended - $begun).TotalSeconds)) }
        }
    }
    $m.Dirty = $true
}

function Get-PcmGuiItemName {
    <# The object an item is about: the name the engine gives (name), else the identity when it looks like an account name, never a DN, a server or a bare GUID. #>
    param($Record)
    if ($null -ne (Get-PcmGuiValue $Record 'name' $null)) { return [string](Get-PcmGuiValue $Record 'name' '') }
    $raw = [string](Get-PcmGuiValue $Record 'identity' '')
    if ($raw -match '\A[\w.+@-]+\z' -and $raw -notmatch '\A[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}\z' -and $raw -notmatch '\A(Batch|Run|Tenant|Exchange)\z') { return $raw }
    return ''
}

function Invoke-PcmGuiItemEvent {
    <# One line of the engine: an object of the current step (its result), a warning or an error, or a detail kept for the technical message. #>
    param($Record)
    $g = $script:Gui; $m = $g.Model
    $rawStatus = [string](Get-PcmGuiValue $Record 'status' 'Info')
    $status = ConvertTo-PcmGuiStatus $rawStatus
    $isSub = Test-PcmGuiOrdinal $rawStatus 'Sub'
    $text = [string](Get-PcmGuiValue $Record 'text' '')
    $identity = Get-PcmGuiItemName $Record
    # The engine writes most lines as '<object>: <what happened>': the object is shown on its own, the line keeps what happened.
    if ($identity -and $text.StartsWith($identity + ': ', [StringComparison]::OrdinalIgnoreCase)) { $text = $text.Substring($identity.Length + 2) }
    $detail = [string](Get-PcmGuiValue $Record 'detail' '')
    $stepId = [string](Get-PcmGuiValue $Record 'stepId' '')
    if (-not $stepId) { $stepId = [string]$m.Current }
    $row = $null
    if ($stepId -and $m.ById.ContainsKey($stepId)) { $row = $m.ById[$stepId] }
    elseif ($m.Steps.Count -and -not $m.Current) { $row = $m.Steps[0] }
    elseif ($m.Steps.Count) { $row = $m.Steps[$m.Steps.Count - 1] }
    if ($text -match '\bbatch\s+([0-9a-fA-F]{8})\b' -and -not $m.BatchId) { $m.BatchId = $Matches[1].ToLowerInvariant() }
    if ($row) {
        $id = [string]$row.Id
        if ($identity -and -not $isSub) {
            $m.CurrentObject = $identity
            $existing = Get-PcmGuiModelObject -Model $m -StepId $id -Identity $identity
            $shown = if ($status -eq 'Info') { 'Running' } else { $status }
            $priority = @{ Pending = 0; Info = 0; Running = 1; Ok = 2; Skip = 2; Warn = 3; Fail = 4; Wait = 1 }
            if ($existing) {
                $keep = [int]$priority[[string]$existing.Status] -gt [int]$priority[$shown] -or ([string]$existing.Status -eq 'Fail')
                Set-PcmGuiRow $existing @{ Result = $text; Status = $(if ($keep) { [string]$existing.Status } else { $shown }) }
            }
            else { Add-PcmGuiModelObject -Model $m -StepId $id -Row (New-PcmGuiRow @{ Name = $identity; Upn = ''; Result = $text; Status = $shown; Identity = $identity }) }
            $count = $m.Objects[$id].Count
            if ([int]$row.Expected -gt 0) { Set-PcmGuiRow $row @{ Counter = ('{0} / {1} objects' -f $count, $row.Expected); Progress = [double][Math]::Min(1.0, $count / [double]$row.Expected) } }
            else { Set-PcmGuiRow $row @{ Counter = (ConvertTo-PcmGuiPlural $count 'object' 'objects') } }
            Set-PcmGuiRow $row @{ Note = $(if ($status -eq 'Info' -and (Test-PcmGuiOrdinal $row.Status 'Running')) { 'Now: ' + $identity } else { [string]$row.Note }) }
        }
        elseif ((Test-PcmGuiOrdinal $status 'Skip') -and -not $identity -and (-not $m.Objects.ContainsKey($id) -or -not $m.Objects[$id].Count)) {
            Set-PcmGuiRow $row @{ Status = 'Skip'; SkipText = $text; Progress = [double]0; Note = ''; NoteBase = '' }
        }
        if ((Test-PcmGuiOrdinalIn $status @('Warn', 'Fail')) -and -not $isSub) {
            if (-not $m.Messages.ContainsKey($id)) { $m.Messages[$id] = New-Object 'Collections.ObjectModel.ObservableCollection[object]' }
            $m.Messages[$id].Add((New-PcmGuiRow @{ Text = $(if ($identity) { $identity + ': ' + $text } else { $text }); Status = $status; Identity = $identity }))
            Set-PcmGuiRow $row @{ Warnings = ([int]$row.Warnings + 1) }
        }
        Add-PcmGuiStepRaw $(if ($detail) { $text + "`r`n" + $detail } else { $text }) $id
        if (Test-PcmGuiOrdinal $row.Status 'Wait') { Set-PcmGuiRow $row @{ Status = 'Running' } }
    }
    $m.Dirty = $true
}

function Invoke-PcmGuiCountEvent {
    <# The engine knows how many objects it works on: the steps that work object by object show "0 / N objects". #>
    param($Record)
    $m = $script:Gui.Model
    $count = [int](Get-PcmGuiValue $Record 'count' 0)
    if ($count -le 0) { return }
    foreach ($id in @(Get-PcmGuiValue $Record 'stepIds' @())) {
        if (-not $m.ById.ContainsKey([string]$id)) { continue }
        $row = $m.ById[[string]$id]
        $done = if ($m.Objects.ContainsKey([string]$id)) { $m.Objects[[string]$id].Count } else { 0 }
        Set-PcmGuiRow $row @{ Expected = $count; Counter = ('{0} / {1} objects' -f $done, $count) }
    }
    $m.Dirty = $true
}

function Invoke-PcmGuiProgressEvent {
    <# Progress of the current step: done of total objects of a phase, or the attempts of a wait. #>
    param($Record)
    $g = $script:Gui; $m = $g.Model
    $stepId = [string](Get-PcmGuiValue $Record 'stepId' '')
    if (-not $stepId) { $stepId = [string]$m.Current }
    if (-not $stepId -or -not $m.ById.ContainsKey($stepId)) { return }
    $row = $m.ById[$stepId]
    $attempt = [int](Get-PcmGuiValue $Record 'attempt' 0); $max = [int](Get-PcmGuiValue $Record 'maxAttempts' 0)
    $next = [int](Get-PcmGuiValue $Record 'nextInSeconds' 0)
    $done = [double](Get-PcmGuiValue $Record 'done' (Get-PcmGuiValue $Record 'current' 0)); $total = [double](Get-PcmGuiValue $Record 'total' 0)
    $phase = [string](Get-PcmGuiValue $Record 'phase' (Get-PcmGuiValue $Record 'text' ''))
    if ($attempt -gt 0 -and $max -gt 0 -or $next -gt 0) {
        $note = $phase
        if ($attempt -gt 0 -and $max -gt 0) { $note += (' - attempt {0} / {1}' -f $attempt, $max) }
        $limit = if ($total -gt 0) { $total } elseif ($max -gt 0 -and $next -gt 0) { [double]$max * $next } else { 0 }
        $data = @{ NoteBase = $note; Note = $note; Progress = [double]$(if ($max -gt 0) { [Math]::Min(1.0, $attempt / [double]$max) } else { $row.Progress }); NextAt = $(if ($next -gt 0) { [datetime]::UtcNow.AddSeconds($next) } else { $null })
            WaitWhat = $phase; WaitAttempt = $attempt; WaitMax = $max; WaitLimit = [double]$limit }
        if ($next -gt 0) { $data.Note = $note + (', next check in {0} s' -f $next); $data.Status = 'Wait' } else { $data.Status = 'Running' }
        Set-PcmGuiRow $row $data
    }
    elseif ($total -gt 0) {
        $eta = [string](Get-PcmGuiValue $Record 'eta' '')
        Set-PcmGuiRow $row @{ Expected = [int]$total; Progress = [double][Math]::Max(0.0, [Math]::Min(1.0, $done / $total)); Counter = ('{0} / {1} objects' -f [int]$done, [int]$total) + $(if ($phase) { ' · ' + $phase } else { '' }); NoteBase = $eta }
    }
    $m.Dirty = $true
}

function Invoke-PcmGuiAskEvent {
    <# A question of the run (Entra Connect in Manual mode): asked here, answered in the file the run named next to its own event file. #>
    param($Record)
    $g = $script:Gui; $run = $g.Run; $m = $g.Model
    $id = [int](Get-PcmGuiValue $Record 'id' 0)
    if (-not $run.Asked.Add($id)) { return }
    $text = [string](Get-PcmGuiValue $Record 'text' '')
    $answer = [string](Get-PcmGuiValue $Record 'answer' '')
    Add-PcmGuiLine ('Question: ' + $text) 'Warn'
    # (no $x = if (...) { row }: an if used as a value would enumerate the dynamic row)
    $current = $null
    if ($m.Current -and $m.ById.ContainsKey($m.Current)) { $current = $m.ById[$m.Current] }
    if ($current) { Set-PcmGuiRow $current @{ Status = 'Wait'; WaitWhat = $text; WaitAttempt = 0; WaitMax = 0; WaitLimit = 0 } }
    $yes = $false
    if (-not $run.StopAsked) { $yes = Show-PcmGuiQuestion -Title ([string](Get-PcmGuiValue $Record 'title' 'PRA Cloud Mailbox')) -Text $text }
    # Only the answer file the run named next to its own event file.
    if ($answer -and $answer.StartsWith($run.EventPath + '.answer-', [StringComparison]::OrdinalIgnoreCase)) {
        [IO.File]::WriteAllText($answer, $(if ($yes) { 'yes' } else { 'no' }))
    }
    Add-PcmGuiLine ('Answer: ' + $(if ($yes) { 'yes' } else { 'no' })) 'Info'
    if ($current) { Set-PcmGuiRow $current @{ Status = 'Running' } }
}

function Invoke-PcmGuiEvent {
    param($Record)
    $g = $script:Gui; $run = $g.Run
    foreach ($field in @('title', 'text')) {
        $shown = [string](Get-PcmGuiValue $Record $field '')
        if ($shown.Length -ge 4) { $run.Seen.Add($shown); while ($run.Seen.Count -gt 400) { $run.Seen.RemoveAt(0) } }
    }
    switch ([string](Get-PcmGuiValue $Record 'kind' '')) {
        'start' {
            $run.LogFile = [string](Get-PcmGuiValue $Record 'logFile' '')
            Add-PcmGuiLine ([string](Get-PcmGuiValue $Record 'title' 'Start'))
            $details = Get-PcmGuiValue $Record 'details' $null
            if ($details) {
                foreach ($property in $details.PSObject.Properties) {
                    Add-PcmGuiLine ('{0}: {1}' -f $property.Name, $property.Value)
                    if ((Test-PcmGuiOrdinal $property.Name 'Batch') -and [string]$property.Value -match '^[0-9a-fA-F]{8}$') { $g.Model.BatchId = ([string]$property.Value).ToLowerInvariant() }
                }
            }
        }
        'plan' { Set-PcmGuiPlanFromEvent $Record }
        'step' {
            Invoke-PcmGuiStepEvent $Record
            if (-not (Get-PcmGuiValue $Record 'state' '')) { Add-PcmGuiLine ('{0}/{1} · {2}' -f [int](Get-PcmGuiValue $Record 'index' 0), [int](Get-PcmGuiValue $Record 'total' 0), (Get-PcmGuiValue $Record 'title' '')) }
        }
        'progress' { Invoke-PcmGuiProgressEvent $Record }
        'count' { Invoke-PcmGuiCountEvent $Record }
        'target' {
            $run.TargetText = [string](Get-PcmGuiValue $Record 'text' '')
            if ($run.TargetText) { Add-PcmGuiLine $run.TargetText }
        }
        'item' {
            $status = [string](Get-PcmGuiValue $Record 'status' 'Info')
            if ($run.Counts.ContainsKey($status)) { $run.Counts[$status]++ }
            Invoke-PcmGuiItemEvent $Record
            Add-PcmGuiLine ((([string](Get-PcmGuiValue $Record 'text' '')) + ' ' + (Get-PcmGuiValue $Record 'identity' '')).Trim()) $status
        }
        'ask' { Invoke-PcmGuiAskEvent $Record }
        'summary' { Add-PcmGuiLine ((Get-PcmGuiValue $Record 'title' '') + ' ' + ((Get-PcmGuiValue $Record 'values' @{}) | ConvertTo-Json -Compress -Depth 5)) ([string](Get-PcmGuiValue $Record 'status' 'Info')) }
        'result' { $run.ResultCount++; $run.Result = $Record }
    }
}

function Show-PcmGuiStepDetail {
    <# The middle pane: what the step does, its objects, its warnings and errors in plain words, and the technical message (collapsed). #>
    param($Row)
    $g = $script:Gui; $c = $g.Controls; $m = $g.Model
    if (-not $Row) { return }
    $id = [string]$Row.Id
    $c.DetTitle.Text = [string]$Row.Title
    $pill = switch ($Row.Status) { 'Ok' { @('Done', 'Ok') } 'Warn' { @('Done with a warning', 'Warn') } 'Fail' { @('Failed', 'Fail') } 'Running' { @('Running', 'Brand') } 'Wait' { @('Waiting', 'Warn') } 'Skip' { @('Skipped', 'Neutral') } default { @('Not started', 'Neutral') } }
    $c.DetPillText.Text = $pill[0]; Set-PcmGuiPill $c.DetPill $c.DetPillText $pill[1]
    $c.DetWhat.Text = [string]$Row.What + $(if ($Row.SkipText) { ' ' + [string]$Row.SkipText } else { '' })
    if (Test-PcmGuiOrdinal $Row.Status 'Wait') {
        $c.WaitBox.Visibility = 'Visible'
        $c.WaitTitle.Text = 'Waiting: ' + [string]$Row.WaitWhat
        $line = if ([int]$Row.WaitAttempt -gt 0 -and [int]$Row.WaitMax -gt 0) { 'Attempt {0} of {1}' -f [int]$Row.WaitAttempt, [int]$Row.WaitMax } else { 'In progress' }
        if ($Row.NextAt -and ([datetime]$Row.NextAt) -gt [datetime]::UtcNow) { $line += (' · next check in {0} s' -f [int][Math]::Ceiling((([datetime]$Row.NextAt) - [datetime]::UtcNow).TotalSeconds)) }
        $c.WaitLine.Text = $line
        $limit = [double]$Row.WaitLimit
        $c.WaitNormal.Text = if ($limit -ge 90) { 'This is normal: it can take up to {0} minutes.' -f [int][Math]::Ceiling($limit / 60) } elseif ($limit -gt 0) { 'This is normal: it can take up to {0} seconds.' -f [int][Math]::Ceiling($limit) } else { 'This is normal: Microsoft 365 takes its time.' }
    }
    else { $c.WaitBox.Visibility = 'Collapsed' }
    if ($m.Objects.ContainsKey($id)) { $c.DetObjList.ItemsSource = $m.Objects[$id]; $c.DetObjHeader.Text = ('OBJECTS  {0}' -f $m.Objects[$id].Count) }
    else { $c.DetObjList.ItemsSource = $null; $c.DetObjHeader.Text = '' }
    if ($m.Messages.ContainsKey($id) -and $m.Messages[$id].Count) { $c.DetMsgs.ItemsSource = $m.Messages[$id]; $c.DetNoMsg.Visibility = 'Collapsed' }
    else { $c.DetMsgs.ItemsSource = $null; $c.DetNoMsg.Visibility = 'Visible' }
    $c.DetRaw.Text = $(if ($m.Raw.ContainsKey($id)) { $m.Raw[$id] -join "`r`n" } else { '' })
}

function Show-PcmGuiObjectCard {
    <# The right pane: the journey of one object across the steps (from the item events), what the snapshot holds about it, and its warnings. #>
    param($Row)
    $g = $script:Gui; $c = $g.Controls; $m = $g.Model
    if (-not $Row) { $c.ObjCard.Visibility = 'Collapsed'; $c.ObjEmpty.Visibility = 'Visible'; return }
    $identity = [string]$Row.Identity
    $c.CardName.Text = ($identity -split '@')[0]
    $c.CardUpn.Text = $identity
    $found = @(Find-PcmGuiMailbox -State $g.State -Text $identity) | Select-Object -First 1
    $c.CardType.Text = if ($found) { if ($found.kind -eq 'User') { 'User mailbox' } else { ([string]$found.kind) + ' mailbox' } } else { 'Mailbox' }
    $journey = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
    $messages = New-Object 'Collections.ObjectModel.ObservableCollection[object]'
    foreach ($step in $m.Steps) {
        $id = [string]$step.Id
        $mine = $null
        $mine = Get-PcmGuiModelObject -Model $m -StepId $id -Identity $identity
        if (-not $mine) {
            # A step that works on the whole batch (connection, readiness, plan) is part of every object's journey.
            if ((-not $m.Objects.ContainsKey($id) -or -not $m.Objects[$id].Count) -and -not (Test-PcmGuiOrdinal $step.Status 'Skip')) {
                $journey.Add((New-PcmGuiRow @{ Title = [string]$step.Label; Caption = $(if ($step.Note) { 'Whole batch · ' + [string]$step.Note } else { 'Whole batch' }); Status = [string]$step.Status }))
            }
            continue
        }
        $journey.Add((New-PcmGuiRow @{ Title = [string]$step.Label; Caption = [string]$mine.Result; Status = [string]$mine.Status }))
        if (Test-PcmGuiOrdinalIn $mine.Status @('Warn', 'Fail')) { $messages.Add((New-PcmGuiRow @{ Text = [string]$mine.Result; Status = [string]$mine.Status })) }
    }
    $c.Journey.ItemsSource = $journey
    $c.CardMsgs.ItemsSource = $messages; $c.CardNoMsg.Visibility = Get-PcmGuiVisibility ($messages.Count -eq 0)
    $c.CardSnap.Text = Get-PcmGuiSnapshotFact -State $g.State -Identity $identity
    $c.ObjEmpty.Visibility = 'Collapsed'; $c.ObjCard.Visibility = 'Visible'
}

function Set-PcmGuiRunPill {
    param([string]$Text, [string]$Level)
    $c = $script:Gui.Controls
    $c.RunPillText.Text = $Text; Set-PcmGuiPill $c.RunPill $c.RunPillText $Level
}

function Update-PcmGuiRunView {
    <# Called by the window timer: elapsed times, the countdown of the waits, the overall bar, the pill and the header line. #>
    $g = $script:Gui; $m = $g.Model; $c = $g.Controls
    if (-not $m) { return }
    $now = [datetime]::UtcNow
    $done = 0; $fraction = [double]0
    foreach ($row in $m.Steps) {
        switch ($row.Status) {
            { (Test-PcmGuiOrdinalIn $_ @('Running', 'Wait')) } {
                $fraction = [double]$row.Progress
                if ($row.Started) { Set-PcmGuiRow $row @{ Elapsed = (Format-PcmGuiClock (($now - [datetime]$row.Started).TotalSeconds)) } }
                if ($row.NextAt -and ([datetime]$row.NextAt) -gt $now) { Set-PcmGuiRow $row @{ Note = ([string]$row.NoteBase + (', next check in {0} s' -f [int][Math]::Ceiling((([datetime]$row.NextAt) - $now).TotalSeconds))) } }
            }
            { (Test-PcmGuiOrdinalIn $_ @('Ok', 'Skip', 'Warn', 'Fail')) } { $done++ }
        }
    }
    $total = $m.Steps.Count
    $c.RunBar.Value = $(if ($total) { [Math]::Min(1.0, ($done + $fraction) / $total) } else { 0.0 })
    $elapsed = if ($g.Run) { $g.Run.Clock.Elapsed.TotalSeconds } elseif ($g.LastRun) { $g.LastRun.Seconds } else { 0 }
    $current = @($m.Steps | Where-Object { Test-PcmGuiOrdinalIn $_.Status @('Running', 'Wait') } | Select-Object -First 1)
    $objectCount = 0
    foreach ($id in $m.Objects.Keys) { $objectCount = [Math]::Max($objectCount, $m.Objects[$id].Count) }
    $expected = 0
    foreach ($row in $m.Steps) { $expected = [Math]::Max($expected, [int]$row.Expected) }
    $objectsText = if ($expected -gt 0) { ' · {0} / {1} mailboxes' -f $objectCount, $expected } elseif ($objectCount) { ' · ' + (ConvertTo-PcmGuiPlural $objectCount 'object' 'objects') } else { '' }
    $c.RunBarText.Text = if ($g.Run) { 'Step {0} of {1}{2} · {3} elapsed' -f $(if ($current.Count) { $current[0].Index } else { [Math]::Min($total, $done + 1) }), $total, $objectsText, (Format-PcmGuiClock $elapsed) } else { '{0} steps · {1}' -f $total, (Format-PcmGuiClock $elapsed) }
    if ($g.Run) {
        $run = $g.Run
        $now_ = if ($m.CurrentObject) { ' · now: ' + $m.CurrentObject } else { '' }
        $c.RunSub.Text = '{0} · {1}{2}{3} · {4} · started {5} on {6}{7}' -f $(if ($run.Request.Mode -eq 'Apply') { 'Apply' } elseif ($run.Request.Action -eq 'Check') { 'Read only' } else { 'Preview' }), (Get-PcmGuiField 'Env'), $(if ($run.Request.Contains('Batch') -and $run.Request.Batch -and $run.Request.Action -ne 'Check') { '' } else { ' · ' + (Get-PcmGuiObjectsText) }), $(if ($run.Config.Keys.Count) { ' · ' + (ConvertTo-PcmGuiPlural $run.Config.Keys.Count 'setting' 'settings') + ' changed for this run' } else { '' }), (Get-PcmGuiSnapshotText), $run.StartedAt.ToString('HH:mm', [Globalization.CultureInfo]::InvariantCulture), $env:COMPUTERNAME, $now_
        Set-PcmGuiRunPill $(if ($run.StopAsked) { 'Stopping' } else { 'Running' }) 'Brand'
    }
    if ($m.Dirty) {
        $m.Dirty = $false
        if ($m.Selected) { Show-PcmGuiStepDetail $m.Selected }
        elseif ($m.Steps.Count) { $m.Selected = $m.Steps[0]; Show-PcmGuiStepDetail $m.Selected }
    }
    elseif ($m.Selected -and (Test-PcmGuiOrdinalIn $m.Selected.Status @('Running', 'Wait'))) { Show-PcmGuiStepDetail $m.Selected }
    if ($g.Page -eq 3 -and $g.Run) {
        $c.VerBar.Value = $c.RunBar.Value
        $label = if ($current.Count) { [string]$current[0].Label } else { 'Starting...' }
        $c.VerEmptyText.Text = if ($current.Count -and $current[0].Note) { $label + ' · ' + $current[0].Note } else { $label }
    }
}
#endregion

#region The run: child process, events, result
function Set-PcmGuiBusy {
    <# While the engine works, the Setup and Task pages cannot be edited. #>
    param([bool]$Busy)
    $c = $script:Gui.Controls
    $c.Page1.IsEnabled = -not $Busy; $c.Page2.IsEnabled = -not $Busy
    foreach ($index in 1..5) { $c["Nav$index"].IsEnabled = -not $Busy -or $index -eq 4 }
}

function Sync-PcmGuiTimer {
    <# The timer runs while something is polled: a run, the Microsoft 365 test, the Exchange test or a group lookup. #>
    $g = $script:Gui
    if ($g.Run -or $g.NetJob -or $g.ExchJob -or $g.GroupLookup) { $g.Timer.Start() } else { $g.Timer.Stop() }
}

function Add-PcmGuiEngineOutput {
    <#
        The console lines of the engine, shown in the activity log as they arrive. A line that repeats an event already shown is skipped, so the log gains what the events do not carry (module
        messages, sign-in instructions, warnings). At most 60 lines per tick; the complete text is always kept in the .stdout.txt file next to the events.
    #>
    param($Run, [switch]$Final)
    $text = $Run.Output.TakeText()
    if (-not $text -and -not ($Final -and ($Run.Partial -or $Run.Hidden -gt 0))) { return }
    $parts = @(([string]$Run.Partial + $text) -split "`r?`n")
    $Run.Partial = [string]$parts[$parts.Count - 1]
    $lines = New-Object 'Collections.Generic.List[string]'
    for ($i = 0; $i -lt $parts.Count - 1; $i++) { $lines.Add([string]$parts[$i]) }
    if ($Run.Partial.Length -gt 8192 -or ($Final -and $Run.Partial)) { $lines.Add($Run.Partial); $Run.Partial = '' }
    $added = 0
    foreach ($raw in $lines) {
        $line = $raw.Trim()
        if (-not $line) { continue }
        $known = $false
        foreach ($seen in $Run.Seen) {
            $shown = $seen.Trim()
            if ($(if ($shown.Length -ge 16) { $line.Contains($shown) } else { (Test-PcmGuiOrdinal $line $shown) })) { $known = $true; break }
        }
        if ($known) { continue }
        if ($added -ge 60) { $Run.Hidden++; continue }
        Add-PcmGuiLine ('engine · ' + $line) 'Info'; $added++
        Add-PcmGuiStepRaw $line
    }
    if ($Run.Hidden -gt 0 -and ($added -lt 60 -or $Final)) { Add-PcmGuiLine ('engine · {0} output line(s) not shown; the full text is in the .stdout.txt file' -f $Run.Hidden) 'Info'; $Run.Hidden = 0 }
}

function Remove-PcmGuiCsvGuard {
    <# The report puts an apostrophe before a text that starts with = + - @ TAB or CR (a spreadsheet would read it as a formula): the shown value is the text without it. A value without a guard is unchanged. #>
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if ($null -ne $Value -and $Value -match "^'[=+\-@\t\r]") { return $Value.Substring(1) }
    return $Value
}

function Read-PcmGuiReport {
    <# The CSV report of a run (delimiter ;, UTF-8 with or without a byte order mark), with the guard the report puts before a formula removed from the shown values. #>
    param([string]$Path)
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter ';' -Encoding UTF8 -ErrorAction Stop)
    foreach ($row in $rows) {
        foreach ($property in @($row.PSObject.Properties)) {
            if ($property.Value -is [string]) { $property.Value = Remove-PcmGuiCsvGuard $property.Value }
        }
    }
    return $rows
}

function ConvertTo-PcmGuiObjectRow {
    <# A row of the report as a row of the preview list. #>
    param($Row)
    $g = $script:Gui
    $final = [string](Get-PcmGuiValue $Row 'FinalStatus' ''); $warnings = [string](Get-PcmGuiValue $Row 'Warnings' ''); $detail = [string](Get-PcmGuiValue $Row 'Detail' '')
    $kind = [string](Get-PcmGuiValue $Row 'Kind' '')
    $status = if (Test-PcmGuiOrdinalIn $final @('Error', 'Failed')) { 'Fail' } elseif (Test-PcmGuiOrdinal $final 'Skipped') { 'Skip' } elseif ($warnings.Trim()) { 'Warn' } else { 'Ok' }
    $first = { param($Text) $one = ($Text.Split('|')[0]).Trim(); if ($one.Length -gt 200) { $one.Substring(0, 197) + '...' } else { $one } }
    $result = switch ($status) {
        'Fail' { & $first $detail }
        'Skip' { if ($detail) { 'Skipped: ' + (& $first $detail) } else { 'Skipped' } }
        'Warn' { 'Ready · ' + (& $first $warnings) }
        default { 'Ready' }
    }
    $identity = [string](Get-PcmGuiValue $Row 'Identity' '')
    $guid = [string](Get-PcmGuiValue $Row 'ObjectGuid' '')
    $match = $null
    if ($guid -and $g.State.Objects) { [void](Get-PcmGuiObjectIndex -State $g.State).TryGetValue($guid, [ref]$match) }
    $name = if ($match -and $match.display_name) { [string]$match.display_name } elseif ($identity) { ($identity -split '@')[0] } else { '(run)' }
    $type = if (Test-PcmGuiOrdinal $kind 'User') { 'User' } elseif (Test-PcmGuiOrdinal $kind 'Shared') { 'Shared' } elseif ($kind) { $kind } else { 'Object' }
    return New-PcmGuiRow @{ Status = $status; Name = $name; Upn = $identity; Type = $type; Result = $result; Plain = $(if ($status -eq 'Ok') { '' } else { [string]$detail + $(if ($warnings) { "`r`n" + $warnings } else { '' }) }) }
}

function Get-PcmGuiRunOutcome {
    <# What a run did, from its result event and the exit code of its process: Valid (a result consistent with the exit code), Level (Success, Pending, Failed) and the first reason of a failure. #>
    param($Result, [int]$ExitCode, [string]$ErrorText = '')
    $issue = ''
    if ($Result) {
        foreach ($item in @(Get-PcmGuiValue $Result 'issues' @())) { if ([string]$item) { $issue = [string]$item; break } }
    }
    $model = if ($script:Gui) { $script:Gui.Model } else { $null }
    if (-not $issue -and $model) {
        foreach ($row in $model.Steps) {
            if (-not $model.Messages.ContainsKey([string]$row.Id)) { continue }
            foreach ($message in $model.Messages[[string]$row.Id]) { if (-not $issue -and (Test-PcmGuiOrdinal $message.Status 'Fail')) { $issue = [string]$message.Text } }
        }
    }
    if (-not $issue -and $ErrorText.Trim()) { $issue = ($ErrorText.Trim() -split "`r?`n")[0].Trim() }
    if (-not $Result) { return [pscustomobject]@{ Valid = $false; Level = 'Failed'; Issue = $(if ($issue) { $issue } else { 'The engine ended without a result.' }) } }
    $code = [int](Get-PcmGuiValue $Result 'exitCode' 1)
    $consistent = $code -eq $ExitCode
    $level = if ($code -eq 1) { 'Failed' } elseif ($code -eq 2) { 'Pending' } else { 'Success' }
    return [pscustomobject]@{ Valid = ($consistent -and $code -ne 1); Level = $level; Issue = $(if ($consistent) { $issue } else { 'The engine result does not match its exit code ({0}).' -f $ExitCode }) }
}

function Start-PcmGuiRun {
    param([ValidateSet('Preview', 'Apply')][string]$Mode, [ValidateSet('', 'Preview', 'Run')][string]$Purpose = '', [switch]$Schedule)
    $g = $script:Gui; $c = $g.Controls
    if ($g.Run) { throw 'A run is going: wait for its end.' }
    if ($g.GroupLookup) { throw 'The group check is still going: wait for its answer.' }
    if (-not $Purpose) { $Purpose = if ($Mode -eq 'Apply') { 'Run' } else { 'Preview' } }
    $problem = Get-PcmGuiSettingsProblem
    if ($problem) { throw $problem }
    $blocker = Get-PcmGuiBlocker
    if ($blocker) { throw $blocker }
    $request = if ($Schedule) { New-PcmGuiRequest -Action Schedule -Mode Apply } else { Get-PcmGuiWindowRequest $Mode }
    $word = $script:GuiActionWord[$request.Action]
    $key = ''
    if ($Mode -eq 'Apply' -and -not $Schedule) {
        if (-not $g.Pending -or -not (Test-PcmGuiConfirmation -Typed ([string]$c.ConfirmBox.Text) -Word $word)) {
            Reset-PcmGuiPreview; Update-PcmGuiView
            throw 'The run is refused: a successful, unchanged preview and the typed confirmation (the name of the action, any letter case) are required.'
        }
        $key = Get-PcmGuiRunKey -Request $request -Config $g.Pending.Config
        if (-not (Test-PcmGuiOrdinal $key $g.Pending.Key)) { Reset-PcmGuiPreview; Update-PcmGuiView; throw 'The run is refused: the task, the objects or the settings changed since the preview. Preview again.' }
    }
    $config = if ($Mode -eq 'Apply' -and -not $Schedule) { $g.Pending.Config } else { New-PcmGuiRunConfig -Root $g.Root -ConfigPath $g.ConfigPath -Overrides (Get-PcmGuiOverride) }
    $engine = Get-PcmGuiEngine
    if (-not $engine.Ok) { throw $engine.Problem }
    $pending = $g.Pending
    $previewKey = if ($key) { $key } else { Get-PcmGuiRunKey -Request $request -Config $config }
    $request = Set-PcmGuiSnapshotPin -Request $request -Pending $(if ($Mode -eq 'Apply' -and -not $Schedule) { $pending } else { $null })
    Reset-PcmGuiPreview
    $folder = Join-Path (Join-Path $g.Root 'logs') 'gui'
    Initialize-PcmGuiFolder (Join-Path $g.Root 'logs'); Initialize-PcmGuiFolder $folder
    $base = Join-Path $folder ([guid]::NewGuid().ToString('N'))
    $command = Get-PcmGuiCommand -Root $g.Root -ConfigPath $config.Path -Request $request -Engine $engine.Path -BaseConfigPath $config.Base
    if ($g.Hooks.NoConsole) { $command.NeedsConsoleWindow = $false }
    $child = Start-PcmGuiChildProcess -Command $command -Root $g.Root -EventPath ($base + '.events.jsonl') -StopPath ($base + '.stop')
    $g.Run = @{
        Request = $request; Config = $config; Purpose = $Purpose; Process = $child.Process; Output = $child.Output; Errors = $child.Errors; Key = $previewKey; Pending = $pending
        EventPath = ($base + '.events.jsonl'); StopPath = ($base + '.stop'); Base = $base; Position = [long]0; Result = $null; ResultCount = 0; LogFile = ''; TargetText = ''
        Clock = [Diagnostics.Stopwatch]::StartNew(); StartedAt = [datetime]::Now; StopAsked = $false; Command = $command
        Counts = @{ Ok = 0; Warn = 0; Fail = 0 }; Partial = ''; Seen = (New-Object 'Collections.Generic.List[string]'); Hidden = 0; Asked = (New-Object 'Collections.Generic.HashSet[int]')
    }
    $g.LastResult = $null; $g.LastRun = $null; $g.Lines.Clear(); $g.PreviewError = ''
    Set-PcmGuiLogHeader ($base + '.stdout.txt')
    if ($Purpose -eq 'Preview') { $g.VerRows.Clear(); $g.Previewed = $false }
    Initialize-PcmGuiRunModel $request
    $c.RunTitle.Text = '{0}{1}' -f $script:GuiActionTitle[$request.Action], $(if ($request.Batch -and $request.Action -ne 'Check') { ' · batch ' + $request.Batch } else { '' })
    $c.RunPillText.Text = ''
    Set-PcmGuiBusy $true
    Add-PcmGuiLine ('Started: ' + $command.Display)
    if ($config.Keys.Count) { Add-PcmGuiLine ('Settings changed for this run only (the configuration file is untouched): ' + ($config.Keys -join ', ')) }
    $g.Timer.Start()
    Update-PcmGuiRunView
    Update-PcmGuiView
}

function Start-PcmGuiSchedule {
    <# The button of the Collect page: registers the daily Collect task. It is the command line action Schedule, run like the other actions in a PowerShell window of its own, which asks for the account and its password there. #>
    $g = $script:Gui
    if ($g.Run) { throw 'A run is going: wait for its end.' }
    if ($g.Action -ne 'Collect') { throw 'Pick the task Collect first.' }
    $problem = Get-PcmGuiSettingsProblem
    if ($problem) { throw $problem }
    $blocker = Get-PcmGuiBlocker
    if ($blocker) { throw $blocker }
    $question = "Schedule the daily Collect on this computer?`n`nA PowerShell window opens and asks for the account that runs it (a domain user in the Exchange role group View-Only Organization Management) and its password. The console checks the account, then registers the task. A task of the same name is replaced.`n`nThe console must have been started as administrator."
    if (-not (Show-PcmGuiQuestion -Title 'Schedule the daily Collect' -Text $question)) { return }
    Start-PcmGuiRun -Mode Apply -Schedule
    Show-PcmGuiPage 4
}

function Get-PcmGuiLatestSnapshotId {
    <# The ID of the last complete snapshot of the snapshot file, read now: 0 when there is none or the file cannot be read. #>
    $g = $script:Gui
    try {
        $path = [string]$g.State.Config.Store.Path
        if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { return [long]0 }
        $connection = Open-PcmStore -Path $path -Root $g.Root -ReadOnly
        try {
            $latest = Get-PcmSnapshot $connection
            return $(if ($latest) { [long]$latest.id } else { [long]0 })
        } finally { Close-PcmStore $connection }
    } catch { return [long]0 }
}

function Set-PcmGuiSnapshotPin {
    <#
        A Check or a Convert of new objects reads one snapshot. The run is given its ID (-Snapshot): the preview reads the last complete snapshot as it is now, and the Apply that follows
        reads the one the preview read, never a snapshot written in between. A request that names a snapshot or a batch is returned as it is.
    #>
    param($Request, $Pending = $null)
    if (-not (Test-PcmGuiOrdinalIn $Request.Action @('Check', 'Convert')) -or $Request.Batch -or [long]$Request.Snapshot -gt 0) { return $Request }
    $id = if ($Pending) { [long]$Pending.Request.Snapshot } else { Get-PcmGuiLatestSnapshotId }
    if ($id -le 0) { return $Request }
    $pinned = [ordered]@{}
    foreach ($key in $Request.Keys) { $pinned[$key] = $Request[$key] }
    $pinned.Snapshot = $id
    return $pinned
}

function Get-PcmGuiGuardedFile {
    <#
        The files a preview and the run that follows it must find unchanged: the configuration file of the run (the copy made for it, and the file it was made from), the engine, the modules,
        the CSV file that lists the objects, and, for the actions that read it, the snapshot file and its .sha256 file. Full paths, once each, in a fixed order.
    #>
    param($Request, $Config)
    $g = $script:Gui
    $paths = [Collections.Generic.List[string]]::new()
    $paths.Add([string]$Config.Path); $paths.Add([string]$Config.Base)
    $paths.Add((Join-Path $g.Root 'Invoke-PraCloudMailbox.ps1'))
    $moduleFolder = Join-Path $g.Root 'module'
    if (Test-Path -LiteralPath $moduleFolder -PathType Container) {
        foreach ($file in @(Get-ChildItem -LiteralPath $moduleFolder -Recurse -File -ErrorAction SilentlyContinue | Where-Object { Test-PcmGuiOrdinalIn $_.Extension @('.psm1', '.psd1', '.ps1', '.dll', '.xaml') })) { $paths.Add($file.FullName) }
    }
    $paths.Add([string]$Request.CsvPath)
    $settings = if ($g.State) { $g.State.Config } else { $null }
    if ($settings) {
        if (Test-PcmGuiOrdinal $Request.Action 'Collect') { $paths.Add([string]$settings.Scope.CsvPath) }
        elseif (Test-PcmGuiOrdinalIn $Request.Action @('Check', 'Convert', 'Recover')) { $paths.Add([string]$settings.Store.Path); $paths.Add([string]$settings.Store.Path + '.sha256') }
    }
    $seen = New-Object 'Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
    $files = [Collections.Generic.List[string]]::new()
    foreach ($path in $paths) {
        if (-not $path) { continue }
        $full = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($path)) { $path } else { Join-Path $g.Root $path }))
        if ($seen.Add($full)) { $files.Add($full) }
    }
    $sorted = $files.ToArray()
    [Array]::Sort($sorted, [StringComparer]::OrdinalIgnoreCase)
    return $sorted
}

function Get-PcmGuiRunKey {
    <# What a preview and the run that follows it must share: the request, the configuration it uses and the content of every file it reads (see Get-PcmGuiGuardedFile). #>
    param($Request, $Config)
    $parts = [Collections.Generic.List[string]]::new()
    foreach ($key in @($Request.Keys | Where-Object { -not (Test-PcmGuiOrdinal $_ 'Mode') } | Sort-Object)) { $parts.Add(('{0}={1}' -f $key, $Request[$key])) }
    $parts.Add('config=' + [string]$Config.Path); $parts.Add('base=' + [string]$Config.Base); $parts.Add('hash=' + [string]$Config.Hash)
    foreach ($file in @(Get-PcmGuiGuardedFile -Request $Request -Config $Config)) {
        $parts.Add($(if (Test-Path -LiteralPath $file -PathType Leaf) { 'file={0}={1}' -f $file.ToLowerInvariant(), (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash } else { 'missing=' + $file.ToLowerInvariant() }))
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))) -replace '-', '').ToLowerInvariant() } finally { $sha.Dispose() }
}

function Stop-PcmGuiRun {
    <# Asks the run to stop: it finishes the object in progress, starts no other one and ends normally (journal, report); the batch can then be continued. #>
    $g = $script:Gui; $run = $g.Run
    if (-not $run -or $run.StopAsked) { return }
    [IO.File]::WriteAllText($run.StopPath, (Get-Date).ToString('o'))
    $run.StopAsked = $true
    Add-PcmGuiLine 'Stop asked: the object in progress is finished, no other one is started; the run then ends normally. Continue the batch to go on.' 'Warn'
    Set-PcmGuiRunPill 'Stopping' 'Warn'
    Update-PcmGuiButtons
}

function Invoke-PcmGuiTick {
    $g = $script:Gui
    if (-not $g -or $g.InTick) { return }
    if ($g.NetJob) { Complete-PcmGuiNetworkCheck }
    if ($g.ExchJob) { Complete-PcmGuiExchangeCheck }
    if ($g.GroupLookup) { Complete-PcmGuiGroupLookup }
    if (-not $g.Run) { Sync-PcmGuiTimer; return }
    $g.InTick = $true
    try {
        $run = $g.Run; $exited = $run.Process.HasExited
        try {
            do {
                $read = Read-PcmGuiEvents -Path $run.EventPath -Position $run.Position
                $old = $run.Position; $run.Position = $read.Position
                foreach ($record in $read.Events) { try { Invoke-PcmGuiEvent $record } catch { Add-PcmGuiLine ('(event not shown: {0})' -f $_.Exception.Message) 'Info' } }
            } while ($exited -and $read.Position -gt $old -and (Get-Item -LiteralPath $run.EventPath).Length -gt $read.Position)
        } catch { Add-PcmGuiLine $_.Exception.Message 'Fail' }
        Add-PcmGuiEngineOutput $run
        Update-PcmGuiRunView
        if ($g.Page -eq 4 -or $g.Page -eq 3) { Update-PcmGuiButtons }
        if ($exited -and $run.Output.IsCompleted -and $run.Errors.IsCompleted) { Add-PcmGuiEngineOutput $run -Final; Complete-PcmGuiRun }
    } finally { $g.InTick = $false }
}

function Complete-PcmGuiRun {
    $g = $script:Gui; $run = $g.Run; $c = $g.Controls
    $g.Timer.Stop()
    $exitCode = -1; $stderr = ''; $rows = @(); $result = $run.Result
    try {
        $null = $run.Output.Result; $stderr = [string]$run.Errors.Result
        $exitCode = $run.Process.ExitCode
        try {
            $csv = [string](Get-PcmGuiValue $result 'csvReport' '')
            if ($csv -and (Test-Path -LiteralPath $csv)) { $rows = @(Read-PcmGuiReport $csv) }
        } catch { Add-PcmGuiLine ('The report cannot be read: ' + $_.Exception.Message) 'Warn' }
        if ($stderr.Trim()) { Add-PcmGuiLine $stderr.Trim() 'Fail' }
    } catch { Add-PcmGuiLine ('The result or the log cannot be read: ' + $_.Exception.Message) 'Fail' }
    $outcome = Get-PcmGuiRunOutcome -Result $result -ExitCode $exitCode -ErrorText $stderr
    $planned = [int](Get-PcmGuiValue $result 'planned' 0)
    $usablePreview = $run.Purpose -eq 'Preview' -and $result -and ($outcome.Valid -or ($outcome.Level -eq 'Failed' -and $planned -gt 0 -and [int](Get-PcmGuiValue $result 'exitCode' 1) -eq 1 -and $rows.Count)) -and $run.Request.Action -ne 'Check'
    $seconds = $run.Clock.Elapsed.TotalSeconds
    try { $run.Process.Dispose() } catch { Write-Verbose $_.Exception.Message }
    $g.Run = $null
    foreach ($row in @($g.Model.Steps)) {
        if (Test-PcmGuiOrdinalIn $row.Status @('Running', 'Wait')) { Set-PcmGuiRow $row @{ Status = $(if ($outcome.Valid) { 'Ok' } else { 'Fail' }); Note = ''; NoteBase = ''; NextAt = $null; Progress = [double]1 } }
        elseif ($row.Status -eq 'Pending' -and $outcome.Valid) { Set-PcmGuiRow $row @{ Status = 'Skip'; Progress = [double]0 } }
    }
    $g.LastResult = if ($result) { $result } else { [pscustomobject]@{ logFile = $run.Base + '.stdout.txt'; batchId = '' } }
    $g.LastRun = [pscustomobject]@{ Request = $run.Request; Result = $result; Valid = $outcome.Valid; Level = $outcome.Level; ExitCode = $exitCode; Base = $run.Base; Seconds = $seconds; Issue = $outcome.Issue; Purpose = $run.Purpose
        Config = $run.Config; StartedAt = $run.StartedAt; Stopped = [bool]$run.StopAsked; Counts = $run.Counts; BatchId = $g.Model.BatchId }
    $g.PreviewResult = $result
    $logFile = [string](Get-PcmGuiValue $result 'logFile' '')
    if ($logFile) { Set-PcmGuiLogHeader $logFile }
    if ($run.Purpose -eq 'Preview') {
        $g.VerRows.Clear()
        # What needs attention first (blocked, warnings, skipped), then the users, then the shared mailboxes; the order of the report inside each group.
        $weight = @{ Fail = 0; Warn = 1; Skip = 2; Ok = 3 }
        $position = 0
        $keys = [Collections.Generic.List[string]]::new(); $objectRows = [Collections.Generic.List[object]]::new()
        foreach ($csvRow in $rows) {
            $objectRow = ConvertTo-PcmGuiObjectRow $csvRow
            $position++
            $keys.Add(('{0}{1}{2:D8}' -f [int]$weight[[string]$objectRow.Status], $(if (Test-PcmGuiOrdinal $objectRow.Type 'Shared') { 1 } else { 0 }), $position))
            $objectRows.Add($objectRow)
        }
        $ordered = $keys.ToArray(); $items = $objectRows.ToArray()
        [Array]::Sort([Array]$ordered, [Array]$items, [Collections.IComparer][StringComparer]::OrdinalIgnoreCase)
        foreach ($objectRow in $items) { $g.VerRows.Add($objectRow) }
        $g.Previewed = [bool]$usablePreview
        $g.PreviewError = if ($usablePreview) { '' } else { $outcome.Issue }
        $c.TargetText.Text = [string]$run.TargetText
        if ($usablePreview) {
            $g.Pending = @{ Key = $run.Key; Config = $run.Config; Request = $run.Request }
        }
    }
    Set-PcmGuiBusy $false
    Read-PcmGuiStateKeepingSettings
    $g.Model.Dirty = $true
    Update-PcmGuiRunView
    $warned = $run.Counts.Warn -gt 0 -or @($g.Model.Steps | Where-Object { Test-PcmGuiOrdinal $_.Status 'Warn' }).Count -gt 0
    if ($outcome.Valid) { Set-PcmGuiRunPill $(if ($run.StopAsked) { 'Stopped' } elseif ($warned) { 'Done with warnings' } else { 'Done' }) $(if ($warned -or $run.StopAsked) { 'Warn' } else { 'Ok' }) }
    elseif ($outcome.Level -eq 'Failed') { Set-PcmGuiRunPill 'Failed' 'Fail' }
    if ($run.Purpose -ne 'Preview') { Update-PcmGuiNext }
    if ($run.Request.Action -eq 'Schedule') { try { Invoke-PcmGuiDetection -Refresh } catch { Add-PcmGuiLine ('The lights could not be refreshed: ' + $_.Exception.Message) 'Warn' } }
    Update-PcmGuiBatchList
    Update-PcmGuiView
    Sync-PcmGuiTimer
    if ($g.Pending) { Add-PcmGuiLine ('Preview validated. Confirm by typing {0} (capitals or not); any change requires a new preview.' -f $script:GuiActionWord[$run.Request.Action]) 'Ok' }
    if ($run.Purpose -eq 'Preview') { Show-PcmGuiPage 3 }
}
#endregion

#region Next steps
function Add-PcmGuiNextStep {
    param([int]$Number, [string]$Text)
    [void]$script:Gui.NextRows.Add((New-PcmGuiRow @{ N = [string]$Number; Text = $Text }))
}

function Add-PcmGuiTimeline {
    param([string]$Title, [string]$Caption, [string]$Status, [bool]$Last)
    [void]$script:Gui.TimelineRows.Add((New-PcmGuiRow @{ Title = $Title; Caption = $Caption; Status = $Status; Last = $Last }))
}

function Get-PcmGuiNextCommand {
    <# The command the engine proposes as the next step, or the one the console builds from the batch. #>
    param($Result, [string]$Action, [string]$BatchId, [string]$Mode = 'Preview')
    $g = $script:Gui
    foreach ($line in @(Get-PcmGuiValue $Result 'nextSteps' @())) { if (([string]$line).IndexOf('Invoke-PraCloudMailbox', [StringComparison]::OrdinalIgnoreCase) -ge 0) { return [string]$line } }
    if (-not $BatchId) { return '' }
    $request = New-PcmGuiRequest -Action $Action -Mode $Mode -Batch $BatchId
    return (Get-PcmGuiCommand -Root $g.Root -ConfigPath $g.ConfigPath -Request $request).Display
}

function Get-PcmGuiRollbackCarry {
    <#
        What the command that continues a rollback carries so that it stays on the objects of the wave: the objects of the interrupted rollback still to roll back, as a CSV file written for
        the command, or else the way the run picked its objects. Nothing when the rollback covers the whole batch.
    #>
    param($Request, [string]$BatchId)
    $g = $script:Gui
    $carry = @{}
    $summary = @(Get-PcmGuiBatchSummary -State $g.State | Where-Object { Test-PcmGuiOrdinal $_.Id $BatchId }) | Select-Object -First 1
    if ($summary -and @($summary.Wave).Count) {
        $folder = Join-Path (Join-Path $g.Root 'logs') 'gui'
        Initialize-PcmGuiFolder (Join-Path $g.Root 'logs'); Initialize-PcmGuiFolder $folder
        $carry.CsvPath = Write-PcmGuiWave -Folder $folder -Name $BatchId -Rows @($summary.Wave)
        return $carry
    }
    if ($summary -and [int]$summary.Undone) { return $carry }
    foreach ($name in @('CsvPath', 'Identity', 'SearchBase', 'Group')) { if ($Request.Contains($name) -and $Request[$name]) { $carry[$name] = [string]$Request[$name] } }
    if ($Request.Contains('Scope') -and $Request.Scope -and -not (Test-PcmGuiOrdinal $Request.Scope 'All')) { $carry.Scope = [string]$Request.Scope }
    return $carry
}

function Get-PcmGuiReportSplit {
    <# The users and the shared mailboxes that a run completed, from its CSV report ($null when the report cannot be read). #>
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        $rows = @(Read-PcmGuiReport $Path | Where-Object { Test-PcmGuiOrdinalIn $_.FinalStatus @('Success', 'Planned') })
        $users = @($rows | Where-Object { Test-PcmGuiOrdinal $_.Kind 'User' }).Count
        return [pscustomobject]@{ Users = $users; Shared = $rows.Count - $users }
    } catch { return $null }
}

function Update-PcmGuiNext {
    $g = $script:Gui; $c = $g.Controls; $last = $g.LastRun
    $g.NextRows.Clear(); $g.TimelineRows.Clear()
    $c.NextCmdBox.Visibility = 'Collapsed'; $c.ContinueCmdBox.Visibility = 'Collapsed'; $c.BatchIdBox.Visibility = 'Collapsed'
    $c.BtnReport.IsEnabled = $false; $c.BtnLog.IsEnabled = $false; $c.BtnGlobalReport.IsEnabled = $false; $c.BtnGlobalReport.Visibility = 'Collapsed'
    $c.HowTitle.Text = 'How to do it'
    if (-not $last) {
        $c.NextPillText.Text = 'Nothing to do yet'; Set-PcmGuiPill $c.NextPill $c.NextPillText 'Neutral'
        $c.NextTitle.Text = 'No run has finished yet'; $c.NextLead.Text = 'Start a task, verify it and run it: what to do next appears here.'
        $c.ResultText.Text = 'No result.'; $c.StatusLine.Text = ''
        return
    }
    $r = $last.Result; $request = $last.Request; $action = [string]$request.Action
    $warnings = [int](Get-PcmGuiValue $r 'warnings' 0); $success = [int](Get-PcmGuiValue $r 'success' 0); $pending = [int](Get-PcmGuiValue $r 'pending' 0)
    $errors = [int](Get-PcmGuiValue $r 'error' 0); $skipped = [int](Get-PcmGuiValue $r 'skipped' 0)
    $batchId = [string](Get-PcmGuiValue $r 'batchId' ''); if (-not $batchId) { $batchId = [string]$last.BatchId }
    if (-not $batchId -and $request.Batch) { $batchId = [string]$request.Batch }
    # A rollback is asked and continued by the ID of the Convert batch it rolls back (the one the batch lists show); the ID the engine gives its own batch is only named.
    $rollbackId = ''
    if ($action -eq 'Recover' -and $request.Batch) { $rollbackId = $batchId; $batchId = [string]$request.Batch }
    $htmlReport = [string](Get-PcmGuiValue $g.LastResult 'htmlReport' ''); $globalReport = [string](Get-PcmGuiValue $g.LastResult 'globalReport' '')
    $c.BtnLog.IsEnabled = [bool](Get-PcmGuiValue $g.LastResult 'logFile' '')
    $c.BtnReport.IsEnabled = [bool]$htmlReport
    if ($globalReport) { $c.BtnGlobalReport.Visibility = 'Visible'; $c.BtnGlobalReport.IsEnabled = $true }
    $steps = @(Get-PcmGuiValue $r 'nextSteps' @() | ForEach-Object { $_ } | Where-Object { $_ })
    $commandLine = @($steps | Where-Object { ([string]$_).IndexOf('Invoke-PraCloudMailbox', [StringComparison]::OrdinalIgnoreCase) -ge 0 })
    $others = @($steps | Where-Object { ([string]$_).IndexOf('Invoke-PraCloudMailbox', [StringComparison]::OrdinalIgnoreCase) -lt 0 })
    $env_ = Get-PcmGuiField 'Env'
    $time = $last.StartedAt.ToString('HH:mm', [Globalization.CultureInfo]::InvariantCulture)
    $addStep = { param($Text) Add-PcmGuiNextStep ($g.NextRows.Count + 1) $Text }
    if (-not $last.Valid -and $last.Level -eq 'Failed') {
        $c.NextPillText.Text = 'Failed'; Set-PcmGuiPill $c.NextPill $c.NextPillText 'Fail'
        $c.NextTitle.Text = 'The run stopped'
        $c.NextLead.Text = 'Nothing is silently half-done: every step before the failure is journaled. Reason: ' + [string]$last.Issue
        $c.HowTitle.Text = 'What to do'
        & $addStep 'Go back to the Run page and click the red step: it explains in plain words what failed. The log has every detail.'
        & $addStep 'Fix the cause (the message above says what to look at).'
        if ($action -in @('Convert', 'Recover') -and $batchId) { & $addStep "Continue the batch: click 'Start a new task', pick the same batch ($batchId), preview and run. Objects already done are recognised and skipped." }
        elseif ($action -eq 'Schedule') { & $addStep "Click 'Start a new task', pick Collect, then click 'Schedule the daily Collect' again." }
        else { & $addStep "Click 'Start a new task', pick the same task again, preview and run." }
        foreach ($line in $others) { & $addStep ([string]$line) }
        $c.ResultText.Text = 'Failed (exit code {0}). {1} done · {2} error(s). The log and the report say which object failed.' -f $last.ExitCode, $success, $errors
        $c.NextPillText.Text = 'Failed'
        $pillLevel = 'Fail'
    } else {
        $pillLevel = 'Ok'
        $stopped = [bool]$last.Stopped
        if ($pending -gt 0 -or $last.Level -eq 'Pending') { $c.NextPillText.Text = $(if ($stopped) { 'Stopped as asked' } else { 'Next step required' }); $pillLevel = 'Warn' }
        elseif ($warnings -gt 0) { $c.NextPillText.Text = 'Done · {0}' -f (ConvertTo-PcmGuiPlural $warnings 'warning' 'warnings'); $pillLevel = 'Warn' }
        else { $c.NextPillText.Text = 'Done' }
        $objectsText = ConvertTo-PcmGuiPlural $success 'mailbox' 'mailboxes'
        switch ($action) {
            'Convert' {
                $c.NextTitle.Text = if ($pending -gt 0) { 'Part of the conversion remains' } else { 'The conversion is complete' }
                $c.NextLead.Text = if ($pending -gt 0) { '{0} are now in Exchange Online, {1} not yet. Continue the batch below: objects already done are skipped.' -f $objectsText, $pending } else { '{0} are now in Exchange Online. Keep the batch ID below: you need it to roll this batch back.' -f $objectsText }
                if ($warnings) { & $addStep ('Open the report and read the {0}. Nothing else needs your action: the other mailboxes are in Exchange Online.' -f (ConvertTo-PcmGuiPlural $warnings 'warning' 'warnings')) }
                else { & $addStep 'Open the report: it lists every mailbox with what was done.' }
                if ($pending -gt 0) { & $addStep ("Continue the batch: Task, Continue a batch, pick batch {0}, preview and run. Or use the command below." -f $batchId) }
                & $addStep ('Write down the batch ID {0} and keep the journal file. You need both to roll this batch back.' -f $batchId)
                & $addStep 'Run Check every day while the mailboxes are in the cloud: it tells you at once if one stops being ready.'
                & $addStep 'When Active Directory, Exchange and Entra Connect are rebuilt, come back here: Task, Roll back a conversion, pick this batch, Preview first.'
            }
            'Recover' {
                $c.NextTitle.Text = if ($pending -gt 0) { 'Part of the rollback remains' } else { 'The rollback is complete' }
                $c.NextLead.Text = if ($pending -gt 0) { '{0} are back on-premises, {1} not yet. Continue the rollback below.' -f $objectsText, $pending } else { '{0} are back on-premises. Their cloud data stays under the hold.' -f $objectsText }
                & $addStep 'Open the report and read what was done for each mailbox.'
                & $addStep 'Run Check on this batch with "Back on-premises" to confirm that every mailbox is where it should be.'
                if ($pending -gt 0) { & $addStep ("Continue the rollback: Task, Roll back a conversion, Continue an interrupted rollback, pick batch {0}." -f $batchId) }
            }
            'Schedule' {
                $taskName = if ($g.State.Config) { [string]$g.State.Config.Collect.Schedule.TaskName } else { 'PraCloudMailbox Collect' }
                $c.NextTitle.Text = 'The daily Collect is scheduled'
                $c.NextLead.Text = 'Windows now runs Collect every day with the account you gave. The light Daily Collect on the Setup page shows the task.'
                & $addStep ("Run the task once now, to see it work: open PowerShell as administrator and run Start-ScheduledTask -TaskName '{0}'." -f $taskName.Replace("'", "''"))
                & $addStep 'Look at the light Daily Collect on the Setup page: it shows the task and the time of its last run.'
                & $addStep 'Back up the snapshot file every day and copy it to the cloud admin server.'
            }
            'Collect' {
                $c.NextTitle.Text = 'The snapshot is written'
                $c.NextLead.Text = 'The snapshot is in the database. After the disaster nothing on-premises can be read, so this is what Convert will use.'
                & $addStep 'Copy the database file (and its backup folder) to the cloud admin server, next to the configuration file.'
                & $addStep 'Keep running Collect often: every day is best. A scheduled task can do it.'
            }
            default {
                $c.NextTitle.Text = if ($errors) { 'The check found objects that are not ready' } else { 'The check is finished' }
                $c.NextLead.Text = if ($errors) { 'Nothing was changed. Open the report: each object that is not ready says why.' } else { 'Nothing was changed. Open the report to read what is ready in Microsoft 365.' }
                & $addStep 'Open the report and read the objects that are not ready or have a warning.'
                & $addStep 'Fix what the report says, then run the check again.'
            }
        }
        foreach ($line in $others) { & $addStep ([string]$line) }
        $split = Get-PcmGuiReportSplit -Path ([string](Get-PcmGuiValue $r 'csvReport' ''))
        $lead = switch ($action) { 'Convert' { 'Converted {0}' } 'Recover' { 'Rolled back {0}' } 'Check' { '{0} checked' } default { '{0} collected' } }
        $counted = '{0}{1}' -f $success, $(if ($split) { ' ({0} users, {1} shared)' -f $split.Users, $split.Shared } else { '' })
        $c.ResultText.Text = if ($action -eq 'Schedule') { 'The daily Collect task is registered · {0} · {1} failed. The report and the log say what was checked.' -f (ConvertTo-PcmGuiPlural $warnings 'warning' 'warnings'), $errors }
            else { '{0} · {1} · {2} failed · {3} skipped. The report lists every mailbox.' -f ($lead -f $counted), (ConvertTo-PcmGuiPlural $warnings 'warning' 'warnings'), $errors, $skipped }
    }
    Set-PcmGuiPill $c.NextPill $c.NextPillText $pillLevel
    $command = Get-PcmGuiNextCommand -Result $r -Action $(if ($action -eq 'Convert' -and $last.Valid -and $pending -eq 0) { 'Recover' } else { $action }) -BatchId $batchId
    if ($last.Valid -and $action -eq 'Convert' -and $pending -eq 0 -and $batchId) {
        $c.NextCmdLabel.Text = 'Or from the command line (later, to roll this batch back: preview first)'
        $c.NextCmd.Text = $command
        $c.NextCmdBox.Visibility = 'Visible'
    }
    elseif ($commandLine.Count) { $c.NextCmdLabel.Text = 'Or from the command line'; $c.NextCmd.Text = [string]$commandLine[$commandLine.Count - 1]; $c.NextCmdBox.Visibility = 'Visible' }
    if (($pending -gt 0 -or -not $last.Valid) -and $action -in @('Convert', 'Recover') -and $batchId) {
        $continueAction = $action
        $convertBatch = $batchId
        $continueOptions = @{ Action = $continueAction; Mode = 'Preview'; Batch = $convertBatch }
        if ($continueAction -eq 'Recover') { foreach ($pair in (Get-PcmGuiRollbackCarry -Request $request -BatchId $convertBatch).GetEnumerator()) { $continueOptions[$pair.Key] = $pair.Value } }
        $c.ContinueCmd.Text = (Get-PcmGuiCommand -Root $g.Root -ConfigPath $g.ConfigPath -Request (New-PcmGuiRequest @continueOptions) -Engine 'pwsh.exe').Display
        $c.ContinueCmdBox.Visibility = 'Visible'
    }
    if ($action -in @('Convert', 'Recover') -and $batchId) {
        $c.BatchIdBox.Visibility = 'Visible'
        $c.BatchIdText.Text = $batchId
        $objectsLine = Get-PcmGuiObjectsText
        if ($objectsLine) { $objectsLine = $objectsLine.Substring(0, 1).ToLowerInvariant() + $objectsLine.Substring(1) }
        $c.BatchLine.Text = '{0} · {1} · {2} · {3}' -f $action, $last.StartedAt.ToString('d MMM yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture), $objectsLine, $(if ($g.State.Snapshot) { 'snapshot #' + $g.State.Snapshot.id } else { 'no snapshot' })
        if ($rollbackId -and -not (Test-PcmGuiOrdinal $rollbackId $batchId)) { $c.BatchLine.Text += ' · rollback batch ' + $rollbackId }
        $c.BtnReportLink.Content = $htmlReport
        $c.ReportRow.Visibility = Get-PcmGuiVisibility ([bool]$htmlReport)
    }
    $c.StatusLine.Text = if ($batchId) { '{0} · {1}' -f $batchId, $env_ } else { $env_ }
    $snapshot = $g.State.Snapshot
    Add-PcmGuiTimeline 'Collect' $(if ($snapshot) { 'Snapshot #{0}, {1}, on {2}' -f $snapshot.id, (Format-PcmGuiTime ([string]$snapshot.finished_utc)), $(if ($snapshot.PSObject.Properties['computer'] -and $snapshot.computer) { $snapshot.computer } else { 'the Exchange server' }) } else { 'No snapshot yet' }) $(if ($snapshot) { 'Ok' } else { 'Skip' }) $false
    $summary = @(Get-PcmGuiBatchSummary -State $g.State | Where-Object { Test-PcmGuiOrdinal $_.Id $batchId }) | Select-Object -First 1
    if ($action -eq 'Convert' -and $last.Valid) { Add-PcmGuiTimeline 'Convert' ('Done {0} on {1}: {2}{3}' -f $time, $env:COMPUTERNAME, (ConvertTo-PcmGuiPlural $success 'mailbox' 'mailboxes'), $(if ($warnings) { ', ' + (ConvertTo-PcmGuiPlural $warnings 'warning' 'warnings') } else { '' })) $(if ($warnings) { 'Warn' } else { 'Ok' }) $false }
    elseif ($summary) { Add-PcmGuiTimeline 'Convert' ('{0} · {1} converted, {2}' -f $summary.Status, $summary.Converted, $summary.Recover.ToLowerInvariant().Replace('-', 'no rollback')) $(if ($summary.Status -eq 'Complete') { 'Ok' } else { 'Warn' }) $false }
    else { Add-PcmGuiTimeline 'Convert' 'Not started: it runs after the disaster.' 'Skip' $false }
    if ($action -eq 'Recover' -and $last.Valid) { Add-PcmGuiTimeline 'Recover' ('Done {0} on {1}: {2} back' -f $time, $env:COMPUTERNAME, (ConvertTo-PcmGuiPlural $success 'mailbox' 'mailboxes')) 'Ok' $true }
    elseif ($summary -and $summary.Recover -like 'Partial*') { Add-PcmGuiTimeline 'Recover' $summary.Recover 'Warn' $true }
    else { Add-PcmGuiTimeline 'Recover' 'Only when the infrastructure is rebuilt. Not before.' 'Skip' $true }
}
#endregion

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDgudFiZ5BDTvke
# 0mAw1+QPlCuClVjwCFMjSvpTMlYjQKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAUnVx5x4+U
# jAOXKi5MOVS53S7WeN2GLWaawHrkj1lk1zANBgkqhkiG9w0BAQEFAASCAgCIjTQj
# Y3CJKpOGLUQOgwY/NKCPs7mgQM4hnE9FY92/hR485wDwfeOEmlx49E9UvH5jK8tw
# JQHA0PPQmzClOUxTTzJ9d9lvzcYzRTBBuRKNT/A8FzIsEz4/NqrLbSGakMK6lqKm
# ns2HDNr8BXR29r3WlXdL0drT0YiFw7vav/0yAufHCKcTSnQ9Lm0SC8ED/Fmdo4w6
# RpN0ZVPhmJpF95gWmUiIp+1VwyWKlheKGNwn1H1AmtQD+zOPCBgdpPPJNg8kYi1v
# w6zcVRW31ESO+0TNs4fZqNPN+nnEg8Cp4VWVaVB9yVsBtvhtLM/gUgBzfmRgbMwW
# y5OSIpJ9mS1sGx1JXjD8hVSzQrnwqa5djAoyZaGYa7BCyWXmrRi9dBlFSz4vreMv
# auJ8KuKCsAgZ1Ropg582wndbSgNwuj7lftNbzmI30ivhpKeSKDMhu+dVWYijhdKy
# gp/SpJv4CbgXwg39PvVQGltmwAcC65KCRiQskkbBdwyQksTwsPfQsjL6e5gqX1R4
# XlnVN/8xyyQ3hg0iWCjajpz30nnuhBtSk8k010jowOU4KD7xZQHKF9s9Uc4vnB+V
# fr6dy8C+0nKAum9shNmyvjVNrWkvymPaCJI5VckXt3vgjKnzYs+650aPNQV6SyAL
# 7Sk6HIpkSjKpcru3Yz4e0Kep2Pgez73TqnnQaqGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEILO709+XIR0GdbLi4KKdzgsIw0q5XToj3W3IgZRRxb2MAghJcCSdvlE0
# vhgPMjAyNjExMTQwODE3NDJaMAMCAQECEE+KGlvP5EEwwZ2qlgHnjDigghGnMIIE
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
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNDA4MTc0MlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIIeR
# 4ICGXfUdyVVOiwur1XGyqn8Yj0ub28eHkwU1FwyrMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBnU0ZzMs9Hmw9YKnIegLzwxGohyuWMEAdffvL5qzZM/gIgcBKc
# iu1mQ5/QnuHHZWN/ndsTtR8dM18UDVZfi9Jjp44=
# SIG # End signature block
