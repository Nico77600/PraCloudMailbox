#Requires -Version 7.4

<#
.SYNOPSIS
    Renders the graphics of the GitHub README from the PRA Cloud Mailbox developer guide, in a light and a
    dark version: banner, principles, how it works, Convert and Recover, states of a mailbox.

.DESCRIPTION
    GitHub renders Markdown only: the custom blocks of the guide (cards, flow) and its theme are lost. This
    tool renders them as images with the CSS of the built HTML guide and the icons of
    tools\Build-Documentation.ps1, so that the README and the guide always look the same. The README shows
    them with <picture>, which picks the light or dark image from the theme of the reader.

    Sources:
      docs\PraCloudMailbox-Guide.md          the cards block of the introduction, the version
      docs\PraCloudMailbox-Guide.html        the CSS (run tools\Build-Documentation.ps1 first)
      tools\Build-Documentation.ps1          the icons

    Screenshots: Microsoft Edge in headless mode, with a temporary profile, 2x resolution. Only local files
    are opened. Output: docs\images\readme-<name>-light.png and readme-<name>-dark.png.

.PARAMETER OutputFolder
    Default: docs\images next to the tools folder.

.PARAMETER KeepWork
    Keeps the work folder (the HTML pages of the graphics) and shows its path.

.EXAMPLE
    .\tools\Build-Documentation.ps1; .\tools\New-ReadmeImages.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0  (from Web Services Client for Exchange 1.0.0)
    Part of : PRA Cloud Mailbox (repository tool, not in the package)
#>
[CmdletBinding()]
param(
    [string]$OutputFolder,
    [switch]$KeepWork
)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
if (-not $OutputFolder) { $OutputFolder = Join-Path $root 'docs\images' }

#region Assets of the guide ------------------------------------------------------------------------
function ConvertTo-ReadmeInline([string]$Text) {
    # Inline Markdown of a guide block (code, bold, italic) -> HTML.
    $h = [System.Net.WebUtility]::HtmlEncode($Text.Trim())
    $h = [regex]::Replace($h, '`([^`]+)`', '<code>$1</code>')
    $h = [regex]::Replace($h, '\*\*([^*]+)\*\*', '<strong>$1</strong>')
    return [regex]::Replace($h, '(?<![\w*])\*([^*\s][^*]*)\*(?![\w*])', '<em>$1</em>')
}

function Get-ReadmeAssets {
    param([string]$Root)
    $builder = Join-Path $Root 'tools\Build-Documentation.ps1'
    $guideHtml = Join-Path $Root 'docs\PraCloudMailbox-Guide.html'
    $guideMd = Join-Path $Root 'docs\PraCloudMailbox-Guide.md'
    if (-not (Test-Path $guideHtml)) { throw 'docs\PraCloudMailbox-Guide.html not found: run tools\Build-Documentation.ps1 first (it holds the CSS of the graphics).' }
    # Icons: the $Icons table of the documentation builder, read without running the builder.
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($builder, [ref]$null, [ref]$null)
    $assign = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] -and $n.Left.Extent.Text -eq '$Icons' }, $true)
    if (-not $assign) { throw "Icon table not found in $builder." }
    $md = [IO.File]::ReadAllText($guideMd) -replace "`r`n", "`n"
    $blocks = foreach ($m in [regex]::Matches($md, '(?s)```(flow|cards)\n(.*?)\n```')) {
        [pscustomobject]@{ Kind = $m.Groups[1].Value; Lines = @($m.Groups[2].Value -split "`n" | Where-Object { $_.Trim() }) }
    }
    [pscustomobject]@{
        Icons   = & ([scriptblock]::Create($assign.Right.Extent.Text))
        Css     = [regex]::Match([IO.File]::ReadAllText($guideHtml), '(?s)<style>(.*?)</style>').Groups[1].Value
        Version = [regex]::Match($md, '(?m)^version:\s*(\S+)').Groups[1].Value
        Flows   = @($blocks | Where-Object Kind -eq 'flow')
        Cards   = @($blocks | Where-Object Kind -eq 'cards')
    }
}

function Get-ReadmeIcon([string]$Name, [string]$Class = 'icon') {
    $path = $assets.Icons[$Name]; if (-not $path) { $path = $assets.Icons['info'] }
    "<svg class=""$Class"" viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.7"" stroke-linecap=""round"" stroke-linejoin=""round"">$path</svg>"
}

function ConvertTo-ReadmeFlow([string[]]$Lines, [switch]$Vertical) {
    # Vertical: the nodes are stacked, icon on the left, with a downward arrow and its label.
    $items = foreach ($l in $Lines) {
        $icon, $title, $sub = $l.Split('|', 3).ForEach({ $_.Trim() })
        $title = [System.Net.WebUtility]::HtmlEncode($title); $sub = [System.Net.WebUtility]::HtmlEncode($sub)
        if ($Vertical) {
            if ($icon -eq 'arrow') {
                $note = if ($sub) { "<span class=""flow-sub"">$sub</span>" } else { '' }
                "<div class=""rb-varrow""><svg viewBox=""0 0 12 30""><path d=""M6 1v26M1 21l5 6 5-6"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-label"">$title</span>$note</div>"
            } else {
                "<div class=""rb-vnode""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div></div>"
            }
        } elseif ($icon -eq 'arrow') {
            $class = if ($title -or $sub) { 'flow-arrow' } else { 'flow-arrow rb-bare' }
            "<div class=""$class""><span class=""flow-label"">$title</span><svg viewBox=""0 0 40 12""><path d=""M0 6h36M31 1l6 5-6 5"" fill=""none"" stroke=""currentColor"" stroke-width=""1.6""/></svg><span class=""flow-sub"">$sub</span></div>"
        } else {
            "<div class=""flow-node""><div class=""flow-icon"">$(Get-ReadmeIcon $icon)</div><div class=""flow-title"">$title</div><div class=""flow-text"">$sub</div></div>"
        }
    }
    $class = if ($Vertical) { 'flow rb-vflow' } else { 'flow rb-flow' }
    "<div class=""$class"">$($items -join '')</div>"
}

function ConvertTo-ReadmeCards([string[]]$Lines, [string]$Class = '') {
    $items = foreach ($l in $Lines) {
        $icon, $title, $text = $l.Split('|', 3).ForEach({ $_.Trim() })
        "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $icon)</div><div><div class=""card-title"">$(ConvertTo-ReadmeInline $title)</div><div class=""card-text"">$(ConvertTo-ReadmeInline $text)</div></div></div>"
    }
    "<div class=""cards $Class"">$($items -join '')</div>"
}

function Get-ReadmePill([string]$Text, [string]$Tone) { "<span class=""rb-pill"" style=""--tone: var(--cp-$Tone)"">$Text</span>" }
#endregion

#region Styles of the graphics, on top of the CSS of the guide -------------------------------------
$Script:ReadmeCss = @'
html, body { background: #ffffff; }
html[data-theme="dark"], html[data-theme="dark"] body { background: #0d1117; }
:root { --cp-info: #0078d4; --cp-violet: #7c3aed; --cp-teal: #0d9488; }
html[data-theme="dark"] { --cp-info: #4da6ff; --cp-violet: #a78bfa; --cp-teal: #2dd4bf; }
body { display: block; margin: 0; padding: 0; }
.canvas { padding: 6px; }
.rb-pill { display: inline-block; padding: 1px 10px; margin: 8px 6px 0 0; border-radius: 999px; font-size: 11.5px; font-weight: 600; line-height: 1.6;
  color: var(--tone); background: color-mix(in srgb, var(--tone) 11%, transparent); border: 1px solid color-mix(in srgb, var(--tone) 38%, transparent); }
.rb-caption { font-size: 11.5px; font-weight: 700; letter-spacing: 0.1em; text-transform: uppercase; color: var(--cp-accent); margin: 0 0 8px 4px; }
.rb-caption span { color: var(--cp-text-muted); font-weight: 600; letter-spacing: 0.04em; text-transform: none; font-size: 12.5px; }
/* Before / after */
.rb-bench { display: grid; gap: 14px; padding: 20px 22px; border-radius: 16px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-row { display: grid; grid-template-columns: 250px minmax(0, 1fr); gap: 18px; align-items: center; }
.rb-row .label { font-size: 13.5px; font-weight: 650; color: var(--cp-text); } .rb-row .label span { display: block; font-weight: 500; font-size: 12px; color: var(--cp-text-muted); margin-top: 2px; }
.rb-bars { display: grid; gap: 6px; }
.rb-bar { display: flex; align-items: center; gap: 10px; font-size: 12.5px; color: var(--cp-text-muted); }
.rb-bar i { display: block; height: 18px; border-radius: 6px; min-width: 6px; }
.rb-bar.old i { background: color-mix(in srgb, var(--cp-text-muted) 45%, transparent); } .rb-bar.new i { background: var(--cp-accent); }
.rb-bar b { color: var(--cp-text); font-weight: 650; } .rb-bar .bad { color: var(--cp-danger); font-weight: 600; }
.rb-legend { display: flex; gap: 18px; font-size: 12px; color: var(--cp-text-muted); margin: 0 0 4px 4px; } .rb-legend i { display: inline-block; width: 12px; height: 12px; border-radius: 4px; margin-right: 6px; vertical-align: -1px; }
/* Banner */
.rb-hero { margin: 0; padding: 32px 36px 30px; }
.rb-hero-grid { position: relative; display: grid; grid-template-columns: minmax(0, 1fr) 240px; gap: 34px; align-items: center; }
.rb-hero h1 { font-size: 35px; }
.rb-hero .lead { margin: 18px 0 0; font-size: 17px; max-width: none; }
.rb-hero .badges { margin: 20px 0 0; }
.rb-stats { position: relative; display: grid; gap: 10px; }
.rb-stat { display: flex; align-items: center; gap: 14px; padding: 12px 16px; border-radius: 14px; background: var(--cp-panel-strong); border: 1px solid var(--cp-border); box-shadow: 0 1px 2px rgba(0, 0, 0, 0.08); }
.rb-stat b { font-size: 30px; line-height: 1; color: var(--cp-accent); font-weight: 750; min-width: 40px; text-align: center; }
.rb-stat span { font-size: 13px; color: var(--cp-text-muted); line-height: 1.35; }
.rb-stat strong { display: block; color: var(--cp-text); font-size: 14px; }
/* Cards and flows */
.cards { margin: 0; }
.rb-cards2 { grid-template-columns: 1fr 1fr; }
.rb-flow { margin: 0; flex-wrap: nowrap; padding: 18px; gap: 4px; }
.rb-flow .flow-node { flex: 1 1 0; min-width: 0; padding: 14px 10px; }
.rb-flow .flow-title { font-size: 13.5px; overflow-wrap: anywhere; }
.rb-flow .flow-arrow { min-width: 0; width: 84px; flex: 0 0 84px; }
.rb-flow .flow-arrow.rb-bare { width: 46px; flex-basis: 46px; }
.rb-flow .flow-sub { max-width: 84px; }
.rb-space { height: 18px; }
/* How it works: vertical pipeline and the three modes */
.rb-hiw { display: grid; grid-template-columns: minmax(0, 1.08fr) minmax(0, 1fr); gap: 16px; align-items: stretch; }
.rb-col { display: flex; flex-direction: column; }
.rb-vflow { flex: 1; flex-direction: column; flex-wrap: nowrap; align-items: stretch; justify-content: center; gap: 0; margin: 0; padding: 16px 18px; }
.rb-vnode { display: flex; align-items: center; gap: 14px; padding: 11px 16px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); }
.rb-vnode .flow-icon { margin: 0; flex-shrink: 0; }
.rb-vnode .flow-text { margin-top: 1px; }
.rb-varrow { display: flex; align-items: center; gap: 10px; min-height: 36px; padding-left: 31px; }
.rb-varrow svg { width: 12px; height: 28px; color: var(--cp-accent); flex-shrink: 0; }
.rb-varrow .flow-sub { max-width: none; font-size: 12px; }
.rb-modes { flex: 1; display: flex; flex-direction: column; gap: 10px; }
.rb-modes .card-item { flex: 1; align-items: center; }
.rb-modes .card-title { display: flex; align-items: center; gap: 8px; }
.rb-chip { font-size: 11px; font-weight: 600; padding: 0 8px; border-radius: 999px; border: 1px solid var(--cp-border); color: var(--cp-text-muted); }
.rb-chip.hot { color: var(--cp-accent-fg); background: var(--cp-accent); border-color: var(--cp-accent); }
'@
#endregion

#region Rendering (Microsoft Edge, headless) -------------------------------------------------------
function Save-Screenshot([string]$Html, [string]$Png, [int]$Width, [int]$Height, [int]$Scale = 1) {
    $url = 'file:///' + ($Html -replace '\\', '/')
    $profilePath = Join-Path $work 'edge-profile'
    if (Test-Path $Png) { Remove-Item $Png -Force }
    # Start-Process, not &: an Edge helper process can keep the output pipe open after the capture.
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$profilePath`"", "--window-size=$Width,$Height", "--force-device-scale-factor=$Scale", "--screenshot=`"$Png`"", "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden
    $deadline = (Get-Date).AddSeconds(45)
    while (-not (Test-Path $Png) -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 300 }
    if (-not $proc.WaitForExit(10000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    if (-not (Test-Path $Png)) { throw "Screenshot not written: $Png" }
}

function Get-PageHeight([string]$Html, [int]$Width) {
    # Height of the .canvas element: the page writes it in body[data-h], read with --dump-dom.
    $url = 'file:///' + ($Html -replace '\\', '/')
    $dom = Join-Path $work ('dom-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.html')
    $edgeArgs = @('--headless=new', '--disable-gpu', '--hide-scrollbars', '--no-first-run', "--user-data-dir=`"$(Join-Path $work 'edge-profile')`"", "--window-size=$Width,2000", '--dump-dom', "`"$url`"")
    $proc = Start-Process -FilePath $edge -ArgumentList $edgeArgs -PassThru -WindowStyle Hidden -RedirectStandardOutput $dom
    if (-not $proc.WaitForExit(45000)) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue }
    # Edge helper processes inherit the output handle: read in shared mode, retry until written.
    $m = $null
    for ($i = 0; $i -lt 20 -and -not ($m -and $m.Success); $i++) {
        $stream = [IO.File]::Open($dom, 'Open', 'Read', 'ReadWrite')
        try { $text = [IO.StreamReader]::new($stream).ReadToEnd() } finally { $stream.Dispose() }
        $m = [regex]::Match($text, 'data-h="(\d+)"')
        if (-not $m.Success) { Start-Sleep -Milliseconds 250 }
    }
    if (-not $m.Success) { throw "Height not measured: $Html" }
    return [int]$m.Groups[1].Value
}

function New-ReadmeGraphic {
    # One graphic, light and dark: HTML page -> height measured by Edge -> 2x screenshot.
    param([string]$Name, [string]$Body, [int]$Width)
    $pages = @{}
    foreach ($theme in 'light', 'dark') {
        $html = "<!doctype html><html lang=""en"" data-theme=""$theme""><head><meta charset=""utf-8""><style>$($assets.Css)`n$($Script:ReadmeCss)</style></head>" +
            "<body><div class=""canvas"" style=""width:$($Width)px"">$Body</div><script>document.body.setAttribute('data-h', Math.ceil(document.querySelector('.canvas').getBoundingClientRect().height));</script></body></html>"
        $pages[$theme] = Join-Path $work "readme-$Name-$theme.html"
        [IO.File]::WriteAllText($pages[$theme], $html, [Text.UTF8Encoding]::new($false))
    }
    $height = Get-PageHeight $pages['light'] $Width
    foreach ($theme in 'light', 'dark') { Save-Screenshot $pages[$theme] (Join-Path $OutputFolder "readme-$Name-$theme.png") $Width $height 2 }
}
#endregion

#region Main ---------------------------------------------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge not found: it takes the screenshots (headless mode).' }
$work = Join-Path ([IO.Path]::GetTempPath()) ('pra-readme-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work, $OutputFolder -Force | Out-Null
$Script:assets = Get-ReadmeAssets -Root $root
if ($assets.Cards.Count -lt 1) { throw 'The guide must hold the cards block of the introduction.' }
$mid = '&middot;'
$Script:ReadmeCss += @"
/* Convert and Recover: four flows */
.rb-flows { display: grid; gap: 14px; }
.rb-flows .rb-caption { margin-top: 2px; }
/* States of a mailbox */
.rb-states { display: grid; grid-template-columns: 128px repeat(3, minmax(0, 1fr)); gap: 10px; padding: 18px; border-radius: 16px; background: var(--cp-surface-soft); border: 1px dashed var(--cp-border-strong); }
.rb-states .rb-h { font-size: 12px; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase; color: var(--cp-accent); padding: 2px 6px 0; }
.rb-states .rb-k { display: flex; align-items: center; gap: 10px; font-weight: 700; font-size: 14px; color: var(--cp-text); }
.rb-states .rb-k .card-icon { margin: 0; }
.rb-cell { padding: 12px 14px; border-radius: 12px; background: var(--cp-surface); border: 1px solid var(--cp-border); font-size: 12.5px; line-height: 1.45; color: var(--cp-text-muted); }
.rb-cell b { display: block; font-size: 13.5px; color: var(--cp-text); margin-bottom: 2px; }
.rb-cell .rb-pill { margin-top: 8px; }
"@

try {
    Write-Host 'Rendering the README graphics (light and dark, 2x)...'

    # Banner: the hero of the guide, with the key figures.
    $badges = @(
        "<span class=""badge badge-accent"">Version $($assets.Version)</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'terminal' 'icon-sm')PowerShell 5.1 &middot; 7.4+</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'server' 'icon-sm')Exchange 2019 &middot; SE &middot; Online</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'refresh' 'icon-sm')Entra Connect 2.5.76+</span>"
        "<span class=""badge"">$(Get-ReadmeIcon 'tag' 'icon-sm')MIT license</span>"
    ) -join ''
    $banner = "<header class=""hero rb-hero""><div class=""rb-hero-grid""><div>" +
        "<div class=""hero-top""><div class=""hero-logo"">$(Get-ReadmeIcon 'lifebuoy')</div><div><div class=""eyebrow"">Exchange disaster recovery $mid scenario 2 $mid Exchange Online</div><h1>PRA Cloud Mailbox</h1></div></div>" +
        "<p class=""lead"">When <strong>Active Directory, Exchange and Entra Connect</strong> are lost, gives every user and shared mailbox a mailbox in <strong>Exchange Online on its existing identity</strong> &mdash; the Teams storage becomes the mailbox &mdash; and <strong>rolls everything back</strong> when the infrastructure is rebuilt, without losing a mail.</p>" +
        "<div class=""badges"">$badges</div></div>" +
        "<div class=""rb-stats"">" +
        "<div class=""rb-stat""><b>4</b><span><strong>actions</strong>Collect, Check, Convert, Recover</span></div>" +
        "<div class=""rb-stat""><b>1</b><span><strong>licence unit</strong>for all the shared mailboxes</span></div>" +
        "<div class=""rb-stat""><b>0</b><span><strong>mail lost</strong>case hold, inactive mailboxes</span></div>" +
        "</div></div></header>"
    New-ReadmeGraphic -Name 'banner' -Body $banner -Width 1080

    # Principles: the cards block of the introduction of the guide.
    New-ReadmeGraphic -Name 'principles' -Body (ConvertTo-ReadmeCards $assets.Cards[0].Lines 'rb-cards2') -Width 1080

    # How it works: the life of the tool as a vertical pipeline, and the four actions.
    $stages = @(
        'database | Collect | Exchange server, every day: mailboxes, addresses, shared mailbox permissions'
        'arrow | copy | one SQLite file to the cloud admin server'
        'check | Check | is every object ready? Entra ID, Exchange Online, licences, app'
        'arrow | disaster | AD, Exchange and Entra Connect lost'
        'cloud | Convert | identities managed in the cloud, Teams storage = mailbox, shared mailboxes'
        'arrow | batch ID | journal: original state of every object'
        'refresh | Recover | infrastructure rebuilt: back to AD, nothing lost'
    )
    $actions = @(
        [pscustomobject]@{ Icon = 'database'; Name = 'Collect'; Chip = '<span class="rb-chip">Windows PowerShell 5.1</span>'; Text = 'On an Exchange server, every day: <strong>mailboxes, GUIDs, X500 addresses and shared mailbox permissions</strong> into a SQLite snapshot.'; Pills = (Get-ReadmePill 'Reads only' 'success') + (Get-ReadmePill 'Scheduled task' 'teal') }
        [pscustomobject]@{ Icon = 'check'; Name = 'Check'; Chip = '<span class="rb-chip hot">read-only</span>'; Text = 'On the cloud admin server: <strong>is every object ready?</strong> Entra ID object, Exchange Online recipient, licences, permissions of the app.'; Pills = (Get-ReadmePill 'Error = refused' 'warning') + (Get-ReadmePill 'Every week' 'teal') }
        [pscustomobject]@{ Icon = 'cloud'; Name = 'Convert'; Chip = '<span class="rb-chip">-Mode Apply</span>'; Text = 'Users first &mdash; <strong>the Teams storage becomes the mailbox</strong> &mdash; then the shared mailboxes one after another, with their permissions.'; Pills = (Get-ReadmePill 'Batch ID' 'violet') + (Get-ReadmePill 'Resumable' 'success') }
        [pscustomobject]@{ Icon = 'refresh'; Name = 'Recover'; Chip = '<span class="rb-chip">-Batch</span>'; Text = 'Users back on-premises, their cloud data <strong>under a case hold</strong>; shared mailboxes <strong>inactive</strong>, recreated by Entra Connect.'; Pills = (Get-ReadmePill 'Licences as before' 'success') + (Get-ReadmePill 'Waves' 'teal') }
    )
    $actionHtml = ($actions | ForEach-Object { "<div class=""card-item""><div class=""card-icon"">$(Get-ReadmeIcon $_.Icon)</div><div><div class=""card-title"">$($_.Name) $($_.Chip)</div><div class=""card-text"">$($_.Text)</div><div>$($_.Pills)</div></div></div>" }) -join ''
    $howItWorks = "<div class=""rb-hiw""><div class=""rb-col""><div class=""rb-caption"">The life of the tool <span>$mid before, during and after the disaster</span></div>$(ConvertTo-ReadmeFlow $stages -Vertical)</div>" +
        "<div class=""rb-col""><div class=""rb-caption"">Four actions <span>$mid preview first, one confirmation</span></div><div class=""rb-modes"">$actionHtml</div></div></div>"
    New-ReadmeGraphic -Name 'how-it-works' -Body $howItWorks -Width 1080

    # Convert and Recover, users and shared mailboxes: the validated order of the steps.
    $flows = [ordered]@{
        'Users &middot; Convert <span>&middot; minutes</span>' = @(
            'mail | GUID cleared | the MailUser no longer points on-premises', 'arrow | | ', 'cloud | Cloud authority | Entra Connect ignores the object', 'arrow | | '
            'key | Exchange plan | licence group, Kiosk or direct', 'arrow | | ', 'check | Cloud mailbox | the Teams storage, chats kept')
        'Users &middot; Recover <span>&middot; once AD, Exchange and Entra Connect are rebuilt</span>' = @(
            'refresh | AD authority | one delta cycle', 'arrow | | ', 'key | Plan removed | licences exactly as before', 'arrow | | '
            'server | On-premises | MailUser, on-premises GUID', 'arrow | | ', 'shield | Case hold | disaster mails kept')
        'Shared mailboxes &middot; Convert <span>&middot; one after another</span>' = @(
            'key | Temporary licence | one unit for all', 'arrow | | ', 'people | SharedMailbox | then the licence goes back', 'arrow | | '
            'tag | Tag | CustomAttribute1 = Converted', 'arrow | | ', 'check | Permissions | FullAccess, SendAs, SendOnBehalf')
        'Shared mailboxes &middot; Recover <span>&middot; one scheduler pause, one delta cycle</span>' = @(
            'shield | Case hold | stamped before anything', 'arrow | | ', 'warning | Identity deleted | scheduler paused, purged', 'arrow | | '
            'folder | Inactive mailbox | content, hold and tag kept', 'arrow | | ', 'refresh | New object | Entra Connect, on-premises GUID')
    }
    $flowHtml = ($flows.Keys | ForEach-Object { "<div><div class=""rb-caption"">$_</div>$(ConvertTo-ReadmeFlow ([string[]]$flows[$_]))</div>" }) -join ''
    New-ReadmeGraphic -Name 'flows' -Body "<div class=""rb-flows"">$flowHtml</div>" -Width 1080

    # States of a mailbox: before the disaster, after Convert, after Recover.
    $onPrem = Get-ReadmePill 'Mail on-premises' 'info'; $online = Get-ReadmePill 'Mail in Exchange Online' 'violet'
    $cell = { param($Title, $Text, $Pill) "<div class=""rb-cell""><b>$Title</b>$Text<br>$Pill</div>" }
    $states = "<div class=""rb-states"">" +
        "<div></div><div class=""rb-h"">Before the disaster</div><div class=""rb-h"">After Convert</div><div class=""rb-h"">After Recover</div>" +
        "<div class=""rb-k""><div class=""card-icon"">$(Get-ReadmeIcon 'user')</div>User</div>" +
        (& $cell 'Mailbox on-premises' 'Entra ID synchronised from AD; Exchange Online: MailUser + Teams storage (ComponentShared)' $onPrem) +
        (& $cell 'Cloud mailbox' 'Entra ID managed in the cloud; the Teams storage became the mailbox (same identity, same addresses)' $online) +
        (& $cell 'Mailbox on-premises again' 'Synchronised from AD; the cloud mailbox is a ComponentShared again, <strong>under the case hold</strong>' $onPrem) +
        "<div class=""rb-k""><div class=""card-icon"">$(Get-ReadmeIcon 'people')</div>Shared</div>" +
        (& $cell 'Shared mailbox on-premises' 'Entra ID synchronised from AD; Exchange Online: MailUser' $onPrem) +
        (& $cell 'Cloud shared mailbox' 'Managed in the cloud, tagged, the permissions of the snapshot granted' $online) +
        (& $cell 'Shared mailbox on-premises again' 'New object created by Entra Connect; the cloud mailbox is <strong>inactive</strong>, under the hold, then under retention' $onPrem) +
        "</div>"
    New-ReadmeGraphic -Name 'states' -Body $states -Width 1080

    Get-ChildItem $OutputFolder -Filter 'readme-*.png' | Select-Object Name, @{ n = 'KB'; e = { [math]::Round($_.Length / 1KB) } } | Format-Table -AutoSize | Out-String | Write-Host
} finally {
    # Edge helper processes of the temporary profile, if any are left.
    Get-CimInstance Win32_Process -Filter "Name='msedge.exe'" | Where-Object { $_.CommandLine -like "*$work*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
    if ($KeepWork) { Write-Host "Work folder: $work" } else { Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue }
}
#endregion