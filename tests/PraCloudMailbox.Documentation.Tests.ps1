<#
.SYNOPSIS
    Documentation tests for PRA Cloud Mailbox (offline; reads the repository files only: no Exchange, Graph or network).
.DESCRIPTION
    Keeps the README, the user guide and the administrator guide honest against the code:
      - every picture exists, in a light and a dark version, has an alternative text, is in the integrity inventory, and none is left unused;
      - no real machine name, user name, tenant or local path appears in the documents;
      - the links between the documents resolve, the versions agree;
      - every command-line parameter, configuration setting, Microsoft Graph permission and Entra role that a document names is the one of the code;
      - every message that a document quotes (in italics) exists in the code;
      - every button of the console that the user guide names exists in the console;
      - the snapshot protection is written in every document that the operator reads;
      - the guides keep short sentences and common words (a simple readability measure).
.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
#>
#Requires -Version 7.5

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Repo = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $script:Package = Join-Path $script:Repo 'package'
    $script:Docs = Join-Path $script:Package 'docs'
    $script:Images = Join-Path $script:Docs 'images'
    function script:Read-Doc([string]$Relative) { [IO.File]::ReadAllText((Join-Path $script:Repo $Relative)) -replace "`r`n", "`n" }
    $script:Readme = Read-Doc 'README.md'
    $script:PackageReadme = Read-Doc 'package\README.md'
    $script:UserGuide = Read-Doc 'package\docs\PraCloudMailbox-UserGuide.md'
    $script:Guide = Read-Doc 'package\docs\PraCloudMailbox-Guide.md'
    $script:Entry = Read-Doc 'package\Invoke-PraCloudMailbox.ps1'
    $script:Cloud = Read-Doc 'package\module\PRA.Cloud.psm1'
    $script:Collect = Read-Doc 'package\module\PRA.Collect.psm1'
    $script:Documents = [ordered]@{ 'README.md' = $script:Readme; 'package\README.md' = $script:PackageReadme; 'UserGuide' = $script:UserGuide; 'Guide' = $script:Guide }
    # Everything the operator can read in the tool: the code, the window and the configuration.
    $codeFiles = @(Get-ChildItem -LiteralPath $script:Package -Recurse -File | Where-Object { $_.Extension -in '.ps1', '.psm1', '.psd1', '.xaml' -and $_.Name -notin 'PackageIntegrity.psd1' })
    $script:CodeText = ($codeFiles | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
    $script:ConsoleText = (@($codeFiles | Where-Object { $_.Name -like 'PRA.Gui*' }) | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
    function script:Get-ImageReferences([string]$Text, [string]$Base) {
        # Markdown images and the src / srcset of HTML pictures, resolved against the folder of the document.
        $found = New-Object 'Collections.Generic.List[string]'
        foreach ($m in [regex]::Matches($Text, '!\[[^\]]*\]\(([^)\s]+)\)')) { $found.Add($m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($Text, '(?:src|srcset)="([^"]+\.png)(?:\?[^"]*)?"')) { $found.Add($m.Groups[1].Value) }
        foreach ($f in $found) { [IO.Path]::GetFullPath((Join-Path $Base ($f -replace '/', '\'))) }
    }
    function script:Get-Slug([string]$Heading) {
        $s = $Heading.ToLowerInvariant() -replace '[^\p{L}\p{N}\s_-]', ''
        return ($s.Trim() -replace '\s', '-')
    }
    function script:Remove-CodeBlocks([string]$Markdown) { [regex]::Replace($Markdown, '(?s)```.*?```', '') }
    function script:Get-ProseSentences([string]$Markdown, [bool]$Tables) {
        $text = $Markdown -replace "`r`n", "`n"
        $text = [regex]::Replace($text, '(?s)^---\n.*?\n---\n', '')
        $text = [regex]::Replace($text, '(?s)```.*?```', '')
        $lines = foreach ($l in ($text -split "`n")) {
            $x = $l.Trim()
            if (-not $x) { continue }
            if ($x.StartsWith('|')) {
                if (-not $Tables -or $x -match '^\|[-| :]+\|$') { continue }
                $x = (($x.Trim('|') -split '\|' | ForEach-Object { $_.Trim() }) -join '. ')
            }
            if ($x.StartsWith('#') -or $x.StartsWith('!') -or $x.StartsWith('<!--')) { continue }
            $x = $x -replace '^(>\s*)+', '' -replace '^\[!\w+\]\s*$', '' -replace '^[-*]\s+(\[[ x]\]\s+)?', '' -replace '^\d+\.\s+', ''
            $x = [regex]::Replace($x, '\[([^\]]*)\]\([^)]*\)', '$1')
            $x = [regex]::Replace($x, '`[^`]*`', 'CODE')
            $x = $x -replace '\*\*|__|\*', ''
            if ($x.Trim()) { $x }
        }
        foreach ($l in $lines) { foreach ($s in [regex]::Split($l, '(?<=[.!?:])\s+')) { $s = $s.Trim(); if ($s -match '\w') { $s } } }
    }
    function script:Get-Readability([string]$Markdown, [bool]$Tables = $false) {
        $sentences = @(Get-ProseSentences $Markdown $Tables)
        $counts = @($sentences | ForEach-Object { @($_ -split '\s+' | Where-Object { $_ -match '\w' }).Count })
        $words = @($sentences | ForEach-Object { $_ -split '\s+' | Where-Object { $_ -match '\w' } })
        $syllables = 0
        foreach ($w in $words) {
            $x = $w.ToLowerInvariant() -replace '[^a-z]', ''
            if ($x.Length -le 3) { $syllables++; continue }
            $x = $x -replace '(?:[^laeiouy]es|ed|[^laeiouy]e)$', '' -replace '^y', ''
            $syllables += [Math]::Max(1, ([regex]::Matches($x, '[aeiouy]{1,2}')).Count)
        }
        [pscustomobject]@{
            Sentences = $sentences.Count; Words = $words.Count; Average = ($counts | Measure-Object -Average).Average
            Over20 = @($counts | Where-Object { $_ -gt 20 }).Count; Over25 = @($counts | Where-Object { $_ -gt 25 }).Count
            FleschReadingEase = 206.835 - 1.015 * ($words.Count / $sentences.Count) - 84.6 * ($syllables / $words.Count)
        }
    }
    function script:Get-ParameterNames {
        $ast = [Management.Automation.Language.Parser]::ParseInput($script:Entry, [ref]$null, [ref]$null)
        $names = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Name.VariablePath.UserPath })
        $aliases = @($ast.ParamBlock.Parameters | ForEach-Object { $_.Attributes | Where-Object { $_.TypeName.Name -eq 'Alias' } | ForEach-Object { $_.PositionalArguments | ForEach-Object { $_.Value } } })
        $names + $aliases + @('Confirm', 'WhatIf', 'Verbose', 'Debug', 'ErrorAction')
    }
}

AfterAll {
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'Pictures' {
    It 'every picture used by the README and the guides exists' {
        $missing = foreach ($entry in @(
                @{ Text = $script:Readme; Base = $script:Repo }, @{ Text = $script:PackageReadme; Base = $script:Package }
                @{ Text = $script:UserGuide; Base = $script:Docs }, @{ Text = $script:Guide; Base = $script:Docs })) {
            foreach ($path in @(Get-ImageReferences $entry.Text $entry.Base)) { if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { $path } }
        }
        @($missing) | Should -BeNullOrEmpty
    }
    It 'every picture exists in a light and a dark version' {
        $names = @(Get-ChildItem -LiteralPath $script:Images -Filter '*.png' | ForEach-Object { $_.BaseName })
        $lonely = foreach ($name in $names) {
            if ($name -match '^(.+)-light$' -and ('{0}-dark' -f $Matches[1]) -notin $names) { "$name has no dark twin" }
            elseif ($name -match '^(.+)-dark$' -and ('{0}-light' -f $Matches[1]) -notin $names) { "$name has no light twin" }
            elseif ($name -notmatch '-(light|dark)$') { "$name is neither light nor dark" }
        }
        @($lonely) | Should -BeNullOrEmpty
    }
    It 'no picture of docs\images is left unused (the twin of a picture that is used counts as used)' {
        $used = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($entry in @(@{ Text = $script:Readme; Base = $script:Repo }, @{ Text = $script:UserGuide; Base = $script:Docs }, @{ Text = $script:Guide; Base = $script:Docs })) {
            foreach ($path in @(Get-ImageReferences $entry.Text $entry.Base)) {
                [void]$used.Add($path)
                [void]$used.Add(($path -replace '-light\.png$', '-dark.png' -replace '(?<=-)dark\.png$', 'dark.png'))
                [void]$used.Add(($path -replace '-dark\.png$', '-light.png'))
            }
        }
        $orphans = @(Get-ChildItem -LiteralPath $script:Images -Filter '*.png' | Where-Object { -not $used.Contains($_.FullName) } | ForEach-Object Name)
        $orphans | Should -BeNullOrEmpty
    }
    It 'the pictures of docs\images are exactly the ones of the integrity inventory (verifier, fixture, manifest) and of .gitattributes' {
        $onDisk = @(Get-ChildItem -LiteralPath $script:Images -Filter '*.png' | ForEach-Object { 'docs\images\' + $_.Name } | Sort-Object { $_ } -CaseSensitive)
        $fixture = @(Get-PcmExpectedInventory | Where-Object { $_ -like 'docs\images\*' } | Sort-Object { $_ } -CaseSensitive)
        $verifier = @([regex]::Matches((Read-Doc 'package\module\Test-PackageIntegrity.ps1'), "(?m)^\s*'(docs\\images\\[^']+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object { $_ } -CaseSensitive)
        $manifest = @([regex]::Matches((Read-Doc 'package\PackageIntegrity.psd1'), "Path = '(docs\\images\\[^']+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object { $_ } -CaseSensitive)
        $attributes = @([regex]::Matches((Read-Doc '.gitattributes'), '(?m)^package/docs/images/(\S+\.png) -text') | ForEach-Object { 'docs\images\' + $_.Groups[1].Value } | Sort-Object { $_ } -CaseSensitive)
        ($fixture -join '|') | Should -BeExactly ($onDisk -join '|')
        ($verifier -join '|') | Should -BeExactly ($onDisk -join '|')
        ($manifest -join '|') | Should -BeExactly ($onDisk -join '|')
        ($attributes -join '|') | Should -BeExactly ($onDisk -join '|')
    }
    It 'every picture of a document has an alternative text' {
        foreach ($name in $script:Documents.Keys) {
            $text = $script:Documents[$name]
            @([regex]::Matches($text, '!\[\s*\]\(')).Count | Should -Be 0 -Because "$name has a markdown picture without alternative text"
            foreach ($m in [regex]::Matches($text, '<img\b[^>]*>')) { $m.Value | Should -Match 'alt="[^"]{8,}"' -Because "$name has an <img> without alternative text" }
        }
    }
    It 'every light picture of the README picture elements has its dark twin' {
        foreach ($m in [regex]::Matches($script:Readme, 'srcset="(package/docs/images/[^"]+)-dark\.png"')) {
            Test-Path -LiteralPath (Join-Path $script:Repo (($m.Groups[1].Value + '-light.png') -replace '/', '\')) | Should -BeTrue -Because $m.Groups[1].Value
        }
    }
    It 'nothing refers to the pictures of the former windows (pra-gui-*, pra-console-*, pra-report-*)' {
        foreach ($name in $script:Documents.Keys) { $script:Documents[$name] | Should -Not -Match 'pra-(gui|console|report)-' -Because $name }
        foreach ($path in 'package\module\Test-PackageIntegrity.ps1', 'tests\IntegrityFixture.ps1', '.gitattributes', 'package\PackageIntegrity.psd1') { (Read-Doc $path) | Should -Not -Match 'pra-(gui|console|report)-' -Because $path }
        @(Get-ChildItem -LiteralPath $script:Images -Filter 'pra-*.png').Count | Should -Be 0
    }
    It 'the user guide shows every page of the console and every task' {
        foreach ($name in 'console-setup', 'console-collect-task', 'console-task-convert-one', 'console-task-convert-group', 'console-task-convert-ou', 'console-task-convert-csv', 'console-task-continue', 'console-task-rollback', 'console-task-rollback-wave', 'console-task-check', 'console-verify-before', 'console-verify-preview', 'console-verify-confirmed', 'console-run', 'console-run-done', 'console-next', 'console-next-failed', 'cli-check', 'cli-convert', 'cli-follow', 'cli-recover', 'report-check', 'report-convert', 'report-global') {
            $script:UserGuide | Should -Match ([regex]::Escape("images/$name-light.png")) -Because "the user guide must show $name"
        }
    }
}

Describe 'No private name in the documents' {
    It 'no real machine, user name, tenant or local path in <Name>' -ForEach @(
        @{ Name = 'README'; Path = 'README.md' }, @{ Name = 'package README'; Path = 'package\README.md' }, @{ Name = 'user guide'; Path = 'package\docs\PraCloudMailbox-UserGuide.md' }
        @{ Name = 'administrator guide'; Path = 'package\docs\PraCloudMailbox-Guide.md' }, @{ Name = 'CHANGELOG'; Path = 'CHANGELOG.md' }
        @{ Name = 'user guide HTML'; Path = 'package\docs\PraCloudMailbox-UserGuide.html' }, @{ Name = 'administrator guide HTML'; Path = 'package\docs\PraCloudMailbox-Guide.html' }
    ) {
        $text = Read-Doc $Path
        $text = [regex]::Replace($text, 'data:image/png;base64,[A-Za-z0-9+/=]+', 'data:image')
        $text | Should -Not -Match '(?i)nfabert|REDMOND|DESKTOP-[A-Z0-9]|C:\\Users\\|\\AppData\\|scratchpad|Temp\\claude'
    }
    It 'the guides name synthetic computers, accounts and domains only' {
        foreach ($name in 'UserGuide', 'Guide') {
            $text = $script:Documents[$name]
            foreach ($m in [regex]::Matches($text, '(?i)\b[a-z0-9-]+\.(?:onmicrosoft\.com|com|local|net|org)\b')) {
                $m.Value | Should -Match '(?i)^(?:[a-z0-9-]+\.)?(?:contoso\.com|contoso\.onmicrosoft\.com|microsoft\.com|office365\.com|microsoftonline\.com|github\.com|learn\.microsoft\.com|office\.com)$|^(?:login|outlook)\.' -Because "$name names a domain that is not synthetic: $($m.Value)"
            }
        }
    }
}

Describe 'Versions and links' {
    It 'the two guides carry the version of the tool' {
        $version = [regex]::Match((Read-Doc 'package\PraCloudMailbox.psd1'), "(?m)^\s*ModuleVersion\s*=\s*'(\d+\.\d+\.\d+)'").Groups[1].Value
        $version | Should -Match '^\d+\.\d+\.\d+$'
        foreach ($guide in $script:UserGuide, $script:Guide) { [regex]::Match($guide, '(?m)^version:\s*(\S+)').Groups[1].Value | Should -BeExactly $version }
        $script:Readme | Should -Match ([regex]::Escape("**Version:** $version"))
        (Read-Doc 'CHANGELOG.md') | Should -Match ('(?m)^## ' + [regex]::Escape($version) + ' ')
    }
    It 'the relative links between the documents resolve, with their heading anchor' {
        $problems = foreach ($entry in @(
                @{ Name = 'README.md'; Text = $script:Readme; Base = $script:Repo }, @{ Name = 'user guide'; Text = $script:UserGuide; Base = $script:Docs }
                @{ Name = 'administrator guide'; Text = $script:Guide; Base = $script:Docs }, @{ Name = 'package README'; Text = $script:PackageReadme; Base = $script:Package })) {
            foreach ($m in [regex]::Matches($entry.Text, '(?<!\!)\[[^\]]+\]\((?!https?:|mailto:)([^)#\s]*)(?:#([^)\s]*))?\)')) {
                $file = if ($m.Groups[1].Value) { [IO.Path]::GetFullPath((Join-Path $entry.Base ($m.Groups[1].Value -replace '/', '\'))) } else { $null }
                if ($file -and -not (Test-Path -LiteralPath $file)) { "$($entry.Name): missing file $($m.Groups[1].Value)"; continue }
                $anchor = $m.Groups[2].Value
                if ($anchor) {
                    $target = if ($file) { Read-Doc ($file.Substring($script:Repo.Length + 1)) } else { $entry.Text }
                    $slugs = @([regex]::Matches((Remove-CodeBlocks $target), '(?m)^#{1,6}\s+(.+)$') | ForEach-Object { Get-Slug $_.Groups[1].Value })
                    $stripped = @($slugs | ForEach-Object { $_ -replace '^\d+-', '' -replace '-{2,}', '-' })
                    if (($anchor -notin $slugs) -and ($anchor -notin $stripped)) { "$($entry.Name): no heading for #$anchor in $($m.Groups[1].Value)" }
                }
            }
        }
        @($problems) | Should -BeNullOrEmpty
    }
    It 'the README, the package README and the administrator guide point to the user guide, and the user guide to the administrator guide' {
        $script:Readme | Should -Match 'PraCloudMailbox-UserGuide\.md'
        $script:PackageReadme | Should -Not -BeNullOrEmpty
        $script:Guide | Should -Match 'PraCloudMailbox-UserGuide\.md'
        $script:UserGuide | Should -Match 'PraCloudMailbox-Guide\.md'
        $script:UserGuide | Should -Match 'PraCloudMailbox-Guide\.html'
    }
    It 'every link of the two HTML guides leads somewhere: the file is next to it and the anchor is an id of the target page' {
        function Read-GuideHtml([string]$Name) { [regex]::Replace((Read-Doc "package\docs\$Name.html"), 'data:image/png;base64,[A-Za-z0-9+/=]+', 'data:image') }
        $html = [ordered]@{ 'PraCloudMailbox-UserGuide' = (Read-GuideHtml 'PraCloudMailbox-UserGuide'); 'PraCloudMailbox-Guide' = (Read-GuideHtml 'PraCloudMailbox-Guide') }
        $problems = foreach ($name in $html.Keys) {
            foreach ($m in [regex]::Matches($html[$name], '<a [^>]*href="([^"]*)"')) {
                $href = $m.Groups[1].Value
                if ($href -match '^(https?:|mailto:)') { continue }
                $file, $anchor = $href -split '#', 2
                $target = if ($file) { Join-Path $script:Docs $file } else { Join-Path $script:Docs "$name.html" }
                if (-not (Test-Path -LiteralPath $target -PathType Leaf)) { "${name}: missing file $href"; continue }
                if ($anchor) {
                    $page = if ($file) { Read-GuideHtml ([IO.Path]::GetFileNameWithoutExtension($file)) } else { $html[$name] }
                    if ($page -notmatch ('\sid="' + [regex]::Escape($anchor) + '"')) { "${name}: no id for $href" }
                }
            }
            @([regex]::Matches($html[$name], '<img\b(?![^>]*\ssrc="[^"]+")[^>]*>')).Count | Should -Be 0 -Because "$name has an <img> without src"
        }
        @($problems) | Should -BeNullOrEmpty
    }
    It 'every mention of a chapter or an annex is a link, in both guides' {
        foreach ($entry in @(@{ Name = 'user guide'; Text = $script:UserGuide }, @{ Name = 'administrator guide'; Text = $script:Guide })) {
            $text = [regex]::Replace($entry.Text, '(?s)^---\n.*?\n---\n', '')
            $text = Remove-CodeBlocks $text
            $text = [regex]::Replace($text, '(?m)^#{1,6}\s.*$', '')
            $text = [regex]::Replace($text, '\[[^\]]*\]\([^)]*\)', '')
            $text = [regex]::Replace($text, '`[^`]*`', '')
            # In the user guide a chapter of the same guide is named by its number (the table of contents says where it is); only the other guide must be linked.
            if ($entry.Name -eq 'user guide') { @([regex]::Matches($text, '(?i)\b(?:administrator guide,? chapter \d+|appendix [A-F])\b') | ForEach-Object { $_.Value }) | Should -BeNullOrEmpty -Because 'the user guide names a chapter of the other guide without a link' }
        }
    }
    It 'the guides are in the integrity inventory, with their HTML' {
        $verifier = Read-Doc 'package\module\Test-PackageIntegrity.ps1'
        foreach ($name in 'docs\PraCloudMailbox-Guide.md', 'docs\PraCloudMailbox-Guide.html', 'docs\PraCloudMailbox-UserGuide.md', 'docs\PraCloudMailbox-UserGuide.html', 'README.md') {
            $verifier | Should -Match ([regex]::Escape("'$name'"))
            (Read-Doc 'package\PackageIntegrity.psd1') | Should -Match ([regex]::Escape("Path = '$name'"))
            @(Get-PcmExpectedInventory) | Should -Contain $name
        }
    }
    It 'the HTML guides were built from the current Markdown (the user guide title and the last chapter are in the HTML)' {
        foreach ($name in 'PraCloudMailbox-UserGuide', 'PraCloudMailbox-Guide') {
            $html = [regex]::Replace((Read-Doc "package\docs\$name.html"), 'data:image/png;base64,[A-Za-z0-9+/=]+', 'data:image')
            $md = Read-Doc "package\docs\$name.md"
            $html | Should -Match 'PRA Cloud Mailbox'
            foreach ($heading in [regex]::Matches($md, '(?m)^## (.+)$') | Select-Object -Last 2) {
                $title = $heading.Groups[1].Value.Trim() -replace '^\d+\.\s+', ''
                $html | Should -Match ([regex]::Escape([Net.WebUtility]::HtmlEncode($title))) -Because "$name.html must hold the chapter '$title'"
            }
        }
    }
}

Describe 'The snapshot is protected: the rule is written everywhere the operator reads' {
    It 'every operator document has a Protect the snapshot part with the four rules' -ForEach @(
        @{ Name = 'README'; Key = 'Readme' }, @{ Name = 'package README'; Key = 'PackageReadme' }, @{ Name = 'user guide'; Key = 'UserGuide' }, @{ Name = 'administrator guide'; Key = 'Guide' }
    ) {
        $text = Get-Variable -Scope Script -Name $Key -ValueOnly
        $text | Should -Match '(?i)protect the snapshot' -Because $Name
        $text | Should -Match '(?i)only copy' -Because "$Name must say that the snapshot is the only copy"
        $text | Should -Match '(?i)off-site' -Because "$Name must ask for an external backup"
        $text | Should -Match '(?i)(Domain Users|access list)' -Because "$Name must say who may read it"
        $text | Should -Match '\.sha256' -Because "$Name must say to keep the .sha256 file"
        $text | Should -Match '(?i)journal' -Because "$Name must say that the journal is not covered by the hash"
        $text | Should -Match '(?i)-AcceptUnverifiedSnapshot' -Because $Name
    }
    It 'the user guide and the administrator guide say what to do when the file changed' {
        $script:UserGuide | Should -Match 'Restore the file from your external backup'
        $script:Guide | Should -Match 'restore the snapshot from the external backup'
        $script:CodeText | Should -Match ([regex]::Escape('Restore the file from your external backup'))
    }
}

Describe 'What the documents quote follows the code' {
    It 'every command-line parameter of the entry script that a document writes exists' {
        $known = Get-ParameterNames
        $problems = foreach ($name in $script:Documents.Keys) {
            foreach ($m in [regex]::Matches($script:Documents[$name], '(?m)Invoke-PraCloudMailbox\.ps1[^\n`"]*')) {
                foreach ($option in [regex]::Matches($m.Value, '(?<![\w-])-([A-Za-z]+)\b')) {
                    if ($option.Groups[1].Value -notin $known) { "$name`: -$($option.Groups[1].Value) is not a parameter of the entry script" }
                }
            }
        }
        @($problems) | Should -BeNullOrEmpty
    }
    It 'every parameter written in code (`-Name`) next to a tool option exists, and the parameters of the next-step builder are documented' {
        $known = Get-ParameterNames
        foreach ($option in '-Identity', '-CsvPath', '-IdentityPath', '-SearchBase', '-Group', '-MaxObjects', '-Scope', '-Batch', '-Expect', '-Once', '-ResolveOnly', '-AcceptUnverifiedSnapshot', '-BaseConfigPath', '-ConfigPath', '-Snapshot') {
            $script:Guide | Should -Match ([regex]::Escape($option)) -Because "the administrator guide must document $option"
            $option.TrimStart('-') | Should -BeIn $known
        }
        foreach ($option in '-Identity', '-CsvPath', '-SearchBase', '-Group', '-MaxObjects', '-Expect', '-Once', '-AcceptUnverifiedSnapshot') { $script:UserGuide | Should -Match ([regex]::Escape($option)) -Because "the user guide must show $option" }
    }
    It 'every configuration setting that a document names (Section.Key) is a setting of the configuration file' {
        $config = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Package 'config\PraCloudMailbox.config.psd1')
        $flat = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        function Add-Flat($Prefix, $Value) { foreach ($key in $Value.Keys) { [void]$flat.Add("$Prefix$key"); if ($Value[$key] -is [hashtable]) { Add-Flat "$Prefix$key." $Value[$key] } } }
        Add-Flat '' $config
        $problems = foreach ($name in $script:Documents.Keys) {
            foreach ($m in [regex]::Matches((Remove-CodeBlocks $script:Documents[$name]), '`((?:Exchange|Scope|Collect|Store|Cloud|Exo|Polling|Licensing|Recover|Resources|Retention|EntraConnect|Logging|Report)\.[A-Za-z]+(?:\.[A-Za-z]+)?)`')) {
                if (-not $flat.Contains($m.Groups[1].Value) -and $m.Groups[1].Value -notmatch '\.\*$' -and $m.Groups[1].Value -notin 'Exchange.ManageAsApp', 'Exchange.ConnectionMode') { "$name`: $($m.Groups[1].Value)" }
            }
        }
        @($problems | Sort-Object -Unique) | Should -BeNullOrEmpty
    }
    It 'every Graph permission named in a document is one of the code, or Exchange.ManageAsApp' {
        $code = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($m in [regex]::Matches($script:Cloud, "Role = '([A-Za-z-]+(?:\.[A-Za-z]+)+)'")) { [void]$code.Add($m.Groups[1].Value) }
        foreach ($m in [regex]::Matches($script:Cloud, "GroupReadRoles = @\(([^)]*)\)")) { foreach ($r in [regex]::Matches($m.Groups[1].Value, "'([^']+)'")) { [void]$code.Add($r.Groups[1].Value) } }
        [void]$code.Add('GroupMember.Read.All'); [void]$code.Add('Exchange.ManageAsApp')
        $code.Count | Should -BeGreaterThan 8
        foreach ($name in $script:Documents.Keys) {
            foreach ($m in [regex]::Matches($script:Documents[$name], '\b[A-Z][A-Za-z-]+\.(?:Read|ReadWrite|ReadBasic|ManageAsApp)(?:\.[A-Za-z]+)?\b')) {
                $code.Contains($m.Value) | Should -BeTrue -Because "$name names $($m.Value), which the code does not use"
            }
        }
    }
    It 'the minimum permissions chapter names every permission, role and cmdlet that the code needs' {
        $start = $script:Guide.IndexOf("`n### 4.5 Minimum permissions`n"); $end = $script:Guide.IndexOf("`n## 5. Installation")
        $start | Should -BeGreaterThan 0; $end | Should -BeGreaterThan $start
        $chapter = $script:Guide.Substring($start, $end - $start)
        $names = New-Object 'Collections.Generic.HashSet[string]'
        foreach ($m in [regex]::Matches($script:Cloud, "Role = '([A-Za-z-]+(?:\.[A-Za-z]+)+)'")) { [void]$names.Add($m.Groups[1].Value) }
        foreach ($role in [regex]::Matches($script:Cloud, "AnyOf = @\(([^)]*)\)") | ForEach-Object { [regex]::Matches($_.Groups[1].Value, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value } }) { [void]$names.Add($role) }
        $collectList = [regex]::Match($script:Collect, '(?s)\$script:ExchangeCommands = @\((.*?)\)').Groups[1].Value
        foreach ($m in [regex]::Matches($collectList, "'([^']+)'")) { [void]$names.Add($m.Groups[1].Value) }
        $baseList = [regex]::Match($script:Cloud, "(?s)function Get-PcmExoCommandName.*?\`$commands = @\((.*?)\)\s*\r?\n\s*if").Groups[1].Value
        foreach ($m in [regex]::Matches($baseList, "'([^']+)'")) { [void]$names.Add($m.Groups[1].Value) }
        foreach ($n in 'Get-CaseHoldPolicy', 'Set-CaseHoldPolicy', 'New-CaseHoldPolicy', 'New-CaseHoldRule', 'Get-ADSyncScheduler', 'Set-ADSyncScheduler', 'Start-ADSyncSyncCycle', 'Exchange.ManageAsApp', 'Exchange Administrator', 'eDiscovery Manager') { [void]$names.Add($n) }
        $names.Count | Should -BeGreaterThan 35
        # Set-ADServerSettings and Get-OrganizationConfig are named through their groups of cmdlets in the sentence of the Collect account and the table of Check.
        $missing = @($names | Where-Object { -not $chapter.Contains($_) })
        $missing | Should -BeNullOrEmpty
    }
    It 'every cloud and directory cmdlet that the code calls and a document names is a real cmdlet of the code' {
        $cmdlets = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach ($m in [regex]::Matches($script:CodeText, '\b(?:Get|Set|New|Add|Remove|Start|Connect|Disconnect|Invoke|Import)-[A-Za-z]+\b')) { [void]$cmdlets.Add($m.Value) }
        foreach ($name in 'UserGuide', 'Guide') {
            foreach ($m in [regex]::Matches((Remove-CodeBlocks $script:Documents[$name]), '`((?:Get|Set|New|Add|Remove|Start)-(?:Mailbox|MailUser|Recipient|User|OrganizationConfig|MailboxPermission|RecipientPermission|MailContact|CaseHold\w+|ADSync\w+|ADPermission|DistributionGroup\w*|ComplianceCase|AdaptiveScope|RetentionCompliance\w+|ServicePrincipal|RoleGroupMember|MgGraph\w*)\b)`')) {
                if ($m.Groups[1].Value -in 'New-ComplianceCase', 'New-AdaptiveScope', 'New-RetentionCompliancePolicy', 'New-RetentionComplianceRule', 'New-ServicePrincipal', 'Add-RoleGroupMember') { continue }
                $cmdlets.Contains($m.Groups[1].Value) | Should -BeTrue -Because "$name names $($m.Groups[1].Value), which the code does not call"
            }
        }
    }
    It 'every Microsoft Entra role that the code asks for is named in the administrator guide' {
        $roles = @([regex]::Matches($script:Cloud, "AnyOf = @\(([^)]*)\)") | ForEach-Object { [regex]::Matches($_.Groups[1].Value, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value } } | Sort-Object -Unique)
        $roles.Count | Should -BeGreaterThan 4
        foreach ($role in $roles) { $script:Guide | Should -Match ([regex]::Escape($role)) }
    }
    It 'the messages that the guides quote in italics exist in the code' {
        # A quoted message is written between asterisks; '…' (or '...') stands for a name or a number that changes.
        $exceptions = @('mail user', "Plan de Reprise d'Activité", 'View-Only Organization Management', 'Active Directory Permissions', 'Organization Management', 'App registrations', 'New registration', 'Grant admin consent', 'Enterprise application')
        $problems = foreach ($name in 'UserGuide', 'Guide') {
            $text = Remove-CodeBlocks ([regex]::Replace($script:Documents[$name], '(?s)^---\n.*?\n---\n', ''))
            $text = [regex]::Replace($text, '\*\*[^*\n]+\*\*', '')
            foreach ($m in [regex]::Matches($text, '(?<![\w*])\*([^*\n]{14,})\*(?![\w*])')) {
                $quote = $m.Groups[1].Value.Trim()
                if (@($quote -split '\s+').Count -lt 3 -or $quote -in $exceptions) { continue }
                foreach ($part in [regex]::Split($quote, '…|\.\.\.')) {
                    $fragment = $part.Trim().Trim(':', '.', '?', ',').Trim()
                    if ($fragment.Length -lt 14) { continue }
                    $found = $script:CodeText.Contains($fragment) -or $script:CodeText.Contains($fragment.Replace("'", "''")) -or $script:CodeText.Contains($fragment.Replace('"', '`"'))
                    if (-not $found) { "$name`: $fragment" }
                }
            }
        }
        @($problems | Sort-Object -Unique) | Should -BeNullOrEmpty
    }
    It 'the buttons and titles of the console that the user guide names are the ones of the console' {
        $labels = 'Choose a task', 'Continue to verification', 'Preview', 'Convert the mailboxes', 'Roll back the mailboxes', 'Collect the snapshot', 'Run the check', 'Check again', 'What next?', 'Open report', 'Open global report', 'Open log',
        'Copy batch ID', 'Copy command', 'Stop after the current object', 'Start a new task', 'Start a new batch', 'Continue a batch', 'Start a rollback', 'Continue an interrupted rollback', 'Follow a batch', 'Check the snapshot',
        'Collect a snapshot', 'Move mailboxes to Exchange Online', 'Roll back a conversion', 'Whole batch', 'Tick all shown', 'Untick all', 'Reset everything', 'Look once, do not wait', 'Change settings',
        'One mailbox (try one first)', 'Everything in the snapshot (configured scope)', 'An OU (picked from the snapshot)', 'A group', 'A CSV file', 'Settings for this run only (the file is never written)',
        'The conversion is complete', 'Part of the conversion remains', 'The rollback is complete', 'Part of the rollback remains', 'The snapshot is written', 'The check is finished', 'The check found objects that are not ready', 'The run stopped',
        'This computer can run', 'Prerequisites', 'What will happen on this computer', 'Which objects?', 'What will happen, in order', 'Settings used', 'WARNINGS AND ERRORS', 'BATCH ID', 'Batch status', 'How to do it', 'Or from the command line'
        $normalised = $script:ConsoleText -replace '&#x2026;', '…' -replace '&#xB7;', '·'
        foreach ($label in $labels) {
            $normalised.IndexOf($label, [StringComparison]::OrdinalIgnoreCase) | Should -BeGreaterOrEqual 0 -Because "the console has no '$label'"
            if ($label -notin 'WARNINGS AND ERRORS', 'BATCH ID') { $script:UserGuide.IndexOf($label, [StringComparison]::OrdinalIgnoreCase) | Should -BeGreaterOrEqual 0 -Because "the user guide must name '$label'" }
        }
        foreach ($button in 'Pick...', 'Change…', 'Done', 'Reload') { $normalised.IndexOf($button, [StringComparison]::Ordinal) | Should -BeGreaterOrEqual 0 }
    }
    It 'the typed confirmation words of the user guide are the ones of the console' {
        foreach ($word in 'CONVERT', 'RECOVER', 'COLLECT') { $script:ConsoleText | Should -Match ([regex]::Escape("= '$word'")); $script:UserGuide | Should -Match ([regex]::Escape($word)) }
    }
    It 'the exit codes of the guides are the ones of the entry script' {
        $script:Entry | Should -Match 'Exit codes : 0 = done, 1 = failed \(or an object is not ready\), 2 = done, next step required\.'
        foreach ($text in $script:Readme, $script:UserGuide, $script:Guide) { $text | Should -Match '(?s)0.{0,20}done.{0,200}1.{0,30}failed.{0,200}2.{0,30}done, next step' }
    }
    It 'no guide names a thing that is gone: Windows PowerShell 5.1 as a requirement, ConnectionMode as a setting, lots, rounds or lab names' {
        foreach ($name in $script:Documents.Keys) {
            $text = $script:Documents[$name]
            $text | Should -Not -Match '(?i)\b(?:lot|round) \d' -Because "$name names an internal lot or round"
            $text | Should -Not -Match '(?i)\bWPF window\b' -Because "$name still calls the console a WPF window"
            $text | Should -Not -Match '(?i)developer guide' -Because "$name uses the former name of the administrator guide"
            foreach ($m in [regex]::Matches($text, '(?i)[^\n]*(?:Windows PowerShell 5\.1|powershell\.exe)[^\n]*')) { $m.Value | Should -Match '(?i)not |no longer|never|1\.x|instead of|refused|former|use .pwsh' -Because "$name mentions Windows PowerShell 5.1 as if it were supported: $($m.Value)" }
            foreach ($m in [regex]::Matches($text, '(?i)[^\n]*ConnectionMode[^\n]*')) { $m.Value | Should -Match '(?i)gone|remove|refused|ignored|1\.x' -Because "$name mentions ConnectionMode as if it still existed" }
        }
    }
}

Describe 'Readability of the user guide' {
    It 'the Start here part keeps short sentences and common words' {
        $a = $script:UserGuide.IndexOf('## 1. Start here'); $b = $script:UserGuide.IndexOf('## 3. Set up once')
        $r = Get-Readability $script:UserGuide.Substring($a, $b - $a)
        $r.Average | Should -BeLessThan 12
        $r.Over20 | Should -Be 0
        $r.FleschReadingEase | Should -BeGreaterThan 80
    }
    It 'the whole user guide, tables included, has no sentence over 25 words and stays easy to read' {
        $r = Get-Readability $script:UserGuide $true
        $r.Over25 | Should -Be 0
        $r.Average | Should -BeLessThan 12
        $r.FleschReadingEase | Should -BeGreaterThan 75
    }
    It 'the user guide has the sections of a beginner guide, in order' {
        $order = '## 1. Start here', '## 2. Before you start', '## 3. Set up once', '## 4. A · Move mailboxes to the cloud', '## 5. B · Go back', '## 6. Check and read the report', '## 7. When something goes wrong', '## 8. Picture dictionary', '## 9. Glossary'
        $at = -1
        foreach ($heading in $order) { $i = $script:UserGuide.IndexOf($heading); $i | Should -BeGreaterThan $at -Because $heading; $at = $i }
        $script:UserGuide | Should -Match '\*\*Stop and ask\.\*\*'
        foreach ($term in 'Snapshot', 'Journal', 'Batch', 'Wave', 'Case hold', 'Inactive mailbox', 'Cloud admin server', 'Run-only settings', '`.sha256` file') { $script:UserGuide | Should -Match ('\| \*\*' + [regex]::Escape($term)) -Because "the glossary must define $term" }
    }
}

Describe 'Troubleshooting: Entra Connect' {
    It 'tells what to check when an object stays not recreated after several delta cycles, and never proposes an Initial cycle' {
        $row = @($script:Guide -split "`n" | Where-Object { $_ -like '| An object stays *not recreated by Entra Connect yet*' })
        $row.Count | Should -Be 1
        $row[0] | Should -Match 'Synchronization Service Manager'
        $row[0] | Should -Match 'metaverse object'
        $row[0] | Should -Match 'Entra connector'
        $script:Guide | Should -Not -Match '(?i)-PolicyType\s+Initial|initial (sync|synchronisation|synchronization|cycle)'
        $script:Cloud | Should -Not -Match '(?i)PolicyType Initial'
    }
    It 'tells what to do when the connector is busy, with the message of the tool' {
        $script:Guide | Should -Match '\| \*Entra Connect is busy with another synchronization cycle\* \|'
        $script:Cloud | Should -Match 'Entra Connect is busy with another synchronization cycle: wait for it to finish, then run Recover -Batch'
    }
}
