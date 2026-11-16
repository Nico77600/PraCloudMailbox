#Requires -Version 7.5
# The one JSON reader and the culture-invariant dates of the whole package.

BeforeAll {
    . (Join-Path $PSScriptRoot 'GlobalState.ps1')
    $script:GlobalBefore = Save-PcmGlobalState
    . (Join-Path $PSScriptRoot 'IntegrityFixture.ps1')
    $script:Repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
    $script:Package = New-PcmTestPackage -Repository $script:Repository
    Import-Module (Join-Path $script:Package 'module\PRA.Common.psm1') -Force
    Import-Module (Join-Path $script:Package 'module\PRA.Cloud.psm1') -Force
}

AfterAll {
    Remove-PcmTestPackage $script:Package
    Restore-PcmGlobalState $script:GlobalBefore
}

Describe 'The JSON reader' {
    It 'keeps a text that looks like a date exactly as it was written' {
        $value = '{"when":"2025-01-31T00:00:00Z","list":["2026-02-03T04:05:06+02:00"]}' | ConvertFrom-PcmJson
        $value.when | Should -BeOfType [string]
        $value.when | Should -BeExactly '2025-01-31T00:00:00Z'
        @($value.list)[0] | Should -BeExactly '2026-02-03T04:05:06+02:00'
    }

    It 'reads an array, an empty text and several lines of a pipeline' {
        @('["a","b"]' | ConvertFrom-PcmJson).Count | Should -Be 2
        ('' | ConvertFrom-PcmJson) | Should -BeNullOrEmpty
        (@('{"a":', '1}') | ConvertFrom-PcmJson).a | Should -Be 1
    }

    It 'refuses a text that is not JSON' {
        { 'not json' | ConvertFrom-PcmJson } | Should -Throw
    }

    It 'is the only reader of JSON in the package' {
        $files = @(Get-ChildItem -LiteralPath (Join-Path $script:Repository 'package') -Recurse -File -Include *.ps1, *.psm1)
        $files.Count | Should -BeGreaterThan 10
        $found = foreach ($file in $files) {
            $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$null)
            foreach ($command in $ast.FindAll({ param($node) $node -is [Management.Automation.Language.CommandAst] -and [string]$node.GetCommandName() -eq 'ConvertFrom-Json' }, $true)) {
                $owner = $command.Parent
                while ($owner -and $owner -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $owner = $owner.Parent }
                if ($owner -and $owner.Name -eq 'ConvertFrom-PcmJson') { continue }
                [pscustomobject]@{ File = $file.Name; Line = $command.Extent.StartLineNumber; Text = $command.Parent.Extent.Text }
            }
        }
        # Only the report of Invoke-ADSyncSingleObjectSync, read on the Entra Connect server where the module of the tool is not loaded, keeps the plain reader.
        $others = @($found | Where-Object { $_.Text -notmatch 'Invoke-ADSyncSingleObjectSync' -or $_.File -notin @('EntraConnect.sample.ps1', 'PRA.Cloud.psm1') })
        @($others | ForEach-Object { '{0}:{1}' -f $_.File, $_.Line }) | Should -BeNullOrEmpty
    }
}

Describe 'A date-like value of a resource survives a restore' {
    BeforeAll {
        $global:PcmDateFake = @{ Set = $null }
        $names = @('Identity', 'EmailAddresses', 'ResourceCapacity', 'ResourceCustom', 'HiddenFromAddressListsEnabled') + @(1..15 | ForEach-Object { "CustomAttribute$_" }) +
            @('RequireSenderAuthenticationEnabled', 'ModerationEnabled', 'SendModerationNotifications', 'AcceptMessagesOnlyFromSendersOrMembers', 'RejectMessagesFromSendersOrMembers', 'ModeratedBy', 'BypassModerationFromSendersOrMembers')
        $declaration = '[CmdletBinding()] param(' + (($names | ForEach-Object { '$' + $_ }) -join ', ') + ') $global:PcmDateFake.Set = $PSBoundParameters; throw ''stopped after Set-Mailbox'''
        . (Join-Path $PSScriptRoot 'FakeTenant.ps1')
        Set-Item -Path Function:\global:Set-Mailbox -Value ([scriptblock]::Create($declaration))
        $settings = New-PraFakeResourceSettings
        $settings.Directory.Office = '2025-01-31T00:00:00Z'
        $script:Record = [pscustomobject]@{ object_guid = 'source'; kind = 'Room'; primary_smtp_address = 'room@contoso.com'; email_addresses_json = '["SMTP:room@contoso.com"]'
            legacy_exchange_dn = ''; resource_settings_json = ($settings | ConvertTo-Json -Depth 8 -Compress)
            custom_attributes_json = '{"CustomAttribute3":"2025-01-31T00:00:00Z","CustomAttribute4":"plain"}' }
    }
    AfterAll { Remove-Variable -Name PcmDateFake -Scope Global -ErrorAction SilentlyContinue }

    It 'reads a directory field that looks like a date as the text it is' {
        $read = Read-PcmResourceSettings -Record $script:Record
        $read.Directory.Office | Should -BeOfType [string]
        $read.Directory.Office | Should -BeExactly '2025-01-31T00:00:00Z'
    }

    It 'writes a custom attribute that looks like a date byte for byte' {
        Mock Confirm-PcmWrite -ModuleName PRA.Cloud { }
        Mock Resolve-PcmResourceRecipients -ModuleName PRA.Cloud { @() }
        { Set-PcmResourceConfiguration -Record $script:Record -Owned ([pscustomobject]@{ Id = 'owned-id' }) -Intent ([pscustomobject]@{ markerAttribute = 'CustomAttribute15'; token = 'abc' }) } | Should -Throw '*stopped after Set-Mailbox*'
        $global:PcmDateFake.Set.CustomAttribute3 | Should -BeExactly '2025-01-31T00:00:00Z'
        $global:PcmDateFake.Set.CustomAttribute4 | Should -BeExactly 'plain'
    }
}
