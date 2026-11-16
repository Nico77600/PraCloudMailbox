<#
.SYNOPSIS
    Test helper: a fake Exchange Online session list (Get-ConnectionInformation, Connect-ExchangeOnline, Disconnect-ExchangeOnline).
.DESCRIPTION
    Dot-source this file in the BeforeAll of a test file that signs in with Connect-PcmCloud, and call Reset-PraFakeExoSession before each test.
    The state is in $global:PraFakeExo: Connections (what Get-ConnectionInformation returns), Disconnects (the sign-outs asked: the connection ID, or
    '*' for a sign-out without a connection ID), InteractiveTenant (the tenant that an interactive sign-in reports) and FailDisconnect (the message
    that a sign-out throws). A test that replaces Connect-ExchangeOnline by a mock calls Add-PraFakeExoConnection from it.
    Remove-PraFakeExoSession removes the functions and the state again.
.NOTES
    Author  : Nicolas Fabert
#>
#Requires -Version 7.5

function global:Reset-PraFakeExoSession {
    $global:PraFakeExo = @{ Connections = [System.Collections.Generic.List[object]]::new(); Disconnects = [System.Collections.Generic.List[string]]::new()
        InteractiveTenant = '33333333-3333-3333-3333-333333333333'; FailDisconnect = '' }
}

function global:Add-PraFakeExoConnection {
    param([string]$TenantId = '33333333-3333-3333-3333-333333333333', [string]$AppId = '', [string]$ConnectionId = ([guid]::NewGuid().ToString()))
    $connection = [pscustomobject]@{ ConnectionId = $ConnectionId; TenantID = $TenantId; AppId = $AppId; State = 'Connected'; TokenStatus = 'Active' }
    $global:PraFakeExo.Connections.Add($connection)
    return $connection
}

function global:Get-ConnectionInformation {
    [CmdletBinding()]
    param()
    return @($global:PraFakeExo.Connections)
}

function global:Connect-ExchangeOnline {
    [CmdletBinding()]
    param($AppId, $CertificateThumbprint, $Organization, $ShowBanner, $CommandName, $UserPrincipalName, [switch]$DisableWAM)
    $null = Add-PraFakeExoConnection -TenantId $(if ($AppId) { [string]$Organization } else { [string]$global:PraFakeExo.InteractiveTenant }) -AppId ([string]$AppId)
}

function global:Connect-IPPSSession {
    [CmdletBinding()]
    param($AppId, $CertificateThumbprint, $Organization, $ShowBanner, $CommandName, $UserPrincipalName, [switch]$EnableSearchOnlySession)
    $null = Add-PraFakeExoConnection -TenantId $(if ($AppId) { [string]$Organization } else { [string]$global:PraFakeExo.InteractiveTenant }) -AppId ([string]$AppId)
}

function global:Disconnect-ExchangeOnline {
    [CmdletBinding(SupportsShouldProcess)]
    param([string]$ConnectionId)
    if ($global:PraFakeExo.FailDisconnect) { throw $global:PraFakeExo.FailDisconnect }
    Write-Host 'Disconnected successfully !'
    $global:PraFakeExo.Disconnects.Add($(if ($ConnectionId) { $ConnectionId } else { '*' }))
    if ($ConnectionId) { $null = $global:PraFakeExo.Connections.RemoveAll([Predicate[object]]{ param($c) $c.ConnectionId -eq $ConnectionId }) } else { $global:PraFakeExo.Connections.Clear() }
}

function global:Remove-PraFakeExoSession {
    foreach ($name in 'Reset-PraFakeExoSession', 'Add-PraFakeExoConnection', 'Get-ConnectionInformation', 'Connect-ExchangeOnline', 'Connect-IPPSSession', 'Disconnect-ExchangeOnline', 'Remove-PraFakeExoSession') {
        Remove-Item -Path "Function:\$name" -ErrorAction SilentlyContinue
    }
    Remove-Variable -Name PraFakeExo -Scope Global -ErrorAction SilentlyContinue
}

Reset-PraFakeExoSession
