function global:Get-MgContext {
    [CmdletBinding()]
    param()
    $global:PraFakeGraphAuth
}

function global:Get-MgRequestContext {
    [CmdletBinding()]
    param()
    $global:PraFakeGraphRequest
}

function global:Set-MgRequestContext {
    [CmdletBinding()]
    param([ValidateRange(0, 10)][int]$MaxRetry)
    if ($PSBoundParameters.ContainsKey('MaxRetry')) { $global:PraFakeGraphRequest.MaxRetry = $MaxRetry }
    $global:PraFakeGraphClientGeneration++
    $global:PraFakeGraphRequest
}

function global:Reset-PraFakeGraphTransport {
    $global:PraFakeGraphAuth = $null
    $global:PraFakeGraphRequest = [pscustomobject]@{ MaxRetry = 3; RetryDelay = 3; RetriesTimeLimit = [timespan]::Zero; ClientTimeout = [timespan]::FromMinutes(5) }
    $global:PraFakeGraphClientGeneration = 0
    & (Get-Module PRA2.Cloud) {
        $script:GraphRequestOwner = $null
        $script:GraphTransportReady = $false
        $script:GraphPreviousMaxRetry = $null
        Initialize-Pra2GraphTransport -Context @{}
    }
}
