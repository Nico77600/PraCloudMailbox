#Requires -Version 5.1
<#
.SYNOPSIS
    Sample of EntraConnect.Mode = 'Script' for PRA Cloud Mailbox: runs one Entra Connect operation on the
    rebuilt server with an account other than the one running the tool.

.DESCRIPTION
    Recover needs three operations on the Entra Connect server: Pause and Resume of the scheduler (around the
    deletion of the shared mailbox identities) and a Delta cycle that it waits for. EntraConnect.Mode =
    'Remoting' does them with Invoke-Command and the current account. When that is not possible (the cloud
    admin server is not in the domain, another account is needed, the server is reached through a jump host
    or an Azure Run Command...), write a script with this contract and set EntraConnect.Mode = 'Script' and
    EntraConnect.ScriptPath to its path:

        -Operation Pause | Resume | Delta     the operation to run
        output                               a short state (one line is best): it goes to the log
        exit 0                               done; any other exit code (or an exception) = failed

    This sample uses PowerShell remoting with a credential saved for the account that runs the tool:

        Get-Credential CONTOSO\svc-pra-sync | Export-Clixml .\config\EntraConnect.cred.xml

    (the file can only be read by the same Windows account on the same computer). The account needs to be
    a member of ADSyncAdmins on the Entra Connect server and allowed to use WinRM there.

.PARAMETER Operation
    Pause, Resume or Delta.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
    Part of : PRA Cloud Mailbox (sample, copy and adapt it)
#>
[CmdletBinding()]
param([Parameter(Mandatory)][ValidateSet('Pause', 'Resume', 'Delta')][string]$Operation)
$ErrorActionPreference = 'Stop'

# ---- Adapt these two values ------------------------------------------------------------------------
$server = 'aadc01.contoso.com'
$credentialFile = Join-Path $PSScriptRoot 'EntraConnect.cred.xml'
# ------------------------------------------------------------------------------------------------------

try {
    $credential = Import-Clixml -LiteralPath $credentialFile
    $state = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock {
        Import-Module ADSync -ErrorAction Stop
        switch ($using:Operation) {
            'Pause' { Set-ADSyncScheduler -SyncCycleEnabled $false }
            'Resume' { Set-ADSyncScheduler -SyncCycleEnabled $true }
            'Delta' {
                # A cycle already running is waited for, then the delta cycle is started and waited for.
                $t0 = Get-Date
                while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 15) { Start-Sleep -Seconds 10 }
                Start-ADSyncSyncCycle -PolicyType Delta | Out-Null
                Start-Sleep -Seconds 10
                $t0 = Get-Date
                while ((Get-ADSyncScheduler).SyncCycleInProgress -and ((Get-Date) - $t0).TotalMinutes -lt 30) { Start-Sleep -Seconds 10 }
            }
        }
        $s = Get-ADSyncScheduler
        'SyncCycleEnabled={0} InProgress={1}' -f $s.SyncCycleEnabled, $s.SyncCycleInProgress
    }
    "$Operation on ${server}: $(@($state) -join ' ')"
    exit 0
} catch {
    "$Operation on $server failed: $($_.Exception.Message)"
    exit 1
}
