<#
.SYNOPSIS
    Test helper: gives the global scope back as a test file found it.
.DESCRIPTION
    The tests replace commands of Exchange Online, Microsoft Graph and the tool by global functions (the modules under test do not see the scope of a test file) and keep
    their fakes in global variables. A global function stays until the session ends and hides the real command from every file that runs after it, in a run of several
    files. Dot-source this file in the BeforeAll of the file, call Save-PcmGlobalState before the first global function or variable is made, and Restore-PcmGlobalState in
    the AfterAll of the file: the functions and the variables that the file added are removed, and the ones that it replaced are put back. The functions of the imported
    modules are not touched. Restore-PcmGlobalState removes the two functions of this file too.
.NOTES
    Author  : Nicolas Fabert
#>
function global:Save-PcmGlobalState {
    [CmdletBinding()]
    param()
    [pscustomobject]@{
        Functions = @(Get-ChildItem -Path Function:\ | Where-Object { -not $_.ModuleName -and $_.Name -notin 'Save-PcmGlobalState', 'Restore-PcmGlobalState' } |
                ForEach-Object { [pscustomobject]@{ Name = $_.Name; Body = [string]$_.ScriptBlock } })
        Variables = @(Get-Variable -Scope Global | ForEach-Object { $_.Name })
    }
}

function global:Restore-PcmGlobalState {
    [CmdletBinding()]
    param([Parameter(Mandatory)][object]$State)
    $functions = @{}
    foreach ($function in $State.Functions) { $functions[$function.Name] = $function.Body }
    foreach ($function in @(Get-ChildItem -Path Function:\ | Where-Object { -not $_.ModuleName })) {
        if (-not $functions.ContainsKey($function.Name)) {
            for ($pass = 0; $pass -lt 4 -and (Test-Path -LiteralPath "Function:\$($function.Name)"); $pass++) { Remove-Item -LiteralPath "Function:\$($function.Name)" -Force -ErrorAction SilentlyContinue }
        } elseif ($functions[$function.Name] -ne [string]$function.ScriptBlock) {
            Set-Item -LiteralPath "Function:\global:$($function.Name)" -Value ([scriptblock]::Create($functions[$function.Name]))
        }
    }
    $variables = @{}
    foreach ($name in $State.Variables) { $variables[$name] = $true }
    foreach ($variable in @(Get-Variable -Scope Global)) {
        if (-not $variables.ContainsKey($variable.Name)) { Remove-Variable -Name $variable.Name -Scope Global -Force -ErrorAction SilentlyContinue }
    }
}
