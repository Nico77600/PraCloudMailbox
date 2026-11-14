# The engine child of a run: the command line, the new console of the run (sign-in window of the Windows account broker),
# the redirected streams and their collectors. Dot-sourced by PRA.Gui.psm1.

$script:PraGuiNativeProcess = $null

function Add-PcmGuiNativeImport {
    <# One kernel32 import, emitted with Reflection.Emit: DllImport with SetLastError = true (the Win32 error is kept for Marshal.GetLastWin32Error / Get-PcmGuiWin32ErrorText). #>
    param([Parameter(Mandatory)][Reflection.Emit.TypeBuilder]$Type, [Parameter(Mandatory)][string]$Name, [string]$EntryPoint = '', [Parameter(Mandatory)][Type]$Return, [Parameter(Mandatory)][Type[]]$Parameters)
    if (-not $EntryPoint) { $EntryPoint = $Name }
    $method = $Type.DefineMethod($Name, 'Public,Static,PinvokeImpl,HideBySig', $Return, $Parameters)
    $method.SetImplementationFlags([Reflection.MethodImplAttributes]::PreserveSig)
    $import = [Runtime.InteropServices.DllImportAttribute]
    $attribute = New-Object Reflection.Emit.CustomAttributeBuilder($import.GetConstructor([Type[]]@([string])), [object[]]@('kernel32.dll'),
        [Reflection.FieldInfo[]]@($import.GetField('EntryPoint'), $import.GetField('CharSet'), $import.GetField('SetLastError')),
        [object[]]@($EntryPoint, [Runtime.InteropServices.CharSet]::Unicode, $true))
    $method.SetCustomAttribute($attribute)
}

function Get-PcmGuiWin32ErrorText {
    <# '<code>: <system message>' for a Win32 error code (Marshal.GetLastWin32Error after a native call). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$Code)
    return ('{0}: {1}' -f $Code, (New-Object ComponentModel.Win32Exception($Code)).Message)
}

function Get-PcmGuiNativeProcessApi {
    <#
    CreateProcessW and its structures, bound in memory with Reflection.Emit in memory (no compiler, temporary file, DLL or cache). Used by the
    window to start the engine child with its own console while its three standard streams stay redirected. Every import keeps the Win32 error (SetLastError = true).
    Returns a table: Api (the static class), StartupInfo (STARTUPINFOW), StartupInfoEx (STARTUPINFOEXW: StartupInfo + lpAttributeList),
    ProcessInformation. Api methods: CreateProcessW (STARTUPINFOW), CreateProcessEx (the same native CreateProcessW with a STARTUPINFOEXW, for a handle list),
    InitializeProcThreadAttributeList, UpdateProcThreadAttribute, DeleteProcThreadAttributeList, GetExitCodeProcess, CloseHandle.
    #>
    if ($script:PraGuiNativeProcess) { return $script:PraGuiNativeProcess }
    $assembly = [Reflection.Emit.AssemblyBuilder]::DefineDynamicAssembly((New-Object Reflection.AssemblyName 'PraNativeProcess'), [Reflection.Emit.AssemblyBuilderAccess]::Run)
    $module = $assembly.DefineDynamicModule('PraNativeProcess')
    $startup = $module.DefineType('PraStartupInfo', 'Public,SequentialLayout,Sealed', [ValueType])
    foreach ($field in @(@('cb', [uint32]), @('lpReserved', [IntPtr]), @('lpDesktop', [IntPtr]), @('lpTitle', [IntPtr]), @('dwX', [uint32]), @('dwY', [uint32]),
            @('dwXSize', [uint32]), @('dwYSize', [uint32]), @('dwXCountChars', [uint32]), @('dwYCountChars', [uint32]), @('dwFillAttribute', [uint32]),
            @('dwFlags', [uint32]), @('wShowWindow', [uint16]), @('cbReserved2', [uint16]), @('lpReserved2', [IntPtr]),
            @('hStdInput', [IntPtr]), @('hStdOutput', [IntPtr]), @('hStdError', [IntPtr]))) { [void]$startup.DefineField($field[0], $field[1], 'Public') }
    $startupType = $startup.CreateType()
    $startupEx = $module.DefineType('PraStartupInfoEx', 'Public,SequentialLayout,Sealed', [ValueType])
    [void]$startupEx.DefineField('StartupInfo', $startupType, 'Public'); [void]$startupEx.DefineField('lpAttributeList', [IntPtr], 'Public')
    $startupExType = $startupEx.CreateType()
    $information = $module.DefineType('PraProcessInformation', 'Public,SequentialLayout,Sealed', [ValueType])
    foreach ($field in @(@('hProcess', [IntPtr]), @('hThread', [IntPtr]), @('dwProcessId', [uint32]), @('dwThreadId', [uint32]))) { [void]$information.DefineField($field[0], $field[1], 'Public') }
    $informationType = $information.CreateType()
    $type = $module.DefineType('Api', 'Public,Class,Sealed,Abstract')
    $createParameters = @([string], [Text.StringBuilder], [IntPtr], [IntPtr], [bool], [uint32], [IntPtr], [string])
    Add-PcmGuiNativeImport -Type $type -Name 'CreateProcessW' -Return ([bool]) -Parameters ([Type[]]($createParameters + @($startupType.MakeByRefType(), $informationType.MakeByRefType())))
    Add-PcmGuiNativeImport -Type $type -Name 'CreateProcessEx' -EntryPoint 'CreateProcessW' -Return ([bool]) -Parameters ([Type[]]($createParameters + @($startupExType.MakeByRefType(), $informationType.MakeByRefType())))
    Add-PcmGuiNativeImport -Type $type -Name 'InitializeProcThreadAttributeList' -Return ([bool]) -Parameters ([Type[]]@([IntPtr], [int], [int], [IntPtr].MakeByRefType()))
    Add-PcmGuiNativeImport -Type $type -Name 'UpdateProcThreadAttribute' -Return ([bool]) -Parameters ([Type[]]@([IntPtr], [uint32], [IntPtr], [IntPtr], [IntPtr], [IntPtr], [IntPtr]))
    Add-PcmGuiNativeImport -Type $type -Name 'DeleteProcThreadAttributeList' -Return ([void]) -Parameters ([Type[]]@([IntPtr]))
    Add-PcmGuiNativeImport -Type $type -Name 'GetExitCodeProcess' -Return ([bool]) -Parameters ([Type[]]@([IntPtr], [uint32].MakeByRefType()))
    Add-PcmGuiNativeImport -Type $type -Name 'CloseHandle' -Return ([bool]) -Parameters ([Type[]]@([IntPtr]))
    $script:PraGuiNativeProcess = @{ Api = $type.CreateType(); StartupInfo = $startupType; StartupInfoEx = $startupExType; ProcessInformation = $informationType }
    return $script:PraGuiNativeProcess
}

function ConvertTo-PcmGuiNativeArgument {
    <# CommandLineToArgvW: double backslashes before quotes and the closing quote. No shell is used. #>
    param([AllowEmptyString()][string]$Value)
    if ($Value.Contains([string][char]0)) { throw 'NUL is not a valid argument.' }
    $escaped = [regex]::Replace($Value, '(\\*)"', '$1$1\"')
    $escaped = [regex]::Replace($escaped, '(\\+)$', '$1$1')
    return '"' + $escaped + '"'
}

# The redirected output of the engine child goes straight to a file with pure .NET code: Stream.CopyToAsync (thread pool, no
# PowerShell code on a background thread, no compiler, no temporary DLL) copies the pipe to <base>.stdout.txt / .stderr.txt as the
# child writes, so a child that prints more than the pipe buffer never blocks whatever the window timer does. The window thread only
# reads what has been written so far, through a second FileStream, and decodes it with the console code page of the child.
function New-PcmGuiStreamCollector {
    <#
    .SYNOPSIS
        Collector of one redirected child stream: IsCompleted (end of stream), TakeText() (text received since the previous
        call) and Result (the whole text; waits for the end of the stream). The raw bytes are kept in -Path.
    #>
    param([Parameter(Mandatory)][IO.Stream]$Stream, [Parameter(Mandatory)][string]$Path, [Text.Encoding]$Encoding)
    # The child writes with the console code page; without a console in this host, UTF-8.
    if (-not $Encoding) { try { $Encoding = [Console]::OutputEncoding } catch { $Encoding = New-Object Text.UTF8Encoding($false) } }
    # Not buffered (bufferSize 1): each write reaches the file at once, so the window sees it.
    $file = New-Object IO.FileStream($Path, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite, 1)
    $state = @{ Stream = $Stream; File = $file; Path = $Path; Task = $Stream.CopyToAsync($file); Offset = [long]0; Decoder = $Encoding.GetDecoder(); Encoding = $Encoding; Closed = $false }
    $collector = New-Object psobject -Property @{ State = $state }
    Add-Member -InputObject $collector -MemberType ScriptMethod -Name Close -Value {
        if (-not $this.State.Closed) {
            $this.State.Closed = $true
            try { $this.State.File.Flush() } catch { Write-Verbose $_.Exception.Message }
            $this.State.File.Dispose(); $this.State.Stream.Dispose()
        }
    }
    Add-Member -InputObject $collector -MemberType ScriptProperty -Name IsCompleted -Value {
        if (-not $this.State.Task.IsCompleted) { return $false }
        $this.Close()
        return $true
    }
    Add-Member -InputObject $collector -MemberType ScriptMethod -Name TakeText -Value {
        $reader = [IO.File]::Open($this.State.Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $count = [int][Math]::Min(1048576, $reader.Length - $this.State.Offset)
            if ($count -le 0) { return '' }
            [void]$reader.Seek($this.State.Offset, [IO.SeekOrigin]::Begin)
            $buffer = New-Object byte[] $count
            $read = $reader.Read($buffer, 0, $count)
            $this.State.Offset += $read
            $chars = New-Object char[] ($this.State.Encoding.GetMaxCharCount($read))
            $length = $this.State.Decoder.GetChars($buffer, 0, $read, $chars, 0)
            return [string]::new($chars, 0, $length)
        } finally { $reader.Dispose() }
    }
    Add-Member -InputObject $collector -MemberType ScriptProperty -Name Result -Value {
        try { $null = $this.State.Task.GetAwaiter().GetResult() } catch { Write-Verbose $_.Exception.Message }
        $this.Close()
        return $this.State.Encoding.GetString([IO.File]::ReadAllBytes($this.State.Path))
    }
    return $collector
}

function Start-PcmGuiNativeConsoleChild {
    <#
    .SYNOPSIS
        Starts the engine child of a cloud action in its OWN, new, visible console (CREATE_NEW_CONSOLE) with stdin, stdout and
        stderr still redirected to anonymous pipes. Returns the process and the three pipe streams.
    .DESCRIPTION
        .NET cannot ask for a new console together with redirected streams (ProcessStartInfo only inherits the console of the
        parent), so CreateProcessW is called directly through a Reflection.Emit binding (no compiler, no DLL).
        The child no longer shares the console of the terminal that started the window: closing or hiding that terminal does not
        reach the engine in the middle of an Apply, and the Windows account broker (WAM) of the interactive sign-in gets the
        window of this new console as its parent.
        Only the three pipe ends of the child are inherited: CreateProcess is given a STARTUPINFOEX with a PROC_THREAD_ATTRIBUTE_HANDLE_LIST
        naming them, so no other inheritable handle of the window process (a pipe end of another child, a file) reaches the engine - an
        extra copy of a pipe end held by the child would delay the end of the stream and the run would never complete. The window itself creates
        no inheritable handle but those three client ends, which it closes as soon as the child owns them.
        A child that dies before the window can attach to it (the process list no longer has it) is not silent: its error output is drained
        to the end of the stream, its exit code is read from the process handle, and both are in the error.
    #>
    param($Command, [string]$Root, [hashtable]$Environment)
    $native = Get-PcmGuiNativeProcessApi
    $out = New-Object IO.Pipes.AnonymousPipeServerStream([IO.Pipes.PipeDirection]::In, [IO.HandleInheritability]::Inheritable)
    $err = New-Object IO.Pipes.AnonymousPipeServerStream([IO.Pipes.PipeDirection]::In, [IO.HandleInheritability]::Inheritable)
    $in = New-Object IO.Pipes.AnonymousPipeServerStream([IO.Pipes.PipeDirection]::Out, [IO.HandleInheritability]::Inheritable)
    $title = [IntPtr]::Zero; $block = [IntPtr]::Zero; $attributeList = [IntPtr]::Zero; $handleList = [IntPtr]::Zero; $attributeListReady = $false
    $processHandle = $null; $threadHandle = $null
    try {
        $variables = [Environment]::GetEnvironmentVariables()
        $merged = @{}
        foreach ($key in $variables.Keys) { $merged[[string]$key] = [string]$variables[$key] }
        foreach ($key in $Environment.Keys) {
            if ($null -eq $Environment[$key]) { [void]$merged.Remove([string]$key) } else { $merged[[string]$key] = [string]$Environment[$key] }
        }
        # The environment block is sorted by upper-cased name, by UTF-16 code unit (never the culture collation).
        $names = [string[]]@($merged.Keys)
        $sortKeys = [string[]]@(foreach ($name in $names) { $name.ToUpperInvariant() })
        [Array]::Sort([Array]$sortKeys, [Array]$names, [Collections.IComparer][StringComparer]::Ordinal)
        $text = (@($names | ForEach-Object { $_ + '=' + $merged[$_] }) -join [string][char]0) + [string][char]0 + [string][char]0
        $block = [Runtime.InteropServices.Marshal]::StringToHGlobalUni($text)
        $title = [Runtime.InteropServices.Marshal]::StringToHGlobalUni('PRA Cloud Mailbox - cloud sign-in (keep this window open until the operation ends)')
        $api = $native.Api
        # The three client ends are the only handles the child may inherit.
        $clientHandles = @($in.ClientSafePipeHandle.DangerousGetHandle(), $out.ClientSafePipeHandle.DangerousGetHandle(), $err.ClientSafePipeHandle.DangerousGetHandle())
        $startupEx = [Activator]::CreateInstance($native.StartupInfoEx)
        $startup = $startupEx.GetType().GetField('StartupInfo').GetValue($startupEx)
        $startup.GetType().GetField('cb').SetValue($startup, [uint32][Runtime.InteropServices.Marshal]::SizeOf($startupEx))
        $startup.GetType().GetField('lpTitle').SetValue($startup, $title)
        $startup.GetType().GetField('dwFlags').SetValue($startup, [uint32]0x100)   # STARTF_USESTDHANDLES
        $startup.GetType().GetField('hStdInput').SetValue($startup, $clientHandles[0])
        $startup.GetType().GetField('hStdOutput').SetValue($startup, $clientHandles[1])
        $startup.GetType().GetField('hStdError').SetValue($startup, $clientHandles[2])
        $startupEx.GetType().GetField('StartupInfo').SetValue($startupEx, $startup)
        # Attribute list with PROC_THREAD_ATTRIBUTE_HANDLE_LIST (0x20002): first call asks the size, second one initialises the list.
        $initialize = New-Object object[] 4
        $initialize[0] = [IntPtr]::Zero; $initialize[1] = [int]1; $initialize[2] = [int]0; $initialize[3] = [IntPtr]::Zero
        [void]$api.GetMethod('InitializeProcThreadAttributeList').Invoke($null, $initialize)
        $size = [IntPtr]$initialize[3]
        if ($size -eq [IntPtr]::Zero) { throw ('The engine child process cannot start (attribute list size: {0}).' -f (Get-PcmGuiWin32ErrorText ([Runtime.InteropServices.Marshal]::GetLastWin32Error()))) }
        $attributeList = [Runtime.InteropServices.Marshal]::AllocHGlobal($size)
        $initialize[0] = $attributeList; $initialize[3] = $size
        if (-not [bool]$api.GetMethod('InitializeProcThreadAttributeList').Invoke($null, $initialize)) {
            throw ('The engine child process cannot start (InitializeProcThreadAttributeList failed: {0}).' -f (Get-PcmGuiWin32ErrorText ([Runtime.InteropServices.Marshal]::GetLastWin32Error())))
        }
        $attributeListReady = $true
        $handleList = [Runtime.InteropServices.Marshal]::AllocHGlobal([IntPtr]::Size * 3)
        for ($i = 0; $i -lt 3; $i++) { [Runtime.InteropServices.Marshal]::WriteIntPtr($handleList, $i * [IntPtr]::Size, [IntPtr]$clientHandles[$i]) }
        $update = New-Object object[] 7
        $update[0] = $attributeList; $update[1] = [uint32]0; $update[2] = [IntPtr]0x20002; $update[3] = $handleList; $update[4] = [IntPtr]([IntPtr]::Size * 3); $update[5] = [IntPtr]::Zero; $update[6] = [IntPtr]::Zero
        if (-not [bool]$api.GetMethod('UpdateProcThreadAttribute').Invoke($null, $update)) {
            throw ('The engine child process cannot start (UpdateProcThreadAttribute failed: {0}).' -f (Get-PcmGuiWin32ErrorText ([Runtime.InteropServices.Marshal]::GetLastWin32Error())))
        }
        $startupEx.GetType().GetField('lpAttributeList').SetValue($startupEx, $attributeList)
        $information = [Activator]::CreateInstance($native.ProcessInformation)
        $commandLine = New-Object Text.StringBuilder ('"' + $Command.FilePath + '" ' + $Command.Arguments)
        # CREATE_NEW_CONSOLE (0x10) | CREATE_UNICODE_ENVIRONMENT (0x400) | EXTENDED_STARTUPINFO_PRESENT (0x80000); bInheritHandles is true, but
        # only the three handles of the list are inherited.
        $arguments = New-Object object[] 10
        $arguments[1] = [Text.StringBuilder]$commandLine; $arguments[2] = [IntPtr]::Zero; $arguments[3] = [IntPtr]::Zero; $arguments[4] = $true
        $arguments[5] = [uint32](0x410 -bor 0x80000); $arguments[6] = [IntPtr]$block; $arguments[7] = [string]$Root; $arguments[8] = $startupEx.psobject.BaseObject; $arguments[9] = $information.psobject.BaseObject
        $created = [bool]$api.GetMethod('CreateProcessEx').Invoke($null, $arguments)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        if (-not $created) { throw ('The engine child process cannot start (CreateProcess failed, Win32 error {0}).' -f (Get-PcmGuiWin32ErrorText $lastError)) }
        $information = $arguments[9]
        $processId = [int]$information.GetType().GetField('dwProcessId').GetValue($information)
        $processHandle = New-Object Microsoft.Win32.SafeHandles.SafeProcessHandle ($information.GetType().GetField('hProcess').GetValue($information), $true)
        $threadHandle = New-Object Microsoft.Win32.SafeHandles.SafeProcessHandle ($information.GetType().GetField('hThread').GetValue($information), $true)
        # The child owns its ends now: close ours. Closing the write end of its stdin gives it an end of file at once.
        $out.DisposeLocalCopyOfClientHandle(); $err.DisposeLocalCopyOfClientHandle(); $in.DisposeLocalCopyOfClientHandle()
        $in.Dispose()
        # Touching SafeHandle makes the Process object open and keep its own handle now, while ours still pins the process (its exit code stays readable).
        $process = $null
        try { $process = [Diagnostics.Process]::GetProcessById($processId); $null = $process.SafeHandle }
        catch {
            # The child is already gone (it died at start, before the window could attach): its error output and its exit code are not lost.
            $attachError = $_.Exception.Message
            $code = [uint32]0
            $codeArguments = New-Object object[] 2
            $codeArguments[0] = $processHandle.DangerousGetHandle(); $codeArguments[1] = $code
            [void]$api.GetMethod('GetExitCodeProcess').Invoke($null, $codeArguments)
            $exitCode = [uint32]$codeArguments[1]
            $errorText = ''
            if ($exitCode -ne 259) {   # 259 = STILL_ACTIVE: never wait for a child that has not ended
                $reader = New-Object IO.StreamReader($err)
                $task = $reader.ReadToEndAsync()
                if ($task.Wait(5000)) { $errorText = $task.Result.Trim() }
            }
            throw ('The engine child process ended before the window could attach to it (exit code {0}; {1}). Error output: {2}' -f $exitCode,$attachError,$(if ($errorText) { $errorText } else { '(none)' }))
        }
        return [pscustomobject]@{ Process = $process; Output = $out; Errors = $err }
    }
    catch {
        $out.Dispose(); $err.Dispose(); $in.Dispose()
        throw
    }
    finally {
        if ($threadHandle) { $threadHandle.Dispose() }
        if ($processHandle) { $processHandle.Dispose() }
        if ($attributeListReady) { try { $native.Api.GetMethod('DeleteProcThreadAttributeList').Invoke($null, @($attributeList)) | Out-Null } catch { Write-Verbose $_.Exception.Message } }
        if ($attributeList -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($attributeList) }
        if ($handleList -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($handleList) }
        if ($title -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($title) }
        if ($block -ne [IntPtr]::Zero) { [Runtime.InteropServices.Marshal]::FreeHGlobal($block) }
    }
}

function Start-PcmGuiChildProcess {
    <#
    .SYNOPSIS
        Starts the engine child of a run, every action in its own visible console (the Microsoft sign-in window and a prompt for credentials need one). Without a console only for the tests and the pictures (the hook NoConsole).
    .DESCRIPTION
        Never -NonInteractive (an interactive sign-in must be allowed to open its window); the child gets an end of file on its standard input instead, so a
        prompt nobody can answer fails at once instead of waiting for a keyboard. The events of the run are written to EventPath (PRA_EVENT_FILE), and
        StopPath names the file whose existence asks the run to stop after the current object (PRA_STOP_FILE).
    #>
    param($Command, [string]$Root, [string]$EventPath, [string]$StopPath = '')
    if (-not (Test-Path -LiteralPath $Command.FilePath -PathType Leaf)) { throw ('{0} is unavailable for the engine child process.' -f [IO.Path]::GetFileName([string]$Command.FilePath)) }
    $needsWindow = ($null -ne $Command.PSObject.Properties['NeedsConsoleWindow']) -and [bool]$Command.NeedsConsoleWindow
    $environment = @{ PRA_EVENT_FILE = $EventPath; PRA_ICONS = 'Ascii'; PRA_STOP_FILE = $(if ($StopPath) { $StopPath } else { $null }) }
    $base = $EventPath -replace '\.events\.jsonl$', ''
    if ($needsWindow) { $child = Start-PcmGuiNativeConsoleChild $Command $Root $environment }
    else {
        $info = New-Object Diagnostics.ProcessStartInfo
        $info.FileName = $Command.FilePath; $info.Arguments = $Command.Arguments; $info.WorkingDirectory = $Root
        $info.UseShellExecute = $false; $info.CreateNoWindow = $true
        $info.RedirectStandardOutput = $true; $info.RedirectStandardError = $true; $info.RedirectStandardInput = $true
        foreach ($key in $environment.Keys) {
            if ($null -eq $environment[$key]) { [void]$info.EnvironmentVariables.Remove($key) } else { $info.EnvironmentVariables[$key] = [string]$environment[$key] }
        }
        $process = [Diagnostics.Process]::Start($info)
        $process.StandardInput.Close()
        $child = [pscustomobject]@{ Process = $process; Output = $process.StandardOutput.BaseStream; Errors = $process.StandardError.BaseStream }
    }
    return [pscustomobject]@{
        Process = $child.Process
        Output = (New-PcmGuiStreamCollector $child.Output ($base + '.stdout.txt'))
        Errors = (New-PcmGuiStreamCollector $child.Errors ($base + '.stderr.txt'))
    }
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB7e5uxE2Dg0K4i
# IQXCvbM6KuN476Cl5VT1Q4HCuW4XT6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDgiwYrBnDJ
# Qaby0WFxAyD/6sPJUJ3rY/lBbAfXWUMb/jANBgkqhkiG9w0BAQEFAASCAgAsxw9P
# AqdviZd3PY7Cqhzf1Nknnpwjn443ZbvxfuJMIMyYQmf5w+lZN3nJqd8H6gG0VPtW
# bSSKj/AxvNKGRMnYBLc5qzXfJr2FuST+iYf6QUsipu8ExhjVJrSevVMXvMHk0SBr
# u0ZgW8AbsEPGJyB641zvsmdMnBa4tSe0OZ+738PirAx1/DAOSGS/BeYBIjlvssE6
# ZwT/jUi+3kMO/PBDvFK3Pqw9kIkO5fxVY37tPyMVPueShzp+cdgpPQO3FIq8jp3i
# RJh/MK0AcqN+GsHPjMafZbL2TNOSB7gGJ7dCGFfE5SutBg0iV+xVVuyNse4FSVrV
# 4NlYouet8SrrSejlI8LU8pBmk4mKxRWhsDOsRmiiErkfnYVoK7HEIwEtRTeeuxWd
# O8Bj7T8spvHcyP8DekxM6dks22M/tUJBW/em9zwkb/Af0F6E+Z7dSzF9OqYGyb+i
# gql0HkgXB0pMj4XgqUpQGYGCn1k/Zz3AvtRESer5gkYd48YEgVA1tkUkdJxROXzM
# i0R+cKLp6zW/gb09xpknj0ccjXs/2fGHnG60VfcUR8rszw5UF/wgdXgafsKVz4ZM
# ScXGsKDAbYWTyJ/TRZ8CYdmwWkiDeaW5U9w2Mn2actKo71yEctC+VSAaKeHjDQ2J
# WmpjUxfGABVvr8d02LP0XRrtxl198odEtrjwl6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIOBGMyqdYDYo4e8qq28SvPS8iwe29Atx4XsKW8geBCjMAgh9Qg+Dlbmw
# pxgPMjAyNjExMTQwODE3MzVaMAMCAQECEA3j5JqNwZ4h7oog51ph4zWgghGnMIIE
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
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNDA4MTczNVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEICP5
# ejD4IPtFbIt5AtZH1BGiDnpF6ocbLqdvuQxHZj18MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiBwZZIi8R3QYPXAwxEvg4YzAC91+znRZ3g/Kv9lAgt6BAIhAMhP
# RERvLay0nii/pgiWjucVUucQPCbfZKuW+HkccoHs
# SIG # End signature block
