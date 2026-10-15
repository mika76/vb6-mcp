# vb6-mcp uninstaller (Phase 6). Run it as yourself:
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File "%ProgramFiles(x86)%\vb6mcp\uninstall.ps1"
#
# Reverses install.ps1:
#   Your part: removes the Claude Code hook and MCP registration (user scope) and VB6's add-in entry (HKCU).
#   Machine part (one UAC prompt): stops the watchdog, unregisters the add-in DLL, deletes "%ProgramFiles(x86)%\vb6mcp".
# Close VB6 first (it has the add-in loaded); the script refuses while VB6 runs unless -Force. Files still in use are
# left for Windows to delete and reported. Your data in %LOCALAPPDATA%\vb6mcp (logs, recovery copies, audit log) is
# kept unless -RemoveData.
param(
    [string]$Name = 'vb6',
    [switch]$Force,
    [switch]$RemoveData,
    [switch]$Machine,
    [string]$Log = ''
)
$ErrorActionPreference = 'Stop'
$dest = Join-Path ${env:ProgramFiles(x86)} 'vb6mcp'
$progId = 'VB6McpBridge.Connect'

if ($Machine) {
    Start-Transcript -LiteralPath $Log -Force | Out-Null
    try {
        Get-CimInstance Win32_Process -Filter "Name='vb6mcp.exe'" | Where-Object { $_.CommandLine -match '\swatchdog\b' } | ForEach-Object {
            Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
            Write-Output "  stopped watchdog (pid $($_.ProcessId))"
        }
        $dll = Join-Path $dest 'VB6McpBridge.dll'
        if (Test-Path $dll) {
            $p = Start-Process -FilePath (Join-Path $env:windir 'SysWOW64\regsvr32.exe') -ArgumentList '/s', '/u', "`"$dll`"" -Wait -PassThru
            Write-Output "  unregistered the add-in DLL (regsvr32 exit $($p.ExitCode))"
        }
        if (Test-Path $dest) {
            # The script itself lives in $dest: deleting a running script file is fine on Windows (it's read into memory)
            Get-ChildItem -LiteralPath $dest -Recurse -File | ForEach-Object {
                $file = $_.FullName                 # inside catch, $_ is the error record
                try { Remove-Item -LiteralPath $file -Force }
                catch { Write-Output "  in use, not deleted: $file (delete it after the program using it has ended; a later install replaces it)" }
            }
            try { Remove-Item -LiteralPath $dest -Recurse -Force; Write-Output "  deleted $dest" } catch { Write-Output "  $dest not empty yet (files in use); delete it after a restart" }
        }
        exit 0
    } catch {
        Write-Output "ERROR: $($_.Exception.Message)"
        exit 1
    } finally { Stop-Transcript | Out-Null }
}

if ((Get-Process VB6 -ErrorAction SilentlyContinue) -and -not $Force) {
    throw 'VB6 is running with the add-in loaded: close it first (or pass -Force; the DLL then stays until VB6 exits)'
}
$exe = Join-Path $dest 'vb6mcp.exe'
if (Test-Path $exe) { & $exe uninstall-hook | ForEach-Object { Write-Output "  $_" } }
if (Get-Command claude -ErrorAction SilentlyContinue) {
    $ErrorActionPreference = 'Continue'
    claude mcp remove $Name -s user *> $null
    if ($LASTEXITCODE -eq 0) { Write-Output "  removed Claude Code MCP server '$Name' (user scope)" }
    $ErrorActionPreference = 'Stop'
}
$k = 'HKCU:\Software\Microsoft\Visual Basic\6.0\Addins\' + $progId
if (Test-Path $k) { Remove-Item $k -Recurse -Force; Write-Output "  removed VB6 add-in entry $k" }

Write-Output 'Removing the Program Files part (a UAC prompt follows)...'
$log = Join-Path $env:TEMP "vb6mcp-uninstall-$(Get-Date -Format yyyyMMddHHmmss).log"
$self = $MyInvocation.MyCommand.Path
# Run the machine part from a temp copy: this script may live in the folder it deletes
$copy = Join-Path $env:TEMP 'vb6mcp-uninstall.ps1'
Copy-Item -LiteralPath $self -Destination $copy -Force
$shell = (Get-Process -Id $PID).Path
$argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$copy`"", '-Machine', '-Log', "`"$log`"")
try { $p = Start-Process -FilePath $shell -ArgumentList $argList -Verb RunAs -Wait -PassThru -WindowStyle Hidden }
catch { throw "the elevated part didn't start (UAC declined?): $($_.Exception.Message). Program Files and the COM registration were left in place" }
if (Test-Path $log) { Get-Content $log | Where-Object { $_ -match '^\s{2}|^ERROR' } | ForEach-Object { Write-Output $_ } }
Remove-Item -LiteralPath $copy -ErrorAction SilentlyContinue
if ($p.ExitCode -ne 0) { throw "the machine part failed (log: $log)" }

if ($RemoveData) {
    $data = Join-Path $env:LOCALAPPDATA 'vb6mcp'
    if (Test-Path $data) { Remove-Item -LiteralPath $data -Recurse -Force; Write-Output "  deleted $data" }
}
Write-Output 'vb6-mcp uninstalled. Restart Claude Code so it stops the server.'
