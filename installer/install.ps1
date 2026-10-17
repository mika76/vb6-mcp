# vb6-mcp installer (Phase 6). Run it as yourself from the unpacked package folder:
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -Roots D:\src,E:\work
#
# What it does:
#   Machine part (one UAC prompt):
#     - copies VB6McpBridge.dll, vb6mcp.exe and roots.txt to "%ProgramFiles(x86)%\vb6mcp".
#       Only administrators can write there. That matters: the add-in and the watchdog run elevated
#       inside/next to VB6, and the roots file says what they may build, register and load (AGENTS.md
#       hard rule 8). In a user-writable folder any program could swap them or widen the roots.
#     - registers the add-in DLL (32-bit regsvr32).
#   Your part (no elevation):
#     - tells VB6 about the add-in (HKCU\Software\Microsoft\Visual Basic\6.0\Addins).
#     - registers the MCP server with Claude Code for your user (claude mcp add --scope user).
#     - installs the Claude Code hook that keeps agents from editing VB6 files with plain Edit/Write.
#
# Running it again upgrades in place: files in use (a running VB6, Claude Code or watchdog) are renamed
# aside and replaced, and the new version is used after VB6 and Claude Code restart. Without -Roots an
# upgrade keeps the existing roots. Works with Windows PowerShell 5.1 and PowerShell 7.
param(
    [string[]]$Roots = @(),
    [string]$Name = 'vb6',
    [switch]$NoHook,
    [switch]$NoClaude,
    # let debug_eval run expressions with side effects, flagged and audited (VB6MCP_EVAL=open); refused by default
    [switch]$EvalOpen,
    # internal: the elevated machine part
    [switch]$Machine,
    [string]$Log = ''
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dest = Join-Path ${env:ProgramFiles(x86)} 'vb6mcp'
$progId = 'VB6McpBridge.Connect'

function Test-Elevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Copies a file over one that may be in use: a running exe or a loaded DLL can't be overwritten, but it can
# be renamed, so the old file moves aside (vb6mcp.exe.old-<time>) and is deleted when it's free again.
function Install-File([string]$source, [string]$target) {
    if (Test-Path -LiteralPath $target) {
        try { Remove-Item -LiteralPath $target -Force }
        catch {
            $aside = "$target.old-$(Get-Date -Format yyyyMMddHHmmss)"
            Move-Item -LiteralPath $target -Destination $aside
            Write-Output "  $([IO.Path]::GetFileName($target)) is in use: moved aside to $([IO.Path]::GetFileName($aside))"
        }
    }
    Copy-Item -LiteralPath $source -Destination $target
    Unblock-File -LiteralPath $target          # drop the downloaded-from-the-internet mark the zip may carry
}

if ($Machine) {
    # ---------------- elevated part ----------------
    Start-Transcript -LiteralPath $Log -Force | Out-Null
    $Roots = @($Roots | ForEach-Object { $_ -split ';' } | Where-Object { $_ })
    try {
        New-Item -ItemType Directory -Force $dest | Out-Null
        Get-ChildItem -LiteralPath $dest -Filter '*.old-*' -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force } catch { }      # leftovers of earlier upgrades, if free now
        }
        if ($Roots.Count -eq 0 -and -not (Test-Path (Join-Path $dest 'roots.txt'))) { throw 'no workspace roots: pass -Roots (the folders that contain your VB6 projects)' }   # before copying anything
        foreach ($f in 'VB6McpBridge.dll', 'vb6mcp.exe', 'uninstall.ps1') {
            Install-File (Join-Path $here $f) (Join-Path $dest $f)
            Write-Output "  installed $f"
        }
        if ($Roots.Count -gt 0) {
            # VB reads ANSI; one absolute folder per line. Read by the add-in (next to its DLL) and the server (next to its exe).
            [IO.File]::WriteAllLines((Join-Path $dest 'roots.txt'), [string[]]$Roots, [Text.Encoding]::GetEncoding(1252))
            Write-Output "  roots: $($Roots -join '; ')"
        } elseif (-not (Test-Path (Join-Path $dest 'roots.txt'))) {
            throw 'no workspace roots: pass -Roots (the folders that contain your VB6 projects)'
        } else {
            Write-Output "  roots kept: $((Get-Content (Join-Path $dest 'roots.txt')) -join '; ')"
        }
        # A watchdog from an earlier install or a development build keeps running (one per session, named mutex);
        # stop it so the new add-in starts the new one. It keeps no state that isn't on disk.
        Get-CimInstance Win32_Process -Filter "Name='vb6mcp.exe'" | Where-Object { $_.CommandLine -match '\swatchdog\b' } | ForEach-Object {
            Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
            Write-Output "  stopped running watchdog (pid $($_.ProcessId))"
        }
        $regsvr = Join-Path $env:windir 'SysWOW64\regsvr32.exe'      # VB6 is 32-bit: the 32-bit registry view
        $p = Start-Process -FilePath $regsvr -ArgumentList '/s', "`"$(Join-Path $dest 'VB6McpBridge.dll')`"" -Wait -PassThru
        if ($p.ExitCode -ne 0) { throw "regsvr32 failed with exit code $($p.ExitCode)" }
        # Check the outcome, not the exit code alone: the ProgID's class must point at the installed DLL
        $hkcr = [Microsoft.Win32.RegistryKey]::OpenBaseKey('ClassesRoot', 'Registry32')
        $clsid = $hkcr.OpenSubKey("$progId\CLSID").GetValue('')
        $server = $hkcr.OpenSubKey("CLSID\$clsid\InprocServer32").GetValue('')
        if (-not $server -or ([IO.Path]::GetFullPath($server) -ne [IO.Path]::GetFullPath((Join-Path $dest 'VB6McpBridge.dll')))) {
            throw "registration check failed: $progId points at '$server'"
        }
        Write-Output "  registered $progId -> $server"
        exit 0
    } catch {
        Write-Output "ERROR: $($_.Exception.Message)"
        exit 1
    } finally { Stop-Transcript | Out-Null }
}

# ---------------- your part ----------------
foreach ($f in 'VB6McpBridge.dll', 'vb6mcp.exe', 'uninstall.ps1') {
    if (-not (Test-Path (Join-Path $here $f))) { throw "$f is missing next to install.ps1: run it from the unpacked package folder" }
}
$vb6Dir = $null
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey('LocalMachine', 'Registry32').OpenSubKey('SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic')
if ($key) { $vb6Dir = $key.GetValue('ProductDir') }
if (-not $vb6Dir) { $vb6Dir = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\VB98' }
if (-not (Test-Path (Join-Path $vb6Dir 'VB6.EXE'))) { Write-Warning "VB6.EXE not found in $vb6Dir; installing anyway (set VB6MCP_VB6_EXE if it lives elsewhere)" }

# with -File, "-Roots D:\src,E:\work" arrives as one string: split it (commas or semicolons)
$Roots = @($Roots | ForEach-Object { $_ -split '[;,]' } | Where-Object { $_.Trim() })
$fullRoots = @()
foreach ($r in $Roots) {
    $full = [IO.Path]::GetFullPath($r).TrimEnd('\')
    if (-not (Test-Path -LiteralPath $full -PathType Container)) { throw "root folder not found: $full" }
    $fullRoots += $full
}

# Check before elevating, so a missing argument costs no UAC prompt and leaves no half-copied install
if ($fullRoots.Count -eq 0 -and -not (Test-Path (Join-Path $dest 'roots.txt'))) {
    throw 'no workspace roots: pass -Roots with the folders that contain your VB6 projects, e.g. -Roots "D:\src;E:\work"'
}
Write-Output 'Installing to Program Files (a UAC prompt follows)...'
$log = Join-Path $env:TEMP "vb6mcp-install-$(Get-Date -Format yyyyMMddHHmmss).log"
$argList = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$($MyInvocation.MyCommand.Path)`"", '-Machine', '-Log', "`"$log`"")
if ($fullRoots.Count -gt 0) { $argList += @('-Roots', "`"$($fullRoots -join ';')`"") }    # -File doesn't parse arrays: one string, split below
$shell = (Get-Process -Id $PID).Path                                  # the same PowerShell this runs in
if (Test-Elevated) { & $shell @argList; $code = $LASTEXITCODE }
else {
    try { $p = Start-Process -FilePath $shell -ArgumentList $argList -Verb RunAs -Wait -PassThru -WindowStyle Hidden }
    catch { throw "the elevated part didn't start (UAC declined?): $($_.Exception.Message). Nothing was installed" }
    $code = $p.ExitCode
}
if (Test-Path $log) { Get-Content $log | Where-Object { $_ -match '^\s{2}|^ERROR' } | ForEach-Object { Write-Output $_ } }
if ($code -ne 0) { throw "the machine part failed (log: $log)" }

# VB6 learns about add-ins from HKCU. Only now that the ProgID is registered: a VB6 that starts while this key
# names an unregistered ProgID deletes the key (Phase 1). VB6 rewrites LoadBehavior from the Add-In Manager.
$k = 'HKCU:\Software\Microsoft\Visual Basic\6.0\Addins\' + $progId
New-Item -Force $k | Out-Null
Set-ItemProperty $k -Name FriendlyName -Value 'VB6 MCP Bridge'
Set-ItemProperty $k -Name Description -Value 'Lets the vb6-mcp MCP server read and drive this IDE over loopback TCP.'
Set-ItemProperty $k -Name LoadBehavior -Value 3 -Type DWord
Set-ItemProperty $k -Name CommandLineSafe -Value 0 -Type DWord
Write-Output "  VB6 add-in entry: $k"

$exe = Join-Path $dest 'vb6mcp.exe'
if (-not $NoClaude) {
    if (Get-Command claude -ErrorAction SilentlyContinue) {
        # 'claude mcp remove' writes to stderr when there's nothing to remove; that must not stop the script
        $ErrorActionPreference = 'Continue'
        claude mcp remove $Name -s user *> $null
        # The server's request log (PLAN.md C6): without VB6MCP_LOG it only reaches Claude Code's MCP logs
        $envArgs = @('-e', "VB6MCP_LOG=$(Join-Path $env:LOCALAPPDATA 'vb6mcp\server.log')")
        if ($EvalOpen) { $envArgs += @('-e', 'VB6MCP_EVAL=open') }
        claude mcp add --scope user $Name @envArgs -- $exe
        $added = $LASTEXITCODE
        $ErrorActionPreference = 'Stop'
        if ($added -ne 0) { Write-Warning "claude mcp add failed ($added); register it yourself: claude mcp add --scope user $Name -- `"$exe`"" }
        else { Write-Output "  Claude Code MCP server '$Name' (user scope) -> $exe" }
    } else {
        Write-Warning "Claude Code CLI not found. Register the server yourself: claude mcp add --scope user $Name -- `"$exe`""
    }
}
if (-not $NoHook) { & $exe install-hook | ForEach-Object { Write-Output "  $_" } }

Write-Output ''
Write-Output "vb6-mcp installed in $dest. Restart VB6 (the add-in loads at startup) and Claude Code (it starts the server)."
