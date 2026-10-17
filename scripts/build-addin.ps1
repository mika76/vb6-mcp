# Builds the VB6McpBridge add-in with VB6.EXE /make. MUST run elevated: building an ActiveX DLL
# registers it (HKLM), and starting VB6.EXE from a non-elevated process triggers a UAC prompt
# (spike 2). Output: addin\VB6McpBridge\build\VB6McpBridge.dll (gitignored).
#
# Binary compatibility: the .vbp points at addin\compat\VB6McpBridge.dll so the ProgID and CLSIDs
# stay stable across builds (AGENTS.md). If that reference copy doesn't exist yet (first build ever),
# this script builds once without compatibility and stores the result as the reference.
#
# Before building: unload the add-in in VB6 (Add-In Manager) or close VB6, otherwise the DLL is locked.
#
# -WatchdogOnly: only republish the watchdog and start it (elevated, like this shell). For when the
# add-in can't be reloaded, e.g. while an IDE is stuck: the add-in starts the watchdog only when it loads.
# -Stay (with -WatchdogOnly): keep the watchdog running with no VB6 open, so Claude sessions can start their own IDEs
# through it (ide_launch) without a UAC prompt each time.
param([switch]$WatchdogOnly, [switch]$Stay)
$ErrorActionPreference = 'Stop'
$root   = Split-Path -Parent $PSScriptRoot
$proj   = Join-Path $root 'addin\VB6McpBridge\VB6McpBridge.vbp'
$outDir = Join-Path $root 'addin\VB6McpBridge\build'
$compat = Join-Path $root 'addin\compat\VB6McpBridge.dll'
# VB6.EXE: VB6MCP_VB6_EXE, else the installer's ProductDir (32-bit view), else the default folder
$vb6 = $env:VB6MCP_VB6_EXE
if (-not $vb6) {
    $dir = (Get-ItemProperty 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic' -ErrorAction SilentlyContinue).ProductDir
    $vb6 = if ($dir) { Join-Path $dir 'VB6.EXE' } else { Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\VB98\VB6.EXE' }
}
if (-not (Test-Path $vb6)) { throw "VB6.EXE not found ($vb6); set VB6MCP_VB6_EXE" }

$elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $elevated) { throw 'Run this from an elevated (Administrator) PowerShell: building registers the DLL.' }
New-Item -ItemType Directory -Force $outDir | Out-Null

function Publish-Watchdog {
    # vb6mcp.exe published next to the DLL, where the add-in looks for it (Sidecar.bas: App.Path\watchdog).
    # It runs elevated, so it lives where the DLL lives. A running watchdog locks its exe: stop it first
    # (it holds no state that isn't on disk; the add-in, or -WatchdogOnly, starts it again).
    $wdDir = Join-Path $outDir 'watchdog'
    Get-Process vb6mcp -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -and $_.Path.StartsWith($wdDir, [StringComparison]::OrdinalIgnoreCase) } |
        ForEach-Object { Write-Host "Stopping running watchdog (pid $($_.Id))"; $_.Kill(); $_.WaitForExit(5000) | Out-Null }
    dotnet publish (Join-Path $root 'src\Vb6Mcp\Vb6Mcp.csproj') -c Release -o $wdDir --nologo -v q | Out-Host
    if ($LASTEXITCODE) { throw 'dotnet publish of the watchdog failed' }
    Write-Host "Watchdog: $(Join-Path $wdDir 'vb6mcp.exe')"
    return (Join-Path $wdDir 'vb6mcp.exe')
}

if ($WatchdogOnly) {
    $exe = Publish-Watchdog
    Start-Process -FilePath $exe -ArgumentList $(if ($Stay) { 'watchdog', '--stay' } else { 'watchdog' }) -WindowStyle Hidden
    Write-Host "Watchdog started$(if ($Stay) { ' (stays running without VB6)' }) (log: %LOCALAPPDATA%\vb6mcp\watchdog.log)"
    return
}

function Invoke-Make([string]$vbp) {
    $log = Join-Path $outDir 'build.log'
    if (Test-Path $log) { Remove-Item -LiteralPath $log }
    & $vb6 /make $vbp /outdir $outDir /out $log | Out-Null
    $text = Get-Content $log -Raw
    Write-Host $text.Trim()
    if ($text -notmatch "Build of '.*' succeeded") { throw "VB6 build failed: $vbp" }
}

if (-not (Test-Path $compat)) {
    Write-Host 'No binary-compatibility reference yet: first build without compatibility.'
    $enc = [Text.Encoding]::GetEncoding(1252)
    $tmp = Join-Path (Split-Path $proj) 'VB6McpBridge.nocompat.vbp'
    $text = [IO.File]::ReadAllText($proj, $enc) -replace 'CompatibleMode="2"', 'CompatibleMode="0"'
    [IO.File]::WriteAllBytes($tmp, $enc.GetBytes($text))
    try { Invoke-Make $tmp } finally { Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue }
    New-Item -ItemType Directory -Force (Split-Path $compat) | Out-Null
    Copy-Item (Join-Path $outDir 'VB6McpBridge.dll') $compat
    Write-Host "Stored binary-compatibility reference: $compat (commit it)"
}

Invoke-Make $proj
Write-Host "Built: $(Join-Path $outDir 'VB6McpBridge.dll')"

Publish-Watchdog | Out-Null

# Register the add-in with VB6 only now that the DLL's ProgID exists: a VB6 instance that starts
# while the Addins key points at an unregistered ProgID deletes the key (found in Phase 1; the
# build's own headless VB6 did exactly that).
& (Join-Path $PSScriptRoot 'register-addin.ps1')
