# Publishes vb6mcp.exe and registers it with Claude Code (local scope = this project folder only).
# Runs non-elevated.
#
# Installs into one fixed folder, build\server\current\, so the registered path never changes (PLAN.md E7):
# the desktop app keeps the command it read at session start and restarts an ended server from it on the
# next call (tested 2026-10-17), so after an install the tool server_restart brings up the new build without
# restarting the session. Claude Code keeps the server running, which locks its exe and DLLs: they can't be
# overwritten but can be renamed (spike 26), so files in use are moved aside as *.old-<stamp> and deleted by
# a later install. (Earlier versions installed each build into its own build\server\<stamp>\ folder and
# registered that path, which needed a session restart; those folders are removed when no longer in use.)
#
# Registration notes (spike 1): a project .mcp.json didn't load in the desktop app, local scope does;
# `-e` is variadic, so the server name must come before it; local-scope servers don't appear in the
# app's Connectors UI (check with `claude mcp get vb6`).
#
# -Roots: workspace folders for the disk tools (ws_projects, ws_file_projects, file_read), stored as
# VB6MCP_ROOTS. Paths outside them are refused (AGENTS.md hard rule 8).
# -NoIdleTimeout: don't set CLAUDE_CODE_MCP_TOOL_IDLE_TIMEOUT (2 minutes) in Claude Code's settings env (spike 42).
# -NoHook: don't register the Claude Code PreToolUse hook (it blocks Edit/Write on VB6 source files in
# every project, so agents use the vb6 tools; registered in the user's ~/.claude/settings.json).
# -EvalOpen: let debug_eval run expressions with side effects (database writes, Shell, files, registry...), flagged and
# audited (VB6MCP_EVAL=open). By default they are refused (PLAN.md E5).
param([string]$Name = 'vb6', [string[]]$Roots = @(), [switch]$NoHook, [switch]$EvalOpen, [switch]$NoIdleTimeout)
$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSScriptRoot
$servers = Join-Path $root 'build\server'
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$staging = Join-Path $servers "staging-$stamp"
$out     = Join-Path $servers 'current'
$exe     = Join-Path $out 'vb6mcp.exe'
$log     = Join-Path $root 'build\vb6mcp.log'

# Workspace roots. The add-in reads %LOCALAPPDATA%\vb6mcp\roots.txt for build_make (an elevated /make only builds
# projects under them); the server gets the same list as VB6MCP_ROOTS. Without -Roots the existing file is kept and
# reused: writing an empty list there crashed this script (null array) after it had already removed the old
# registration, leaving no server at all. So the roots are worked out before anything is changed.
$rootsFile = Join-Path $env:LOCALAPPDATA 'vb6mcp\roots.txt'
$Roots = @($Roots | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })   # -File passes "a,b" as one string
if ($Roots.Count -gt 0) {
    $Roots = @($Roots | ForEach-Object { [IO.Path]::GetFullPath($_) })
    New-Item -ItemType Directory -Force (Split-Path $rootsFile) | Out-Null
    [IO.File]::WriteAllLines($rootsFile, [string[]]$Roots, [Text.Encoding]::GetEncoding(1252))   # VB reads ANSI
} elseif (Test-Path $rootsFile) {
    $Roots = @([IO.File]::ReadAllLines($rootsFile, [Text.Encoding]::GetEncoding(1252)) | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    Write-Host "Workspace roots (kept from $rootsFile): $($Roots -join '; ')"
} else {
    Write-Warning "No workspace roots: pass -Roots D:\Projects,... (disk tools and build_make refuse paths outside them)"
}

dotnet publish (Join-Path $root 'src\Vb6Mcp\Vb6Mcp.csproj') -c Release -o $staging --nologo -v q
if ($LASTEXITCODE) { throw 'dotnet publish failed' }

# Copy into current\: a file in use (the running server's exe or DLLs) is renamed aside first
New-Item -ItemType Directory -Force $out | Out-Null
$moved = 0
Get-ChildItem $staging -Recurse -File | ForEach-Object {
    $src = $_   # inside catch, $_ is the error record
    $dst = Join-Path $out $src.FullName.Substring($staging.Length + 1)
    New-Item -ItemType Directory -Force (Split-Path $dst) | Out-Null
    if (Test-Path -LiteralPath $dst) {
        try { Remove-Item -LiteralPath $dst -ErrorAction Stop }
        catch { Rename-Item -LiteralPath $dst -NewName ((Split-Path $dst -Leaf) + ".old-$stamp"); $moved++ }
    }
    Copy-Item -LiteralPath $src.FullName -Destination $dst
}
Remove-Item -LiteralPath $staging -Recurse -Force
if ($moved -gt 0) { Write-Host "$moved file(s) in use were moved aside; a running server keeps its copy until it restarts" }

Push-Location $root
# 'claude mcp remove' writes to stderr when there is nothing to remove; with ErrorActionPreference
# 'Stop' Windows PowerShell 5.1 turns that into a terminating error, which once stopped the script
# after removing the old registration and before adding the new one.
$ErrorActionPreference = 'Continue'
try {
    claude mcp remove $Name -s local *> $null
    claude mcp remove vb6spike -s local *> $null      # spike 1 test server
    $envArgs = @('-e', "VB6MCP_LOG=$log")
    if ($Roots.Count -gt 0) { $envArgs += @('-e', "VB6MCP_ROOTS=$($Roots -join ';')") }
    if ($EvalOpen) { $envArgs += @('-e', 'VB6MCP_EVAL=open') }
    claude mcp add --scope local $Name @envArgs -- $exe
    if ($LASTEXITCODE) { throw "claude mcp add failed ($LASTEXITCODE)" }
    claude mcp get $Name
} finally { Pop-Location; $ErrorActionPreference = 'Stop' }

# Remove files moved aside by earlier installs, and the per-install folders of older versions of this script,
# once nothing uses them (a running exe can't be deleted; those stay until next time)
Get-ChildItem $out -Recurse -File -Filter '*.old-*' | Where-Object { $_.Name -notlike "*.old-$stamp" } | ForEach-Object {
    $old = $_
    try { Remove-Item -LiteralPath $old.FullName -ErrorAction Stop } catch { }
}
Get-ChildItem $servers -Directory | Where-Object { $_.FullName -ne $out } | ForEach-Object {
    $old = $_
    try { Remove-Item -LiteralPath $old.FullName -Recurse -ErrorAction Stop; Write-Host "removed old install $($old.Name)" }
    catch { Write-Host "kept old install $($old.Name) (in use)" }
}
# the hook step also sets Claude Code's idle timeout to 2 minutes, unless set already or -NoIdleTimeout (spike 42)
if (-not $NoHook) { if ($NoIdleTimeout) { & $exe install-hook --no-idle-timeout } else { & $exe install-hook } }
Write-Host "Installed $exe as MCP server '$Name'. To use it: ask the agent to call server_restart (the next vb6 call starts the new build), or restart the Claude session."
