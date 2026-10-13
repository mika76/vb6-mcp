# Publishes vb6mcp.exe and registers it with Claude Code (local scope = this project folder only).
# Runs non-elevated.
#
# Each install goes to its own folder build\server\<yyyyMMdd-HHmmss>\ : Claude Code keeps the server
# running, which locks its exe, so publishing over the previous copy would fail while a session is
# open. The new path is registered (replacing the old registration); older folders are deleted when
# they're no longer in use. Restart the Claude session afterwards so it starts the new copy.
#
# Registration notes (spike 1): a project .mcp.json didn't load in the desktop app, local scope does;
# `-e` is variadic, so the server name must come before it; local-scope servers don't appear in the
# app's Connectors UI (check with `claude mcp get vb6`).
#
# -Roots: workspace folders for the disk tools (ws_projects, ws_file_projects, file_read), stored as
# VB6MCP_ROOTS. Paths outside them are refused (AGENTS.md hard rule 8).
param([string]$Name = 'vb6', [string[]]$Roots = @())
$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSScriptRoot
$servers = Join-Path $root 'build\server'
$out     = Join-Path $servers (Get-Date -Format 'yyyyMMdd-HHmmss')
$exe     = Join-Path $out 'vb6mcp.exe'
$log     = Join-Path $root 'build\vb6mcp.log'

dotnet publish (Join-Path $root 'src\Vb6Mcp\Vb6Mcp.csproj') -c Release -o $out --nologo -v q
if ($LASTEXITCODE) { throw 'dotnet publish failed' }

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
    claude mcp add --scope local $Name @envArgs -- $exe
    if ($LASTEXITCODE) { throw "claude mcp add failed ($LASTEXITCODE)" }
    claude mcp get $Name
} finally { Pop-Location; $ErrorActionPreference = 'Stop' }

# Remove older installs that aren't running (a running exe can't be deleted; those stay until next time)
Get-ChildItem $servers -Directory | Where-Object { $_.FullName -ne $out } | ForEach-Object {
    $old = $_   # inside catch, $_ is the error record
    try { Remove-Item -LiteralPath $old.FullName -Recurse -ErrorAction Stop; Write-Host "removed old install $($old.Name)" }
    catch { Write-Host "kept old install $($old.Name) (in use)" }
}
Write-Host "Installed $exe as MCP server '$Name'. Restart the Claude session to use it."
