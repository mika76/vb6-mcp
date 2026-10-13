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
param([string]$Name = 'vb6')
$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $PSScriptRoot
$servers = Join-Path $root 'build\server'
$out     = Join-Path $servers (Get-Date -Format 'yyyyMMdd-HHmmss')
$exe     = Join-Path $out 'vb6mcp.exe'
$log     = Join-Path $root 'build\vb6mcp.log'

dotnet publish (Join-Path $root 'src\Vb6Mcp\Vb6Mcp.csproj') -c Release -o $out --nologo -v q
if ($LASTEXITCODE) { throw 'dotnet publish failed' }

Push-Location $root
try {
    claude mcp remove $Name -s local 2>$null | Out-Null
    claude mcp remove vb6spike -s local 2>$null | Out-Null      # spike 1 test server
    claude mcp add --scope local $Name -e "VB6MCP_LOG=$log" -- $exe
    claude mcp get $Name
} finally { Pop-Location }

# Remove older installs that aren't running (a running exe can't be deleted; those stay until next time)
Get-ChildItem $servers -Directory | Where-Object { $_.FullName -ne $out } | ForEach-Object {
    try { Remove-Item -LiteralPath $_.FullName -Recurse -ErrorAction Stop; Write-Host "removed old install $($_.Name)" }
    catch { Write-Host "kept old install $($_.Name) (in use)" }
}
Write-Host "Installed $exe as MCP server '$Name'. Restart the Claude session to use it."
