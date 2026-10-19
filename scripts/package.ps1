# Builds the distributable package (Phase 6): dist\vb6mcp-<version>\ and dist\vb6mcp-<version>.zip with
#   VB6McpBridge.dll   the add-in, built by build-addin.ps1 (elevated; this script only checks it's up to date)
#   vb6mcp.exe         MCP server and watchdog: one self-contained single-file win-x86 exe (no .NET needed on the
#                      target; ~48 MB compressed; spike 26)
#   install.ps1, uninstall.ps1, INSTALL.txt
#   LICENSE, CHANGELOG.md
# Runs non-elevated. Usage:  .\scripts\package.ps1
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$version = ([xml](Get-Content (Join-Path $root 'src\Vb6Mcp\Vb6Mcp.csproj'))).Project.PropertyGroup.Version | Where-Object { $_ } | Select-Object -First 1
$out = Join-Path $root "dist\vb6mcp-$version"
$zip = "$out.zip"

# The add-in can only be built elevated (VB6 /make registers it): require a build newer than its source
$dll = Join-Path $root 'addin\VB6McpBridge\build\VB6McpBridge.dll'
if (-not (Test-Path $dll)) { throw "no add-in build ($dll): run scripts\build-addin.ps1 from an elevated PowerShell first" }
$newestSource = Get-ChildItem (Join-Path $root 'addin\VB6McpBridge') -File | Where-Object { $_.Extension -in '.bas', '.cls', '.vbp', '.frm' } |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($newestSource.LastWriteTime -gt (Get-Item $dll).LastWriteTime) {
    throw "the add-in build is older than $($newestSource.Name): run scripts\build-addin.ps1 (elevated) again"
}

if (Test-Path $out) { Remove-Item -LiteralPath $out -Recurse -Force }
New-Item -ItemType Directory -Force $out | Out-Null
Write-Host "Publishing vb6mcp.exe (self-contained, single file)..."
dotnet publish (Join-Path $root 'src\Vb6Mcp\Vb6Mcp.csproj') -c Release -r win-x86 --self-contained true `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:EnableCompressionInSingleFile=true `
    -p:DebugType=none -p:GenerateDocumentationFile=false -o (Join-Path $out 'publish') --nologo -v q
if ($LASTEXITCODE) { throw 'dotnet publish failed' }
Move-Item (Join-Path $out 'publish\vb6mcp.exe') $out
Remove-Item -LiteralPath (Join-Path $out 'publish') -Recurse -Force
Copy-Item $dll $out
Copy-Item (Join-Path $root 'installer\install.ps1'), (Join-Path $root 'installer\uninstall.ps1'), (Join-Path $root 'installer\INSTALL.txt'), (Join-Path $root 'LICENSE'), (Join-Path $root 'CHANGELOG.md') $out

if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $out '*') -DestinationPath $zip
Get-ChildItem $out | ForEach-Object { '{0,12:N0}  {1}' -f $_.Length, $_.Name }
Write-Host "Package: $zip ($('{0:N1}' -f ((Get-Item $zip).Length / 1MB)) MB)"
