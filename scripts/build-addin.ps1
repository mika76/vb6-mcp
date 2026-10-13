# Builds the VB6McpBridge add-in with VB6.EXE /make. MUST run elevated: building an ActiveX DLL
# registers it (HKLM), and starting VB6.EXE from a non-elevated process triggers a UAC prompt
# (spike 2). Output: addin\VB6McpBridge\build\VB6McpBridge.dll (gitignored).
#
# Binary compatibility: the .vbp points at addin\compat\VB6McpBridge.dll so the ProgID and CLSIDs
# stay stable across builds (AGENTS.md). If that reference copy doesn't exist yet (first build ever),
# this script builds once without compatibility and stores the result as the reference.
#
# Before building: unload the add-in in VB6 (Add-In Manager) or close VB6, otherwise the DLL is locked.
$ErrorActionPreference = 'Stop'
$root   = Split-Path -Parent $PSScriptRoot
$proj   = Join-Path $root 'addin\VB6McpBridge\VB6McpBridge.vbp'
$outDir = Join-Path $root 'addin\VB6McpBridge\build'
$compat = Join-Path $root 'addin\compat\VB6McpBridge.dll'
$vb6    = 'C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE'

$elevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $elevated) { throw 'Run this from an elevated (Administrator) PowerShell: building registers the DLL.' }
New-Item -ItemType Directory -Force $outDir | Out-Null

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

# Register the add-in with VB6 only now that the DLL's ProgID exists: a VB6 instance that starts
# while the Addins key points at an unregistered ProgID deletes the key (found in Phase 1; the
# build's own headless VB6 did exactly that).
& (Join-Path $PSScriptRoot 'register-addin.ps1')
