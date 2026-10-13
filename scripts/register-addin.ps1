# Tells VB6 about the VB6McpBridge add-in (HKCU, no elevation needed). The COM registration of the
# DLL itself happens during the elevated build (scripts\build-addin.ps1).
#
# Also disables the spike add-in (VB6McpSpike.Connect): both would write the same discovery file
# (named by the VB6 process id).
#
# Notes:
# - VB6 rewrites LoadBehavior from the Add-In Manager checkboxes (spike 12b), so the user's
#   choices there win over this script.
# - Run this only AFTER the DLL is registered (build-addin.ps1 calls it at the end): a VB6 that
#   starts while this key names an unregistered ProgID deletes the key.
# - VB6 reads this list only at startup: restart VB6 to see a newly registered add-in.
$ErrorActionPreference = 'Stop'
$base = 'HKCU:\Software\Microsoft\Visual Basic\6.0\Addins'

$k = Join-Path $base 'VB6McpBridge.Connect'
New-Item -Force $k | Out-Null
Set-ItemProperty $k -Name FriendlyName -Value 'VB6 MCP Bridge'
Set-ItemProperty $k -Name Description -Value 'Lets the vb6-mcp MCP server read and drive this IDE over loopback TCP.'
Set-ItemProperty $k -Name LoadBehavior -Value 3 -Type DWord      # loaded + load on startup
Set-ItemProperty $k -Name CommandLineSafe -Value 0 -Type DWord
Write-Host "Registered add-in: $k (LoadBehavior=3)"

$spike = Join-Path $base 'VB6McpSpike.Connect'
if (Test-Path $spike) {
    Set-ItemProperty $spike -Name LoadBehavior -Value 0 -Type DWord
    Write-Host "Disabled spike add-in: $spike (unload it in Add-In Manager if it's loaded now)"
}
