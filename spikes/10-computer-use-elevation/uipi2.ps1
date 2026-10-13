. "$PSScriptRoot/uipi-types.ps1"
$vbpid=[uint32](Get-Process VB6)[0].Id
$h=([U]::Find("wndclass_desked_gsk",$vbpid))[0]
$fgw=[U]::GetForegroundWindow(); [uint32]$x=0
$fgThread=[U]::GetWindowThreadProcessId($fgw,[ref]$x); $me=[U2]::GetCurrentThreadId()
$att=[U2]::AttachThreadInput($me,$fgThread,$true)
[U2]::ShowWindow($h,9) | Out-Null; [U2]::BringWindowToTop($h) | Out-Null
$fg=[U]::SetForegroundWindow($h)
[U2]::AttachThreadInput($me,$fgThread,$false) | Out-Null
Start-Sleep -Milliseconds 500
$isFg=([U]::GetForegroundWindow() -eq $h)
"attach=$att SetForegroundWindow=$fg VB6 foreground=$isFg"
if(-not $isFg){
  "Waiting up to 20s for VB6 to become the foreground window (click it)..."
  $t=[DateTime]::Now; while(([DateTime]::Now-$t).TotalSeconds -lt 20 -and [U]::GetForegroundWindow() -ne $h){Start-Sleep -Milliseconds 200}
  $isFg=([U]::GetForegroundWindow() -eq $h); "VB6 foreground now=$isFg"; Start-Sleep -Milliseconds 500
}
if(-not $isFg){ "INCONCLUSIVE: could not get VB6 to foreground"; exit }
$before=([U]::Find("#32768",$vbpid)).Count
[U]::keybd_event(0x12,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x46,0,0,[UIntPtr]::Zero)
[U]::keybd_event(0x46,0,2,[UIntPtr]::Zero); [U]::keybd_event(0x12,0,2,[UIntPtr]::Zero)
Start-Sleep -Milliseconds 700
$after=([U]::Find("#32768",$vbpid)).Count
"synthetic Alt+F into elevated VB6: menu popups before=$before after=$after -> " + $(if($after -gt $before){"INPUT DELIVERED"}else{"INPUT BLOCKED"})
1..2 | % { [U]::keybd_event(0x1B,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x1B,0,2,[UIntPtr]::Zero); Start-Sleep -Milliseconds 150 }
"after Escape: menu popups=" + ([U]::Find("#32768",$vbpid)).Count
