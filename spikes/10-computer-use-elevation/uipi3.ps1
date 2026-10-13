. "$PSScriptRoot/uipi-types.ps1"
Add-Type -TypeDefinition @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class U3 {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  public static List<string> Snap(uint p){ var l=new List<string>(); EnumWindows((h,x)=>{uint pid;GetWindowThreadProcessId(h,out pid);
    if(pid==p&&IsWindowVisible(h)){var s=new StringBuilder(256);GetClassName(h,s,256);l.Add(h.ToString("X")+" "+s);} return true;},IntPtr.Zero); return l;}
}
"@
$vbpid=[uint32](Get-Process VB6)[0].Id
$h=([U]::Find("wndclass_desked_gsk",$vbpid))[0]
$fgw=[U]::GetForegroundWindow(); [uint32]$x=0; $t=[U]::GetWindowThreadProcessId($fgw,[ref]$x); $me=[U2]::GetCurrentThreadId()
[U2]::AttachThreadInput($me,$t,$true)|Out-Null; [U2]::BringWindowToTop($h)|Out-Null; [U]::SetForegroundWindow($h)|Out-Null; [U2]::AttachThreadInput($me,$t,$false)|Out-Null
Start-Sleep -Milliseconds 500
"VB6 foreground=" + ([U]::GetForegroundWindow() -eq $h)
$before=[U3]::Snap($vbpid)
[U]::keybd_event(0x12,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x46,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x46,0,2,[UIntPtr]::Zero); [U]::keybd_event(0x12,0,2,[UIntPtr]::Zero)
Start-Sleep -Milliseconds 800
$after=[U3]::Snap($vbpid)
"visible VB6 top-level windows before=$($before.Count) after=$($after.Count)"
"new windows after Alt+F: " + (($after | ? { $before -notcontains $_ }) -join ', ')
"still foreground=" + ([U]::GetForegroundWindow() -eq $h)
1..2 | % { [U]::keybd_event(0x1B,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x1B,0,2,[UIntPtr]::Zero); Start-Sleep -Milliseconds 150 }
"after Escape: " + ([U3]::Snap($vbpid)).Count
