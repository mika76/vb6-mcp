Add-Type -TypeDefinition @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic; using System.Diagnostics;
public static class U {
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc cb, IntPtr l);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)] public static extern IntPtr SendMessageTimeout(IntPtr h, uint m, IntPtr w, StringBuilder l, uint f, uint t, out IntPtr r);
  [DllImport("user32.dll", SetLastError=true)] public static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte sc, uint f, UIntPtr e);
  [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr OpenProcess(uint a, bool i, uint pid);
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool OpenProcessToken(IntPtr p, uint a, out IntPtr t);
  [DllImport("advapi32.dll", SetLastError=true)] public static extern bool GetTokenInformation(IntPtr t, int c, IntPtr b, int l, out int r);
  [DllImport("advapi32.dll")] public static extern IntPtr GetSidSubAuthority(IntPtr s, uint i);
  [DllImport("advapi32.dll")] public static extern IntPtr GetSidSubAuthorityCount(IntPtr s);
  public static string IL(uint pid){
    IntPtr p=OpenProcess(0x1000,false,pid); if(p==IntPtr.Zero) return "OpenProcess failed err="+Marshal.GetLastWin32Error();
    IntPtr t; if(!OpenProcessToken(p,8,out t)) return "OpenProcessToken failed err="+Marshal.GetLastWin32Error();
    int n; GetTokenInformation(t,25,IntPtr.Zero,0,out n); IntPtr b=Marshal.AllocHGlobal(n);
    if(!GetTokenInformation(t,25,b,n,out n)) return "GTI failed";
    IntPtr sid=Marshal.ReadIntPtr(b); int c=Marshal.ReadByte(GetSidSubAuthorityCount(sid));
    int rid=Marshal.ReadInt32(GetSidSubAuthority(sid,(uint)(c-1)));
    return rid>=0x3000?"High (elevated) 0x"+rid.ToString("X"):rid>=0x2000?"Medium 0x"+rid.ToString("X"):"Low 0x"+rid.ToString("X");
  }
  public static List<IntPtr> Find(string cls, uint pidFilter){ var l=new List<IntPtr>();
    EnumWindows((h,x)=>{ var s=new StringBuilder(256); GetClassName(h,s,256); uint pid; GetWindowThreadProcessId(h,out pid);
      if(s.ToString()==cls && IsWindowVisible(h) && (pidFilter==0||pid==pidFilter)) l.Add(h); return true;},IntPtr.Zero); return l; }
  public static string Title(IntPtr h){var s=new StringBuilder(512);GetWindowText(h,s,512);return s.ToString();}
}
"@
"Me (this shell): " + [U]::IL($PID)
$claude = Get-Process -Name claude -ErrorAction SilentlyContinue | Select-Object -First 1
if($claude){ "Claude app ($($claude.Id)): " + [U]::IL($claude.Id) }
$vb = Get-Process -Name VB6 -ErrorAction SilentlyContinue
if(-not $vb){ "No VB6 process found"; exit }
foreach($p in $vb){ "VB6 pid $($p.Id): " + [U]::IL($p.Id) }
$vbpid=[uint32]$vb[0].Id
$main = [U]::Find("wndclass_desked_gsk",$vbpid)
if($main.Count -eq 0){ "VB6 main window not found"; exit }
$h=$main[0]; "VB6 main window: 0x{0:X} '{1}'" -f [int64]$h, [U]::Title($h)

# Test 1: SendMessage WM_GETTEXT (0x0D)
$sb=New-Object System.Text.StringBuilder 512; $r=[IntPtr]::Zero
$ok=[U]::SendMessageTimeout($h,0x0D,[IntPtr]512,$sb,2,1000,[ref]$r); $e=[Runtime.InteropServices.Marshal]::GetLastWin32Error()
"T1 SendMessage WM_GETTEXT: ok=$($ok -ne [IntPtr]::Zero) err=$e text='$sb'"

# Test 2: PostMessage WM_MOUSEMOVE (0x200), harmless
$ok=[U]::PostMessage($h,0x200,[IntPtr]0,[IntPtr]0); $e=[Runtime.InteropServices.Marshal]::GetLastWin32Error()
"T2 PostMessage WM_MOUSEMOVE: ok=$ok err=$e  (err 5 = blocked by UIPI)"

# Test 3: synthetic input: Alt+F should open VB6's File menu (#32768 popup)
$before=([U]::Find("#32768",$vbpid)).Count
$fg=[U]::SetForegroundWindow($h); Start-Sleep -Milliseconds 400
$isFg = ([U]::GetForegroundWindow() -eq $h)
"T3 SetForegroundWindow ok=$fg, VB6 is foreground=$isFg"
[U]::keybd_event(0x12,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x46,0,0,[UIntPtr]::Zero)
[U]::keybd_event(0x46,0,2,[UIntPtr]::Zero); [U]::keybd_event(0x12,0,2,[UIntPtr]::Zero)
Start-Sleep -Milliseconds 600
$after=([U]::Find("#32768",$vbpid)).Count
"T3 synthetic Alt+F: menu popups before=$before after=$after  -> " + $(if($after -gt $before){"INPUT DELIVERED"}elseif($isFg){"INPUT BLOCKED (VB6 was foreground)"}else{"INCONCLUSIVE (VB6 not foreground)"})
# cleanup: Escape twice
1..2 | % { [U]::keybd_event(0x1B,0,0,[UIntPtr]::Zero); [U]::keybd_event(0x1B,0,2,[UIntPtr]::Zero); Start-Sleep -Milliseconds 150 }
"after Escape: menu popups=" + ([U]::Find("#32768",$vbpid)).Count
