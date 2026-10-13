# Read-side window helpers (medium integrity is fine: reads cross the elevation boundary).
# Dot-source: . .\vbwin.ps1 ; then Get-VbTitle <pid>, Get-VbDialogs <pid>, Get-VbImmediate <pid>
if (-not ('VbWin' -as [type])) {
Add-Type -TypeDefinition @"
using System; using System.Text; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class VbWin {
 public delegate bool P(IntPtr h, IntPtr l);
 [DllImport("user32.dll")] static extern bool EnumWindows(P p, IntPtr l);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, P p, IntPtr l);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] static extern bool IsHungAppWindow(IntPtr h);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
 static string Cls(IntPtr h){var s=new StringBuilder(64);GetClassName(h,s,64);return s.ToString();}
 static string Cap(IntPtr h){var s=new StringBuilder(512);GetWindowText(h,s,512);return s.ToString();}
 public static IntPtr Main(uint pid){IntPtr m=IntPtr.Zero;EnumWindows((h,x)=>{uint p;GetWindowThreadProcessId(h,out p);if(p==pid&&Cls(h)=="wndclass_desked_gsk")m=h;return true;},IntPtr.Zero);return m;}
 public static string Title(uint pid){return Cap(Main(pid));}
 public static bool Hung(uint pid){return IsHungAppWindow(Main(pid));}
 public static string Immediate(uint pid){IntPtr r=IntPtr.Zero;EnumChildWindows(Main(pid),(h,x)=>{if(Cls(h)=="VbaWindow"&&Cap(h)=="Immediate")r=h;return true;},IntPtr.Zero);return r.ToInt64().ToString("X");}
 public static List<string> Dialogs(uint pid){var l=new List<string>();EnumWindows((h,x)=>{uint p;GetWindowThreadProcessId(h,out p);if(p==pid&&IsWindowVisible(h)&&Cls(h)=="#32770"){var kids=new List<string>();EnumChildWindows(h,(c,y)=>{if(IsWindowVisible(c))kids.Add(c.ToInt64().ToString("X")+":"+Cls(c)+":'"+Cap(c).Replace("\r\n"," ")+"'");return true;},IntPtr.Zero);l.Add(h.ToInt64().ToString("X")+" '"+Cap(h)+"' -> "+string.Join(" | ",kids));}return true;},IntPtr.Zero);return l;}
}
"@
}
function Get-VbTitle([uint32]$ProcessId) { [VbWin]::Title($ProcessId) }
function Get-VbHung([uint32]$ProcessId) { [VbWin]::Hung($ProcessId) }
function Get-VbDialogs([uint32]$ProcessId) { [VbWin]::Dialogs($ProcessId) }
function Get-VbImmediate([uint32]$ProcessId) { [VbWin]::Immediate($ProcessId) }
