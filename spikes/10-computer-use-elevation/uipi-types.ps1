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
Add-Type -TypeDefinition @"
using System; using System.Runtime.InteropServices;
public static class U2 {
  [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
  [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool f);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
}
"@
