using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Vb6Mcp.Core;

/// <summary>
/// Finds VB6.EXE and reads process elevation. Machine-specific values come from configuration or
/// the registry, never from code, so the server works on other people's setups (PLAN §0.2).
/// </summary>
public static class Vb6Install
{
    /// <summary>
    /// VB6.EXE, in order: VB6MCP_VB6_EXE; the installer's ProductDir (32-bit registry view);
    /// the .vbp file association; the default install folder. Null if none exists.
    /// </summary>
    public static string? FindExe()
    {
        foreach (var candidate in Candidates())
            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate)) return Path.GetFullPath(candidate);
        return null;
    }

    private static IEnumerable<string?> Candidates()
    {
        yield return Environment.GetEnvironmentVariable("VB6MCP_VB6_EXE");
        using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
        using (var setup = hklm.OpenSubKey(@"SOFTWARE\Microsoft\VisualStudio\6.0\Setup\Microsoft Visual Basic"))
            if (setup?.GetValue("ProductDir") is string dir) yield return Path.Combine(dir, "VB6.EXE");
        using (var cmd = Registry.ClassesRoot.OpenSubKey(@"VisualBasic.Project\shell\open\command"))
            if (cmd?.GetValue(null) is string open) yield return ExeFromCommand(open);
        yield return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft Visual Studio\VB98\VB6.EXE");
    }

    /// <summary>"C:\...\vb6.exe" "%1" → C:\...\vb6.exe</summary>
    public static string ExeFromCommand(string command)
    {
        var c = command.Trim();
        if (c.StartsWith('"')) { int end = c.IndexOf('"', 1); return end > 0 ? c[1..end] : c.Trim('"'); }
        int space = c.IndexOf(' ');
        return space > 0 ? c[..space] : c;
    }

    /// <summary>True if the process runs at high integrity or above (elevated); null if it can't be read.</summary>
    public static bool? IsElevated(int pid)
    {
        IntPtr proc = OpenProcess(0x1000 /*PROCESS_QUERY_LIMITED_INFORMATION*/, false, pid);
        if (proc == IntPtr.Zero) return null;
        try
        {
            if (!OpenProcessToken(proc, 0x8 /*TOKEN_QUERY*/, out var token)) return null;
            try
            {
                GetTokenInformation(token, 25 /*TokenIntegrityLevel*/, IntPtr.Zero, 0, out int size);
                IntPtr buf = Marshal.AllocHGlobal(size);
                try
                {
                    if (!GetTokenInformation(token, 25, buf, size, out _)) return null;
                    IntPtr sid = Marshal.ReadIntPtr(buf);
                    int count = Marshal.ReadByte(GetSidSubAuthorityCount(sid));
                    return Marshal.ReadInt32(GetSidSubAuthority(sid, (uint)(count - 1))) >= 0x3000;
                }
                finally { Marshal.FreeHGlobal(buf); }
            }
            finally { CloseHandle(token); }
        }
        finally { CloseHandle(proc); }
    }

    [DllImport("kernel32.dll")] private static extern IntPtr OpenProcess(uint access, bool inherit, int pid);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr h);
    [DllImport("advapi32.dll")] private static extern bool OpenProcessToken(IntPtr process, uint access, out IntPtr token);
    [DllImport("advapi32.dll")] private static extern bool GetTokenInformation(IntPtr token, int cls, IntPtr buf, int len, out int ret);
    [DllImport("advapi32.dll")] private static extern IntPtr GetSidSubAuthority(IntPtr sid, uint index);
    [DllImport("advapi32.dll")] private static extern IntPtr GetSidSubAuthorityCount(IntPtr sid);
}
