using System.Globalization;
using System.Text.RegularExpressions;

namespace Vb6Mcp.Core;

/// <summary>
/// An error from a VB6 /make log. <see cref="Line"/> is 1-based in the code as the IDE shows it
/// (without the hidden header and Attribute lines), like code_read; null when the error has no location.
/// <see cref="Kind"/>: "compile", "compatibility" (binary compatibility break, with <see cref="Class"/>
/// and <see cref="Member"/>), "load" (VB couldn't load a form or control: <see cref="File"/> is the <c>.log</c> VB wrote
/// next to it, whose lines the build tool adds), or "other".
/// </summary>
public sealed record BuildError(string Kind, string? File, int? Line, string Message, string? Class = null, string? Member = null);

/// <summary>A parsed /make log.</summary>
public sealed record BuildLogResult(bool Succeeded, IReadOnlyList<BuildError> Errors);

/// <summary>
/// Parses the log VB6.EXE /make writes with /out (formats from spike 21):
/// <list type="bullet">
/// <item><c>Build of 'Basic.exe' succeeded.</c> / <c>Build of 'Basic.exe' failed.</c></item>
/// <item><c>Compile Error in File '...\modMath.bas', Line 8 : Variable not defined</c>, where the line
/// is the IDE's code line counted from 0 (IDE line 9; file line 10 in a module, 40 in a form)</item>
/// <item><c>There is a compatibility error between the current project and the version-compatible
/// component: 'Hello' in the 'clsGreeter' class module has arguments ...</c> (no dialog, the build fails)</item>
/// <item><c>The file '...\Basic.vbp' is already open.</c> (the project is open in an IDE)</item>
/// </list>
/// The tools always return the raw log too.
/// </summary>
public static partial class BuildLog
{
    public static BuildLogResult Parse(string log)
    {
        var errors = new List<BuildError>();
        bool ok = false;
        foreach (var raw in log.Replace("\r\n", "\n").Split('\n'))
        {
            var line = raw.Trim();
            if (line.Length == 0) continue;
            var status = Status().Match(line);
            if (status.Success) { ok |= status.Groups[1].Value.Equals("succeeded", StringComparison.OrdinalIgnoreCase); continue; }
            var m = ErrorAt().Match(line);
            if (m.Success)
            {
                // /make counts the IDE's code lines from 0 (spike 21: three samples)
                int zeroBased = int.Parse(m.Groups["line"].Value, CultureInfo.InvariantCulture);
                errors.Add(new BuildError("compile", m.Groups["file"].Value, zeroBased + 1, m.Groups["msg"].Value.Trim()));
                continue;
            }
            // "Errors during load. Refer to 'D:\x\frmMain.log' for details" (spike 27: a licensed control missing)
            var load = LoadLog().Match(line);
            if (load.Success) { errors.Add(new BuildError("load", load.Groups["file"].Value, null, line)); continue; }
            var c = Compat().Match(line);
            if (c.Success) { errors.Add(new BuildError("compatibility", null, null, line, c.Groups["cls"].Value, c.Groups["member"].Value)); continue; }
            if (!errors.Any(e => e.Message == line)) errors.Add(new BuildError("other", null, null, line));
        }
        return new BuildLogResult(ok && errors.Count == 0, errors);
    }

    [GeneratedRegex(@"^Build of '.*' (succeeded|failed)\.?$", RegexOptions.IgnoreCase)]
    private static partial Regex Status();

    [GeneratedRegex(@"in File '(?<file>[^']+)', Line (?<line>\d+)\s*:\s*(?<msg>.*)$", RegexOptions.IgnoreCase)]
    private static partial Regex ErrorAt();

    [GeneratedRegex(@"Refer to '(?<file>[^']+\.log)'", RegexOptions.IgnoreCase)]
    private static partial Regex LoadLog();

    [GeneratedRegex(@"compatibility error.*?'(?<member>[^']+)' in the '(?<cls>[^']+)' (class|form|user control|module)", RegexOptions.IgnoreCase)]
    private static partial Regex Compat();
}
