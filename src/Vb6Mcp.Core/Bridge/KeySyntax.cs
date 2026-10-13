using System.Globalization;
using System.Text;

namespace Vb6Mcp.Core.Bridge;

/// <summary>
/// The key syntax of app_keys, turned into the add-in's key events (docs/protocol.md "KEYS"). Plain
/// text is typed as characters; braces name keys, optionally with modifiers:
/// <c>{ENTER}</c>, <c>{TAB}</c>, <c>{F5}</c>, <c>{CTRL+S}</c>, <c>{SHIFT+TAB}</c>, <c>{ALT+F}</c>,
/// <c>{CTRL+SHIFT+END}</c>, <c>{LBRACE}</c>/<c>{RBRACE}</c> for literal braces.
/// Events: <c>S&lt;mods&gt;</c> modifier state (1 Shift, 2 Ctrl, 4 Alt), <c>D&lt;vk&gt;</c>/<c>U&lt;vk&gt;</c>
/// key down/up, <c>C&lt;code&gt;</c> a Windows-1252 character, separated by ';'.
/// </summary>
public static class KeySyntax
{
    private static readonly Dictionary<string, int> s_keys = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ENTER"] = 0x0D, ["RETURN"] = 0x0D, ["TAB"] = 0x09, ["ESC"] = 0x1B, ["ESCAPE"] = 0x1B, ["BACKSPACE"] = 0x08, ["BS"] = 0x08,
        ["SPACE"] = 0x20, ["DELETE"] = 0x2E, ["DEL"] = 0x2E, ["INSERT"] = 0x2D, ["INS"] = 0x2D, ["HOME"] = 0x24, ["END"] = 0x23,
        ["PGUP"] = 0x21, ["PAGEUP"] = 0x21, ["PGDN"] = 0x22, ["PAGEDOWN"] = 0x22,
        ["UP"] = 0x26, ["DOWN"] = 0x28, ["LEFT"] = 0x25, ["RIGHT"] = 0x27,
    };

    /// <summary>Parses <paramref name="keys"/>; throws <see cref="FormatException"/> with a readable message.</summary>
    public static string ToEvents(string keys)
    {
        var ev = new List<string>();
        int i = 0;
        while (i < keys.Length)
        {
            char c = keys[i];
            if (c == '{')
            {
                int end = keys.IndexOf('}', i + 1);
                if (end < 0) throw new FormatException($"'{{' at position {i + 1} has no closing '}}' (use {{LBRACE}} for a literal brace)");
                AddChord(ev, keys[(i + 1)..end]);
                i = end + 1;
                continue;
            }
            if (c == '}') throw new FormatException($"'}}' at position {i + 1} has no opening '{{' (use {{RBRACE}} for a literal brace)");
            if (c == '\r') { i++; continue; }
            if (c == '\n') { ev.Add("D13"); ev.Add("U13"); i++; continue; }          // a line break means Enter
            AddChar(ev, c);
            i++;
        }
        if (ev.Count == 0) throw new FormatException("no keys given");
        return string.Join(";", ev);
    }

    private static void AddChar(List<string> ev, char c)
    {
        byte[] bytes;
        try { bytes = Vb1252.Strict.GetBytes(c.ToString()); }
        catch (EncoderFallbackException) { bytes = []; }
        if (bytes.Length != 1 || bytes[0] == 0) throw new FormatException($"character '{c}' can't be typed (not in Windows-1252)");
        ev.Add("C" + bytes[0].ToString(CultureInfo.InvariantCulture));
    }

    private static void AddChord(List<string> ev, string chord)
    {
        var parts = chord.Split('+', StringSplitOptions.TrimEntries);
        int mods = 0;
        foreach (var m in parts[..^1])
            mods |= m.ToUpperInvariant() switch
            {
                "SHIFT" => 1, "CTRL" or "CONTROL" => 2, "ALT" => 4,
                _ => throw new FormatException($"unknown modifier '{m}' in {{{chord}}} (SHIFT, CTRL, ALT)"),
            };
        var key = parts[^1];
        if (key.Equals("LBRACE", StringComparison.OrdinalIgnoreCase) && mods == 0) { AddChar(ev, '{'); return; }
        if (key.Equals("RBRACE", StringComparison.OrdinalIgnoreCase) && mods == 0) { AddChar(ev, '}'); return; }
        int vk = VirtualKey(key) ?? throw new FormatException($"unknown key '{key}' in {{{chord}}}");
        if (mods != 0) ev.Add("S" + mods.ToString(CultureInfo.InvariantCulture));
        ev.Add("D" + vk.ToString(CultureInfo.InvariantCulture));
        ev.Add("U" + vk.ToString(CultureInfo.InvariantCulture));
        if (mods != 0) ev.Add("S0");
    }

    private static int? VirtualKey(string key)
    {
        if (s_keys.TryGetValue(key, out var vk)) return vk;
        if (key.Length == 1 && char.IsAsciiLetterOrDigit(key[0])) return char.ToUpperInvariant(key[0]);   // VK_A.. / VK_0..
        if (key.Length >= 2 && (key[0] is 'F' or 'f') && int.TryParse(key[1..], NumberStyles.None, CultureInfo.InvariantCulture, out var f) && f is >= 1 and <= 12)
            return 0x70 + f - 1;
        return null;
    }
}
