using Vb6Mcp.Core.Bridge;

namespace Vb6Mcp.Tests;

public static class KeySyntaxTests
{
    [Test]
    public static void TextAndNamedKeys()
    {
        Assert.Equal("C72;C105", KeySyntax.ToEvents("Hi"));
        Assert.Equal("C97;D13;U13", KeySyntax.ToEvents("a{ENTER}"));
        Assert.Equal("D9;U9;D116;U116;D38;U38", KeySyntax.ToEvents("{tab}{F5}{Up}"));
        Assert.Equal("C123;C125", KeySyntax.ToEvents("{LBRACE}{RBRACE}"));
        Assert.Equal("C120;D13;U13;C121", KeySyntax.ToEvents("x\r\ny"), "a line break is Enter");
        Assert.Equal("C233", KeySyntax.ToEvents("é"), "1252 characters are typed as their 1252 code");
    }

    [Test]
    public static void RealInputTypesMappedCharactersAsKeys()
    {
        // US-like layout for the test: letters and digits have keys, 'é' has none (stays a character)
        (int, bool)? Map(char c) => char.IsAsciiLetterOrDigit(c) ? (char.ToUpperInvariant(c), char.IsUpper(c)) : null;
        Assert.Equal("D66;U66;S1;D66;U66;S0;D49;U49;C233;D13;U13", KeySyntax.ToEvents("bB1é{ENTER}", Map),
            "Form_KeyDown should see KeyCode 66 for b, not VK_PACKET (improvement notes 54)");
    }

    [Test]
    public static void ChordsSetAndClearModifiers()
    {
        Assert.Equal("S2;D83;U83;S0", KeySyntax.ToEvents("{CTRL+S}"));
        Assert.Equal("S1;D9;U9;S0", KeySyntax.ToEvents("{SHIFT+TAB}"));
        Assert.Equal("S4;D70;U70;S0", KeySyntax.ToEvents("{ALT+F}"));
        Assert.Equal("S3;D35;U35;S0", KeySyntax.ToEvents("{CTRL+SHIFT+END}"));
    }

    [Test]
    public static void BadInputIsExplained()
    {
        foreach (var (input, expect) in new[]
        {
            ("{ENTER", "no closing"), ("a}", "no opening"), ("{BOGUS}", "unknown key"), ("{WIN+R}", "unknown modifier"),
            ("€x中", "Windows-1252"), ("", "no keys"), ("{F13}", "unknown key"),
        })
        {
            try { KeySyntax.ToEvents(input); throw new InvalidOperationException($"'{input}' should fail"); }
            catch (FormatException ex) { Assert.True(ex.Message.Contains(expect), $"'{input}': {ex.Message}"); }
        }
    }
}
