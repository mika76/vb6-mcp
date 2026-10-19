using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Word-start matching for com_search (improvement notes 44: "Ml" matched every MSXML library).</summary>
public static class WordMatchTests
{
    [Test]
    public static void TheTextMustStartAWord()
    {
        Assert.Equal(WordMatch.Rank.Word, WordMatch.Of("Ml", "MlText.ocx"));               // camel case splits MlText
        Assert.Equal(WordMatch.Rank.Word, WordMatch.Of("ml", "Ml Controls 1.0"));
        Assert.Equal(WordMatch.Rank.WordPrefix, WordMatch.Of("Script", "Microsoft Scripting Runtime"));
        Assert.Equal(WordMatch.Rank.Substring, WordMatch.Of("Ml", "Microsoft XML, v6.0 (MSXML6)"));
        Assert.Equal(WordMatch.Rank.Substring, WordMatch.Of("Ml", "MSXML2.DOMDocument"));
        Assert.Equal(WordMatch.Rank.None, WordMatch.Of("Ml", "Microsoft ActiveX Data Objects"));
        Assert.Equal(WordMatch.Rank.WordPrefix, WordMatch.Of("Common Cont", "Microsoft Windows Common Controls 6.0 (SP6)"));   // spans words
        Assert.Equal(WordMatch.Rank.WordPrefix, WordMatch.Of("ADODB.Rec", "ADODB.Recordset.6.0"));
        Assert.Equal(WordMatch.Rank.Word, WordMatch.Of("Ml", "Other", "MlGrid.ocx"));   // best of several names
        Assert.Equal(WordMatch.Rank.WordPrefix, WordMatch.Of("Ml", "MSXML", "Mls Library"));
        Assert.Equal(WordMatch.Rank.None, WordMatch.Of("Ml", null, ""));
    }
}
