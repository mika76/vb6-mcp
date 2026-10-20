using Vb6Mcp.Core.Files;

namespace Vb6Mcp.Tests;

/// <summary>Names for a running form's windows from its .frm designer block (improvement notes 42).</summary>
public static class FrmDesignTests
{
    const string Frm =
        "VERSION 5.00\r\n" +
        "Object = \"{00000000-0000-0000-0000-000000000000}#1.0#0\"; \"XmText.ocx\"\r\n" +
        "Begin VB.Form frmDemo \r\n" +
        "   Caption         =   \"Demo\"\r\n" +
        "   ClientHeight    =   3000\r\n" +
        "   ClientWidth     =   4500\r\n" +
        "   BeginProperty Font \r\n" +
        "      Size            =   8.25\r\n" +
        "   EndProperty\r\n" +
        "   Begin XmTextCtrl.XmText txtCode \r\n" +
        "      Height          =   315\r\n" +
        "      Left            =   1500\r\n" +
        "      Top             =   300\r\n" +
        "      Width           =   1500\r\n" +
        "   End\r\n" +
        "   Begin VB.Frame fraStatus \r\n" +
        "      Height          =   900\r\n" +
        "      Left            =   150\r\n" +
        "      Top             =   900\r\n" +
        "      Width           =   3000\r\n" +
        "      Begin VB.OptionButton optStatus \r\n" +
        "         Height          =   300\r\n" +
        "         Index           =   1\r\n" +
        "         Left            =   1500\r\n" +
        "         Top             =   300\r\n" +
        "         Width           =   1200\r\n" +
        "      End\r\n" +
        "   End\r\n" +
        "   Begin VB.PictureBox picPixels \r\n" +
        "      Height          =   600\r\n" +
        "      Left            =   3300\r\n" +
        "      ScaleMode       =   3  'Pixel\r\n" +
        "      Top             =   900\r\n" +
        "      Width           =   900\r\n" +
        "      Begin VB.CommandButton cmdIn \r\n" +
        "         Height          =   20\r\n" +
        "         Left            =   5\r\n" +
        "         Top             =   5\r\n" +
        "         Width           =   40\r\n" +
        "      End\r\n" +
        "   End\r\n" +
        "End\r\n" +
        "Attribute VB_Name = \"frmDemo\"\r\n";

    [Test]
    public static void DesignerBlockIsReadAndWindowsAreNamed()
    {
        var form = FrmDesign.Parse(Frm)!;
        Assert.Equal("frmDemo:txtCode,fraStatus,picPixels", form.Name + ":" + string.Join(",", form.Children.Select(c => c.Name)));
        Assert.Equal("optStatus(1)", form.Children[1].Children[0].Name);
        Assert.True(form.Children[0].IsActiveX && !form.Children[1].IsActiveX, "UserControl vs intrinsic");

        // windows as app_tree sees them at 96 DPI (15 twips per pixel), positions relative to the parent window
        var windows = new List<FrmDesign.Window>
        {
            new(10, 1, 100, 20, 100, 21),    // txtCode
            new(11, 10, 0, 0, 100, 21),      // the UserControl's own inner text box: not in the .frm
            new(20, 1, 10, 60, 200, 60),     // fraStatus
            new(21, 20, 100, 20, 80, 20),    // optStatus(1), in the frame's (inherited) twips
            new(30, 1, 220, 60, 60, 40),     // picPixels
            new(31, 30, 6, 6, 40, 20),       // cmdIn, in pixels, 1 px off for the border
            new(40, 1, 300, 150, 50, 50),    // moved by code: no match
        };
        var map = FrmDesign.Match(form, 1, windows, 15);
        Assert.Equal("10=txtCode,20=fraStatus,21=optStatus(1),30=picPixels,31=cmdIn",
            string.Join(",", map.OrderBy(kv => kv.Key).Select(kv => $"{kv.Key}={kv.Value.Name}")));
    }
}
