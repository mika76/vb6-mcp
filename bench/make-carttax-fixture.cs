#:project ../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
#:property BuiltInComInteropSupport=true
// Benchmark T16 (a harder T14: a bug the code doesn't show): writes tests\fixtures\RateLibSrc and tests\fixtures\CartTax.
//   RateLibSrc\   source of RateLib.dll: clsRates.RateFor(code) by category code (FOOD 5%, BOOKS 0%, anything else 20%)
//                 and Codes() = "FOOD,STD,BOOKS". Never copied into a run: the agent gets the compiled DLL only.
//   CartTax\      an order form with a category combo; lib\RateLib.dll (committed) is the compiled DLL it references.
// The DELIBERATE bug, marked only here: the form passes UCase$(Left$(category, 4)) as the code, so "Books" becomes
// "BOOK", which RateLib doesn't know, and books get the standard 20%. Reading the form doesn't show it is wrong; the
// codes are only in the DLL, and only at run time: it keeps them shifted by one letter (CPPLT for BOOKS), because a
// strings dump of the DLL showed plain literals (dry run 3). Found by calling Codes() or RateFor in the running program.
//   Books 10 x 1 -> should be 10.00, shows 12.00; then Food 10 x 2 -> 31.00 (33.00); then Standard 10 x 1 -> 43.00 (45.00)
// Steps (repo root):
//   dotnet run bench\make-carttax-fixture.cs sources     then build tests\fixtures\RateLibSrc\RateLib.vbp (build_make)
//   dotnet run bench\make-carttax-fixture.cs link        copies the DLL to CartTax\lib, adds the reference; build CartTax\Cart.vbp
using Vb6Mcp.Core;

if (!Directory.Exists(Path.Combine("tests", "fixtures"))) throw new Exception("run from the repo root");
var fixtures = Path.GetFullPath(Path.Combine("tests", "fixtures"));

void Put(string rel, params string[] lines)
{
    var path = Path.Combine(fixtures, rel);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    var text = string.Join("\r\n", lines) + (rel.EndsWith(".vbp") ? "" : "\r\n");   // hard rule 1: 1252, CR LF
    File.WriteAllBytes(path, Vb1252.Encoding.GetBytes(text));
}

string[] Settings(string startup, string name, string exeName, int startMode) =>
[
    $"Startup=\"{startup}\"", "Command32=\"\"", $"Name=\"{name}\"", $"ExeName32=\"{exeName}\"", "HelpContextID=\"0\"", "CompatibleMode=\"0\"",
    "MajorVer=1", "MinorVer=0", "RevisionVer=0", "AutoIncrementVer=0", "ServerSupportFiles=0", "CompilationType=0", "OptimizationType=0",
    "FavorPentiumPro(tm)=0", "CodeViewDebugInfo=0", "NoAliasing=0", "BoundsCheck=0", "OverflowCheck=0", "FlPointCheck=0", "FDIVCheck=0",
    "UnroundedFP=0", $"StartMode={startMode}", "Unattended=0", "Retained=0", "ThreadPerObject=0", "MaxNumberOfThreads=1",
];

string[] Control(string type, string name, int tab, int left, int top, int width, int height, params string[] props) =>
[
    $"   Begin VB.{type} {name} ", .. props.Select(p => "      " + p),
    $"      Height          =   {height}", $"      Left            =   {left}", $"      TabIndex        =   {tab}",
    $"      Top             =   {top}", $"      Width           =   {width}", "   End",
];

switch (args.FirstOrDefault())
{
    case "sources":
        Put(@"RateLibSrc\clsRates.cls",
            "VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "  Persistable = 0  'NotPersistable", "  DataBindingBehavior = 0  'vbNone",
            "  DataSourceBehavior  = 0  'vbNone", "  MTSTransactionMode  = 0  'NotAnMTSObject", "END",
            "Attribute VB_Name = \"clsRates\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = True",
            "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = True",
            "Option Explicit",
            "",
            "' The category codes RateFor knows, comma-separated.",
            "Public Function Codes() As String",
            "    Codes = Plain(\"GPPE\") & \",\" & Plain(\"TUE\") & \",\" & Plain(\"CPPLT\")",
            "End Function",
            "",
            "' The tax rate for a category code; unknown codes get the standard rate.",
            "Public Function RateFor(ByVal code As String) As Double",
            "    Select Case code",
            "    Case Plain(\"GPPE\"): RateFor = 0.05",
            "    Case Plain(\"CPPLT\"): RateFor = 0",
            "    Case Else: RateFor = 0.2",
            "    End Select",
            "End Function",
            "",
            "' Codes are kept shifted by one letter",
            "Private Function Plain(ByVal s As String) As String",
            "    Dim i As Long",
            "    For i = 1 To Len(s)",
            "        Plain = Plain & Chr$(Asc(Mid$(s, i, 1)) - 1)",
            "    Next",
            "End Function");
        Put(@"RateLibSrc\RateLib.vbp", ["Type=OleDll", "Class=clsRates; clsRates.cls", .. Settings("(None)", "RateLib", "RateLib.dll", 1)]);

        Put(@"CartTax\frmCart.frm",
        [
            "VERSION 5.00",
            "Begin VB.Form frmCart ",
            "   Caption         =   \"Order\"",
            "   ClientHeight    =   3240",
            "   ClientLeft      =   60",
            "   ClientTop       =   405",
            "   ClientWidth     =   7200",
            "   LinkTopic       =   \"Form1\"",
            "   ScaleHeight     =   3240",
            "   ScaleWidth      =   7200",
            "   StartUpPosition =   3  'Windows Default",
            .. Control("TextBox", "txtTotal", 8, 240, 2640, 6735, 375, "Locked          =   -1  'True"),
            .. Control("ListBox", "lstLines", 7, 240, 960, 6735, 1425),
            .. Control("CommandButton", "cmdClear", 6, 5760, 450, 1215, 375, "Caption         =   \"Clear\""),
            .. Control("CommandButton", "cmdAdd", 5, 4440, 450, 1215, 375, "Caption         =   \"Add\""),
            .. Control("ComboBox", "cboCategory", 4, 2880, 480, 1455, 315, "Style           =   2  'Dropdown List"),
            .. Control("TextBox", "txtQty", 2, 1560, 480, 1215, 315),
            .. Control("TextBox", "txtPrice", 1, 240, 480, 1215, 315),
            .. Control("Label", "lblCategory", 3, 2880, 240, 1455, 255, "Caption         =   \"Category\""),
            .. Control("Label", "lblQty", 9, 1560, 240, 1215, 255, "Caption         =   \"Quantity\""),
            .. Control("Label", "lblPrice", 0, 240, 240, 1215, 255, "Caption         =   \"Price\""),
            "End",
            "Attribute VB_Name = \"frmCart\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False",
            "Option Explicit",
            "",
            "' Gross amounts of the lines (net plus tax)",
            "Private m_lines As Collection",
            "Private m_rates As New RateLib.clsRates",
            "",
            "Private Sub Form_Load()",
            "    Set m_lines = New Collection",
            "    cboCategory.AddItem \"Food\"",
            "    cboCategory.AddItem \"Standard\"",
            "    cboCategory.AddItem \"Books\"",
            "    cboCategory.ListIndex = 1",
            "    ShowTotal",
            "End Sub",
            "",
            "' The category's code for RateLib",
            "Private Function CategoryCode() As String",
            "    CategoryCode = UCase$(Left$(cboCategory.Text, 4))",
            "End Function",
            "",
            "Private Sub cmdAdd_Click()",
            "    Dim net As Currency",
            "    If Val(txtQty.Text) <= 0 Then",
            "        MsgBox \"Enter a quantity\", vbExclamation",
            "        Exit Sub",
            "    End If",
            "    net = CCur(Val(txtPrice.Text)) * CLng(Val(txtQty.Text))",
            "    m_lines.Add net * (1 + m_rates.RateFor(CategoryCode()))",
            "    lstLines.AddItem txtQty.Text & \" x \" & txtPrice.Text & \" (\" & cboCategory.Text & \")\"",
            "    ShowTotal",
            "End Sub",
            "",
            "Private Sub cmdClear_Click()",
            "    Set m_lines = New Collection",
            "    lstLines.Clear",
            "    ShowTotal",
            "End Sub",
            "",
            "Private Sub ShowTotal()",
            "    Dim i As Long, total As Currency",
            "    For i = 1 To m_lines.Count",
            "        total = total + m_lines(i)",
            "    Next",
            "    txtTotal.Text = \"Total: \" & Format$(total, \"0.00\")",
            "End Sub",
        ]);
        Put(@"CartTax\Cart.vbp", ["Type=Exe", "Form=frmCart.frm", .. Settings("frmCart", "Cart", "Cart.exe", 0)]);
        Console.WriteLine("sources written; now build tests\\fixtures\\RateLibSrc\\RateLib.vbp, then run with 'link'");
        break;

    case "link":
    {
        var built = Path.Combine(fixtures, "RateLibSrc", "RateLib.dll");
        var lib = Vb6Mcp.Core.Files.TypeLibInfo.Read(built) ?? throw new Exception("build RateLib first: no type library in " + built);
        Directory.CreateDirectory(Path.Combine(fixtures, "CartTax", "lib"));
        File.Copy(built, Path.Combine(fixtures, "CartTax", "lib", "RateLib.dll"), true);
        // the reference line as VB writes it (hard rule 13: the GUID comes from the DLL's typelib)
        var vbp = Path.Combine(fixtures, "CartTax", "Cart.vbp");
        var c = Vb1252.Encoding.GetString(File.ReadAllBytes(vbp));
        // RateLib has no binary compatibility: each build is a new LIBID, so a rebuilt DLL replaces the old reference
        c = string.Join("\r\n", c.Split("\r\n").Where(l => !l.StartsWith("Reference=")));
        c = c.Replace("Form=frmCart.frm", $"Reference=*\\G{lib.LibId.ToString("B").ToUpperInvariant()}#{lib.Major:x}.{lib.Minor:x}#0#lib\\RateLib.dll#RateLib\r\nForm=frmCart.frm");
        File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes(c));
        Console.WriteLine($"RateLib {lib.LibId:B} {lib.Major}.{lib.Minor} copied to CartTax\\lib and referenced; register that copy (com_register), then build CartTax\\Cart.vbp");
        break;
    }
    default:
        Console.WriteLine("usage: dotnet run bench\\make-carttax-fixture.cs sources|link");
        break;
}
