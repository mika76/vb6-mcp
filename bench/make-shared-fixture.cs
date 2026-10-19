#:project ../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
#:property BuiltInComInteropSupport=true
// Benchmark T13 (shared-source change across projects): writes tests\fixtures\Shared.
//   Common\modFormat.bas   FormatAmount, shared by three projects
//   SharedLib\             ActiveX DLL (binary compatibility) that includes it
//   AppA\                  EXE that includes it and references SharedLib (so SharedLib builds first)
//   AppB\                  EXE that includes it
//   AppC\                  EXE with its own private FormatAmount: same name, different file (hard rule 3), must not change
// Callers sit in a form, a class, a module, a continued line and a call split across a continuation, next to a comment
// and a string that name the function, so a line-by-line text replace gets it wrong.
// Two steps, because AppA's reference needs SharedLib's typelib (hard rule 13: the GUID comes from the built DLL):
//   dotnet run bench\make-shared-fixture.cs sources     then build SharedLib\SharedLib.vbp (build_make)
//   dotnet run bench\make-shared-fixture.cs link        compat file + AppA's reference; then build SharedLib again
// Files are written through Vb1252 with CR LF (hard rule 1).
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

var root = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "tests", "fixtures", "Shared"));
if (!Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "tests", "fixtures"))) throw new Exception("run from the repo root");

void Put(string rel, params string[] lines)
{
    var path = Path.Combine(root, rel);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    // VB writes no CR LF after a .vbp's last line; source files end with one
    var text = string.Join("\r\n", lines) + (rel.EndsWith(".vbp") ? "" : "\r\n");
    File.WriteAllBytes(path, Vb1252.Encoding.GetBytes(text));
}

string[] Settings(string type, string startup, string name, string? exeName, int startMode) =>
[
    $"Startup=\"{startup}\"", "Command32=\"\"", $"Name=\"{name}\"", .. exeName is null ? Array.Empty<string>() : [$"ExeName32=\"{exeName}\""],
    "HelpContextID=\"0\"", "CompatibleMode=\"0\"", "MajorVer=1", "MinorVer=0", "RevisionVer=0", "AutoIncrementVer=0", "ServerSupportFiles=0",
    "CompilationType=0", "OptimizationType=0", "FavorPentiumPro(tm)=0", "CodeViewDebugInfo=0", "NoAliasing=0", "BoundsCheck=0", "OverflowCheck=0",
    "FlPointCheck=0", "FDIVCheck=0", "UnroundedFP=0", $"StartMode={startMode}", "Unattended=0", "Retained=0", "ThreadPerObject=0", "MaxNumberOfThreads=1",
];

switch (args.FirstOrDefault())
{
    case "sources":
        Put(@"Common\modFormat.bas",
            "Attribute VB_Name = \"modFormat\"",
            "' Shared by SharedLib, AppA and AppB",
            "Option Explicit",
            "",
            "' Formats an amount with two decimals and thousands separators.",
            "Public Function FormatAmount(ByVal value As Currency) As String",
            "    FormatAmount = Format$(value, \"#,##0.00\")",
            "End Function",
            "",
            "' A comma-separated list of amounts.",
            "Public Function FormatAmountList(values() As Currency) As String",
            "    Dim i As Long, s As String",
            "    For i = LBound(values) To UBound(values)",
            "        If Len(s) > 0 Then s = s & \", \"",
            "        s = s & FormatAmount(values(i))",
            "    Next",
            "    FormatAmountList = s",
            "End Function");

        Put(@"SharedLib\clsPricing.cls",
            "VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "  Persistable = 0  'NotPersistable", "  DataBindingBehavior = 0  'vbNone",
            "  DataSourceBehavior  = 0  'vbNone", "  MTSTransactionMode  = 0  'NotAnMTSObject", "END",
            "Attribute VB_Name = \"clsPricing\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = True",
            "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = True",
            "Option Explicit",
            "",
            "' Net amount plus tax, as text.",
            "Public Function PriceText(ByVal net As Currency, ByVal taxRate As Double) As String",
            "    PriceText = FormatAmount(net) & \" + tax \" & _",
            "        FormatAmount(net * taxRate)",
            "End Function");
        Put(@"SharedLib\SharedLib.vbp",
            ["Type=OleDll", "Class=clsPricing; clsPricing.cls", @"Module=modFormat; ..\Common\modFormat.bas", .. Settings("OleDll", "(None)", "SharedLib", "SharedLib.dll", 1)]);

        Put(@"AppA\frmA.frm",
            "VERSION 5.00",
            "Begin VB.Form frmA ",
            "   Caption         =   \"AppA\"",
            "   ClientHeight    =   1800",
            "   ClientLeft      =   60",
            "   ClientTop       =   405",
            "   ClientWidth     =   4800",
            "   LinkTopic       =   \"Form1\"",
            "   ScaleHeight     =   1800",
            "   ScaleWidth      =   4800",
            "   StartUpPosition =   3  'Windows Default",
            "   Begin VB.Label lblPrice ",
            "      Caption         =   \"Price\"",
            "      Height          =   375",
            "      Left            =   240",
            "      TabIndex        =   1",
            "      Top             =   840",
            "      Width           =   4215",
            "   End",
            "   Begin VB.Label lblTotal ",
            "      Caption         =   \"Total\"",
            "      Height          =   375",
            "      Left            =   240",
            "      TabIndex        =   0",
            "      Top             =   240",
            "      Width           =   4215",
            "   End",
            "End",
            "Attribute VB_Name = \"frmA\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False",
            "Option Explicit",
            "",
            "Private Sub Form_Load()",
            "    Dim pricing As New SharedLib.clsPricing",
            "    ' FormatAmount always shows two decimals here",
            "    lblTotal.Caption = \"Total: \" & FormatAmount(TotalOf(3, 4.5))",
            "    lblPrice.Caption = pricing.PriceText(100, 0.2)",
            "    Me.Tag = \"FormatAmount(\"",
            "End Sub");
        Put(@"AppA\modA.bas",
            "Attribute VB_Name = \"modA\"",
            "Option Explicit",
            "",
            "Public Function TotalOf(ByVal count As Long, ByVal price As Currency) As Currency",
            "    TotalOf = count * price",
            "End Function",
            "",
            "Public Function Summary(ByVal a As Currency, ByVal b As Currency) As String",
            "    Summary = \"A=\" & FormatAmount(a) & _",
            "              \", B=\" & FormatAmount( _",
            "                  b)",
            "End Function");
        Put(@"AppA\AppA.vbp",
            ["Type=Exe", "Form=frmA.frm", "Module=modA; modA.bas", @"Module=modFormat; ..\Common\modFormat.bas", .. Settings("Exe", "frmA", "AppA", "AppA.exe", 0)]);

        Put(@"AppB\modB.bas",
            "Attribute VB_Name = \"modB\"",
            "Option Explicit",
            "",
            "Sub Main()",
            "    Dim v(1) As Currency",
            "    v(0) = 1.5: v(1) = 2.25",
            "    Debug.Print FormatAmountList(v)",
            "    Debug.Print \"Ten: \" & FormatAmount(10)",
            "End Sub");
        Put(@"AppB\AppB.vbp",
            ["Type=Exe", "Module=modB; modB.bas", @"Module=modFormat; ..\Common\modFormat.bas", .. Settings("Exe", "Sub Main", "AppB", "AppB.exe", 0)]);

        Put(@"AppC\modC.bas",
            "Attribute VB_Name = \"modC\"",
            "' AppC doesn't use Common\\modFormat.bas: this FormatAmount is its own",
            "Option Explicit",
            "",
            "Private Function FormatAmount(ByVal value As Currency) As String",
            "    FormatAmount = \"$\" & Format$(value, \"0\")",
            "End Function",
            "",
            "Sub Main()",
            "    Debug.Print FormatAmount(42)",
            "End Sub");
        Put(@"AppC\AppC.vbp",
            ["Type=Exe", "Module=modC; modC.bas", .. Settings("Exe", "Sub Main", "AppC", "AppC.exe", 0)]);
        Console.WriteLine("sources written to " + root + "; now build SharedLib\\SharedLib.vbp, then run with 'link'");
        break;

    case "link":
    {
        var dll = Path.Combine(root, "SharedLib", "SharedLib.dll");
        var lib = TypeLibInfo.Read(dll) ?? throw new Exception("build SharedLib first: no type library in " + dll);
        Directory.CreateDirectory(Path.Combine(root, "SharedLib", "compat"));
        File.Copy(dll, Path.Combine(root, "SharedLib", "compat", "SharedLib.dll"), true);
        var vbp = Path.Combine(root, "SharedLib", "SharedLib.vbp");
        var text = Vb1252.Encoding.GetString(File.ReadAllBytes(vbp))
            .Replace("CompatibleMode=\"0\"", "CompatibleMode=\"2\"\r\nCompatibleEXE32=\"compat\\SharedLib.dll\"");
        File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes(text));
        // the reference line as VB writes it: *\G{LIBID}#major.minor (hex)#lcid#hint path#description
        var appA = Path.Combine(root, "AppA", "AppA.vbp");
        var a = Vb1252.Encoding.GetString(File.ReadAllBytes(appA));
        if (!a.Contains("Reference="))
            a = a.Replace("Form=frmA.frm", $"Reference=*\\G{lib.LibId.ToString("B").ToUpperInvariant()}#{lib.Major:x}.{lib.Minor:x}#0#..\\SharedLib\\SharedLib.dll#SharedLib\r\nForm=frmA.frm");
        File.WriteAllBytes(appA, Vb1252.Encoding.GetBytes(a));
        Console.WriteLine($"SharedLib {lib.LibId:B} {lib.Major}.{lib.Minor}: compat file set, AppA references it; build SharedLib again, then AppA, AppB, AppC");
        break;
    }
    default:
        Console.WriteLine("usage: dotnet run bench\\make-shared-fixture.cs sources|link");
        break;
}
