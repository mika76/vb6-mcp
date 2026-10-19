#:project ../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
#:property BuiltInComInteropSupport=true
// Benchmark T15 (binary compatibility): writes tests\fixtures\Compat.
//   PriceLib\    ActiveX DLL, binary compatibility (compat\PriceLib.dll committed): clsTax.AddTax(net) adds 20%
//   ClientApp\   EXE that references PriceLib and shows AddTax(100) = 120 in a text box when it starts
// The harness builds both per run, then asks for a new method (fine) and a new parameter on AddTax (breaks binary
// compatibility, so the compiled ClientApp.exe would stop working): the agent must keep compatibility and surface
// the conflict instead of accepting the break (hard rule 13).
// Two steps, because ClientApp's reference needs PriceLib's typelib (hard rule 13: the GUID comes from the built DLL):
//   dotnet run bench\make-compat-fixture.cs sources     then build PriceLib\PriceLib.vbp (build_make)
//   dotnet run bench\make-compat-fixture.cs link        compat file + ClientApp's reference; build PriceLib again, then ClientApp
using Vb6Mcp.Core;

if (!Directory.Exists(Path.Combine("tests", "fixtures"))) throw new Exception("run from the repo root");
var root = Path.GetFullPath(Path.Combine("tests", "fixtures", "Compat"));

void Put(string rel, params string[] lines)
{
    var path = Path.Combine(root, rel);
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

switch (args.FirstOrDefault())
{
    case "sources":
        Put(@"PriceLib\clsTax.cls",
            "VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "  Persistable = 0  'NotPersistable", "  DataBindingBehavior = 0  'vbNone",
            "  DataSourceBehavior  = 0  'vbNone", "  MTSTransactionMode  = 0  'NotAnMTSObject", "END",
            "Attribute VB_Name = \"clsTax\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = True",
            "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = True",
            "Option Explicit",
            "",
            "' The net amount plus 20% tax.",
            "Public Function AddTax(ByVal net As Currency) As Currency",
            "    AddTax = net * 1.2",
            "End Function");
        Put(@"PriceLib\PriceLib.vbp", ["Type=OleDll", "Class=clsTax; clsTax.cls", .. Settings("(None)", "PriceLib", "PriceLib.dll", 1)]);

        Put(@"ClientApp\frmClient.frm",
            "VERSION 5.00",
            "Begin VB.Form frmClient ",
            "   Caption         =   \"Client\"",
            "   ClientHeight    =   1200",
            "   ClientLeft      =   60",
            "   ClientTop       =   405",
            "   ClientWidth     =   3600",
            "   LinkTopic       =   \"Form1\"",
            "   ScaleHeight     =   1200",
            "   ScaleWidth      =   3600",
            "   StartUpPosition =   3  'Windows Default",
            "   Begin VB.TextBox txtResult ",
            "      Height          =   375",
            "      Left            =   240",
            "      Locked          =   -1  'True",
            "      TabIndex        =   0",
            "      Top             =   360",
            "      Width           =   3015",
            "   End",
            "End",
            "Attribute VB_Name = \"frmClient\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
            "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False",
            "Option Explicit",
            "",
            "Private Sub Form_Load()",
            "    Dim tax As New PriceLib.clsTax",
            "    txtResult.Text = CStr(tax.AddTax(100))",
            "End Sub");
        Put(@"ClientApp\ClientApp.vbp", ["Type=Exe", "Form=frmClient.frm", .. Settings("frmClient", "ClientApp", "ClientApp.exe", 0)]);
        Console.WriteLine("sources written to " + root + "; now build PriceLib\\PriceLib.vbp, then run with 'link'");
        break;

    case "link":
    {
        var dll = Path.Combine(root, "PriceLib", "PriceLib.dll");
        var lib = Vb6Mcp.Core.Files.TypeLibInfo.Read(dll) ?? throw new Exception("build PriceLib first: no type library in " + dll);
        Directory.CreateDirectory(Path.Combine(root, "PriceLib", "compat"));
        File.Copy(dll, Path.Combine(root, "PriceLib", "compat", "PriceLib.dll"), true);
        var vbp = Path.Combine(root, "PriceLib", "PriceLib.vbp");
        var text = Vb1252.Encoding.GetString(File.ReadAllBytes(vbp));
        if (!text.Contains("CompatibleEXE32"))
            text = text.Replace("CompatibleMode=\"0\"", "CompatibleMode=\"2\"\r\nCompatibleEXE32=\"compat\\PriceLib.dll\"");
        File.WriteAllBytes(vbp, Vb1252.Encoding.GetBytes(text));
        // the reference line as VB writes it: *\G{LIBID}#major.minor (hex)#lcid#hint path#description
        var client = Path.Combine(root, "ClientApp", "ClientApp.vbp");
        var c = Vb1252.Encoding.GetString(File.ReadAllBytes(client));
        if (!c.Contains("Reference="))
            c = c.Replace("Form=frmClient.frm", $"Reference=*\\G{lib.LibId.ToString("B").ToUpperInvariant()}#{lib.Major:x}.{lib.Minor:x}#0#..\\PriceLib\\PriceLib.dll#PriceLib\r\nForm=frmClient.frm");
        File.WriteAllBytes(client, Vb1252.Encoding.GetBytes(c));
        Console.WriteLine($"PriceLib {lib.LibId:B} {lib.Major}.{lib.Minor}: compat file set, ClientApp references it; build PriceLib again, then ClientApp");
        break;
    }
    default:
        Console.WriteLine("usage: dotnet run bench\\make-compat-fixture.cs sources|link");
        break;
}
