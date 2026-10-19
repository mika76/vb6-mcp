#:project ../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
// Benchmark T14 (debug a reported bug): writes tests\fixtures\Cart, a small order form with a DELIBERATE bug.
// The bug isn't marked anywhere in the VB6 files (the agent reads them): frmCart keeps one module-level
// "Private m_line As New clsLine" and adds that same object for every line, so every line in the collection
// becomes the last one entered. The list box still looks right (its text is taken when the line is added); only
// the total is wrong, and only once two different lines have been added.
//   10 x 1, 20 x 2, 5 x 4  -> should be 68.00 (third line 10% off), shows 58.00
//   3 x 2, 7 x 1           -> should be 13.00, shows 14.00
// The discount rule in modPricing ("position > 2") is right and must stay. The fix is one or two lines in frmCart.
// Usage (repo root): dotnet run bench\make-cart-fixture.cs
using Vb6Mcp.Core;

if (!Directory.Exists(Path.Combine("tests", "fixtures"))) throw new Exception("run from the repo root");
var root = Path.GetFullPath(Path.Combine("tests", "fixtures", "Cart"));

void Put(string rel, params string[] lines)
{
    var path = Path.Combine(root, rel);
    Directory.CreateDirectory(Path.GetDirectoryName(path)!);
    var text = string.Join("\r\n", lines) + (rel.EndsWith(".vbp") ? "" : "\r\n");   // hard rule 1: 1252, CR LF
    File.WriteAllBytes(path, Vb1252.Encoding.GetBytes(text));
}

string[] Control(string type, string name, int tab, int left, int top, int width, int height, params string[] props) =>
[
    $"   Begin VB.{type} {name} ", .. props.Select(p => "      " + p),
    $"      Height          =   {height}", $"      Left            =   {left}", $"      TabIndex        =   {tab}",
    $"      Top             =   {top}", $"      Width           =   {width}", "   End",
];

Put("frmCart.frm",
[
    "VERSION 5.00",
    "Begin VB.Form frmCart ",
    "   Caption         =   \"Order\"",
    "   ClientHeight    =   3240",
    "   ClientLeft      =   60",
    "   ClientTop       =   405",
    "   ClientWidth     =   5880",
    "   LinkTopic       =   \"Form1\"",
    "   ScaleHeight     =   3240",
    "   ScaleWidth      =   5880",
    "   StartUpPosition =   3  'Windows Default",
    .. Control("TextBox", "txtTotal", 6, 240, 2640, 5415, 375, "Locked          =   -1  'True"),
    .. Control("ListBox", "lstLines", 5, 240, 960, 5415, 1425),
    .. Control("CommandButton", "cmdClear", 4, 4440, 450, 1215, 375, "Caption         =   \"Clear\""),
    .. Control("CommandButton", "cmdAdd", 3, 3120, 450, 1215, 375, "Caption         =   \"Add\""),
    .. Control("TextBox", "txtQty", 2, 1680, 480, 1215, 315),
    .. Control("TextBox", "txtPrice", 1, 240, 480, 1215, 315),
    .. Control("Label", "lblQty", 7, 1680, 240, 1215, 255, "Caption         =   \"Quantity\""),
    .. Control("Label", "lblPrice", 0, 240, 240, 1215, 255, "Caption         =   \"Price\""),
    "End",
    "Attribute VB_Name = \"frmCart\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
    "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False",
    "Option Explicit",
    "",
    "Private m_lines As Collection",
    "Private m_line As New clsLine",
    "",
    "Private Sub Form_Load()",
    "    Set m_lines = New Collection",
    "    ShowTotal",
    "End Sub",
    "",
    "Private Sub cmdAdd_Click()",
    "    If Val(txtQty.Text) <= 0 Then",
    "        MsgBox \"Enter a quantity\", vbExclamation",
    "        Exit Sub",
    "    End If",
    "    m_line.Price = CCur(Val(txtPrice.Text))",
    "    m_line.Quantity = CLng(Val(txtQty.Text))",
    "    m_lines.Add m_line",
    "    lstLines.AddItem m_line.Describe",
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
    "    txtTotal.Text = \"Total: \" & Format$(CartTotal(m_lines), \"0.00\")",
    "End Sub",
]);

Put("clsLine.cls",
    "VERSION 1.0 CLASS", "BEGIN", "  MultiUse = -1  'True", "  Persistable = 0  'NotPersistable", "  DataBindingBehavior = 0  'vbNone",
    "  DataSourceBehavior  = 0  'vbNone", "  MTSTransactionMode  = 0  'NotAnMTSObject", "END",
    "Attribute VB_Name = \"clsLine\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = True",
    "Attribute VB_PredeclaredId = False", "Attribute VB_Exposed = False",
    "' One order line",
    "Option Explicit",
    "",
    "Public Price As Currency",
    "Public Quantity As Long",
    "",
    "Public Function Amount() As Currency",
    "    Amount = Price * Quantity",
    "End Function",
    "",
    "Public Function Describe() As String",
    "    Describe = Quantity & \" x \" & Format$(Price, \"0.00\")",
    "End Function");

Put("modPricing.bas",
    "Attribute VB_Name = \"modPricing\"",
    "Option Explicit",
    "",
    "' The third and later lines of an order get 10% off.",
    "Public Function CartTotal(lines As Collection) As Currency",
    "    Dim i As Long, total As Currency",
    "    For i = 1 To lines.Count",
    "        total = total + LineTotal(lines(i), i)",
    "    Next",
    "    CartTotal = total",
    "End Function",
    "",
    "Private Function LineTotal(ByVal item As clsLine, ByVal position As Long) As Currency",
    "    LineTotal = item.Amount",
    "    If position > 2 Then LineTotal = LineTotal * 0.9",
    "End Function");

Put("Cart.vbp",
    "Type=Exe", "Form=frmCart.frm", "Module=modPricing; modPricing.bas", "Class=clsLine; clsLine.cls",
    "Startup=\"frmCart\"", "Command32=\"\"", "Name=\"Cart\"", "ExeName32=\"Cart.exe\"", "HelpContextID=\"0\"", "CompatibleMode=\"0\"",
    "MajorVer=1", "MinorVer=0", "RevisionVer=0", "AutoIncrementVer=0", "ServerSupportFiles=0", "CompilationType=0", "OptimizationType=0",
    "FavorPentiumPro(tm)=0", "CodeViewDebugInfo=0", "NoAliasing=0", "BoundsCheck=0", "OverflowCheck=0", "FlPointCheck=0", "FDIVCheck=0",
    "UnroundedFP=0", "StartMode=0", "Unattended=0", "Retained=0", "ThreadPerObject=0", "MaxNumberOfThreads=1");
Console.WriteLine("written to " + root);
