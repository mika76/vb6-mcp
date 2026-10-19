#:project ../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
// Benchmark T18 (stepping): writes tests\fixtures\Order, an order form that recalculates the whole total after each Add.
// The DELIBERATE bug, marked only here: LineAmount(price As Currency, qty As Long) takes price ByRef (VB's default) and
// applies the bulk discount by assigning to it, so the discount is written back into Prices(). The first total is
// right; every later recalculation discounts the bulk line again. Seen by stepping into OrderTotal / LineAmount and
// watching Prices(1) change when the call returns. The fix: ByVal (or a local copy).
//   2 x 10 -> 18.00 (right); then 5 x 1 -> should be 23.00, shows 21.20; then 1 x 1 -> 24.00, shows 20.58
// Usage (repo root): dotnet run bench\make-order-fixture.cs
using Vb6Mcp.Core;

if (!Directory.Exists(Path.Combine("tests", "fixtures"))) throw new Exception("run from the repo root");
var root = Path.GetFullPath(Path.Combine("tests", "fixtures", "Order"));
Directory.CreateDirectory(root);

void Put(string rel, params string[] lines)
{
    var text = string.Join("\r\n", lines) + (rel.EndsWith(".vbp") ? "" : "\r\n");   // hard rule 1: 1252, CR LF
    File.WriteAllBytes(Path.Combine(root, rel), Vb1252.Encoding.GetBytes(text));
}

string[] Control(string type, string name, int tab, int left, int top, int width, int height, params string[] props) =>
[
    $"   Begin VB.{type} {name} ", .. props.Select(p => "      " + p),
    $"      Height          =   {height}", $"      Left            =   {left}", $"      TabIndex        =   {tab}",
    $"      Top             =   {top}", $"      Width           =   {width}", "   End",
];

Put("modOrder.bas",
    "Attribute VB_Name = \"modOrder\"",
    "Option Explicit",
    "",
    "Public Prices() As Currency",
    "Public Qtys() As Long",
    "Public LineCount As Long",
    "",
    "Public Sub AddLine(ByVal price As Currency, ByVal qty As Long)",
    "    LineCount = LineCount + 1",
    "    ReDim Preserve Prices(1 To LineCount)",
    "    ReDim Preserve Qtys(1 To LineCount)",
    "    Prices(LineCount) = price",
    "    Qtys(LineCount) = qty",
    "End Sub",
    "",
    "Public Sub ClearLines()",
    "    LineCount = 0",
    "    Erase Prices",
    "    Erase Qtys",
    "End Sub",
    "",
    "Public Function OrderTotal() As Currency",
    "    Dim i As Long, total As Currency",
    "    For i = 1 To LineCount",
    "        total = total + LineAmount(Prices(i), Qtys(i))",
    "    Next",
    "    OrderTotal = total",
    "End Function",
    "",
    "' 10 or more of an item: 10% off that line.",
    "Private Function LineAmount(price As Currency, qty As Long) As Currency",
    "    If qty >= 10 Then price = price * 0.9",
    "    LineAmount = price * qty",
    "End Function");

Put("frmOrder.frm",
[
    "VERSION 5.00",
    "Begin VB.Form frmOrder ",
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
    "Attribute VB_Name = \"frmOrder\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
    "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False",
    "Option Explicit",
    "",
    "Private Sub Form_Load()",
    "    ShowTotal",
    "End Sub",
    "",
    "Private Sub cmdAdd_Click()",
    "    If Val(txtQty.Text) <= 0 Then",
    "        MsgBox \"Enter a quantity\", vbExclamation",
    "        Exit Sub",
    "    End If",
    "    AddLine CCur(Val(txtPrice.Text)), CLng(Val(txtQty.Text))",
    "    lstLines.AddItem txtQty.Text & \" x \" & Format$(CCur(Val(txtPrice.Text)), \"0.00\")",
    "    ShowTotal",
    "End Sub",
    "",
    "Private Sub cmdClear_Click()",
    "    ClearLines",
    "    lstLines.Clear",
    "    ShowTotal",
    "End Sub",
    "",
    "Private Sub ShowTotal()",
    "    txtTotal.Text = \"Total: \" & Format$(OrderTotal(), \"0.00\")",
    "End Sub",
]);

Put("Order.vbp",
    "Type=Exe", "Form=frmOrder.frm", "Module=modOrder; modOrder.bas",
    "Startup=\"frmOrder\"", "Command32=\"\"", "Name=\"Order\"", "ExeName32=\"Order.exe\"", "HelpContextID=\"0\"", "CompatibleMode=\"0\"",
    "MajorVer=1", "MinorVer=0", "RevisionVer=0", "AutoIncrementVer=0", "ServerSupportFiles=0", "CompilationType=0", "OptimizationType=0",
    "FavorPentiumPro(tm)=0", "CodeViewDebugInfo=0", "NoAliasing=0", "BoundsCheck=0", "OverflowCheck=0", "FlPointCheck=0", "FDIVCheck=0",
    "UnroundedFP=0", "StartMode=0", "Unattended=0", "Retained=0", "ThreadPerObject=0", "MaxNumberOfThreads=1");
Console.WriteLine("written to " + root);
