#:project ../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
// Benchmark T17 (breakpoint debugging): writes tests\fixtures\Shop, a shop form whose price list is a binary data file.
//   products.dat   records as VB's Put writes a Product: Name As String * 12 (ANSI, space-padded), Price As Currency
//                  (Int64, value x 10000), Active As Boolean (Int16, -1/0); a Long record count first
// The DELIBERATE bug, marked only here: Tea is in the file twice, an old discontinued record (3.20, Active False) before
// the current one (2.80). FindPrice returns the first name match and ignores Active, so Tea costs 3.20. The code
// looks right; the duplicate is only in the binary data, which the running program shows (breakpoint in FindPrice,
// Locals / debug_eval of Products(i)). The fix: skip inactive records in FindPrice.
//   Tea 1 -> should be 2.80, shows 3.20; then Coffee 2 -> 7.80 (8.20); then Juice 1 -> 10.80 (11.20)
// Usage (repo root): dotnet run bench\make-shop-fixture.cs
using Vb6Mcp.Core;

if (!Directory.Exists(Path.Combine("tests", "fixtures"))) throw new Exception("run from the repo root");
var root = Path.GetFullPath(Path.Combine("tests", "fixtures", "Shop"));
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

// The data file, in Put's layout (no padding between fields in a Binary file)
var products = new (string Name, decimal Price, bool Active)[]
{
    ("Coffee", 2.50m, true), ("Tea", 3.20m, false), ("Water", 1.00m, true), ("Tea", 2.80m, true), ("Juice", 3.00m, true),
};
using (var w = new BinaryWriter(File.Create(Path.Combine(root, "products.dat"))))
{
    w.Write(products.Length);                                          // Long
    foreach (var p in products)
    {
        w.Write(Vb1252.Encoding.GetBytes(p.Name.PadRight(12)));       // String * 12
        w.Write((long)(p.Price * 10000));                              // Currency
        w.Write((short)(p.Active ? -1 : 0));                           // Boolean
    }
}

Put("modProducts.bas",
    "Attribute VB_Name = \"modProducts\"",
    "Option Explicit",
    "",
    "Public Type Product",
    "    Name As String * 12",
    "    Price As Currency",
    "    Active As Boolean",
    "End Type",
    "",
    "Public Products() As Product",
    "",
    "' Reads the price list (a Long count, then the records).",
    "Public Sub LoadProducts(ByVal path As String)",
    "    Dim f As Integer, n As Long, i As Long",
    "    f = FreeFile",
    "    Open path For Binary Access Read As #f",
    "    Get #f, , n",
    "    ReDim Products(1 To n)",
    "    For i = 1 To n",
    "        Get #f, , Products(i)",
    "    Next",
    "    Close #f",
    "End Sub",
    "",
    "' The price of a product by name.",
    "Public Function FindPrice(ByVal productName As String) As Currency",
    "    Dim i As Long",
    "    For i = 1 To UBound(Products)",
    "        If RTrim$(Products(i).Name) = productName Then",
    "            FindPrice = Products(i).Price",
    "            Exit Function",
    "        End If",
    "    Next",
    "End Function");

Put("frmShop.frm",
[
    "VERSION 5.00",
    "Begin VB.Form frmShop ",
    "   Caption         =   \"Shop\"",
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
    .. Control("TextBox", "txtQty", 2, 1800, 480, 1095, 315, "Text            =   \"1\""),
    .. Control("ComboBox", "cboProduct", 1, 240, 480, 1455, 315, "Style           =   2  'Dropdown List"),
    .. Control("Label", "lblQty", 7, 1800, 240, 1095, 255, "Caption         =   \"Quantity\""),
    .. Control("Label", "lblProduct", 0, 240, 240, 1455, 255, "Caption         =   \"Product\""),
    "End",
    "Attribute VB_Name = \"frmShop\"", "Attribute VB_GlobalNameSpace = False", "Attribute VB_Creatable = False",
    "Attribute VB_PredeclaredId = True", "Attribute VB_Exposed = False",
    "Option Explicit",
    "",
    "Private m_total As Currency",
    "",
    "Private Sub Form_Load()",
    "    Dim i As Long",
    "    LoadProducts App.Path & \"\\products.dat\"",
    "    For i = 1 To UBound(Products)",
    "        If Not Listed(RTrim$(Products(i).Name)) Then cboProduct.AddItem RTrim$(Products(i).Name)",
    "    Next",
    "    cboProduct.ListIndex = 0",
    "    ShowTotal",
    "End Sub",
    "",
    "' True if the combo already has this product",
    "Private Function Listed(ByVal productName As String) As Boolean",
    "    Dim i As Long",
    "    For i = 0 To cboProduct.ListCount - 1",
    "        If cboProduct.List(i) = productName Then Listed = True",
    "    Next",
    "End Function",
    "",
    "Private Sub cmdAdd_Click()",
    "    Dim amount As Currency",
    "    If Val(txtQty.Text) <= 0 Then",
    "        MsgBox \"Enter a quantity\", vbExclamation",
    "        Exit Sub",
    "    End If",
    "    amount = FindPrice(cboProduct.Text) * CLng(Val(txtQty.Text))",
    "    m_total = m_total + amount",
    "    lstLines.AddItem txtQty.Text & \" x \" & cboProduct.Text & \" = \" & Format$(amount, \"0.00\")",
    "    ShowTotal",
    "End Sub",
    "",
    "Private Sub cmdClear_Click()",
    "    m_total = 0",
    "    lstLines.Clear",
    "    ShowTotal",
    "End Sub",
    "",
    "Private Sub ShowTotal()",
    "    txtTotal.Text = \"Total: \" & Format$(m_total, \"0.00\")",
    "End Sub",
]);

Put("Shop.vbp",
    "Type=Exe", "Form=frmShop.frm", "Module=modProducts; modProducts.bas",
    "Startup=\"frmShop\"", "Command32=\"\"", "Name=\"Shop\"", "ExeName32=\"Shop.exe\"", "HelpContextID=\"0\"", "CompatibleMode=\"0\"",
    "MajorVer=1", "MinorVer=0", "RevisionVer=0", "AutoIncrementVer=0", "ServerSupportFiles=0", "CompilationType=0", "OptimizationType=0",
    "FavorPentiumPro(tm)=0", "CodeViewDebugInfo=0", "NoAliasing=0", "BoundsCheck=0", "OverflowCheck=0", "FlPointCheck=0", "FDIVCheck=0",
    "UnroundedFP=0", "StartMode=0", "Unattended=0", "Retained=0", "ThreadPerObject=0", "MaxNumberOfThreads=1");
Console.WriteLine("written to " + root);
