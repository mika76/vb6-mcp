using Vb6Mcp.Core;

namespace Vb6Mcp.Tests;

public static class VbLocalsTests
{
    static string Names(IReadOnlyList<LocalDecl> d) => string.Join(",", d.Select(x => $"{x.Name}:{x.Kind}:{x.DeclaredType}{(x.IsArray ? "[]" : "")}"));

    [Test]
    public static void FunctionWithParamsLocalsAndArrays()
    {
        // tests/fixtures/Debug/modCalc.bas, Inner
        var text = "' doc comment\r\nPrivate Function Inner(ByVal n As Long, ByRef tag As String) As Long\r\n" +
                   "    Dim i As Long, acc As Long, items(2) As Variant, flag As Boolean, when As Date\r\n" +
                   "    items(0) = \"a\": items(1) = 2: items(2) = Null\r\n    Inner = acc\r\nEnd Function";
        Assert.Equal("n:Parameter:Long,tag:Parameter:String,Inner:ReturnValue:Long,i:Local:Long,acc:Local:Long,items:Local:Variant[],flag:Local:Boolean,when:Local:Date",
                     Names(VbLocals.Parse(text)));
    }

    [Test]
    public static void ContinuationsCommentsStaticAndReDim()
    {
        var text = "Public Sub Go(Optional ByVal x As Integer = 3, _\r\n    ParamArray rest() As Variant)\r\n" +
                   "    Dim s As String * 10 ' fixed: a, b\r\n    Static count&\r\n    Dim o As New Collection, f As Scripting.FileSystemObject\r\n" +
                   "    ReDim buf(1 To 10) As Byte\r\n    ReDim Preserve buf(1 To 20)\r\n    Rem Dim notThis As Long\r\n" +
                   "    Dim q As Long: q = 1\r\nEnd Sub\r\nDim afterEnd As Long";
        Assert.Equal("x:Parameter:Integer,rest:Parameter:Variant[],s:Local:String * 10,count:Static:,o:Local:Collection,f:Local:Scripting.FileSystemObject,buf:Local:Byte[],q:Local:Long",
                     Names(VbLocals.Parse(text)));
    }

    [Test]
    public static void PropertyGetHasAReturnValueLetDoesNot()
    {
        Assert.Equal("Value:ReturnValue:Long", Names(VbLocals.Parse("Public Property Get Value() As Long\r\n    Value = m_v\r\nEnd Property")));
        Assert.Equal("v:Parameter:Long", Names(VbLocals.Parse("Public Property Let Value(ByVal v As Long)\r\n    m_v = v\r\nEnd Property")));
    }

    [Test]
    public static void StringsWithColonsAndQuotesDontSplit()
    {
        var text = "Sub A()\r\n    MsgBox \"a: b ' not a comment\": Dim z As Long\r\n    Call F(x:=1): Dim w\r\nEnd Sub";
        Assert.Equal("z:Local:Long,w:Local:", Names(VbLocals.Parse(text)));
    }
}
