#:project ../../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
#:property PublishAot=false
#:property BuiltInComInteropSupport=true
// Spike 36 (PLAN D3/D4): a copy of the Basic fixture whose form uses a control VB6 can't load, to watch "Errors during
// load" / "Continue Loading Project?" while the watchdog's recovery poll and add-in verbs run.
// The control's library is MSCOMCTL's own LIBID, read from the OCX's type library (hard rule 13: no hand-written GUIDs),
// with a major version that isn't registered, so VB resolves nothing (spike 25: GUID + major through HKCR\TypeLib).
// Files are written through Vb1252 (hard rule 1).
// Usage: dotnet run make.cs <target folder>
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;

var target = Path.GetFullPath(args[0]);
var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "fixtures", "Basic"));
if (!Directory.Exists(fixture)) fixture = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "tests", "fixtures", "Basic"));
var ocx = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "SysWOW64", "MSCOMCTL.OCX");
var lib = TypeLibInfo.Read(ocx) ?? throw new Exception("no type library in " + ocx);
int major = lib.Major + 7;   // not registered (checked below)
if (TypeLibRegistry.All().Any(t => t.LibId == lib.LibId && t.Major == major)) throw new Exception($"{major}.0 is registered after all");
var guid = lib.LibId.ToString("B").ToUpperInvariant();

Directory.CreateDirectory(target);
foreach (var f in new[] { "clsCounter.cls", "modMath.bas" }) File.Copy(Path.Combine(fixture, f), Path.Combine(target, f), true);
string Read(string f) => Vb1252.Encoding.GetString(File.ReadAllBytes(Path.Combine(fixture, f)));
void Write(string f, string text) => File.WriteAllBytes(Path.Combine(target, f), Vb1252.Encoding.GetBytes(text));

var vbp = Read("Basic.vbp").Replace("Name=\"Basic\"", "Name=\"LoadErr\"");
vbp = vbp.Replace("\r\nForm=", $"\r\nObject={guid}#{major}.0#0; MSCOMCTL.OCX\r\nForm=");
Write("LoadErr.vbp", vbp);

var frm = Read("frmMain.frm");
frm = frm.Replace("VERSION 5.00\r\n", $"VERSION 5.00\r\nObject = \"{guid}#{major}.0#0\"; \"MSCOMCTL.OCX\"\r\n");
int at = frm.IndexOf("\r\nEnd\r\nAttribute VB_Name", StringComparison.Ordinal);
frm = frm.Insert(at, "\r\n   Begin MSComctlLib.ProgressBar ProgressBar1\r\n      Height          =   255\r\n      Left            =   120\r\n      TabIndex        =   9\r\n      Top             =   120\r\n      Width           =   2000\r\n   End");
Write("frmMain.frm", frm);
// optional: more forms with the same control (does VB show one "Errors during load" per form?)
int extra = args.Length > 1 ? int.Parse(args[1]) : 0;
for (int i = 1; i <= extra; i++)
{
    var name = $"frmExtra{i}";
    Write(name + ".frm", frm.Replace("Begin VB.Form frmMain", $"Begin VB.Form {name}").Replace("Attribute VB_Name = \"frmMain\"", $"Attribute VB_Name = \"{name}\""));
    vbp = vbp.Replace("\r\nForm=frmMain.frm", $"\r\nForm=frmMain.frm\r\nForm={name}.frm");
}
if (extra > 0) Write("LoadErr.vbp", vbp);
Console.WriteLine($"{target}\\LoadErr.vbp: Object {guid}#{major}.0 (not registered), ProgressBar1 on frmMain and {extra} more form(s)");
