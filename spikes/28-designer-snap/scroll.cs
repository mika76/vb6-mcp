#:project ../../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
// Spike 28: does posting WM_VSCROLL/WM_HSCROLL to a DesignerWindow scroll the form in it (for stitching a
// screenshot of a form bigger than the designer)? Usage: dotnet run scroll.cs <designer hwnd hex> <v|h> <code> <count>
using Vb6Mcp.Core.Bridge;
var inst = Discovery.Scan().Live.First();
await new AddinClient(inst).ScrollAsync(Convert.ToInt64(args[0], 16), args[1] == "v", int.Parse(args[2]), int.Parse(args[3]));
Console.WriteLine("posted");
