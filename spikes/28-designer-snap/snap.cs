#:project ../../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
// Spike 28: does the add-in's PrintWindow of a designer's ThunderForm return the whole form, also the part
// scrolled or clipped out of the designer window? Usage: dotnet run snap.cs <hwnd hex>
using Vb6Mcp.Core.Bridge;
var inst = Discovery.Scan().Live.First();
var path = await new AddinClient(inst).SnapAsync(Convert.ToInt64(args[0], 16), "spike" + DateTime.Now.Ticks);
for (int i = 0; i < 50 && !File.Exists(path); i++) await Task.Delay(100);
Console.WriteLine(path + " exists=" + File.Exists(path));
