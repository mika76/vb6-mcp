#:project ../../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
// Phase 7 B8 spike: does VBComponent.SaveAs(own path) save one component quietly, without the .vbp?
// Usage: dotnet run savecomp.cs <pid> <project> <component>
using Vb6Mcp.Core.Bridge;
var inst = Discovery.Scan().Live.First(i => i.Pid == int.Parse(args[0]));
await new AddinClient(inst).SaveComponentAsync(args[1], args[2]);
Console.WriteLine("queued");
