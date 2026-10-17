#:project ../../src/Vb6Mcp.Core/Vb6Mcp.Core.csproj
// Spike 29: the raw CONTROLS records (progId, className, controlType) the add-in returns for each control kind.
using Vb6Mcp.Core.Bridge;
var inst = Discovery.Scan().Live.First();
var (_, payload) = await new AddinClient(inst).DesignAsync("CONTROLS", [args[0], args[1]]);
foreach (var r in Vb6Mcp.Core.Protocol.Wire.Records(payload).DistinctBy(r => r[2])) Console.WriteLine(string.Join(" | ", r.Take(6)));
