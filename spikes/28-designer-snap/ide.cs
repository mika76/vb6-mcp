#:project ../../src/Vb6Mcp.UI/Vb6Mcp.UI.csproj
// Reads an IDE's title, mode and dialogs without the add-in (channel C reads cross elevation, spike 10).
var s = Vb6Mcp.UI.IdeWindows.Read(int.Parse(args[0]));
Console.WriteLine($"title='{s.Title}' mode={s.Mode} responds={s.Responds} hung={s.IsHung}");
foreach (var d in s.Dialogs) Console.WriteLine($"dialog '{d.Caption}': {string.Join(" / ", d.Texts)} [{string.Join(", ", d.Buttons.Select(b => b.Caption))}]");
