using System.Drawing;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// res_read (spike 35): a project's resource file (.res, ResFile32), read-only. VB code loads these with LoadResString,
/// LoadResPicture and LoadResData by id, so an agent changing UI text or icons needs to see which ids exist and what they
/// hold. Changing a .res stays with VB's Resource Editor (the hook refuses Edit/Write on .res).
/// </summary>
public sealed partial class WorkspaceTools
{
    private Tool ResReadTool() => new("res_read",
        "A VB6 resource file (.res; or pass the .vbp and its ResFile32 is used), read-only: every resource (type, name or id, language, size), " +
        "the string table (id and text, as LoadResString(id) returns it) and the version info. extract (\"TYPE/name\" keys such as BITMAP/101, " +
        "GROUP_ICON/MAINICON, CUSTOM/7, or [\"all\"]) takes resources out: bitmaps as .bmp, icon and cursor groups as .ico/.cur, others as raw " +
        "data, written to out_dir (a folder under the roots, nothing overwritten), or without out_dir the first picture shown as an image.",
        Schema.Object(("path", "string", "Full path of the .res file, or of a .vbp with a ResFile32 line.", true),
                      ("extract", "array", "Resources to take out, as TYPE/name keys from the listing, or [\"all\"].", false),
                      ("out_dir", "string", "Folder under the roots to write extracted resources to.", false)),
        (a, _) => ResReadAsync(a));

    private Task<ToolResult> ResReadAsync(JsonObject a)
    {
        var path = PathKey.Normalize(a.Req("path"));
        RequireUnderRoots(path);
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        if (path.EndsWith(".vbp", StringComparison.OrdinalIgnoreCase))
        {
            var vbp = VbpFile.Load(path);
            var res = vbp.Settings.GetValueOrDefault("ResFile32") is { Length: > 0 } r ? PathKey.Normalize(r, Path.GetDirectoryName(path)!) : null;
            path = res ?? throw new ToolException($"{Path.GetFileName(path)} has no resource file (no ResFile32 line)");
            RequireUnderRoots(path);
            if (!File.Exists(path)) throw new ToolException("the project's resource file isn't there: " + path);
        }
        var bytes = File.ReadAllBytes(path);
        IReadOnlyList<ResEntry> entries;
        try { entries = ResReader.Entries(bytes); }
        catch (Exception ex) when (ex is InvalidDataException or ArgumentException) { throw new ToolException($"{Path.GetFileName(path)}: {ex.Message}"); }

        var strings = ResReader.Strings(bytes, entries);
        object? version = null;
        if (entries.FirstOrDefault(e => e.TypeId == 16) is { } ve)
            try { var v = ResReader.Version(bytes, ve); version = new { fileVersion = v.FileVersion, productVersion = v.ProductVersion, strings = v.Strings }; }
            catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException) { version = new { error = "unreadable VERSION resource" }; }

        var wanted = a.List("extract");
        bool all = wanted.Any(w => w.Equals("all", StringComparison.OrdinalIgnoreCase));
        // individual ICON/CURSOR images belong to their group: "all" takes the groups, which make proper .ico/.cur files
        bool Extractable(ResEntry e) => e.TypeId is not (6 or 16 or 3 or 1);
        var unknown = wanted.Where(w => !all && !entries.Any(e => e.Key.Equals(w, StringComparison.OrdinalIgnoreCase))).ToList();
        if (unknown.Count > 0) throw new ToolException($"not in {Path.GetFileName(path)}: {string.Join(", ", unknown)}; it has {string.Join(", ", entries.Select(e => e.Key).Take(40))}");
        string? outDir = a.Str("out_dir") is { Length: > 0 } od ? PathKey.Normalize(od) : null;
        if (outDir is not null) { RequireUnderRoots(outDir); Directory.CreateDirectory(outDir); }

        var written = new List<string>();
        (byte[] Png, string Name)? shown = null;
        foreach (var e in entries.Where(e => all ? Extractable(e) : wanted.Any(w => e.Key.Equals(w, StringComparison.OrdinalIgnoreCase))))
        {
            var (data, ext) = e.TypeId switch
            {
                2 => (ResReader.Bitmap(bytes, e), "bmp"),
                14 => (ResReader.IconGroup(bytes, e, entries), "ico"),
                12 => (ResReader.IconGroup(bytes, e, entries), "cur"),
                _ => (bytes.AsSpan(e.Offset, e.Size).ToArray(), FrxReader.ImageType(bytes.AsSpan(e.Offset, e.Size).ToArray()) ?? "bin"),
            };
            if (outDir is not null)
            {
                var file = FreeName(Path.Combine(outDir, $"{Path.GetFileNameWithoutExtension(path)}_{e.Type}_{e.Name}.{ext}"));
                File.WriteAllBytes(file, data);
                written.Add(file);
            }
            else if (ext is "bmp" or "ico" or "cur" or "gif" or "jpg" or "png")
                shown ??= (ToPng(new FrxBlock("picture", data, ext, null, null)) ?? throw new ToolException($"{e.Key} can't be shown; pass out_dir to save it"), e.Key);
        }
        var result = new
        {
            source = "disk", path,
            resources = entries.Select(e => new { key = e.Key, type = e.Type, name = e.Name, language = $"0x{e.Language:X4}", bytes = e.Size }),
            strings = strings.Take(500).Select(s => new { id = s.Id, language = $"0x{s.Language:X4}", text = s.Text }),
            stringsTruncated = strings.Count > 500 ? true : (bool?)null,
            version, written = written.Count > 0 ? written : null, shown = shown?.Name,
        };
        if (written.Count > 0) Audit.Write("res_read", 0, $"{path} -> {string.Join(", ", written)}");
        return Task.FromResult(shown is { } s ? ToolResult.Image(s.Png, result) : ToolResult.Ok(result));
    }
}
