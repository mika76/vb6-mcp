using System.Drawing;
using System.Text.Json.Nodes;
using Vb6Mcp.Core;
using Vb6Mcp.Core.Files;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// form_resources (spike 34): the pictures, lists and long strings a form keeps in its .frx (or .ctx/.pgx/.dsx/.dox), read
/// from disk. The owner's case: taking an image out of a form. The .frx holds the original image (GIF stays GIF), which
/// VBIDE's SavePicture wouldn't give, and it works for closed projects. Read-only: hard rule 2.
/// </summary>
public sealed partial class WorkspaceTools
{
    private Tool FormResourcesTool() => new("form_resources",
        "What a form, UserControl or property page keeps in its binary companion file (.frx/.ctx/.pgx): pictures, list items, long or " +
        "multi-line texts, fonts and ActiveX controls' saved state, " +
        "read from disk: control, property, kind, image type and size. extract (\"Control.Property\" names, or \"all\") takes them out: " +
        "pictures in their original format (GIF, JPEG, ICO, BMP...) and control state as .bin files in out_dir (a folder under the roots; existing files aren't " +
        "overwritten), or without out_dir the first picture shown as an image. Passwords in long texts (connection strings) are masked. " +
        "Read-only; change pictures and lists with form_set_property (@picture:, List).",
        Schema.Object(("path", "string", "Full path of the .frm/.ctl/.pag/.dsr/.dob file.", true),
                      ("extract", "array", "Properties to take out, as Control.Property (e.g. imgLogo.Picture, frmMain.Icon), or [\"all\"].", false),
                      ("out_dir", "string", "Folder under the roots to write extracted pictures to (default: show the first one).", false)),
        (a, _) => FormResourcesAsync(a));

    private async Task<ToolResult> FormResourcesAsync(JsonObject a)
    {
        var path = PathKey.Normalize(a.Req("path"));
        RequireUnderRoots(path);
        if (!File.Exists(path)) throw new ToolException("file not found: " + path);
        if (Path.GetExtension(path).ToLowerInvariant() is not (".frm" or ".ctl" or ".pag" or ".dsr" or ".dob"))
            throw new ToolException("path must be a designer file (.frm .ctl .pag .dsr .dob); its .frx is found from it");
        var entries = FrxReader.Entries(path);
        var form = Path.GetFileNameWithoutExtension(path);
        string NameOf(FrxEntry e) => $"{(e.Control.Length > 0 ? e.Control : form)}.{e.Property}";

        var wanted = a.List("extract");
        bool all = wanted.Any(w => w.Equals("all", StringComparison.OrdinalIgnoreCase));
        var unknown = wanted.Where(w => !all && !entries.Any(e => NameOf(e).Equals(w, StringComparison.OrdinalIgnoreCase))).ToList();
        if (unknown.Count > 0) throw new ToolException($"not in {Path.GetFileName(path)}: {string.Join(", ", unknown)}; it has {string.Join(", ", entries.Select(NameOf).Take(30))}");
        string? outDir = a.Str("out_dir") is { Length: > 0 } od ? PathKey.Normalize(od) : null;
        if (outDir is not null) { RequireUnderRoots(outDir); Directory.CreateDirectory(outDir); }

        var rows = new List<object>();
        var written = new List<string>();
        (byte[] Png, string Name)? shown = null;
        foreach (var e in entries)
        {
            FrxBlock b;
            try { b = FrxReader.Read(e); }
            catch (Exception ex) when (ex is InvalidDataException or IOException) { rows.Add(new { name = NameOf(e), error = ex.Message }); continue; }
            bool take = all ? b.Kind == "picture" : wanted.Any(w => w.Equals(NameOf(e), StringComparison.OrdinalIgnoreCase));
            string? file = null;
            string? note = null;
            if (take && b.Data.Length > 0)
            {
                if (outDir is not null)
                {
                    // pictures keep their own format; an ActiveX control's state or a font is saved as it is stored
                    file = FreeName(Path.Combine(outDir, $"{form}_{NameOf(e).Replace('.', '_').Replace('(', '_').Replace(")", "")}.{b.ImageType ?? "bin"}"));
                    File.WriteAllBytes(file, b.Data);
                    written.Add(file);
                }
                else if (b.Kind == "picture")
                    shown ??= (ToPng(b) ?? throw new ToolException($"{NameOf(e)} is a {b.ImageType ?? "unknown"} picture that can't be shown; pass out_dir to save it"), NameOf(e));
                else note = "pass out_dir to save it";
            }
            rows.Add(new
            {
                name = NameOf(e), kind = b.Kind, imageType = b.ImageType, bytes = b.Data.Length > 0 ? b.Data.Length : (int?)null,
                key = e.Key, items = b.Items, text = b.Text is null ? null : FrxReader.MaskSecrets(b.Text), file, note,
            });
        }
        var loaded = await FindInIdesAsync(path);
        var result = new
        {
            source = "disk", path, resources = rows, written = written.Count > 0 ? written : null, shown = shown?.Name,
            warning = loaded.Any(l => l.dirty) ? "the form has unsaved changes in a running IDE: the .frx on disk is the last saved state" : null,
        };
        Audit.Write("form_resources", 0, $"{path}{(written.Count > 0 ? " -> " + string.Join(", ", written) : "")}");
        return shown is { } s ? ToolResult.Image(s.Png, result) : ToolResult.Ok(result);
    }

    /// <summary>A path that doesn't exist yet: the name, or name-2, name-3...</summary>
    private static string FreeName(string path)
    {
        if (!File.Exists(path)) return path;
        for (int i = 2; ; i++)
        {
            var p = Path.Combine(Path.GetDirectoryName(path)!, $"{Path.GetFileNameWithoutExtension(path)}-{i}{Path.GetExtension(path)}");
            if (!File.Exists(p)) return p;
        }
    }

    /// <summary>The picture as PNG for showing, or null when System.Drawing can't read it.</summary>
    private static byte[]? ToPng(FrxBlock b)
    {
        try
        {
            using var ms = new MemoryStream(b.Data);
            using var bmp = b.ImageType is "ico" or "cur" ? new Icon(ms).ToBitmap() : new Bitmap(Image.FromStream(ms));
            return UI.Screenshots.ToPng(bmp, 1024);
        }
        catch (Exception ex) when (ex is ArgumentException or System.Runtime.InteropServices.ExternalException or OutOfMemoryException) { return null; }
    }
}
