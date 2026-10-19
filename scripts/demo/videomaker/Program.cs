// Renders the demo video from the timeline scripts/demo/video.js records: a chat panel (the prompt, each tool call
// and its result, the answer) next to a desktop with the IDE and the program's form, a drawn cursor that moves to
// the control each call acts on, and captions. Writes an MP4 (Media Foundation) and optionally a GIF.
// usage: videomaker <timeline dir> <out.mp4> [out.gif]
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using VideoMaker;

const int W = 1280, H = 720, Fps = 30, GifEvery = 3;   // the GIF gets every 3rd frame (10 fps)
const int ChatW = 420, Pad = 20;
var dir = args[0];
var events = JsonNode.Parse(File.ReadAllText(Path.Combine(dir, "timeline.json")))!.AsArray().Select(e => e!.AsObject()).ToList();

// ---------- steps: each event becomes a timed step ----------
var steps = new List<Step>();
steps.Add(new Step("title", 4.5, null));
foreach (var e in events)
{
    var t = (string)e["t"]!;
    switch (t)
    {
        case "prompt": steps.Add(new Step(t, 3.6, e)); steps.Add(new Step("wait", 0.8, null)); break;
        case "tool": steps.Add(new Step(t, 0.55, e)); break;
        case "shot": steps.Add(new Step(t, 0.35, e)); break;
        case "layout": steps.Add(new Step(t, 0, e)); break;
        case "cursor": steps.Add(new Step(t, 0.55, e)); break;
        case "click": steps.Add(new Step(t, 0.25, e)); break;
        case "caption": steps.Add(new Step(t, (double)e["hold"]!, e)); break;
        case "hide-app": steps.Add(new Step(t, 0.3, e)); break;
        case "answer": steps.Add(new Step(t, 3.0, e)); steps.Add(new Step("wait", 4.0, null)); break;
    }
}
double total = steps.Sum(s => s.Duration);
Console.WriteLine($"{steps.Count} steps, {total:F1} s");

// ---------- state ----------
var s = new State();
var images = new Dictionary<string, Bitmap>();
Bitmap Img(string f) => images.TryGetValue(f, out var b) ? b : images[f] = new Bitmap(Path.Combine(dir, f));

// desktop geometry: the IDE scaled to the desktop's width, the program's form at full size over its lower right
var desk = new Rectangle(ChatW, 0, W - ChatW, H);
// The IDE without its docked Project/Properties panes on the right (they don't change in this demo): the code
// window and the Immediate window come out about a third larger.
Rectangle IdeCrop(Bitmap ide) => new(0, 0, (int)(ide.Width * 0.74), ide.Height);
RectangleF IdeRect(Bitmap ide)
{
    var c = IdeCrop(ide);
    float sc = Math.Min((desk.Width - 2f * Pad) / c.Width, (H - 130f) / c.Height);
    return new RectangleF(desk.Left + (desk.Width - c.Width * sc) / 2, Pad, c.Width * sc, c.Height * sc);
}
RectangleF AppRect(Bitmap app) => new(W - Pad - app.Width, H - 82 - app.Height, app.Width, app.Height);
PointF? Target(string name)
{
    if (s.App is null || !s.Controls.TryGetValue(name, out var c)) return null;
    var r = AppRect(s.App);
    return new PointF(r.X + s.ClientOrigin.X + c.X + c.W / 2f, r.Y + s.ClientOrigin.Y + c.Y + c.H / 2f + 2);
}

using var fTitle = new Font("Segoe UI", 40f, FontStyle.Bold, GraphicsUnit.Pixel);
using var fSub = new Font("Segoe UI", 20f, GraphicsUnit.Pixel);
using var fBody = new Font("Segoe UI", 15f, GraphicsUnit.Pixel);
using var fSmall = new Font("Segoe UI", 13f, GraphicsUnit.Pixel);
using var fLabel = new Font("Segoe UI", 12f, FontStyle.Bold, GraphicsUnit.Pixel);
using var fMono = new Font("Consolas", 13f, GraphicsUnit.Pixel);
using var fCaption = new Font("Segoe UI", 18f, GraphicsUnit.Pixel);
var cBg = Color.FromArgb(30, 30, 30);
var cDesk = Color.FromArgb(32, 58, 90);
var cTool = Color.FromArgb(86, 182, 255);
var cResult = Color.FromArgb(150, 150, 150);
var cText = Color.FromArgb(225, 225, 225);

using var mp4 = new Mp4Writer(args[1], W, H, Fps);
using var gif = args.Length > 2 ? new GifWriter(args[2], W * 3 / 4, H * 3 / 4) : null;
using var frame = new Bitmap(W, H, PixelFormat.Format32bppRgb);
using var g = Graphics.FromImage(frame);
using var small = gif is null ? null : new Bitmap(W * 3 / 4, H * 3 / 4, PixelFormat.Format32bppRgb);
using var gs = small is null ? null : Graphics.FromImage(small);
g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
g.SmoothingMode = SmoothingMode.AntiAlias;

var pngAt = (Environment.GetEnvironmentVariable("VIDEOMAKER_PNG") ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries)
    .Select(x => double.Parse(x, System.Globalization.CultureInfo.InvariantCulture)).ToList();
int frames = (int)Math.Ceiling(total * Fps), stepIndex = -1;
double stepStart = 0;
for (int f = 0; f < frames; f++)
{
    double now = f / (double)Fps;
    while (stepIndex + 1 < steps.Count && (stepIndex < 0 || now >= stepStart + steps[stepIndex].Duration))
    {
        if (stepIndex >= 0) { stepStart += steps[stepIndex].Duration; End(steps[stepIndex]); }
        stepIndex++;
        Begin(steps[stepIndex]);
        if (steps[stepIndex].Duration > 0) break;
    }
    var st = steps[stepIndex];
    double p = st.Duration > 0 ? Math.Clamp((now - stepStart) / st.Duration, 0, 1) : 1;
    Draw(st, p);
    mp4.Write(frame);
    // VIDEOMAKER_PNG="12.5,40": also save the frames at those seconds, to check the layout without a player
    if (pngAt.Any(t => f == (int)(t * Fps))) frame.Save(Path.Combine(dir, $"check-{now:F1}.png"), ImageFormat.Png);
    if (gif is not null && f % GifEvery == 0)
    {
        gs!.InterpolationMode = InterpolationMode.HighQualityBicubic;
        gs.DrawImage(frame, 0, 0, small!.Width, small.Height);
        gif.Add(small, 100 * GifEvery / Fps);
    }
}
Console.WriteLine($"{frames} frames -> {args[1]}{(gif is null ? "" : " and " + args[2])}");

void Begin(Step st)
{
    var e = st.Event;
    switch (st.Kind)
    {
        case "title": s.Title = true; break;
        case "prompt": s.Title = false; s.Chat.Add(new Chat("you", (string)e!["text"]!, "")); break;
        case "tool": s.Caption = null; s.Chat.Add(new Chat("tool", (string)e!["call"]!, (string?)e["result"] ?? "")); break;
        case "shot":
            var target = (string)e!["target"]!;
            if (target == "ide") { s.Ide = Img((string)e["file"]!); s.AppFocus = false; }
            else
            {
                s.App = Img((string)e["file"]!); s.AppShown = true; s.AppFocus = true;
                if (e["clientOrigin"] is JsonObject co) s.ClientOrigin = new Point((int)co["x"]!, (int)co["y"]!);
            }
            break;
        case "layout":
            s.Controls.Clear();
            foreach (var (k, v) in e!["controls"]!.AsObject()) s.Controls[k] = new Ctl((int)v!["x"]!, (int)v["y"]!, (int)v["w"]!, (int)v["h"]!);
            break;
        case "cursor":
            if (Target((string)e!["target"]!) is { } to)
            {
                if (!s.CursorShown) { s.Cursor = new PointF(desk.Left + desk.Width / 2f, H / 2f); s.CursorShown = true; }
                s.From = s.Cursor; s.To = to;
            }
            else s.To = s.From = s.Cursor;
            break;
        case "click": s.ClickAt = s.Cursor; break;
        case "caption": s.Caption = (string)e!["text"]!; s.Tone = (string?)e["tone"] ?? "info"; break;
        case "hide-app": s.AppShown = false; s.AppFocus = false; s.CursorShown = false; break;
        case "answer": s.Caption = null; s.Chat.Add(new Chat("agent", (string)e!["text"]!, "")); break;
    }
}

void End(Step st)
{
    if (st.Kind == "cursor") s.Cursor = s.To;
    if (st.Kind == "click") s.ClickAt = null;
}

void Draw(Step st, double p)
{
    g.Clear(cBg);
    if (s.Title) { DrawTitle(p); return; }

    // desktop
    using (var b = new SolidBrush(cDesk)) g.FillRectangle(b, desk);
    if (s.Ide is { } ide) { var r = IdeRect(ide); Shadow(r); g.InterpolationMode = InterpolationMode.HighQualityBicubic; g.DrawImage(ide, r, IdeCrop(ide), GraphicsUnit.Pixel); }
    if (s.AppShown && s.AppFocus && s.App is { } app) { var r = AppRect(app); Shadow(r); g.InterpolationMode = InterpolationMode.NearestNeighbor; g.DrawImage(app, r); }

    // cursor and click
    if (st.Kind == "cursor") { double e = Ease(p); s.Cursor = new PointF((float)(s.From.X + (s.To.X - s.From.X) * e), (float)(s.From.Y + (s.To.Y - s.From.Y) * e)); }
    if (s.ClickAt is { } c && st.Kind == "click")
    {
        float rad = (float)(6 + 16 * p);
        using var pen = new Pen(Color.FromArgb((int)(220 * (1 - p)), 255, 210, 80), 3);
        g.DrawEllipse(pen, c.X - rad, c.Y - rad, 2 * rad, 2 * rad);
    }
    if (s.CursorShown && s.AppShown && s.AppFocus) DrawCursor(s.Cursor);

    // caption
    if (s.Caption is { } cap)
    {
        var color = s.Tone switch { "bug" => Color.FromArgb(255, 150, 120), "ok" => Color.FromArgb(130, 220, 130), _ => Color.White };
        var size = g.MeasureString(cap, fCaption, desk.Width - 60);
        var box = new RectangleF(desk.Left + (desk.Width - size.Width) / 2 - 14, H - 62, size.Width + 28, size.Height + 12);
        using (var b = new SolidBrush(Color.FromArgb(215, 15, 15, 15))) FillRounded(box, 8, b);
        using (var b = new SolidBrush(color)) g.DrawString(cap, fCaption, b, new RectangleF(box.X + 14, box.Y + 6, size.Width + 2, size.Height));
    }

    DrawChat(st, p);
}

void DrawChat(Step st, double p)
{
    using (var b = new SolidBrush(Color.FromArgb(24, 24, 24))) g.FillRectangle(b, 0, 0, ChatW, H);
    using (var b = new SolidBrush(Color.FromArgb(45, 45, 45))) g.FillRectangle(b, ChatW - 1, 0, 1, H);
    using (var b = new SolidBrush(cResult)) g.DrawString("Agent session (reconstructed)", fLabel, b, 16, 12);
    const float x = 16, w = ChatW - 32;
    // lay the items out, then scroll so the newest is visible
    var heights = new List<float>();
    foreach (var item in s.Chat)
    {
        float h = item.Kind switch
        {
            "tool" => g.MeasureString("vb6 > " + item.Text, fMono, (int)w).Height + (item.Result.Length > 0 ? g.MeasureString(item.Result, fSmall, (int)w - 12).Height : 0) + 8,
            _ => g.MeasureString(item.Text, fBody, (int)w - 20).Height + 40,
        };
        heights.Add(h);
    }
    float top = 40, sum = heights.Sum(), room = H - top - 12;
    float y = top - Math.Max(0, sum - room);
    var clip = g.Clip;
    g.SetClip(new RectangleF(0, top, ChatW - 1, H - top));
    for (int i = 0; i < s.Chat.Count; i++)
    {
        var item = s.Chat[i];
        bool last = i == s.Chat.Count - 1;
        if (y + heights[i] < top) { y += heights[i]; continue; }
        switch (item.Kind)
        {
            case "you":
            case "agent":
                {
                    var text = item.Text;
                    if (last && st.Kind is "prompt" or "answer") text = text[..(int)(text.Length * Math.Min(1, p * 1.15))];
                    var bubble = new RectangleF(x, y + 18, w, heights[i] - 26);
                    using (var b = new SolidBrush(item.Kind == "you" ? Color.FromArgb(44, 62, 92) : Color.FromArgb(40, 40, 40))) FillRounded(bubble, 8, b);
                    using (var b = new SolidBrush(cResult)) g.DrawString(item.Kind == "you" ? "You" : "Agent", fLabel, b, x + 2, y + 2);
                    using (var b = new SolidBrush(cText)) g.DrawString(text, fBody, b, new RectangleF(x + 10, y + 24, w - 20, heights[i]));
                    break;
                }
            default:
                {
                    var callH = g.MeasureString("vb6 > " + item.Text, fMono, (int)w).Height;
                    int alpha = last && st.Kind == "tool" ? (int)(255 * Math.Min(1, p * 3)) : 255;
                    using (var b = new SolidBrush(Color.FromArgb(alpha, cTool))) g.DrawString("vb6 > " + item.Text, fMono, b, new RectangleF(x, y, w, callH + 2));
                    if (item.Result.Length > 0 && !(last && st.Kind == "tool" && p < 0.5))
                        using (var b = new SolidBrush(cResult)) g.DrawString(item.Result, fSmall, b, new RectangleF(x + 12, y + callH, w - 12, heights[i]));
                    break;
                }
        }
        y += heights[i];
    }
    g.Clip = clip;
}

void DrawTitle(double p)
{
    g.Clear(cBg);
    using var white = new SolidBrush(Color.White);
    using var grey = new SolidBrush(Color.FromArgb(170, 170, 170));
    using var dim = new SolidBrush(Color.FromArgb(130, 130, 130));
    g.DrawString("vb6-mcp", fTitle, white, 110, 210);
    g.DrawString("An AI agent finds and fixes a bug in a VB6 program, working inside the IDE.", fSub, grey, 114, 272);
    g.DrawString("The screenshots are real: a scripted session through the vb6 tools, one tool call per step.\n" +
                 "The chat panel and the mouse cursor are drawn: the tools send input to the controls directly.", fBody, dim, new RectangleF(114, 330, 1000, 60));
}

void DrawCursor(PointF at)
{
    var pts = new[] { new PointF(0, 0), new PointF(0, 17), new PointF(4.5f, 13), new PointF(7.5f, 20), new PointF(10, 19), new PointF(7, 12), new PointF(12.5f, 12) }
        .Select(q => new PointF(at.X + q.X, at.Y + q.Y)).ToArray();
    g.FillPolygon(Brushes.White, pts);
    using var pen = new Pen(Color.Black, 1.2f);
    g.DrawPolygon(pen, pts);
}

void Shadow(RectangleF r)
{
    using var b = new SolidBrush(Color.FromArgb(70, 0, 0, 0));
    g.FillRectangle(b, r.X + 5, r.Y + 6, r.Width, r.Height);
}

void FillRounded(RectangleF r, float rad, Brush b)
{
    using var path = new GraphicsPath();
    path.AddArc(r.X, r.Y, 2 * rad, 2 * rad, 180, 90);
    path.AddArc(r.Right - 2 * rad, r.Y, 2 * rad, 2 * rad, 270, 90);
    path.AddArc(r.Right - 2 * rad, r.Bottom - 2 * rad, 2 * rad, 2 * rad, 0, 90);
    path.AddArc(r.X, r.Bottom - 2 * rad, 2 * rad, 2 * rad, 90, 90);
    path.CloseFigure();
    g.FillPath(b, path);
}

static double Ease(double t) => t < 0.5 ? 2 * t * t : 1 - Math.Pow(-2 * t + 2, 2) / 2;

record Step(string Kind, double Duration, JsonObject? Event);
record Chat(string Kind, string Text, string Result);
record Ctl(int X, int Y, int W, int H);

sealed class State
{
    public bool Title;
    public Bitmap? Ide, App;
    public bool AppShown, AppFocus;
    public Point ClientOrigin = new(4, 27);
    public Dictionary<string, Ctl> Controls = [];
    public List<Chat> Chat = [];
    public string? Caption, Tone;
    public PointF Cursor, From, To;
    public bool CursorShown;
    public PointF? ClickAt;
}
