using System.Text;

namespace Vb6Mcp.Core;

/// <summary>
/// Line diff (Myers' O((N+M)·D) algorithm) with unified-diff output, for sync_diff: the IDE's text of
/// a component against the file on disk. Written here because the project takes no third-party
/// packages (AGENTS.md principle 1); it only needs to be right and fast enough for source files.
/// </summary>
public static class LineDiff
{
    /// <summary>One edit: Equal, Delete (only in <c>a</c>) or Insert (only in <c>b</c>), with 0-based positions.</summary>
    public readonly record struct Edit(char Op, int A, int B, string Text);

    public static IReadOnlyList<string> Lines(string text) => text.Length == 0 ? [] : text.Replace("\r\n", "\n").Split('\n');

    /// <summary>The edit script turning <paramref name="a"/> into <paramref name="b"/>.</summary>
    public static IReadOnlyList<Edit> Diff(IReadOnlyList<string> a, IReadOnlyList<string> b)
    {
        int n = a.Count, m = b.Count, max = n + m;
        var v = new int[2 * max + 2];
        var trace = new List<int[]>();
        int found = -1;
        for (int d = 0; d <= max && found < 0; d++)
        {
            trace.Add((int[])v.Clone());
            for (int k = -d; k <= d; k += 2)
            {
                int x = k == -d || (k != d && v[max + k - 1] < v[max + k + 1]) ? v[max + k + 1] : v[max + k - 1] + 1;
                int y = x - k;
                while (x < n && y < m && a[x] == b[y]) { x++; y++; }
                v[max + k] = x;
                if (x >= n && y >= m) { found = d; break; }
            }
        }
        // Walk back through the saved states to recover the path
        var edits = new List<Edit>();
        int cx = n, cy = m;
        for (int d = found; d > 0; d--)
        {
            var pv = trace[d];
            int k = cx - cy;
            int prevK = k == -d || (k != d && pv[max + k - 1] < pv[max + k + 1]) ? k + 1 : k - 1;
            int px = pv[max + prevK], py = px - prevK;
            while (cx > px && cy > py) { cx--; cy--; edits.Add(new Edit('=', cx, cy, a[cx])); }
            if (cx == px) { cy--; edits.Add(new Edit('+', cx, cy, b[cy])); }
            else { cx--; edits.Add(new Edit('-', cx, cy, a[cx])); }
        }
        while (cx > 0 && cy > 0) { cx--; cy--; edits.Add(new Edit('=', cx, cy, a[cx])); }
        edits.Reverse();
        return edits;
    }

    /// <summary>Unified diff with <paramref name="context"/> lines around each change; empty when equal.</summary>
    public static string Unified(string a, string b, string aName, string bName, int context = 3)
    {
        var edits = Diff(Lines(a), Lines(b));
        if (edits.All(e => e.Op == '=')) return "";
        var sb = new StringBuilder().Append("--- ").AppendLine(aName).Append("+++ ").AppendLine(bName);
        int i = 0;
        while (i < edits.Count)
        {
            if (edits[i].Op == '=') { i++; continue; }
            int start = Math.Max(0, i - context), end = i;
            // extend the hunk while changes are within 2*context of each other
            while (end < edits.Count)
            {
                int nextChange = end;
                while (nextChange < edits.Count && edits[nextChange].Op == '=') nextChange++;
                if (nextChange >= edits.Count || nextChange - end > 2 * context) break;
                end = nextChange;
                while (end < edits.Count && edits[end].Op != '=') end++;
            }
            int stop = Math.Min(edits.Count, end + context);
            var hunk = edits.Skip(start).Take(stop - start).ToList();
            int aStart = hunk.First().A, bStart = hunk.First().B;
            int aLen = hunk.Count(e => e.Op != '+'), bLen = hunk.Count(e => e.Op != '-');
            sb.Append("@@ -").Append(aStart + 1).Append(',').Append(aLen).Append(" +").Append(bStart + 1).Append(',').Append(bLen).AppendLine(" @@");
            foreach (var e in hunk) sb.Append(e.Op == '=' ? ' ' : e.Op).AppendLine(e.Text);
            i = stop;
        }
        return sb.ToString();
    }
}
