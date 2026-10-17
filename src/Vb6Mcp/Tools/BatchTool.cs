using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using Vb6Mcp.Mcp;

namespace Vb6Mcp.Tools;

/// <summary>
/// batch_run: several tool calls in one MCP call, run in order by the server through the same tool code as single
/// calls (ownership guard, design-mode checks, roots, audit and expected_hash all apply per call). Each MCP call is a
/// model turn, and the turns, not the tools, cost the tokens: the free-hand Opus runs (benchmark T12) made 285-326 calls,
/// many of them small chains like app_click, app_wait, app_screenshot. No new add-in verb: this only composes tools.
/// </summary>
public static class BatchTool
{
    public const int MaxCalls = 50;

    /// <summary>Characters kept of a successful call's text with compact (enough for "ok", a hash or a short status).</summary>
    public const int CompactChars = 160;

    /// <summary>Tools a batch won't run: itself, and long or disruptive ones that should stay single, visible calls.</summary>
    public static readonly HashSet<string> Excluded = new(StringComparer.Ordinal)
    {
        "batch_run", "ide_launch", "ide_kill", "build_make", "build_in_ide", "server_restart",
    };

    public static Tool Create(IReadOnlyList<Tool> tools)
    {
        var byName = tools.ToDictionary(t => t.Name, StringComparer.Ordinal);
        var schema = Schema.Object(("calls", "array", $"The calls, in order: [{{\"tool\": \"app_click\", \"args\": {{...}}}}, ...]. At most {MaxCalls}.", true),
                                   ("continue_on_error", "boolean", "Run the remaining calls after one fails (default: stop at the first error).", false),
                                   ("compact", "boolean", $"Show only the first {CompactChars} characters of each successful call's text except the last call's (errors and images in full). Saves tokens when the steps' results don't matter.", false));
        schema["properties"]!["calls"]!["items"] = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = new JsonObject { ["tool"] = new JsonObject { ["type"] = "string" }, ["args"] = new JsonObject { ["type"] = "object" } },
            ["required"] = new JsonArray("tool"),
        };
        return new Tool("batch_run",
            "Run several vb6 tool calls in one call, in order, e.g. app_click then app_wait then app_screenshot, or run_start, " +
            "clicks and run_end. Each call behaves exactly as on its own (same checks). Stops at the first error unless " +
            "continue_on_error. A later call can use a value from an earlier call's result: an argument that is exactly " +
            "\"$2.hash\" becomes the hash field of call 2's result (paths like $1.positions[0].left or $3.added.name work too), " +
            $"e.g. code_read then code_replace_lines with expected_hash \"$1.hash\". Not allowed in a batch: {string.Join(", ", Excluded.Where(e => e != "batch_run"))}.",
            schema,
            (a, ct) => RunAsync(byName, a, ct));
    }

    private static async Task<ToolResult> RunAsync(Dictionary<string, Tool> tools, JsonObject a, CancellationToken ct)
    {
        var list = a["calls"] switch
        {
            JsonArray arr => arr,
            JsonValue v when v.TryGetValue(out string? text) => JsonNode.Parse(text) as JsonArray,   // some clients send the list as a JSON string
            _ => null,
        } ?? throw new ToolException("calls must be a list of {tool, args} objects");
        if (list.Count == 0) throw new ToolException("calls is empty");
        if (list.Count > MaxCalls) throw new ToolException($"at most {MaxCalls} calls per batch ({list.Count} given)");
        // check the whole list before running anything
        var calls = list.Select((n, i) =>
        {
            if (n is not JsonObject o || o["tool"]?.GetValue<string>() is not { Length: > 0 } name) throw new ToolException($"calls[{i}] needs a tool name");
            if (Excluded.Contains(name)) throw new ToolException($"calls[{i}]: {name} can't run in a batch; call it on its own");
            if (!tools.TryGetValue(name, out var tool)) throw new ToolException($"calls[{i}]: unknown tool {name}");
            var args = o["args"] switch { null => new JsonObject(), JsonObject x => (JsonObject)x.DeepClone(), _ => throw new ToolException($"calls[{i}]: args must be an object") };
            foreach (var r in References(args).Where(r => r.Call > i))
                throw new ToolException($"calls[{i}] ({name}) refers to {r.Text}, but call {r.Call} {(r.Call == i + 1 ? "is this call" : "runs later")}; a reference can only use an earlier call");
            return (Name: name, Tool: tool, Args: args);
        }).ToList();
        bool keepGoing = a.Bool("continue_on_error") == true;

        var results = new List<(string Label, ToolResult Result)>();
        var outputs = new List<(JsonNode? Json, bool Failed)>();
        int failed = 0;
        for (int i = 0; i < calls.Count; i++)
        {
            ct.ThrowIfCancellationRequested();
            ToolResult r;
            try { r = await calls[i].Tool.InvokeAsync(Resolve(calls[i].Args, outputs), ct); }
            catch (ToolException ex) { r = ToolResult.Fail(ex.Message); }   // a reference that can't be resolved
            results.Add(($"[{i + 1}] {calls[i].Name}", r));
            outputs.Add((FirstJson(r), r.IsError));
            if (r.IsError) { failed++; if (!keepGoing) break; }
        }
        return ToolResult.Batch(results, new
        {
            ran = results.Count, failed, notRun = calls.Count - results.Count,
            note = calls.Count > results.Count ? $"stopped at call {results.Count} ({calls[results.Count - 1].Name}), which failed; the rest didn't run" : null,
        }, a.Bool("compact") == true ? CompactChars : null);
    }

    // ---- references to earlier results (PLAN.md E4): "$2.hash", "$1.positions[0].left"

    // a path starts with a letter, so text like "$1.00" in code is never taken for a reference
    static readonly Regex RefPattern = new(@"^\$(\d+)\.([A-Za-z_]\w*(?:\[\d+\]|\.[A-Za-z_]\w*)*)$", RegexOptions.CultureInvariant);

    private static IEnumerable<(int Call, string Path, string Text)> References(JsonNode? node) => node switch
    {
        JsonObject o => o.SelectMany(kv => References(kv.Value)),
        JsonArray a => a.SelectMany(References),
        JsonValue v when v.TryGetValue(out string? s) && RefPattern.Match(s) is { Success: true } m => [(int.Parse(m.Groups[1].Value), m.Groups[2].Value, s)],
        _ => [],
    };

    /// <summary>A copy of the arguments with every reference replaced by the value it names.</summary>
    private static JsonObject Resolve(JsonObject args, List<(JsonNode? Json, bool Failed)> outputs)
    {
        JsonNode? Walk(JsonNode? node) => node switch
        {
            JsonObject o => new JsonObject(o.Select(kv => KeyValuePair.Create(kv.Key, Walk(kv.Value)))),
            JsonArray a => new JsonArray(a.Select(Walk).ToArray()),
            JsonValue v when v.TryGetValue(out string? s) && RefPattern.Match(s) is { Success: true } m => Lookup(int.Parse(m.Groups[1].Value), m.Groups[2].Value, s),
            _ => node?.DeepClone(),
        };
        JsonNode? Lookup(int call, string path, string text)
        {
            if (call < 1 || call > outputs.Count) throw new ToolException($"{text}: there is no call {call} before this one");
            var (json, bad) = outputs[call - 1];
            if (bad) throw new ToolException($"{text}: call {call} failed, so its result has no values");
            JsonNode? at = json ?? throw new ToolException($"{text}: call {call}'s result isn't JSON (an image or plain text)");
            foreach (Match part in Regex.Matches(path, @"([A-Za-z_]\w*)|\[(\d+)\]"))
            {
                at = part.Groups[1].Success
                    ? (at as JsonObject)?.FirstOrDefault(kv => kv.Key == part.Value).Value ?? (at as JsonObject)?.FirstOrDefault(kv => kv.Key.Equals(part.Value, StringComparison.OrdinalIgnoreCase)).Value
                    : at is JsonArray arr && int.Parse(part.Groups[2].Value) is var ix && ix < arr.Count ? arr[ix] : null;
                if (at is null)
                    throw new ToolException($"{text}: call {call}'s result has no {path}" + (json is JsonObject top ? $" (its fields: {string.Join(", ", top.Select(kv => kv.Key).Take(20))})" : ""));
            }
            return at!.DeepClone();
        }
        return (JsonObject)Walk(args)!;
    }

    /// <summary>The first text item of a result parsed as JSON, or null.</summary>
    private static JsonNode? FirstJson(ToolResult r)
    {
        var text = r.ToJson()["content"]!.AsArray().Select(c => (string?)c?["text"]).FirstOrDefault(t => t is not null);
        if (text is null) return null;
        try { return JsonNode.Parse(text); } catch (JsonException) { return null; }
    }
}
