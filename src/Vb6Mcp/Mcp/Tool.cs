using System.Text.Json;
using System.Text.Json.Nodes;

namespace Vb6Mcp.Mcp;

/// <summary>An MCP tool: name, description, JSON schema of its arguments, and the handler.</summary>
public sealed class Tool(string name, string description, JsonObject inputSchema, Func<JsonObject, CancellationToken, Task<ToolResult>> handler)
{
    public string Name { get; } = name;

    public JsonObject Definition() => new()
    {
        ["name"] = Name,
        ["description"] = description,
        ["inputSchema"] = inputSchema.DeepClone(),
    };

    /// <summary>The same tool with <paramref name="check"/> run before the handler (it throws ToolException to refuse).</summary>
    public Tool Guarded(Action<JsonObject> check) => new(Name, description, inputSchema, (a, ct) => { check(a); return handler(a, ct); });

    /// <summary>
    /// The same tool with <paramref name="before"/> run first (it can't refuse) and its value handed to <paramref name="after"/>
    /// together with the result, also for failed calls (ide crash notices ride along with whatever the agent called next).
    /// </summary>
    public Tool Around<T>(Func<T> before, Func<T, ToolResult, ToolResult> after)
        => new(Name, description, inputSchema, async (a, ct) => { var state = before(); return after(state, await InvokeAsync(a, ct).ConfigureAwait(false)); });

    /// <summary>Runs the handler. Expected failures (<see cref="ToolException"/>) and unexpected ones become isError results.</summary>
    public async Task<ToolResult> InvokeAsync(JsonObject args, CancellationToken ct)
    {
        try { return await handler(args, ct).ConfigureAwait(false); }
        catch (ToolException ex) { return ToolResult.Fail(ex.Message); }
        // Known failure types already carry a message written for the agent
        catch (Exception ex) when (ex is Core.Bridge.AddinErrorException or Core.Bridge.AddinBusyException or Core.Bridge.AddinUnreachableException)
        { return ToolResult.Fail(ex.Message); }
        catch (Exception ex) { return ToolResult.Fail($"{ex.GetType().Name}: {ex.Message}"); }
    }
}

/// <summary>Result of a tool call: MCP content items plus the isError flag.</summary>
public sealed class ToolResult
{
    private readonly JsonArray _content = [];
    public bool IsError { get; private init; }

    /// <summary>A successful result whose single text item is <paramref name="value"/> serialized as JSON.</summary>
    public static ToolResult Ok(object value) => Text(JsonSerializer.Serialize(value, McpServer.Json));

    public static ToolResult Text(string text)
    {
        var r = new ToolResult();
        r._content.Add(new JsonObject { ["type"] = "text", ["text"] = text });
        return r;
    }

    /// <summary>An image (MCP image content) followed by a JSON text item describing it.</summary>
    public static ToolResult Image(byte[] png, object info)
    {
        var r = new ToolResult();
        r._content.Add(new JsonObject { ["type"] = "image", ["data"] = Convert.ToBase64String(png), ["mimeType"] = "image/png" });
        r._content.Add(new JsonObject { ["type"] = "text", ["text"] = JsonSerializer.Serialize(info, McpServer.Json) });
        return r;
    }

    /// <summary>
    /// This result with one more field: added to the JSON object of the last text item, or (for an error or plain text)
    /// appended as a line. Used for cross-cutting notes such as reloads applied before a compile.
    /// </summary>
    public ToolResult With(string key, JsonNode? value)
    {
        if (value is null || _content.Count == 0 || _content[^1] is not JsonObject last || (string?)last["type"] != "text") return this;
        string text = (string?)last["text"] ?? "";
        if (!IsError && text.StartsWith('{') && JsonNode.Parse(text) is JsonObject obj)
        {
            obj[key] = value;
            last["text"] = obj.ToJsonString(McpServer.Json);
        }
        else last["text"] = text + "\n" + key + ": " + value.ToJsonString();
        return this;
    }

    /// <summary>
    /// batch_run's result: each call's content in order, its text items prefixed with a label ("[2] app_click: ...") and
    /// errors marked, images kept as they are, then a JSON summary. isError when any call failed.
    /// </summary>
    public static ToolResult Batch(IEnumerable<(string Label, ToolResult Result)> items, object summary, int? clip = null)
    {
        var r = new ToolResult();
        bool anyError = false;
        var list = items.ToList();
        for (int i = 0; i < list.Count; i++)
        {
            var (label, item) = list[i];
            anyError |= item.IsError;
            // clip: successful calls other than the last one keep only the start of their text (batch_run compact)
            bool cut = clip is int max && !item.IsError && i < list.Count - 1;
            foreach (var c in item._content)
            {
                var copy = c!.DeepClone().AsObject();
                if ((string?)copy["type"] == "text")
                {
                    var text = (string?)copy["text"] ?? "";
                    if (cut && text.Length > clip!.Value) text = text[..clip.Value] + $"... ({text.Length} chars, cut by compact)";
                    copy["text"] = $"{label}{(item.IsError ? " (error)" : "")}: {text}";
                }
                r._content.Add(copy);
            }
        }
        r._content.Add(new JsonObject { ["type"] = "text", ["text"] = JsonSerializer.Serialize(summary, McpServer.Json) });
        return anyError ? new ToolResult(r._content, true) : r;
    }

    private ToolResult() { }
    private ToolResult(JsonArray content, bool isError) { _content = content; IsError = isError; }

    public static ToolResult Fail(string message)
    {
        var r = new ToolResult { IsError = true };
        r._content.Add(new JsonObject { ["type"] = "text", ["text"] = message });
        return r;
    }

    public JsonObject ToJson() => new() { ["content"] = _content.DeepClone(), ["isError"] = IsError };
}

/// <summary>An expected tool failure with a message meant for the agent (e.g. "several IDEs are running, pass pid").</summary>
public sealed class ToolException(string message) : Exception(message);

/// <summary>Helpers for building input schemas and reading arguments.</summary>
public static class Schema
{
    public static JsonObject Object(params (string Name, string Type, string Description, bool Required)[] props)
    {
        var properties = new JsonObject();
        foreach (var (n, t, d, _) in props)
        {
            properties[n] = new JsonObject { ["type"] = t, ["description"] = d };
            if (t == "array") properties[n]!["items"] = new JsonObject { ["type"] = "string" };   // arrays are string lists
        }
        var schema = new JsonObject { ["type"] = "object", ["properties"] = properties };
        var required = props.Where(p => p.Required).Select(p => (JsonNode)p.Name).ToArray();
        if (required.Length > 0) schema["required"] = new JsonArray(required);
        return schema;
    }

    public static string? Str(this JsonObject args, string name) => args[name] is JsonValue v && v.TryGetValue(out string? s) ? s : null;

    public static string Req(this JsonObject args, string name)
        => args.Str(name) is { Length: > 0 } s ? s : throw new ToolException($"missing required argument '{name}'");

    /// <summary>A boolean argument; also accepts "true"/"false" strings (some clients send those).</summary>
    public static bool? Bool(this JsonObject args, string name)
        => args[name] is JsonValue v ? (v.TryGetValue(out bool b) ? b : v.TryGetValue(out string? s) && bool.TryParse(s, out var p) ? p : null) : null;

    /// <summary>A string-list argument: a JSON array, or one string with items separated by ';' or line breaks (some clients send those).</summary>
    public static IReadOnlyList<string> List(this JsonObject args, string name) => args[name] switch
    {
        JsonArray a => a.Select(n => n is JsonValue v && v.TryGetValue(out string? s) ? s : null).OfType<string>().Where(s => s.Length > 0).ToList(),
        JsonValue v when v.TryGetValue(out string? s) => s.Split([';', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries),
        _ => [],
    };

    public static int? Int(this JsonObject args, string name)
        => args[name] is JsonValue v ? (v.TryGetValue(out int i) ? i : v.TryGetValue(out double d) ? (int)d : v.TryGetValue(out string? s) && int.TryParse(s, out var p) ? p : null) : null;
}
