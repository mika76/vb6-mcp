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
        foreach (var (n, t, d, _) in props) properties[n] = new JsonObject { ["type"] = t, ["description"] = d };
        var schema = new JsonObject { ["type"] = "object", ["properties"] = properties };
        var required = props.Where(p => p.Required).Select(p => (JsonNode)p.Name).ToArray();
        if (required.Length > 0) schema["required"] = new JsonArray(required);
        return schema;
    }

    public static string? Str(this JsonObject args, string name) => args[name] is JsonValue v && v.TryGetValue(out string? s) ? s : null;

    public static string Req(this JsonObject args, string name)
        => args.Str(name) is { Length: > 0 } s ? s : throw new ToolException($"missing required argument '{name}'");

    public static int? Int(this JsonObject args, string name)
        => args[name] is JsonValue v ? (v.TryGetValue(out int i) ? i : v.TryGetValue(out double d) ? (int)d : v.TryGetValue(out string? s) && int.TryParse(s, out var p) ? p : null) : null;
}
