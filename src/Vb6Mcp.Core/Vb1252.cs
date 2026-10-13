using System.Text;

namespace Vb6Mcp.Core;

/// <summary>
/// Windows-1252: the only encoding VB6 source files and the wire protocol use (AGENTS.md hard rule 1).
/// .NET only ships it after registering <see cref="CodePagesEncodingProvider"/>.
/// </summary>
public static class Vb1252
{
    private static readonly Lazy<Encoding> s_encoding = new(() =>
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(1252);
    });

    public static Encoding Encoding => s_encoding.Value;

    private static readonly Lazy<Encoding> s_strict = new(() =>
    {
        _ = s_encoding.Value;                       // registers the provider
        return Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
    });

    /// <summary>1252 that throws for characters it can't represent (instead of writing '?').</summary>
    public static Encoding Strict => s_strict.Value;
}
