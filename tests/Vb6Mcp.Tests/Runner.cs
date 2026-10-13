// Minimal test runner (AGENTS.md: no test framework package, keep it ~100 lines).
//   dotnet run --project tests/Vb6Mcp.Tests                 -> uncategorised (unit) tests
//   dotnet run --project tests/Vb6Mcp.Tests -- --cat IDE     -> only tests in category IDE
//   dotnet run --project tests/Vb6Mcp.Tests -- --all         -> everything
// Exit code = number of failures (0 = all passed).

using System.Diagnostics;
using System.Reflection;

namespace Vb6Mcp.Tests;

/// <summary>Marks a public static or instance method (no parameters) as a test.</summary>
[AttributeUsage(AttributeTargets.Method)]
public sealed class TestAttribute : Attribute { }

/// <summary>Puts a test in a category; categorised tests only run when requested (needs VB6, elevation, ...).</summary>
[AttributeUsage(AttributeTargets.Method)]
public sealed class CategoryAttribute(string name) : Attribute { public string Name { get; } = name; }

/// <summary>Assertion helpers; failures throw <see cref="AssertFailed"/>.</summary>
public static class Assert
{
    public static void True(bool condition, string message = "expected true")
    { if (!condition) throw new AssertFailed(message); }

    public static void Equal<T>(T expected, T actual, string? what = null)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
            throw new AssertFailed($"{what ?? "value"}: expected <{expected}> but was <{actual}>");
    }

    public static void SequenceEqual<T>(IEnumerable<T> expected, IEnumerable<T> actual, string? what = null)
    {
        var e = expected.ToArray(); var a = actual.ToArray();
        if (!e.SequenceEqual(a))
            throw new AssertFailed($"{what ?? "sequence"}: expected [{string.Join(", ", e)}] but was [{string.Join(", ", a)}]");
    }

    public static TEx Throws<TEx>(Action action) where TEx : Exception
    {
        try { action(); }
        catch (TEx ex) { return ex; }
        catch (Exception ex) { throw new AssertFailed($"expected {typeof(TEx).Name} but got {ex.GetType().Name}: {ex.Message}"); }
        throw new AssertFailed($"expected {typeof(TEx).Name} but nothing was thrown");
    }
}

public sealed class AssertFailed(string message) : Exception(message);

public static class Runner
{
    public static int Main(string[] args)
    {
        bool all = args.Contains("--all");
        var cats = args.SkipWhile(a => a != "--cat").Skip(1).FirstOrDefault()?.Split(',', StringSplitOptions.RemoveEmptyEntries) ?? [];

        var tests = typeof(Runner).Assembly.GetTypes()
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance))
            .Where(m => m.GetCustomAttribute<TestAttribute>() != null)
            .Select(m => (Method: m, Category: m.GetCustomAttribute<CategoryAttribute>()?.Name))
            .Where(t => all || (cats.Length == 0 ? t.Category == null : cats.Contains(t.Category, StringComparer.OrdinalIgnoreCase)))
            .OrderBy(t => t.Method.DeclaringType!.Name).ThenBy(t => t.Method.Name)
            .ToList();

        int failed = 0;
        var total = Stopwatch.StartNew();
        foreach (var (method, category) in tests)
        {
            string name = $"{method.DeclaringType!.Name}.{method.Name}" + (category is null ? "" : $" [{category}]");
            var sw = Stopwatch.StartNew();
            try
            {
                object? instance = method.IsStatic ? null : Activator.CreateInstance(method.DeclaringType!);
                var result = method.Invoke(instance, null);
                if (result is Task task) task.GetAwaiter().GetResult();
                Console.WriteLine($"  PASS  {name} ({sw.ElapsedMilliseconds} ms)");
            }
            catch (Exception ex)
            {
                failed++;
                var inner = ex is TargetInvocationException tie && tie.InnerException != null ? tie.InnerException : ex;
                Console.WriteLine($"  FAIL  {name}: {inner.GetType().Name}: {inner.Message}");
                if (inner is not AssertFailed)
                    foreach (var frame in (inner.StackTrace ?? "").Split('\n').Take(4))
                        Console.WriteLine("        " + frame.Trim());
            }
        }
        Console.WriteLine($"{tests.Count - failed}/{tests.Count} passed in {total.ElapsedMilliseconds} ms");
        return failed;
    }
}
