namespace ExcelPowerPivotMcp.Common;

using System.Reflection;

/// <summary>
/// Helper to read embedded resources from the assembly.
/// Used for single-file deployment where Resources folder is embedded.
/// </summary>
public static class EmbeddedResources
{
    private static readonly Assembly Assembly = typeof(EmbeddedResources).Assembly;
    private const string ResourcePrefix = "ExcelPowerPivotMcp.Resources.";

    /// <summary>
    /// Get all embedded resource names under Resources folder.
    /// </summary>
    public static IEnumerable<string> GetResourceNames()
    {
        return Assembly.GetManifestResourceNames()
            .Where(n => n.StartsWith(ResourcePrefix, StringComparison.OrdinalIgnoreCase))
            .Select(n => n.Substring(ResourcePrefix.Length));
    }

    /// <summary>
    /// Read an embedded resource as string.
    /// </summary>
    /// <param name="resourceName">Name like "common_workflows.md"</param>
    public static string? ReadResource(string resourceName)
    {
        var fullName = ResourcePrefix + resourceName;
        using var stream = Assembly.GetManifestResourceStream(fullName);
        if (stream == null) return null;
        
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }

    /// <summary>
    /// Check if an embedded resource exists.
    /// </summary>
    public static bool ResourceExists(string resourceName)
    {
        var fullName = ResourcePrefix + resourceName;
        return Assembly.GetManifestResourceNames().Contains(fullName);
    }
}
