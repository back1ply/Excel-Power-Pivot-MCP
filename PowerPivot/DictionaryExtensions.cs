namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// Extension methods for dictionary access to reduce boilerplate ContainsKey patterns.
/// </summary>
public static class DictionaryExtensions
{
    /// <summary>
    /// Get a value from dictionary, returning default if key doesn't exist or value is wrong type.
    /// </summary>
    public static T? Get<T>(this Dictionary<string, object?> dict, string key)
    {
        return dict.TryGetValue(key, out var val) && val is T typed ? typed : default;
    }

    /// <summary>
    /// Get a value from dictionary with a fallback default value.
    /// </summary>
    public static T Get<T>(this Dictionary<string, object?> dict, string key, T defaultValue)
    {
        return dict.TryGetValue(key, out var val) && val is T typed ? typed : defaultValue;
    }

    /// <summary>
    /// Get a string value, trying multiple keys in order (for DMV column name variations).
    /// </summary>
    public static string? GetFirstString(this Dictionary<string, object?> dict, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (dict.TryGetValue(key, out var val) && val is string s)
                return s;
        }
        return null;
    }

    /// <summary>
    /// Get any value from multiple possible keys, returning first match.
    /// </summary>
    public static object? GetFirst(this Dictionary<string, object?> dict, params string[] keys)
    {
        foreach (var key in keys)
        {
            if (dict.TryGetValue(key, out var val) && val != null)
                return val;
        }
        return null;
    }
}
