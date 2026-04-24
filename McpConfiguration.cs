using ExcelPowerPivotMcp.Common;

namespace ExcelPowerPivotMcp;

/// <summary>
/// Configuration for the MCP server.
/// Values can be overridden via environment variables.
/// </summary>
public class McpConfiguration
{
    /// <summary>
    /// Server name reported in MCP protocol.
    /// Override with MCP_SERVER_NAME environment variable.
    /// </summary>
    public string ServerName { get; set; } = "excel-powerpivot-mcp";

    /// <summary>
    /// Server version reported in MCP protocol.
    /// Override with MCP_SERVER_VERSION environment variable.
    /// </summary>
    public string ServerVersion { get; set; } = typeof(McpConfiguration).Assembly.GetName().Version?.ToString(3) ?? "1.0.0";

    /// <summary>
    /// Server description for MCP protocol.
    /// </summary>
    public string ServerDescription { get; set; } = "MCP Server for Excel Power Pivot - enables AI assistants to interact with Power Pivot data models";

    /// <summary>
    /// Maximum rows returned by DAX queries.
    /// Override with MCP_MAX_QUERY_ROWS environment variable.
    /// </summary>
    public int MaxQueryRows { get; set; } = 1000;

    /// <summary>
    /// Path to resources folder containing markdown documentation.
    /// Override with MCP_RESOURCES_PATH environment variable.
    /// </summary>
    public string? ResourcesPath { get; set; }

    // ===== TIMEOUT SETTINGS =====
    
    /// <summary>
    /// Default timeout for DAX query execution in seconds.
    /// Override with MCP_QUERY_TIMEOUT_SECONDS environment variable.
    /// Inspired by ms (200s) but using more conservative default for Excel.
    /// </summary>
    public int QueryTimeoutSeconds { get; set; } = 120;

    /// <summary>
    /// Default timeout for DAX validation (syntax check) in seconds.
    /// Override with MCP_VALIDATION_TIMEOUT_SECONDS environment variable.
    /// Should be shorter than query timeout since no data is returned.
    /// Inspired by ms (10s default).
    /// </summary>
    public int ValidationTimeoutSeconds { get; set; } = 10;

    /// <summary>
    /// Timeout for DMV metadata queries in seconds.
    /// Override with MCP_DMV_TIMEOUT_SECONDS environment variable.
    /// </summary>
    public int DmvTimeoutSeconds { get; set; } = 30;

    // ===== DAX FORMATTER SETTINGS =====

    /// <summary>
    /// Timeout for DAX formatter API calls in seconds.
    /// Override with MCP_DAX_FORMATTER_TIMEOUT_SECONDS environment variable.
    /// Default: 10 seconds (external API call, should be fast).
    /// </summary>
    public int DaxFormatterTimeoutSeconds { get; set; } = 10;

    /// <summary>
    /// Number of retry attempts for DAX formatter API calls.
    /// Override with MCP_DAX_FORMATTER_RETRY_COUNT environment variable.
    /// Default: 2 (one initial attempt + 2 retries = 3 total attempts).
    /// </summary>
    public int DaxFormatterRetryCount { get; set; } = 2;

    /// <summary>
    /// Base delay between DAX formatter retry attempts in milliseconds.
    /// Override with MCP_DAX_FORMATTER_RETRY_DELAY_MS environment variable.
    /// Uses exponential backoff: delay * 2^attempt.
    /// Default: 500ms (fast retry for network blips).
    /// </summary>
    public int DaxFormatterRetryDelayMs { get; set; } = 500;

    /// <summary>
    /// Maximum delay between DAX formatter retries in milliseconds.
    /// Override with MCP_DAX_FORMATTER_MAX_RETRY_DELAY_MS environment variable.
    /// Default: 3000ms (cap exponential backoff).
    /// </summary>
    public int DaxFormatterMaxRetryDelayMs { get; set; } = 3000;

    /// <summary>
    /// Enable in-memory caching of DAX formatter results.
    /// Override with MCP_DAX_FORMATTER_CACHE_ENABLED environment variable.
    /// Default: true (reduces API calls for repeated expressions).
    /// </summary>
    public bool DaxFormatterCacheEnabled { get; set; } = true;

    /// <summary>
    /// Maximum number of entries in the DAX formatter cache.
    /// Override with MCP_DAX_FORMATTER_CACHE_MAX_ENTRIES environment variable.
    /// Default: 100 (reasonable memory footprint ~10-50KB).
    /// </summary>
    public int DaxFormatterCacheMaxEntries { get; set; } = 100;

    /// <summary>
    /// Cache entry time-to-live in minutes.
    /// Override with MCP_DAX_FORMATTER_CACHE_TTL_MINUTES environment variable.
    /// Default: 60 minutes (formatter results are deterministic).
    /// </summary>
    public int DaxFormatterCacheTtlMinutes { get; set; } = 60;

    // ===== RETRY SETTINGS =====
    
    /// <summary>
    /// Number of connection retry attempts.
    /// Override with MCP_CONNECTION_RETRY_COUNT environment variable.
    /// Inspired by mx (default 3).
    /// </summary>
    public int ConnectionRetryCount { get; set; } = 3;

    /// <summary>
    /// Base delay between connection retry attempts in milliseconds.
    /// Override with MCP_CONNECTION_RETRY_DELAY_MS environment variable.
    /// Uses exponential backoff: delay * 2^attempt.
    /// Inspired by mx (1000ms default).
    /// </summary>
    public int ConnectionRetryDelayMs { get; set; } = 1000;

    /// <summary>
    /// Maximum delay between retries in milliseconds (caps exponential backoff).
    /// Override with MCP_CONNECTION_MAX_RETRY_DELAY_MS environment variable.
    /// Inspired by mx (8000ms default).
    /// </summary>
    public int ConnectionMaxRetryDelayMs { get; set; } = 8000;

    /// <summary>
    /// Timeout for each connection attempt in milliseconds.
    /// Override with MCP_CONNECTION_TIMEOUT_MS environment variable.
    /// Default: 10000ms (10 seconds) to accommodate Excel COM initialization delays.
    /// Excel COM calls can be slow on first invocation while the data model loads.
    /// </summary>
    public int ConnectionTimeoutMs { get; set; } = 10000;

    /// <summary>
    /// Current configuration instance. Set after calling Load().
    /// </summary>
    public static McpConfiguration Current { get; private set; } = new();

    /// <summary>
    /// Helper to safely parse integer environment variables with validation warning.
    /// </summary>
    private static bool TryParseEnvInt(string varName, out int value, int minValue = 0)
    {
        value = 0;
        var envValue = Environment.GetEnvironmentVariable(varName);
        if (string.IsNullOrEmpty(envValue))
            return false;

        if (!int.TryParse(envValue, out value))
        {
            Console.Error.WriteLine($"[Warning] Invalid environment variable '{varName}' = '{envValue}'. Expected an integer. Using default value.");
            return false;
        }

        if (value < minValue)
        {
            Console.Error.WriteLine($"[Warning] Environment variable '{varName}' = '{value}' is below minimum ({minValue}). Using default value.");
            return false;
        }

        return true;
    }

    /// <summary>
    /// Helper to safely parse boolean environment variables with validation warning.
    /// </summary>
    private static bool TryParseEnvBool(string varName, out bool value)
    {
        value = false;
        var envValue = Environment.GetEnvironmentVariable(varName);
        if (string.IsNullOrEmpty(envValue))
            return false;

        if (!bool.TryParse(envValue, out value))
        {
            Console.Error.WriteLine($"[Warning] Invalid environment variable '{varName}' = '{envValue}'. Expected 'true' or 'false'. Using default value.");
            return false;
        }

        return true;
    }

    /// <summary>
    /// Load configuration from environment variables with fallback defaults.
    /// Also sets the Current instance for global access.
    /// </summary>
    public static McpConfiguration Load()
    {
        var config = new McpConfiguration();

        // Override from environment variables
        var serverName = Environment.GetEnvironmentVariable("MCP_SERVER_NAME");
        if (!string.IsNullOrEmpty(serverName))
            config.ServerName = serverName;

        var serverVersion = Environment.GetEnvironmentVariable("MCP_SERVER_VERSION");
        if (!string.IsNullOrEmpty(serverVersion))
            config.ServerVersion = serverVersion;

        if (TryParseEnvInt("MCP_MAX_QUERY_ROWS", out var maxRows, minValue: 1))
            config.MaxQueryRows = maxRows;

        var resourcesPath = Environment.GetEnvironmentVariable("MCP_RESOURCES_PATH");
        if (!string.IsNullOrEmpty(resourcesPath))
            config.ResourcesPath = resourcesPath;

        // Timeout settings
        if (TryParseEnvInt("MCP_QUERY_TIMEOUT_SECONDS", out var queryTimeout, minValue: 1))
            config.QueryTimeoutSeconds = queryTimeout;

        if (TryParseEnvInt("MCP_VALIDATION_TIMEOUT_SECONDS", out var valTimeout, minValue: 1))
            config.ValidationTimeoutSeconds = valTimeout;

        if (TryParseEnvInt("MCP_DMV_TIMEOUT_SECONDS", out var dmvTimeout, minValue: 1))
            config.DmvTimeoutSeconds = dmvTimeout;

        // Retry settings
        if (TryParseEnvInt("MCP_CONNECTION_RETRY_COUNT", out var retryCount, minValue: 0))
            config.ConnectionRetryCount = retryCount;

        if (TryParseEnvInt("MCP_CONNECTION_RETRY_DELAY_MS", out var retryDelay, minValue: 0))
            config.ConnectionRetryDelayMs = retryDelay;

        if (TryParseEnvInt("MCP_CONNECTION_MAX_RETRY_DELAY_MS", out var maxRetryDelay, minValue: 1))
            config.ConnectionMaxRetryDelayMs = maxRetryDelay;

        if (TryParseEnvInt("MCP_CONNECTION_TIMEOUT_MS", out var connTimeout, minValue: 1))
            config.ConnectionTimeoutMs = connTimeout;

        // DAX Formatter settings
        if (TryParseEnvInt("MCP_DAX_FORMATTER_TIMEOUT_SECONDS", out var formatterTimeout, minValue: 1))
            config.DaxFormatterTimeoutSeconds = formatterTimeout;

        if (TryParseEnvInt("MCP_DAX_FORMATTER_RETRY_COUNT", out var formatterRetryCount, minValue: 0))
            config.DaxFormatterRetryCount = formatterRetryCount;

        if (TryParseEnvInt("MCP_DAX_FORMATTER_RETRY_DELAY_MS", out var formatterRetryDelay, minValue: 0))
            config.DaxFormatterRetryDelayMs = formatterRetryDelay;

        if (TryParseEnvInt("MCP_DAX_FORMATTER_MAX_RETRY_DELAY_MS", out var formatterMaxRetryDelay, minValue: 1))
            config.DaxFormatterMaxRetryDelayMs = formatterMaxRetryDelay;

        if (TryParseEnvBool("MCP_DAX_FORMATTER_CACHE_ENABLED", out var cacheEnabled))
            config.DaxFormatterCacheEnabled = cacheEnabled;

        if (TryParseEnvInt("MCP_DAX_FORMATTER_CACHE_MAX_ENTRIES", out var cacheMaxEntries, minValue: 1))
            config.DaxFormatterCacheMaxEntries = cacheMaxEntries;

        if (TryParseEnvInt("MCP_DAX_FORMATTER_CACHE_TTL_MINUTES", out var cacheTtl, minValue: 1))
            config.DaxFormatterCacheTtlMinutes = cacheTtl;

        // Set as current instance for global access
        Current = config;
        
        return config;
    }

    /// <summary>
    /// Determine the resources path, checking multiple locations.
    /// </summary>
    public string GetResourcesPath()
    {
        // If explicitly set, use that
        if (!string.IsNullOrEmpty(ResourcesPath) && Directory.Exists(ResourcesPath))
            return ResourcesPath!;

        // Check executable directory
        var exeDir = AppDomain.CurrentDomain.BaseDirectory;
        var exeResources = Path.Combine(exeDir, "Resources");
        if (Directory.Exists(exeResources))
            return exeResources;

        // Fallback to current directory
        var currentResources = Path.Combine(Directory.GetCurrentDirectory(), "Resources");
        if (Directory.Exists(currentResources))
            return currentResources;

        // Return the exe path even if it doesn't exist (for error reporting)
        return exeResources;
    }
}
