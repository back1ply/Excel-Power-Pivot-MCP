namespace ExcelPowerPivotMcp.Common.DataStructures;

/// <summary>
/// Standardized error response format for MCP tools.
/// Provides consistent error reporting with actionable hints.
/// </summary>
public class ToolErrorResponse
{
    /// <summary>
    /// Human-readable error message describing what went wrong.
    /// </summary>
    public required string Error { get; init; }

    /// <summary>
    /// Classification of the error type for programmatic handling.
    /// Common values: InvalidArgument, NotFound, Timeout, ConnectionLost, SyntaxError, PermissionDenied
    /// </summary>
    public string? ErrorType { get; init; }

    /// <summary>
    /// Actionable suggestion to help the user resolve the error.
    /// Example: "Use list_tables to see available table names"
    /// </summary>
    public string? Hint { get; init; }

    /// <summary>
    /// Optional technical details for debugging (stack trace, inner exception, etc.).
    /// </summary>
    public string? Details { get; init; }

    /// <summary>
    /// Indicates if the error was due to a timeout.
    /// </summary>
    public bool Timeout { get; init; }

    /// <summary>
    /// Create a simple error response with just a message.
    /// </summary>
    public static ToolErrorResponse Simple(string error) => new() { Error = error };

    /// <summary>
    /// Create an invalid argument error response.
    /// </summary>
    public static ToolErrorResponse InvalidArgument(string error, string? hint = null) => new()
    {
        Error = error,
        ErrorType = "InvalidArgument",
        Hint = hint
    };

    /// <summary>
    /// Create a not found error response.
    /// </summary>
    public static ToolErrorResponse NotFound(string resource, string? hint = null) => new()
    {
        Error = $"{resource} not found",
        ErrorType = "NotFound",
        Hint = hint ?? "Use the appropriate list_* tool to see available items"
    };

    /// <summary>
    /// Create a timeout error response.
    /// </summary>
    public static ToolErrorResponse TimeoutError(string operation, string message, string? hint = null) => new()
    {
        Error = $"{operation} timed out: {message}",
        ErrorType = "Timeout",
        Timeout = true,
        Hint = hint
    };

    /// <summary>
    /// Create a syntax error response (typically for DAX).
    /// </summary>
    public static ToolErrorResponse SyntaxError(string message, string? hint = null) => new()
    {
        Error = message,
        ErrorType = "SyntaxError",
        Hint = hint
    };

    /// <summary>
    /// Create a general exception error response.
    /// </summary>
    public static ToolErrorResponse FromException(Exception ex, string? hint = null) => new()
    {
        Error = ex.Message,
        ErrorType = ex.GetType().Name,
        Hint = hint,
        Details = ex.StackTrace
    };
}
