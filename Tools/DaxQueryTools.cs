using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.Infrastructure.Services;
using ExcelPowerPivotMcp.Core.Services;

namespace ExcelPowerPivotMcp.Tools;

/// <summary>
/// MCP tools for DAX query execution and validation.
/// </summary>
[McpServerToolType]
public class DaxQueryTools
{
    private readonly IDaxService _daxService;
    private readonly IDaxFormatterService _daxFormatterService;
    private readonly McpConfiguration _config;

    public DaxQueryTools(IDaxService daxService, IDaxFormatterService daxFormatterService, McpConfiguration config)
    {
        _daxService = daxService;
        _daxFormatterService = daxFormatterService;
        _config = config;
    }

    [McpServerTool(Name = "run_dax")]
    [Description("Execute a DAX query (EVALUATE) or preview a table by name. Default max 100 rows.")]
    public async Task<string> RunDax(
        [Description("DAX query to execute. Must start with EVALUATE.")] string? daxQuery = null,
        [Description("Table name for quick preview (alternative to full DAX query)")] string? tableName = null,
        [Description("Maximum rows to return (default: 100, max: 1000)")] int maxRows = 100)
    {
        try
        {
            if (string.IsNullOrEmpty(daxQuery) && string.IsNullOrEmpty(tableName))
            {
                return JsonSerializer.Serialize(new { error = "Either dax_query or table_name is required" });
            }

            // Clamp max_rows
            maxRows = Math.Clamp(maxRows, 1, _config.MaxQueryRows);

            List<Dictionary<string, object?>> results;
            string executedQuery;
            
            if (!string.IsNullOrEmpty(tableName))
            {
                // Table preview mode
                results = await _daxService.PreviewTableAsync(tableName, maxRows);
                executedQuery = $"EVALUATE TOPN({maxRows}, '{tableName}')";
            }
            else
            {
                // Full DAX query mode
                executedQuery = daxQuery!;
                results = await _daxService.ExecuteDaxAsync(daxQuery!, maxRows);
                
                // Apply row limit
                if (results.Count > maxRows)
                {
                    results = results.Take(maxRows).ToList();
                }
            }

            var result = new
            {
                success = true,
                rowCount = results.Count,
                data = results,
                query = executedQuery,
                hint = results.Count == maxRows 
                    ? $"Results may be truncated at {maxRows} rows. Use max_rows to increase or add TOPN/ORDER BY to your query."
                    : null
            };
            
            return JsonSerializer.Serialize(result);
        }
        catch (TimeoutException ex)
        {
            return JsonSerializer.Serialize(new { error = $"Query timed out: {ex.Message}", timeout = true });
        }
        catch (Exception ex)
        {
            // Enhanced error hints based on error message analysis
            string hint = GetDaxErrorHint(ex.Message);
            return JsonSerializer.Serialize(new { error = ex.Message, hint });
        }
    }

    [McpServerTool(Name = "validate_dax")]
    [Description("Validate DAX syntax without execution. Returns formatted DAX.")]
    public async Task<string> ValidateDax(
        [Description("DAX expression to validate")] string expression)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                return JsonSerializer.Serialize(new
                {
                    success = false,
                    error = "expression is required",
                    errorType = "InvalidArgument"
                });
            }

            // Use formatter service which validates syntax as a side effect
            var (formatted, warning) = await _daxFormatterService.FormatDaxAsync(expression);

            if (warning != null)
            {
                // Formatter failed but syntax may still be valid
                return JsonSerializer.Serialize(new
                {
                    success = true, // Assume valid unless proven otherwise
                    message = "Syntax not validated (formatter unavailable), but expression accepted",
                    warning,
                    original = expression
                });
            }

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "DAX syntax is valid",
                formatted,
                original = expression
            });
        }
        catch (InvalidOperationException ex) when (ex.Message.Contains("DAX syntax error"))
        {
            // DAX formatter found syntax errors
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = ex.Message,
                errorType = "SyntaxError",
                hint = "Check for missing brackets [], quotes, or parentheses. Verify function names and table/column references."
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                success = false,
                error = ex.Message,
                errorType = ex.GetType().Name
            });
        }
    }

    /// <summary>
    /// Provides specific, actionable hints based on DAX error messages.
    /// </summary>
    private static string GetDaxErrorHint(string errorMessage)
    {
        var lowerError = errorMessage.ToLowerInvariant();

        // Table/Column/Measure not found errors
        if (lowerError.Contains("cannot find") || lowerError.Contains("does not exist"))
        {
            if (lowerError.Contains("table"))
                return "Use list_tables to see available table names (case-sensitive).";
            if (lowerError.Contains("column"))
                return "Use list_columns to see available column names in the table (case-sensitive).";
            if (lowerError.Contains("measure"))
                return "Use list_measures to see available measures.";
            return "Use list_tables or list_columns to verify names exist.";
        }

        // Type mismatch / scalar errors
        if (lowerError.Contains("expect") && (lowerError.Contains("scalar") || lowerError.Contains("single")))
        {
            return "Query returned multiple values but expected a single value. Wrap your expression in an aggregation like SUM(), MIN(), MAX(), or use ROW() for EVALUATE queries.";
        }

        // Syntax errors
        if (lowerError.Contains("syntax error") || lowerError.Contains("unexpected token"))
        {
            return "Check for: (1) missing/extra brackets [], (2) unmatched quotes or parentheses, (3) misspelled DAX functions. Use validate_dax for syntax checking.";
        }

        // Circular dependency
        if (lowerError.Contains("circular") || lowerError.Contains("dependency"))
        {
            return "Circular reference detected. A measure cannot reference itself directly or indirectly through other measures.";
        }

        // Data type errors
        if (lowerError.Contains("type") && lowerError.Contains("mismatch"))
        {
            return "Data type mismatch. Check that numeric operations use numeric columns, text operations use text columns, etc.";
        }

        // Context transition errors
        if (lowerError.Contains("context") || lowerError.Contains("row context"))
        {
            return "Row context issue. Use CALCULATE() to convert row context to filter context, or use iterator functions like SUMX(), FILTER().";
        }

        // Ambiguous column reference
        if (lowerError.Contains("ambiguous"))
        {
            return "Ambiguous column reference. Fully qualify the column with table name: TableName[ColumnName].";
        }

        // Default
        return "For more details, use: (1) validate_dax to check syntax, (2) get_model_summary to review model structure, (3) list_tables and list_columns to verify names.";
    }
}
