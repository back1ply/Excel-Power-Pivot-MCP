using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Common.Utils;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Async implementation of DAX query execution against the Power Pivot data model.
/// All operations are dispatched to the STA thread via IExcelDispatcher.
/// </summary>
public class AdoDaxService : IDaxService
{
    // ADO Connection State constants
    private const int AdStateClosed = 0;
    private const int AdStateOpen = 1;

    private readonly IPowerPivotConnection _connection;
    private readonly IExcelDispatcher _dispatcher;
    private readonly McpConfiguration _config;

    public AdoDaxService(IPowerPivotConnection connection, IExcelDispatcher dispatcher, McpConfiguration config)
    {
        _connection = connection;
        _dispatcher = dispatcher;
        _config = config;
    }

    public async Task<List<Dictionary<string, object?>>> ExecuteDaxAsync(string daxQuery, int maxRows, CancellationToken ct = default)
    {
        // Use configured timeout to prevent hangs on invalid DAX syntax
        var timeout = TimeSpan.FromSeconds(_config.QueryTimeoutSeconds);
        
        return await _dispatcher.RunAsync(() =>
        {
            return _connection.ExecuteWithConnectionValidation(() =>
            {
                dynamic adoConn = _connection.AdoConnection!;
                
                // Check connection state
                int state = (int)adoConn.State;

                if (state == AdStateClosed)
                {
                    throw new InvalidOperationException("ADO Connection is closed. Reconnect the workbook.");
                }
                
                // Execute DAX query via Excel's internal ADO connection
                dynamic? recordset = null;
                try
                {
                    recordset = adoConn.Execute(daxQuery);

                    // Use helper to extract results with max rows limit
                    return AdoRecordsetHelper.ExtractResults(recordset, maxRows, useOrdinalIgnoreCase: false);
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    // Parse Analysis Services error messages for actionable feedback
                    string hint = ParseDaxError(comEx.Message);
                    throw new InvalidOperationException($"DAX execution failed: {hint}", comEx);
                }
                finally
                {
                    // Recordset is a COM object, should be released
                    if (recordset != null)
                    {
                        ComHelper.SafeRelease(recordset);
                    }
                }
            }, "executing DAX query");
        }, timeout, ct);
    }

    public async Task<List<Dictionary<string, object?>>> PreviewTableAsync(string tableName, int maxRows, CancellationToken ct = default)
    {
        // Clean up any existing delimiters and sanitize to prevent DAX injection
        string cleanName = tableName.Trim('[', ']', '\'');
        string sanitizedName = DaxHelpers.SanitizeValue(cleanName);
        // DAX table references use single quotes: 'TableName'
        var query = $"EVALUATE TOPN({maxRows}, '{sanitizedName}')";
        return await ExecuteDaxAsync(query, maxRows, ct);
    }

    public async Task<ColumnProfile> GetColumnProfileAsync(string tableName, string columnName, CancellationToken ct = default)
    {
        // Use DaxHelpers for clean column reference building (already sanitizes internally)
        string fullRef = DaxHelpers.BuildColumnReference(tableName, columnName);
        string cleanTableName = tableName.Trim('[', ']', '\'');
        string sanitizedTableName = DaxHelpers.SanitizeValue(cleanTableName);
        string tName = $"'{sanitizedTableName}'";

        var profile = new ColumnProfile();

        // 1. Stats query
        var statsQuery = $@"
            EVALUATE 
            ROW(
                ""Min"", MIN({fullRef}),
                ""Max"", MAX({fullRef}),
                ""DistinctCount"", DISTINCTCOUNT({fullRef}),
                ""Count"", COUNT({fullRef}),
                ""NullCount"", COUNTROWS({tName}) - COUNT({fullRef})
            )";

        try
        {
            var stats = await ExecuteDaxAsync(statsQuery, 100, ct); // Stats always return 1 row
            if (stats.Count > 0)
            {
                var row = stats[0];
                // DAX ROW() may return keys with brackets like [Min] or without
                object? GetValue(string key) =>
                    row.TryGetValue(key, out var val) ? val :
                    row.TryGetValue($"[{key}]", out var brVal) ? brVal : null;
                
                profile.Min = GetValue("Min");
                profile.Max = GetValue("Max");
                profile.DistinctCount = Convert.ToInt64(GetValue("DistinctCount") ?? 0, System.Globalization.CultureInfo.InvariantCulture);
                profile.Count = Convert.ToInt64(GetValue("Count") ?? 0, System.Globalization.CultureInfo.InvariantCulture);
                profile.NullCount = Convert.ToInt64(GetValue("NullCount") ?? 0, System.Globalization.CultureInfo.InvariantCulture);
            }
        }
        catch (Exception ex) 
        { 
            profile.Error = ex.Message; 
        }

        // 2. Sample values query
        try
        {
            var sampleQuery = $"EVALUATE TOPN(50, VALUES({fullRef}), {fullRef}, 1)";
            var samples = await ExecuteDaxAsync(sampleQuery, 50, ct);
            foreach (var r in samples)
            {
                if (r.Count > 0) profile.DistinctValues.Add(r.Values.First());
            }
        }
        catch (Exception ex) 
        { 
            profile.Error += " | Sample Error: " + ex.Message; 
        }

        return profile;
    }

    /// <summary>
    /// Parse DAX error messages for actionable feedback to help LLM self-correct.
    /// </summary>
    private static string ParseDaxError(string errorMessage)
    {
        var lowerError = errorMessage.ToLowerInvariant();
        
        // Column not found
        if (lowerError.Contains("cannot find") && lowerError.Contains("column"))
        {
            // Try to extract column name from error like "The column '[Table].[Column]' cannot be found"
            var match = System.Text.RegularExpressions.Regex.Match(
                errorMessage, @"\[([^\]]+)\]\.\[([^\]]+)\]");
            if (match.Success)
            {
                return $"Column not found: [{match.Groups[1].Value}].[{match.Groups[2].Value}]. " +
                       "Use list_columns to verify the exact column name (case-sensitive).";
            }
            return $"{errorMessage}. Use list_columns to verify column names.";
        }
        
        // Table not found
        if (lowerError.Contains("cannot find") && lowerError.Contains("table"))
        {
            return $"{errorMessage}. Use list_tables to verify the exact table name.";
        }
        
        // Measure not found
        if (lowerError.Contains("cannot find") && lowerError.Contains("measure"))
        {
            return $"{errorMessage}. Use list_measures to verify the exact measure name.";
        }
        
        // Syntax error
        if (lowerError.Contains("syntax error") || lowerError.Contains("unexpected token"))
        {
            return $"{errorMessage}. Check for missing brackets [], quotes, or parentheses in your DAX.";
        }
        
        // Circular reference
        if (lowerError.Contains("circular"))
        {
            return $"{errorMessage}. A measure cannot reference itself (directly or indirectly).";
        }
        
        // Function not found
        if (lowerError.Contains("function") && lowerError.Contains("not"))
        {
            return $"{errorMessage}. Verify the DAX function name spelling.";
        }
        
        // Default: return original with hint
        return $"{errorMessage}. Use get_model_summary to review the data model structure.";
    }
}
