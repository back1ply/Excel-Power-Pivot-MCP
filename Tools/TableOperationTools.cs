using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.PowerPivot;
using ExcelPowerPivotMcp.Core.Services;
using ExcelPowerPivotMcp.Infrastructure.Services;
using ExcelPowerPivotMcp.Common;

namespace ExcelPowerPivotMcp.Tools;

/// <summary>
/// MCP tools for Excel table and Power Query operations.
/// </summary>
[McpServerToolType]
public class TableOperationTools
{
    private readonly ITableService _tableService;
    private readonly ITableComService _tableComService;

    // M-Code preview length: truncate to 97 chars + "..." = 100 chars total for readable display
    private const int MCodePreviewLength = 97;

    public TableOperationTools(ITableService tableService, ITableComService tableComService)
    {
        _tableService = tableService;
        _tableComService = tableComService;
    }

    /// <summary>
    /// Lists all Excel tables in the workbook and their data model status.
    /// </summary>
    /// <returns>JSON containing table count, model status, and table details.</returns>
    [McpServerTool(Name = "list_excel_tables")]
    [Description("List Excel tables and their data model status.")]
    public async Task<string> ListExcelTables()
    {
        var tables = await _tableService.GetExcelTablesAsync();
        var inModel = tables.Count(t => t.ContainsKey("isInModel") && (bool)t["isInModel"]!);

        var result = new
        {
            tableCount = tables.Count,
            inDataModel = inModel,
            availableToAdd = tables.Count - inModel,
            tables
        };
        return JsonSerializer.Serialize(result);
    }

    /// <summary>
    /// Adds an Excel table to the Power Pivot data model.
    /// </summary>
    /// <param name="tableName">Name of the Excel table to add.</param>
    /// <param name="usePowerQuery">Whether to use Power Query method (more robust).</param>
    /// <param name="queryName">Optional query name (must match tableName if provided).</param>
    /// <param name="mCode">Optional M-code for Power Query transformation.</param>
    /// <param name="confirm">Confirmation flag required for execution.</param>
    /// <returns>JSON result with operation status and details.</returns>
    [McpServerTool(Name = "add_excel_table_to_model")]
    [Description("Add an Excel table to the data model. Requires confirm=true.")]
    public async Task<string> AddExcelTableToModel(
        [Description("Name of the Excel table to add")] string tableName,
        [Description("Set to true to confirm adding table.")] bool confirm = false)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            throw new ArgumentException("table_name is required");
            
        if (!confirm)
        {
            return JsonSerializer.Serialize(new
            {
                requiresConfirmation = true,
                message = $"This will add table '{tableName}' to the data model. Call again with confirm=true to proceed.",
            });
        }

        var request = new Common.DataStructures.TableAddToModel
        {
            TableName = tableName,
            UsePowerQuery = false
        };

        var actualTableName = await _tableService.AddTableToModelAsync(request);

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Successfully added table '{actualTableName}'.",
            data = new { tableName = actualTableName, requestedName = tableName, method = "direct" }
        });
    }

    [McpServerTool(Name = "add_power_query_table_to_model")]
    [Description("Add a table via Power Query (M-Code). Requires confirm=true.")]
    public async Task<string> AddPowerQueryTableToModel(
        [Description("Name of the resulting table in the model")] string tableName,
        [Description("M-Code for the query.")] string mCode,
        [Description("Name for the query (optional, defaults to tableName)")] string? queryName = null,
        [Description("Set to true to confirm adding table.")] bool confirm = false)
    {
        if (string.IsNullOrWhiteSpace(tableName))
            throw new ArgumentException("table_name is required");
        if (string.IsNullOrWhiteSpace(mCode))
             throw new ArgumentException("m_code is required");

        if (!confirm)
        {
            return JsonSerializer.Serialize(new
            {
                requiresConfirmation = true,
                message = $"This will add table '{tableName}' via Power Query. Call again with confirm=true to proceed.",
            });
        }

        var request = new Common.DataStructures.TableAddToModel
        {
            TableName = tableName,
            UsePowerQuery = true,
            QueryName = queryName,
            MCode = mCode
        };

        var actualTableName = await _tableService.AddTableToModelAsync(request);

        var response = new
        {
            success = true,
            message = actualTableName.Equals(tableName, StringComparison.OrdinalIgnoreCase)
                ? $"Successfully added table '{actualTableName}' via Power Query."
                : $"Successfully added table '{actualTableName}' via Power Query. Note: Excel renamed the table from requested name '{tableName}'.",
            data = new { tableName = actualTableName, requestedName = tableName, method = "power_query" }
        };

        return JsonSerializer.Serialize(response);
    }



    // NOTE: set_column_description tool removed - Excel COM API's ModelTableColumn object
    // does NOT have a Description property. Only available properties are: DataType, Name,
    // Application, Creator, Parent. Column descriptions can only be set via the Power Pivot UI.

    /// <summary>
    /// Refreshes a specific table in the data model from its source.
    /// </summary>
    /// <param name="tableName">Name of the table to refresh.</param>
    /// <returns>JSON result with operation status.</returns>
    [McpServerTool(Name = "refresh_table")]
    [Description("Refresh a specific table from its data source.")]
    public async Task<string> RefreshTable(
        [Description("Name of the table to refresh")] string tableName)
    {
        await _tableComService.RefreshTableAsync(tableName);

        var result = new
        {
            success = true,
            message = $"Successfully refreshed table '{tableName}'."
        };
        return JsonSerializer.Serialize(result);
    }

    /// <summary>
    /// Deletes a table from the Power Pivot data model.
    /// </summary>
    /// <param name="tableName">Name of the table to delete.</param>
    /// <param name="confirm">Confirmation flag required for execution.</param>
    /// <returns>JSON result with operation status.</returns>
    [McpServerTool(Name = "delete_table")]
    [Description("Delete a table from the data model. Requires confirm=true.")]
    public async Task<string> DeleteTable(
        [Description("Name of the table to delete")] string tableName,
        [Description("Set to true to confirm deletion. Required for safety.")] bool confirm = false)
    {
        if (!confirm)
        {
            return JsonSerializer.Serialize(new
            {
                requiresConfirmation = true,
                message = $"This will permanently delete table '{tableName}' from the data model. Call again with confirm=true to proceed.",
                warning = "This will break conflicting measures and relationships!"
            });
        }

        await _tableService.DeleteTableAsync(tableName);

        var result = new
        {
            success = true,
            message = $"Successfully deleted table '{tableName}'. Remember to call save_workbook to persist changes."
        };
        return JsonSerializer.Serialize(result);
    }
}
