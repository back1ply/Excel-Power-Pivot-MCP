using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.PowerPivot;
using ExcelPowerPivotMcp.Common;

namespace ExcelPowerPivotMcp.Tools;

/// <summary>
/// MCP tools for measure CRUD operations.
/// </summary>
[McpServerToolType]
public class MeasureCrudTools
{
    private readonly Core.Services.IMeasureService _measureService;
    private readonly Core.Services.IDaxFormatterService _daxFormatterService;

    public MeasureCrudTools(
        Core.Services.IMeasureService measureService,
        Core.Services.IDaxFormatterService daxFormatterService)
    {
        _measureService = measureService;
        _daxFormatterService = daxFormatterService;
    }

    [McpServerTool(Name = "create_measure")]
    [Description("Create a DAX measure in a table. Call save_workbook to persist.")]
    public async Task<string> CreateMeasure(
        [Description("Name of the table to add the measure to")] string tableName,
        [Description("Name of the new measure")] string measureName,
        [Description("DAX expression for the measure (e.g., 'SUM(Sales[Amount])')")] string expression,
        [Description("Optional description for the measure")] string? description = null,
        [Description("Auto-format DAX for readability (default: true). Set to false for faster performance.")] bool autoFormat = true)
    {
        try
        {
            // Parameter validation
            if (string.IsNullOrWhiteSpace(tableName))
                return JsonSerializer.Serialize(new { error = "table_name is required" });
            if (string.IsNullOrWhiteSpace(measureName))
                return JsonSerializer.Serialize(new { error = "measure_name is required" });
            if (string.IsNullOrWhiteSpace(expression))
                return JsonSerializer.Serialize(new { error = "expression is required" });
                
            // Auto-format DAX expression if enabled
            string? formatterWarning = null;
            var formattedExpression = expression;
            if (autoFormat)
            {
                (formattedExpression, formatterWarning) = await _daxFormatterService.FormatDaxAsync(expression);
            }

            var request = new Common.DataStructures.MeasureCreate
            {
                TableName = tableName,
                MeasureName = measureName,
                Expression = formattedExpression,
                Description = description
            };

            await _measureService.CreateMeasureAsync(request);

            var result = new
            {
                success = true,
                message = $"Successfully created measure '{measureName}' in table '{tableName}'. Remember to call save_workbook to persist changes.",
                formatted = autoFormat,
                formatterWarning = formatterWarning,
                data = request
            };
            return JsonSerializer.Serialize(result);
        }
        catch (TimeoutException ex)
        {
            return JsonSerializer.Serialize(new { error = $"Operation timed out: {ex.Message}", timeout = true });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool(Name = "update_measure")]
    [Description("Update a measure's expression, description, or name. Call save_workbook to persist.")]
    public async Task<string> UpdateMeasure(
        [Description("Current name of the measure to update")] string measureName,
        [Description("New DAX expression")] string? newExpression = null,
        [Description("New description")] string? newDescription = null,
        [Description("New name for the measure")] string? newName = null,
        [Description("Auto-format DAX for readability (default: true). Set to false for faster performance.")] bool autoFormat = true)
    {
        try
        {
            // Parameter validation
            if (string.IsNullOrWhiteSpace(measureName))
                return JsonSerializer.Serialize(new { error = "measure_name is required" });
                
            if (string.IsNullOrEmpty(newExpression) && newDescription == null && string.IsNullOrEmpty(newName))
            {
                return JsonSerializer.Serialize(new { error = "At least one of new_expression, new_description, or new_name must be provided" });
            }

            // Auto-format DAX expression if provided and enabled
            string? formattedExpression = null;
            string? formatterWarning = null;
            if (!string.IsNullOrEmpty(newExpression))
            {
                if (autoFormat)
                {
                    (formattedExpression, formatterWarning) = await _daxFormatterService.FormatDaxAsync(newExpression!);
                }
                else
                {
                    formattedExpression = newExpression;
                }
            }

            var request = new Common.DataStructures.MeasureUpdate
            {
                MeasureName = measureName,
                Expression = formattedExpression,
                Description = newDescription,
                NewName = newName
            };

            await _measureService.UpdateMeasureAsync(request);

            var changes = new List<string>();
            if (!string.IsNullOrEmpty(formattedExpression)) changes.Add("expression");
            if (newDescription != null) changes.Add("description");
            if (!string.IsNullOrEmpty(newName)) changes.Add("name");

            var result = new
            {
                success = true,
                message = $"Successfully updated measure '{measureName}'. Remember to call save_workbook to persist changes.",
                formatted = autoFormat && !string.IsNullOrEmpty(newExpression),
                formatterWarning = formatterWarning,
                updatedFields = changes,
                newName = newName ?? measureName
            };
            return JsonSerializer.Serialize(result);
        }
        catch (TimeoutException ex)
        {
            return JsonSerializer.Serialize(new { error = $"Operation timed out: {ex.Message}", timeout = true });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }

    [McpServerTool(Name = "delete_measure")]
    [Description("Delete a measure. Requires confirm=true.")]
    public async Task<string> DeleteMeasure(
        [Description("Name of the measure to delete")] string measureName,
        [Description("Set to true to confirm deletion. Required for safety.")] bool confirm = false)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(measureName))
                return JsonSerializer.Serialize(new { error = "measure_name is required" });
                
            if (!confirm)
            {
                return JsonSerializer.Serialize(new
                {
                    requiresConfirmation = true,
                    message = $"This will permanently delete measure '{measureName}'. Call again with confirm=true to proceed.",
                    warning = "Other measures referencing this one will break! Use get_dependencies to check first."
                });
            }
            
            await _measureService.DeleteMeasureAsync(measureName);

            var result = new
            {
                success = true,
                message = $"Successfully deleted measure '{measureName}'. Remember to call save_workbook to persist changes."
            };
            return JsonSerializer.Serialize(result);
        }
        catch (TimeoutException ex)
        {
            return JsonSerializer.Serialize(new { error = $"Operation timed out: {ex.Message}", timeout = true });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { error = ex.Message });
        }
    }
}
