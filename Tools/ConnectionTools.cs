using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.PowerPivot;
using ExcelPowerPivotMcp.Infrastructure.Services;
using ExcelPowerPivotMcp.Common;

namespace ExcelPowerPivotMcp.Tools;

/// <summary>
/// MCP tools for workbook discovery and connection management.
/// </summary>
[McpServerToolType]
public class ConnectionTools
{
    private readonly IExcelInteropService _excelService;

    public ConnectionTools(IExcelInteropService excelService)
    {
        _excelService = excelService;
    }

    [McpServerTool(Name = "discover_workbooks")]
    [Description("List open Excel workbooks with Power Pivot models.")]
    public async Task<string> DiscoverWorkbooks()
    {
        var workbooks = await _excelService.DiscoverWorkbooksAsync();
        
        if (workbooks.Count == 0)
        {
            throw new InvalidOperationException("No Excel workbooks found. Make sure Excel is running with at least one workbook open.");
        }

        var result = new
        {
            workbooks = workbooks.Select(w => new
            {
                name = w.Name,
                path = w.FullPath,
                hasDataModel = w.HasDataModel
            }).ToList(),
            count = workbooks.Count,
            withDataModel = workbooks.Count(w => w.HasDataModel)
        };
        
        return JsonSerializer.Serialize(result);
    }

    [McpServerTool(Name = "connect_workbook")]
    [Description("Connect to a workbook by path or name.")]
    public async Task<string> ConnectWorkbook(
        [Description("Full path to the Excel workbook (optional if workbook_name provided)")] string? workbookPath = null,
        [Description("Workbook name for lookup (e.g., 'Sales.xlsx'). Supports partial matching.")] string? workbookName = null)
    {
        if (string.IsNullOrEmpty(workbookPath) && string.IsNullOrEmpty(workbookName))
        {
            throw new ArgumentException("Either workbook_path or workbook_name is required");
        }

        string connectedPath;
        
        if (!string.IsNullOrEmpty(workbookPath))
        {
            await _excelService.ConnectAsync(workbookPath!);
            connectedPath = workbookPath!;
        }
        else
        {
            var workbooks = await _excelService.DiscoverWorkbooksAsync();
            
            if (workbooks.Count == 0)
            {
                throw new InvalidOperationException("No Excel workbooks found. Make sure Excel is running.");
            }

            var match = workbooks.FirstOrDefault(w => 
                w.Name.Equals(workbookName, StringComparison.OrdinalIgnoreCase) ||
                w.Name.Contains(workbookName!, StringComparison.OrdinalIgnoreCase));

            if (match == null)
            {
                var names = string.Join(", ", workbooks.Select(w => $"'{w.Name}'"));
                throw new InvalidOperationException($"Workbook '{workbookName}' not found. Available: {names}");
            }

            if (!match.HasDataModel)
            {
                throw new InvalidOperationException($"Workbook '{match.Name}' does not have a Power Pivot data model.");
            }

            await _excelService.ConnectAsync(match.FullPath);
            connectedPath = match.FullPath;
        }

        return JsonSerializer.Serialize(new { success = true, message = $"Connected to Power Pivot model in: {connectedPath}" });
    }

    [McpServerTool(Name = "get_connection_status")]
    [Description("Get current connection status and workbook path.")]
    public string GetConnectionStatus()
    {
        return JsonSerializer.Serialize(new
        {
            connected = _excelService.IsConnected,
            workbook = _excelService.GetConnectedWorkbookPath() ?? "None"
        });
    }

    [McpServerTool(Name = "save_workbook")]
    [Description("Save the workbook to persist changes.")]
    public async Task<string> SaveWorkbook()
    {
        await _excelService.SaveWorkbookAsync();
        
        return JsonSerializer.Serialize(new
        {
            success = true,
            message = $"Successfully saved workbook: {_excelService.GetConnectedWorkbookPath()}"
        });
    }

    [McpServerTool(Name = "refresh_model")]
    [Description("Refresh all tables from their data sources.")]
    public async Task<string> RefreshModel()
    {
        await _excelService.RefreshModelAsync();

        return JsonSerializer.Serialize(new
        {
            success = true,
            message = "Successfully triggered model refresh."
        });
    }
}
