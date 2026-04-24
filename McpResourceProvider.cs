using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.Core.Services;
using ExcelPowerPivotMcp.Infrastructure.Services;

namespace ExcelPowerPivotMcp;

/// <summary>
/// MCP Resources for schema discovery.
/// Resources expose static/semi-static metadata that LLMs can read before calling tools.
/// 
/// Note: Resources that need to call async services use .GetAwaiter().GetResult()
/// because MCP resources are currently synchronous in the SDK.
/// </summary>
[McpServerResourceType]
public class McpResourceProvider
{
    private readonly IModelMetadataService _metadataService;
    private readonly IExcelInteropService _excelService;
    private readonly IInMemoryLogReader? _logReader; // Optional

    public McpResourceProvider(
        IModelMetadataService metadataService,
        IExcelInteropService excelService,
        IServiceProvider serviceProvider) // Use ServiceProvider for optional dependency
    {
        _metadataService = metadataService;
        _excelService = excelService;
        _logReader = serviceProvider.GetService(typeof(IInMemoryLogReader)) as IInMemoryLogReader;
    }

    [McpServerResource(
        UriTemplate = "model://schema",
        Name = "Model Schema",
        MimeType = "application/json")]
    [Description("Current Power Pivot data model schema (tables, columns, measures, relationships). Read this before writing DAX queries to understand the model structure.")]
    public async Task<string> GetModelSchemaAsync()
    {
        if (!_excelService.IsConnected)
        {
            return JsonSerializer.Serialize(new
            {
                connected = false,
                message = "No workbook connected. Use connect_workbook first.",
                hint = "1. Use discover_workbooks to find available workbooks, 2. Use connect_workbook to connect"
            });
        }

        try
        {
            // Now fully async :)
            var tables = await _metadataService.GetTablesAsync();
            var measures = await _metadataService.GetMeasuresAsync(null);
            var relationships = await _metadataService.GetRelationshipsAsync();

            // Build columns per table
            var tableData = new List<object>();
            foreach (var t in tables)
            {
                var columns = await _metadataService.GetColumnsAsync(t.Name);
                tableData.Add(new
                {
                    name = t.Name,
                    rowCount = t.RowCount,
                    isVisible = t.IsVisible,
                    columns = columns.Select(c => new
                    {
                        name = c.Name,
                        dataType = c.DataType,
                        isCalculated = c.IsCalculated
                    }).ToList()
                });
            }

            var measureData = measures.Select(m => new
            {
                name = m.Name,
                table = m.TableName,
                expression = m.Expression,
                formatString = m.FormatString
            }).ToList();

            var relationshipData = relationships.Select(r => new
            {
                from = $"{r.ForeignKeyTable}.{r.ForeignKeyColumn}",
                to = $"{r.PrimaryKeyTable}.{r.PrimaryKeyColumn}",
                isActive = r.IsActive
            }).ToList();

            return JsonSerializer.Serialize(new
            {
                connected = true,
                workbook = _excelService.GetConnectedWorkbookPath(),
                tableCount = tables.Count,
                measureCount = measures.Count,
                relationshipCount = relationships.Count,
                tables = tableData,
                measures = measureData,
                relationships = relationshipData
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new
            {
                connected = true,
                error = $"Failed to retrieve schema: {ex.Message}",
                hint = "Connection may be stale. Try reconnecting with connect_workbook."
            });
        }
    }

    [McpServerResource(
        UriTemplate = "model://instructions",
        Name = "Power Pivot Instructions",
        MimeType = "text/markdown")]
    [Description("Comprehensive guide for working with Excel Power Pivot data models via MCP. Covers connection workflow, available operations, limitations, and error handling.")]
    public static string GetInstructions()
    {
        return ExcelPowerPivotMcp.Common.EmbeddedResources.ReadResource("excel_powerpivot_instructions.md")
            ?? "# Excel Power Pivot MCP Instructions\n\nNo instructions available.";
    }

    [McpServerResource(
        UriTemplate = "model://dax-guide",
        Name = "DAX Query Guide",
        MimeType = "text/markdown")]
    [Description("DAX query reference for Excel Power Pivot. Includes EVALUATE syntax, DEFINE MEASURE patterns, aggregation functions, and Excel-specific limitations.")]
    public static string GetDaxGuide()
    {
        return ExcelPowerPivotMcp.Common.EmbeddedResources.ReadResource("dax_query_excel_guide.md")
            ?? "# DAX Query Guide\n\nNo guide available.";
    }

    [McpServerResource(
        UriTemplate = "model://best-practices",
        Name = "Measure Best Practices",
        MimeType = "text/markdown")]
    [Description("Best practices for creating and managing Power Pivot measures. Covers naming conventions, performance optimization, error handling patterns, and format strings.")]
    public static string GetBestPractices()
    {
        return ExcelPowerPivotMcp.Common.EmbeddedResources.ReadResource("powerpivot_measure_best_practices.md")
            ?? "# Measure Best Practices\n\nNo best practices available.";
    }

    [McpServerResource(
        UriTemplate = "model://workflows",
        Name = "Common Workflows",
        MimeType = "text/markdown")]
    [Description("Step-by-step workflows for common Power Pivot tasks: exploring models, creating measures, adding tables, building relationships.")]
    public static string GetWorkflows()
    {
        return ExcelPowerPivotMcp.Common.EmbeddedResources.ReadResource("common_workflows.md")
            ?? "# Common Workflows\n\nNo workflows available.";
    }

    [McpServerResource(
        UriTemplate = "model://performance-guide",
        Name = "Performance Optimization Guide",
        MimeType = "text/markdown")]
    [Description("Comprehensive guide for optimizing Power Pivot data models. Covers Vertipaq compression, data type optimization, cardinality management, relationship optimization, and diagnostic workflows using get_storage_statistics and analyze_vertipaq_compression tools.")]
    public static string GetPerformanceGuide()
    {
        return ExcelPowerPivotMcp.Common.EmbeddedResources.ReadResource("performance_optimization_guide.md")
            ?? "# Performance Optimization Guide\n\nNo performance guide available.";
    }

    [McpServerResource(
        UriTemplate = "model://logs",
        Name = "Server Logs",
        MimeType = "text/plain")]
    [Description("Recent server diagnostic logs (last 100 entries). Use this to debug connection issues, failed queries, or unexpected behavior.")]
    public string GetServerLogs()
    {
        // This can be synchronous as it just reads from memory
        if (_logReader == null)
        {
            return "Log reader not available.";
        }
        return _logReader.GetLogs();
    }

    [McpServerResource(
        UriTemplate = "model://dirty-state",
        Name = "Unsaved Changes Status",
        MimeType = "application/json")]
    [Description("Check if the data model has unsaved changes. Returns isDirty=true if save_workbook is needed after mutations like create_measure, delete_relationship, etc.")]
    public static string GetDirtyState()
    {
        var connection = ExcelPowerPivotMcp.PowerPivot.PowerPivotConnection.Instance;
        
        return JsonSerializer.Serialize(new
        {
            isDirty = connection.IsDirty,
            connected = connection.IsConnected,
            workbook = connection.ConnectedWorkbook ?? "Not connected",
            hint = connection.IsDirty 
                ? "Call save_workbook to persist changes." 
                : "No unsaved changes."
        });
    }
}
