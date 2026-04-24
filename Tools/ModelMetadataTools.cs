using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Server;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Tools;

/// <summary>
/// MCP tools for querying model metadata (tables, columns, measures, relationships).
/// </summary>
[McpServerToolType]
public class ModelMetadataTools
{
    private readonly Core.Services.IModelMetadataService _metadataService;
    private readonly Core.Services.IDataProfileService _profileService;
    private readonly Core.Services.ITableService _tableService;
    private readonly Infrastructure.Services.IDmvService _dmvService;

    public ModelMetadataTools(
        Core.Services.IModelMetadataService metadataService, 
        Core.Services.IDataProfileService profileService,
        Core.Services.ITableService tableService,
        Infrastructure.Services.IDmvService dmvService)
    {
        _metadataService = metadataService;
        _profileService = profileService;
        _tableService = tableService;
        _dmvService = dmvService;
    }

    [McpServerTool(Name = "list_tables")]
    [Description("List all tables in the data model. Options: include_columns, include_stats.")]
    public async Task<string> GetTables(
        [Description("Include column names for each table (default: false)")] bool includeColumns = false,
        [Description("Include row counts for each table (default: false)")] bool includeStats = false)
    {
        var tables = await _metadataService.GetTablesAsync();
        
        if (includeColumns || includeStats)
        {
            var enrichedTables = new List<Dictionary<string, object?>>();
            
            foreach (var table in tables)
            {
                var tableInfo = new Dictionary<string, object?>
                {
                    ["name"] = table.Name,
                    ["caption"] = table.Caption,
                    ["isVisible"] = table.IsVisible
                };
                
                if (includeColumns)
                {
                    var columns = await _metadataService.GetColumnsAsync(table.Name);
                    tableInfo["columns"] = columns.Select(c => c.Name).ToList();
                }
                
                if (includeStats)
                {
                    tableInfo["rowCount"] = table.RowCount;
                }
                
                enrichedTables.Add(tableInfo);
            }
            
            return JsonSerializer.Serialize(new { tableCount = enrichedTables.Count, tables = enrichedTables });
        }

        return JsonSerializer.Serialize(new { tableCount = tables.Count, tables });
    }

    [McpServerTool(Name = "list_columns")]
    [Description("List columns in a table with data types. Options: include_expressions (for calculated column DAX).")]
    public async Task<string> GetColumns(
        [Description("Name of the table to get columns for")] string tableName,
        [Description("Include DAX expressions for calculated columns (default: false)")] bool includeExpressions = false)
    {
        var columns = await _metadataService.GetColumnsAsync(tableName);

        // Add calculated column expressions if requested
        if (includeExpressions)
        {
            var calcColumns = await _profileService.GetCalculatedColumnsAsync(tableName);
            var calcLookup = calcColumns.ToDictionary(
                c => c.Name, 
                c => c.Expression);
            
            foreach (var col in columns)
            {
                if (calcLookup.TryGetValue(col.Name, out var expr))
                {
                    col.Expression = expr;
                    col.IsCalculated = true;
                }
            }
        }

        return JsonSerializer.Serialize(new { tableName, columnCount = columns.Count, columns });
    }

    [McpServerTool(Name = "list_measures")]
    [Description("List measures in the data model. Options: table_name, include_expressions, include_metadata.")]
    public async Task<string> GetMeasures(
        [Description("Filter measures by table name")] string? tableName = null,
        [Description("Include DAX expressions (default: false)")] bool includeExpressions = false,
        [Description("Include visibility and format strings (default: false)")] bool includeMetadata = false)
    {
        var measures = await _metadataService.GetMeasuresAsync(tableName);

        if (!includeExpressions && !includeMetadata)
        {
            // Simplify output
             var simpleMeasures = measures.Select(m => new { m.Name, m.TableName }).ToList();
             return JsonSerializer.Serialize(new { measureCount = simpleMeasures.Count, measures = simpleMeasures });
        }
        
        // Filter properties if needed, but returning full DTO is fine for now
        return JsonSerializer.Serialize(new { measureCount = measures.Count, measures });
    }

    [McpServerTool(Name = "list_relationships")]
    [Description("List relationships between tables (one-to-many, active/inactive).")]
    public async Task<string> GetRelationships()
    {
         // Uses ModelMetadataService which tries DMV first, then COM fallback
         var relationships = await _metadataService.GetRelationshipsAsync();
         return JsonSerializer.Serialize(new { relationshipCount = relationships.Count, relationships });
    }

    [McpServerTool(Name = "get_model_summary")]
    [Description("Get a model overview. Options: detail_level='quick', 'standard', or 'full'.")]
    public async Task<string> GetModelSummary(
        [Description("Level of detail: 'quick', 'standard' (default), or 'full'")] string detailLevel = "standard")
    {
        var level = detailLevel.ToLowerInvariant();
        var tables = await _metadataService.GetTablesAsync();
        var measures = await _metadataService.GetMeasuresAsync(null);
        var relationships = await _metadataService.GetRelationshipsAsync();
        
        // Get catalog info for compatibility level
        var catalogInfo = await _dmvService.GetCatalogInfoAsync();
        
        // Build response based on detail level
        var result = new Dictionary<string, object?>
        {
            ["tableCount"] = tables.Count,
            ["measureCount"] = measures.Count,
            ["relationshipCount"] = relationships.Count,
            ["compatibilityLevel"] = catalogInfo.TryGetValue("CompatibilityLevel", out var cl) ? cl : null,
            ["lastModified"] = catalogInfo.TryGetValue("LastModified", out var lm) ? lm : null
        };
        
        if (level == "quick")
        {
            // Just names
            result["tables"] = tables.Select(t => t.Name).ToList();
            result["measures"] = measures.Select(m => new { m.Name, m.TableName }).ToList();
        }
        else if (level == "full")
        {
            // Everything including expressions and columns with types
            var fullTables = new List<Dictionary<string, object?>>();
            foreach (var table in tables)
            {
                var columns = await _metadataService.GetColumnsAsync(table.Name);
                fullTables.Add(new Dictionary<string, object?>
                {
                    ["name"] = table.Name,
                    ["rowCount"] = table.RowCount,
                    ["columns"] = columns.Select(c => new { c.Name, c.DataType, c.IsCalculated, c.Expression }).ToList()
                });
            }
            result["tables"] = fullTables;
            result["measures"] = measures.Select(m => new { m.Name, m.TableName, m.Expression, m.FormatString }).ToList();
            result["relationships"] = relationships;
        }
        else // "standard" (default)
        {
            // Tables with column names (most useful for authoring)
            var standardTables = new List<Dictionary<string, object?>>();
            foreach (var table in tables)
            {
                var columns = await _metadataService.GetColumnsAsync(table.Name);
                standardTables.Add(new Dictionary<string, object?>
                {
                    ["name"] = table.Name,
                    ["columns"] = columns.Select(c => c.Name).ToList()
                });
            }
            result["tables"] = standardTables;
            result["measures"] = measures.Select(m => new { m.Name, m.TableName }).ToList();
            result["relationships"] = relationships.Select(r => $"{r.ForeignKeyTable}.{r.ForeignKeyColumn} → {r.PrimaryKeyTable}.{r.PrimaryKeyColumn}").ToList();
        }
        
        return JsonSerializer.Serialize(result);
    }

    [McpServerTool(Name = "list_kpis")]
    [Description("List KPIs with their targets and status expressions.")]
    public async Task<string> GetKpis()
    {
        var kpis = await _metadataService.GetKpisAsync();

        if (kpis.Count == 0)
        {
            return JsonSerializer.Serialize(new { kpiCount = 0, kpis = new List<object>(), message = "No KPIs defined in the model" });
        }
        return JsonSerializer.Serialize(new { kpiCount = kpis.Count, kpis });
    }

    [McpServerTool(Name = "analyze_column")]
    [Description("Profile a column: min/max, distinct count, nulls, sample values.")]
    public async Task<string> ProfileColumn(
        [Description("Name of the table")] string tableName,
        [Description("Name of the column to profile")] string columnName)
    {
        var profile = await _profileService.GetColumnProfileAsync(tableName, columnName);
        return JsonSerializer.Serialize(profile);
    }

    [McpServerTool(Name = "list_power_queries")]
    [Description("List Power Query (M) definitions for workbook connections.")]
    public async Task<string> GetPowerQueries()
    {
         var queries = await _tableService.GetPowerQueriesAsync();
         return JsonSerializer.Serialize(new { queryCount = queries.Count, queries });
    }

    [McpServerTool(Name = "list_dependencies")]
    [Description("List dependencies for an object. Use before deletion to check impact.")]
    public async Task<string> GetDependencies(
        [Description("Filter to dependencies involving this object (optional)")] string? objectName = null)
    {
        var deps = await _dmvService.GetDependenciesAsync(objectName);
        
        // Group by object for easier reading
        var grouped = deps
            .GroupBy(d => d["Object"]?.ToString() ?? "Unknown")
            .Select(g => new 
            {
                Object = g.Key,
                Type = g.First()["ObjectType"]?.ToString(),
                Table = g.First()["Table"]?.ToString(),
                References = g.Select(d => new 
                {
                    Name = d["ReferencedObject"]?.ToString(),
                    Type = d["ReferencedObjectType"]?.ToString(),
                    Table = d["ReferencedTable"]?.ToString()
                }).ToList()
            })
            .ToList();
        
        return JsonSerializer.Serialize(new { dependencyCount = deps.Count, objects = grouped });
    }

    [McpServerTool(Name = "list_hierarchies")]
    [Description("List user-defined hierarchies (drill-down paths). Not all models have these.")]
    public async Task<string> GetHierarchies()
    {
        var hierarchies = await _dmvService.GetHierarchiesAsync();

        if (hierarchies.Count == 0)
        {
            return JsonSerializer.Serialize(new {
                hierarchyCount = 0,
                hierarchies = new List<object>(),
                message = "No user-defined hierarchies in this model"
            });
        }

        return JsonSerializer.Serialize(new { hierarchyCount = hierarchies.Count, hierarchies });
    }
}
