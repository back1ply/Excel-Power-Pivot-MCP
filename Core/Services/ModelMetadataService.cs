using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Common.Utils;
using ExcelPowerPivotMcp.Infrastructure.Services;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Async implementation of model metadata queries.
/// </summary>
public class ModelMetadataService : IModelMetadataService
{
    private readonly IDmvService _dmvService;
    private readonly IRelationshipComService _relationshipComService;
    private readonly IExcelInteropService _excelService;
    private readonly IDaxService _daxService;
    private readonly ILogger<ModelMetadataService> _logger;

    public ModelMetadataService(IDmvService dmvService, IRelationshipComService relationshipComService, IExcelInteropService excelService, IDaxService daxService, ILogger<ModelMetadataService> logger)
    {
        _dmvService = dmvService;
        _relationshipComService = relationshipComService;
        _excelService = excelService;
        _daxService = daxService;
        _logger = logger;
    }

    public async Task<List<TableMetadata>> GetTablesAsync(CancellationToken ct = default) 
        => await _dmvService.GetTablesAsync(ct);
    
    public async Task<List<ColumnMetadata>> GetColumnsAsync(string tableName, CancellationToken ct = default) 
        => await _dmvService.GetColumnsAsync(tableName, ct);
    
    public async Task<List<MeasureMetadata>> GetMeasuresAsync(string? tableName, CancellationToken ct = default) 
        => await _dmvService.GetMeasuresAsync(tableName, ct);
    
    public async Task<List<Dictionary<string, object?>>> GetKpisAsync(CancellationToken ct = default)
    {
        var query = @"
            SELECT 
                [MEASUREGROUP_NAME],
                [KPI_NAME],
                [KPI_CAPTION],
                [KPI_VALUE],
                [KPI_GOAL],
                [KPI_STATUS],
                [KPI_TREND],
                [KPI_DESCRIPTION]
            FROM $SYSTEM.MDSCHEMA_KPIS
            WHERE [CUBE_NAME] = 'Model'
        ";
        return await _dmvService.ExecuteDmvAsync(query, ct);
    }

    public async Task<ModelSummary> GetModelSummaryAsync(CancellationToken ct = default)
    {
        var summary = new ModelSummary
        {
            ConnectedWorkbook = _excelService.GetConnectedWorkbookPath()
        };

        try
        {
            summary.Tables = await GetTablesAsync(ct);
            summary.TableCount = summary.Tables.Count;

            summary.Measures = await GetMeasuresAsync(null, ct);
            summary.MeasureCount = summary.Measures.Count;

            // Relationships - COM API is required for column-level FK/PK info
            // (TMSCHEMA not available in Power Pivot, MDSCHEMA only has table-level cardinality)
            var rels = await _relationshipComService.GetRelationshipsAsync(ct);
            
            summary.RelationshipCount = rels.Count;
            // Simplify for summary output
            summary.Relationships = rels.Select(r => new Dictionary<string, string> {
                {"from", $"{r.ForeignKeyTable}.{r.ForeignKeyColumn}"},
                {"to", $"{r.PrimaryKeyTable}.{r.PrimaryKeyColumn}"}
            }).ToList();

            // Row Counts
            summary.TableRowCounts = await GetAllTableRowCountsAsync(summary.Tables, ct);
            summary.TotalRows = summary.TableRowCounts
                .Where(t => t["rowCount"] is long or int)
                .Sum(t => Convert.ToInt64(t["rowCount"], System.Globalization.CultureInfo.InvariantCulture));
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Model summary failed");
        }

        return summary;
    }

    public async Task<List<RelationshipMetadata>> GetRelationshipsAsync(CancellationToken ct = default)
    {
        // COM API is required for column-level relationship info in Power Pivot
        // TMSCHEMA requires compat level 1200+ (not available in Excel)
        // MDSCHEMA_MEASUREGROUP_DIMENSIONS only provides table-level cardinality
        return await _relationshipComService.GetRelationshipsAsync(ct);
    }

    private async Task<List<Dictionary<string, object?>>> GetAllTableRowCountsAsync(List<TableMetadata> tables, CancellationToken ct = default)
    {
        var results = new List<Dictionary<string, object?>>();
        var tableNames = tables
            .Select(t => t.Name)
            .Where(n => !string.IsNullOrEmpty(n) && !n.StartsWith('$'))
            .ToList();
        
        if (tableNames.Count == 0) return results;
        
        // Build a single UNION query for all tables
        var rowExpressions = tableNames
            .Select(name => $"ROW(\"Table\", \"{DaxHelpers.SanitizeValue(name)}\", \"RowCount\", COUNTROWS('{DaxHelpers.SanitizeValue(name)}'))");
        
        var batchQuery = $"EVALUATE UNION({string.Join(", ", rowExpressions)})";
        
        try
        {
            var queryResults = await _daxService.ExecuteDaxAsync(batchQuery, 1000, ct);
            foreach (var row in queryResults)
            {
                if (row.TryGetValue("Table", out var tName) && row.TryGetValue("RowCount", out var rc))
                {
                    results.Add(new Dictionary<string, object?>
                    {
                        ["table"] = tName,
                        ["rowCount"] = rc
                    });
                }
            }
        }
        catch (Exception ex)
        {
            // Fall back to individual queries
            _logger.LogWarning(ex, "Batch row count failed, falling back to individual queries");
            foreach (var tableName in tableNames)
            {
                try
                {
                    var query = $"EVALUATE ROW(\"RowCount\", COUNTROWS('{DaxHelpers.SanitizeValue(tableName)}'))";
                    var result = await _daxService.ExecuteDaxAsync(query, 1, ct);
                    if (result.Count > 0 && result[0].TryGetValue("RowCount", out var rc))
                    {
                        results.Add(new Dictionary<string, object?>
                        {
                            ["table"] = tableName,
                            ["rowCount"] = rc
                        });
                    }
                }
                catch { /* Skip tables that fail */ }
            }
        }
        
        return results;
    }
}
