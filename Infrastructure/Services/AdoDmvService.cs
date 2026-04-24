using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Common.Utils;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Async implementation of DMV (Data Management Views) queries against the Power Pivot model.
/// All operations are dispatched to the STA thread via IExcelDispatcher.
/// </summary>
public class AdoDmvService : IDmvService
{
    private readonly IPowerPivotConnection _connection;
    private readonly IExcelDispatcher _dispatcher;

    public AdoDmvService(IPowerPivotConnection connection, IExcelDispatcher dispatcher)
    {
        _connection = connection;
        _dispatcher = dispatcher;
    }

    public async Task<List<Dictionary<string, object?>>> ExecuteDmvAsync(string query, CancellationToken ct = default)
    {
        return await _dispatcher.RunAsync(() =>
        {
            return _connection.ExecuteWithConnectionValidation(() =>
            {
                dynamic adoConn = _connection.AdoConnection!;
                dynamic? recordset = null;

                try
                {
                    // For DMVs, we use the same Execute method on ADO connection
                    recordset = adoConn.Execute(query);

                    // Use helper to extract results with OrdinalIgnoreCase comparison (DMV standard)
                    return AdoRecordsetHelper.ExtractResults(recordset, maxRows: 0, useOrdinalIgnoreCase: true);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"DMV Failed: {ex.Message}", ex);
                }
                finally
                {
                    if (recordset != null)
                    {
                        ComHelper.SafeRelease(recordset);
                    }
                }
            }, "executing DMV query");
        }, ct);
    }

    public async Task<List<TableMetadata>> GetTablesAsync(CancellationToken ct = default)
    {
        var query = @"
            SELECT 
                [DIMENSION_UNIQUE_NAME] AS [Name],
                [DIMENSION_CAPTION] AS [Caption],
                [DIMENSION_IS_VISIBLE] AS [IsVisible],
                [DESCRIPTION],
                [DIMENSION_CARDINALITY] AS [RowCount]
            FROM $SYSTEM.MDSCHEMA_DIMENSIONS
            WHERE [DIMENSION_TYPE] <> 2
                AND [CUBE_NAME] = 'Model'
        ";
        var rows = await ExecuteDmvAsync(query, ct);
        
        return rows.Select(r => new TableMetadata
        {
            Name = r["Name"]?.ToString() ?? "Unknown",
            Caption = r["Caption"]?.ToString(),
            IsVisible = Convert.ToBoolean(r["IsVisible"], System.Globalization.CultureInfo.InvariantCulture),
            Description = r["Description"]?.ToString(),
            RowCount = r.TryGetValue("RowCount", out var rc) && rc != null ? Convert.ToInt64(rc, System.Globalization.CultureInfo.InvariantCulture) : null
        }).ToList();
    }

    public async Task<List<ColumnMetadata>> GetColumnsAsync(string tableName, CancellationToken ct = default)
    {
        var sanitized = DaxHelpers.SanitizeTableName(tableName);
        var query = $@"
            SELECT 
                [DIMENSION_UNIQUE_NAME] AS [TableName],
                [LEVEL_NAME] AS [Name],
                [LEVEL_CAPTION] AS [Caption],
                [LEVEL_DBTYPE] AS [DATA_TYPE],
                [DESCRIPTION],
                [LEVEL_IS_VISIBLE]
            FROM $SYSTEM.MDSCHEMA_LEVELS
            WHERE [CUBE_NAME] = 'Model'
                AND [DIMENSION_UNIQUE_NAME] = '{sanitized}'
                AND [LEVEL_NUMBER] = 1
        ";
        var rows = await ExecuteDmvAsync(query, ct);
        
        return rows.Select(r => new ColumnMetadata
        {
            Name = r["Name"]?.ToString() ?? "Unknown",
            TableName = tableName,
            Caption = r["Caption"]?.ToString(),
            DataType = r.TryGetValue("DATA_TYPE", out var dt) ? DaxHelpers.GetDataTypeName(Convert.ToInt32(dt, System.Globalization.CultureInfo.InvariantCulture)) : "Unknown",
            Description = r["Description"]?.ToString(),
            IsHidden = r.TryGetValue("LEVEL_IS_VISIBLE", out var hid) && !Convert.ToBoolean(hid, System.Globalization.CultureInfo.InvariantCulture)
        }).ToList();
    }

    public async Task<List<MeasureMetadata>> GetMeasuresAsync(string? tableName = null, CancellationToken ct = default)
    {
        var whereClause = "[MEASURE_AGGREGATOR] = 0";
        if (!string.IsNullOrEmpty(tableName))
        {
            whereClause += $" AND [MEASUREGROUP_NAME] = '{DaxHelpers.SanitizeValue(tableName!)}'";
        }

        var query = $@"
            SELECT 
                [MEASUREGROUP_NAME],
                [MEASURE_NAME],
                [MEASURE_CAPTION],
                [DATA_TYPE],
                [EXPRESSION],
                [MEASURE_IS_VISIBLE],
                [DEFAULT_FORMAT_STRING],
                [DESCRIPTION]
            FROM $SYSTEM.MDSCHEMA_MEASURES
            WHERE {whereClause}
        ";

        var rows = await ExecuteDmvAsync(query, ct);
        return rows.Select(r => new MeasureMetadata
        {
            Name = r["MEASURE_NAME"]?.ToString() ?? "Unknown",
            TableName = r["MEASUREGROUP_NAME"]?.ToString() ?? "Unknown",
            Caption = r["MEASURE_CAPTION"]?.ToString(),
            DataType = r.TryGetValue("DATA_TYPE", out var dt) ? DaxHelpers.GetDataTypeName(Convert.ToInt32(dt, System.Globalization.CultureInfo.InvariantCulture)) : "Unknown",
            Expression = r["EXPRESSION"]?.ToString(),
            FormatString = r["DEFAULT_FORMAT_STRING"]?.ToString(),
            Description = r["DESCRIPTION"]?.ToString(),
            IsVisible = Convert.ToBoolean(r.TryGetValue("MEASURE_IS_VISIBLE", out var vis) ? vis : true, System.Globalization.CultureInfo.InvariantCulture)
        }).ToList();
    }

    public async Task<List<Dictionary<string, object?>>> GetDependenciesAsync(string? objectName = null, CancellationToken ct = default)
    {
        var query = @"
            SELECT 
                [OBJECT] AS [Object],
                [OBJECT_TYPE] AS [ObjectType],
                [TABLE] AS [Table],
                [REFERENCED_OBJECT] AS [ReferencedObject],
                [REFERENCED_OBJECT_TYPE] AS [ReferencedObjectType],
                [REFERENCED_TABLE] AS [ReferencedTable]
            FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY
        ";
        
        var rows = await ExecuteDmvAsync(query, ct);
        
        // Filter by object name if specified
        if (!string.IsNullOrEmpty(objectName))
        {
            rows = rows.Where(r => 
                string.Equals(r["Object"]?.ToString(), objectName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(r["ReferencedObject"]?.ToString(), objectName, StringComparison.OrdinalIgnoreCase)
            ).ToList();
        }
        
        return rows;
    }

    public async Task<List<Dictionary<string, object?>>> GetHierarchiesAsync(CancellationToken ct = default)
    {
        // HIERARCHY_ORIGIN values (from OLAP_HIERARCHY_ORIGIN enumeration):
        // 1 = MD_USER_DEFINED (user-created multi-level hierarchies)
        // 2 = MD_SYSTEM_ENABLED (attribute hierarchies, auto-generated from columns)
        // We filter for user-defined hierarchies only (HIERARCHY_ORIGIN = 1)
        var query = @"
            SELECT 
                [DIMENSION_UNIQUE_NAME] AS [TableName],
                [HIERARCHY_NAME] AS [Name],
                [HIERARCHY_CAPTION] AS [Caption],
                [HIERARCHY_IS_VISIBLE] AS [IsVisible],
                [DESCRIPTION]
            FROM $SYSTEM.MDSCHEMA_HIERARCHIES
            WHERE [CUBE_NAME] = 'Model'
              AND [HIERARCHY_ORIGIN] = 1
        ";
        
        return await ExecuteDmvAsync(query, ct);
    }

    public async Task<Dictionary<string, object?>> GetCatalogInfoAsync(CancellationToken ct = default)
    {
        var query = @"
            SELECT
                [CATALOG_NAME] AS [Name],
                [COMPATIBILITY_LEVEL] AS [CompatibilityLevel],
                [DATE_MODIFIED] AS [LastModified]
            FROM $SYSTEM.DBSCHEMA_CATALOGS
        ";

        var rows = await ExecuteDmvAsync(query, ct);
        return rows.FirstOrDefault() ?? new Dictionary<string, object?>();
    }
}
