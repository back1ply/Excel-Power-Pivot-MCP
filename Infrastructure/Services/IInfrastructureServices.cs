using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Handles pure Excel Application interactions (Discovery, Connection, File Ops).
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IExcelInteropService
{
    Task<List<PowerPivotWorkbook>> DiscoverWorkbooksAsync(CancellationToken ct = default);
    Task ConnectAsync(string workbookPath, CancellationToken ct = default);
    Task SaveWorkbookAsync(CancellationToken ct = default);
    Task RefreshModelAsync(CancellationToken ct = default);
    string? GetConnectedWorkbookPath();
    bool IsConnected { get; }
    Task<List<Dictionary<string, object?>>> GetExcelTablesAsync(CancellationToken ct = default);
}

/// <summary>
/// Handles Power Pivot measure operations via COM.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IMeasureComService
{
    Task CreateMeasureAsync(MeasureCreate def, CancellationToken ct = default);
    Task UpdateMeasureAsync(MeasureUpdate def, CancellationToken ct = default);
    Task DeleteMeasureAsync(string measureName, CancellationToken ct = default);
}

/// <summary>
/// Handles Power Pivot relationship operations via COM.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IRelationshipComService
{
    Task CreateRelationshipAsync(RelationshipCreate def, CancellationToken ct = default);
    Task DeleteRelationshipAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, CancellationToken ct = default);
    Task SetRelationshipActiveAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, bool active, CancellationToken ct = default);
    Task<List<RelationshipMetadata>> GetRelationshipsAsync(CancellationToken ct = default);
}

/// <summary>
/// Handles Power Pivot table operations via COM.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface ITableComService
{
    Task<string> AddTableToModelAsync(TableAddToModel def, CancellationToken ct = default);
    Task DeleteTableAsync(string tableName, CancellationToken ct = default);
    Task<List<ColumnMetadata>> GetCalculatedColumnsAsync(string? tableName = null, CancellationToken ct = default);
    Task RefreshTableAsync(string tableName, CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> GetPowerQueriesAsync(CancellationToken ct = default);
}

/// <summary>
/// Handles Metadata/Schema queries (DMV - Read Only).
/// All operations are async to avoid blocking on the STA thread.
///
/// Note: Relationships are NOT available via DMV in Power Pivot (requires compat level 1200+ for TMSCHEMA).
/// MDSCHEMA_MEASUREGROUP_DIMENSIONS only provides table-level cardinality, not column-level FK/PK.
/// Use IRelationshipComService.GetRelationshipsAsync() instead.
/// </summary>
public interface IDmvService
{
    Task<List<TableMetadata>> GetTablesAsync(CancellationToken ct = default);
    Task<List<ColumnMetadata>> GetColumnsAsync(string tableName, CancellationToken ct = default);
    Task<List<MeasureMetadata>> GetMeasuresAsync(string? tableName = null, CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> ExecuteDmvAsync(string query, CancellationToken ct = default);

    // DMV methods
    Task<List<Dictionary<string, object?>>> GetDependenciesAsync(string? objectName = null, CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> GetHierarchiesAsync(CancellationToken ct = default);
    Task<Dictionary<string, object?>> GetCatalogInfoAsync(CancellationToken ct = default);
}

/// <summary>
/// Handles Data queries (DAX - Read Only).
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IDaxService
{
    Task<List<Dictionary<string, object?>>> ExecuteDaxAsync(string daxQuery, int maxRows, CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> PreviewTableAsync(string tableName, int maxRows, CancellationToken ct = default);
    Task<ColumnProfile> GetColumnProfileAsync(string tableName, string columnName, CancellationToken ct = default);
}
