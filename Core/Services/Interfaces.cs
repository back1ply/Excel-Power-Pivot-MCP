using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Core service for measure CRUD operations.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IMeasureService
{
    Task CreateMeasureAsync(MeasureCreate def, CancellationToken ct = default);
    Task UpdateMeasureAsync(MeasureUpdate def, CancellationToken ct = default);
    Task DeleteMeasureAsync(string measureName, CancellationToken ct = default);
    Task<List<MeasureMetadata>> GetMeasuresAsync(string? tableName = null, CancellationToken ct = default);
}

/// <summary>
/// Core service for relationship operations.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IRelationshipService
{
    Task CreateRelationshipAsync(RelationshipCreate def, CancellationToken ct = default);
    Task DeleteRelationshipAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, CancellationToken ct = default);
    Task<List<RelationshipMetadata>> GetRelationshipsAsync(CancellationToken ct = default);
    Task SetRelationshipActiveAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, bool active, CancellationToken ct = default);
}

/// <summary>
/// Core service for table operations.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface ITableService
{
    Task<List<TableMetadata>> GetTablesAsync(CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> GetExcelTablesAsync(CancellationToken ct = default);
    Task<string> AddTableToModelAsync(TableAddToModel def, CancellationToken ct = default);
    Task DeleteTableAsync(string tableName, CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> GetPowerQueriesAsync(CancellationToken ct = default);
}

// NOTE: ICalculatedColumnService removed - Excel COM API doesn't support creating calculated columns
// Calculated columns can only be created through the Power Pivot window UI

/// <summary>
/// Core service for model metadata queries.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IModelMetadataService
{
    Task<List<TableMetadata>> GetTablesAsync(CancellationToken ct = default);
    Task<List<ColumnMetadata>> GetColumnsAsync(string tableName, CancellationToken ct = default);
    Task<List<MeasureMetadata>> GetMeasuresAsync(string? tableName, CancellationToken ct = default);
    Task<List<Dictionary<string, object?>>> GetKpisAsync(CancellationToken ct = default);
    Task<ModelSummary> GetModelSummaryAsync(CancellationToken ct = default);
    Task<List<RelationshipMetadata>> GetRelationshipsAsync(CancellationToken ct = default);
}

/// <summary>
/// Core service for data profiling.
/// All operations are async to avoid blocking on the STA thread.
/// </summary>
public interface IDataProfileService
{
    Task<List<Dictionary<string, object?>>> PreviewTableAsync(string tableName, int maxRows, CancellationToken ct = default);
    Task<ColumnProfile> GetColumnProfileAsync(string tableName, string columnName, CancellationToken ct = default);
    Task<List<ColumnMetadata>> GetCalculatedColumnsAsync(string? tableName, CancellationToken ct = default);
}

/// <summary>
/// Core service for DAX expression formatting.
/// Provides resilient formatting with retry logic, caching, and graceful degradation.
/// </summary>
public interface IDaxFormatterService
{
    /// <summary>
    /// Format a DAX expression for readability.
    /// </summary>
    /// <param name="expression">The DAX expression to format</param>
    /// <param name="ct">Cancellation token</param>
    /// <returns>
    /// Tuple containing:
    /// - formatted: The formatted expression (or original if formatting fails)
    /// - warning: Optional warning message if formatting failed
    /// </returns>
    /// <remarks>
    /// This method never throws - it always returns a result with graceful degradation.
    /// If formatting fails for any reason, the original expression is returned with a warning.
    /// DAX syntax errors are thrown as InvalidOperationException for the LLM to fix.
    /// </remarks>
    Task<(string formatted, string? warning)> FormatDaxAsync(string expression, CancellationToken ct = default);

    /// <summary>
    /// Clear the formatter cache.
    /// </summary>
    void ClearCache();

    /// <summary>
    /// Get cache statistics for monitoring.
    /// </summary>
    /// <returns>Dictionary with cache stats (hits, misses, size, etc.)</returns>
    Dictionary<string, object> GetCacheStats();
}
