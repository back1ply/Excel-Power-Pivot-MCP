using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.PowerPivot;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Async implementation of Power Pivot relationship operations via COM.
/// All COM operations are:
/// - Dispatched to the STA thread via IExcelDispatcher
/// - Wrapped in ComScope for proper COM object cleanup
/// - Protected by ExecuteWithConnectionValidation for stale-connection detection
/// </summary>
public class RelationshipComService : IRelationshipComService
{
    private readonly IPowerPivotConnection _connection;
    private readonly IExcelDispatcher _dispatcher;
    private readonly McpConfiguration _config;
    private readonly ILogger<RelationshipComService> _logger;

    // Default timeout for COM operations (30 seconds)
    private static TimeSpan ComOperationTimeout => TimeSpan.FromSeconds(30);

    public RelationshipComService(IPowerPivotConnection connection, IExcelDispatcher dispatcher, McpConfiguration config, ILogger<RelationshipComService> logger)
    {
        _connection = connection;
        _dispatcher = dispatcher;
        _config = config;
        _logger = logger;
    }

    public async Task CreateRelationshipAsync(RelationshipCreate def, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                // Use TrackProperty for defensive COM access
                dynamic model = scope.TrackProperty(() => workbook.Model);
                dynamic modelTables = scope.TrackProperty(() => model.ModelTables);
                dynamic modelRelationships = scope.TrackProperty(() => model.ModelRelationships);

                // Check if relationship already exists
                bool exists = false;
                ExcelComHelpers.EnumerateCollectionWithCleanup(modelRelationships, (Action<dynamic>)(rel =>
                {
                    if ((string)rel.ForeignKeyTable.Name == def.ForeignTable &&
                        (string)rel.ForeignKeyColumn.Name == def.ForeignColumn &&
                        (string)rel.PrimaryKeyTable.Name == def.PrimaryTable &&
                        (string)rel.PrimaryKeyColumn.Name == def.PrimaryColumn)
                    {
                        exists = true;
                    }
                }));

                if (exists)
                {
                    throw new ArgumentException(
                        $"Relationship already exists: {def.ForeignTable}.{def.ForeignColumn} → {def.PrimaryTable}.{def.PrimaryColumn}. " +
                        "Use set_relationship_active to modify it, or delete_relationship first.");
                }

                dynamic fkTable = ExcelComHelpers.FindByNameOrThrow(modelTables, def.ForeignTable, "Foreign key table");
                scope.TrackDynamic(fkTable);
                dynamic pkTable = ExcelComHelpers.FindByNameOrThrow(modelTables, def.PrimaryTable, "Primary key table");
                scope.TrackDynamic(pkTable);

                dynamic fkTableColumns = scope.TrackProperty(() => fkTable.ModelTableColumns);
                dynamic pkTableColumns = scope.TrackProperty(() => pkTable.ModelTableColumns);

                dynamic fkCol = ExcelComHelpers.FindByNameOrThrow(fkTableColumns, def.ForeignColumn, $"FK Column '{def.ForeignTable}'");
                scope.TrackDynamic(fkCol);
                dynamic pkCol = ExcelComHelpers.FindByNameOrThrow(pkTableColumns, def.PrimaryColumn, $"PK Column '{def.PrimaryTable}'");
                scope.TrackDynamic(pkCol);

                try
                {
                    modelRelationships.Add(fkCol, pkCol);
                    _connection.MarkDirty();
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    string hint = ParseRelationshipError(ex, def);
                    throw new InvalidOperationException(hint, ex);
                }
            }, "creating relationship");
        }, ct);
    }

    private static string ParseRelationshipError(System.Runtime.InteropServices.COMException ex, RelationshipCreate def)
    {
        string errorDetails = ex.Message.ToLowerInvariant();

        if (errorDetails.Contains("storageexception") || errorDetails.Contains("couldn't get data"))
        {
            return $"Failed to create relationship: {def.ForeignTable}.{def.ForeignColumn} → {def.PrimaryTable}.{def.PrimaryColumn}.\n\n" +
                   "Common causes:\n" +
                   $"1. The primary key column '{def.PrimaryColumn}' in '{def.PrimaryTable}' may contain duplicate values. " +
                   "Use analyze_column to check for uniqueness.\n" +
                   "2. Data types may be incompatible between the two columns. Use list_columns to verify.\n" +
                   "3. A relationship path already exists between these tables, creating an ambiguous filter path. " +
                   "Use list_relationships to see existing relationships.";
        }
        if (errorDetails.Contains("type") || errorDetails.Contains("mismatch"))
        {
            return $"Data type mismatch: Columns '{def.ForeignColumn}' and '{def.PrimaryColumn}' must have compatible data types.";
        }
        return $"Failed to create relationship: {ex.Message}";
    }

    public async Task DeleteRelationshipAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelRelationships = scope.TrackDynamic(model.ModelRelationships);

                // FIXED: Find-then-delete pattern to avoid collection mutation during enumeration
                // Using reverse iteration for safety (though we break after finding)
                dynamic? targetRel = null;
                int count = modelRelationships.Count;

                for (int i = count; i >= 1; i--)
                {
                    dynamic rel = modelRelationships.Item[i];
                    try
                    {
                        if ((string)rel.ForeignKeyTable.Name == foreignTable &&
                            (string)rel.ForeignKeyColumn.Name == foreignColumn &&
                            (string)rel.PrimaryKeyTable.Name == primaryTable &&
                            (string)rel.PrimaryKeyColumn.Name == primaryColumn)
                        {
                            targetRel = rel;
                            scope.TrackDynamic(targetRel);
                            break;
                        }
                    }
                    finally
                    {
                        if (targetRel == null || !ReferenceEquals(targetRel, rel))
                        {
                            ComHelper.SafeRelease(rel);
                        }
                    }
                }

                if (targetRel == null)
                {
                    throw new ArgumentException("Relationship not found. Use list_relationships to see available relationships.");
                }

                targetRel.Delete();
                _connection.MarkDirty();

            }, "deleting relationship");
        }, ct);
    }


    public async Task SetRelationshipActiveAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, bool active, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelRelationships = scope.TrackDynamic(model.ModelRelationships);

                bool found = false;
                ExcelComHelpers.EnumerateCollectionWithCleanup(modelRelationships, (Action<dynamic>)(rel =>
                {
                    if (found) return;

                    if ((string)rel.ForeignKeyTable.Name == foreignTable &&
                        (string)rel.ForeignKeyColumn.Name == foreignColumn &&
                        (string)rel.PrimaryKeyTable.Name == primaryTable &&
                        (string)rel.PrimaryKeyColumn.Name == primaryColumn)
                    {
                        rel.Active = active;
                        found = true;
                    }
                }));

                if (!found)
                {
                    throw new ArgumentException("Relationship not found. Use list_relationships to see available relationships.");
                }
                _connection.MarkDirty();
            }, "setting relationship active status");
        }, ct);
    }

    public async Task<List<RelationshipMetadata>> GetRelationshipsAsync(CancellationToken ct = default)
    {
        return await _dispatcher.RunAsync(() =>
        {
            return _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelRelationships = scope.TrackDynamic(model.ModelRelationships);

                return ExcelComHelpers.EnumerateCollectionWithCleanup(modelRelationships, (Func<dynamic, RelationshipMetadata>)(rel =>
                {
                    return new RelationshipMetadata
                    {
                        ForeignKeyTable = (string)rel.ForeignKeyTable.Name,
                        ForeignKeyColumn = (string)rel.ForeignKeyColumn.Name,
                        PrimaryKeyTable = (string)rel.PrimaryKeyTable.Name,
                        PrimaryKeyColumn = (string)rel.PrimaryKeyColumn.Name,
                        IsActive = rel.Active
                    };
                }));
            }, "retrieving relationships");
        }, ct);
    }
}
