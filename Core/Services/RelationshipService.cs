using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Infrastructure.Services;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Async implementation of relationship operations.
/// Uses COM API for all operations since DMV doesn't provide column-level FK/PK info in Power Pivot.
/// </summary>
public class RelationshipService : IRelationshipService
{
    private readonly IRelationshipComService _comService;

    public RelationshipService(IRelationshipComService comService)
    {
        _comService = comService;
    }

    public async Task CreateRelationshipAsync(RelationshipCreate def, CancellationToken ct = default)
    {
        await _comService.CreateRelationshipAsync(def, ct);
    }

    public async Task DeleteRelationshipAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, CancellationToken ct = default)
    {
        await _comService.DeleteRelationshipAsync(foreignTable, foreignColumn, primaryTable, primaryColumn, ct);
    }

    public async Task<List<RelationshipMetadata>> GetRelationshipsAsync(CancellationToken ct = default)
    {
        // COM API is required for column-level relationship info in Power Pivot
        // TMSCHEMA requires compat level 1200+ (not available in Excel)
        // MDSCHEMA_MEASUREGROUP_DIMENSIONS only provides table-level cardinality
        return await _comService.GetRelationshipsAsync(ct);
    }

    public async Task SetRelationshipActiveAsync(string foreignTable, string foreignColumn, string primaryTable, string primaryColumn, bool active, CancellationToken ct = default)
    {
        await _comService.SetRelationshipActiveAsync(foreignTable, foreignColumn, primaryTable, primaryColumn, active, ct);
    }
}
