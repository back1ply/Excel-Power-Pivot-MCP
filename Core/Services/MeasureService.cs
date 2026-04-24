using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Infrastructure.Services;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Async implementation of measure CRUD operations.
/// </summary>
public class MeasureService : IMeasureService
{
    private readonly IMeasureComService _comService;
    private readonly IDmvService _dmvService;

    public MeasureService(IMeasureComService comService, IDmvService dmvService)
    {
        _comService = comService;
        _dmvService = dmvService;
    }

    public async Task CreateMeasureAsync(MeasureCreate def, CancellationToken ct = default)
    {
        await _comService.CreateMeasureAsync(def, ct);
    }

    public async Task UpdateMeasureAsync(MeasureUpdate def, CancellationToken ct = default)
    {
        await _comService.UpdateMeasureAsync(def, ct);
    }

    public async Task DeleteMeasureAsync(string measureName, CancellationToken ct = default)
    {
        await _comService.DeleteMeasureAsync(measureName, ct);
    }

    public async Task<List<MeasureMetadata>> GetMeasuresAsync(string? tableName = null, CancellationToken ct = default)
    {
        // Purely metadata read -> Use DMV
        return await _dmvService.GetMeasuresAsync(tableName, ct);
    }
}
