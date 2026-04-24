using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Infrastructure.Services;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Async implementation of table operations.
/// </summary>
public class TableService : ITableService
{
    private readonly ITableComService _comService;
    private readonly IDmvService _dmvService;
    private readonly IExcelInteropService _excelService;

    public TableService(ITableComService comService, IDmvService dmvService, IExcelInteropService excelService)
    {
        _comService = comService;
        _dmvService = dmvService;
        _excelService = excelService;
    }

    public async Task<List<TableMetadata>> GetTablesAsync(CancellationToken ct = default)
    {
        return await _dmvService.GetTablesAsync(ct);
    }

    public async Task<List<Dictionary<string, object?>>> GetExcelTablesAsync(CancellationToken ct = default)
    {
        return await _excelService.GetExcelTablesAsync(ct);
    }
    
    public async Task<string> AddTableToModelAsync(TableAddToModel def, CancellationToken ct = default)
    {
        return await _comService.AddTableToModelAsync(def, ct);
    }

    public async Task DeleteTableAsync(string tableName, CancellationToken ct = default)
    {
        await _comService.DeleteTableAsync(tableName, ct);
    }
    
    public async Task<List<Dictionary<string, object?>>> GetPowerQueriesAsync(CancellationToken ct = default)
    {
        return await _comService.GetPowerQueriesAsync(ct);
    }
}
