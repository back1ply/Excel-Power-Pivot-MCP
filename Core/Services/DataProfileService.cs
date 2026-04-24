using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Infrastructure.Services;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Core.Services;

/// <summary>
/// Async implementation of data profiling operations.
/// </summary>
public class DataProfileService : IDataProfileService
{
    private readonly IDaxService _daxService;
    private readonly ITableComService _tableComService;
    private readonly IDmvService _dmvService;
    private readonly ILogger<DataProfileService> _logger;

    public DataProfileService(IDaxService daxService, ITableComService tableComService, IDmvService dmvService, ILogger<DataProfileService> logger)
    {
        _daxService = daxService;
        _tableComService = tableComService;
        _dmvService = dmvService;
        _logger = logger;
    }

    public async Task<List<Dictionary<string, object?>>> PreviewTableAsync(string tableName, int maxRows, CancellationToken ct = default)
    {
        return await _daxService.PreviewTableAsync(tableName, maxRows, ct);
    }

    public async Task<ColumnProfile> GetColumnProfileAsync(string tableName, string columnName, CancellationToken ct = default)
    {
        return await _daxService.GetColumnProfileAsync(tableName, columnName, ct);
    }

    public async Task<List<ColumnMetadata>> GetCalculatedColumnsAsync(string? tableName, CancellationToken ct = default)
    {
        try
        {
            // COM service is more reliable for calculated columns in Power Pivot
            return await _tableComService.GetCalculatedColumnsAsync(tableName, ct);
        }
        catch (Exception ex)
        {
            // Graceful degradation: return empty list if calculated columns unavailable
            // This can happen if the model has no calculated columns or COM interop fails
            _logger.LogWarning(ex, "Failed to get calculated columns");
            return new List<ColumnMetadata>();
        }
    }
}
