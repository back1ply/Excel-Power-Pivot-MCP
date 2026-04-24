using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.PowerPivot;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.Infrastructure.Services;

/// <summary>
/// Async implementation of Power Pivot measure operations via COM.
/// All COM operations are:
/// - Dispatched to the STA thread via IExcelDispatcher
/// - Wrapped in ComScope for proper COM object cleanup
/// - Protected by ExecuteWithConnectionValidation for stale-connection detection
/// </summary>
public class MeasureComService : IMeasureComService
{
    private readonly IPowerPivotConnection _connection;
    private readonly IExcelDispatcher _dispatcher;
    private readonly McpConfiguration _config;
    private readonly ILogger<MeasureComService> _logger;

    // Default timeout for COM operations (30 seconds)
    private static TimeSpan ComOperationTimeout => TimeSpan.FromSeconds(30);

    public MeasureComService(IPowerPivotConnection connection, IExcelDispatcher dispatcher, McpConfiguration config, ILogger<MeasureComService> logger)
    {
        _connection = connection;
        _dispatcher = dispatcher;
        _config = config;
        _logger = logger;
    }

    public async Task CreateMeasureAsync(MeasureCreate def, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                // Use TrackProperty for defensive COM access - prevents leaks on exception paths
                dynamic model = scope.TrackProperty(() => workbook.Model);
                dynamic modelTables = scope.TrackProperty(() => model.ModelTables);
                dynamic modelMeasures = scope.TrackProperty(() => model.ModelMeasures);
                dynamic formatGeneral = scope.TrackProperty(() => model.ModelFormatGeneral);

                dynamic targetTable = ExcelComHelpers.FindByNameOrThrow(modelTables, def.TableName, "Table");
                scope.TrackDynamic(targetTable);

                dynamic newMeasure = scope.TrackProperty(() =>
                    modelMeasures.Add(def.MeasureName, targetTable, def.Expression, formatGeneral));

                if (!string.IsNullOrEmpty(def.Description))
                {
                    newMeasure.Description = def.Description;
                }
                _connection.MarkDirty();
            }, "creating measure");
        }, ComOperationTimeout, ct);
    }

    public async Task UpdateMeasureAsync(MeasureUpdate def, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelMeasures = scope.TrackDynamic(model.ModelMeasures);

                dynamic measure = ExcelComHelpers.FindByNameOrThrow(modelMeasures, def.MeasureName, "Measure");
                scope.TrackDynamic(measure);

                if (!string.IsNullOrEmpty(def.Expression)) measure.Formula = def.Expression;
                if (def.Description != null) measure.Description = def.Description;
                if (!string.IsNullOrEmpty(def.NewName) && def.NewName != def.MeasureName) measure.Name = def.NewName;
                _connection.MarkDirty();
            }, "updating measure");
        }, ComOperationTimeout, ct);
    }

    public async Task DeleteMeasureAsync(string measureName, CancellationToken ct = default)
    {
        await _dispatcher.RunAsync(() =>
        {
            _connection.ExecuteWithConnectionValidation(() =>
            {
                using var scope = new ComScope();

                dynamic workbook = _connection.Workbook!;
                dynamic model = scope.TrackDynamic(workbook.Model);
                dynamic modelMeasures = scope.TrackDynamic(model.ModelMeasures);

                dynamic measure = ExcelComHelpers.FindByNameOrThrow(modelMeasures, measureName, "Measure");
                scope.TrackDynamic(measure);
                measure.Delete();
                _connection.MarkDirty();
            }, "deleting measure");
        }, ComOperationTimeout, ct);
    }
}
