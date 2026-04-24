using Moq;
using Microsoft.Extensions.Logging;
using ExcelPowerPivotMcp.Infrastructure.Services;
using ExcelPowerPivotMcp.Core.Services;

namespace ExcelPowerPivotMcp.Tests.Fixtures;

/// <summary>
/// Base class for service unit tests with common mock setup.
/// Provides pre-configured mocks for infrastructure dependencies.
/// </summary>
public abstract class ServiceTestBase
{
    // Decomposed COM service mocks (SRP-compliant)
    protected Mock<IMeasureComService> MockMeasureComService { get; }
    protected Mock<IRelationshipComService> MockRelationshipComService { get; }
    protected Mock<ITableComService> MockTableComService { get; }

    protected Mock<IDmvService> MockDmvService { get; }
    protected Mock<IDaxService> MockDaxService { get; }
    protected Mock<IExcelInteropService> MockExcelService { get; }

    protected Mock<ILogger<DataProfileService>> MockDataProfileLogger { get; }
    protected Mock<ILogger<DaxFormatterService>> MockDaxFormatterLogger { get; }
    protected Mock<ILogger<ModelMetadataService>> MockModelMetadataLogger { get; }

    protected ServiceTestBase()
    {
        // Decomposed COM service mocks
        MockMeasureComService = new Mock<IMeasureComService>();
        MockRelationshipComService = new Mock<IRelationshipComService>();
        MockTableComService = new Mock<ITableComService>();

        MockDmvService = new Mock<IDmvService>();
        MockDaxService = new Mock<IDaxService>();
        MockExcelService = new Mock<IExcelInteropService>();

        MockDataProfileLogger = new Mock<ILogger<DataProfileService>>();
        MockDaxFormatterLogger = new Mock<ILogger<DaxFormatterService>>();
        MockModelMetadataLogger = new Mock<ILogger<ModelMetadataService>>();
    }

    /// <summary>
    /// Reset all mocks between tests to ensure test isolation.
    /// Call this in test cleanup if needed.
    /// </summary>
    protected void ResetMocks()
    {
        MockMeasureComService.Reset();
        MockRelationshipComService.Reset();
        MockTableComService.Reset();
        MockDmvService.Reset();
        MockDaxService.Reset();
        MockExcelService.Reset();
    }
}
