using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Moq;
using Xunit;
using FluentAssertions;
using ExcelPowerPivotMcp.Core.Services;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Tests.Fixtures;

namespace ExcelPowerPivotMcp.Tests.Core.Services;

/// <summary>
/// Unit tests for ModelMetadataService.
/// Tests model metadata aggregation with mocked infrastructure dependencies.
/// </summary>
public class ModelMetadataServiceTests : ServiceTestBase
{
    private ModelMetadataService CreateService()
    {
        return new ModelMetadataService(
            MockDmvService.Object,
            MockRelationshipComService.Object,
            MockExcelService.Object,
            MockDaxService.Object,
            MockModelMetadataLogger.Object);
    }

    #region GetTablesAsync Tests

    [Fact]
    public async Task GetTablesAsync_HasTables_ReturnsDmvServiceResults()
    {
        // Arrange
        var expectedTables = new List<TableMetadata>
        {
            TestDataFactory.CreateTableMetadata("Table1", 100),
            TestDataFactory.CreateTableMetadata("Table2", 200)
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedTables);

        // Act
        var result = await service.GetTablesAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedTables);
    }

    [Fact]
    public async Task GetTablesAsync_NoTables_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<TableMetadata>());

        // Act
        var result = await service.GetTablesAsync();

        // Assert
        result.Should().BeEmpty();
    }

    #endregion

    #region GetColumnsAsync Tests

    [Fact]
    public async Task GetColumnsAsync_ValidTableName_ReturnsDmvServiceResults()
    {
        // Arrange
        var tableName = "Sales";
        var expectedColumns = new List<ColumnMetadata>
        {
            TestDataFactory.CreateColumnMetadata("ProductKey", tableName),
            TestDataFactory.CreateColumnMetadata("Amount", tableName)
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetColumnsAsync(tableName, It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedColumns);

        // Act
        var result = await service.GetColumnsAsync(tableName);

        // Assert
        result.Should().BeEquivalentTo(expectedColumns);
    }

    [Fact]
    public async Task GetColumnsAsync_PassesTableNameToDmvService()
    {
        // Arrange
        var tableName = "Products";
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetColumnsAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<ColumnMetadata>());

        // Act
        await service.GetColumnsAsync(tableName);

        // Assert
        MockDmvService.Verify(
            s => s.GetColumnsAsync(tableName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    #endregion

    #region GetMeasuresAsync Tests

    [Fact]
    public async Task GetMeasuresAsync_NoTableName_ReturnsAllMeasures()
    {
        // Arrange
        var expectedMeasures = new List<MeasureMetadata>
        {
            TestDataFactory.CreateMeasureMetadata("Measure1", "Table1"),
            TestDataFactory.CreateMeasureMetadata("Measure2", "Table2")
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetMeasuresAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedMeasures);

        // Act
        var result = await service.GetMeasuresAsync(null);

        // Assert
        result.Should().BeEquivalentTo(expectedMeasures);
    }

    [Fact]
    public async Task GetMeasuresAsync_WithTableName_PassesTableNameToDmvService()
    {
        // Arrange
        var tableName = "Sales";
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetMeasuresAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<MeasureMetadata>());

        // Act
        await service.GetMeasuresAsync(tableName);

        // Assert
        MockDmvService.Verify(
            s => s.GetMeasuresAsync(tableName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    #endregion

    #region GetKpisAsync Tests

    [Fact]
    public async Task GetKpisAsync_ExecutesDmvQuery()
    {
        // Arrange
        var expectedKpis = new List<Dictionary<string, object?>>
        {
            new() { ["KPI_NAME"] = "Revenue", ["KPI_VALUE"] = "SUM([Amount])" }
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.ExecuteDmvAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedKpis);

        // Act
        var result = await service.GetKpisAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedKpis);
        MockDmvService.Verify(
            s => s.ExecuteDmvAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetKpisAsync_ReturnsKpiData()
    {
        // Arrange
        var expectedKpis = new List<Dictionary<string, object?>>
        {
            new() { ["KPI_NAME"] = "SalesTarget", ["KPI_STATUS"] = "Good" }
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.ExecuteDmvAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedKpis);

        // Act
        var result = await service.GetKpisAsync();

        // Assert
        result.Should().HaveCount(1);
        result[0]["KPI_NAME"].Should().Be("SalesTarget");
    }

    #endregion

    #region GetRelationshipsAsync Tests

    [Fact]
    public async Task GetRelationshipsAsync_ReturnsComServiceResults()
    {
        // Arrange
        var expectedRelationships = new List<RelationshipMetadata>
        {
            TestDataFactory.CreateRelationshipMetadata("Sales", "ProductKey", "Product", "ProductKey", true)
        };
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedRelationships);

        // Act
        var result = await service.GetRelationshipsAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedRelationships);
    }

    [Fact]
    public async Task GetRelationshipsAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Failed to get relationships"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.GetRelationshipsAsync());
    }

    #endregion

    #region GetModelSummaryAsync Tests

    [Fact]
    public async Task GetModelSummaryAsync_SuccessfulQuery_ReturnsCompleteModel()
    {
        // Arrange
        var workbookPath = "C:\\Test\\Workbook.xlsx";
        var tables = new List<TableMetadata>
        {
            TestDataFactory.CreateTableMetadata("Sales", 1000),
            TestDataFactory.CreateTableMetadata("Products", 500)
        };
        var measures = new List<MeasureMetadata>
        {
            TestDataFactory.CreateMeasureMetadata("TotalSales", "Sales")
        };
        var relationships = new List<RelationshipMetadata>
        {
            TestDataFactory.CreateRelationshipMetadata("Sales", "ProductKey", "Products", "ProductKey", true)
        };
        var service = CreateService();

        MockExcelService
            .Setup(s => s.GetConnectedWorkbookPath())
            .Returns(workbookPath);
        MockDmvService
            .Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(tables);
        MockDmvService
            .Setup(s => s.GetMeasuresAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(measures);
        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(relationships);

        // Act
        var result = await service.GetModelSummaryAsync();

        // Assert
        result.ConnectedWorkbook.Should().Be(workbookPath);
        result.TableCount.Should().Be(2);
        result.MeasureCount.Should().Be(1);
        result.RelationshipCount.Should().Be(1);
    }

    [Fact]
    public async Task GetModelSummaryAsync_IncludesConnectedWorkbookPath()
    {
        // Arrange
        var workbookPath = "C:\\Excel\\Model.xlsx";
        var service = CreateService();

        MockExcelService
            .Setup(s => s.GetConnectedWorkbookPath())
            .Returns(workbookPath);
        MockDmvService
            .Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<TableMetadata>());
        MockDmvService
            .Setup(s => s.GetMeasuresAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<MeasureMetadata>());
        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<RelationshipMetadata>());

        // Act
        var result = await service.GetModelSummaryAsync();

        // Assert
        result.ConnectedWorkbook.Should().Be(workbookPath);
    }

    [Fact]
    public async Task GetModelSummaryAsync_SimplifiesRelationshipsForSummary()
    {
        // Arrange
        var relationships = new List<RelationshipMetadata>
        {
            TestDataFactory.CreateRelationshipMetadata("Sales", "ProductKey", "Products", "ProductKey", true)
        };
        var service = CreateService();

        MockExcelService.Setup(s => s.GetConnectedWorkbookPath()).Returns("test.xlsx");
        MockDmvService.Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<TableMetadata>());
        MockDmvService.Setup(s => s.GetMeasuresAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<MeasureMetadata>());
        MockRelationshipComService.Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(relationships);

        // Act
        var result = await service.GetModelSummaryAsync();

        // Assert
        result.Relationships.Should().HaveCount(1);
        var rel = result.Relationships[0];
        rel["from"].Should().Be("Sales.ProductKey");
        rel["to"].Should().Be("Products.ProductKey");
    }

    #endregion
}
