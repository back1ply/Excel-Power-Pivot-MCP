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
/// Unit tests for TableService.
/// Tests table operations with mocked infrastructure dependencies.
/// </summary>
public class TableServiceTests : ServiceTestBase
{
    private TableService CreateService()
    {
        return new TableService(MockTableComService.Object, MockDmvService.Object, MockExcelService.Object);
    }

    #region GetTablesAsync Tests

    [Fact]
    public async Task GetTablesAsync_HasTables_ReturnsAllTables()
    {
        // Arrange
        var expectedTables = new List<TableMetadata>
        {
            TestDataFactory.CreateTableMetadata("Table1", 1000),
            TestDataFactory.CreateTableMetadata("Table2", 2000)
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedTables);

        // Act
        var result = await service.GetTablesAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedTables);
        MockDmvService.Verify(s => s.GetTablesAsync(It.IsAny<CancellationToken>()), Times.Once);
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

    [Fact]
    public async Task GetTablesAsync_DmvServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetTablesAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("DMV query failed"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.GetTablesAsync());
    }

    #endregion

    #region GetExcelTablesAsync Tests

    [Fact]
    public async Task GetExcelTablesAsync_HasExcelTables_ReturnsAllExcelTables()
    {
        // Arrange
        var expectedTables = new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "ExcelTable1", ["Range"] = "A1:C10" },
            new() { ["Name"] = "ExcelTable2", ["Range"] = "E1:G20" }
        };
        var service = CreateService();

        MockExcelService
            .Setup(s => s.GetExcelTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedTables);

        // Act
        var result = await service.GetExcelTablesAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedTables);
        MockExcelService.Verify(s => s.GetExcelTablesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task GetExcelTablesAsync_NoExcelTables_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockExcelService
            .Setup(s => s.GetExcelTablesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<Dictionary<string, object?>>());

        // Act
        var result = await service.GetExcelTablesAsync();

        // Assert
        result.Should().BeEmpty();
    }

    [Fact]
    public async Task GetExcelTablesAsync_ExcelServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockExcelService
            .Setup(s => s.GetExcelTablesAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Excel COM error"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.GetExcelTablesAsync());
    }

    #endregion

    #region AddTableToModelAsync Tests

    [Fact]
    public async Task AddTableToModelAsync_ValidTable_CallsComServiceOnce()
    {
        // Arrange
        var tableDef = TestDataFactory.CreateTableAddToModel("NewTable");
        var service = CreateService();

        MockTableComService
            .Setup(s => s.AddTableToModelAsync(It.IsAny<TableAddToModel>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync("NewTable");

        // Act
        await service.AddTableToModelAsync(tableDef);

        // Assert
        MockTableComService.Verify(
            s => s.AddTableToModelAsync(tableDef, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task AddTableToModelAsync_WithTableDefinition_PassesDefinitionToComService()
    {
        // Arrange
        var tableDef = TestDataFactory.CreateTableAddToModel("CustomTable");
        var service = CreateService();

        TableAddToModel? capturedArg = null;
        MockTableComService
            .Setup(s => s.AddTableToModelAsync(It.IsAny<TableAddToModel>(), It.IsAny<CancellationToken>()))
            .Callback<TableAddToModel, CancellationToken>((t, ct) => capturedArg = t)
            .ReturnsAsync("CustomTable");

        // Act
        await service.AddTableToModelAsync(tableDef);

        // Assert
        capturedArg.Should().NotBeNull();
        capturedArg!.TableName.Should().Be("CustomTable");
    }

    [Fact]
    public async Task AddTableToModelAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var tableDef = TestDataFactory.CreateTableAddToModel();
        var service = CreateService();

        MockTableComService
            .Setup(s => s.AddTableToModelAsync(It.IsAny<TableAddToModel>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Table already exists"));

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.AddTableToModelAsync(tableDef));

        exception.Message.Should().Contain("already exists");
    }

    #endregion

    #region DeleteTableAsync Tests

    [Fact]
    public async Task DeleteTableAsync_ValidTableName_CallsComServiceOnce()
    {
        // Arrange
        var tableName = "TableToDelete";
        var service = CreateService();

        MockTableComService
            .Setup(s => s.DeleteTableAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteTableAsync(tableName);

        // Assert
        MockTableComService.Verify(
            s => s.DeleteTableAsync(tableName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task DeleteTableAsync_TableWithSpaces_PassesCorrectNameToComService()
    {
        // Arrange
        var tableName = "Table With Spaces";
        var service = CreateService();

        string? capturedTableName = null;
        MockTableComService
            .Setup(s => s.DeleteTableAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Callback<string, CancellationToken>((name, ct) => capturedTableName = name)
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteTableAsync(tableName);

        // Assert
        capturedTableName.Should().Be("Table With Spaces");
    }

    [Fact]
    public async Task DeleteTableAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var tableName = "NonExistentTable";
        var service = CreateService();

        MockTableComService
            .Setup(s => s.DeleteTableAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new KeyNotFoundException("Table not found"));

        // Act & Assert
        await Assert.ThrowsAsync<KeyNotFoundException>(() => service.DeleteTableAsync(tableName));
    }

    #endregion

    #region GetPowerQueriesAsync Tests

    [Fact]
    public async Task GetPowerQueriesAsync_HasQueries_ReturnsAllQueries()
    {
        // Arrange
        var expectedQueries = new List<Dictionary<string, object?>>
        {
            new() { ["Name"] = "Query1", ["Formula"] = "let Source = ..." },
            new() { ["Name"] = "Query2", ["Formula"] = "let Source = ..." }
        };
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetPowerQueriesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedQueries);

        // Act
        var result = await service.GetPowerQueriesAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedQueries);
        MockTableComService.Verify(s => s.GetPowerQueriesAsync(It.IsAny<CancellationToken>()), Times.Once);
    }

    [Fact]
    public async Task GetPowerQueriesAsync_NoQueries_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetPowerQueriesAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<Dictionary<string, object?>>());

        // Act
        var result = await service.GetPowerQueriesAsync();

        // Assert
        result.Should().BeEmpty();
    }

    [Fact]
    public async Task GetPowerQueriesAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetPowerQueriesAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Failed to get Power Queries"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(() => service.GetPowerQueriesAsync());
    }

    #endregion
}
