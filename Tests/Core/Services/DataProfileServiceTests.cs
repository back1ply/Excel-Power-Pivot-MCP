using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Moq;
using Xunit;
using FluentAssertions;
using ExcelPowerPivotMcp.Core.Services;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using ExcelPowerPivotMcp.Tests.Fixtures;

namespace ExcelPowerPivotMcp.Tests.Core.Services;

/// <summary>
/// Unit tests for DataProfileService.
/// Tests data profiling operations with mocked infrastructure dependencies.
/// </summary>
public class DataProfileServiceTests : ServiceTestBase
{
    private DataProfileService CreateService()
    {
        return new DataProfileService(MockDaxService.Object, MockTableComService.Object, MockDmvService.Object, MockDataProfileLogger.Object);
    }

    #region PreviewTableAsync Tests

    [Fact]
    public async Task PreviewTableAsync_ValidTable_ReturnsDaxServiceResults()
    {
        // Arrange
        var tableName = "Sales";
        var maxRows = 100;
        var expectedData = new List<Dictionary<string, object?>>
        {
            new() { ["ProductKey"] = 1, ["Amount"] = 100.50 },
            new() { ["ProductKey"] = 2, ["Amount"] = 250.75 }
        };
        var service = CreateService();

        MockDaxService
            .Setup(s => s.PreviewTableAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedData);

        // Act
        var result = await service.PreviewTableAsync(tableName, maxRows);

        // Assert
        result.Should().BeEquivalentTo(expectedData);
        MockDaxService.Verify(
            s => s.PreviewTableAsync(tableName, maxRows, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task PreviewTableAsync_WithMaxRows_PassesMaxRowsToDaxService()
    {
        // Arrange
        var tableName = "Products";
        var service = CreateService();

        int? capturedMaxRows = null;
        MockDaxService
            .Setup(s => s.PreviewTableAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .Callback<string, int, CancellationToken>((t, max, ct) => capturedMaxRows = max)
            .ReturnsAsync(new List<Dictionary<string, object?>>());

        // Act
        await service.PreviewTableAsync(tableName, 500);

        // Assert
        capturedMaxRows.Should().Be(500);
    }

    [Theory]
    [InlineData(10)]
    [InlineData(100)]
    [InlineData(1000)]
    public async Task PreviewTableAsync_DifferentMaxRows_PassesCorrectLimitToDaxService(int maxRows)
    {
        // Arrange
        var service = CreateService();

        int? capturedMaxRows = null;
        MockDaxService
            .Setup(s => s.PreviewTableAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .Callback<string, int, CancellationToken>((t, max, ct) => capturedMaxRows = max)
            .ReturnsAsync(new List<Dictionary<string, object?>>());

        // Act
        await service.PreviewTableAsync("TestTable", maxRows);

        // Assert
        capturedMaxRows.Should().Be(maxRows);
    }

    [Fact]
    public async Task PreviewTableAsync_DaxServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockDaxService
            .Setup(s => s.PreviewTableAsync(It.IsAny<string>(), It.IsAny<int>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("DAX query failed"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.PreviewTableAsync("TestTable", 100));
    }

    #endregion

    #region GetColumnProfileAsync Tests

    [Fact]
    public async Task GetColumnProfileAsync_ValidColumn_ReturnsProfileData()
    {
        // Arrange
        var tableName = "Sales";
        var columnName = "Amount";
        var expectedProfile = TestDataFactory.CreateColumnProfile(1000, 500, 10);
        var service = CreateService();

        MockDaxService
            .Setup(s => s.GetColumnProfileAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedProfile);

        // Act
        var result = await service.GetColumnProfileAsync(tableName, columnName);

        // Assert
        result.Should().BeEquivalentTo(expectedProfile);
        MockDaxService.Verify(
            s => s.GetColumnProfileAsync(tableName, columnName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetColumnProfileAsync_NumericColumn_ReturnsMinMaxValues()
    {
        // Arrange
        var expectedProfile = TestDataFactory.CreateColumnProfile();
        var service = CreateService();

        MockDaxService
            .Setup(s => s.GetColumnProfileAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedProfile);

        // Act
        var result = await service.GetColumnProfileAsync("Sales", "Amount");

        // Assert
        result.Min.Should().NotBeNull();
        result.Max.Should().NotBeNull();
    }

    [Fact]
    public async Task GetColumnProfileAsync_StringColumn_ReturnsDistinctCount()
    {
        // Arrange
        var expectedProfile = TestDataFactory.CreateColumnProfile(distinctCount: 250);
        var service = CreateService();

        MockDaxService
            .Setup(s => s.GetColumnProfileAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedProfile);

        // Act
        var result = await service.GetColumnProfileAsync("Products", "Category");

        // Assert
        result.DistinctCount.Should().Be(250);
    }

    [Fact]
    public async Task GetColumnProfileAsync_DaxServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockDaxService
            .Setup(s => s.GetColumnProfileAsync(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Column profiling failed"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.GetColumnProfileAsync("TestTable", "TestColumn"));
    }

    #endregion

    #region GetCalculatedColumnsAsync Tests

    [Fact]
    public async Task GetCalculatedColumnsAsync_NoTableName_ReturnsAllCalculatedColumns()
    {
        // Arrange
        var expectedColumns = new List<ColumnMetadata>
        {
            TestDataFactory.CreateColumnMetadata("CalcColumn1", "Table1"),
            TestDataFactory.CreateColumnMetadata("CalcColumn2", "Table2")
        };
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetCalculatedColumnsAsync(null, It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedColumns);

        // Act
        var result = await service.GetCalculatedColumnsAsync(null);

        // Assert
        result.Should().BeEquivalentTo(expectedColumns);
        MockTableComService.Verify(
            s => s.GetCalculatedColumnsAsync(null, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetCalculatedColumnsAsync_WithTableName_ReturnsFilteredColumns()
    {
        // Arrange
        var tableName = "Sales";
        var expectedColumns = new List<ColumnMetadata>
        {
            TestDataFactory.CreateColumnMetadata("FullPrice", tableName),
            TestDataFactory.CreateColumnMetadata("Discount", tableName)
        };
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetCalculatedColumnsAsync(tableName, It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedColumns);

        // Act
        var result = await service.GetCalculatedColumnsAsync(tableName);

        // Assert
        result.Should().HaveCount(2);
        MockTableComService.Verify(
            s => s.GetCalculatedColumnsAsync(tableName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetCalculatedColumnsAsync_ComServiceThrows_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetCalculatedColumnsAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("COM API error"));

        // Act
        var result = await service.GetCalculatedColumnsAsync(null);

        // Assert
        result.Should().BeEmpty();
    }

    [Fact]
    public async Task GetCalculatedColumnsAsync_NoCalculatedColumns_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockTableComService
            .Setup(s => s.GetCalculatedColumnsAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<ColumnMetadata>());

        // Act
        var result = await service.GetCalculatedColumnsAsync(null);

        // Assert
        result.Should().BeEmpty();
    }

    #endregion
}
