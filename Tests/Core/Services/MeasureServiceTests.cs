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
/// Unit tests for MeasureService.
/// Tests all CRUD operations with mocked infrastructure dependencies.
/// </summary>
public class MeasureServiceTests : ServiceTestBase
{
    private MeasureService CreateService()
    {
        return new MeasureService(MockMeasureComService.Object, MockDmvService.Object);
    }

    #region CreateMeasureAsync Tests

    [Fact]
    public async Task CreateMeasureAsync_ValidMeasure_CallsComServiceOnce()
    {
        // Arrange
        var measureDef = TestDataFactory.CreateMeasureCreate();
        var service = CreateService();

        MockMeasureComService
            .Setup(s => s.CreateMeasureAsync(It.IsAny<MeasureCreate>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.CreateMeasureAsync(measureDef);

        // Assert
        MockMeasureComService.Verify(
            s => s.CreateMeasureAsync(measureDef, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task CreateMeasureAsync_WithDescription_PassesDescriptionToComService()
    {
        // Arrange
        var measureDef = TestDataFactory.CreateMeasureCreate(
            description: "Test measure description");
        var service = CreateService();

        MeasureCreate? capturedArg = null;
        MockMeasureComService
            .Setup(s => s.CreateMeasureAsync(It.IsAny<MeasureCreate>(), It.IsAny<CancellationToken>()))
            .Callback<MeasureCreate, CancellationToken>((m, ct) => capturedArg = m)
            .Returns(Task.CompletedTask);

        // Act
        await service.CreateMeasureAsync(measureDef);

        // Assert
        capturedArg.Should().NotBeNull();
        capturedArg!.Description.Should().Be("Test measure description");
    }

    [Fact]
    public async Task CreateMeasureAsync_WithFormatString_PassesFormatStringToComService()
    {
        // Arrange
        var measureDef = TestDataFactory.CreateMeasureCreate(
            formatString: "#,##0.00");
        var service = CreateService();

        MeasureCreate? capturedArg = null;
        MockMeasureComService
            .Setup(s => s.CreateMeasureAsync(It.IsAny<MeasureCreate>(), It.IsAny<CancellationToken>()))
            .Callback<MeasureCreate, CancellationToken>((m, ct) => capturedArg = m)
            .Returns(Task.CompletedTask);

        // Act
        await service.CreateMeasureAsync(measureDef);

        // Assert
        capturedArg.Should().NotBeNull();
        capturedArg!.FormatString.Should().Be("#,##0.00");
    }

    [Fact]
    public async Task CreateMeasureAsync_CancellationRequested_ThrowsOperationCanceledException()
    {
        // Arrange
        var measureDef = TestDataFactory.CreateMeasureCreate();
        var service = CreateService();
        var cts = new CancellationTokenSource();
        cts.Cancel();

        MockMeasureComService
            .Setup(s => s.CreateMeasureAsync(It.IsAny<MeasureCreate>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new OperationCanceledException());

        // Act & Assert
        await Assert.ThrowsAsync<OperationCanceledException>(
            () => service.CreateMeasureAsync(measureDef, cts.Token));
    }

    [Fact]
    public async Task CreateMeasureAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var measureDef = TestDataFactory.CreateMeasureCreate();
        var service = CreateService();
        var expectedException = new InvalidOperationException("COM API error");

        MockMeasureComService
            .Setup(s => s.CreateMeasureAsync(It.IsAny<MeasureCreate>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(expectedException);

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.CreateMeasureAsync(measureDef));

        exception.Message.Should().Be("COM API error");
    }

    #endregion

    #region UpdateMeasureAsync Tests

    [Fact]
    public async Task UpdateMeasureAsync_ValidUpdate_CallsComServiceOnce()
    {
        // Arrange
        var updateDef = TestDataFactory.CreateMeasureUpdate(
            measureName: "TestMeasure",
            expression: "COUNT([Rows])");
        var service = CreateService();

        MockMeasureComService
            .Setup(s => s.UpdateMeasureAsync(It.IsAny<MeasureUpdate>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.UpdateMeasureAsync(updateDef);

        // Assert
        MockMeasureComService.Verify(
            s => s.UpdateMeasureAsync(updateDef, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task UpdateMeasureAsync_ChangesExpression_PassesNewExpressionToComService()
    {
        // Arrange
        var updateDef = TestDataFactory.CreateMeasureUpdate(
            measureName: "TestMeasure",
            expression: "AVERAGE([Amount])");
        var service = CreateService();

        MeasureUpdate? capturedArg = null;
        MockMeasureComService
            .Setup(s => s.UpdateMeasureAsync(It.IsAny<MeasureUpdate>(), It.IsAny<CancellationToken>()))
            .Callback<MeasureUpdate, CancellationToken>((m, ct) => capturedArg = m)
            .Returns(Task.CompletedTask);

        // Act
        await service.UpdateMeasureAsync(updateDef);

        // Assert
        capturedArg.Should().NotBeNull();
        capturedArg!.Expression.Should().Be("AVERAGE([Amount])");
    }

    [Fact]
    public async Task UpdateMeasureAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var updateDef = TestDataFactory.CreateMeasureUpdate(
            measureName: "TestMeasure",
            expression: "INVALID DAX");
        var service = CreateService();

        MockMeasureComService
            .Setup(s => s.UpdateMeasureAsync(It.IsAny<MeasureUpdate>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Invalid DAX expression"));

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.UpdateMeasureAsync(updateDef));

        exception.Message.Should().Contain("Invalid DAX");
    }

    #endregion

    #region DeleteMeasureAsync Tests

    [Fact]
    public async Task DeleteMeasureAsync_ValidMeasureName_CallsComServiceOnce()
    {
        // Arrange
        var measureName = "TestMeasure";
        var service = CreateService();

        MockMeasureComService
            .Setup(s => s.DeleteMeasureAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteMeasureAsync(measureName);

        // Assert
        MockMeasureComService.Verify(
            s => s.DeleteMeasureAsync(measureName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task DeleteMeasureAsync_EmptyMeasureName_CallsComServiceWithEmptyString()
    {
        // Arrange
        var measureName = string.Empty;
        var service = CreateService();

        MockMeasureComService
            .Setup(s => s.DeleteMeasureAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteMeasureAsync(measureName);

        // Assert
        MockMeasureComService.Verify(
            s => s.DeleteMeasureAsync(string.Empty, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task DeleteMeasureAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var measureName = "NonExistentMeasure";
        var service = CreateService();

        MockMeasureComService
            .Setup(s => s.DeleteMeasureAsync(It.IsAny<string>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new KeyNotFoundException("Measure not found"));

        // Act & Assert
        await Assert.ThrowsAsync<KeyNotFoundException>(
            () => service.DeleteMeasureAsync(measureName));
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
        var result = await service.GetMeasuresAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedMeasures);
        MockDmvService.Verify(
            s => s.GetMeasuresAsync(null, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetMeasuresAsync_WithTableName_ReturnsFilteredMeasures()
    {
        // Arrange
        var tableName = "SalesTable";
        var expectedMeasures = new List<MeasureMetadata>
        {
            TestDataFactory.CreateMeasureMetadata("TotalSales", tableName),
            TestDataFactory.CreateMeasureMetadata("AvgSales", tableName)
        };
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetMeasuresAsync(tableName, It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedMeasures);

        // Act
        var result = await service.GetMeasuresAsync(tableName);

        // Assert
        result.Should().HaveCount(2);
        result.Should().AllSatisfy(m => m.TableName.Should().Be(tableName));
        MockDmvService.Verify(
            s => s.GetMeasuresAsync(tableName, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetMeasuresAsync_EmptyResult_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetMeasuresAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<MeasureMetadata>());

        // Act
        var result = await service.GetMeasuresAsync();

        // Assert
        result.Should().BeEmpty();
    }

    [Fact]
    public async Task GetMeasuresAsync_DmvServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockDmvService
            .Setup(s => s.GetMeasuresAsync(It.IsAny<string?>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("DMV query failed"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.GetMeasuresAsync());
    }

    #endregion
}
