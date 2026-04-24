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
/// Unit tests for RelationshipService.
/// Tests all relationship operations with mocked COM service.
/// </summary>
public class RelationshipServiceTests : ServiceTestBase
{
    private RelationshipService CreateService()
    {
        return new RelationshipService(MockRelationshipComService.Object);
    }

    #region CreateRelationshipAsync Tests

    [Fact]
    public async Task CreateRelationshipAsync_ValidRelationship_CallsComServiceOnce()
    {
        // Arrange
        var relationshipDef = TestDataFactory.CreateRelationshipCreate();
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.CreateRelationshipAsync(It.IsAny<RelationshipCreate>(), It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.CreateRelationshipAsync(relationshipDef);

        // Assert
        MockRelationshipComService.Verify(
            s => s.CreateRelationshipAsync(relationshipDef, It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task CreateRelationshipAsync_SameTableRelationship_CallsComServiceWithSameTable()
    {
        // Arrange
        var relationshipDef = TestDataFactory.CreateRelationshipCreate(
            foreignTable: "Employee",
            foreignColumn: "ManagerId",
            primaryTable: "Employee",
            primaryColumn: "EmployeeId");
        var service = CreateService();

        RelationshipCreate? capturedArg = null;
        MockRelationshipComService
            .Setup(s => s.CreateRelationshipAsync(It.IsAny<RelationshipCreate>(), It.IsAny<CancellationToken>()))
            .Callback<RelationshipCreate, CancellationToken>((r, ct) => capturedArg = r)
            .Returns(Task.CompletedTask);

        // Act
        await service.CreateRelationshipAsync(relationshipDef);

        // Assert
        capturedArg.Should().NotBeNull();
        capturedArg!.ForeignTable.Should().Be("Employee");
        capturedArg.PrimaryTable.Should().Be("Employee");
    }

    [Fact]
    public async Task CreateRelationshipAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var relationshipDef = TestDataFactory.CreateRelationshipCreate();
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.CreateRelationshipAsync(It.IsAny<RelationshipCreate>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Relationship already exists"));

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.CreateRelationshipAsync(relationshipDef));

        exception.Message.Should().Contain("already exists");
    }

    [Fact]
    public async Task CreateRelationshipAsync_CancellationRequested_ThrowsOperationCanceledException()
    {
        // Arrange
        var relationshipDef = TestDataFactory.CreateRelationshipCreate();
        var service = CreateService();
        var cts = new CancellationTokenSource();
        cts.Cancel();

        MockRelationshipComService
            .Setup(s => s.CreateRelationshipAsync(It.IsAny<RelationshipCreate>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new OperationCanceledException());

        // Act & Assert
        await Assert.ThrowsAsync<OperationCanceledException>(
            () => service.CreateRelationshipAsync(relationshipDef, cts.Token));
    }

    #endregion

    #region DeleteRelationshipAsync Tests

    [Fact]
    public async Task DeleteRelationshipAsync_ValidRelationship_CallsComServiceOnce()
    {
        // Arrange
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.DeleteRelationshipAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteRelationshipAsync("Sales", "ProductKey", "Product", "ProductKey");

        // Assert
        MockRelationshipComService.Verify(
            s => s.DeleteRelationshipAsync("Sales", "ProductKey", "Product", "ProductKey", It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task DeleteRelationshipAsync_WithAllParameters_PassesCorrectValuesToComService()
    {
        // Arrange
        var service = CreateService();
        string? capturedForeignTable = null;
        string? capturedForeignColumn = null;
        string? capturedPrimaryTable = null;
        string? capturedPrimaryColumn = null;

        MockRelationshipComService
            .Setup(s => s.DeleteRelationshipAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .Callback<string, string, string, string, CancellationToken>((ft, fc, pt, pc, ct) =>
            {
                capturedForeignTable = ft;
                capturedForeignColumn = fc;
                capturedPrimaryTable = pt;
                capturedPrimaryColumn = pc;
            })
            .Returns(Task.CompletedTask);

        // Act
        await service.DeleteRelationshipAsync("Orders", "CustomerKey", "Customer", "CustomerKey");

        // Assert
        capturedForeignTable.Should().Be("Orders");
        capturedForeignColumn.Should().Be("CustomerKey");
        capturedPrimaryTable.Should().Be("Customer");
        capturedPrimaryColumn.Should().Be("CustomerKey");
    }

    [Fact]
    public async Task DeleteRelationshipAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.DeleteRelationshipAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<CancellationToken>()))
            .ThrowsAsync(new KeyNotFoundException("Relationship not found"));

        // Act & Assert
        await Assert.ThrowsAsync<KeyNotFoundException>(
            () => service.DeleteRelationshipAsync("Sales", "ProductKey", "Product", "ProductKey"));
    }

    #endregion

    #region GetRelationshipsAsync Tests

    [Fact]
    public async Task GetRelationshipsAsync_HasRelationships_ReturnsAllRelationships()
    {
        // Arrange
        var expectedRelationships = new List<RelationshipMetadata>
        {
            TestDataFactory.CreateRelationshipMetadata("Sales", "ProductKey", "Product", "ProductKey", true),
            TestDataFactory.CreateRelationshipMetadata("Sales", "CustomerKey", "Customer", "CustomerKey", true)
        };
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(expectedRelationships);

        // Act
        var result = await service.GetRelationshipsAsync();

        // Assert
        result.Should().BeEquivalentTo(expectedRelationships);
        MockRelationshipComService.Verify(
            s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()),
            Times.Once);
    }

    [Fact]
    public async Task GetRelationshipsAsync_NoRelationships_ReturnsEmptyList()
    {
        // Arrange
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ReturnsAsync(new List<RelationshipMetadata>());

        // Act
        var result = await service.GetRelationshipsAsync();

        // Assert
        result.Should().BeEmpty();
    }

    [Fact]
    public async Task GetRelationshipsAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.GetRelationshipsAsync(It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Failed to query relationships"));

        // Act & Assert
        await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.GetRelationshipsAsync());
    }

    #endregion

    #region SetRelationshipActiveAsync Tests

    [Fact]
    public async Task SetRelationshipActiveAsync_SetToActive_CallsComServiceWithTrue()
    {
        // Arrange
        var service = CreateService();
        bool? capturedActive = null;

        MockRelationshipComService
            .Setup(s => s.SetRelationshipActiveAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<bool>(),
                It.IsAny<CancellationToken>()))
            .Callback<string, string, string, string, bool, CancellationToken>((ft, fc, pt, pc, active, ct) =>
                capturedActive = active)
            .Returns(Task.CompletedTask);

        // Act
        await service.SetRelationshipActiveAsync("Sales", "ProductKey", "Product", "ProductKey", true);

        // Assert
        capturedActive.Should().BeTrue();
    }

    [Fact]
    public async Task SetRelationshipActiveAsync_SetToInactive_CallsComServiceWithFalse()
    {
        // Arrange
        var service = CreateService();
        bool? capturedActive = null;

        MockRelationshipComService
            .Setup(s => s.SetRelationshipActiveAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<bool>(),
                It.IsAny<CancellationToken>()))
            .Callback<string, string, string, string, bool, CancellationToken>((ft, fc, pt, pc, active, ct) =>
                capturedActive = active)
            .Returns(Task.CompletedTask);

        // Act
        await service.SetRelationshipActiveAsync("Sales", "ProductKey", "Product", "ProductKey", false);

        // Assert
        capturedActive.Should().BeFalse();
    }

    [Fact]
    public async Task SetRelationshipActiveAsync_ValidParameters_PassesAllParametersToComService()
    {
        // Arrange
        var service = CreateService();
        string? capturedForeignTable = null;
        string? capturedForeignColumn = null;
        string? capturedPrimaryTable = null;
        string? capturedPrimaryColumn = null;
        bool? capturedActive = null;

        MockRelationshipComService
            .Setup(s => s.SetRelationshipActiveAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<bool>(),
                It.IsAny<CancellationToken>()))
            .Callback<string, string, string, string, bool, CancellationToken>((ft, fc, pt, pc, active, ct) =>
            {
                capturedForeignTable = ft;
                capturedForeignColumn = fc;
                capturedPrimaryTable = pt;
                capturedPrimaryColumn = pc;
                capturedActive = active;
            })
            .Returns(Task.CompletedTask);

        // Act
        await service.SetRelationshipActiveAsync("Orders", "CustomerKey", "Customer", "CustomerKey", true);

        // Assert
        capturedForeignTable.Should().Be("Orders");
        capturedForeignColumn.Should().Be("CustomerKey");
        capturedPrimaryTable.Should().Be("Customer");
        capturedPrimaryColumn.Should().Be("CustomerKey");
        capturedActive.Should().BeTrue();
    }

    [Fact]
    public async Task SetRelationshipActiveAsync_ComServiceThrows_PropagatesException()
    {
        // Arrange
        var service = CreateService();

        MockRelationshipComService
            .Setup(s => s.SetRelationshipActiveAsync(
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<string>(),
                It.IsAny<bool>(),
                It.IsAny<CancellationToken>()))
            .ThrowsAsync(new InvalidOperationException("Cannot change relationship state"));

        // Act & Assert
        var exception = await Assert.ThrowsAsync<InvalidOperationException>(
            () => service.SetRelationshipActiveAsync("Sales", "ProductKey", "Product", "ProductKey", true));

        exception.Message.Should().Contain("Cannot change");
    }

    #endregion
}
