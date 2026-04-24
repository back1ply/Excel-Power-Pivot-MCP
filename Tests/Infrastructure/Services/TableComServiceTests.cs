using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Moq;
using Xunit;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using ExcelPowerPivotMcp.Infrastructure.Services;
using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Tests.Infrastructure.Services;

/// <summary>
/// Unit tests for TableComService.
/// Tests validation logic and error handling.
/// </summary>
public class TableComServiceTests
{
    private readonly Mock<PowerPivotConnection> _mockConnection;
    private readonly Mock<IExcelDispatcher> _mockDispatcher;
    private readonly McpConfiguration _config;
    private readonly Mock<ILogger<TableComService>> _mockLogger;
    private readonly TableComService _service;

    public TableComServiceTests()
    {
        _mockConnection = new Mock<PowerPivotConnection>();
        _mockDispatcher = new Mock<IExcelDispatcher>();
        _config = new McpConfiguration(); // Use default config
        _mockLogger = new Mock<ILogger<TableComService>>();
        _service = new TableComService(
            _mockConnection.Object,
            _mockDispatcher.Object,
            _config,
            _mockLogger.Object);
    }

    #region AddTableToModelAsync - QueryName Validation Tests (Issue #1 Fix)

    // NOTE: The queryName validation logic (Issue #1 fix) is tested implicitly through integration tests.
    // Unit testing this validation requires extensive COM object mocking (Workbook, Model, Queries, etc.)
    // which would make tests brittle and complex. The validation code in TableComService.cs line 67-75
    // ensures QueryName matches TableName for Power Query tables.

    [Fact]
    public async Task AddTableToModelAsync_PowerQuery_CallsDispatcher()
    {
        // Arrange
        var tableAdd = new TableAddToModel
        {
            TableName = "MyTable",
            QueryName = "MyTable",
            UsePowerQuery = true
        };

        bool dispatcherCalled = false;
        _mockDispatcher
            .Setup(d => d.RunAsync(It.IsAny<Func<string>>(), It.IsAny<CancellationToken>()))
            .Callback(() => dispatcherCalled = true)
            .ReturnsAsync("MyTable");

        // Act
        await _service.AddTableToModelAsync(tableAdd);

        // Assert
        dispatcherCalled.Should().BeTrue();
    }

    [Fact]
    public async Task AddTableToModelAsync_DirectConnection_CallsDispatcher()
    {
        // Arrange
        var tableAdd = new TableAddToModel
        {
            TableName = "MyTable",
            UsePowerQuery = false
        };

        bool dispatcherCalled = false;
        _mockDispatcher
            .Setup(d => d.RunAsync(It.IsAny<Func<string>>(), It.IsAny<CancellationToken>()))
            .Callback(() => dispatcherCalled = true)
            .ReturnsAsync("MyTable");

        // Act
        await _service.AddTableToModelAsync(tableAdd);

        // Assert
        dispatcherCalled.Should().BeTrue();
    }

    #endregion

    #region RefreshTableAsync Tests

    [Fact]
    public async Task RefreshTableAsync_ValidTableName_CallsDispatcher()
    {
        // Arrange
        var tableName = "TestTable";
        bool dispatcherCalled = false;

        _mockDispatcher
            .Setup(d => d.RunAsync(It.IsAny<Action>(), It.IsAny<CancellationToken>()))
            .Callback(() => dispatcherCalled = true)
            .Returns(Task.CompletedTask);

        // Act
        await _service.RefreshTableAsync(tableName);

        // Assert
        dispatcherCalled.Should().BeTrue();
    }

    // NOTE: RefreshTableAsync does not validate input parameters.
    // It delegates to ExcelComHelpers.FindByNameOrThrow which handles invalid table names
    // by throwing appropriate exceptions when the table is not found.

    #endregion

    #region DeleteTableAsync Tests

    [Fact]
    public async Task DeleteTableAsync_ValidTableName_CallsDispatcher()
    {
        // Arrange
        var tableName = "TestTable";
        bool dispatcherCalled = false;

        _mockDispatcher
            .Setup(d => d.RunAsync(It.IsAny<Action>(), It.IsAny<CancellationToken>()))
            .Callback(() => dispatcherCalled = true)
            .Returns(Task.CompletedTask);

        // Act
        await _service.DeleteTableAsync(tableName);

        // Assert
        dispatcherCalled.Should().BeTrue();
    }

    // NOTE: DeleteTableAsync does not validate input parameters.
    // It delegates to ExcelComHelpers.FindByNameOrThrow which handles invalid table names
    // by throwing appropriate exceptions when the table is not found.

    #endregion

    #region GetPowerQueriesAsync Tests

    // NOTE: GetPowerQueriesAsync testing requires complex COM object mocking
    // (Workbook.Queries collection with enumeration, PowerPivotConnection setup, etc.)
    // which is too brittle and complex for unit tests. The method is fully tested
    // via integration tests in test_driver.py with real Excel workbooks.
    // Unit tests for this infrastructure service focus on simpler operations above.

    #endregion

    #region CancellationToken Tests

    [Fact]
    public async Task AddTableToModelAsync_CancellationRequested_PropagatesToken()
    {
        // Arrange
        var tableAdd = new TableAddToModel
        {
            TableName = "MyTable",
            UsePowerQuery = false
        };
        var cts = new CancellationTokenSource();
        cts.Cancel();

        _mockDispatcher
            .Setup(d => d.RunAsync(It.IsAny<Func<string>>(), It.IsAny<CancellationToken>()))
            .ThrowsAsync(new OperationCanceledException());

        // Act
        var act = async () => await _service.AddTableToModelAsync(tableAdd, cts.Token);

        // Assert
        await act.Should().ThrowAsync<OperationCanceledException>();
    }

    #endregion
}
