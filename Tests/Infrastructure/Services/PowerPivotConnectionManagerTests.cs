using System;
using Moq;
using Xunit;
using FluentAssertions;
using Microsoft.Extensions.Logging;
using ExcelPowerPivotMcp.Infrastructure.Services;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Tests.Infrastructure.Services;

/// <summary>
/// Unit tests for PowerPivotConnectionManager.
/// Tests logger injection improvement from Issue #4 fix.
/// </summary>
public class PowerPivotConnectionManagerTests
{
    private readonly Mock<ILogger<PowerPivotConnection>> _mockLogger;

    public PowerPivotConnectionManagerTests()
    {
        _mockLogger = new Mock<ILogger<PowerPivotConnection>>();

        // Reset singleton before each test
        PowerPivotConnection.Reset();
    }

    #region Constructor Tests (Issue #4 Fix)

    [Fact]
    public void Constructor_InjectsLogger_LoggerIsAvailable()
    {
        // Act
        var manager = new PowerPivotConnectionManager(_mockLogger.Object);

        // Assert
        manager.Should().NotBeNull();
        // Logger should be set via SetLogger() in constructor
        // We can't directly verify this without exposing internal state,
        // but we can verify the manager works correctly
    }

    [Fact]
    public void Constructor_WithNullLogger_DoesNotThrow()
    {
        // Act
        var act = () => new PowerPivotConnectionManager(null!);

        // Assert
        // Should not throw - logger is passed to PowerPivotConnection.SetLogger
        // which accepts nullable ILogger
        act.Should().NotThrow();
    }

    #endregion

    #region Current Property Tests

    [Fact]
    public void Current_ReturnsInstance()
    {
        // Arrange
        var manager = new PowerPivotConnectionManager(_mockLogger.Object);

        // Act
        var current = manager.Current;

        // Assert
        current.Should().NotBeNull();
        current.Should().BeSameAs(PowerPivotConnection.Instance);
    }

    [Fact]
    public void Current_AfterReset_ReturnsNewInstance()
    {
        // Arrange
        var manager = new PowerPivotConnectionManager(_mockLogger.Object);
        var first = manager.Current;

        // Act
        PowerPivotConnection.Reset();
        var second = manager.Current;

        // Assert
        second.Should().NotBeNull();
        second.Should().NotBeSameAs(first); // New instance after reset
    }

    #endregion

    #region IsConnected Property Tests

    [Fact]
    public void IsConnected_InitialState_ReturnsFalse()
    {
        // Arrange
        var manager = new PowerPivotConnectionManager(_mockLogger.Object);

        // Act
        var isConnected = manager.IsConnected;

        // Assert
        isConnected.Should().BeFalse();
    }

    #endregion

    #region ConnectedWorkbook Property Tests

    [Fact]
    public void ConnectedWorkbook_InitialState_ReturnsNull()
    {
        // Arrange
        var manager = new PowerPivotConnectionManager(_mockLogger.Object);

        // Act
        var workbook = manager.ConnectedWorkbook;

        // Assert
        workbook.Should().BeNull();
    }

    #endregion

    #region Reset Method Tests

    [Fact]
    public void Reset_DisconnectsCurrentInstance()
    {
        // Arrange
        var manager = new PowerPivotConnectionManager(_mockLogger.Object);
        var before = manager.Current;

        // Act
        manager.Reset();
        var after = manager.Current;

        // Assert
        after.Should().NotBeSameAs(before);
        after.IsConnected.Should().BeFalse();
    }

    #endregion

    #region IsFatalComError Static Method Tests

    [Theory]
    [InlineData(unchecked((int)0x80010108))] // RPC_E_DISCONNECTED
    [InlineData(unchecked((int)0x800401FD))] // CO_E_OBJNOTCONNECTED
    [InlineData(unchecked((int)0x80010007))] // RPC_E_SERVER_DIED
    [InlineData(unchecked((int)0x8001010D))] // RPC_E_SERVER_DIED_DNE
    [InlineData(unchecked((int)0x800706BA))] // RPC_S_SERVER_UNAVAILABLE
    [InlineData(unchecked((int)0x800706BE))] // RPC_S_CALL_FAILED
    public void IsFatalComError_FatalHResults_ReturnsTrue(int hResult)
    {
        // Arrange
        var exception = new System.Runtime.InteropServices.COMException("Test", hResult);

        // Act
        var isFatal = PowerPivotConnectionManager.IsFatalComError(exception);

        // Assert
        isFatal.Should().BeTrue($"HRESULT 0x{hResult:X} should be considered fatal");
    }

    [Theory]
    [InlineData(unchecked((int)0x80004005))] // E_FAIL (generic)
    [InlineData(unchecked((int)0x80070005))] // E_ACCESSDENIED
    [InlineData(0x00000000)] // S_OK (doesn't need unchecked)
    public void IsFatalComError_NonFatalHResults_ReturnsFalse(int hResult)
    {
        // Arrange
        var exception = new System.Runtime.InteropServices.COMException("Test", hResult);

        // Act
        var isFatal = PowerPivotConnectionManager.IsFatalComError(exception);

        // Assert
        isFatal.Should().BeFalse($"HRESULT 0x{hResult:X} should not be considered fatal");
    }

    #endregion
}
