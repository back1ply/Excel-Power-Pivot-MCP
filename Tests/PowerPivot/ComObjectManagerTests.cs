using System;
using Xunit;
using FluentAssertions;
using ExcelPowerPivotMcp.PowerPivot;

namespace ExcelPowerPivotMcp.Tests.PowerPivot;

/// <summary>
/// Unit tests for ComObjectManager (ComScope).
/// Tests COM object lifecycle management.
/// </summary>
public class ComObjectManagerTests
{
    #region ComScope Track Tests

    [Fact]
    public void Track_NullObject_ReturnsNull()
    {
        // Arrange
        using var scope = new ComScope();

        // Act
        // Track<T> where T : class accepts T? and returns T? (nullable reference type)
        var result = scope.Track<object>(null);

        // Assert
        result.Should().BeNull();
    }

    [Fact]
    public void Track_NonComObject_ReturnsObjectWithoutTracking()
    {
        // Arrange
        using var scope = new ComScope();
        var regularObject = new object();

        // Act
        var result = scope.Track(regularObject);

        // Assert
        result.Should().BeSameAs(regularObject);
    }

    [Fact]
    public void TrackDynamic_NullObject_ReturnsNull()
    {
        // Arrange
        using var scope = new ComScope();

        // Act
        // TrackDynamic accepts dynamic and handles null internally
        dynamic? result = scope.TrackDynamic(null!);

        // Assert
        ((object?)result).Should().BeNull();
    }

    [Fact]
    public void TrackDynamic_NonNullObject_ReturnsObject()
    {
        // Arrange
        using var scope = new ComScope();
        var obj = new { Name = "Test" };

        // Act
        dynamic result = scope.TrackDynamic(obj);

        // Assert
        ((object)result).Should().BeSameAs(obj);
    }

    #endregion

    #region ComScope Dispose Tests

    [Fact]
    public void Dispose_MultipleCalls_DoesNotThrow()
    {
        // Arrange
        var scope = new ComScope();

        // Act & Assert
        scope.Dispose();
        var act = () => scope.Dispose();
        act.Should().NotThrow();
    }

    [Fact]
    public void Dispose_WithTrackedObjects_Completes()
    {
        // Arrange
        var scope = new ComScope();
        scope.Track(new object());
        scope.Track(new object());

        // Act & Assert
        var act = () => scope.Dispose();
        act.Should().NotThrow();
    }

    #endregion

    #region ComHelper SafeRelease Tests

    [Fact]
    public void SafeRelease_NullObject_DoesNotThrow()
    {
        // Act & Assert
        var act = () => ComHelper.SafeRelease(null);
        act.Should().NotThrow();
    }

    [Fact]
    public void SafeRelease_NonComObject_DoesNotThrow()
    {
        // Arrange
        var regularObject = new object();

        // Act & Assert
        var act = () => ComHelper.SafeRelease(regularObject);
        act.Should().NotThrow();
    }

    #endregion

    #region TrackProperty Tests

    [Fact]
    public void TrackProperty_AccessorThrows_PropagatesException()
    {
        // Arrange
        using var scope = new ComScope();

        // Act & Assert
        var act = () => scope.TrackProperty(() => throw new InvalidOperationException("Test exception"));
        act.Should().Throw<InvalidOperationException>().WithMessage("Test exception");
    }

    [Fact]
    public void TrackProperty_ValidAccessor_ReturnsValue()
    {
        // Arrange
        using var scope = new ComScope();
        var expectedValue = new { Name = "Test" };

        // Act
        var result = scope.TrackProperty(() => (dynamic)expectedValue);

        // Assert
        ((object)result).Should().BeSameAs(expectedValue);
    }

    [Fact]
    public void TrackProperty_NullValue_ReturnsNull()
    {
        // Arrange
        using var scope = new ComScope();

        // Act
        var result = scope.TrackProperty(() => null!);

        // Assert
        ((object?)result).Should().BeNull();
    }

    #endregion

    #region Integration Tests

    [Fact]
    public void ComScope_TrackMultipleObjects_AllReleasedOnDispose()
    {
        // Arrange
        var scope = new ComScope();
        var obj1 = new object();
        var obj2 = new object();
        var obj3 = new object();

        // Act
        scope.Track(obj1);
        scope.Track(obj2);
        scope.Track(obj3);

        // Assert - dispose should handle all objects
        var act = () => scope.Dispose();
        act.Should().NotThrow();
    }

    [Fact]
    public void ComScope_WithUsing_DisposesAutomatically()
    {
        // Act & Assert
        var act = () =>
        {
            using var scope = new ComScope();
            scope.Track(new object());
            scope.Track(new object());
            // Automatic disposal when scope exits
        };

        act.Should().NotThrow();
    }

    #endregion
}
