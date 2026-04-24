using System;
using System.IO;
using Xunit;
using FluentAssertions;
using ExcelPowerPivotMcp;

namespace ExcelPowerPivotMcp.Tests.Configuration;

/// <summary>
/// Unit tests for McpConfiguration environment variable parsing.
/// Tests validation logic added in Issue #5 fix.
/// </summary>
public class McpConfigurationTests : IDisposable
{
    private readonly StringWriter _errorOutput;
    private readonly TextWriter _originalError;

    public McpConfigurationTests()
    {
        // Capture stderr to test warning messages
        _errorOutput = new StringWriter();
        _originalError = Console.Error;
        Console.SetError(_errorOutput);
    }

    public void Dispose()
    {
        Console.SetError(_originalError);
        _errorOutput.Dispose();

        // Clean up environment variables
        Environment.SetEnvironmentVariable("MCP_SERVER_NAME", null);
        Environment.SetEnvironmentVariable("MCP_SERVER_VERSION", null);
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", null);
        Environment.SetEnvironmentVariable("MCP_QUERY_TIMEOUT_SECONDS", null);
        Environment.SetEnvironmentVariable("MCP_DAX_FORMATTER_CACHE_ENABLED", null);
        Environment.SetEnvironmentVariable("MCP_CONNECTION_RETRY_COUNT", null);
    }

    #region Integer Environment Variable Tests

    [Fact]
    public void Load_ValidIntegerEnvironmentVariable_AppliesValue()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", "5000");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.MaxQueryRows.Should().Be(5000);
        _errorOutput.ToString().Should().BeEmpty(); // No warnings
    }

    [Fact]
    public void Load_InvalidIntegerEnvironmentVariable_UsesDefaultAndWarns()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", "invalid123");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.MaxQueryRows.Should().Be(1000); // Default value
        _errorOutput.ToString().Should().Contain("[Warning]");
        _errorOutput.ToString().Should().Contain("MCP_MAX_QUERY_ROWS");
        _errorOutput.ToString().Should().Contain("invalid123");
        _errorOutput.ToString().Should().Contain("Expected an integer");
    }

    [Fact]
    public void Load_BelowMinimumInteger_UsesDefaultAndWarns()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", "-100");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.MaxQueryRows.Should().Be(1000); // Default value
        _errorOutput.ToString().Should().Contain("[Warning]");
        _errorOutput.ToString().Should().Contain("MCP_MAX_QUERY_ROWS");
        _errorOutput.ToString().Should().Contain("-100");
        _errorOutput.ToString().Should().Contain("below minimum");
    }

    [Fact]
    public void Load_ZeroValue_UsesDefaultWhenMinimumIsOne()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_QUERY_TIMEOUT_SECONDS", "0");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.QueryTimeoutSeconds.Should().Be(120); // Default value
        _errorOutput.ToString().Should().Contain("[Warning]");
    }

    [Fact]
    public void Load_ZeroValue_AcceptsWhenMinimumIsZero()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_CONNECTION_RETRY_COUNT", "0");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.ConnectionRetryCount.Should().Be(0); // Zero is valid for retry count
        _errorOutput.ToString().Should().BeEmpty(); // No warnings
    }

    [Fact]
    public void Load_EmptyEnvironmentVariable_UsesDefault()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", "");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.MaxQueryRows.Should().Be(1000); // Default value
        _errorOutput.ToString().Should().BeEmpty(); // No warnings for empty
    }

    #endregion

    #region Boolean Environment Variable Tests

    [Theory]
    [InlineData("true", true)]
    [InlineData("True", true)]
    [InlineData("TRUE", true)]
    [InlineData("false", false)]
    [InlineData("False", false)]
    [InlineData("FALSE", false)]
    public void Load_ValidBooleanEnvironmentVariable_AppliesValue(string envValue, bool expected)
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_DAX_FORMATTER_CACHE_ENABLED", envValue);

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.DaxFormatterCacheEnabled.Should().Be(expected);
        _errorOutput.ToString().Should().BeEmpty(); // No warnings
    }

    [Theory]
    [InlineData("yes")]
    [InlineData("1")]
    [InlineData("invalid")]
    [InlineData("on")]
    public void Load_InvalidBooleanEnvironmentVariable_UsesDefaultAndWarns(string invalidValue)
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_DAX_FORMATTER_CACHE_ENABLED", invalidValue);

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.DaxFormatterCacheEnabled.Should().BeTrue(); // Default value
        _errorOutput.ToString().Should().Contain("[Warning]");
        _errorOutput.ToString().Should().Contain("MCP_DAX_FORMATTER_CACHE_ENABLED");
        _errorOutput.ToString().Should().Contain(invalidValue);
        _errorOutput.ToString().Should().Contain("'true' or 'false'");
    }

    #endregion

    #region Multiple Environment Variables

    [Fact]
    public void Load_MultipleValidVariables_AppliesAllValues()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", "2000");
        Environment.SetEnvironmentVariable("MCP_QUERY_TIMEOUT_SECONDS", "60");
        Environment.SetEnvironmentVariable("MCP_DAX_FORMATTER_CACHE_ENABLED", "false");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.MaxQueryRows.Should().Be(2000);
        config.QueryTimeoutSeconds.Should().Be(60);
        config.DaxFormatterCacheEnabled.Should().BeFalse();
        _errorOutput.ToString().Should().BeEmpty(); // No warnings
    }

    [Fact]
    public void Load_MixOfValidAndInvalid_AppliesValidAndWarnsAboutInvalid()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_MAX_QUERY_ROWS", "2000"); // Valid
        Environment.SetEnvironmentVariable("MCP_QUERY_TIMEOUT_SECONDS", "abc"); // Invalid

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.MaxQueryRows.Should().Be(2000); // Applied
        config.QueryTimeoutSeconds.Should().Be(120); // Default (invalid)

        var output = _errorOutput.ToString();
        output.Should().NotContain("MCP_MAX_QUERY_ROWS"); // No warning for valid
        output.Should().Contain("MCP_QUERY_TIMEOUT_SECONDS"); // Warning for invalid
    }

    #endregion

    #region Default Values

    [Fact]
    public void Load_NoEnvironmentVariables_UsesAllDefaults()
    {
        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.ServerName.Should().Be("excel-powerpivot-mcp");
        config.MaxQueryRows.Should().Be(1000);
        config.QueryTimeoutSeconds.Should().Be(120);
        config.ValidationTimeoutSeconds.Should().Be(10);
        config.DmvTimeoutSeconds.Should().Be(30);
        config.ConnectionRetryCount.Should().Be(3);
        config.ConnectionRetryDelayMs.Should().Be(1000);
        config.ConnectionMaxRetryDelayMs.Should().Be(8000);
        config.ConnectionTimeoutMs.Should().Be(10000);
        config.DaxFormatterTimeoutSeconds.Should().Be(10);
        config.DaxFormatterRetryCount.Should().Be(2);
        config.DaxFormatterRetryDelayMs.Should().Be(500);
        config.DaxFormatterMaxRetryDelayMs.Should().Be(3000);
        config.DaxFormatterCacheEnabled.Should().BeTrue();
        config.DaxFormatterCacheMaxEntries.Should().Be(100);
        config.DaxFormatterCacheTtlMinutes.Should().Be(60);
        _errorOutput.ToString().Should().BeEmpty(); // No warnings
    }

    #endregion

    #region Current Instance

    [Fact]
    public void Load_SetsCurrentInstance()
    {
        // Act
        var config = McpConfiguration.Load();

        // Assert
        McpConfiguration.Current.Should().BeSameAs(config);
    }

    #endregion

    #region String Environment Variables

    [Fact]
    public void Load_ValidServerName_AppliesValue()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_SERVER_NAME", "custom-server-name");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.ServerName.Should().Be("custom-server-name");
    }

    [Fact]
    public void Load_ValidServerVersion_AppliesValue()
    {
        // Arrange
        Environment.SetEnvironmentVariable("MCP_SERVER_VERSION", "9.9.9");

        // Act
        var config = McpConfiguration.Load();

        // Assert
        config.ServerVersion.Should().Be("9.9.9");
    }

    #endregion
}
