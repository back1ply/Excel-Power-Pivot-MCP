using Xunit;
using ExcelPowerPivotMcp.Common.Utils;

namespace ExcelPowerPivotMcp.Tests.Common.Utils;

public class DaxHelpersTests
{
    [Theory]
    [InlineData("TableName", "[TableName]")]
    [InlineData("Table Name", "[Table Name]")]
    [InlineData("[AlreadyBracketed]", "[AlreadyBracketed]")]
    public void SanitizeTableName_AddsBracketsCorrectly(string input, string expected)
    {
        // Act
        var result = DaxHelpers.SanitizeTableName(input);

        // Assert
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("Simple", "'Table'[Simple]")]
    [InlineData("With Space", "'Table'[With Space]")]
    [InlineData("With'Quote", "'Table'[With''Quote]")]
    public void BuildColumnReference_EscapesCorrectly(string column, string expectedSuffix)
    {
        // Arrange
        string table = "Table";

        // Act
        var result = DaxHelpers.BuildColumnReference(table, column);

        // Assert
        Assert.EndsWith(expectedSuffix, result);
    }
}
