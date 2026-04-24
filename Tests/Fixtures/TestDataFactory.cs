using ExcelPowerPivotMcp.Common.DataStructures;
using ExcelPowerPivotMcp.Common.DataStructures.Metadata;
using System.Collections.Generic;

namespace ExcelPowerPivotMcp.Tests.Fixtures;

/// <summary>
/// Factory for creating test data objects with sensible defaults.
/// Reduces test boilerplate and ensures consistent test data.
/// </summary>
public static class TestDataFactory
{
    public static MeasureCreate CreateMeasureCreate(
        string tableName = "TestTable",
        string measureName = "TestMeasure",
        string expression = "SUM([Amount])",
        string? description = null,
        string? formatString = null)
    {
        return new MeasureCreate
        {
            TableName = tableName,
            MeasureName = measureName,
            Expression = expression,
            Description = description,
            FormatString = formatString
        };
    }

    public static MeasureUpdate CreateMeasureUpdate(
        string measureName = "TestMeasure",
        string? expression = null,
        string? newName = null,
        string? description = null)
    {
        return new MeasureUpdate
        {
            MeasureName = measureName,
            Expression = expression,
            NewName = newName,
            Description = description
        };
    }

    public static MeasureMetadata CreateMeasureMetadata(
        string name = "TestMeasure",
        string tableName = "TestTable",
        string expression = "SUM([Amount])")
    {
        return new MeasureMetadata
        {
            Name = name,
            TableName = tableName,
            Expression = expression,
            DataType = "Integer",
            IsVisible = true
        };
    }

    public static RelationshipCreate CreateRelationshipCreate(
        string foreignTable = "Sales",
        string foreignColumn = "ProductKey",
        string primaryTable = "Product",
        string primaryColumn = "ProductKey")
    {
        return new RelationshipCreate
        {
            ForeignTable = foreignTable,
            ForeignColumn = foreignColumn,
            PrimaryTable = primaryTable,
            PrimaryColumn = primaryColumn
        };
    }

    public static RelationshipMetadata CreateRelationshipMetadata(
        string foreignTable = "Sales",
        string foreignColumn = "ProductKey",
        string primaryTable = "Product",
        string primaryColumn = "ProductKey",
        bool isActive = true)
    {
        return new RelationshipMetadata
        {
            ForeignKeyTable = foreignTable,
            ForeignKeyColumn = foreignColumn,
            PrimaryKeyTable = primaryTable,
            PrimaryKeyColumn = primaryColumn,
            IsActive = isActive
        };
    }

    public static TableMetadata CreateTableMetadata(
        string name = "TestTable",
        long rowCount = 1000)
    {
        return new TableMetadata
        {
            Name = name,
            IsVisible = true,
            RowCount = rowCount
        };
    }

    public static ColumnMetadata CreateColumnMetadata(
        string name = "TestColumn",
        string tableName = "TestTable",
        string dataType = "String")
    {
        return new ColumnMetadata
        {
            Name = name,
            TableName = tableName,
            DataType = dataType,
            IsHidden = false,
            IsCalculated = false
        };
    }

    public static ColumnProfile CreateColumnProfile(
        long count = 100,
        long distinctCount = 50,
        long nullCount = 5)
    {
        return new ColumnProfile
        {
            Count = count,
            DistinctCount = distinctCount,
            NullCount = nullCount,
            Min = 1,
            Max = 100,
            DistinctValues = new List<object?> { 1, 2, 3 }
        };
    }

    public static TableAddToModel CreateTableAddToModel(
        string tableName = "TestTable")
    {
        return new TableAddToModel
        {
            TableName = tableName
        };
    }
}
