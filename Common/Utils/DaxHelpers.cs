using System;

namespace ExcelPowerPivotMcp.Common.Utils;

public static class DaxHelpers
{
    /// <summary>
    /// ADO/OLEDB data type codes returned by DMV queries (DBTYPE values).
    /// </summary>
    public static class DmvDataTypes
    {
        public const int DbInteger = 1;        // DBTYPE_I4
        public const int DbReal = 2;           // DBTYPE_R8
        public const int DbCurrency = 3;       // DBTYPE_CY
        public const int DbBoolean = 8;        // DBTYPE_BOOL
        public const int DbBigInt = 20;        // DBTYPE_I8
        public const int DbWideString = 130;   // DBTYPE_WSTR
        public const int DbDecimal = 131;      // DBTYPE_NUMERIC
        public const int DbDateTime = 135;     // DBTYPE_DBTIMESTAMP
    }

    public static string GetDataTypeName(int dataType)
    {
        return dataType switch
        {
            DmvDataTypes.DbInteger => "Integer",
            DmvDataTypes.DbReal => "Real",
            DmvDataTypes.DbCurrency => "Currency",
            DmvDataTypes.DbBoolean => "Boolean",
            DmvDataTypes.DbBigInt => "Int64",
            DmvDataTypes.DbWideString => "String",
            DmvDataTypes.DbDecimal => "Decimal",
            DmvDataTypes.DbDateTime => "DateTime",
            _ => $"Unknown ({dataType})"
        };
    }
    
    public static string SanitizeValue(string value)
    {
        return value.Replace("'", "''");
    }

    public static string SanitizeTableName(string tableName)
    {
        var name = tableName.StartsWith('[') ? tableName : $"[{tableName}]";
        return SanitizeValue(name);
    }

    public static string EscapeLikeWildcards(string value)
    {
        return value.Replace("%", "[%]").Replace("_", "[_]");
    }
    
    /// <summary>
    /// Build a properly formatted DAX column reference from table and column names.
    /// Handles escaping and proper bracket/quote formatting.
    /// </summary>
    /// <param name="tableName">Table name (will be cleaned of existing brackets/quotes)</param>
    /// <param name="columnName">Column name</param>
    /// <returns>Formatted reference like 'TableName'[ColumnName]</returns>
    public static string BuildColumnReference(string tableName, string columnName)
    {
        var cleanTable = tableName.Trim('[', ']', '\'');
        var cleanColumn = SanitizeValue(columnName);
        return $"'{cleanTable}'[{cleanColumn}]";
    }

}
