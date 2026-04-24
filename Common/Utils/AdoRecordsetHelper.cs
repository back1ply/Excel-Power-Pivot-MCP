using System;
using System.Collections.Generic;

namespace ExcelPowerPivotMcp.Common.Utils;

/// <summary>
/// Helper utility for extracting data from ADO recordsets.
/// Reduces code duplication between AdoDaxService and AdoDmvService.
/// </summary>
public static class AdoRecordsetHelper
{
    /// <summary>
    /// Extract results from an ADO recordset into a list of dictionaries.
    /// </summary>
    /// <param name="recordset">Dynamic ADO recordset object</param>
    /// <param name="maxRows">Maximum number of rows to return (0 = unlimited)</param>
    /// <param name="useOrdinalIgnoreCase">Use OrdinalIgnoreCase string comparer for dictionary keys</param>
    /// <returns>List of dictionaries representing rows</returns>
    public static List<Dictionary<string, object?>> ExtractResults(
        dynamic recordset,
        int maxRows = 0,
        bool useOrdinalIgnoreCase = false)
    {
        var results = new List<Dictionary<string, object?>>();

        if (recordset == null) return results;

        // Check for EOF (empty recordset)
        if ((bool)recordset.EOF) return results;

        // Get field count and names
        dynamic fields = recordset.Fields;
        int fieldCount = (int)fields.Count;

        var colNames = new List<string>();
        for (int i = 0; i < fieldCount; i++)
        {
            colNames.Add((string)fields[i].Name);
        }

        // Create string comparer based on parameter
        var comparer = useOrdinalIgnoreCase
            ? StringComparer.OrdinalIgnoreCase
            : StringComparer.Ordinal;

        // Read data by iterating recordset
        recordset.MoveFirst();
        int count = 0;

        while (!(bool)recordset.EOF)
        {
            // Check max rows limit (if specified)
            if (maxRows > 0 && count >= maxRows) break;

            var row = new Dictionary<string, object?>(comparer);
            for (int i = 0; i < fieldCount; i++)
            {
                object? value = fields[i].Value;
                // Convert DBNull to null for consistency
                row[colNames[i]] = value == DBNull.Value ? null : value;
            }
            results.Add(row);
            count++;

            recordset.MoveNext();
        }

        return results;
    }
}
