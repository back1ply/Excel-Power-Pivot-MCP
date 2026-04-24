using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;

namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// Helper methods for working with Excel COM collections.
/// These helpers reduce boilerplate when iterating 1-indexed COM collections.
/// </summary>
public static class ExcelComHelpers
{
    /// <summary>
    /// Find an item in a 1-indexed COM collection by its Name property.
    /// Uses case-insensitive comparison and handles DAX-style delimiters.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <param name="name">The name to search for</param>
    /// <returns>The matching item, or null if not found</returns>
    public static dynamic? FindByName(dynamic collection, string name)
    {
        int count = collection.Count;

        // Clean the search name of DAX delimiters (brackets and single quotes)
        string cleanSearchName = name.TrimStart('[', '\'').TrimEnd(']', '\'');

        for (int i = 1; i <= count; i++)
        {
            dynamic item = collection.Item[i];
            string itemName = item.Name;

            // Direct match
            if (itemName.Equals(name, StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }

            // Match cleaned name (handles '[TableName]', "'TableName'", or "TableName")
            if (itemName.Equals(cleanSearchName, StringComparison.OrdinalIgnoreCase))
            {
                return item;
            }
        }
        return null;
    }

    /// <summary>
    /// Find an item in a 1-indexed COM collection by its Name property.
    /// Throws ArgumentException if not found.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <param name="name">The name to search for</param>
    /// <param name="itemType">Type description for error message (e.g., "Table", "Measure")</param>
    /// <returns>The matching item</returns>
    /// <exception cref="ArgumentException">Thrown when item is not found</exception>
    public static dynamic FindByNameOrThrow(dynamic collection, string name, string itemType)
    {
        var result = FindByName(collection, name);
        if (result == null)
        {
            // Actionable error message - tells AI what to do next
            HashSet<string> availableNames = GetAllNames(collection);
            var suggestion = availableNames.Count > 0 
                ? $" Available: {string.Join(", ", availableNames.Cast<string>().Take(5))}{(availableNames.Count > 5 ? "..." : "")}"
                : " No items found in collection.";
            throw new ArgumentException($"{itemType} '{name}' not found.{suggestion} Use list_tables or list_measures to see available items.");
        }
        return result;
    }

    /// <summary>
    /// Try to find an item in a 1-indexed COM collection by its Name property.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <param name="name">The name to search for</param>
    /// <param name="result">The matching item if found</param>
    /// <returns>True if found, false otherwise</returns>
    public static bool TryFindByName(dynamic collection, string name, out dynamic? result)
    {
        result = FindByName(collection, name);
        return result != null;
    }

    /// <summary>
    /// Check if an item with the given name exists in a 1-indexed COM collection.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <param name="name">The name to search for</param>
    /// <returns>True if an item with that name exists</returns>
    public static bool ContainsName(dynamic collection, string name)
    {
        return FindByName(collection, name) != null;
    }

    /// <summary>
    /// Enumerate all items in a 1-indexed COM collection.
    /// Note: Caller should handle COM object cleanup if needed.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <returns>Enumerable of all items in the collection</returns>
    public static IEnumerable<dynamic> EnumerateCollection(dynamic collection)
    {
        int count = collection.Count;
        for (int i = 1; i <= count; i++)
        {
            yield return collection.Item[i];
        }
    }

    /// <summary>
    /// Enumerate all items in a 1-indexed COM collection with automatic cleanup.
    /// Each COM object is released after the action is executed.
    /// Use this when you don't need to hold references to items after processing.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <param name="action">Action to perform on each item</param>
    public static void EnumerateCollectionWithCleanup(dynamic collection, Action<dynamic> action)
    {
        int count = collection.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic item = collection.Item[i];
            try
            {
                action(item);
            }
            finally
            {
                ComHelper.SafeRelease(item);
            }
        }
    }

    /// <summary>
    /// Enumerate all items in a 1-indexed COM collection with automatic cleanup and result collection.
    /// Each COM object is released after the transform is executed.
    /// </summary>
    /// <typeparam name="T">Result type</typeparam>
    /// <param name="collection">A COM collection with Count and Item[index] properties</param>
    /// <param name="transform">Transform function applied to each item</param>
    /// <returns>List of transformed results</returns>
    public static List<T> EnumerateCollectionWithCleanup<T>(dynamic collection, Func<dynamic, T> transform)
    {
        var results = new List<T>();
        int count = collection.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic item = collection.Item[i];
            try
            {
                results.Add(transform(item));
            }
            finally
            {
                ComHelper.SafeRelease(item);
            }
        }
        return results;
    }

    /// <summary>
    /// Get all names from a 1-indexed COM collection.
    /// </summary>
    /// <param name="collection">A COM collection with Count and Item[index] properties where items have Name</param>
    /// <returns>HashSet of names (case-insensitive)</returns>
    public static HashSet<string> GetAllNames(dynamic collection)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int count = collection.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic item = collection.Item[i];
            try
            {
                names.Add((string)item.Name);
            }
            finally
            {
                ComHelper.SafeRelease(item);
            }
        }
        return names;
    }

    /// <summary>
    /// Get the count of items in a COM collection.
    /// </summary>
    public static int GetCount(dynamic collection)
    {
        return (int)collection.Count;
    }
    
    public static List<Dictionary<string, object?>> GetExcelTables(dynamic workbook, ILogger? logger = null)
    {
        var tables = new List<Dictionary<string, object?>>();
        
        // Scope ensures all tracked COM objects are released 
        using var scope = new ComScope();
        
        try
        {
            // Track intermediate objects to avoid leaks
            // "One Dot Rule": Assign every COM interaction to a variable and track it
            
            dynamic model = scope.TrackDynamic(workbook.Model);
            dynamic modelTables = scope.TrackDynamic(model.ModelTables);
            var modelTableNames = GetAllNames(modelTables);

            dynamic worksheets = scope.TrackDynamic(workbook.Worksheets);
            int sheetCount = worksheets.Count;

            // Manual loop for ComScope tracking (EnumerateCollection yields untracked objects)
            for (int i = 1; i <= sheetCount; i++)
            {
                dynamic sheet = scope.TrackDynamic(worksheets.Item[i]);
                string sheetName = sheet.Name;
                
                dynamic listObjects = scope.TrackDynamic(sheet.ListObjects);
                int listCount = listObjects.Count;

                for (int j = 1; j <= listCount; j++)
                {
                    dynamic listObject = scope.TrackDynamic(listObjects.Item[j]);
                    string tableName = listObject.Name;
                    
                    dynamic range = scope.TrackDynamic(listObject.Range);
                    dynamic rows = scope.TrackDynamic(range.Rows);
                    dynamic columns = scope.TrackDynamic(range.Columns);
                    
                    string address = range.Address;
                    int rowCount = rows.Count - 1; 
                    int colCount = columns.Count;

                    tables.Add(new Dictionary<string, object?>
                    {
                        ["name"] = tableName,
                        ["worksheet"] = sheetName,
                        ["address"] = address,
                        ["rowCount"] = rowCount,
                        ["columnCount"] = colCount,
                        ["isInModel"] = modelTableNames.Contains(tableName)
                    });
                }
            }
        }
        catch (COMException comEx)
        {
            // COM errors accessing worksheets or tables - log warning and return partial results
            logger?.LogWarning(comEx, "COM error while enumerating Excel tables: {HResult:X8}", comEx.HResult);
            // Return what we've collected so far - graceful degradation
        }
        catch (InvalidCastException ex)
        {
            // Expected when workbook structure is unexpected (e.g., protected sheets)
            logger?.LogWarning(ex, "Type conversion error accessing workbook structure");
        }
        catch (NullReferenceException ex)
        {
            // Expected when optional properties are missing
            logger?.LogWarning(ex, "Null reference accessing workbook structure");
        }
        // Let critical exceptions (OutOfMemoryException, etc.) propagate

        return tables;
    }
}

