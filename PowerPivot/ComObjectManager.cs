using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics.CodeAnalysis;

namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// Manages COM object lifecycle to prevent memory leaks.
/// Tracks COM objects and releases them in reverse order when disposed.
/// </summary>
/// <example>
/// <code>
/// using var scope = new ComScope();
/// var excelApp = scope.TrackDynamic(GetExcelApplication());
/// var workbooks = scope.TrackDynamic(excelApp.Workbooks);
/// // ... work with COM objects
/// // All tracked objects are released when scope is disposed
/// </code>
/// </example>
public sealed class ComScope : IDisposable

{
    private readonly Stack<object> _comObjects = new();
    private bool _disposed;

    /// <summary>
    /// Track a COM object for later release.
    /// Returns the same object for fluent usage.
    /// </summary>
    /// <remarks>
    /// Returns the same object that was passed in. If null is passed, null is returned.
    /// The generic constraint ensures the compiler knows T is a reference type.
    /// </remarks>
    [return: NotNullIfNotNull(nameof(comObject))]
    public T? Track<T>(T? comObject) where T : class
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            _comObjects.Push(comObject);
        }
        return comObject;
    }


    /// <summary>
    /// Track a dynamic COM object for later release.
    /// </summary>
    /// <remarks>
    /// Returns the same object that was passed in. If null is passed, null is returned.
    /// </remarks>
    [return: NotNullIfNotNull(nameof(comObject))]
    public dynamic TrackDynamic(dynamic comObject)
    {
        if (comObject != null)
        {
            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    _comObjects.Push(comObject);
                }
            }
            catch (InvalidCastException)
            {
                // Expected when checking dynamic objects that may not be COM
                // Track it anyway to be safe
                _comObjects.Push(comObject);
            }
            catch (NotSupportedException)
            {
                // Can occur with certain dynamic proxy objects
                // Track it anyway to be safe
                _comObjects.Push(comObject);
            }
        }
        return comObject!;
    }


    /// <summary>
    /// Safely access a COM property and track the result.
    /// If the property access throws, nothing is leaked.
    /// If tracking fails after access, the object is released.
    /// </summary>
    /// <param name="accessor">Lambda that accesses the COM property</param>
    /// <returns>The tracked COM object</returns>
    /// <example>
    /// // Instead of: dynamic model = scope.TrackDynamic(workbook.Model);
    /// // Use: dynamic model = scope.TrackProperty(() => workbook.Model);
    /// </example>
    public dynamic TrackProperty(Func<dynamic> accessor)
    {
        dynamic? result = null;
        try
        {
            result = accessor();
            if (result != null)
            {
                try
                {
                    if (Marshal.IsComObject(result))
                    {
                        _comObjects.Push(result);
                    }
                }
                catch (InvalidCastException)
                {
                    // Expected when checking dynamic objects that may not be COM
                    // Track it anyway to be safe
                    _comObjects.Push(result);
                }
                catch (NotSupportedException)
                {
                    // Can occur with certain dynamic proxy objects
                    // Track it anyway to be safe
                    _comObjects.Push(result);
                }
            }
            return result!;
        }
        catch
        {
            // If accessor threw after partially succeeding, clean up
            if (result != null)
            {
                try
                {
                    if (Marshal.IsComObject(result))
                    {
                        Marshal.ReleaseComObject(result);
                    }
                }
                catch (InvalidCastException)
                {
                    // Ignore - not a COM object
                }
                catch (NotSupportedException)
                {
                    // Ignore - proxy object doesn't support release
                }
            }
            throw;
        }
    }


    /// <summary>
    /// Release all tracked COM objects in reverse order.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;

        while (_comObjects.Count > 0)
        {
            var obj = _comObjects.Pop();
            if (obj != null)
            {
                try
                {
                    if (Marshal.IsComObject(obj))
                    {
                        Marshal.ReleaseComObject(obj);
                    }
                }
                catch (InvalidCastException)
                {
                    // Ignore - not a COM object
                }
                catch (NotSupportedException)
                {
                    // Ignore - proxy object doesn't support release
                }
            }
        }
    }
}

/// <summary>
/// Static helper for one-off COM object releases.
/// </summary>
public static class ComHelper
{
    /// <summary>
    /// Safely release a COM object if it is one.
    /// </summary>
    public static void SafeRelease(object? comObject)
    {
        if (comObject != null)
        {
            try
            {
                if (Marshal.IsComObject(comObject))
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (InvalidCastException)
            {
                // Ignore - not a COM object
            }
            catch (NotSupportedException)
            {
                // Ignore - proxy object doesn't support release
            }
        }
    }

}
