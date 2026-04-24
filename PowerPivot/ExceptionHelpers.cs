using System;

namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// Helper methods for consistent exception handling across the codebase.
/// Provides standardized wrapping patterns for COM and other operations.
/// </summary>
public static class ExceptionHelpers
{
    private static string TranslateExceptionMessage(Exception ex)
    {
        string msg = ex.Message;
        
        // Improve common DAX/OleDb error codes
        if (msg.Contains("0xC113000A")) 
            return $"{msg} (Hint: The table was not found or the schema has changed. Use get_model_summary to check structure.)";
        if (msg.Contains("0x80040E14")) 
            return $"{msg} (Hint: Syntax error in DAX expression. Check columns and measures exist.)";
        if (msg.Contains("0x80004005")) 
            return $"{msg} (Hint: Unspecified error. Check if Excel is in edit mode or a dialog is open.)";

        return msg;
    }

    /// <summary>
    /// Execute an action with standard exception wrapping.
    /// InvalidOperationException and ArgumentException pass through unchanged.
    /// Other exceptions are wrapped in InvalidOperationException with context.
    /// </summary>
    /// <param name="action">The action to execute</param>
    /// <param name="context">Context for error message (e.g., "creating measure")</param>
    public static void WrapExceptions(Action action, string context)
    {
        try
        {
            action();
        }
        catch (Exception ex) when (ex is not InvalidOperationException && ex is not ArgumentException)
        {
            throw new InvalidOperationException($"Error {context}: {TranslateExceptionMessage(ex)}", ex);
        }
    }

    /// <summary>
    /// Execute a function with standard exception wrapping and return its result.
    /// InvalidOperationException and ArgumentException pass through unchanged.
    /// Other exceptions are wrapped in InvalidOperationException with context.
    /// </summary>
    /// <typeparam name="T">Return type</typeparam>
    /// <param name="func">The function to execute</param>
    /// <param name="context">Context for error message (e.g., "executing DAX query")</param>
    /// <returns>The function result</returns>
    public static T WrapExceptions<T>(Func<T> func, string context)
    {
        try
        {
            return func();
        }
        catch (Exception ex) when (ex is not InvalidOperationException && ex is not ArgumentException)
        {
            throw new InvalidOperationException($"Error {context}: {TranslateExceptionMessage(ex)}", ex);
        }
    }

    /// <summary>
    /// Wrap a COM-specific operation with appropriate error handling.
    /// Provides more detailed context for COM failures.
    /// </summary>
    /// <param name="action">The COM action to execute</param>
    /// <param name="operation">Description of the operation (e.g., "accessing Excel workbook")</param>
    public static void WrapComOperation(Action action, string operation)
    {
        try
        {
            action();
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            string details = TranslateExceptionMessage(ex);
            throw new InvalidOperationException(
                $"COM error during {operation}. Excel may be busy or the object is no longer valid. Details: {details}", 
                ex);
        }
        catch (Exception ex) when (ex is not InvalidOperationException && ex is not ArgumentException)
        {
            throw new InvalidOperationException($"Error {operation}: {TranslateExceptionMessage(ex)}", ex);
        }
    }

    /// <summary>
    /// Wrap a COM-specific operation with appropriate error handling and return result.
    /// </summary>
    /// <typeparam name="T">Return type</typeparam>
    /// <param name="func">The COM function to execute</param>
    /// <param name="operation">Description of the operation</param>
    /// <returns>The function result</returns>
    public static T WrapComOperation<T>(Func<T> func, string operation)
    {
        try
        {
            return func();
        }
        catch (System.Runtime.InteropServices.COMException ex)
        {
            string details = TranslateExceptionMessage(ex);
            throw new InvalidOperationException(
                $"COM error during {operation}. Excel may be busy or the object is no longer valid. Details: {details}", 
                ex);
        }
        catch (Exception ex) when (ex is not InvalidOperationException && ex is not ArgumentException)
        {
            throw new InvalidOperationException($"Error {operation}: {TranslateExceptionMessage(ex)}", ex);
        }
    }
}
