using System;
using System.Runtime.InteropServices;

namespace ExcelPowerPivotMcp.PowerPivot;

/// <summary>
/// PowerPivotConnection partial class - Discovery operations.
/// Contains methods for discovering Excel instances and workbooks.
/// </summary>
public partial class PowerPivotConnection
{
    /// <summary>
    /// Get the running Excel application via COM
    /// </summary>
    private static object? GetExcelApplication()
    {
        try
        {
            // Excel.Application CLSID
            var excelClsid = new Guid("00024500-0000-0000-C000-000000000046");
            GetActiveObject(ref excelClsid, IntPtr.Zero, out object excelApp);
            return excelApp;
        }
        catch (COMException)
        {
            return null;
        }
    }

    /// <summary>
    /// Discover Excel workbooks that have Power Pivot data models
    /// </summary>
    public List<PowerPivotWorkbook> DiscoverWorkbooks()
    {
        var workbooks = new List<PowerPivotWorkbook>();

        try
        {
            using var scope = new ComScope();
            
            var excelApp = GetExcelApplication();
            if (excelApp == null) return workbooks;
            
            scope.TrackDynamic(excelApp);

            dynamic dynExcel = excelApp;
            var workbooksCollection = scope.TrackDynamic(dynExcel.Workbooks);
            if (workbooksCollection == null) return workbooks;

            int count = ExcelComHelpers.GetCount(workbooksCollection);

            for (int i = 1; i <= count; i++)
            {
                dynamic workbook = scope.TrackDynamic(workbooksCollection.Item[i]);
                if (workbook == null) continue;

                string name = workbook.Name;
                string fullPath = workbook.FullName;
                bool hasModel = HasDataModel(workbook);

                workbooks.Add(new PowerPivotWorkbook(name, fullPath, hasModel));
            }
        }
        catch (Exception ex)
        {
            // Log error but don't throw
            Console.Error.WriteLine($"Error discovering workbooks: {ex.Message}");
        }

        return workbooks;
    }

    /// <summary>
    /// Check if a workbook has a data model
    /// </summary>
    private static bool HasDataModel(object workbook)
    {
        try
        {
            using var scope = new ComScope();
            dynamic dynWorkbook = workbook;
            var model = scope.TrackDynamic(dynWorkbook.Model);
            return model != null;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"HasDataModel check failed: {ex.Message}");
            return false;
        }
    }
    /// <summary>
    /// Refresh the data model (equivalent to Data > Refresh All in Excel)
    /// </summary>
    public void RefreshModel()
    {
        ExecuteWithConnectionValidation(() =>
        {
            dynamic dynWorkbook = _workbook!;
            // RefreshAll refreshes all data connections in the workbook
            dynWorkbook.RefreshAll();
        }, "refreshing model");
    }
}
