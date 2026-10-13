
public void Main()
{
    Excel.Application excelApp = null;
    Excel.Workbook workbook = null;
    Excel.Worksheet worksheet = null;

    try
    {
        // Read the dynamic Excel file path from SSIS
        string filePath =
            (string)Dts.Variables["User::ExcelFilePath"].Value;

        // Open Excel
        excelApp = new Excel.Application();
        excelApp.DisplayAlerts = false;

        workbook = excelApp.Workbooks.Open(filePath);

        // Select the first worksheet
        worksheet = (Excel.Worksheet)workbook.Worksheets[1];

        // Delete the second row, preserving the header
        Excel.Range secondRow =
            (Excel.Range)worksheet.Rows[2];

        secondRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

        // Save the changes
        workbook.Save();
        workbook.Close(false);

        Dts.TaskResult = (int)ScriptResults.Success;
    }
    catch (Exception ex)
    {
        Dts.Events.FireError(
            0,
            "Delete Second Excel Row",
            ex.ToString(),
            "",
            0);

        if (workbook != null)
            workbook.Close(false);

        Dts.TaskResult = (int)ScriptResults.Failure;
    }
    finally
    {
        if (worksheet != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);

        if (workbook != null)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);

        if (excelApp != null)
        {
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
        }
    }
}
