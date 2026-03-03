using System;
using Aspose.Cells;

class RemoveThreadedComments
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all comments (including threaded comments) from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Save the workbook after comments have been removed
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}