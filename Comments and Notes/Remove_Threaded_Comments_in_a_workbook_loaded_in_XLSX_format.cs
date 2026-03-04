using System;
using Aspose.Cells;

class RemoveThreadedCommentsDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and clear their comments
        // Worksheet.ClearComments removes both regular and threaded comments
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Save the workbook after comments have been removed
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}