using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.ClearComments(); // Removes both legacy and threaded comments from the sheet
            }

            // Save the workbook after removing the comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}