using System;
using Aspose.Cells;

namespace RemoveThreadedCommentsDemo
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
                // This method removes every comment attached to the worksheet,
                // which also clears any threaded comment chains.
                sheet.ClearComments();
            }

            // Save the workbook after removing the comments
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}