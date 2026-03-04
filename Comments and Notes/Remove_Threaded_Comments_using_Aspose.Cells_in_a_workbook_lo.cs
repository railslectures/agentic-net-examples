using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and clear all comments (including threaded comments)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clears all comments in the current worksheet
                sheet.ClearComments();
            }

            // Save the workbook after removing threaded comments
            string outputPath = "OutputWorkbook_NoThreadedComments.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Threaded comments have been removed and workbook saved to: " + outputPath);
        }
    }
}