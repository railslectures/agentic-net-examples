using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string inputPath = "input.xlsx";

        // Path where the workbook without threaded comments will be saved
        string outputPath = "output.xlsx";

        // Load the workbook (XLSX format)
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and clear all comments (including threaded comments)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // This method removes both regular and threaded comments from the worksheet
            sheet.ClearComments();
        }

        // Save the modified workbook
        workbook.Save(outputPath, SaveFormat.Xlsx);
    }
}