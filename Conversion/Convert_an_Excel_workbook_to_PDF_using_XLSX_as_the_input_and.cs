using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the output PDF file
            string destPath = "output.pdf";

            // Load the workbook and save it as PDF
            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}