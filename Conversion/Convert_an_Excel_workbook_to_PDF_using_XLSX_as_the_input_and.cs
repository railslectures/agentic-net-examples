using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertExcelToPdf
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path where the PDF will be saved
            string destPath = "output.pdf";

            // Load the workbook and save it as PDF
            Workbook workbook = new Workbook(sourcePath);
            workbook.Save(destPath, SaveFormat.Pdf);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}