using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToHtmlDemo
    {
        public static void Main()
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path where the HTML output will be saved
            string destPath = "output.html";

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Save the workbook as HTML
            workbook.Save(destPath, SaveFormat.Html);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{destPath}'");
        }
    }
}