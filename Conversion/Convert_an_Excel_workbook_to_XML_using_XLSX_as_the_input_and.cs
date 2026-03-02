using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExcelToXmlConverter
    {
        public static void Run()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the resulting XML file
            string outputPath = "output.xml";

            // Load the Excel workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Create XML save options – you can customize options as needed
            XmlSaveOptions saveOptions = new XmlSaveOptions
            {
                SheetNameAsElementName = true,
                HasHeaderRow = true
            };

            // Save the workbook as an XML file using the specified options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Excel workbook '{sourcePath}' has been successfully converted to XML at '{outputPath}'.");
        }

        // Entry point
        public static void Main(string[] args)
        {
            Run();
        }
    }
}