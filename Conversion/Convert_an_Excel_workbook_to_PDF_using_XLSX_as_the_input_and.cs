using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    public class ExcelToPdfConverter
    {
        public static void Main()
        {
            // Path to the source Excel file (XLSX format)
            string sourcePath = "input.xlsx";

            // Desired output PDF file path
            string destPath = "output.pdf";

            // Convert the Excel workbook to PDF using Aspose.Cells ConversionUtility
            // This utilizes the provided Convert(string source, string saveAs) method.
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}