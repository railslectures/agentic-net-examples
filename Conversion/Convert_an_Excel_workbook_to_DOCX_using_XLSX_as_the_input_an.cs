using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Desired output path with DOCX extension
            string destPath = "output.docx";

            // Convert the Excel workbook to DOCX using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine("Conversion completed successfully.");
        }
    }
}