using System;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Desired output path for the XPS file
            string outputPath = "output.xps";

            // Convert the Excel workbook to XPS using Aspose.Cells ConversionUtility
            ConversionUtility.Convert(sourcePath, outputPath);

            Console.WriteLine("Excel file has been successfully converted to XPS.");
        }
    }
}