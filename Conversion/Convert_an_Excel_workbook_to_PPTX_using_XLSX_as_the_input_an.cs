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

            // Desired output PowerPoint file (PPTX)
            string destPath = "output.pptx";

            try
            {
                // Convert the Excel workbook to PPTX using the built‑in conversion utility
                ConversionUtility.Convert(sourcePath, destPath);

                Console.WriteLine($"Conversion successful: '{sourcePath}' → '{destPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}