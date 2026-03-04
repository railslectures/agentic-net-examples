using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output PDF file path
        string destPath = "output.pdf";

        // Convert the Excel workbook to PDF using the provided utility method
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}