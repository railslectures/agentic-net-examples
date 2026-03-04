using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired path for the output XPS file
        string destPath = "output.xps";

        // Convert the Excel workbook to XPS format
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}