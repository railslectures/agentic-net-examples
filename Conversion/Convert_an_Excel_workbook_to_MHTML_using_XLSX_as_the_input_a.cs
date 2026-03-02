using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired output path with .mht extension (MHTML format)
        string destPath = "output.mht";

        // Convert the Excel workbook to MHTML using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion completed successfully.");
    }
}