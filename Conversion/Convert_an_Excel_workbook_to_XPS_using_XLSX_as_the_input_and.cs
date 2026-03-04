using System;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourcePath = "input.xlsx";

        // Desired output path for the XPS file
        string destPath = "output.xps";

        // Convert the Excel workbook to XPS format using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Excel file has been successfully converted to XPS.");
    }
}