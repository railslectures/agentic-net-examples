using System;
using Aspose.Cells.Utility;

class ConvertXlsxToOds
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Desired path for the output ODS file
        string destPath = "output.ods";

        // Convert the XLSX workbook to ODS format using Aspose.Cells ConversionUtility
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("Conversion from XLSX to ODS completed successfully.");
    }
}