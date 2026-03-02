using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook (XLSX)
            string sourcePath = "input.xlsx";

            // Convert to OpenDocument Spreadsheet (ODS)
            string odsPath = "output.ods";
            ConversionUtility.Convert(sourcePath, odsPath);
            Console.WriteLine($"Converted to ODS: {odsPath}");

            // Convert to StarOffice Calc Spreadsheet (SXC)
            string sxcPath = "output.sxc";
            ConversionUtility.Convert(sourcePath, sxcPath);
            Console.WriteLine($"Converted to SXC: {sxcPath}");

            // Convert to OpenDocument Flat XML Spreadsheet (FODS)
            string fodsPath = "output.fods";
            ConversionUtility.Convert(sourcePath, fodsPath);
            Console.WriteLine($"Converted to FODS: {fodsPath}");
        }
    }
}