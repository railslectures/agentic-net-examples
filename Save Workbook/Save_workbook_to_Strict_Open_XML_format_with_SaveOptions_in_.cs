using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello Strict OOXML");

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Create OOXML save options for XLSX format
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        // Example: set compression level (optional)
        saveOptions.CompressionType = OoxmlCompressionType.Level6;

        // Save the workbook as a strict OOXML XLSX file using the save options
        workbook.Save("StrictWorkbook.xlsx", saveOptions);
    }
}