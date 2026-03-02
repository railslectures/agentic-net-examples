using System;
using Aspose.Cells;

class SaveStrictOpenXml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
        workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

        // Save the workbook in XLSX format using default options
        workbook.Save("StrictWorkbook.xlsx", SaveFormat.Xlsx);
    }
}