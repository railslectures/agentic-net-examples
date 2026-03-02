using System;
using Aspose.Cells;

namespace AsposeCellsStrictSaveDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Hello");
            worksheet.Cells["B1"].PutValue("Strict OOXML");

            // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Save the workbook in XLSX format with strict compliance
            workbook.Save("StrictWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}