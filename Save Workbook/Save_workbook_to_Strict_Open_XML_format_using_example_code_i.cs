using System;
using Aspose.Cells;

namespace AsposeCellsStrictOoxmlDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Set the OOXML compliance level to ISO/IEC 29500:2008 Strict
            workbook.Settings.Compliance = OoxmlCompliance.Iso29500_2008_Strict;

            // Add some data to the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Hello, Strict OOXML!");

            // Save the workbook in XLSX format (the compliance setting will be applied)
            workbook.Save("StrictComplianceDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}