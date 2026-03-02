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
        sheet.Cells["A1"].PutValue("Compression Test");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create OOXML save options and set maximum compression (Level9)
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level9;

        // Save the workbook as XLSX with the specified compression options
        workbook.Save("CompressedWorkbook.xlsx", saveOptions);
    }
}