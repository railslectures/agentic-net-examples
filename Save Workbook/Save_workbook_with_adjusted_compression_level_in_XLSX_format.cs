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

        // Configure OOXML save options with the desired compression level
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level9; // maximum compression

        // Save the workbook as XLSX using the specified compression options
        workbook.Save("CompressedWorkbook.xlsx", saveOptions);
    }
}