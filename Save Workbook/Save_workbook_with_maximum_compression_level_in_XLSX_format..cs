using System;
using Aspose.Cells;

class SaveWithMaxCompression
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, Aspose!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Create OOXML save options
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();

        // Set the compression type to the maximum level (Level9)
        saveOptions.CompressionType = OoxmlCompressionType.Level9;

        // Save the workbook as XLSX using the specified compression options
        workbook.Save("MaxCompressed.xlsx", saveOptions);
    }
}