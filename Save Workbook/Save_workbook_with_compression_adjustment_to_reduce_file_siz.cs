using System;
using Aspose.Cells;

class CompressionDemo
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Compression Test");
        worksheet.Cells["A2"].PutValue(DateTime.Now);

        // Configure save options with the highest compression level to reduce file size
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.CompressionType = OoxmlCompressionType.Level9; // maximum compression

        // Save the workbook as XLSX using the configured compression options
        workbook.Save("CompressedDemo.xlsx", saveOptions);
    }
}