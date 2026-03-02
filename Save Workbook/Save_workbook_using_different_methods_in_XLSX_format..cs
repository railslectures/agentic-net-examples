using System;
using System.IO;
using Aspose.Cells;

public class SaveWorkbookDemo
{
    public static void Run()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // -------------------------------------------------
        // 1. Save using only the file name (extension determines format)
        // -------------------------------------------------
        workbook.Save("Method1.xlsx"); // Saves as XLSX because of the .xlsx extension

        // -------------------------------------------------
        // 2. Save using file name and explicit SaveFormat enumeration
        // -------------------------------------------------
        workbook.Save("Method2.xlsx", SaveFormat.Xlsx);

        // -------------------------------------------------
        // 3. Save to a MemoryStream specifying the SaveFormat
        // -------------------------------------------------
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);
            // Optionally write the stream content to a physical file for verification
            File.WriteAllBytes("Method3_FromStream.xlsx", stream.ToArray());
        }

        // -------------------------------------------------
        // 4. Save to a MemoryStream using OoxmlSaveOptions
        // -------------------------------------------------
        OoxmlSaveOptions ooxmlOptions = new OoxmlSaveOptions();
        // Example option: set compression level (optional)
        ooxmlOptions.CompressionType = OoxmlCompressionType.Level6;

        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, ooxmlOptions);
            File.WriteAllBytes("Method4_FromStream.xlsx", stream.ToArray());
        }

        // Clean up resources
        workbook.Dispose();
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        SaveWorkbookDemo.Run();
    }
}