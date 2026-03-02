using System;
using System.IO;
using Aspose.Cells;

public class SaveWorkbookToStreamDemo
{
    public static void Main(string[] args)
    {
        Run();
    }

    public static void Run()
    {
        // Create a new workbook (default format is Xlsx)
        Workbook workbook = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");
        sheet.Cells["A2"].PutValue(123);
        sheet.Cells["B2"].PutValue(456);

        // Create a memory stream to hold the workbook data
        using (MemoryStream stream = new MemoryStream())
        {
            // Save the workbook to the stream in XLSX format
            workbook.Save(stream, SaveFormat.Xlsx);

            // Reset the stream position if you need to read from it later
            stream.Position = 0;

            // Optional: write the stream content to a physical file for verification
            using (FileStream file = new FileStream("WorkbookFromStream.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
        }

        // Clean up resources
        workbook.Dispose();
    }
}