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

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to a memory stream in XLSX format using default settings
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);

            // Reset the stream position if it will be read later
            stream.Position = 0;

            // Example: write the stream content to a physical file for verification
            using (FileStream file = new FileStream("SavedFromStream.xlsx", FileMode.Create, FileAccess.Write))
            {
                stream.CopyTo(file);
            }
        }

        // Dispose the workbook when done
        workbook.Dispose();
    }
}