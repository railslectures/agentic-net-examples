using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (uses Workbook() constructor rule)
        Workbook workbook = new Workbook();

        // Add some sample data
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Save the workbook to a memory stream in XLSX format (uses Save(Stream, SaveFormat) rule)
        using (MemoryStream stream = new MemoryStream())
        {
            workbook.Save(stream, SaveFormat.Xlsx);

            // Convert the stream content to a byte array
            byte[] workbookBytes = stream.ToArray();

            // Example usage of the byte array
            Console.WriteLine($"Workbook saved to byte array. Size = {workbookBytes.Length} bytes");
        }

        // Clean up
        workbook.Dispose();
    }
}