using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class WorkbookToByteArray
    {
        /// <summary>
        /// Creates a workbook, adds sample data, saves it to a memory stream in XLSX format,
        /// and returns the resulting byte array.
        /// </summary>
        public static byte[] GetWorkbookBytes()
        {
            // Initialize a new workbook (default format is Xlsx)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            // Create a memory stream to hold the workbook data
            using (MemoryStream stream = new MemoryStream())
            {
                // Save the workbook to the stream in XLSX format
                workbook.Save(stream, SaveFormat.Xlsx);

                // Return the stream contents as a byte array
                return stream.ToArray();
            }
        }

        // Example usage
        public static void Run()
        {
            byte[] xlsxBytes = GetWorkbookBytes();
            Console.WriteLine($"Workbook saved to byte array. Size: {xlsxBytes.Length} bytes");

            // Optionally write the byte array to a file for verification
            File.WriteAllBytes("GeneratedWorkbook.xlsx", xlsxBytes);
            Console.WriteLine("Workbook written to 'GeneratedWorkbook.xlsx'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToByteArray.Run();
        }
    }
}