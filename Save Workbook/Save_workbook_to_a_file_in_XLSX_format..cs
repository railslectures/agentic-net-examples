using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook instance
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data to cells
        worksheet.Cells["A1"].PutValue("Hello");
        worksheet.Cells["B1"].PutValue("World");

        // Save the workbook to an XLSX file using the Save(string, SaveFormat) overload
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}