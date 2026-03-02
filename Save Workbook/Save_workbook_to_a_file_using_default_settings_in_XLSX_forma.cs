using System;
using Aspose.Cells;

namespace AsposeCellsSaveExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the Workbook() constructor rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue("World");

            // Save the workbook to an XLSX file using default settings (Save(string) rule)
            workbook.Save("output.xlsx");

            Console.WriteLine("Workbook saved as output.xlsx");
        }
    }
}