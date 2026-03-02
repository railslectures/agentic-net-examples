using System;
using Aspose.Cells;

namespace AsposeCellsSaveMultipleFormats
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (default constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Age");
            sheet.Cells["A2"].PutValue("John");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Jane");
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["A4"].PutValue("Doe");
            sheet.Cells["B4"].PutValue(35);

            // Save the workbook in various formats using default settings
            workbook.Save("Output.xlsx", SaveFormat.Xlsx);               // Excel Open XML
            workbook.Save("Output.xls", SaveFormat.Excel97To2003);       // Excel 97-2003
            workbook.Save("Output.csv", SaveFormat.Csv);                // CSV
            workbook.Save("Output.pdf", SaveFormat.Pdf);                // PDF
            workbook.Save("Output.html", SaveFormat.Html);              // HTML
            workbook.Save("Output.ods", SaveFormat.Ods);                // OpenDocument Spreadsheet
            workbook.Save("Output.tif", SaveFormat.Tiff);               // TIFF image
            workbook.Save("Output.svg", SaveFormat.Svg);                // SVG image
            workbook.Save("Output.md", SaveFormat.Markdown);            // Markdown
            workbook.Save("Output.epub", SaveFormat.Epub);              // EPUB ebook

            Console.WriteLine("Workbook saved in multiple formats successfully.");
        }
    }
}