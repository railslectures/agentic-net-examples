using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Input Excel file (XLSX) and output PDF file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Load the workbook from the XLSX file (uses Workbook(string) constructor)
        Workbook workbook = new Workbook(inputPath);

        // Save the workbook as PDF (uses Workbook.Save(string, SaveFormat))
        workbook.Save(outputPath, SaveFormat.Pdf);

        Console.WriteLine("Excel workbook successfully converted to PDF.");
    }
}