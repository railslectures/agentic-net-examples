using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file (replace with actual path)
            string inputPath = "input.xlsx";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's default locale to German for demonstration
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) notation
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized formats
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Set the formula using the German localized notation
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the difference
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Demonstrate GetFormula with localization options
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula: " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula: " + cell.GetFormula(false, true));

            // Save the modified workbook (replace with desired output path)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}