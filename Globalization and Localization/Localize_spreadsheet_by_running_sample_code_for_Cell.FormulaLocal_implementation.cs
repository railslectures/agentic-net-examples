using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX workbook
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook with default options (parsing formulas on open)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's locale to German (for demonstration of localization)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) notation
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Now set the formula using the German localized notation
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the difference
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Demonstrate GetFormula with localization flags
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula: " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula: " + cell.GetFormula(false, true));

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}