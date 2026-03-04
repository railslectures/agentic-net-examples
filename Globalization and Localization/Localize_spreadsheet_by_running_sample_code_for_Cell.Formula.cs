using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's default locale to German (de-DE) for demonstration
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Display the formula in standard (English) format
            Console.WriteLine("Standard Formula: " + cell.Formula);

            // Display the formula in the localized (German) format
            Console.WriteLine("Localized Formula (FormulaLocal): " + cell.FormulaLocal);

            // Set a formula using the German localized syntax
            // In German Excel the SUM function is "SUMME"
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Show the updated formulas
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula (FormulaLocal): " + cell.FormulaLocal);

            // Optionally, demonstrate GetFormula with localization flags
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula: " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula: " + cell.GetFormula(false, true));

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}