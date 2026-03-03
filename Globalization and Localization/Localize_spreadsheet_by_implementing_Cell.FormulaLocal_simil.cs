using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file (must exist)
            string inputPath = "input.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's locale to German (de-DE) for demonstration
            // This influences how FormulaLocal is interpreted and displayed
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Example 1: Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("After setting standard formula:");
            Console.WriteLine("Standard Formula   : " + cell.Formula);
            Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);
            Console.WriteLine();

            // Example 2: Set a formula using the localized (German) syntax via FormulaLocal
            // In German Excel, the SUM function is called "SUMME"
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the conversion
            Console.WriteLine("After setting localized formula:");
            Console.WriteLine("Standard Formula   : " + cell.Formula);
            Console.WriteLine("Localized Formula  : " + cell.FormulaLocal);
            Console.WriteLine();

            // Demonstrate GetFormula with localization flags
            // GetFormula(false, false) -> standard (English) A1 notation
            // GetFormula(false, true)  -> localized (German) A1 notation
            Console.WriteLine("Using GetFormula:");
            Console.WriteLine("English formula    : " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula  : " + cell.GetFormula(false, true));
            Console.WriteLine();

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}