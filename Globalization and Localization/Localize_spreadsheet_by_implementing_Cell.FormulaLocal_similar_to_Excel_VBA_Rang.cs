using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing workbook (XLSX format)
            // Replace "input.xlsx" with the path to your source file
            Workbook workbook = new Workbook("input.xlsx");

            // Set the workbook's locale (region) to German.
            // This influences how formulas are represented in the local language.
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and cell A1
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Example 1: Read the formula in both standard (English) and localized (German) formats
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Example 2: Set a formula using the localized (German) syntax
            // In German, the SUM function is "SUMME"
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Verify that the standard formula has been translated automatically
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Recalculate the workbook so that the new formula result is computed
            workbook.CalculateFormula();

            // Display the calculated value of the cell
            Console.WriteLine("\nCalculated Value of A1: " + cell.Value);

            // Save the modified workbook to a new file
            // Replace "output.xlsx" with the desired output path
            workbook.Save("output.xlsx");
        }
    }
}