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

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale (example: German)
            workbook.Settings.Region = CountryCode.Germany;

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access a specific cell (A1)
            Cell cell = sheet.Cells["A1"];

            // If the cell already contains a formula, display both the standard
            // (English) and the localized (German) versions
            Console.WriteLine("Original formulas:");
            Console.WriteLine("Standard (Formula)   : " + cell.Formula);
            Console.WriteLine("Localized (FormulaLocal) : " + cell.FormulaLocal);

            // Set a formula using the localized syntax (German function name)
            // Example: SUMME is the German equivalent of SUM
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // After setting, the standard Formula property is automatically updated
            Console.WriteLine("\nAfter assigning FormulaLocal:");
            Console.WriteLine("Standard (Formula)   : " + cell.Formula);
            Console.WriteLine("Localized (FormulaLocal) : " + cell.FormulaLocal);

            // Recalculate the workbook so the formula result is refreshed
            workbook.CalculateFormula();

            // Show the calculated value
            Console.WriteLine("\nCalculated value in A1: " + cell.Value);

            // Save the modified workbook to a new file
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}