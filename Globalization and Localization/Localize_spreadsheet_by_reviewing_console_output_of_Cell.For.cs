using System;
using Aspose.Cells;
using System.Globalization;

namespace AsposeCellsFormulaLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the existing XLSX workbook
            string inputPath = "input.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale to German (de-DE) to see localized formulas
            workbook.Settings.CultureInfo = new CultureInfo("de-DE");

            // Ensure formulas are parsed (in case they were stored as raw strings)
            workbook.ParseFormulas(false);

            // Get the first worksheet (you can iterate all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Iterate through all used cells
            foreach (Cell cell in cells)
            {
                // Process only cells that contain a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Standard (English) formula
                    string standardFormula = cell.Formula;

                    // Localized formula using the FormulaLocal property
                    string localizedFormula = cell.FormulaLocal;

                    // Localized formula using GetFormula with isLocal = true
                    string localizedViaGet = cell.GetFormula(false, true);

                    // Output the information to the console
                    Console.WriteLine($"Cell {cell.Name}:");
                    Console.WriteLine($"  Standard Formula : {standardFormula}");
                    Console.WriteLine($"  FormulaLocal     : {localizedFormula}");
                    Console.WriteLine($"  GetFormula(local): {localizedViaGet}");
                    Console.WriteLine();
                }
            }

            // Example: modify a formula using the localized format
            // (Assuming cell A1 originally has a SUM formula)
            Cell targetCell = worksheet.Cells["A1"];
            if (!string.IsNullOrEmpty(targetCell.Formula))
            {
                // Set a new formula using the German function name "SUMME"
                targetCell.FormulaLocal = "=SUMME(B1:C1)";
                Console.WriteLine($"After setting FormulaLocal on {targetCell.Name}:");
                Console.WriteLine($"  New Standard Formula : {targetCell.Formula}");
                Console.WriteLine($"  New Localized Formula: {targetCell.FormulaLocal}");
            }

            // Save the modified workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}