using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file (replace with actual path)
            string inputPath = "input.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook region to German to obtain German localized formulas
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet (adjust if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Iterate through all cells that contain data/formulas
            foreach (Cell cell in worksheet.Cells)
            {
                // Process only cells that have a formula
                if (cell.IsFormula)
                {
                    // Display the cell address, the standard (English) formula, and the localized formula
                    Console.WriteLine($"Cell {cell.Name}");
                    Console.WriteLine($"  Standard Formula : {cell.Formula}");
                    Console.WriteLine($"  Localized Formula: {cell.FormulaLocal}");
                }
            }

            // Demonstrate setting a formula using the localized syntax (German)
            Cell demoCell = worksheet.Cells["B2"];
            demoCell.FormulaLocal = "=SUMME(A1:A5)"; // German function name for SUM

            // Show the effect of setting FormulaLocal
            Console.WriteLine("\nAfter setting FormulaLocal on B2:");
            Console.WriteLine($"  Standard Formula : {demoCell.Formula}");
            Console.WriteLine($"  Localized Formula: {demoCell.FormulaLocal}");

            // Recalculate to ensure the formula result is up‑to‑date
            workbook.CalculateFormula();

            // Output the calculated value of the demo cell
            Console.WriteLine($"  Calculated Value : {demoCell.Value}");

            // Save the modified workbook (optional)
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}