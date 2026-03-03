using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source XLSX file
            string inputPath = "input.xlsx";

            // Load the workbook (uses the provided load rule)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook locale to German (example of localization)
            // This influences how FormulaLocal is interpreted and displayed
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell (A1)
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) syntax
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized formats
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Now set the formula using the localized (German) syntax via FormulaLocal
            // In German, the SUM function is "SUMME"
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // After setting FormulaLocal, the standard Formula property reflects the English equivalent
            Console.WriteLine("\nAfter assigning FormulaLocal:");
            Console.WriteLine("Standard Formula : " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Optionally, retrieve formulas using GetFormula with explicit flags
            string englishFormula = cell.GetFormula(false, false); // not R1C1, not local
            string localFormula   = cell.GetFormula(false, true);  // not R1C1, local
            Console.WriteLine("\nGetFormula results:");
            Console.WriteLine("English : " + englishFormula);
            Console.WriteLine("Local   : " + localFormula);

            // Recalculate the workbook to ensure the formula result is up‑to‑date
            workbook.CalculateFormula();

            // Output the calculated value of the cell
            Console.WriteLine("\nCalculated Value: " + cell.Value);

            // Save the modified workbook (uses the provided save rule)
            string outputPath = "localized_output.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}