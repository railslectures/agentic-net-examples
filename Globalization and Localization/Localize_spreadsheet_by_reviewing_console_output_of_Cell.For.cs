using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to be examined
            string inputPath = "input.xlsx";

            // Load the workbook (creation and loading follow the standard Aspose.Cells API)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook region to German to see localized function names (e.g., SUM -> SUMME)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet (you can iterate all worksheets if needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Determine the used range of the worksheet
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            Console.WriteLine("=== Formula Localization Review ===");
            Console.WriteLine($"Workbook region set to: {workbook.Settings.Region}");
            Console.WriteLine();

            // Iterate through all cells that contain formulas
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Process only cells that actually have a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Standard (English) formula
                        string standardFormula = cell.Formula;

                        // Localized formula according to the workbook's region
                        string localizedFormula = cell.FormulaLocal;

                        // Alternative way using GetFormula with isLocal = true
                        string getFormulaLocal = cell.GetFormula(false, true);

                        // Output the information
                        Console.WriteLine($"Cell {cell.Name}:");
                        Console.WriteLine($"  Standard Formula : {standardFormula}");
                        Console.WriteLine($"  Localized Formula: {localizedFormula}");
                        Console.WriteLine($"  GetFormula(true): {getFormulaLocal}");
                        Console.WriteLine();
                    }
                }
            }

            // Optionally, save the workbook after any modifications (not required for review)
            // workbook.Save("output.xlsx");
        }
    }
}