using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the XLSX file to be examined
            string inputPath = "SampleWorkbook.xlsx";

            // Load the workbook (standard load, formulas will be parsed on open)
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook region to German to see localized function names (e.g., SUM -> SUMME)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet (adjust if needed)
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Iterate through all used cells and display both standard and localized formulas
            Console.WriteLine("Cell\tStandard Formula\tLocalized Formula");
            foreach (Cell cell in cells)
            {
                // Only interested in cells that actually contain a formula
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    // Standard (English) formula
                    string standardFormula = cell.Formula;

                    // Locale‑formatted formula (German in this case)
                    string localizedFormula = cell.FormulaLocal;

                    // Also demonstrate GetFormula with isLocal = true
                    string getFormulaLocal = cell.GetFormula(false, true);

                    Console.WriteLine($"{cell.Name}\t{standardFormula}\t{localizedFormula}");

                    // Verify that GetFormula with isLocal matches FormulaLocal
                    if (localizedFormula != getFormulaLocal)
                    {
                        Console.WriteLine($"[Info] GetFormula(local) differs: {getFormulaLocal}");
                    }
                }
            }

            // Optionally, save the workbook after any modifications (none made here)
            // workbook.Save("LocalizedOutput.xlsx");
        }
    }
}