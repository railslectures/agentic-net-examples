using System;
using Aspose.Cells;

namespace FormulaLocalLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the workbook (XLSX) to be loaded.
            string inputPath = "input.xlsx";
            // Path where the modified workbook will be saved.
            string outputPath = "output.xlsx";

            // Load the workbook from the specified file.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -----------------------------------------------------------------
            // 1. Set the workbook's default locale to German (Germany).
            // -----------------------------------------------------------------
            workbook.Settings.Region = CountryCode.Germany;

            // -----------------------------------------------------------------
            // 2. Demonstrate reading FormulaLocal when the formula is set in
            //    standard (English) format.
            // -----------------------------------------------------------------
            Cell cellA1 = cells["A1"];
            cellA1.Formula = "=SUM(B1:C1)"; // English formula.
            Console.WriteLine("=== English formula set ===");
            Console.WriteLine($"Standard Formula : {cellA1.Formula}");
            Console.WriteLine($"Localized Formula: {cellA1.FormulaLocal}");

            // -----------------------------------------------------------------
            // 3. Set a formula using the German localized function name via
            //    FormulaLocal and observe the conversion to the standard format.
            // -----------------------------------------------------------------
            Cell cellA2 = cells["A2"];
            cellA2.FormulaLocal = "=SUMME(B2:C2)"; // German formula.
            Console.WriteLine("\n=== German formula set via FormulaLocal ===");
            Console.WriteLine($"Standard Formula : {cellA2.Formula}");
            Console.WriteLine($"Localized Formula: {cellA2.FormulaLocal}");

            // Populate the referenced cells with sample data.
            cells["B1"].PutValue(10);
            cells["C1"].PutValue(20);
            cells["B2"].PutValue(5);
            cells["C2"].PutValue(15);

            // -----------------------------------------------------------------
            // 4. Use SettableGlobalizationSettings to define a custom localized
            //    function name for SUM (e.g., "SOMME" for French) and enable
            //    bidirectional mapping.
            // -----------------------------------------------------------------
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            customSettings.SetLocalFunctionName("SUM", "SOMME", true); // French SUM.
            workbook.Settings.GlobalizationSettings = customSettings;

            // Apply the custom localized function in a formula.
            Cell cellA3 = cells["A3"];
            cellA3.FormulaLocal = "=SOMME(B3:C3)";
            cells["B3"].PutValue(7);
            cells["C3"].PutValue(13);
            Console.WriteLine("\n=== Custom localized function (French) via FormulaLocal ===");
            Console.WriteLine($"Standard Formula : {cellA3.Formula}");
            Console.WriteLine($"Localized Formula: {cellA3.FormulaLocal}");

            // -----------------------------------------------------------------
            // 5. Use FormulaParseOptions with LocaleDependent = true to set a
            //    locale‑formatted formula directly (e.g., French date format).
            // -----------------------------------------------------------------
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                LocaleDependent = true,
                R1C1Style = false
            };
            Cell cellA4 = cells["A4"];
            // French date format example.
            cellA4.SetFormula("=TEXTE(AUJOURDHUI();\"[$-fr-FR]dddd, dd mmmm yyyy\")", parseOptions);
            Console.WriteLine("\n=== Locale‑dependent formula set via FormulaParseOptions ===");
            Console.WriteLine($"Standard Formula : {cellA4.Formula}");
            Console.WriteLine($"Localized Formula: {cellA4.FormulaLocal}");

            // -----------------------------------------------------------------
            // 6. Calculate all formulas to ensure they are evaluated correctly.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // Display calculated results.
            Console.WriteLine("\n=== Calculated Results ===");
            Console.WriteLine($"A1 (SUM) result : {cellA1.Value}");
            Console.WriteLine($"A2 (SUMME) result: {cellA2.Value}");
            Console.WriteLine($"A3 (SOMME) result: {cellA3.Value}");
            Console.WriteLine($"A4 (date text)   : {cellA4.Value}");

            // -----------------------------------------------------------------
            // 7. Save the modified workbook.
            // -----------------------------------------------------------------
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}