using System;
using System.Globalization;
using Aspose.Cells;

namespace FormulaLocalLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Scenario 1: Display standard and localized formulas
            // ------------------------------------------------------------
            // Set workbook region to German to see German localization
            workbook.Settings.Region = CountryCode.Germany;

            // Put a sample formula in English (standard format)
            Cell cellA1 = cells["A1"];
            cellA1.Formula = "=SUM(B1:C1)";

            Console.WriteLine("Scenario 1:");
            Console.WriteLine($"Standard Formula (Formula): {cellA1.Formula}");
            Console.WriteLine($"Localized Formula (FormulaLocal): {cellA1.FormulaLocal}");
            Console.WriteLine();

            // ------------------------------------------------------------
            // Scenario 2: Set formula using localized (German) syntax
            // ------------------------------------------------------------
            // Assign a German formula directly via FormulaLocal
            cellA1.FormulaLocal = "=SUMME(B1:C1)";

            Console.WriteLine("Scenario 2:");
            Console.WriteLine($"After setting FormulaLocal:");
            Console.WriteLine($"Standard Formula (Formula): {cellA1.Formula}");
            Console.WriteLine($"Localized Formula (FormulaLocal): {cellA1.FormulaLocal}");
            Console.WriteLine();

            // ------------------------------------------------------------
            // Scenario 3: Retrieve formulas with GetFormula (localized flag)
            // ------------------------------------------------------------
            Console.WriteLine("Scenario 3:");
            Console.WriteLine($"GetFormula (standard): {cellA1.GetFormula(false, false)}");
            Console.WriteLine($"GetFormula (localized): {cellA1.GetFormula(false, true)}");
            Console.WriteLine();

            // ------------------------------------------------------------
            // Scenario 4: Use custom globalization settings to map a function
            // ------------------------------------------------------------
            // Create custom settings that map English "AVERAGE" to French "MOYENNE"
            SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
            customSettings.SetLocalFunctionName("AVERAGE", "MOYENNE", true);
            workbook.Settings.GlobalizationSettings = customSettings;
            // Set culture to French so that localized function names are recognized
            workbook.Settings.CultureInfo = new CultureInfo("fr-FR");

            // Put sample data
            cells["B2"].PutValue(10);
            cells["B3"].PutValue(20);
            cells["B4"].PutValue(30);

            // Use the localized function name in a formula
            Cell cellB1 = cells["B1"];
            cellB1.Formula = "=MOYENNE(B2:B4)";

            // Calculate to verify that the mapping works
            workbook.CalculateFormula();

            Console.WriteLine("Scenario 4:");
            Console.WriteLine($"Formula using localized function: {cellB1.Formula}");
            Console.WriteLine($"Result (should be 20): {cellB1.Value}");
            Console.WriteLine();

            // ------------------------------------------------------------
            // Scenario 5: Parse a locale‑dependent formula using FormulaParseOptions
            // ------------------------------------------------------------
            FormulaParseOptions parseOptions = new FormulaParseOptions
            {
                LocaleDependent = false,
                R1C1Style = false
            };

            Cell cellC1 = cells["C1"];
            // Set formula without leading '=' when using FormulaParseOptions
            cellC1.SetFormula("TEXT(TODAY(),\"dd/mm/yyyy\")", parseOptions);

            Console.WriteLine("Scenario 5:");
            Console.WriteLine($"Locale‑dependent formula set in C1: {cellC1.Formula}");
            Console.WriteLine();

            // ------------------------------------------------------------
            // Scenario 6: Retrieve the localized name of a standard function via settings
            // ------------------------------------------------------------
            string localizedSumName = customSettings.GetLocalFunctionName("SUM");
            Console.WriteLine("Scenario 6:");
            Console.WriteLine($"Localized name for 'SUM' in current settings: {localizedSumName}");
            Console.WriteLine();

            // ------------------------------------------------------------
            // Save the workbook with all changes
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}