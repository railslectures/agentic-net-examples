using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook using LoadOptions (allows specifying culture if needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            // Example: use German culture for loading (optional)
            loadOptions.CultureInfo = new CultureInfo("de-DE");
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Create an instance of SettableGlobalizationSettings to customize localization
            SettableGlobalizationSettings locSettings = new SettableGlobalizationSettings();

            // Change the list separator from comma to semicolon
            locSettings.SetListSeparator(';');

            // Set custom display strings for boolean values
            locSettings.SetBooleanValueString(true, "WAHR");   // German for TRUE
            locSettings.SetBooleanValueString(false, "FALSCH"); // German for FALSE

            // Map standard function names to their German equivalents
            locSettings.SetLocalFunctionName("SUM", "SUMME", true);
            locSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);

            // Map a built‑in name (e.g., "Total") to a German term
            locSettings.SetLocalBuiltInName("Total", "Gesamt", true);

            // Apply the customized globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = locSettings;

            // Demonstrate the effect by using the localized function name in a formula
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);
            // Use the German function name "SUMME"
            sheet.Cells["A1"].Formula = "=SUMME(B1:B3)";

            // Calculate formulas so the result is stored in the cell
            workbook.CalculateFormula();

            // Output the calculated result to the console
            Console.WriteLine($"Result of localized SUMME formula: {sheet.Cells["A1"].Value}");

            // Save the localized workbook
            workbook.Save("localized_output.xlsx");
        }
    }
}