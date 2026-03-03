using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook with a specific culture (e.g., German)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("de-DE"); // German uses comma as decimal separator
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Create an instance of SettableGlobalizationSettings to customize localization
            SettableGlobalizationSettings locSettings = new SettableGlobalizationSettings();

            // Example: change the list separator from comma to semicolon
            locSettings.SetListSeparator(';');

            // Example: customize boolean display strings
            locSettings.SetBooleanValueString(true, "WAHR");   // German for TRUE
            locSettings.SetBooleanValueString(false, "FALSCH"); // German for FALSE

            // Example: map standard function names to localized names
            locSettings.SetLocalFunctionName("SUM", "SUMME", true);          // SUM -> SUMME
            locSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true); // AVERAGE -> MITTELWERT

            // Example: map a built‑in name (e.g., "Total") to a localized version
            locSettings.SetLocalBuiltInName("Total", "Gesamt", true);

            // Apply the localization settings to the workbook
            workbook.Settings.GlobalizationSettings = locSettings;

            // Demonstrate usage of the localized function name in a formula
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["B1"].PutValue(10);
            sheet.Cells["B2"].PutValue(20);
            sheet.Cells["B3"].PutValue(30);

            // Use the localized function name "SUMME" (German for SUM)
            sheet.Cells["A1"].Formula = "=SUMME(B1:B3)";

            // Calculate formulas so that the result is stored in the cell
            workbook.CalculateFormula();

            // Output the calculated result to the console
            Console.WriteLine($"Result of localized SUMME formula: {sheet.Cells["A1"].Value}");

            // Retrieve and display the localized name for the built‑in "Total"
            string localizedTotal = locSettings.GetLocalBuiltInName("Total");
            Console.WriteLine($"Localized built‑in name for 'Total': {localizedTotal}");

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}