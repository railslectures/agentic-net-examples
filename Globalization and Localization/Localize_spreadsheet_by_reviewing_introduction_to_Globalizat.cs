using System;
using Aspose.Cells;

class GlobalizationDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook (create rule)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create an instance of SettableGlobalizationSettings (create rule)
        SettableGlobalizationSettings settings = new SettableGlobalizationSettings();

        // Customize the list separator (e.g., use semicolon instead of comma)
        settings.SetListSeparator(';');

        // Customize the display strings for boolean values
        settings.SetBooleanValueString(true, "TRUE_LOCAL");
        settings.SetBooleanValueString(false, "FALSE_LOCAL");

        // Map standard function names to localized names (bidirectional mapping)
        settings.SetLocalFunctionName("SUM", "SOMME", true);        // Example: French for SUM
        settings.SetLocalFunctionName("AVERAGE", "MOYENNE", true); // Example: French for AVERAGE

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = settings;

        // Demonstrate the effect of the localized function names
        Worksheet ws = workbook.Worksheets[0];
        ws.Cells["B1"].PutValue(10);
        ws.Cells["B2"].PutValue(20);
        ws.Cells["B3"].PutValue(30);

        // Use the localized function names in formulas
        ws.Cells["A1"].Formula = "=SOMME(B1:B3)";   // Localized SUM
        ws.Cells["A2"].Formula = "=MOYENNE(B1:B3)"; // Localized AVERAGE

        // Calculate the formulas (create rule)
        workbook.CalculateFormula();

        // Output the calculated results
        Console.WriteLine($"Result of SOMME: {ws.Cells["A1"].Value}");
        Console.WriteLine($"Result of MOYENNE: {ws.Cells["A2"].Value}");

        // Save the modified workbook (save rule)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}