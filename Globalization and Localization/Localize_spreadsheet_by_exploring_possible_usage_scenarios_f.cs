using System;
using Aspose.Cells;

class FormulaLocalDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Set the workbook region to German to demonstrate localization
        workbook.Settings.Region = CountryCode.Germany;

        // Scenario 1: Write a formula in standard (English) format
        Cell cellA1 = worksheet.Cells["A1"];
        cellA1.Formula = "=SUM(B1:C1)"; // English function name
        Console.WriteLine("Standard Formula (A1): " + cellA1.Formula);
        Console.WriteLine("Localized Formula (A1): " + cellA1.FormulaLocal);

        // Scenario 2: Write a formula using the localized (German) name
        Cell cellA2 = worksheet.Cells["A2"];
        cellA2.FormulaLocal = "=SUMME(B1:C1)"; // German function name for SUM
        Console.WriteLine("\nAfter setting FormulaLocal (A2):");
        Console.WriteLine("Standard Formula (A2): " + cellA2.Formula);
        Console.WriteLine("Localized Formula (A2): " + cellA2.FormulaLocal);

        // Scenario 3: Use custom globalization settings to map a new localized function name
        SettableGlobalizationSettings customSettings = new SettableGlobalizationSettings();
        customSettings.SetLocalFunctionName("AVERAGE", "MITTELWERT", true);
        workbook.Settings.GlobalizationSettings = customSettings;

        // Populate sample data for the AVERAGE calculation
        worksheet.Cells["B1"].PutValue(10);
        worksheet.Cells["B2"].PutValue(20);
        worksheet.Cells["B3"].PutValue(30);

        Cell cellA3 = worksheet.Cells["A3"];
        cellA3.FormulaLocal = "=MITTELWERT(B1:B3)"; // Use the localized name
        workbook.CalculateFormula(); // Evaluate the formula
        Console.WriteLine("\nResult of localized AVERAGE (A3): " + cellA3.Value);

        // Scenario 4: Set a locale‑dependent formula (French date format)
        worksheet.Cells["A4"].Formula = "TEXT(TODAY(),\"[$-fr-FR]dddd, dd mmmm yyyy\")";
        Console.WriteLine("\nLocale‑dependent formula set in A4: " + worksheet.Cells["A4"].Formula);

        // Save the workbook with all changes
        workbook.Save("Output.xlsx");
    }
}