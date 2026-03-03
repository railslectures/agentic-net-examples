using System;
using Aspose.Cells;

class FormulaLocalDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the workbook region to German to demonstrate built‑in localization
        workbook.Settings.Region = CountryCode.Germany;

        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];

        // 1. Set a formula using the standard (English) syntax
        cell.Formula = "=SUM(B1:C1)";
        Console.WriteLine("Standard Formula: " + cell.Formula);
        Console.WriteLine("Localized Formula (German): " + cell.FormulaLocal);

        // 2. Set a formula using the localized (German) function name
        cell.FormulaLocal = "=SUMME(B1:C1)";
        Console.WriteLine("\nAfter assigning FormulaLocal:");
        Console.WriteLine("Standard Formula: " + cell.Formula);
        Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

        // 3. Use custom globalization settings to map a standard function to an Italian name
        var customSettings = new SettableGlobalizationSettings();
        customSettings.SetLocalFunctionName("SUM", "SOMMA", true); // map SUM ↔ SOMMA
        workbook.Settings.GlobalizationSettings = customSettings;

        // Populate data for the Italian SUM example
        worksheet.Cells["B2"].PutValue(1);
        worksheet.Cells["B3"].PutValue(2);
        worksheet.Cells["B4"].PutValue(3);
        worksheet.Cells["B5"].PutValue(4);

        // Apply a formula using the Italian localized name
        worksheet.Cells["A2"].FormulaLocal = "=SOMMA(B2:B5)";
        workbook.CalculateFormula();
        Console.WriteLine("\nResult of Italian SUM (SOMMA) in A2: " + worksheet.Cells["A2"].Value);

        // 4. Use FormulaParseOptions to indicate that the formula string is locale‑formatted
        var parseOptions = new FormulaParseOptions
        {
            LocaleDependent = true,
            R1C1Style = false
        };
        // Example: French date format inside TEXT function (argument separator is ';' for French)
        worksheet.Cells["A3"].SetFormula("TEXT(TODAY();\"[$-fr-FR]dddd, dd mmmm yyyy\")", parseOptions);
        Console.WriteLine("\nLocalized date formula set in A3 (FormulaLocal): " + worksheet.Cells["A3"].FormulaLocal);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}