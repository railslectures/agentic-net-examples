using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Set the workbook's locale to German to demonstrate localization
        workbook.Settings.Region = CountryCode.Germany;

        // Access the first worksheet and a target cell (A1)
        Worksheet worksheet = workbook.Worksheets[0];
        Cell cell = worksheet.Cells["A1"];

        // Set a formula using the standard (English) notation
        cell.Formula = "=SUM(B1:C1)";

        // Display the formula in both standard and localized forms
        Console.WriteLine("Standard Formula: " + cell.Formula);
        Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

        // Now set the formula using the German localized function name
        cell.FormulaLocal = "=SUMME(B1:C1)";

        // Show the updated formulas after assigning the localized version
        Console.WriteLine("\nAfter setting FormulaLocal:");
        Console.WriteLine("Standard Formula: " + cell.Formula);
        Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

        // Optionally, retrieve formulas with explicit flags
        Console.WriteLine("\nUsing GetFormula:");
        Console.WriteLine("English formula: " + cell.GetFormula(false, false));
        Console.WriteLine("Localized formula: " + cell.GetFormula(false, true));

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}