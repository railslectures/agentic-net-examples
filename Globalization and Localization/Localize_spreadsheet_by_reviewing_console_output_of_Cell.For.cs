using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file (replace with your actual file)
        string inputPath = "sample.xlsx";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Set the workbook locale to German (de-DE) for localization
        workbook.Settings.Region = CountryCode.Germany;
        // Alternatively you can set CultureInfo directly:
        // workbook.Settings.CultureInfo = new CultureInfo("de-DE");

        // Access the first worksheet and cell A1
        Worksheet sheet = workbook.Worksheets[0];
        Cell cell = sheet.Cells["A1"];

        // Display the formula in standard (English) format and the localized format
        Console.WriteLine("Standard Formula: " + cell.Formula);
        Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

        // Set a formula using the German localized function name
        cell.FormulaLocal = "=SUMME(B1:C1)";

        // Show the formulas after assigning the localized version
        Console.WriteLine("\nAfter setting FormulaLocal:");
        Console.WriteLine("Standard Formula: " + cell.Formula);
        Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

        // Retrieve formulas via GetFormula to demonstrate the isLocal flag
        Console.WriteLine("\nUsing GetFormula:");
        Console.WriteLine("English (isLocal = false): " + cell.GetFormula(false, false));
        Console.WriteLine("German (isLocal = true): " + cell.GetFormula(false, true));

        // Calculate the workbook to ensure the formula evaluates correctly
        workbook.CalculateFormula();

        // Save the modified workbook (optional)
        workbook.Save("localized_output.xlsx");
    }
}