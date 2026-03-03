using System;
using Aspose.Cells;

namespace AsposeCellsFormulaLocalDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Set the workbook's locale (example: German)
            workbook.Settings.Region = CountryCode.Germany;

            // Access the first worksheet and a target cell
            Worksheet worksheet = workbook.Worksheets[0];
            Cell cell = worksheet.Cells["A1"];

            // Set a formula using the standard (English) notation
            cell.Formula = "=SUM(B1:C1)";

            // Display the formula in both standard and localized forms
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Set a formula using the localized (German) notation
            cell.FormulaLocal = "=SUMME(B1:C1)";

            // Display the formulas again to show the conversion
            Console.WriteLine("\nAfter setting FormulaLocal:");
            Console.WriteLine("Standard Formula: " + cell.Formula);
            Console.WriteLine("Localized Formula: " + cell.FormulaLocal);

            // Optionally, demonstrate GetFormula with localization flags
            Console.WriteLine("\nUsing GetFormula:");
            Console.WriteLine("English formula: " + cell.GetFormula(false, false));
            Console.WriteLine("Localized formula: " + cell.GetFormula(false, true));

            // Save the modified workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to {outputPath}");
        }
    }
}