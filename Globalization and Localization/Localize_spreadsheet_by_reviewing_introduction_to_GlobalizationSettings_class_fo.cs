using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings inheriting from SettableGlobalizationSettings
    // to demonstrate overriding of boolean strings, error strings and custom separators.
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        public CustomGlobalizationSettings()
        {
            // Example: change list separator from comma to semicolon
            SetListSeparator(';');

            // Example: change boolean display strings
            SetBooleanValueString(true, "ИСТИНА");
            SetBooleanValueString(false, "ЛОЖЬ");

            // Example: map standard function name "SUM" to a local name "СУММА"
            SetLocalFunctionName("SUM", "СУММА", true);
        }

        // Override boolean display strings (alternative way)
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Override error value strings for Russian equivalents
        public override string GetErrorValueString(string err)
        {
            return err switch
            {
                "#NAME?" => "#ИМЯ?",
                "#DIV/0!" => "#ДЕЛ/0!",
                "#REF!" => "#ССЫЛКА!",
                "#VALUE!" => "#ЗНАЧ!",
                "#N/A" => "#Н/Д",
                "#NUM!" => "#ЧИСЛО!",
                "#NULL!" => "#ПУСТО!",
                _ => base.GetErrorValueString(err)
            };
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Apply custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Demonstrate localized boolean and error values
            Cells cells = workbook.Worksheets[0].Cells;

            // Boolean values
            cells["A1"].PutValue(true);
            cells["A2"].PutValue(false);

            // Error values
            string[] errors = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]); // Row 0, columns C..I
            }

            // Use the localized function name in a formula
            cells["B1"].PutValue(10);
            cells["B2"].PutValue(20);
            cells["B3"].PutValue(30);
            cells["C1"].Formula = "=СУММА(B1:B3)"; // "SUM" localized as "СУММА"

            // Calculate formulas to reflect the localized function
            workbook.CalculateFormula();

            // Output some results to console for verification
            Console.WriteLine($"A1 (bool): {cells["A1"].StringValue}");
            Console.WriteLine($"A2 (bool): {cells["A2"].StringValue}");
            for (int i = 0; i < errors.Length; i++)
            {
                Console.WriteLine($"Error cell {i + 2}: {cells[0, i + 2].StringValue}");
            }
            Console.WriteLine($"C1 (localized SUM result): {cells["C1"].StringValue}");

            // Save the localized workbook
            workbook.Save("localized_output.xlsx");
        }
    }
}