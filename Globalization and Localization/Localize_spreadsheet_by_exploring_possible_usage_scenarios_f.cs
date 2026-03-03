using System;
using Aspose.Cells;

namespace LocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            string inputPath = "sample.xlsx";
            Workbook wb = new Workbook(inputPath);

            // Apply custom globalization settings for Boolean and error values
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet and its cells
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // -------------------------------------------------
            // Boolean localization example
            // -------------------------------------------------
            cells["A1"].PutValue(true);   // Will be displayed as "ИСТИНА"
            cells["A2"].PutValue(false);  // Will be displayed as "ЛОЖЬ"

            // -------------------------------------------------
            // Error value localization example
            // -------------------------------------------------
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };

            // Place error strings starting from column C (index 2) in row 1
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // -------------------------------------------------
            // Display localized values in the console
            // -------------------------------------------------
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the workbook to verify that localization is persisted
            string outputPath = "localized_output.xlsx";
            wb.Save(outputPath);
        }
    }

    // Custom globalization settings overriding Boolean and error string representations
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values (example: Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize standard Excel error strings (example: Russian equivalents)
        public override string GetErrorValueString(string err)
        {
            switch (err)
            {
                case "#NAME?":   return "#ИМЯ?";
                case "#DIV/0!":  return "#ДЕЛ/0!";
                case "#REF!":    return "#ССЫЛКА!";
                case "#VALUE!":  return "#ЗНАЧ!";
                case "#N/A":     return "#Н/Д";
                case "#NUM!":    return "#ЧИСЛО!";
                case "#NULL!":   return "#ПУСТО!";
                default:         return base.GetErrorValueString(err);
            }
        }
    }
}