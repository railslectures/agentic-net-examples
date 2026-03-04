using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize boolean and error values
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        // Override boolean display strings
        public override string GetBooleanValueString(bool bv)
        {
            // Example: Russian localization
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Override error display strings
        public override string GetErrorValueString(string err)
        {
            // Map standard English error texts to Russian equivalents
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

    class Program
    {
        static void Main()
        {
            // Create a new workbook (uses the provided create rule)
            Workbook wb = new Workbook();

            // Assign custom globalization settings to the workbook
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet and its cells collection
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with boolean values
            cells[0, 0].PutValue(true);   // A1
            cells[0, 1].PutValue(false);  // B1

            // Populate cells with standard error strings
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                // Starting from C1 (column index 2)
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Display localized values in the console
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                // StringValue reflects the localized representation
                Console.WriteLine($"Cell[0,{col}] ({cells[0, col].Name}): {cells[0, col].StringValue}");
            }

            // Save the workbook to disk (uses the provided save rule)
            wb.Save("LocalizedWorkbook.xlsx");
        }
    }
}