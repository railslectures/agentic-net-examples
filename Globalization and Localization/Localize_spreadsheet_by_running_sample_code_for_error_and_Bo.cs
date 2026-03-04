using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize Boolean and error values
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values (e.g., Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error strings (e.g., Russian)
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
            // Path to the source XLSX file (must exist)
            string inputPath = "input.xlsx";

            // Load the workbook (using the standard load constructor)
            Workbook wb = new Workbook(inputPath);

            // Access the first worksheet and its cells
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Populate Boolean values
            cells[0, 0].PutValue(true);   // A1
            cells[0, 1].PutValue(false);  // B1

            // Populate error strings
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                // Starting from column C (index 2)
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Apply the custom globalization settings to the workbook
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Display localized values in the console
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}] (localized): {cells[0, col].StringValue}");
            }

            // Save the localized workbook
            string outputPath = "output.xlsx";
            wb.Save(outputPath);
        }
    }
}