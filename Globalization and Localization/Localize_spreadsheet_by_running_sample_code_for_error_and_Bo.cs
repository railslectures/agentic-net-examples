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
            switch (err)
            {
                case "#NAME?": return "#ИМЯ?";
                case "#DIV/0!": return "#ДЕЛ/0!";
                case "#REF!": return "#ССЫЛКА!";
                case "#VALUE!": return "#ЗНАЧ!";
                case "#N/A": return "#Н/Д";
                case "#NUM!": return "#ЧИСЛО!";
                case "#NULL!": return "#ПУСТО!";
                default: return base.GetErrorValueString(err);
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the existing XLSX file to be localized
            string inputPath = "input.xlsx";

            // Load the workbook (using the standard constructor)
            Workbook wb = new Workbook(inputPath);

            // Access the first worksheet and its cells
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Boolean values and error strings
            cells[0, 0].PutValue(true);   // Boolean TRUE
            cells[0, 1].PutValue(false);  // Boolean FALSE

            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };

            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]); // Place each error string in the first row
            }

            // Apply the custom globalization settings to the workbook
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Display localized values in the console
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}]: {cells[0, col].StringValue}");
            }

            // Save the localized workbook
            string outputPath = "output.xlsx";
            wb.Save(outputPath);
        }
    }
}