using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings for Russian language
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values
        public override string GetBooleanValueString(bool value)
        {
            return value ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize error values
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

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string inputFile = "input.xlsx";

            // Load the workbook (XLSX format)
            Workbook workbook = new Workbook(inputFile);

            // Apply the custom Russian globalization settings
            workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

            // Example: write some test data to demonstrate localization
            Cells cells = workbook.Worksheets[0].Cells;
            cells[0, 0].PutValue(true);   // Boolean TRUE
            cells[0, 1].PutValue(false);  // Boolean FALSE

            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };

            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]); // Insert error strings
            }

            // Display localized values in the console
            for (int i = 0; i < 9; i++)
            {
                Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
            }

            // Save the localized workbook
            string outputFile = "output.xlsx";
            workbook.Save(outputFile);
        }
    }
}