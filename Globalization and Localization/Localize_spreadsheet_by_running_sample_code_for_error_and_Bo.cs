using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings that localize Boolean and error values.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values.
        public override string GetBooleanValueString(bool bv)
        {
            // Example: Russian localization.
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize error strings.
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

    public class Program
    {
        public static void Main()
        {
            // Path to the source XLSX file.
            string inputFile = "sample.xlsx";

            // Load the workbook (using the standard constructor – complies with lifecycle rules).
            Workbook wb = new Workbook(inputFile);

            // Apply the custom globalization settings to the workbook.
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Access the first worksheet.
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate cells with Boolean values and error strings for demonstration.
            cells[0, 0].PutValue(true);   // Boolean TRUE
            cells[0, 1].PutValue(false);  // Boolean FALSE

            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };

            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]);
            }

            // Display the localized string values in the console.
            for (int i = 0; i < 9; i++)
            {
                Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
            }

            // Save the localized workbook.
            string outputFile = "localized_output.xlsx";
            wb.Save(outputFile);
        }
    }
}