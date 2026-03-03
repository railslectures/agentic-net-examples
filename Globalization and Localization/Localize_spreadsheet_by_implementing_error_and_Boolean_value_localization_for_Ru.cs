using System;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings for Russian language
    public class RussianGlobalizationSettings : GlobalizationSettings
    {
        // Localize Boolean values
        public override string GetBooleanValueString(bool bv)
        {
            // TRUE -> ИСТИНА, FALSE -> ЛОЖЬ
            return bv ? "ИСТИНА" : "ЛОЖЬ";
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
            string inputPath = "input.xlsx";

            // LoadOptions with Russian culture (affects number/date formatting)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.CultureInfo = new CultureInfo("ru-RU");

            // Load the workbook using the specified LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Apply the custom Russian globalization settings
            workbook.Settings.GlobalizationSettings = new RussianGlobalizationSettings();

            // Example: put some Boolean and error values to see localization
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            cells[0, 0].PutValue(true);   // Will display "ИСТИНА"
            cells[0, 1].PutValue(false);  // Will display "ЛОЖЬ"

            string[] errors = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
            for (int i = 0; i < errors.Length; i++)
            {
                cells[0, i + 2].PutValue(errors[i]); // Localized error strings will appear
            }

            // Save the localized workbook
            string outputPath = "output.xlsx";
            workbook.Save(outputPath);
        }
    }
}