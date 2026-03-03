using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook wb = new Workbook(inputPath); // load rule

        // Prepare error strings and boolean values
        string[] errors = new string[] { "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!" };
        Cells cells = wb.Worksheets[0].Cells;

        // Boolean values
        cells[0, 0].PutValue(true);
        cells[0, 1].PutValue(false);

        // Error values
        for (int i = 0; i < errors.Length; i++)
        {
            cells[0, i + 2].PutValue(errors[i]);
        }

        // Apply custom globalization settings for localization
        wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Display localized strings for each cell in the first row
        for (int i = 0; i < 9; i++)
        {
            Console.WriteLine($"Cell[0,{i}]: {cells[0, i].StringValue}");
        }

        // Save the modified workbook
        wb.Save("output.xlsx"); // save rule
    }

    // Custom globalization settings overriding boolean and error string representations
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetBooleanValueString(bool bv)
        {
            // Russian localization for boolean values
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        public override string GetErrorValueString(string err)
        {
            // Russian localization for common Excel error values
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
}