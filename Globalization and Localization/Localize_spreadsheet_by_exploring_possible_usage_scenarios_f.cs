using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize boolean and error values
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize boolean values (e.g., Russian)
        public override string GetBooleanValueString(bool bv)
        {
            return bv ? "ИСТИНА" : "ЛОЖЬ";
        }

        // Localize common Excel error strings (e.g., Russian equivalents)
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
            // -----------------------------------------------------------------
            // Step 1: Create a workbook and populate it with boolean values and
            //         error strings. This workbook will be saved and later loaded.
            // -----------------------------------------------------------------
            Workbook wbCreate = new Workbook();
            Cells cellsCreate = wbCreate.Worksheets[0].Cells;

            // Boolean values
            cellsCreate[0, 0].PutValue(true);   // A1
            cellsCreate[0, 1].PutValue(false);  // B1

            // Error strings (standard English)
            string[] errors = new string[]
            {
                "#NAME?", "#DIV/0!", "#REF!", "#VALUE!", "#N/A", "#NUM!", "#NULL!"
            };
            for (int i = 0; i < errors.Length; i++)
            {
                cellsCreate[0, i + 2].PutValue(errors[i]); // C1 onward
            }

            // Save the initial workbook (XLSX format)
            string originalPath = "original.xlsx";
            wbCreate.Save(originalPath);

            // -----------------------------------------------------------------
            // Step 2: Load the workbook from the saved XLSX file.
            // -----------------------------------------------------------------
            Workbook wb = new Workbook(originalPath);
            Cells cells = wb.Worksheets[0].Cells;

            // -----------------------------------------------------------------
            // Step 3: Apply custom globalization settings for localization.
            // -----------------------------------------------------------------
            wb.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // -----------------------------------------------------------------
            // Step 4: Display localized values using StringValue.
            //         Boolean values and error strings will appear in the
            //         localized form defined in CustomGlobalizationSettings.
            // -----------------------------------------------------------------
            Console.WriteLine("Localized cell values:");
            for (int col = 0; col < 9; col++)
            {
                Console.WriteLine($"Cell[0,{col}] ({cells[0, col].Name}): {cells[0, col].StringValue}");
            }

            // -----------------------------------------------------------------
            // Step 5: Save the workbook with applied localization.
            // -----------------------------------------------------------------
            string localizedPath = "localized.xlsx";
            wb.Save(localizedPath);
        }
    }
}