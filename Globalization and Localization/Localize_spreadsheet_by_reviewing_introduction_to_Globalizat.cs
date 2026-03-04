using System;
using System.Globalization;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook with German culture settings.
        // LoadOptions.CultureInfo influences how numbers, dates, etc., are parsed.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.CultureInfo = new CultureInfo("de-DE");
        Workbook workbook = new Workbook("input.xlsx", loadOptions);

        // Apply custom globalization settings to the workbook.
        // This overrides how boolean values, error messages, and function names are displayed.
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Demonstrate the effect of the custom settings.
        // Assume A1 and A2 contain boolean values, A3 contains an error value.
        Cells cells = workbook.Worksheets[0].Cells;
        Console.WriteLine($"A1 (boolean): {cells["A1"].StringValue}");
        Console.WriteLine($"A2 (boolean): {cells["A2"].StringValue}");
        Console.WriteLine($"A3 (error)   : {cells["A3"].StringValue}");

        // Use a localized function name in a formula to verify mapping.
        cells["B1"].PutValue(10);
        cells["B2"].PutValue(20);
        cells["B3"].PutValue(30);
        cells["C1"].Formula = "=SUMME(B1:B3)"; // "SUMME" is the German localized name for SUM.
        workbook.CalculateFormula();
        Console.WriteLine($"C1 (SUMME result): {cells["C1"].Value}");

        // Save the workbook with the applied globalization settings.
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings that localize booleans, errors, and function names.
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize boolean display strings.
        public override string GetBooleanValueString(bool value)
        {
            return value ? "WAHR" : "FALSCH"; // German for TRUE/FALSE
        }

        // Localize common Excel error messages.
        public override string GetErrorValueString(string error)
        {
            switch (error)
            {
                case "#DIV/0!": return "#DIV/0!";
                case "#NAME?":  return "#NAME?";
                case "#REF!":   return "#BEZUG!";
                case "#VALUE!": return "#WERT!";
                case "#N/A":    return "#NV";
                case "#NUM!":   return "#ZAHL!";
                case "#NULL!":  return "#NULL!";
                default:        return base.GetErrorValueString(error);
            }
        }

        // Map the standard function name "SUM" to its German equivalent "SUMME".
        public override string GetLocalFunctionName(string standardName)
        {
            if (standardName.Equals("SUM", StringComparison.OrdinalIgnoreCase))
                return "SUMME";
            return base.GetLocalFunctionName(standardName);
        }
    }
}