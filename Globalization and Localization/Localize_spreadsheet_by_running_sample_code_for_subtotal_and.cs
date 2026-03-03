using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings to localize subtotal and grand‑total labels
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize the text shown for the grand‑total row (e.g., in Subtotal operation)
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            // Example: replace the default English label with a French one
            return functionType == ConsolidationFunction.Sum ? "Total Général" : base.GetGrandTotalName(functionType);
        }

        // Localize the text shown for the total row produced by Subtotal (uses GetTotalName)
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Example: replace the default English label with a German one
            return functionType == ConsolidationFunction.Sum ? "Zwischensumme" : base.GetTotalName(functionType);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range on which to calculate subtotals (adjust as needed)
            // Here we assume data occupies A1:B5
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1

            // Perform subtotal: group by column 0 (A) and calculate Sum on column 1 (B)
            // The resulting total rows will display the localized labels defined above
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}