using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Apply custom globalization settings to localize subtotal and grand total labels
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the range on which to apply subtotal (example assumes data in A1:B5)
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1

        // Apply subtotal:
        //   - group by column 0 (first column)
        //   - use Sum function on the data column
        //   - show subtotal and grand total labels
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings to provide localized labels
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Localize the Grand Total label for the Sum function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            // Example: French localization
            return "Total Général";
        }

        // Localize Subtotal labels (obsolete method but used by Subtotal operation)
        public override string GetSubTotalName(PivotFieldSubtotalType subTotalType)
        {
            switch (subTotalType)
            {
                case PivotFieldSubtotalType.Sum:
                    return "Sous‑total Somme";
                case PivotFieldSubtotalType.Count:
                    return "Sous‑total Compte";
                case PivotFieldSubtotalType.Average:
                    return "Sous‑total Moyenne";
                case PivotFieldSubtotalType.Max:
                    return "Sous‑total Max";
                case PivotFieldSubtotalType.Min:
                    return "Sous‑total Min";
                default:
                    return base.GetSubTotalName(subTotalType);
            }
        }
    }
}