using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Assign custom globalization settings that localize subtotal and grand‑total labels
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the range on which to apply Subtotal (e.g., A1:B5)
        // Rows are zero‑based, so rows 0‑4 and columns 0‑1 correspond to A1:B5
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

        // Apply Subtotal:
        //   - group by column 0 (first column)
        //   - use Sum as the consolidation function
        //   - include subtotals and grand total
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, true, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }

    // Custom globalization settings that override the default labels
    class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override subtotal label for specific subtotal types
        public override string GetSubTotalName(PivotFieldSubtotalType subTotalType)
        {
            switch (subTotalType)
            {
                case PivotFieldSubtotalType.Sum:
                    return "Custom Subtotal (Sum)";
                case PivotFieldSubtotalType.Count:
                    return "Custom Subtotal (Count)";
                case PivotFieldSubtotalType.Average:
                    return "Custom Subtotal (Avg)";
                default:
                    return base.GetSubTotalName(subTotalType);
            }
        }

        // Override grand‑total label based on the consolidation function
        public override string GetGrandTotalName(ConsolidationFunction functionType)
        {
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Grand Total (Sum)";
            if (functionType == ConsolidationFunction.Count)
                return "Custom Grand Total (Count)";
            return base.GetGrandTotalName(functionType);
        }
    }
}