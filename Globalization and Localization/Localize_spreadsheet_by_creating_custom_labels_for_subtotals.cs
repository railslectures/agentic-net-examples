using System;
using Aspose.Cells;

public class CustomGlobalizationSettings : GlobalizationSettings
{
    // Override the total name for specific consolidation functions
    public override string GetTotalName(ConsolidationFunction functionType)
    {
        switch (functionType)
        {
            case ConsolidationFunction.Sum:
                return "Subtotal Sum";
            case ConsolidationFunction.Count:
                return "Subtotal Count";
            case ConsolidationFunction.Average:
                return "Subtotal Average";
            default:
                return base.GetTotalName(functionType);
        }
    }
}

public class Program
{
    public static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Assign custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

        // Define the cell area on which to apply subtotals (e.g., A1:B5)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0‑4, columns 0‑1

        // Apply subtotal:
        //   - Group by column 0 (first column)
        //   - Use SUM as the consolidation function
        //   - Add subtotal for column 1 (second column)
        //   - Replace existing subtotals, no page break, place summary below
        workbook.Worksheets[0].Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}