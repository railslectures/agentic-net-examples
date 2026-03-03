using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create a SettableGlobalizationSettings instance to customize labels
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Define a custom label for the subtotal of the SUM function
        // This label will appear instead of the default "Sum"
        globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Subtotal");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Get the first worksheet and its cells collection
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define the range on which the subtotal will be applied (A1:B5)
        // Rows are zero‑based, so row 0‑4 and columns 0‑1 cover A1:B5
        CellArea range = CellArea.CreateCellArea(0, 0, 4, 1);

        // Apply subtotal:
        // - Group by column 0 (the first column)
        // - Use SUM as the consolidation function
        // - Calculate totals for column 1 (the second column)
        // - Replace existing subtotals if any, no page break, place summary below data
        cells.Subtotal(range, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

        // Save the modified workbook with the custom subtotal label
        workbook.Save("output.xlsx");
    }
}