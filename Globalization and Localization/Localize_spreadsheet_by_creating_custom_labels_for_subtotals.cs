using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Create a SettableGlobalizationSettings instance to customize labels
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Set a custom label for the Sum subtotal (used by Subtotal operation)
        globalization.SetTotalName(ConsolidationFunction.Sum, "My Subtotal");

        // Apply the custom globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Define the range on which to apply the subtotal (e.g., A1:B5)
        // Rows are zero‑based, so rows 0‑4 and columns 0‑1 cover A1:B5
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

        // Apply subtotal:
        // - group by column 0 (first column)
        // - calculate Sum for column 1 (second column)
        // - include subtotal rows, hide detail rows, and keep the original data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // Save the modified workbook with the custom subtotal label
        workbook.Save("output.xlsx");
    }
}