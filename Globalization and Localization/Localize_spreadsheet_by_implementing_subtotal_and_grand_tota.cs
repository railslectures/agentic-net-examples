using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

class LocalizeSubtotalAndGrandTotal
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Create a SettableGlobalizationSettings instance to customize built‑in texts
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

        // Localize the "Grand Total" label for the SUM function (e.g., Japanese)
        globalization.SetGrandTotalName(ConsolidationFunction.Sum, "合計");

        // Localize the generic "Total" label for the SUM function
        globalization.SetTotalName(ConsolidationFunction.Sum, "合計");

        // Create a SettablePivotGlobalizationSettings instance for pivot‑specific texts
        SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();

        // Localize the "Grand Total" label in pivot tables
        pivotGlobalization.SetTextOfGrandTotal("合計合計");

        // Localize the "Subtotal" label for the SUM subtotal type in pivot tables
        pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "小計");

        // Assign the pivot globalization settings to the main globalization object
        globalization.PivotSettings = pivotGlobalization;

        // Apply the customized globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // (Optional) Demonstrate that the settings take effect by creating a subtotal
        // on a range – the label will appear in the localized language when the file is opened.
        Worksheet sheet = workbook.Worksheets[0];
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // A1:B5
        sheet.Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

        // Save the modified workbook
        string outputPath = "output_localized.xlsx";
        workbook.Save(outputPath);
    }
}