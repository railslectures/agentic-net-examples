using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Create a SettableGlobalizationSettings instance to customize built‑in texts
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // ----- Grand Total Localization -----
            // Change the grand total label for the SUM function (e.g., Japanese "合計")
            globalization.SetGrandTotalName(ConsolidationFunction.Sum, "合計");

            // Change the total label for the SUM function (used in subtotals outside pivot tables)
            globalization.SetTotalName(ConsolidationFunction.Sum, "合計");

            // ----- Subtotal Localization for Pivot Tables -----
            // Create a SettablePivotGlobalizationSettings instance and assign it to the PivotSettings property
            SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();

            // Change the text displayed for the "Sum" subtotal type in pivot tables (e.g., Japanese "小計")
            pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "小計");

            // Assign the pivot globalization settings to the main globalization object
            globalization.PivotSettings = pivotGlobalization;

            // Apply the customized globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}