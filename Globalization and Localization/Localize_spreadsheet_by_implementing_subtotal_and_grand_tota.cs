using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Settings;

namespace LocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Create globalization settings that allow us to change built‑in texts
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Localize the "Total" label for the Sum function (e.g., French)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Total");

            // Localize the "Grand Total" label for the Sum function
            globalization.SetGrandTotalName(ConsolidationFunction.Sum, "Total Général");

            // ------------------------------------------------------------
            // Create pivot‑specific globalization settings for subtotals
            // ------------------------------------------------------------
            SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();

            // Localize the "Sum" subtotal label
            pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Sous‑total Somme");

            // Localize the "Count" subtotal label
            pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Sous‑total Compte");

            // Attach the pivot globalization to the main settings
            globalization.PivotSettings = pivotGlobalization;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // (Optional) Create a simple pivot table to demonstrate the effect
            // ------------------------------------------------------------
            // Ensure there is data for the pivot table
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("A");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("B");
            cells["B3"].PutValue(200);
            cells["A4"].PutValue("A");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("B");
            cells["B5"].PutValue(250);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "DemoPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field
            PivotField dataField = pivot.DataFields[dataFieldIdx];
            dataField.Function = ConsolidationFunction.Sum; // Use Sum aggregation

            // Refresh and calculate the pivot table so that localized labels appear
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // Save the localized workbook
            // ------------------------------------------------------------
            workbook.Save("output_localized.xlsx");
        }
    }
}