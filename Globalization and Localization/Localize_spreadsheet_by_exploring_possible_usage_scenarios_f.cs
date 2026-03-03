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
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Globalize subtotal/total labels for regular Subtotal feature
            // ------------------------------------------------------------
            // Create a SettableGlobalizationSettings instance to modify built‑in texts
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Change the label for the grand total of the SUM function
            globalSettings.SetGrandTotalName(ConsolidationFunction.Sum, "Sum Grand Total (Localized)");

            // Change the label for the total (non‑grand) of the SUM function
            globalSettings.SetTotalName(ConsolidationFunction.Sum, "Sum Total (Localized)");

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // Add sample data for Subtotal demonstration
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");
            cells["A2"].PutValue("North");
            cells["B2"].PutValue(1200);
            cells["A3"].PutValue("South");
            cells["B3"].PutValue(800);
            cells["A4"].PutValue("East");
            cells["B4"].PutValue(1500);
            cells["A5"].PutValue("West");
            cells["B5"].PutValue(700);

            // Apply Subtotal: group by column 0 (Category) and calculate SUM on column 1 (Amount)
            // The generated total rows will use the localized labels set above
            CellArea dataRange = CellArea.CreateCellArea(0, 0, 4, 1);
            cells.Subtotal(dataRange, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // ------------------------------------------------------------
            // 2. Globalize Grand Total and Subtotal labels for PivotTable
            // ------------------------------------------------------------
            // Create a SettablePivotGlobalizationSettings instance
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Localize the "Grand Total" label in the pivot table
            pivotSettings.SetTextOfGrandTotal("Grand Total (Localized)");

            // Localize the "Subtotal" label for the SUM type
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Sum Subtotal (Localized)");

            // Assign the pivot globalization settings to the workbook's globalization settings
            globalSettings.PivotSettings = pivotSettings;

            // Build a simple pivot table to demonstrate the localized labels
            // (The same worksheet is used; pivot will be placed starting at D1)
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Row field: Category (column 0)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            // Data field: Amount (column 1) with SUM aggregation
            int dataFieldPos = pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            PivotField dataField = pivotTable.DataFields[dataFieldPos];
            dataField.Function = ConsolidationFunction.Sum;

            // Refresh and calculate the pivot table so that labels appear
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // Save the modified workbook (replace with your desired output path)
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}