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
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Create custom globalization settings for regular subtotals
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Customize the label for the grand total of the SUM function
            globalSettings.SetGrandTotalName(ConsolidationFunction.Sum, "My Grand Sum");

            // Customize the label for the total (subtotal row) of the SUM function
            globalSettings.SetTotalName(ConsolidationFunction.Sum, "My Total Sum");

            // ------------------------------------------------------------
            // 2. Create custom pivot table globalization settings
            // ------------------------------------------------------------
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Change the text displayed for the pivot table's "Grand Total" label
            pivotSettings.SetTextOfGrandTotal("My Pivot Grand Total");

            // Change the text displayed for a specific subtotal type in the pivot table
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "My Pivot Subtotal Sum");

            // Assign the settings to the workbook
            globalSettings.PivotSettings = pivotSettings;
            workbook.Settings.GlobalizationSettings = globalSettings;

            // ------------------------------------------------------------
            // 3. Apply Subtotal operation to demonstrate localized total labels
            // ------------------------------------------------------------
            // Define the data range (A1:B5) – adjust as needed for your data
            CellArea dataRange = CellArea.CreateCellArea(0, 0, 4, 1);

            // Apply subtotal:
            //   - group by column 0 (Region)
            //   - use SUM as the consolidation function
            //   - include column 0 in the subtotal
            //   - replace existing subtotals, keep grand total, and keep data
            cells.Subtotal(dataRange, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // ------------------------------------------------------------
            // 4. Create a pivot table to demonstrate localized pivot labels
            // ------------------------------------------------------------
            // Add a pivot table based on the same data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D1", "MyPivotTable");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Region as row field
            int dataFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Sales as data field
            PivotField dataField = pivotTable.DataFields[dataFieldIdx];
            dataField.Function = ConsolidationFunction.Sum; // Ensure SUM is used

            // Refresh and calculate to apply the custom labels
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // 5. Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}