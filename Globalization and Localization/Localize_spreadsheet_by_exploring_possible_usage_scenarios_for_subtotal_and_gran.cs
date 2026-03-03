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
            // Load an existing workbook (XLSX format)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // 1. Create a SettableGlobalizationSettings instance to customize
            //    grand total and total labels for standard subtotals.
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

            // Customize the grand total label for SUM function
            globalSettings.SetGrandTotalName(ConsolidationFunction.Sum, "My Custom Grand Total");

            // Customize the total label for SUM function (used in subtotals)
            globalSettings.SetTotalName(ConsolidationFunction.Sum, "My Custom Total");

            // ------------------------------------------------------------
            // 2. Create a SettablePivotGlobalizationSettings instance to
            //    customize the text of PivotField subtotal types.
            // ------------------------------------------------------------
            SettablePivotGlobalizationSettings pivotSettings = new SettablePivotGlobalizationSettings();

            // Customize the text shown for the "Sum" subtotal type
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "My Custom Sum Subtotal");
            // Optionally customize other subtotal types
            pivotSettings.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "My Custom Avg Subtotal");

            // Attach the pivot settings to the main globalization settings
            globalSettings.PivotSettings = pivotSettings;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalSettings;

            // ------------------------------------------------------------
            // 3. Apply a regular Subtotal operation on the worksheet.
            //    The customized total/subtotal labels will be reflected in the
            //    generated total rows.
            // ------------------------------------------------------------
            // Define the range to subtotal (e.g., A1:B5)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1);
            // Subtotal by the first column (index 0) using SUM, add a total row
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // ------------------------------------------------------------
            // 4. Create a PivotTable to demonstrate the localized "Grand Total"
            //    and "Subtotal" labels within a pivot context.
            // ------------------------------------------------------------
            // Ensure there is enough data for the pivot
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
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "MyPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            // Row field: Category
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            // Data field: Amount (SUM)
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1);
            PivotField dataField = pivot.DataFields[dataFieldIdx];
            dataField.Function = ConsolidationFunction.Sum;

            // Refresh and calculate to apply the globalization settings
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 5. Save the modified workbook.
            // ------------------------------------------------------------
            workbook.Save("output_localized.xlsx");
        }
    }
}