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
            // 1. Create a SettableGlobalizationSettings instance to customize
            //    total and grand total labels for standard subtotals.
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize the total name for the SUM function (used by Cells.Subtotal)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // Customize the grand total name for the SUM function (used by PivotTables)
            globalization.SetGrandTotalName(ConsolidationFunction.Sum, "Custom Grand Sum");

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // 2. Apply a regular subtotal on a data range.
            //    The total row will display the custom total name set above.
            // ------------------------------------------------------------
            // Assume data is in A1:B6 (headers + 5 rows of data)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 5, 1);
            // Subtotal on column 0 (Region) using SUM on column 1 (Sales)
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // ------------------------------------------------------------
            // 3. Create a PivotTable to demonstrate the custom grand total label.
            // ------------------------------------------------------------
            // Add sample data if the loaded workbook does not contain any.
            // (This block can be removed if the source file already has data.)
            if (cells["A1"].Value == null)
            {
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
            }

            // Define the source range for the pivot table
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            int dataFieldIdx = pivot.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field
            PivotField dataField = pivot.DataFields[dataFieldIdx];
            dataField.Function = ConsolidationFunction.Sum; // Use SUM aggregation

            // Refresh and calculate to apply the custom grand total label
            pivot.RefreshData();
            pivot.CalculateData();

            // ------------------------------------------------------------
            // 4. Save the modified workbook.
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}