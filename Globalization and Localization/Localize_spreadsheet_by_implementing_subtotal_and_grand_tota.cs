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
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // ------------------------------------------------------------
            // 1. Create a SettableGlobalizationSettings instance to localize
            //    total and grand total labels for Subtotal operations.
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Localize the total name for the SUM function (e.g., German)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Summe");
            // Localize the grand total name for the SUM function
            globalization.SetGrandTotalName(ConsolidationFunction.Sum, "Gesamtsumme");

            // ------------------------------------------------------------
            // 2. Create a SettablePivotGlobalizationSettings instance to
            //    localize pivot table subtotal and grand total texts.
            // ------------------------------------------------------------
            SettablePivotGlobalizationSettings pivotGlobalization = new SettablePivotGlobalizationSettings();

            // Localize the "Grand Total" label (e.g., French)
            pivotGlobalization.SetTextOfGrandTotal("Total Général");

            // Localize various subtotal types (e.g., Spanish)
            pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Sum, "Suma");
            pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Count, "Recuento");
            pivotGlobalization.SetTextOfSubTotal(PivotFieldSubtotalType.Average, "Promedio");

            // Assign the pivot globalization to the main globalization settings
            globalization.PivotSettings = pivotGlobalization;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // 3. Demonstrate Subtotal operation – the labels will use the
            //    localized total and grand total names defined above.
            // ------------------------------------------------------------
            // Define the range for subtotal (assumes data in columns A and B, rows 1-5)
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);
            // Apply subtotal: group by column 0 (A), sum column 1 (B), include subtotals and grand total
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // ------------------------------------------------------------
            // 4. Create a pivot table to show localized subtotal and grand total
            //    texts in the pivot table.
            // ------------------------------------------------------------
            // Ensure there is enough data for the pivot table
            // (If the loaded workbook already contains data, this step can be omitted)
            // Here we add a simple data set if needed.
            cells["A7"].PutValue("Category");
            cells["B7"].PutValue("Amount");
            cells["A8"].PutValue("A");
            cells["B8"].PutValue(100);
            cells["A9"].PutValue("A");
            cells["B9"].PutValue(150);
            cells["A10"].PutValue("B");
            cells["B10"].PutValue(200);
            cells["A11"].PutValue("B");
            cells["B11"].PutValue(250);

            // Add a pivot table based on the data range A7:B11, place it at D7
            int pivotIndex = worksheet.PivotTables.Add("A7:B11", "D7", "LocalizedPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            int dataFieldIndex = pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Amount as data field
            PivotField dataField = pivotTable.DataFields[dataFieldIndex];
            dataField.Function = ConsolidationFunction.Sum; // Summarize by Sum

            // Refresh and calculate the pivot table to apply localization
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ------------------------------------------------------------
            // 5. Save the modified workbook
            // ------------------------------------------------------------
            workbook.Save("output_localized.xlsx");
        }
    }
}