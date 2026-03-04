using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Prepare sample data for subtotal and pie chart
            // -------------------------------------------------
            // Header
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Amount");

            // Data rows
            string[] categories = { "Apple", "Banana", "Cherry", "Date", "Elderberry" };
            double[] amounts = { 1200, 850, 430, 670, 290 };

            for (int i = 0; i < categories.Length; i++)
            {
                cells[i + 1, 0].PutValue(categories[i]);   // Column A
                cells[i + 1, 1].PutValue(amounts[i]);     // Column B
            }

            // -------------------------------------------------
            // 2. Apply custom globalization settings
            // -------------------------------------------------
            // Create a SettableGlobalizationSettings instance to customize subtotal label
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
            // Change the total name for SUM function (used by Subtotal)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum");

            // Create a SettableChartGlobalizationSettings instance to customize chart "Other" label
            SettableChartGlobalizationSettings chartGlobalization = new SettableChartGlobalizationSettings();
            chartGlobalization.SetOtherName("Other_Custom");

            // Assign the chart globalization to the main settings
            globalization.ChartSettings = chartGlobalization;

            // Apply the settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // -------------------------------------------------
            // 3. Add a subtotal (using the customized total name)
            // -------------------------------------------------
            // Define the range that includes the data (A1:B6)
            CellArea dataArea = CellArea.CreateCellArea(0, 0, categories.Length, 1);
            // Apply subtotal: group by Category (column 0), sum the Amount (column 1)
            // The last parameter 'true' indicates that the total row will be added
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // -------------------------------------------------
            // 4. Create a pie chart that will use the "Other" label
            // -------------------------------------------------
            // Add a new chart (pie) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 10, 0, 25, 10);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set chart data source: categories as X values, amounts as Y values
            pieChart.NSeries.Add("B2:B6", true);
            pieChart.NSeries.CategoryData = "A2:A6";

            // Set a title for clarity
            pieChart.Title.Text = "Sales Distribution";

            // -------------------------------------------------
            // 5. Save the modified workbook
            // -------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}