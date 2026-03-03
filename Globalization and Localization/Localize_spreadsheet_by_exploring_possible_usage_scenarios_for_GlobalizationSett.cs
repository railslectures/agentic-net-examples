using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace GlobalizationDemo
{
    // Custom globalization settings to change the subtotal total name
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Provide a custom label for the Sum total in subtotals
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Sum Total";
            return base.GetTotalName(functionType);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 1. Configure globalization settings for subtotal labels
            // -----------------------------------------------------------------
            var customSettings = new CustomGlobalizationSettings();

            // -----------------------------------------------------------------
            // 2. Configure chart globalization settings (e.g., "Other" label)
            // -----------------------------------------------------------------
            var chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetOtherName("Other Category"); // custom label for "Other" in charts
            customSettings.ChartSettings = chartSettings;

            // Apply the combined globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = customSettings;

            // -----------------------------------------------------------------
            // 3. Add a subtotal to demonstrate the custom total name
            // -----------------------------------------------------------------
            // Define the data area (assumes data starts at A1)
            int lastRow = sheet.Cells.MaxDataRow;
            int lastCol = sheet.Cells.MaxDataColumn;
            CellArea dataArea = CellArea.CreateCellArea(0, 0, lastRow, lastCol);

            // Subtotal on the first column (index 0) using Sum, grouping by column 0
            sheet.Cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // -----------------------------------------------------------------
            // 4. Create a pie chart to show the custom "Other" label usage
            // -----------------------------------------------------------------
            // Ensure there is some sample data for the chart
            sheet.Cells["A2"].PutValue("Category A");
            sheet.Cells["A3"].PutValue("Category B");
            sheet.Cells["A4"].PutValue("Category C");
            sheet.Cells["A5"].PutValue("Category D");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(25);
            sheet.Cells["B5"].PutValue(25);

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 10, 0, 20, 5);
            Chart pieChart = sheet.Charts[chartIndex];
            pieChart.NSeries.Add("B2:B5", true);
            pieChart.NSeries.CategoryData = "A2:A5";
            pieChart.Title.Text = "Sales Distribution";

            // -----------------------------------------------------------------
            // 5. Save the modified workbook
            // -----------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}