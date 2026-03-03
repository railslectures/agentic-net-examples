using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGlobalizationDemo
{
    // Custom globalization settings to localize subtotal total name and chart labels
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public CustomGlobalizationSettings()
        {
            // Initialize chart globalization settings and customize the "Other" label
            var chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetOtherName("Other (Custom)");
            this.ChartSettings = chartSettings;
        }

        // Override the total name used in subtotals (e.g., "Sum of Sales")
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            // Provide a custom label for the SUM function; fallback to base for others
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Sum Total";
            return base.GetTotalName(functionType);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with actual path if needed)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // -----------------------------------------------------------------
            // Scenario 1: Subtotal with localized total name
            // -----------------------------------------------------------------
            // Define the data range (A1:B5) for subtotal operation
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1);
            // Perform subtotal: group by column 0 (Region), sum column 1 (Sales)
            // The resulting total row will display the custom total name defined above
            sheet.Cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // -----------------------------------------------------------------
            // Scenario 2: Pie chart with customized "Other" label
            // -----------------------------------------------------------------
            // Add a pie chart to visualize the sales data
            int chartIdx = sheet.Charts.Add(ChartType.Pie, 10, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIdx];

            // Set data series and categories
            pieChart.NSeries.Add("B2:B5", true);
            pieChart.NSeries.CategoryData = "A2:A5";

            // Set chart title
            pieChart.Title.Text = "Sales Distribution";

            // The "Other" slice label (if applicable) will use the custom text
            // defined in SettableChartGlobalizationSettings (i.e., "Other (Custom)")

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}