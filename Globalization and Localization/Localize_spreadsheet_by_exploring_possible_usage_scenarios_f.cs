using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGlobalizationDemo
{
    // Custom globalization settings that configure subtotal total name and chart "Other" label
    public class CustomGlobalizationSettings : SettableGlobalizationSettings
    {
        public CustomGlobalizationSettings()
        {
            // Set a custom total name for the SUM function used in subtotals
            SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // Create chart-specific globalization settings
            var chartSettings = new SettableChartGlobalizationSettings();

            // Set a custom label for the "Other" category that appears in pie charts
            chartSettings.SetOtherName("Other Category");

            // Assign the chart settings to the parent globalization settings
            this.ChartSettings = chartSettings;
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // -----------------------------------------------------------------
            // 1. Demonstrate subtotal with customized total name
            // -----------------------------------------------------------------
            // Assume data is in A1:B5 (headers in row 1)
            // Add a subtotal that groups by the first column and calculates SUM on the second column
            CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0-4, cols 0-1
            // Parameters: area, column index to group by, function, totalColumns, replace, pageBreaks, summaryBelowData
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // Retrieve and display the custom total name for verification
            var gSettings = (SettableGlobalizationSettings)workbook.Settings.GlobalizationSettings;
            string customTotalName = gSettings.GetTotalName(ConsolidationFunction.Sum);
            Console.WriteLine("Custom total name for SUM: " + customTotalName);

            // -----------------------------------------------------------------
            // 2. Create a pie chart that will use the custom "Other" label
            // -----------------------------------------------------------------
            // Add sample data for the chart (categories in C1:C5, values in D1:D5)
            cells["C1"].PutValue("Category");
            cells["D1"].PutValue("Value");
            cells["C2"].PutValue("A");
            cells["C3"].PutValue("B");
            cells["C4"].PutValue("C");
            cells["C5"].PutValue("D");
            cells["D2"].PutValue(30);
            cells["D3"].PutValue(20);
            cells["D4"].PutValue(15);
            cells["D5"].PutValue(35);

            // Add a pie chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 7, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIndex];

            // Set the data series (values) and category labels
            pieChart.NSeries.Add("D2:D5", true);
            pieChart.NSeries.CategoryData = "C2:C5";

            // Set a chart title
            pieChart.Title.Text = "Sample Pie Chart";

            // The chart will display the custom "Other" label if the chart groups data,
            // but we can demonstrate that the setting is applied by reading it back:
            var chartGlobalSettings = ((SettableGlobalizationSettings)workbook.Settings.GlobalizationSettings).ChartSettings;
            string otherLabel = chartGlobalSettings.GetOtherName();
            Console.WriteLine("Custom chart 'Other' label: " + otherLabel);

            // -----------------------------------------------------------------
            // Save the modified workbook
            // -----------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}