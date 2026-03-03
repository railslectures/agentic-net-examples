using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // ------------------------------------------------------------
            // 1. Create a SettableGlobalizationSettings instance to customize
            //    subtotal labels (pivot/table totals) and chart labels.
            // ------------------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize the total name for the SUM function (used in subtotals/pivot totals)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // ------------------------------------------------------------
            // 2. Create a SettableChartGlobalizationSettings instance to
            //    customize labels that appear on a pie chart (series name,
            //    legend total, "Other" slice, etc.).
            // ------------------------------------------------------------
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();

            // Example customizations for a pie chart
            chartGlobals.SetSeriesName("Custom Series");               // Name shown for the series
            chartGlobals.SetLegendTotalName("Custom Total");          // Legend entry for the total slice
            chartGlobals.SetOtherName("Other Category");              // Label for the "Other" slice
            chartGlobals.SetChartTitleName("Custom Pie Chart Title"); // Chart title (if used)

            // Assign the chart globalization settings to the main globalization object
            globalization.ChartSettings = chartGlobals;

            // ------------------------------------------------------------
            // 3. Apply the globalization settings to the workbook.
            // ------------------------------------------------------------
            workbook.Settings.GlobalizationSettings = globalization;

            // ------------------------------------------------------------
            // 4. (Optional) If you need to ensure that a pie chart exists,
            //    you can create one here. This step is only for demonstration
            //    and can be omitted if the workbook already contains a chart.
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            // Create a simple pie chart if none exists
            if (sheet.Charts.Count == 0)
            {
                // Add sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
                Chart pieChart = sheet.Charts[chartIndex];
                pieChart.NSeries.Add("B2:B4", true);
                pieChart.NSeries.CategoryData = "A2:A4";
                pieChart.Title.Text = "Demo Pie Chart";
            }

            // ------------------------------------------------------------
            // 5. Save the modified workbook.
            // ------------------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}