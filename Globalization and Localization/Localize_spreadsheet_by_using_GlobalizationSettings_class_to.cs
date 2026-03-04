using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsLocalizationDemo
{
    // Custom globalization settings that combines SettableGlobalizationSettings
    // with chart-specific settings.
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public CustomGlobalizationSettings(SettableGlobalizationSettings baseSettings,
                                           SettableChartGlobalizationSettings chartSettings)
        {
            // Apply the base (subtotal) settings.
            this.PivotSettings = baseSettings.PivotSettings;
            this.ChartSettings = chartSettings;

            // Copy other possible settings from baseSettings if needed.
            // Since SettableGlobalizationSettings derives from GlobalizationSettings,
            // we can assign it directly.
            // However, WorkbookSettings expects a GlobalizationSettings instance,
            // so we expose the configured baseSettings via this wrapper.
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook.
            Workbook workbook = new Workbook("input.xlsx");

            // ------------------------------------------------------------
            // 1. Configure subtotal (total) label customization.
            // ------------------------------------------------------------
            // Create a SettableGlobalizationSettings instance.
            SettableGlobalizationSettings subtotalSettings = new SettableGlobalizationSettings();

            // Customize the total name for the SUM function (used in Subtotal).
            subtotalSettings.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // Optionally customize other function totals.
            subtotalSettings.SetTotalName(ConsolidationFunction.Average, "Custom Average Total");

            // ------------------------------------------------------------
            // 2. Configure pie chart label customization.
            // ------------------------------------------------------------
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

            // Set custom series name (appears in legend).
            chartSettings.SetSeriesName("Custom Series");

            // Set custom chart title.
            chartSettings.SetChartTitleName("Custom Pie Chart Title");

            // Set custom legend total name (used for "Total" entry in legend, if applicable).
            chartSettings.SetLegendTotalName("Custom Total");

            // Set custom "Other" label (used when chart aggregates small slices).
            chartSettings.SetOtherName("Other (Custom)");

            // ------------------------------------------------------------
            // 3. Combine the settings and apply to the workbook.
            // ------------------------------------------------------------
            // Create a wrapper that holds both subtotal and chart settings.
            // Here we directly assign the SettableGlobalizationSettings instance
            // to the workbook and then replace its ChartSettings.
            workbook.Settings.GlobalizationSettings = subtotalSettings;
            // Attach the chart-specific settings.
            subtotalSettings.ChartSettings = chartSettings;

            // ------------------------------------------------------------
            // 4. (Optional) Demonstrate that the chart picks up the settings.
            // ------------------------------------------------------------
            // Assume the workbook already contains a pie chart on the first worksheet.
            // If not, we create a simple pie chart to illustrate the effect.
            Worksheet sheet = workbook.Worksheets[0];

            // Check if there is at least one chart; if not, add one.
            if (sheet.Charts.Count == 0)
            {
                // Add sample data for the pie chart.
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // Add a pie chart.
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
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