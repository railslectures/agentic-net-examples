using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // -------------------------------------------------
        // 1. Customize chart labels (pie chart example)
        // -------------------------------------------------
        // Create a SettableChartGlobalizationSettings instance
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

        // Set custom texts for chart elements
        chartSettings.SetSeriesName("Custom Series");
        chartSettings.SetChartTitleName("Custom Pie Chart");
        chartSettings.SetLegendTotalName("Custom Total");
        chartSettings.SetOtherName("Other Category");

        // -------------------------------------------------
        // 2. Customize subtotal label for Sum function
        // -------------------------------------------------
        // Create a SettableGlobalizationSettings instance
        SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();

        // Assign the chart globalization settings
        globalSettings.ChartSettings = chartSettings;

        // Set a custom total name for the Sum consolidation function
        globalSettings.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalSettings;

        // -------------------------------------------------
        // 3. (Optional) Refresh chart to ensure labels are applied
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        if (sheet.Charts.Count > 0)
        {
            Chart chart = sheet.Charts[0];
            // Trigger a refresh; the actual label text comes from globalization settings
            chart.Title.Text = chart.Title.Text;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}