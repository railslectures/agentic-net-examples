using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsLocalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // -------------------------------------------------
            // 1. Create and configure SettableGlobalizationSettings
            // -------------------------------------------------
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize subtotal label for SUM function (used in Subtotal operations)
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Subtotal");

            // -------------------------------------------------
            // 2. Create and configure SettableChartGlobalizationSettings
            // -------------------------------------------------
            SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();

            // Customize various chart related texts
            chartSettings.SetSeriesName("Custom Series");
            chartSettings.SetChartTitleName("Custom Pie Chart Title");
            chartSettings.SetLegendIncreaseName("Custom Increase");
            chartSettings.SetLegendDecreaseName("Custom Decrease");
            chartSettings.SetLegendTotalName("Custom Total");
            chartSettings.SetOtherName("Custom Other");

            // Assign the chart globalization settings to the main globalization object
            globalization.ChartSettings = chartSettings;

            // -------------------------------------------------
            // 3. Apply the globalization settings to the workbook
            // -------------------------------------------------
            workbook.Settings.GlobalizationSettings = globalization;

            // -------------------------------------------------
            // 4. (Optional) Ensure there is a pie chart to demonstrate the effect
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];

            // If the workbook already contains charts, we skip creation.
            // Otherwise, create a simple pie chart for demonstration.
            if (sheet.Charts.Count == 0)
            {
                // Sample data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(25);

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
                Chart pieChart = sheet.Charts[chartIndex];
                pieChart.NSeries.Add("B2:B4", true);
                pieChart.NSeries.CategoryData = "A2:A4";
                pieChart.Title.Text = "Demo Pie Chart";
            }

            // -------------------------------------------------
            // 5. Save the localized workbook
            // -------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}