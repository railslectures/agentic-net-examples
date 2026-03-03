using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Settings;

namespace AsposeCellsGlobalizationDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Create and configure globalization settings
            // -------------------------------------------------
            // SettableGlobalizationSettings allows us to change built‑in texts such as total labels
            SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();

            // Customize the subtotal/total label for the SUM function
            globalization.SetTotalName(ConsolidationFunction.Sum, "Custom Sum Total");

            // -------------------------------------------------
            // Customize chart related texts
            // -------------------------------------------------
            // SettableChartGlobalizationSettings provides setters for chart labels
            SettableChartGlobalizationSettings chartGlobals = new SettableChartGlobalizationSettings();

            // Change the default series name shown in legends and data labels
            chartGlobals.SetSeriesName("Custom Series");

            // Change the text used for the "Total" entry in a pie chart legend
            chartGlobals.SetLegendTotalName("Custom Total");

            // Change the label used for the "Other" slice (values that don't fit into other categories)
            chartGlobals.SetOtherName("Other Category");

            // Assign the chart globalization settings to the main globalization object
            globalization.ChartSettings = chartGlobals;

            // Apply the globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = globalization;

            // -------------------------------------------------
            // (Optional) Add a pie chart to demonstrate the settings
            // -------------------------------------------------
            // Prepare sample data for the pie chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["A4"].PutValue("Cherry");
            sheet.Cells["B4"].PutValue(25);

            // Add a pie chart (if the workbook already contains one, this will just add another)
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
            Chart pieChart = sheet.Charts[chartIndex];
            pieChart.NSeries.Add("B2:B4", true);
            pieChart.NSeries.CategoryData = "A2:A4";
            pieChart.Title.Text = "Demo Pie Chart";

            // -------------------------------------------------
            // Save the modified workbook
            // -------------------------------------------------
            workbook.Save("output.xlsx");
        }
    }
}