using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class LocalizePieChartOtherLabel
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create chart globalization settings and set custom text for the "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Custom Other");

        // Create a SettableGlobalizationSettings instance and assign the chart settings
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // (Optional) Create a pie chart to demonstrate the effect
        Worksheet sheet = workbook.Worksheets[0];
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart pieChart = sheet.Charts[chartIndex];
        // Sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(10);
        // Add data series; the "Other" slice will use the custom label
        pieChart.NSeries.Add("B2:B4", true);
        pieChart.NSeries.CategoryData = "A2:A4";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}