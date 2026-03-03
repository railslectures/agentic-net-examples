using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class LocalizePieChartOtherLabel
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Create a pie chart (optional, for demonstration)
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart pieChart = sheet.Charts[chartIndex];
        // Sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(45);
        sheet.Cells["B4"].PutValue(25);
        pieChart.NSeries.Add("B2:B4", true);
        pieChart.NSeries.CategoryData = "A2:A4";

        // Create chart globalization settings and set custom "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Miscellaneous Items");

        // Create overall globalization settings and assign the chart settings
        SettableGlobalizationSettings globalSettings = new SettableGlobalizationSettings();
        globalSettings.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalSettings;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}