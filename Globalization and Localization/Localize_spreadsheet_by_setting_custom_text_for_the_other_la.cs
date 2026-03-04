using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create chart globalization settings and set custom text for the "Other" label
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Miscellaneous Items");

        // Combine chart settings with overall globalization settings
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // Ensure there is a pie chart in the workbook to demonstrate the custom "Other" label
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the pie chart (if not already present)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["A5"].PutValue("D");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(25);
        sheet.Cells["B5"].PutValue(25);

        // Add a pie chart that uses the data
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 15, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}