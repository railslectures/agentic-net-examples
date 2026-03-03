using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class LocalizePieChartOtherLabel
{
    static void Main()
    {
        // Load an existing workbook (or create a new one if the file does not exist)
        // Replace "input.xlsx" with the path to your source file.
        Workbook workbook = new Workbook("input.xlsx");

        // -------------------------------------------------
        // Prepare data for a pie chart (if the sheet does not already contain it)
        // -------------------------------------------------
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Apples");
        sheet.Cells["A3"].PutValue("Bananas");
        sheet.Cells["A4"].PutValue("Other"); // This will be aggregated as "Other" in the chart
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(40);
        sheet.Cells["B3"].PutValue(30);
        sheet.Cells["B4"].PutValue(30);

        // Add a pie chart if it does not already exist
        // (Assumes no chart at index 0; adjust as needed)
        int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 10);
        Chart pieChart = sheet.Charts[chartIndex];
        pieChart.NSeries.Add("B2:B4", true);
        pieChart.NSeries.CategoryData = "A2:A4";
        pieChart.Title.Text = "Fruit Distribution";

        // -------------------------------------------------
        // Create globalization settings and set custom "Other" label
        // -------------------------------------------------
        // Create chart-specific globalization settings
        SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
        chartSettings.SetOtherName("Miscellaneous Items"); // Custom text for "Other" label

        // Create the top‑level globalization settings and attach the chart settings
        SettableGlobalizationSettings globalization = new SettableGlobalizationSettings();
        globalization.ChartSettings = chartSettings;

        // Apply the globalization settings to the workbook
        workbook.Settings.GlobalizationSettings = globalization;

        // -------------------------------------------------
        // Save the modified workbook
        // -------------------------------------------------
        // Replace "output.xlsx" with the desired output path.
        workbook.Save("output.xlsx");
    }
}