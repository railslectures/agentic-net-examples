using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsGlobalizationDemo
{
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        public CustomGlobalizationSettings()
        {
            var chartSettings = new SettableChartGlobalizationSettings();
            chartSettings.SetOtherName("Otros");
            this.ChartSettings = chartSettings;
        }

        public override string GetTotalName(ConsolidationFunction functionType)
        {
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Sum Total";
            return base.GetTotalName(functionType);
        }
    }

    class Program
    {
        static void Main()
        {
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            CellArea dataArea = CellArea.CreateCellArea(0, 0, 4, 1); // A1:B5
            cells.Subtotal(dataArea, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 10);
            Chart pieChart = sheet.Charts[chartIndex];

            pieChart.NSeries.Add("B1:B5", true);
            pieChart.NSeries.CategoryData = "A1:A5";

            // Show data labels for the first series to display the custom "Other" label.
            if (pieChart.NSeries.Count > 0)
            {
                pieChart.NSeries[0].DataLabels.ShowValue = true;
            }

            workbook.Save("output.xlsx");
        }
    }
}