using System;
using Aspose.Cells;

namespace AsposeCellsGlobalizationDemo
{
    // Custom globalization settings to provide custom subtotal labels
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override GetTotalName to return a custom label for the Sum function
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            if (functionType == ConsolidationFunction.Sum)
                return "Custom Subtotal";
            // For other functions fall back to the default implementation
            return base.GetTotalName(functionType);
        }
    }

    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range on which to calculate subtotals (e.g., A1:B5)
            // Adjust the range as needed for your data
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0-4, columns 0-1

            // Perform subtotal: group by column 0 (first column), use Sum, and show subtotals
            // Parameters: area, column index to group by, function, array of subtotal columns, replace, pageBreak, summaryBelow
            workbook.Worksheets[0].Cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 1 }, true, false, true);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}