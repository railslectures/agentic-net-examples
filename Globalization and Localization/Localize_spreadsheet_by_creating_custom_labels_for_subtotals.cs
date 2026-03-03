using System;
using Aspose.Cells;

namespace AsposeCellsGlobalizationDemo
{
    // Custom globalization settings to provide custom subtotal labels
    public class CustomGlobalizationSettings : GlobalizationSettings
    {
        // Override GetTotalName to return a custom label for each consolidation function
        public override string GetTotalName(ConsolidationFunction functionType)
        {
            switch (functionType)
            {
                case ConsolidationFunction.Sum:
                    return "Custom Sum Total";
                case ConsolidationFunction.Average:
                    return "Custom Average Total";
                case ConsolidationFunction.Count:
                    return "Custom Count Total";
                case ConsolidationFunction.Max:
                    return "Custom Max Total";
                case ConsolidationFunction.Min:
                    return "Custom Min Total";
                default:
                    // Fallback to the base implementation for any other functions
                    return base.GetTotalName(functionType);
            }
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Get the first worksheet (you can adjust the index as needed)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Apply the custom globalization settings to the workbook
            workbook.Settings.GlobalizationSettings = new CustomGlobalizationSettings();

            // Define the range on which to apply subtotals.
            // Here we use the used range of columns A and B (0‑based indexes).
            int firstRow = cells.MinRow;
            int firstColumn = cells.MinColumn;
            int lastRow = cells.MaxDataRow;
            int lastColumn = cells.MaxDataColumn;

            // Create a CellArea covering the used range
            CellArea area = CellArea.CreateCellArea(firstRow, firstColumn, lastRow, lastColumn);

            // Apply subtotal:
            // - group by column 0 (first column)
            // - use Sum function for the subtotal
            // - include column 0 in the subtotal calculation (new int[] { 0 })
            // - replace existing subtotals: true
            // - use outline style: false
            // - page break after each subtotal: true
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}