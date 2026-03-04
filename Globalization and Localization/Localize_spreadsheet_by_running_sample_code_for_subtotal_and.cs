using System;
using Aspose.Cells;

namespace AsposeCellsLocalizationDemo
{
    public class Program
    {
        public static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the range on which to apply Subtotal (example: A1:B5)
            CellArea area = CellArea.CreateCellArea(0, 0, 4, 1); // rows 0-4, columns 0-1 (A1:B5)

            // Apply Subtotal:
            //   - group by column 0 (A)
            //   - use Sum function
            //   - replace existing subtotals: true
            //   - page break between groups: false
            //   - summary below data: true
            cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, false, true);

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}