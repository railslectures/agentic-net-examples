using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Define the range for the Subtotal operation (A1:B5)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 1);

        // Apply Subtotal:
        // - group by column 0 (A)
        // - use SUM as the consolidation function for column 1 (B)
        // - include subtotals, grand total, and replace existing data
        cells.Subtotal(area, 0, ConsolidationFunction.Sum, new int[] { 0 }, true, true, true);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}