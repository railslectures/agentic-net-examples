using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input.xlsx";
            string destPath = "output.xps";

            try
            {
                Workbook workbook = new Workbook(sourcePath);
                workbook.Save(destPath, SaveFormat.Xps);
                Console.WriteLine($"Conversion successful: '{sourcePath}' -> '{destPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}