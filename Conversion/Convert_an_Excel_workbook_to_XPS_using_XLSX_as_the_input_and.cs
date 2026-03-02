using System;
using Aspose.Cells;

namespace AsposeCellsConversionDemo
{
    public class XlsxToXpsConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string destinationPath = "output.xps";

            try
            {
                Workbook workbook = new Workbook(sourcePath);
                workbook.Save(destinationPath, SaveFormat.Xps);
                Console.WriteLine($"Conversion successful: '{sourcePath}' → '{destinationPath}'");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }

        public static void Main(string[] args)
        {
            Run();
        }
    }
}