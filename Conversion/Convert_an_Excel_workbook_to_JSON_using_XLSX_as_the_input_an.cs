using System;
using Aspose.Cells;

namespace AsposeCellsJsonConversion
{
    public class ExcelToJsonConverter
    {
        public static void Run()
        {
            string sourcePath = "input.xlsx";
            string jsonPath = "output.json";

            Workbook workbook = new Workbook(sourcePath);

            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportEmptyCells = true,
                HasHeaderRow = true
            };

            workbook.Save(jsonPath, jsonOptions);

            Console.WriteLine($"Excel workbook '{sourcePath}' has been successfully converted to JSON at '{jsonPath}'.");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExcelToJsonConverter.Run();
        }
    }
}