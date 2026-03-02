using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ConvertXlsxToCsvWithBlankRowSeparators
    {
        public static void Run()
        {
            // Path to the source Excel file (XLSX)
            string sourcePath = "input.xlsx";

            // Path for the output CSV file
            string destPath = "output.csv";

            try
            {
                // Load the workbook from the XLSX file
                Workbook workbook = new Workbook(sourcePath);

                // Configure CSV save options to keep separators for blank rows
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    // Ensure that a separator (comma) is written for each blank row
                    KeepSeparatorsForBlankRow = true,

                    // Optional: set the separator explicitly (default is comma)
                    Separator = ','
                };

                // Save the workbook as CSV using the configured options
                workbook.Save(destPath, csvOptions);

                Console.WriteLine($"Conversion completed. CSV saved to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ConvertXlsxToCsvWithBlankRowSeparators.Run();
        }
    }
}