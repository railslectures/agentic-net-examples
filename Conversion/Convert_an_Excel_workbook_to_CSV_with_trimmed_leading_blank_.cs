using System;
using Aspose.Cells;

public class ExcelToCsvConverter
{
    /// <summary>
    /// Converts an XLSX workbook to CSV while trimming leading blank rows and columns.
    /// </summary>
    /// <param name="inputXlsxPath">Full path to the source XLSX file.</param>
    /// <param name="outputCsvPath">Full path where the CSV file will be saved.</param>
    public static void Convert(string inputXlsxPath, string outputCsvPath)
    {
        // Load the Excel workbook from the specified file.
        Workbook workbook = new Workbook(inputXlsxPath);

        // Configure CSV save options to trim leading blank rows and columns.
        TxtSaveOptions saveOptions = new TxtSaveOptions
        {
            TrimLeadingBlankRowAndColumn = true
        };

        // Save the workbook as CSV using the configured options.
        workbook.Save(outputCsvPath, saveOptions);
    }

    // Example usage.
    public static void Main()
    {
        string sourceFile = "input.xlsx";   // Replace with your actual XLSX file path.
        string destFile   = "output.csv";   // Desired CSV output path.

        try
        {
            Convert(sourceFile, destFile);
            Console.WriteLine($"Conversion completed: '{sourceFile}' -> '{destFile}'");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}