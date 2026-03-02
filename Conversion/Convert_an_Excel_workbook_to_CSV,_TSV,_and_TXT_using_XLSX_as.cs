using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string sourcePath = "input.xlsx";

        // Load the workbook (create/load rule)
        Workbook workbook = new Workbook(sourcePath);

        // Convert to CSV using the Save method with SaveFormat.Csv
        workbook.Save("output.csv", SaveFormat.Csv);

        // Convert to TSV using the Save method with SaveFormat.Tsv
        workbook.Save("output.tsv", SaveFormat.Tsv);

        // Convert to a generic TXT file (space‑separated) using TxtSaveOptions
        TxtSaveOptions txtOptions = new TxtSaveOptions(SaveFormat.Csv); // base format
        txtOptions.SeparatorString = " "; // custom separator for TXT
        workbook.Save("output.txt", txtOptions);

        Console.WriteLine("Workbook successfully converted to CSV, TSV, and TXT.");
    }
}