using System;
using Aspose.Cells;
using Aspose.Cells.Revisions;

class ReviewTabDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Enable shared workbook mode to allow revision tracking (similar to using the Review tab)
        workbook.Settings.Shared = true;

        // Highlight changes – this mimics the "Track Changes" feature found under the Review tab
        // The first parameter enables highlighting, the second enables showing comments
        workbook.Worksheets.RevisionLogs.HighlightChanges(new HighlightChangesOptions(true, true));

        // Optionally, accept all existing revisions (as if the user clicked "Accept All Changes")
        if (workbook.HasRevisions)
        {
            workbook.AcceptAllRevisions();
        }

        // Save the modified workbook
        string outputPath = "output.xlsx";
        workbook.Save(outputPath, SaveFormat.Xlsx);

        Console.WriteLine($"Workbook processed and saved to '{outputPath}'.");
    }
}