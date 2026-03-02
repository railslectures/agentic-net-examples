using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

public class WorkbookExport
{
    public void ExportToStream(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        var saveOptions = new OoxmlSaveOptions();
        workbook.Save(outputStream, saveOptions);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        using var memoryStream = new MemoryStream();
        var exporter = new WorkbookExport();
        exporter.ExportToStream(memoryStream);

        // Optionally write the result to a file for verification
        File.WriteAllBytes("output.xlsx", memoryStream.ToArray());
        Console.WriteLine("Workbook exported successfully.");
    }
}