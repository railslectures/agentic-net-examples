using System;
using System.IO;
using Aspose.Cells;

public class WorkbookResponseSaver
{
    // Saves a newly created workbook to the provided output stream
    // using default OOXML (XLSX) save options.
    public void SaveWorkbookToStream(Stream outputStream)
    {
        // Create a new workbook and add some sample data.
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        // Default save options for OOXML (XLSX) format.
        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

        // Write the workbook to the provided stream.
        workbook.Save(outputStream, saveOptions);
    }

    // Entry point for the console application.
    public static void Main()
    {
        var saver = new WorkbookResponseSaver();

        // Save to a file named "output.xlsx" in the current directory.
        using (var fileStream = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write))
        {
            saver.SaveWorkbookToStream(fileStream);
        }

        Console.WriteLine("Workbook saved to output.xlsx");
    }
}