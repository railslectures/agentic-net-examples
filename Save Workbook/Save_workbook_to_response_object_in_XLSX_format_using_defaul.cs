using Aspose.Cells;
using Aspose.Cells.Saving;
using System;
using System.IO;

public class WorkbookResponseExample
{
    public void ExportWorkbook(Stream outputStream, string fileName = "Sample.xlsx")
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
        workbook.Save(outputStream, saveOptions);

        if (outputStream.CanSeek)
        {
            outputStream.Position = 0;
        }
    }
}

public class Program
{
    public static void Main()
    {
        var example = new WorkbookResponseExample();
        using (var ms = new MemoryStream())
        {
            example.ExportWorkbook(ms);
            File.WriteAllBytes("Sample.xlsx", ms.ToArray());
        }

        Console.WriteLine("Workbook exported successfully.");
    }
}