using System;
using System.IO;
using Aspose.Cells;

public class ExcelExportHandler
{
    public void ExportWorkbook(Stream outputStream, string fileName = "Sample.xlsx")
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        var saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);
        workbook.Save(outputStream, saveOptions);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        var handler = new ExcelExportHandler();
        using (var fs = new FileStream("Sample.xlsx", FileMode.Create, FileAccess.Write))
        {
            handler.ExportWorkbook(fs);
        }
        Console.WriteLine("Workbook exported.");
    }
}