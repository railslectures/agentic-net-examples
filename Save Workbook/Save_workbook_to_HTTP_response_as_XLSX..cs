using System;
using System.IO;
using System.Threading.Tasks;
using Aspose.Cells;

public static class WorkbookExport
{
    public static async Task ExportToStreamAsync(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        workbook.Save(outputStream, SaveFormat.Xlsx);
        await outputStream.FlushAsync();
    }
}

public class Program
{
    public static async Task Main(string[] args)
    {
        using var fileStream = new FileStream("output.xlsx", FileMode.Create, FileAccess.Write);
        await WorkbookExport.ExportToStreamAsync(fileStream);
    }
}