using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

public class WorkbookPdfExport
{
    public void ExportToPdf(Stream outputStream)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello");
        sheet.Cells["B1"].PutValue("World");

        var pdfOptions = new PdfSaveOptions();
        workbook.Save(outputStream, pdfOptions);
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        using (var outputStream = new FileStream("output.pdf", FileMode.Create, FileAccess.Write))
        {
            var exporter = new WorkbookPdfExport();
            exporter.ExportToPdf(outputStream);
        }

        Console.WriteLine("PDF exported successfully.");
    }
}