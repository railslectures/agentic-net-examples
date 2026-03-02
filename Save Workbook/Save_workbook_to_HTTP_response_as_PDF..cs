using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExamples
{
    public class SaveWorkbookToPdf
    {
        public static byte[] Run()
        {
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Name");
            sheet.Cells["B1"].PutValue("Score");
            sheet.Cells["A2"].PutValue("Alice");
            sheet.Cells["B2"].PutValue(85);
            sheet.Cells["A3"].PutValue("Bob");
            sheet.Cells["B3"].PutValue(92);

            var pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            using (var ms = new MemoryStream())
            {
                workbook.Save(ms, pdfOptions);
                return ms.ToArray();
            }
        }
    }

    class Program
    {
        static void Main()
        {
            byte[] pdfBytes = SaveWorkbookToPdf.Run();
            File.WriteAllBytes("output.pdf", pdfBytes);
            Console.WriteLine("PDF saved as output.pdf");
        }
    }
}