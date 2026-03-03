using System;
using Aspose.Cells;

namespace AsposeCellsCommentDemo
{
    class Program
    {
        static void Main()
        {
            Run();
        }

        static void Run()
        {
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            int commentIndex = worksheet.Comments.Add("A1");
            Comment comment = worksheet.Comments[commentIndex];
            comment.Note = "This is a newly added comment.";
            comment.Author = "Aspose.Cells";

            string outputPath = "output_with_comment.xlsx";
            workbook.Save(outputPath);
        }
    }
}