using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentsReader
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook workbook = new Workbook("InputWorkbook.xlsx");
            Worksheet worksheet = workbook.Worksheets[0];
            CommentCollection comments = worksheet.Comments;

            foreach (Comment comment in comments)
            {
                string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                Console.WriteLine($"Comment for cell {cellName}:");
                Console.WriteLine($"- {comment.Note} (by {comment.Author})");
            }

            Console.WriteLine("Finished reading comments.");
            Console.ReadKey();
        }
    }
}