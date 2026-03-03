using System;
using Aspose.Cells;

namespace ThreadedCommentsReader
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the input XLSX file containing threaded comments
            string inputPath = "Input.xlsx";

            // Load the workbook (lifecycle rule: load)
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet: {sheet.Name}");

                // Access the comment collection of the current worksheet
                CommentCollection comments = sheet.Comments;

                // Loop through each comment in the collection
                foreach (Comment comment in comments)
                {
                    // Check if the comment is a threaded comment
                    if (comment.IsThreadedComment)
                    {
                        // Get the cell address of the comment
                        string cellName = CellsHelper.CellIndexToName(comment.Row, comment.Column);
                        Console.WriteLine($"  Threaded comment at cell {cellName}:");

                        // Retrieve all threaded comments for this cell
                        ThreadedCommentCollection threadedComments = comments.GetThreadedComments(comment.Row, comment.Column);

                        // Iterate through the threaded comments and display details
                        foreach (ThreadedComment tc in threadedComments)
                        {
                            Console.WriteLine($"    Author: {tc.Author.Name}");
                            Console.WriteLine($"    Notes : {tc.Notes}");
                            Console.WriteLine($"    Row   : {tc.Row}, Column: {tc.Column}");
                            Console.WriteLine($"    Created: {tc.CreatedTime}");
                        }
                    }
                }
            }

            // (Optional) Save the workbook after processing if modifications were made
            // workbook.Save("Output.xlsx");
        }
    }
}