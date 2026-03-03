using System;
using Aspose.Cells;

class ThreadedCommentExample
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Create a threaded comment author
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add("Demo Author", "demo_user", "demo_provider");
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell B2 (row 1, column 1)
        worksheet.Comments.AddThreadedComment(1, 1, "This is a threaded comment.", author);

        // Retrieve and display the threaded comments for verification
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments(1, 1);
        foreach (ThreadedComment comment in threadedComments)
        {
            Console.WriteLine($"Comment at ({comment.Row}, {comment.Column}): {comment.Notes} (Author: {comment.Author.Name})");
        }

        // Save the workbook with the new threaded comment
        workbook.Save("Output.xlsx");
    }
}