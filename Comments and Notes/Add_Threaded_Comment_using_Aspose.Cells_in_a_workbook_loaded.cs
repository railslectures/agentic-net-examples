using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a threaded comment author to the workbook
        int authorIndex = workbook.Worksheets.ThreadedCommentAuthors.Add(
            "John Doe",               // Author name
            "john.doe@example.com",   // User ID (e.g., email)
            "EXAMPLE_PROVIDER");      // Provider ID
        ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIndex];

        // Add a threaded comment to cell B2 using the author
        worksheet.Comments.AddThreadedComment("B2", "This is a threaded comment.", author);

        // Retrieve and display the threaded comments for verification (optional)
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments("B2");
        foreach (ThreadedComment comment in threadedComments)
        {
            Console.WriteLine($"Author: {comment.Author.Name}, Text: {comment.Notes}");
        }

        // Save the workbook with the new threaded comment
        workbook.Save("Output.xlsx");
    }
}