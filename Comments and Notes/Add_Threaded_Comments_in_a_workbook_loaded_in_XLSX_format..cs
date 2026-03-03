using System;
using Aspose.Cells;

class ThreadedCommentsDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("Input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Add authors for threaded comments
        int authorIndex1 = workbook.Worksheets.ThreadedCommentAuthors.Add("Alice", "alice@example.com", "PROV1");
        ThreadedCommentAuthor author1 = workbook.Worksheets.ThreadedCommentAuthors[authorIndex1];

        int authorIndex2 = workbook.Worksheets.ThreadedCommentAuthors.Add("Bob", "bob@example.com", "PROV2");
        ThreadedCommentAuthor author2 = workbook.Worksheets.ThreadedCommentAuthors[authorIndex2];

        // Add a threaded comment to cell B2 using row/column indices (row 1, column 1)
        worksheet.Comments.AddThreadedComment(1, 1, "Initial comment by Alice", author1);

        // Add a reply to the same cell using the cell name overload
        worksheet.Comments.AddThreadedComment("B2", "Reply by Bob", author2);

        // Retrieve and display all threaded comments for cell B2
        ThreadedCommentCollection threadedComments = worksheet.Comments.GetThreadedComments("B2");
        foreach (ThreadedComment tc in threadedComments)
        {
            Console.WriteLine($"Author: {tc.Author.Name}, Text: {tc.Notes}");
        }

        // Save the workbook with the new threaded comments
        workbook.Save("Output.xlsx");
    }
}