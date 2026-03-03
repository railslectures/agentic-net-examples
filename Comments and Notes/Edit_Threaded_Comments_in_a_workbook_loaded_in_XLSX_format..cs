using System;
using Aspose.Cells;

namespace AsposeCellsThreadedCommentEdit
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (lifecycle rule: load)
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the comment collection of the worksheet
            CommentCollection comments = worksheet.Comments;

            // Ensure there is at least one threaded comment author
            // If no authors exist, add a default one
            int authorIdx = workbook.Worksheets.ThreadedCommentAuthors.Add("Default Author", "defaultUser", "defaultProvider");
            ThreadedCommentAuthor author = workbook.Worksheets.ThreadedCommentAuthors[authorIdx];

            // Example: Add a new threaded comment to cell B2 (row 1, column 1)
            comments.AddThreadedComment(1, 1, "Initial threaded comment.", author);

            // Retrieve all threaded comments for cell B2
            var threadedComments = comments.GetThreadedComments(1, 1);

            // Update the text (Notes) of each threaded comment
            foreach (ThreadedComment tc in threadedComments)
            {
                // Append a suffix to indicate the comment was edited
                tc.Notes = tc.Notes + " (Edited on " + DateTime.Now.ToString("yyyy-MM-dd") + ")";
            }

            // Optionally, add a reply to the existing threaded comment
            comments.AddThreadedComment(1, 1, "This is a reply to the edited comment.", author);

            // Save the modified workbook (lifecycle rule: save)
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Threaded comments edited and workbook saved to: " + outputPath);
        }
    }
}