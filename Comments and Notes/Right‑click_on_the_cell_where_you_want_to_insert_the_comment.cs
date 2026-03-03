using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook (simulates opening the file before right‑clicking a cell)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Specify the cell where the comment should be inserted (e.g., B2)
        string targetCell = "B2";

        // Add a comment to the specified cell using the CommentCollection.Add(string) method
        int commentIndex = worksheet.Comments.Add(targetCell);

        // Retrieve the newly created Comment object
        Comment comment = worksheet.Comments[commentIndex];

        // Set comment properties (text, author, visibility, etc.)
        comment.Note = "This comment was added programmatically, equivalent to right‑click → Insert Comment.";
        comment.Author = "Aspose.Cells";
        comment.IsVisible = true; // Makes the comment visible without hovering

        // Save the workbook with the new comment
        workbook.Save("output.xlsx");
    }
}