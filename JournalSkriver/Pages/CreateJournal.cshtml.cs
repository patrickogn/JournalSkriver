using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.IO;
using Xceed.Words.NET;

namespace JournalSkriver.Pages
{
    public class CreateJournalModel : PageModel
    {
        [BindProperty]
        public string Title { get; set; } = string.Empty;

        [BindProperty]
        public string Content { get; set; } = string.Empty;

        public IActionResult OnPost()
        {
            // Validate input
            if (string.IsNullOrWhiteSpace(Title) || string.IsNullOrWhiteSpace(Content))
            {
                ModelState.AddModelError(string.Empty, "Title and content are required.");
                return Page();
            }

            // Create a memory stream for the Word document
            using var stream = new MemoryStream();

            // ? FIX: Specify the stream type explicitly to avoid CS0411
            using (var doc = DocX.Create(stream))
            {
                // Add the journal content to the Word document
                doc.InsertParagraph(Title)
                    .FontSize(20)
                    .Bold()
                    .SpacingAfter(20);

                doc.InsertParagraph(Content)
                    .FontSize(12);

                doc.Save();
            }

            // Reset stream position for reading
            stream.Position = 0;

            // Prepare a valid filename
            var fileName = $"{Title.Replace(" ", "_")}.docx";

            // Return the Word file as a download
            return File(
                stream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                fileName
            );
        }
    }
}