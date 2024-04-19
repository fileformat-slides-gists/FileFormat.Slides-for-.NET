using DocumentFormat.OpenXml.Presentation;
using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using System;
using System.Collections.Generic;
using System.Text;


namespace FileFormat.Slides.Examples
{
    public class CommentExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";
        /// <summary>
        /// Initializes a new instance of the <see cref="CommentExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public CommentExamples()
        {
            if (!System.IO.Directory.Exists(newDocsDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(newDocsDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(newDocsDirectory)}' " +
                    $"created successfully.");
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(newDocsDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(newDocsDirectory)}' " +
                    $"cleaned up.");
            }
        }
        /// <summary>
        /// Creates comment in a slide of a PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void CreateCommentInASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get 1st slide
                Slide slide = presentation.GetSlides()[0];
                // Create comment
                Comment comment1 = new Comment();
                // Set comment author
                comment1.AuthorId = 1;
                // Add comment content
                comment1.Text = "2nd Programmatic comment in an existing presentation";
                // Set comment time
                comment1.InsertedAt = DateTime.Now;
                // Add comment to slide
                slide.AddComment(comment1);
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Remove a comment from a slide.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void RemoveACommentFromASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get 1st slide
                Slide slide = presentation.GetSlides()[0];
                // Get comments of a slide
                var comments = slide.GetComments();
                // Remove 1st comment
                comments[0].Remove();
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Add comment with new comment author.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void AddCommentWithExistingCommentAuthor(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get 1st slide
                Slide slide = presentation.GetSlides()[0];
                // Create comment
                Comment comment1 = new Comment();
                // Get existing saved comment author 
                CommentAuthor author = presentation.GetCommentAuthors()[0];
                // Set authorId
                comment1.AuthorId = author.Id;
                // Add comment content
                comment1.Text = "2nd Programmatic comment in an existing presentation";
                // Set comment time
                comment1.InsertedAt = DateTime.Now;
                // Add comment to slide
                slide.AddComment(comment1);
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
    }
}
