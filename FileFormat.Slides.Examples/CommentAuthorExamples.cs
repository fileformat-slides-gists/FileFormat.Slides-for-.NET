using DocumentFormat.OpenXml.Presentation;
using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using System;
using System.Collections.Generic;
using System.Text;

namespace FileFormat.Slides.Examples
{
    public class CommentAuthorExamples
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
        public CommentAuthorExamples()
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
        /// Method to add new comment author.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void AddCommentAuthor(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Create new comment author
                CommentAuthor author = new CommentAuthor();
                author.Name = "hp";
                author.InitialLetter = "h";
                author.ColorIndex = 2;
                author.Id = 1;
                presentation.CreateAuthor(author);               
                // Save presentation
                presentation.Save();
                

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Method to remove comment authors.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void RemoveCommentAuthor(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Get existing comment authors
                List<CommentAuthor> authors = presentation.GetCommentAuthors();
                // Remove comment authors
                foreach (CommentAuthor author in authors)
                {
                    presentation.RemoveCommentAuthor(author);
                }

                presentation.Save();
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }

    }
}
