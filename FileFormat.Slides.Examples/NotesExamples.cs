using DocumentFormat.OpenXml.Presentation;
using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using System;
using System.Collections.Generic;
using System.Text;


namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and removing comments in a slide of a PowerPoint presentations
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Slides">FileFormat.Slides</a> library.
    /// </summary>
    public class NotesExamples
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
        public NotesExamples()
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
        /// Creates notes in a slide of a PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void CreateNotesInASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get 1st slide
                Slide slide = presentation.GetSlides()[0];
                // Create Note
                slide.AddNote("Here is my first note");                
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Remove notes from a slide.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void RemoveNotesFromASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get 1st slide
                Slide slide = presentation.GetSlides()[0];
                //Remove notes
                slide.RemoveNote();
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Export all notes of presentation to a text file.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void ExportNotesToTextFile(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Export all notes of presentation to a text file
                presentation.SaveAllNotesToTextFile("notes.txt");
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
