using FileFormat.Slides.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying slides in a Presentation
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Slides">FileFormat.Slides</a> library.
    /// </summary>
    public class SlideExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";
        /// <summary>
        /// Initializes a new instance of the <see cref="SlideExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public SlideExamples ()
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
        /// This method create slides in a new presentation
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void CreateNewSlideInNewPresentation (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                //Create instances of text shapes and set their texts.
                TextShape shape = new TextShape();
                shape.Text = "Title: Here is my first title From FF";
                TextShape shape2 = new TextShape();
                shape2.Text = "Body : Here is my first title From FF";    
                // Set yAxis of 2nd text shape
                shape2.Y = 25.9;
                // Create slide
                Slide slide = new Slide();
                // Set background color of slide because default background color is black.
                slide.BackgroundColor = Colors.Silver;
                // Add text shapes.
                slide.AddTextShapes(shape);
                slide.AddTextShapes(shape2);               
                // Adding slides
                presentation.AppendSlide(slide); 
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method creates new slide in an existing presentation
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void CreateNewSlideInExistingPresentation (string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                //Create instances of text shapes and set their texts.
                TextShape shape = new TextShape();
                shape.Text = "Title: Here is my first title From FF";
                TextShape shape2 = new TextShape();
                shape2.Text = "Body : Here is my first title From FF";
                shape2.Y = 25.9;
                // Create new slide
                Slide slide = new Slide();
                // Set background color of slide because default background color is black.
                slide.BackgroundColor = Colors.Silver;
                // Add text shapes.
                slide.AddTextShapes(shape);
                slide.AddTextShapes(shape2);
                // Adding slides
                presentation.AppendSlide(slide);
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method removes slide in an existing presentation
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void RemoveSlideInAnExistingPresentation (string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Remove slide at first index
                presentation.RemoveSlide(0);
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
