using FileFormat.Slides.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Text segments or shapes in a Presentation
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Slides">FileFormat.Slides</a> library.
    /// </summary>
    public class TextExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";

        /// <summary>
        /// Initializes a new instance of the <see cref="TextExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public TextExamples ()
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
        /// This method adds text segment or shape in the silde of a new PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void CreateNewTextShapeInNewSlide (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Create an instance of TextShape
                TextShape shape = new TextShape();
                // Set text of the TextShape
                shape.Text = "FileFormat.Slides Text Shape Example";             
                // Create Slide
                Slide slide = new Slide();
                // Add text shapes.
                slide.AddTextShapes(shape);               
                // Add slide to the presentation
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
        /// This method Sets the background color of a text shape
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void SetBackgroundColorOfTextShape (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Create an instance of TextShape
                TextShape shape = new TextShape();
                // Set text of the TextShape
                shape.Text = "FileFormat.Slides Text Shape Example";
                // Set background color
                shape.BackgroundColor = "5f7200";
                // Create slide
                Slide slide = new Slide();
                // Add text shapes.
                slide.AddTextShapes(shape);
                // Adding slides
                presentation.AppendSlide(slide);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method sets the font family of a text shape
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void SetFontFamilyofTextShape (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Create an instance of TextShape
                TextShape shape = new TextShape();
                // Set text of the TextShape
                shape.Text = "FileFormat.Slides Text Shape Example";
                // Set font family
                shape.FontFamily = "Baguet Script";
                // First slide
                Slide slide = new Slide();
                // Add text shapes.
                slide.AddTextShapes(shape);
                // Adding slides
                presentation.AppendSlide(slide);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method sets the text color of text shape
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void SetTextColorOfTextShape (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Create an instance of TextShape
                TextShape shape = new TextShape();
                // Set text of the TextShape
                shape.Text = "FileFormat.Slides Text Shape Example";
                // Set text color
                shape.TextColor = "980078";
                // First slide
                Slide slide = new Slide();
                // Add text shapes.
                slide.AddTextShapes(shape);
                // Adding slides
                presentation.AppendSlide(slide);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method sets font size of the text
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void SetFontSizeOfTextShape (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                // Create an instance of new presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                // Create an instance of TextShape
                TextShape shape = new TextShape();
                // Set text of the TextShape
                shape.Text = "FileFormat.Slides Text Shape Example";
                // Set font size
                shape.FontSize = 80;
                // First slide
                Slide slide = new Slide();
                // Add text shapes.
                slide.AddTextShapes(shape);
                // Adding slides
                presentation.AppendSlide(slide);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// Add text shape to an existing slide
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void AddNewTextShapeExistingSlide (string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                TextShape shape = new TextShape();
                shape.Text = "Adding new text shape in an existing slide";
                shape.Y = 100.0;
                // First slide
                Slide slide = presentation.GetSlides()[1];
                // Add text shapes.
                slide.AddTextShapes(shape);
                // Adding slides
                presentation.AppendSlide(slide);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
    }
}
