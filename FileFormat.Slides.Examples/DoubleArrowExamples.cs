using FileFormat.Slides;
using FileFormat.Slides.Common;
using System;
using System.Collections.Generic;



namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying DoubleArrow segments or shapes in a Presentation
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Slides">FileFormat.Slides</a> library.
    /// </summary>
    public class DoubleArrowExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";

        /// <summary>
        /// Initializes a new instance of the <see cref="DoubleArrowExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public DoubleArrowExamples()
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
        /// This method adds DoubleArrow segment or shape in the silde of a new PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void DrawNewDoubleArrowShapeInNewSlide(string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Create an instance of DoubleArrow
                
                DoubleArrow DoubleArrow = new DoubleArrow();
                // Set height and width
                DoubleArrow.Width = 400.0;
                DoubleArrow.Height = 400.0;
                // Set Y position
                DoubleArrow.Y = 100.0;
                // First slide
                Slide slide = presentation.GetSlides()[1];
                // Add DoubleArrow shapes.
                slide.DrawDoubleArrow(DoubleArrow);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method adds DoubleArrow segment or shape in the silde of a new PowerPoint presentation with animation.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void DrawNewDoubleArrowShapeWithAnimation(string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Create an instance of DoubleArrow

                DoubleArrow DoubleArrow = new DoubleArrow();
                // Set height and width
                DoubleArrow.Width = 400.0;
                DoubleArrow.Height = 400.0;
                // Set Y position
                DoubleArrow.Y = 100.0;
                // First slide
                DoubleArrow.Animation = Common.Enumerations.AnimationType.FlyIn;
                Slide slide = presentation.GetSlides()[1];
                // Add DoubleArrow shapes.
                slide.DrawDoubleArrow(DoubleArrow);
                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method Sets the background color of a DoubleArrow shape
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void SetBackgroundColorOfDoubleArrow(string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get the slides
                Slide slide = presentation.GetSlides()[1];
                // Get 1st DoubleArrow
                DoubleArrow DoubleArrow = slide.DoubleArrows[0];
                // Set background of the DoubleArrow
                DoubleArrow.BackgroundColor = "289876";

                // Save the PPT or PPTX
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }

        /// <summary>
        /// Remove DoubleArrow shape from an existing slide
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void RemoveDoubleArrowShapeExistingSlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
            // Get the slides
            Slide slide = presentation.GetSlides()[1];
            // Get 1st DoubleArrow
            DoubleArrow DoubleArrow = slide.DoubleArrows[0];
            // Remove DoubleArrow
            DoubleArrow.Remove();
            // Save the PPT or PPTX
            presentation.Save();

        }
    }
}
