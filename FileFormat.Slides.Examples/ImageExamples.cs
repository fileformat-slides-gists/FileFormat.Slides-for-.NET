using FileFormat.Slides.Common;
using System;
using System.Collections.Generic;
using System.Text;

namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying Presentation images
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Slides">FileFormat.Slides</a> library.
    /// </summary>
    public class ImageExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";
        private const string imagesDirectory = "../../../Presentations/Images";
        /// <summary>
        /// Initializes a new instance of the <see cref="ImageExamples"/> class.
        /// Prepares the directory 'Presentations/Image' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// Prepares the directory 'Presentations/Existing/Images' to store images to be added
        /// to the presentations.
        /// </summary>
        public ImageExamples ()
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
        /// This method adds up an image in a slide
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        /// <param name="imagename">Picture or image name</param>
        public void AddImageInASlide (string documentDirectory = newDocsDirectory, string filename = "test.pptx", string imagename="sample.jpg")
        {
           try
            {
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                
                // Create slide
                Slide slide = new Slide();
                // Add text shapes.
                Image image1 = new Image($"{imagesDirectory}/{imagename}");
                // Set xAxis
                image1.X = 180.0;
                // Set yAxis
                image1.Y = 128.0;
                // Set Width
                image1.Width = 195.0;
                // Set Height
                image1.Height = 55.0;
                // Add image to slide
                slide.AddImage(image1);
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
        /// This method updates the existing image of a slide
        /// </summary>
        /// <param name="documentDirectory">Path of presentation folder</param>
        /// <param name="filename">Presentation name</param>
        /// <param name="xAxis">xAxis value in double</param>
        /// <param name="yAxis">xAxis value in double</param>
        public void UpdateImageInExistingSlide (string documentDirectory = existingDocsDirectory, string filename = "test.pptx", double xAxis=0.0, double yAxis=0.0)
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get Slides
                var slides = presentation.GetSlides();
                var slide = slides[1];
                // Get Images from slide
                List<Image> images = slide.Images;
                // Choose desired image
                var image = slide.Images[0];
                // Set xAxis
                image.X = xAxis;
                // Set yAxis
                image.Y = yAxis;
                // Update image
                image.Update();
                // Save presentation
                presentation.Save();
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }

        }
        /// <summary>
        /// This method removes the image from presentation
        /// </summary>
        /// <param name="documentDirectory">Path of presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void RemoveImageInExistingSlide (string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get Slides
                var slides = presentation.GetSlides();
                var slide = slides[1];
                // Get Images from slide
                List<Image> images = slide.Images;
                // Choose desired image
                var image = slide.Images[0];
                // Remove image
                image.Remove();               
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
