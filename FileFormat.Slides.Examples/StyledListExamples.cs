using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using System;
using System.Collections.Generic;
using System.Text;

namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides C# code examples for creating, reading, and modifying slides in a Presentation
    /// using the <a href="https://www.nuget.org/packages/FileFormat.Slides">FileFormat.Slides</a> library.
    /// </summary>
    public class StyledListExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";
        /// <summary>
        /// Initializes a new instance of the <see cref="StyledListExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public StyledListExamples ()
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
        /// This method creates bulleted list in a slide
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void CreateBulletedListInASlide (string documentDirectory = newDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                //Create instances of text shapes and set their properties.
                TextShape shape = new TextShape();
                shape.Y = 20.0;
                shape.FontSize = 80;
                shape.FontFamily = "Baguet Script";
                shape.Text = "Bulleted List Example";
                TextShape shape2 = new TextShape();
                shape2.FontSize = 100;
                shape2.Y = 180.0;
                // Create instance of Bulleted List
                StyledList list = new StyledList(ListType.Bulleted);
                // Add items to list
                list.AddListItem("USA");
                list.AddListItem("Canada");
                list.AddListItem("Brazil");
                list.AddListItem("Mexico");
                // Assign list to text shape
                shape2.TextList = list;
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
        /// This method creates numbered list in a slide
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void CreateNumberedListInASlide (string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {

                // Create instance of presentation
                Presentation presentation = Presentation.Create($"{documentDirectory}/{filename}");
                //Create instances of text shapes and set their properties.
                TextShape shape = new TextShape();
                shape.Y = 20.0;
                shape.FontSize = 80;
                shape.FontFamily = "Amasis MT Pro Black";
                shape.Text = "Numbered List Example";
                TextShape shape2 = new TextShape();
                shape2.FontSize = 100;
                shape2.Y = 180.0;
                // Create instance of Bulleted List
                StyledList list = new StyledList(ListType.Bulleted);
                // Add items to list
                list.AddListItem("Tennis");
                list.AddListItem("Cricket");
                list.AddListItem("Hockey");
                list.AddListItem("Football");
                list.AddListItem("Snooker");
                // Assign list to text shape
                shape2.TextList = list;
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
        /// This method adds the list items in an existing numbered or bulleted list.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void AddListItemsInAnExistingList(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {                
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get desired slide
                Slide slide = presentation.GetSlides()[3];
                // Create instance of Bulleted List
                TextShape shape = slide.TextShapes[0];
                StyledList list = shape.TextList;
                // Add items to list
                list.AddListItem("Tennis");
                list.AddListItem("Cricket");
                list.AddListItem("Hockey");
                list.AddListItem("Football");
                list.AddListItem("Snooker");
                // Update the list
                list.Update();
                // Save presentation
                presentation.Save();

            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        /// <summary>
        /// This method removes the list items in an existing numbered or bulleted list.
        /// </summary>
        /// <param name="documentDirectory">Path of the presentation folder</param>
        /// <param name="filename">Presentation name</param>
        public void RemoveListItemsInAnExistingList (string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get desired slide
                Slide slide = presentation.GetSlides()[3];
                // Create instance of Bulleted List
                TextShape shape = slide.TextShapes[0];
                StyledList list = shape.TextList;
                // Remove a range of list items
                list.ListItems.RemoveRange(0,4);
                // Or you can remove all items like 
                list.ListItems.RemoveAt(1);
                // Update the list
                list.Update();
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
