using DocumentFormat.OpenXml.Presentation;
using FileFormat.Slides.Common;
using FileFormat.Slides.Common.Enumerations;
using System;
using System.Collections.Generic;
using System.Text;

namespace FileFormat.Slides.Examples
{
    /// <summary>
    /// Provides examples for creating tables in PowerPoint presentations.
    /// </summary>
    public class TableExamples
    {
        private const string newDocsDirectory = "../../../Presentations/New";
        private const string existingDocsDirectory = "../../../Presentations/Existing";
        /// <summary>
        /// Initializes a new instance of the <see cref="TableExamples"/> class.
        /// Prepares the directory 'Presentations/New' for storing or loading PowerPoint(PPT or PPTX) presentations
        /// at the root of the project.
        /// If the directory doesn't exist, it is created. If it already exists,
        /// existing files are deleted, and the directory is cleaned up.
        /// </summary>
        public TableExamples()
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
        /// Creates a simple table in a slide of a PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>

        public void CreateSimpleTableInASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get desired slide
                Slide slide = presentation.GetSlides()[0];
                // Create a new table
                Table table = new Table();
                // Define table columns
                TableColumn col1 = new TableColumn();
                col1.Name = "ID";
                table.Columns.Add(col1);
                TableColumn col2 = new TableColumn();
                col2.Name = "Name";
                table.Columns.Add(col2);
                TableColumn col3 = new TableColumn();
                col3.Name = "City";
                table.Columns.Add(col3);

                //1st row        
                TableRow row1 = new TableRow(table);
                // Create cells of first row
                TableCell cell11 = new TableCell(row1);
                cell11.Text = "907";
                cell11.ID = col1.Name;
                row1.AddCell(cell11);
                TableCell cell12 = new TableCell(row1);
                cell12.Text = "John";
                cell12.ID = col2.Name;
                row1.AddCell(cell12);
                TableCell cell13 = new TableCell(row1);
                cell13.Text = "Chicago";
                cell13.ID = col3.Name;
                // Add cells to row
                row1.AddCell(cell13);
                // Add row to table
                table.AddRow(row1);

                //2nd Row
                TableRow row2 = new TableRow(table);
                TableCell cell21 = new TableCell(row2);
                cell21.Text = "908";
                cell21.ID = col1.Name;
                row2.AddCell(cell21);
                TableCell cell22 = new TableCell(row2);
                cell22.Text = "Chris";
                cell22.ID = col2.Name;
                row2.AddCell(cell22);
                TableCell cell23 = new TableCell(row2);
                cell23.Text = "New York";
                cell23.ID = col3.Name;
                row2.AddCell(cell23);
                table.AddRow(row2);
                //Set Table dimensions
                table.Width = 500.0;
                table.Height = 200.0;
                table.X = 300.0;
                table.Y = 500.0;
                slide.AddTable(table);

                presentation.Save();
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        
        /// <summary>
        /// Creates a table with stylings applied in a slide of a PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>
        public void CreateTableWithTableStylingsInASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get desired slide
                Slide slide = presentation.GetSlides()[1];

                // Assign values to the properties of Stylings
                Stylings stylings = new Stylings();
                stylings.FontSize = 14;
                stylings.Alignment = FileFormat.Slides.Common.Enumerations.TextAlignment.Left;
                stylings.FontFamily = "Baguet Script";
                stylings.TextColor = Colors.Red;

                // Create a new table
                Table table = new Table();

                table.TableStylings = stylings;
                // Define table columns
                TableColumn col1 = new TableColumn();
                col1.Name = "ID";
                table.Columns.Add(col1);
                TableColumn col2 = new TableColumn();
                col2.Name = "Name";
                table.Columns.Add(col2);
                TableColumn col3 = new TableColumn();
                col3.Name = "City";
                table.Columns.Add(col3);

                //1st row        
                TableRow row1 = new TableRow(table);
                // Create cells of first row
                TableCell cell11 = new TableCell(row1);
                cell11.Text = "907";
                cell11.ID = col1.Name;
                row1.AddCell(cell11);
                TableCell cell12 = new TableCell(row1);
                cell12.Text = "John";
                cell12.ID = col2.Name;
                row1.AddCell(cell12);
                TableCell cell13 = new TableCell(row1);
                cell13.Text = "Chicago";
                cell13.ID = col3.Name;
                // Add cells to row
                row1.AddCell(cell13);
                // Add row to table
                table.AddRow(row1);

                //2nd Row
                TableRow row2 = new TableRow(table);
                TableCell cell21 = new TableCell(row2);
                cell21.Text = "908";
                cell21.ID = col1.Name;
                row2.AddCell(cell21);
                TableCell cell22 = new TableCell(row2);
                cell22.Text = "Chris";
                cell22.ID = col2.Name;
                row2.AddCell(cell22);
                TableCell cell23 = new TableCell(row2);
                cell23.Text = "New York";
                cell23.ID = col3.Name;
                row2.AddCell(cell23);
                table.AddRow(row2);
                //Set Table dimensions
                table.Width = 500.0;
                table.Height = 200.0;
                table.X = 300.0;
                table.Y = 500.0;
                slide.AddTable(table);

                presentation.Save();
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
       
        /// <summary>
        /// Creates a table with stylings applied to rows in a slide of a PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>
        public void CreateTableWithRowStylingsInASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get desired slide
                Slide slide = presentation.GetSlides()[1];                

                // Create a new table
                Table table = new Table();
               
                // Define table columns
                TableColumn col1 = new TableColumn();
                col1.Name = "ID";
                table.Columns.Add(col1);
                TableColumn col2 = new TableColumn();
                col2.Name = "Name";
                table.Columns.Add(col2);
                TableColumn col3 = new TableColumn();
                col3.Name = "City";
                table.Columns.Add(col3);

                // Assign values to the properties of Stylings
                Stylings stylings = new Stylings();
                stylings.FontSize = 14;
                stylings.Alignment = FileFormat.Slides.Common.Enumerations.TextAlignment.Left;
                stylings.FontFamily = "Baguet Script";
                stylings.TextColor = Colors.Red;

                //1st row        
                TableRow row1 = new TableRow(table);
                // Apply row stylings
                row1.RowStylings= stylings;
                // Create cells of first row
                TableCell cell11 = new TableCell(row1);
                cell11.Text = "907";
                cell11.ID = col1.Name;
                row1.AddCell(cell11);
                TableCell cell12 = new TableCell(row1);
                cell12.Text = "John";
                cell12.ID = col2.Name;
                row1.AddCell(cell12);
                TableCell cell13 = new TableCell(row1);
                cell13.Text = "Chicago";
                cell13.ID = col3.Name;
                // Add cells to row
                row1.AddCell(cell13);
                // Add row to table
                table.AddRow(row1);

                //2nd Row
                TableRow row2 = new TableRow(table);
                TableCell cell21 = new TableCell(row2);
                cell21.Text = "908";
                cell21.ID = col1.Name;
                row2.AddCell(cell21);
                TableCell cell22 = new TableCell(row2);
                cell22.Text = "Chris";
                cell22.ID = col2.Name;
                row2.AddCell(cell22);
                TableCell cell23 = new TableCell(row2);
                cell23.Text = "New York";
                cell23.ID = col3.Name;
                row2.AddCell(cell23);
                table.AddRow(row2);
                //Set Table dimensions
                table.Width = 500.0;
                table.Height = 200.0;
                table.X = 300.0;
                table.Y = 500.0;
                slide.AddTable(table);

                presentation.Save();
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }
        
        /// <summary>
        /// Creates a table with stylings applied to cells in a slide of a PowerPoint presentation.
        /// </summary>
        /// <param name="documentDirectory">The directory where the PowerPoint presentation is located. Default is 'Presentations/Existing'.</param>
        /// <param name="filename">The name of the PowerPoint file. Default is 'test.pptx'.</param>
        public void CreateTableWithCellStylingsInASlide(string documentDirectory = existingDocsDirectory, string filename = "test.pptx")
        {
            try
            {
                // Create instance of presentation
                Presentation presentation = Presentation.Open($"{documentDirectory}/{filename}");
                // Get desired slide
                Slide slide = presentation.GetSlides()[1];

                // Create a new table
                Table table = new Table();

                // Define table columns
                TableColumn col1 = new TableColumn();
                col1.Name = "ID";
                table.Columns.Add(col1);
                TableColumn col2 = new TableColumn();
                col2.Name = "Name";
                table.Columns.Add(col2);
                TableColumn col3 = new TableColumn();
                col3.Name = "City";
                table.Columns.Add(col3);               

                //1st row        
                TableRow row1 = new TableRow(table);                
                // Create cells of first row

                // Assign values to the properties of Stylings
                Stylings stylings = new Stylings();
                stylings.FontSize = 14;
                stylings.Alignment = FileFormat.Slides.Common.Enumerations.TextAlignment.Left;
                stylings.FontFamily = "Baguet Script";
                stylings.TextColor = Colors.Red;

                TableCell cell11 = new TableCell(row1);
                cell11.CellStylings = stylings;
                cell11.Text = "907";
                cell11.ID = col1.Name;
                row1.AddCell(cell11);
                TableCell cell12 = new TableCell(row1);
                cell12.Text = "John";
                cell12.ID = col2.Name;
                row1.AddCell(cell12);
                TableCell cell13 = new TableCell(row1);
                cell13.Text = "Chicago";
                cell13.ID = col3.Name;
                // Add cells to row
                row1.AddCell(cell13);
                // Add row to table
                table.AddRow(row1);

                //2nd Row
                TableRow row2 = new TableRow(table);
                TableCell cell21 = new TableCell(row2);
                cell21.Text = "908";
                cell21.ID = col1.Name;
                row2.AddCell(cell21);
                TableCell cell22 = new TableCell(row2);
                cell22.Text = "Chris";
                cell22.ID = col2.Name;
                row2.AddCell(cell22);
                TableCell cell23 = new TableCell(row2);
                cell23.Text = "New York";
                cell23.ID = col3.Name;
                row2.AddCell(cell23);
                table.AddRow(row2);
                //Set Table dimensions
                table.Width = 500.0;
                table.Height = 200.0;
                table.X = 300.0;
                table.Y = 500.0;
                slide.AddTable(table);

                presentation.Save();
            }
            catch (System.Exception ex)
            {
                throw new FileFormat.Slides.Common.FileFormatException("An error occurred.", ex);
            }
        }


    }

}
