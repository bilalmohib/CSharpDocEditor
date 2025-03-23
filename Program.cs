using System;
using System.IO;

namespace CSharpDoc
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("OpenXML Word Document Utility Demo");
            Console.WriteLine("----------------------------------");

            string outputPath = Path.Combine(Environment.CurrentDirectory, "WordDocumentDemo.docx");
            
            try
            {
                // Create a new document
                WordDocumentUtils docUtils = new WordDocumentUtils();
                docUtils.CreateDocument(outputPath);
                
                // Add a title
                docUtils.AddParagraph("Word Document Generation Demo", "Heading1");
                
                // Add a table of contents
                docUtils.AddTableOfContents("Table of Contents");
                
                // Add a section heading
                docUtils.AddParagraph("1. Tables Example", "Heading1");
                
                // Create a table
                var table = docUtils.CreateTable(3, 3);
                
                // Set table content
                docUtils.SetCellText(table, 0, 0, "Header 1");
                docUtils.SetCellText(table, 0, 1, "Header 2");
                docUtils.SetCellText(table, 0, 2, "Header 3");
                docUtils.SetCellText(table, 1, 0, "Row 1, Cell 1");
                docUtils.SetCellText(table, 1, 1, "Row 1, Cell 2");
                docUtils.SetCellText(table, 1, 2, "Row 1, Cell 3");
                docUtils.SetCellText(table, 2, 0, "Row 2, Cell 1");
                docUtils.SetCellText(table, 2, 1, "Row 2, Cell 2");
                docUtils.SetCellText(table, 2, 2, "Row 2, Cell 3");
                
                // Add a caption to the table
                docUtils.AddCaption("Sample Data Table", "Table");
                
                // Add some paragraphs
                docUtils.AddParagraph("This is a paragraph demonstrating the text content generation.");
                
                // Add a section heading
                docUtils.AddParagraph("2. Math Formula Example", "Heading1");
                
                // Add a math formula - a simple quadratic formula
                // Note: This is a simplified example. Real math formulas require more complex OMML
                docUtils.AddMathFormula("x = (-b ± √(b² - 4ac)) / (2a)");
                
                // Add a caption for the formula
                docUtils.AddCaption("Quadratic Formula", "Equation");
                
                // Add a section heading
                docUtils.AddParagraph("3. Images Example", "Heading1");
                
                // Add an image - this would require an actual image file to exist
                // For this example, we'll comment it out to avoid errors
                // Uncomment and provide a valid image path to test
                /*
                string imagePath = "sample_image.jpg";
                if (File.Exists(imagePath))
                {
                    docUtils.InsertImage(imagePath, 400, 300);
                    docUtils.AddCaption("Sample Image", "Figure");
                }
                */
                
                // Add a paragraph explaining this is a demo
                docUtils.AddParagraph("The image section is commented out in the code as it requires an actual image file.");
                
                // Save and close the document
                docUtils.SaveAndClose();
                
                Console.WriteLine($"Document created successfully: {outputPath}");
                Console.WriteLine("Note: To see the Table of Contents populated, open the document in Word");
                Console.WriteLine("and right-click on the TOC area, then select 'Update Field'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating document: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
    }
} 