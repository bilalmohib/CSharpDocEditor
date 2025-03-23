using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using M = DocumentFormat.OpenXml.Math;

namespace CSharpDoc
{
    /// <summary>
    /// Utility class for creating and editing Word documents using OpenXML
    /// </summary>
    public class WordDocumentUtils
    {
        private WordprocessingDocument _document;
        private MainDocumentPart _mainPart;
        private Body _body;
        private StyleDefinitionsPart _stylesPart;

        /// <summary>
        /// Creates a new Word document with the specified path
        /// </summary>
        /// <param name="filePath">The path where the document will be saved</param>
        public void CreateDocument(string filePath)
        {
            _document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            _mainPart = _document.AddMainDocumentPart();
            _mainPart.Document = new Document();
            _body = _mainPart.Document.AppendChild(new Body());
            
            // Create the StyleDefinitionsPart
            _stylesPart = _mainPart.AddNewPart<StyleDefinitionsPart>();
            _stylesPart.Styles = new Styles();
            
            // Add default styles
            AddDefaultStyles();
        }

        /// <summary>
        /// Opens an existing Word document
        /// </summary>
        /// <param name="filePath">The path to the existing document</param>
        public void OpenDocument(string filePath)
        {
            _document = WordprocessingDocument.Open(filePath, true);
            _mainPart = _document.MainDocumentPart;
            _body = _mainPart.Document.Body;
            _stylesPart = _mainPart.StyleDefinitionsPart;
            
            if (_stylesPart == null)
            {
                _stylesPart = _mainPart.AddNewPart<StyleDefinitionsPart>();
                _stylesPart.Styles = new Styles();
                AddDefaultStyles();
            }
        }

        /// <summary>
        /// Saves and closes the document
        /// </summary>
        public void SaveAndClose()
        {
            _document.Close();
        }

        /// <summary>
        /// Adds a paragraph to the document
        /// </summary>
        /// <param name="text">The text content</param>
        /// <param name="styleName">Optional style name</param>
        /// <returns>The created paragraph</returns>
        public Paragraph AddParagraph(string text, string styleName = null)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            Text textElement = new Text(text);
            
            run.AppendChild(textElement);
            para.AppendChild(run);
            
            if (!string.IsNullOrEmpty(styleName))
            {
                para.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleName });
            }
            
            _body.AppendChild(para);
            return para;
        }

        /// <summary>
        /// Creates a table in the document
        /// </summary>
        /// <param name="rows">Number of rows</param>
        /// <param name="cols">Number of columns</param>
        /// <param name="styleName">Optional table style name</param>
        /// <returns>The created table</returns>
        public Table CreateTable(int rows, int cols, string styleName = null)
        {
            Table table = new Table();
            
            TableProperties tableProps = new TableProperties();
            if (!string.IsNullOrEmpty(styleName))
            {
                tableProps.TableStyle = new TableStyle { Val = styleName };
            }
            
            // Set table borders
            TableBorders tableBorders = new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 }
            );
            
            tableProps.AppendChild(tableBorders);
            table.AppendChild(tableProps);
            
            // Create the grid
            TableGrid tg = new TableGrid();
            for (int i = 0; i < cols; i++)
            {
                tg.AppendChild(new GridColumn());
            }
            table.AppendChild(tg);
            
            // Add rows and cells
            for (int i = 0; i < rows; i++)
            {
                TableRow tr = new TableRow();
                for (int j = 0; j < cols; j++)
                {
                    TableCell tc = new TableCell();
                    tc.AppendChild(new Paragraph(new Run(new Text(""))));
                    tr.AppendChild(tc);
                }
                table.AppendChild(tr);
            }
            
            _body.AppendChild(table);
            return table;
        }

        /// <summary>
        /// Sets text in a table cell
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="rowIndex">Row index (0-based)</param>
        /// <param name="colIndex">Column index (0-based)</param>
        /// <param name="text">The text to set</param>
        public void SetCellText(Table table, int rowIndex, int colIndex, string text)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            TableRow row = table.Elements<TableRow>().ElementAt(rowIndex);
            TableCell cell = row.Elements<TableCell>().ElementAt(colIndex);
            
            // Clear existing content
            cell.RemoveAllChildren();
            
            // Add new paragraph with text
            Paragraph para = new Paragraph(new Run(new Text(text)));
            cell.AppendChild(para);
        }

        /// <summary>
        /// Adds a caption to a table or image
        /// </summary>
        /// <param name="captionText">The caption text</param>
        /// <param name="sequenceType">Type of sequence (Table, Figure, etc.)</param>
        /// <returns>The created paragraph containing the caption</returns>
        public Paragraph AddCaption(string captionText, string sequenceType)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            
            // Add label and auto-number field
            run.AppendChild(new Text($"{sequenceType} "));
            para.AppendChild(run);
            
            // Add SEQ field for auto-numbering
            var fieldChar1 = new FieldChar { FieldCharType = FieldCharValues.Begin };
            var fieldCode = new FieldCode { Space = SpaceProcessingModeValues.Preserve };
            fieldCode.Text = $" SEQ {sequenceType} \\* ARABIC ";
            var fieldChar2 = new FieldChar { FieldCharType = FieldCharValues.End };
            
            Run fieldRun1 = new Run(fieldChar1);
            Run fieldRun2 = new Run(fieldCode);
            Run fieldRun3 = new Run(fieldChar2);
            
            para.AppendChild(fieldRun1);
            para.AppendChild(fieldRun2);
            para.AppendChild(fieldRun3);
            
            // Add caption text
            Run captionRun = new Run(new Text($": {captionText}"));
            para.AppendChild(captionRun);
            
            // Set style for caption
            para.ParagraphProperties = new ParagraphProperties(
                new ParagraphStyleId { Val = "Caption" }
            );
            
            _body.AppendChild(para);
            return para;
        }

        /// <summary>
        /// Inserts an image into the document
        /// </summary>
        /// <param name="imagePath">Path to the image file</param>
        /// <param name="width">Desired width in pixels</param>
        /// <param name="height">Desired height in pixels</param>
        /// <returns>The paragraph containing the image</returns>
        public Paragraph InsertImage(string imagePath, int width, int height)
        {
            MainDocumentPart mainPart = _document.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(GetImagePartType(imagePath));
            
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            
            string relationshipId = mainPart.GetIdOfPart(imagePart);
            
            // Convert pixels to EMUs (English Metric Units)
            long emuWidth = width * 9525;
            long emuHeight = height * 9525;
            
            // Create image element
            var element = new Drawing(
                new DW.Inline(
                    new DW.Extent { Cx = emuWidth, Cy = emuHeight },
                    new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new DW.DocProperties { Id = 1U, Name = "Picture 1" },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties { Id = 0U, Name = Path.GetFileName(imagePath) },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip { Embed = relationshipId },
                                    new A.Stretch(new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset { X = 0L, Y = 0L },
                                        new A.Extents { Cx = emuWidth, Cy = emuHeight }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    ) { Preset = A.ShapeTypeValues.Rectangle })
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                    )
                ) { DistanceFromTop = 0U, DistanceFromBottom = 0U, DistanceFromLeft = 0U, DistanceFromRight = 0U }
            );
            
            Paragraph para = new Paragraph(new Run(element));
            _body.AppendChild(para);
            return para;
        }

        /// <summary>
        /// Adds a mathematical formula to the document
        /// </summary>
        /// <param name="equation">The equation text (using Office Math Markup)</param>
        /// <returns>The paragraph containing the equation</returns>
        public Paragraph AddMathFormula(string equation)
        {
            Paragraph para = new Paragraph();
            Run run = new Run();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Run(
                    new M.Text { Text = equation }
                ) 
                { RunProperties = new M.RunProperties() }
            );
            
            run.AppendChild(officeMath);
            para.AppendChild(run);
            _body.AppendChild(para);
            
            return para;
        }

        /// <summary>
        /// Adds a table of contents to the document
        /// </summary>
        /// <param name="title">The title for the TOC</param>
        public void AddTableOfContents(string title)
        {
            // Add TOC title
            Paragraph tocTitle = new Paragraph(
                new ParagraphProperties(
                    new ParagraphStyleId { Val = "Heading1" }
                ),
                new Run(new Text(title))
            );
            _body.AppendChild(tocTitle);
            
            // Add the TOC
            Paragraph tocParagraph = new Paragraph();
            Run tocRun = new Run();
            
            // Add the TOC field
            var fieldChar1 = new FieldChar { FieldCharType = FieldCharValues.Begin };
            
            var fieldCode = new FieldCode { Space = SpaceProcessingModeValues.Preserve };
            fieldCode.Text = " TOC \\o \"1-3\" \\h \\z \\u ";
            
            var fieldChar2 = new FieldChar { FieldCharType = FieldCharValues.Separate };
            var fieldChar3 = new FieldChar { FieldCharType = FieldCharValues.End };
            
            Run fieldRun1 = new Run(fieldChar1);
            Run fieldRun2 = new Run(fieldCode);
            Run fieldRun3 = new Run(fieldChar2);
            Run fieldRun4 = new Run(new Text("Table of Contents placeholder. Right-click and select 'Update Field' to generate."));
            Run fieldRun5 = new Run(fieldChar3);
            
            tocParagraph.AppendChild(fieldRun1);
            tocParagraph.AppendChild(fieldRun2);
            tocParagraph.AppendChild(fieldRun3);
            tocParagraph.AppendChild(fieldRun4);
            tocParagraph.AppendChild(fieldRun5);
            
            _body.AppendChild(tocParagraph);
        }

        /// <summary>
        /// Adds default styles to the document
        /// </summary>
        private void AddDefaultStyles()
        {
            Styles styles = _stylesPart.Styles;
            
            // Add default paragraph style
            Style style = new Style { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            style.AppendChild(new StyleName { Val = "Normal" });
            style.AppendChild(new PrimaryStyle());
            styles.AppendChild(style);
            
            // Add Heading 1 style
            Style heading1Style = new Style { Type = StyleValues.Paragraph, StyleId = "Heading1" };
            heading1Style.AppendChild(new StyleName { Val = "Heading 1" });
            heading1Style.AppendChild(new BasedOn { Val = "Normal" });
            heading1Style.AppendChild(new NextParagraphStyle { Val = "Normal" });
            heading1Style.AppendChild(new PrimaryStyle());
            
            StyleRunProperties heading1RunProperties = new StyleRunProperties(
                new Bold(),
                new FontSize { Val = "28" },
                new Color { Val = "2E74B5" }
            );
            heading1Style.AppendChild(heading1RunProperties);
            
            styles.AppendChild(heading1Style);
            
            // Add Caption style
            Style captionStyle = new Style { Type = StyleValues.Paragraph, StyleId = "Caption" };
            captionStyle.AppendChild(new StyleName { Val = "Caption" });
            captionStyle.AppendChild(new BasedOn { Val = "Normal" });
            captionStyle.AppendChild(new PrimaryStyle());
            
            StyleRunProperties captionRunProperties = new StyleRunProperties(
                new Italic(),
                new FontSize { Val = "20" },
                new Color { Val = "5A5A5A" }
            );
            captionStyle.AppendChild(captionRunProperties);
            
            styles.AppendChild(captionStyle);
            
            // Add Table style
            Style tableStyle = new Style { Type = StyleValues.Table, StyleId = "TableNormal", Default = true };
            tableStyle.AppendChild(new StyleName { Val = "Normal Table" });
            tableStyle.AppendChild(new PrimaryStyle());
            
            styles.AppendChild(tableStyle);
        }

        /// <summary>
        /// Determines the ImagePartType based on the file extension
        /// </summary>
        private ImagePartType GetImagePartType(string imagePath)
        {
            string extension = Path.GetExtension(imagePath).ToLower();
            
            switch (extension)
            {
                case ".png": return ImagePartType.Png;
                case ".jpg":
                case ".jpeg": return ImagePartType.Jpeg;
                case ".gif": return ImagePartType.Gif;
                case ".bmp": return ImagePartType.Bmp;
                case ".tiff": return ImagePartType.Tiff;
                default: throw new NotSupportedException($"Image format {extension} is not supported.");
            }
        }
    }
} 