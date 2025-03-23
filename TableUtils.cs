using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;

namespace CSharpDoc
{
    /// <summary>
    /// Defines table border styles
    /// </summary>
    public enum TableBorderStyle
    {
        None,
        Single,
        Thick,
        Double,
        Dotted,
        Dashed,
        DotDash,
        DotDotDash
    }

    /// <summary>
    /// Defines table alignment options
    /// </summary>
    public enum TableAlignment
    {
        Left,
        Center,
        Right
    }

    /// <summary>
    /// Utility class for advanced table operations in Word documents
    /// </summary>
    public class TableUtils
    {
        private WordprocessingDocument _document;
        
        /// <summary>
        /// Initializes a new instance of the TableUtils class
        /// </summary>
        /// <param name="document">The WordprocessingDocument to work with</param>
        public TableUtils(WordprocessingDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// Creates a table with the specified number of rows and columns
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
            
            // Set default table borders
            TableBorders tableBorders = new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 12 }
            );
            
            tableProps.AppendChild(tableBorders);
            
            // Set default table width
            TableWidth tableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            tableProps.AppendChild(tableWidth);
            
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
            
            _document.MainDocumentPart.Document.Body.AppendChild(table);
            return table;
        }
        
        /// <summary>
        /// Sets text in a table cell
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="rowIndex">Row index (0-based)</param>
        /// <param name="colIndex">Column index (0-based)</param>
        /// <param name="text">The text to set</param>
        /// <param name="styleName">Optional paragraph style</param>
        public void SetCellText(Table table, int rowIndex, int colIndex, string text, string styleName = null)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            TableRow row = table.Elements<TableRow>().ElementAt(rowIndex);
            TableCell cell = row.Elements<TableCell>().ElementAt(colIndex);
            
            // Clear existing content
            cell.RemoveAllChildren();
            
            // Add new paragraph with text
            Paragraph para = new Paragraph();
            Run run = new Run(new Text(text));
            para.AppendChild(run);
            
            // Apply style if specified
            if (!string.IsNullOrEmpty(styleName))
            {
                para.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = styleName });
            }
            
            cell.AppendChild(para);
        }
        
        /// <summary>
        /// Sets cell properties such as shading, borders, etc.
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="rowIndex">Row index (0-based)</param>
        /// <param name="colIndex">Column index (0-based)</param>
        /// <param name="backgroundColor">Background color in hex format (e.g., "FF0000" for red)</param>
        /// <param name="verticalAlignment">Vertical text alignment</param>
        public void SetCellProperties(
            Table table, 
            int rowIndex, 
            int colIndex, 
            string backgroundColor = null,
            VerticalAlignmentValues? verticalAlignment = null)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            TableRow row = table.Elements<TableRow>().ElementAt(rowIndex);
            TableCell cell = row.Elements<TableCell>().ElementAt(colIndex);
            
            TableCellProperties cellProps = cell.TableCellProperties;
            if (cellProps == null)
            {
                cellProps = new TableCellProperties();
                cell.AppendChild(cellProps);
            }
            
            // Set background color if specified
            if (!string.IsNullOrEmpty(backgroundColor))
            {
                Shading shading = new Shading()
                {
                    Color = "auto",
                    Fill = backgroundColor,
                    Val = ShadingPatternValues.Clear
                };
                
                cellProps.Shading = shading;
            }
            
            // Set vertical alignment if specified
            if (verticalAlignment.HasValue)
            {
                TableCellVerticalAlignment vAlign = new TableCellVerticalAlignment() { Val = verticalAlignment.Value };
                cellProps.TableCellVerticalAlignment = vAlign;
            }
        }
        
        /// <summary>
        /// Merges cells in a table
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="startRow">Start row index (0-based)</param>
        /// <param name="startCol">Start column index (0-based)</param>
        /// <param name="endRow">End row index (0-based)</param>
        /// <param name="endCol">End column index (0-based)</param>
        public void MergeCells(Table table, int startRow, int startCol, int endRow, int endCol)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            // Validate merge range
            if (startRow > endRow || startCol > endCol)
            {
                throw new ArgumentException("End position must be after start position");
            }
            
            // Handle horizontal merge (within a single row)
            if (startRow == endRow)
            {
                TableRow row = table.Elements<TableRow>().ElementAt(startRow);
                
                // Get all cells in the merge range
                IEnumerable<TableCell> cells = row.Elements<TableCell>().Skip(startCol).Take(endCol - startCol + 1);
                
                // First cell gets grid span
                TableCell firstCell = cells.First();
                TableCellProperties cellProps = firstCell.TableCellProperties;
                if (cellProps == null)
                {
                    cellProps = new TableCellProperties();
                    firstCell.AppendChild(cellProps);
                }
                
                GridSpan gridSpan = new GridSpan() { Val = endCol - startCol + 1 };
                cellProps.GridSpan = gridSpan;
                
                // Remove the other cells
                foreach (TableCell cell in cells.Skip(1))
                {
                    row.RemoveChild(cell);
                }
            }
            // Handle vertical merge (within a single column)
            else if (startCol == endCol)
            {
                // Get all rows in the merge range
                IEnumerable<TableRow> rows = table.Elements<TableRow>().Skip(startRow).Take(endRow - startRow + 1);
                
                // First cell gets vertical merge start
                TableCell firstCell = rows.First().Elements<TableCell>().ElementAt(startCol);
                TableCellProperties firstCellProps = firstCell.TableCellProperties;
                if (firstCellProps == null)
                {
                    firstCellProps = new TableCellProperties();
                    firstCell.AppendChild(firstCellProps);
                }
                
                VerticalMerge verticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };
                firstCellProps.VerticalMerge = verticalMerge;
                
                // Other cells get vertical merge continue
                foreach (TableRow row in rows.Skip(1))
                {
                    TableCell cell = row.Elements<TableCell>().ElementAt(startCol);
                    TableCellProperties cellProps = cell.TableCellProperties;
                    if (cellProps == null)
                    {
                        cellProps = new TableCellProperties();
                        cell.AppendChild(cellProps);
                    }
                    
                    verticalMerge = new VerticalMerge() { Val = MergedCellValues.Continue };
                    cellProps.VerticalMerge = verticalMerge;
                }
            }
            // Handle rectangular area merge (both horizontal and vertical)
            else
            {
                // This is more complex and requires both vertical and horizontal merging
                // First, merge each row horizontally
                for (int i = startRow; i <= endRow; i++)
                {
                    MergeCells(table, i, startCol, i, endCol);
                }
                
                // Then merge the resulting cells vertically
                MergeCells(table, startRow, startCol, endRow, startCol);
            }
        }
        
        /// <summary>
        /// Sets the table borders
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="borderStyle">Border style</param>
        /// <param name="borderSize">Border size</param>
        /// <param name="borderColor">Border color in hex format (e.g., "000000" for black)</param>
        public void SetTableBorders(Table table, TableBorderStyle borderStyle, UInt32Value borderSize, string borderColor)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            BorderValues borderValue = ConvertToBorderValues(borderStyle);
            
            TableProperties tableProps = table.TableProperties;
            if (tableProps == null)
            {
                tableProps = new TableProperties();
                table.AppendChild(tableProps);
            }
            
            TableBorders tableBorders = new TableBorders(
                new TopBorder { Val = new EnumValue<BorderValues>(borderValue), Size = borderSize, Color = borderColor },
                new BottomBorder { Val = new EnumValue<BorderValues>(borderValue), Size = borderSize, Color = borderColor },
                new LeftBorder { Val = new EnumValue<BorderValues>(borderValue), Size = borderSize, Color = borderColor },
                new RightBorder { Val = new EnumValue<BorderValues>(borderValue), Size = borderSize, Color = borderColor },
                new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(borderValue), Size = borderSize, Color = borderColor },
                new InsideVerticalBorder { Val = new EnumValue<BorderValues>(borderValue), Size = borderSize, Color = borderColor }
            );
            
            tableProps.TableBorders = tableBorders;
        }
        
        /// <summary>
        /// Sets the table alignment
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="alignment">Table alignment</param>
        public void SetTableAlignment(Table table, TableAlignment alignment)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            TableProperties tableProps = table.TableProperties;
            if (tableProps == null)
            {
                tableProps = new TableProperties();
                table.AppendChild(tableProps);
            }
            
            TableJustification tableJustification = new TableJustification();
            
            switch (alignment)
            {
                case TableAlignment.Left:
                    tableJustification.Val = TableRowAlignmentValues.Left;
                    break;
                case TableAlignment.Center:
                    tableJustification.Val = TableRowAlignmentValues.Center;
                    break;
                case TableAlignment.Right:
                    tableJustification.Val = TableRowAlignmentValues.Right;
                    break;
            }
            
            tableProps.TableJustification = tableJustification;
        }
        
        /// <summary>
        /// Sets the width of a table column
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="colIndex">Column index (0-based)</param>
        /// <param name="width">Width value</param>
        /// <param name="unit">Width unit (e.g., "pct" for percentage, "dxa" for twips)</param>
        public void SetColumnWidth(Table table, int colIndex, string width, string unit)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            TableWidthUnitValues widthUnit = TableWidthUnitValues.Pct; // Default to percentage
            
            switch (unit.ToLower())
            {
                case "pct":
                case "percent":
                    widthUnit = TableWidthUnitValues.Pct;
                    break;
                case "dxa":
                case "twips":
                    widthUnit = TableWidthUnitValues.Dxa;
                    break;
                case "auto":
                    widthUnit = TableWidthUnitValues.Auto;
                    break;
            }
            
            // Apply width to each cell in the column
            IEnumerable<TableRow> rows = table.Elements<TableRow>();
            foreach (TableRow row in rows)
            {
                if (colIndex < row.Elements<TableCell>().Count())
                {
                    TableCell cell = row.Elements<TableCell>().ElementAt(colIndex);
                    TableCellProperties cellProps = cell.TableCellProperties;
                    if (cellProps == null)
                    {
                        cellProps = new TableCellProperties();
                        cell.AppendChild(cellProps);
                    }
                    
                    cellProps.TableCellWidth = new TableCellWidth { Width = width, Type = widthUnit };
                }
            }
            
            // Also update the grid
            TableGrid grid = table.Elements<TableGrid>().FirstOrDefault();
            if (grid != null && colIndex < grid.Elements<GridColumn>().Count())
            {
                GridColumn column = grid.Elements<GridColumn>().ElementAt(colIndex);
                column.Width = width;
            }
        }
        
        /// <summary>
        /// Applies alternating row shading to a table
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="evenRowColor">Background color for even rows</param>
        /// <param name="oddRowColor">Background color for odd rows</param>
        /// <param name="headerRowColor">Background color for the header row (optional)</param>
        public void ApplyAlternatingRowShading(Table table, string evenRowColor, string oddRowColor, string headerRowColor = null)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            IEnumerable<TableRow> rows = table.Elements<TableRow>();
            
            int rowIndex = 0;
            foreach (TableRow row in rows)
            {
                string backgroundColor;
                
                if (rowIndex == 0 && !string.IsNullOrEmpty(headerRowColor))
                {
                    backgroundColor = headerRowColor;
                }
                else if (rowIndex % 2 == 0)
                {
                    backgroundColor = evenRowColor;
                }
                else
                {
                    backgroundColor = oddRowColor;
                }
                
                // Apply shading to all cells in the row
                foreach (TableCell cell in row.Elements<TableCell>())
                {
                    TableCellProperties cellProps = cell.TableCellProperties;
                    if (cellProps == null)
                    {
                        cellProps = new TableCellProperties();
                        cell.AppendChild(cellProps);
                    }
                    
                    Shading shading = new Shading()
                    {
                        Color = "auto",
                        Fill = backgroundColor,
                        Val = ShadingPatternValues.Clear
                    };
                    
                    cellProps.Shading = shading;
                }
                
                rowIndex++;
            }
        }
        
        /// <summary>
        /// Sets header row properties
        /// </summary>
        /// <param name="table">The table</param>
        /// <param name="headerRowIndex">Header row index (usually 0)</param>
        /// <param name="fontColor">Text color in hex format</param>
        /// <param name="backgroundColor">Background color in hex format</param>
        /// <param name="bold">Whether to make the text bold</param>
        public void SetHeaderRow(Table table, int headerRowIndex, string fontColor, string backgroundColor, bool bold = true)
        {
            if (table == null) throw new ArgumentNullException(nameof(table));
            
            TableRow headerRow = table.Elements<TableRow>().ElementAt(headerRowIndex);
            
            foreach (TableCell cell in headerRow.Elements<TableCell>())
            {
                // Apply cell properties
                TableCellProperties cellProps = cell.TableCellProperties;
                if (cellProps == null)
                {
                    cellProps = new TableCellProperties();
                    cell.AppendChild(cellProps);
                }
                
                // Set shading
                Shading shading = new Shading()
                {
                    Color = "auto",
                    Fill = backgroundColor,
                    Val = ShadingPatternValues.Clear
                };
                
                cellProps.Shading = shading;
                
                // Set vertical alignment to center
                TableCellVerticalAlignment vAlign = new TableCellVerticalAlignment() { Val = VerticalAlignmentValues.Center };
                cellProps.TableCellVerticalAlignment = vAlign;
                
                // Apply text formatting to each paragraph in the cell
                foreach (Paragraph para in cell.Elements<Paragraph>())
                {
                    ParagraphProperties paraProps = para.ParagraphProperties;
                    if (paraProps == null)
                    {
                        paraProps = new ParagraphProperties();
                        para.AppendChild(paraProps);
                    }
                    
                    // Center align the text
                    Justification justification = new Justification() { Val = JustificationValues.Center };
                    paraProps.Justification = justification;
                    
                    // Apply run properties to each run
                    foreach (Run run in para.Elements<Run>())
                    {
                        RunProperties runProps = run.RunProperties;
                        if (runProps == null)
                        {
                            runProps = new RunProperties();
                            run.PrependChild(runProps);
                        }
                        
                        // Set text color
                        Color color = new Color() { Val = fontColor };
                        runProps.Color = color;
                        
                        // Set bold if specified
                        if (bold)
                        {
                            Bold boldProperty = new Bold();
                            runProps.Bold = boldProperty;
                        }
                    }
                }
            }
        }
        
        /// <summary>
        /// Converts TableBorderStyle enum to OpenXML BorderValues
        /// </summary>
        private BorderValues ConvertToBorderValues(TableBorderStyle style)
        {
            switch (style)
            {
                case TableBorderStyle.None:
                    return BorderValues.None;
                case TableBorderStyle.Single:
                    return BorderValues.Single;
                case TableBorderStyle.Thick:
                    return BorderValues.Thick;
                case TableBorderStyle.Double:
                    return BorderValues.Double;
                case TableBorderStyle.Dotted:
                    return BorderValues.Dotted;
                case TableBorderStyle.Dashed:
                    return BorderValues.Dashed;
                case TableBorderStyle.DotDash:
                    return BorderValues.DotDash;
                case TableBorderStyle.DotDotDash:
                    return BorderValues.DotDotDash;
                default:
                    return BorderValues.Single;
            }
        }
    }
} 