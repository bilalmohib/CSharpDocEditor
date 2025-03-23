using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;

namespace CSharpDoc
{
    /// <summary>
    /// Utility class for working with mathematical formulas in Word documents
    /// </summary>
    public class MathUtils
    {
        private WordprocessingDocument _document;
        
        /// <summary>
        /// Initializes a new instance of the MathUtils class
        /// </summary>
        /// <param name="document">The WordprocessingDocument to work with</param>
        public MathUtils(WordprocessingDocument document)
        {
            _document = document ?? throw new ArgumentNullException(nameof(document));
        }
        
        /// <summary>
        /// Adds a simple equation to the document
        /// </summary>
        /// <param name="equation">The equation text</param>
        /// <returns>The paragraph containing the equation</returns>
        public Paragraph AddSimpleEquation(string equation)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Run(
                    new M.RunProperties(new M.Text { Text = equation })
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds a fraction to the document
        /// </summary>
        /// <param name="numerator">Numerator text</param>
        /// <param name="denominator">Denominator text</param>
        /// <returns>The paragraph containing the fraction</returns>
        public Paragraph AddFraction(string numerator, string denominator)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Fraction(
                    new M.Numerator(
                        new M.Run(
                            new M.Text { Text = numerator }
                        )
                    ),
                    new M.Denominator(
                        new M.Run(
                            new M.Text { Text = denominator }
                        )
                    )
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds a radical (square root) to the document
        /// </summary>
        /// <param name="baseText">The text inside the radical</param>
        /// <returns>The paragraph containing the radical</returns>
        public Paragraph AddRadical(string baseText)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Radical(
                    new M.RadicalProperties(),
                    new M.Base(
                        new M.Run(
                            new M.Text { Text = baseText }
                        )
                    )
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds a radical with a degree (nth root) to the document
        /// </summary>
        /// <param name="baseText">The text inside the radical</param>
        /// <param name="degree">The degree of the radical (e.g., "3" for cube root)</param>
        /// <returns>The paragraph containing the radical</returns>
        public Paragraph AddRadicalWithDegree(string baseText, string degree)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Radical(
                    new M.RadicalProperties(),
                    new M.Degree(
                        new M.Run(
                            new M.Text { Text = degree }
                        )
                    ),
                    new M.Base(
                        new M.Run(
                            new M.Text { Text = baseText }
                        )
                    )
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds a superscript (like x²) to the document
        /// </summary>
        /// <param name="baseText">The base text</param>
        /// <param name="superText">The superscript text</param>
        /// <returns>The paragraph containing the superscript</returns>
        public Paragraph AddSuperscript(string baseText, string superText)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.SuperScript(
                    new M.Base(
                        new M.Run(
                            new M.Text { Text = baseText }
                        )
                    ),
                    new M.SuperArgument(
                        new M.Run(
                            new M.Text { Text = superText }
                        )
                    )
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds a subscript (like x₁) to the document
        /// </summary>
        /// <param name="baseText">The base text</param>
        /// <param name="subText">The subscript text</param>
        /// <returns>The paragraph containing the subscript</returns>
        public Paragraph AddSubscript(string baseText, string subText)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.SubScript(
                    new M.Base(
                        new M.Run(
                            new M.Text { Text = baseText }
                        )
                    ),
                    new M.SubArgument(
                        new M.Run(
                            new M.Text { Text = subText }
                        )
                    )
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds an integral to the document
        /// </summary>
        /// <param name="integrand">The integrand expression</param>
        /// <param name="lowerLimit">Lower limit</param>
        /// <param name="upperLimit">Upper limit</param>
        /// <returns>The paragraph containing the integral</returns>
        public Paragraph AddIntegral(string integrand, string lowerLimit, string upperLimit)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Delimiter(
                    new M.DelimiterProperties(
                        new M.BeginChar { Val = string.Empty },
                        new M.EndChar { Val = string.Empty }
                    ),
                    new M.Run(
                        new M.Text { Text = "∫" }
                    )
                )
            );
            
            // Add lower limit
            if (!string.IsNullOrEmpty(lowerLimit))
            {
                M.SubScript subScript = new M.SubScript(
                    new M.Base(
                        new M.Run(
                            new M.Text { Text = "∫" }
                        )
                    ),
                    new M.SubArgument(
                        new M.Run(
                            new M.Text { Text = lowerLimit }
                        )
                    )
                );
                
                // Add upper limit if provided
                if (!string.IsNullOrEmpty(upperLimit))
                {
                    M.SuperScript superScript = new M.SuperScript(
                        new M.Base(subScript),
                        new M.SuperArgument(
                            new M.Run(
                                new M.Text { Text = upperLimit }
                            )
                        )
                    );
                    
                    officeMath = new M.OfficeMath(superScript);
                }
                else
                {
                    officeMath = new M.OfficeMath(subScript);
                }
            }
            
            // Add integrand
            if (!string.IsNullOrEmpty(integrand))
            {
                M.Run integrandRun = new M.Run(
                    new M.Text { Text = " " + integrand }
                );
                
                officeMath.AppendChild(integrandRun);
            }
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds a matrix to the document
        /// </summary>
        /// <param name="rows">Number of rows</param>
        /// <param name="columns">Number of columns</param>
        /// <param name="values">Values for the matrix cells (row by row)</param>
        /// <returns>The paragraph containing the matrix</returns>
        public Paragraph AddMatrix(int rows, int columns, string[,] values)
        {
            if (values == null)
                throw new ArgumentNullException(nameof(values));
                
            if (values.GetLength(0) != rows || values.GetLength(1) != columns)
                throw new ArgumentException("The values array dimensions must match the specified rows and columns.");
                
            Paragraph para = new Paragraph();
            
            M.Matrix matrix = new M.Matrix();
            
            // Set the matrix properties
            M.MatrixProperties matrixProps = new M.MatrixProperties(
                new M.ColumnCount { Val = columns },
                new M.RowCount { Val = rows }
            );
            matrix.AppendChild(matrixProps);
            
            // Add the rows and values
            for (int i = 0; i < rows; i++)
            {
                M.MatrixRow matrixRow = new M.MatrixRow();
                
                for (int j = 0; j < columns; j++)
                {
                    M.MatrixColumn matrixColumn = new M.MatrixColumn(
                        new M.Run(
                            new M.Text { Text = values[i, j] }
                        )
                    );
                    
                    matrixRow.AppendChild(matrixColumn);
                }
                
                matrix.AppendChild(matrixRow);
            }
            
            M.OfficeMath officeMath = new M.OfficeMath(matrix);
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
        
        /// <summary>
        /// Adds an equation with parentheses to the document
        /// </summary>
        /// <param name="content">The content inside parentheses</param>
        /// <returns>The paragraph containing the equation</returns>
        public Paragraph AddParentheses(string content)
        {
            Paragraph para = new Paragraph();
            
            M.OfficeMath officeMath = new M.OfficeMath(
                new M.Delimiter(
                    new M.DelimiterProperties(
                        new M.BeginChar { Val = "(" },
                        new M.EndChar { Val = ")" }
                    ),
                    new M.Run(
                        new M.Text { Text = content }
                    )
                )
            );
            
            Run run = new Run();
            run.AppendChild(officeMath);
            para.AppendChild(run);
            
            _document.MainDocumentPart.Document.Body.AppendChild(para);
            return para;
        }
    }
} 