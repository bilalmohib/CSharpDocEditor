using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CSharpDoc
{
    /// <summary>
    /// Utility class for extracting and applying styles from existing Word documents
    /// </summary>
    public class StyleExtractor
    {
        /// <summary>
        /// Extracts styles from a template document and applies them to a target document
        /// </summary>
        /// <param name="templatePath">Path to the template document containing the styles</param>
        /// <param name="targetDocument">Target WordprocessingDocument to apply styles to</param>
        public static void ApplyStylesFromTemplate(string templatePath, WordprocessingDocument targetDocument)
        {
            using (WordprocessingDocument templateDoc = WordprocessingDocument.Open(templatePath, false))
            {
                StyleDefinitionsPart templateStylesPart = templateDoc.MainDocumentPart.StyleDefinitionsPart;
                if (templateStylesPart == null)
                {
                    throw new InvalidOperationException("Template document does not contain any styles.");
                }

                MainDocumentPart targetMainPart = targetDocument.MainDocumentPart;
                
                // Get existing styles part or create a new one
                StyleDefinitionsPart targetStylesPart = targetMainPart.StyleDefinitionsPart;
                if (targetStylesPart == null)
                {
                    targetStylesPart = targetMainPart.AddNewPart<StyleDefinitionsPart>();
                }
                
                // Clone all styles from template
                Styles newStyles = (Styles)templateStylesPart.Styles.CloneNode(true);
                targetStylesPart.Styles = newStyles;
                
                // If the document has numbering styles, copy those too
                if (templateDoc.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    NumberingDefinitionsPart numberingPart;
                    if (targetMainPart.NumberingDefinitionsPart == null)
                    {
                        numberingPart = targetMainPart.AddNewPart<NumberingDefinitionsPart>();
                    }
                    else
                    {
                        numberingPart = targetMainPart.NumberingDefinitionsPart;
                    }
                    
                    numberingPart.Numbering = (Numbering)templateDoc.MainDocumentPart.NumberingDefinitionsPart.Numbering.CloneNode(true);
                }
                
                // If the document has theme part, copy it too
                if (templateDoc.MainDocumentPart.ThemePart != null)
                {
                    ThemePart themePart;
                    if (targetMainPart.ThemePart == null)
                    {
                        themePart = targetMainPart.AddNewPart<ThemePart>();
                    }
                    else
                    {
                        themePart = targetMainPart.ThemePart;
                    }
                    
                    themePart.Theme = (DocumentFormat.OpenXml.Drawing.Theme)templateDoc.MainDocumentPart.ThemePart.Theme.CloneNode(true);
                }
                
                // If the document has fontTable part, copy it too
                if (templateDoc.MainDocumentPart.FontTablePart != null)
                {
                    FontTablePart fontTablePart;
                    if (targetMainPart.FontTablePart == null)
                    {
                        fontTablePart = targetMainPart.AddNewPart<FontTablePart>();
                    }
                    else
                    {
                        fontTablePart = targetMainPart.FontTablePart;
                    }
                    
                    fontTablePart.FontTable = (FontTable)templateDoc.MainDocumentPart.FontTablePart.FontTable.CloneNode(true);
                }
            }
        }

        /// <summary>
        /// Gets a list of available style IDs from a document
        /// </summary>
        /// <param name="documentPath">Path to the Word document</param>
        /// <returns>List of style IDs and their names</returns>
        public static Dictionary<string, string> GetAvailableStyles(string documentPath)
        {
            Dictionary<string, string> styleList = new Dictionary<string, string>();
            
            using (WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, false))
            {
                StyleDefinitionsPart stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
                if (stylesPart != null && stylesPart.Styles != null)
                {
                    foreach (Style style in stylesPart.Styles.Elements<Style>())
                    {
                        if (style.StyleId != null && style.StyleName != null)
                        {
                            string styleId = style.StyleId.Value;
                            string styleName = style.StyleName.Val.Value;
                            styleList[styleId] = styleName;
                        }
                    }
                }
            }
            
            return styleList;
        }

        /// <summary>
        /// Extends the WordDocumentUtils class to support template-based styling
        /// </summary>
        /// <param name="utils">The WordDocumentUtils instance</param>
        /// <param name="templatePath">Path to the template document</param>
        /// <param name="document">The WordprocessingDocument to apply styles to</param>
        public static void ApplyTemplateToDocument(string templatePath, WordprocessingDocument document)
        {
            ApplyStylesFromTemplate(templatePath, document);
        }
    }
} 