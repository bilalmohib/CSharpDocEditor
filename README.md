# C# Word Document Utilities using OpenXML

This library provides a set of utility functions for creating and manipulating Word documents programmatically using the OpenXML SDK. The utilities enable you to:

- Create and open Word documents
- Apply styles from template documents
- Generate tables with advanced formatting
- Add captions to tables and images with automatic numbering
- Insert and format images
- Create mathematical formulas
- Generate table of contents (TOC)

## Prerequisites

- .NET Framework 4.7+ or .NET Core 3.1+
- DocumentFormat.OpenXml NuGet package

## Installation

Add the OpenXML SDK to your project:

```bash
Install-Package DocumentFormat.OpenXml
```

## Usage Examples

### Basic Document Creation

```csharp
// Create a new document
WordDocumentUtils docUtils = new WordDocumentUtils();
docUtils.CreateDocument("MyDocument.docx");

// Add content
docUtils.AddParagraph("Hello World", "Heading1");
docUtils.AddParagraph("This is a sample paragraph.");

// Save and close
docUtils.SaveAndClose();
```

### Working with Tables

```csharp
// Create a table with 3 rows and 4 columns
var table = docUtils.CreateTable(3, 4);

// Set table content
docUtils.SetCellText(table, 0, 0, "Header 1");
docUtils.SetCellText(table, 0, 1, "Header 2");
// ... set other cells

// Add a caption to the table
docUtils.AddCaption("Sample Table", "Table");
```

### Advanced Table Formatting

```csharp
// Create a TableUtils instance
TableUtils tableUtils = new TableUtils(document);

// Create a table
var table = tableUtils.CreateTable(5, 3, "TableGrid");

// Format the table
tableUtils.SetTableBorders(table, TableBorderStyle.Single, 12, "000000");
tableUtils.SetTableAlignment(table, TableAlignment.Center);

// Apply alternating row colors
tableUtils.ApplyAlternatingRowShading(table, "F2F2F2", "FFFFFF", "4472C4");

// Customize header row
tableUtils.SetHeaderRow(table, 0, "FFFFFF", "4472C4", true);

// Merge cells
tableUtils.MergeCells(table, 1, 0, 1, 2);
```

### Working with Styles from Templates

```csharp
// Create a document
WordDocumentUtils docUtils = new WordDocumentUtils();
docUtils.CreateDocument("MyDocument.docx");

// Apply styles from a template
StyleExtractor.ApplyTemplateToDocument("Template.docx", docUtils._document);

// Get available styles
var styles = StyleExtractor.GetAvailableStyles("Template.docx");
foreach (var style in styles)
{
    Console.WriteLine($"Style ID: {style.Key}, Name: {style.Value}");
}
```

### Creating Mathematical Formulas

```csharp
// Create a MathUtils instance
MathUtils mathUtils = new MathUtils(document);

// Add a fraction
mathUtils.AddFraction("x + y", "2");

// Add a radical (square root)
mathUtils.AddRadical("x^2 + y^2");

// Add a matrix
string[,] matrixValues = {
    { "a", "b", "c" },
    { "d", "e", "f" },
    { "g", "h", "i" }
};
mathUtils.AddMatrix(3, 3, matrixValues);

// Add an integral
mathUtils.AddIntegral("f(x) dx", "0", "1");
```

## Class Overview

- **WordDocumentUtils**: Main utility class for Word document operations
- **StyleExtractor**: Utilities for extracting and applying styles from templates
- **TableUtils**: Advanced table creation and formatting utilities
- **MathUtils**: Utilities for creating mathematical formulas

## License

MIT 