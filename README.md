# HwpSharp

HwpSharp is a .NET library for parsing and manipulating HWP (Hangul Word Processor) files.

## Attribution

This library is largely based on the work from https://github.com/forcom/HwpSharp. We acknowledge and are grateful for their contribution to the open-source community.

## Features

- Parse HWP 5.0 file format
- Read document metadata and properties
- Access document content including text, paragraphs, and sections
- Support for multi-section documents

## Installation

You can install HwpSharp via NuGet:

```
dotnet add package HwpSharp
```

## Usage

Basic usage example:

```csharp
using HwpSharp.Hwp5;

// Open an HWP document
var document = new Document("path/to/document.hwp");

// Access document information
var sectionCount = document.DocumentInformation.DocumentProperty.SectionCount;
var startPageNumber = document.DocumentInformation.DocumentProperty.StartPageNumber;

// Read text from the first paragraph of the first section
var firstParagraph = document.BodyText.Sections[0].DataRecords
    .Where(r => r.TagId == ParagraphText.ParagraphTextTagId)
    .Cast<ParagraphText>()
    .FirstOrDefault();

var text = firstParagraph?.Text;
```

## Development

To build and run tests:

```
dotnet restore
dotnet build
dotnet test
```

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgements

We would like to express our gratitude to the original authors and contributors of the [forcom/HwpSharp](https://github.com/forcom/HwpSharp) project, which served as the foundation for this library.
