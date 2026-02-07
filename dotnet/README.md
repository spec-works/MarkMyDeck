# MarkMyDeck

A .NET library for converting CommonMark/GitHub Flavored Markdown to Microsoft PowerPoint (.pptx) presentations. The library targets .NET Standard 2.1 and .NET 10, and the CLI targets .NET 10.

## Features

- **Markdown → PowerPoint**: Convert CommonMark markdown to PowerPoint (.pptx) presentations using the Open XML SDK

### Currently Supported

✅ **Slide Generation**
- `# H1` and `## H2` headings create new slides
- `### H3` through `###### H6` render as styled text on the current slide
- `---` (thematic breaks) force a new slide
- Content between headings renders on the current slide

✅ **Block Elements**
- Headings (ATX: `# H1` through `###### H6`)
- Paragraphs
- Code blocks (fenced with ```) with syntax highlighting
- Block quotes (`>`)
- Thematic breaks (`---`, `***`, `___`) as slide separators
- Lists (ordered, unordered, nested)
- Tables (with headers, borders, and shading)

✅ **Inline Elements**
- Bold (`**text**` or `__text__`)
- Italic (`*text*` or `_text_`)
- Bold + Italic (`***text***`)
- Inline code (`` `code` ``)
- Links (`[text](url)`)
- Images (`![alt](url)`)
- Hard line breaks

✅ **Styling**
- Customizable fonts and colors
- Configurable heading/title sizes
- Code syntax highlighting for JSON, TypeSpec, and Bash
- Widescreen 16:9 slide format

## Installation

```bash
dotnet add package SpecWorks.MarkMyDeck
```

## Quick Start

```csharp
using MarkMyDeck;

// Convert markdown string to .pptx file
string markdown = "# Hello World\n\nThis is **bold** text.\n\n---\n\n# Slide 2\n\n- Item 1\n- Item 2";
MarkdownConverter.ConvertToPptx(markdown, "output.pptx");
```

### Convert to Byte Array

```csharp
byte[] pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);
```

### Stream-based Conversion

```csharp
using var inputStream = File.OpenRead("input.md");
using var outputStream = File.Create("output.pptx");
MarkdownConverter.ConvertToPptx(inputStream, outputStream);
```

### Async Conversion

```csharp
await MarkdownConverter.ConvertToPptxAsync(markdown, "output.pptx");
```

### Custom Styling

```csharp
using MarkMyDeck.Configuration;

var options = new ConversionOptions
{
    Styles = new SlideStyleConfiguration
    {
        DefaultFontName = "Arial",
        DefaultFontSize = 20,
        TitleFontSize = 40,
        TitleColor = "2E74B5",
        CodeFontName = "Fira Code",
        CodeFontSize = 16,
        CodeBackgroundColor = "282C34"
    },
    DocumentTitle = "My Presentation",
    Author = "John Doe"
};

MarkdownConverter.ConvertToPptx(markdown, "styled.pptx", options);
```

## Command-Line Interface

MarkMyDeck includes a CLI for converting markdown files from the command line.

### Installation

```bash
dotnet tool install --global SpecWorks.MarkMyDeck.CLI
```

Or run directly from the project:

```bash
dotnet run --project src/MarkMyDeck.CLI/MarkMyDeck.CLI.csproj -- convert -i input.md
```

### Usage

```bash
# Basic conversion
markmydeck convert -i README.md

# Specify output file
markmydeck convert -i input.md -o output.pptx

# Custom font and size
markmydeck convert -i document.md --font "Arial" --font-size 20

# With title metadata
markmydeck convert -i document.md --title "My Presentation"

# Verbose output
markmydeck convert -i document.md -v

# Force overwrite
markmydeck convert -i document.md --force

# View version
markmydeck version
```

### CLI Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--input` | `-i` | Input markdown file path (required) |
| `--output` | `-o` | Output file path (default: same name with .pptx) |
| `--verbose` | `-v` | Enable verbose output |
| `--force` | - | Overwrite output file if it exists |
| `--font` | `-f` | Default font name |
| `--font-size` | `-s` | Default font size (6-72 points) |
| `--title` | `-t` | Presentation title metadata |

## Architecture

MarkMyDeck uses a three-stage conversion pipeline:

1. **Parse**: Markdown is parsed into an AST using [Markdig](https://github.com/xoofx/markdig)
2. **Render**: The AST is traversed and converted to OpenXML Presentation elements
3. **Save**: The presentation is saved using [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK)

### Slide Strategy

- `# H1` and `## H2` headings create new slides
- `### H3` through `###### H6` are rendered as styled text within the current slide
- `---` thematic breaks force a new slide
- All other content (paragraphs, lists, tables, code blocks) renders on the current slide

## Project Structure

```
MarkMyDeck/
├── src/
│   ├── MarkMyDeck/                    # Core library (netstandard2.1;net10.0)
│   │   ├── MarkdownConverter.cs       # Public API
│   │   ├── Converters/
│   │   │   ├── OpenXmlPresentationRenderer.cs  # Main renderer
│   │   │   ├── BlockRenderers/        # Block element renderers
│   │   │   └── InlineRenderers/       # Inline element renderers
│   │   ├── SyntaxHighlighting/        # Syntax highlighting
│   │   ├── OpenXml/
│   │   │   ├── PresentationBuilder.cs # OpenXML presentation builder
│   │   │   └── SlideManager.cs        # Slide content management
│   │   └── Configuration/             # Configuration classes
│   └── MarkMyDeck.CLI/                # Command-line tool (net10.0)
└── tests/
    └── MarkMyDeck.Tests/              # Unit tests (net10.0)
```

## Requirements

- Library: .NET Standard 2.1 or .NET 10.0
- CLI: .NET 10.0
- Dependencies:
  - Markdig 0.37.0
  - DocumentFormat.OpenXml 3.1.0
  - ColorCode.Core 2.0.15

## Building from Source

```bash
git clone https://github.com/spec-works/MarkMyDeck.git
cd MarkMyDeck/dotnet
dotnet build
dotnet test
```

## License

MIT License

## Acknowledgments

- Built with [Markdig](https://github.com/xoofx/markdig) by Alexandre Mutel
- Uses [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) by Microsoft
- Implements [CommonMark](https://commonmark.org/) specification
- Inspired by [MarkMyWord](https://github.com/spec-works/MarkMyWord)
