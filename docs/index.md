# MarkMyDeck Documentation

Conversion from Markdown to PowerPoint presentations according to [CommonMark 0.31.2](https://spec.commonmark.org/0.31.2/) and [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/).

## What is MarkMyDeck?

MarkMyDeck is a .NET library and command-line tool for converting CommonMark-formatted Markdown into Microsoft PowerPoint (.pptx) presentations. It provides high-fidelity conversion with support for syntax highlighting, tables, lists, and all standard Markdown elements.

## Installation

### Library

Install the library via NuGet:

```bash
dotnet add package SpecWorks.MarkMyDeck
```

### CLI Tool

Install the command-line tool globally:

```bash
dotnet tool install --global SpecWorks.MarkMyDeck.CLI
```

## Features

- ✅ **CommonMark 0.31.2 Compliant** - Full implementation of CommonMark specification
- ✅ **ECMA-376 Compliant** - Standard Office Open XML presentation format
- ✅ **Syntax Highlighting** - Code blocks with language-specific highlighting
- ✅ **Tables** - Full table support with alignment
- ✅ **Links** - Preserve hyperlinks
- ✅ **Styles** - Proper heading levels, emphasis, and formatting
- ✅ **Lists** - Ordered and unordered lists with bullet rendering
- ✅ **Code Blocks** - Fenced code blocks with background styling
- ✅ **Inline Formatting** - Bold, italic, inline code, line breaks
- ✅ **Slide Separation** - Headings and thematic breaks create new slides
- ✅ **CLI Tool** - Command-line interface for batch processing
- ✅ **Type-Safe API** - Strong typing with nullable reference types
- ✅ **Multi-Target** - Supports .NET Standard 2.1 and .NET 10.0

## Quick Start

### Library Usage

```csharp
using MarkMyDeck;

// Convert Markdown to PowerPoint bytes
var markdown = "# My Presentation\n\nHello, World!\n\n# Slide 2\n\n- Item 1\n- Item 2";
byte[] pptxBytes = MarkdownConverter.ConvertToPptxBytes(markdown);
File.WriteAllBytes("output.pptx", pptxBytes);
```

### Stream-based Usage

```csharp
using MarkMyDeck;

var markdown = File.ReadAllText("slides.md");
using var outputStream = File.Create("presentation.pptx");
MarkdownConverter.ConvertToPptx(markdown, outputStream);
```

### CLI Usage

```bash
markmydeck convert slides.md -o presentation.pptx
```

## Slide Structure

MarkMyDeck uses heading levels to determine slide boundaries:

- **`# Heading 1`** creates a new slide with a title
- **`---` (thematic break)** creates a slide break
- Content between headings becomes the slide body

### Example

```markdown
# Welcome

This is the first slide.

# Features

- Feature 1
- Feature 2
- Feature 3

---

# Code Example

​```csharp
Console.WriteLine("Hello!");
​```
```

This produces three slides: a welcome slide, a features slide with bullet points, and a code example slide.

## Supported Markdown Elements

| Element | Support | Notes |
|---------|---------|-------|
| Headings (H1-H6) | ✅ | H1 creates new slides |
| Paragraphs | ✅ | Body text on slides |
| Bold | ✅ | `**bold**` |
| Italic | ✅ | `*italic*` |
| Inline Code | ✅ | Rendered in Cascadia Code font |
| Code Blocks | ✅ | With syntax highlighting and background |
| Links | ✅ | Clickable hyperlinks |
| Unordered Lists | ✅ | Bullet point rendering |
| Ordered Lists | ✅ | Numbered list rendering |
| Tables | ✅ | Full table with header row |
| Block Quotes | ✅ | Styled quote blocks |
| Thematic Breaks | ✅ | Slide separators |
| Line Breaks | ✅ | Preserved in output |

## Syntax Highlighting

Code blocks with language identifiers receive syntax highlighting:

````markdown
```json
{"key": "value"}
```

```csharp
var x = 42;
```

```bash
echo "Hello"
```
````

Supported languages include JSON, C#, Bash, HTTP, TypeSpec, and more via the ColorCode library.

## Specifications

This project implements:

- [CommonMark 0.31.2](https://spec.commonmark.org/0.31.2/) - Markdown parsing
- [GitHub Flavored Markdown](https://github.github.com/gfm/) - Extended Markdown features
- [ECMA-376 5th edition](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) - Office Open XML presentation format
- [ISO/IEC 29500](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/) - Office implementation information

## API Reference

See the [API Reference](api/) for detailed documentation of all public types and methods.
