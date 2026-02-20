# MarkMyDeck CLI — Detailed Reference

## Architecture

MarkMyDeck uses a three-stage conversion pipeline:

1. **Parse** — Markdown is parsed into an AST using [Markdig](https://github.com/xoofx/markdig)
2. **Render** — The AST is traversed and converted to OpenXML Presentation elements
3. **Save** — The presentation is saved using [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK)

```
Markdown Text
     ↓
  Markdig Parser
     ↓
  AST (Syntax Tree)
     ↓
  OpenXmlPresentationRenderer
     ↓
  OpenXML Presentation Elements
     ↓
  PowerPoint Document (.pptx)
```

## Slide Generation Strategy — Detailed

The renderer walks the AST and uses these rules to decide slide boundaries:

| Markdown Element | Slide Behavior |
|-----------------|----------------|
| `# H1` | Starts a new slide; heading text becomes the slide title |
| `## H2` | Starts a new slide; heading text becomes the slide title |
| `### H3` | Renders as large styled text on the current slide |
| `#### H4` | Renders as medium styled text on the current slide |
| `##### H5` | Renders as small styled text on the current slide |
| `###### H6` | Renders as small styled text on the current slide |
| `---` / `***` / `___` | Forces a new slide with no title |
| Any other block | Renders on the current slide |

**Important:** If content appears before the first heading or thematic break, it is placed on an implicit first slide with no title.

## Presentation Format

- **Aspect ratio:** 16:9 widescreen (default)
- **Default font:** Calibri
- **Default font size:** 18pt body text
- **Title font size:** 36pt

## Custom Styling via CLI

```bash
# Change the font
markmydeck convert -i input.md --font "Segoe UI"

# Change font size
markmydeck convert -i input.md --font-size 16

# Combine options
markmydeck convert -i input.md --font "Arial" --font-size 20 --title "Q4 Results"
```

## Programmatic Styling (C# Library)

If using the MarkMyDeck library directly instead of the CLI:

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
    Author = "Jane Doe"
};

MarkdownConverter.ConvertToPptx(markdown, "styled.pptx", options);
```

## Syntax Highlighting Details

### JSON

```json
{
  "name": "MarkMyDeck",
  "version": "1.0.0",
  "enabled": true,
  "count": 42
}
```

Highlighted elements: property names, string values, numbers, booleans (`true`/`false`/`null`).

### TypeSpec

```typespec
@service
namespace MyApi;

model Pet {
  name: string;
  age: int32;
}

@route("/pets")
interface Pets {
  list(): Pet[];
}
```

Highlighted elements: keywords, types, decorators, comments.

### Bash

```bash
#!/bin/bash
echo "Converting files..."
for file in *.md; do
    markmydeck convert -i "$file" -o "${file%.md}.pptx"
done
```

Highlighted elements: commands, keywords, variables, strings, comments.

### Unsupported Languages

Code blocks with other language identifiers (e.g., `python`, `java`, `csharp`) render as plain monospace text without color highlighting.

## Example: Complete Presentation Markdown

```markdown
# Quarterly Business Review

Q4 2025 Results and Q1 2026 Outlook

---

## Revenue Summary

| Metric | Q3 | Q4 | Change |
|--------|------|------|--------|
| Revenue | $10M | $12M | +20% |
| Users | 50K | 65K | +30% |
| NPS | 72 | 78 | +6 |

---

## Key Achievements

- Launched v2.0 of the platform
- Expanded to 3 new markets
- Reduced churn by **15%**
- Achieved SOC 2 compliance

---

## Technical Highlights

### API Performance

```json
{
  "p50_latency_ms": 12,
  "p99_latency_ms": 85,
  "uptime_percent": 99.97
}
```

---

## Q1 2026 Roadmap

1. Mobile app launch
2. Enterprise SSO integration
3. Advanced analytics dashboard
4. Partner API program

---

# Thank You

Questions? Visit [our docs](https://example.com/docs)
```

This produces 6 slides:
1. Title slide: "Quarterly Business Review" with subtitle text
2. Revenue Summary table
3. Key Achievements bullet list
4. Technical Highlights with code block
5. Q1 2026 Roadmap numbered list
6. Thank You slide with link

## Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| Markdig | 0.37.0 | Markdown parsing |
| DocumentFormat.OpenXml | 3.1.0 | PowerPoint generation |
| ColorCode.Core | 2.0.15 | Syntax highlighting |

## Troubleshooting

### "dotnet tool" command not found

Ensure the .NET SDK (not just runtime) is installed. The `dotnet tool` command requires the SDK.

### Output file already exists

Use `--force` to overwrite, or specify a different output path with `-o`.

### Images not appearing

- Ensure image files exist at the path specified relative to the input markdown file
- URL-based images require network access at conversion time
- Supported formats: PNG, JPG/JPEG

### Code block overflows the slide

Keep code blocks to approximately 15 lines. Longer blocks will be clipped at the bottom of the slide. Consider splitting into multiple slides using `---`.

### Table columns too wide

Tables with more than 5 columns may not fit the slide width. Shorten header text or reduce the number of columns.
