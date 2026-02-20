---
name: markmydeck-cli
description: >
  Convert Markdown files to Microsoft PowerPoint (.pptx) presentations using the MarkMyDeck CLI tool.
  Use this skill when the user wants to create a PowerPoint deck from Markdown, generate slides from
  documentation, or convert .md files to .pptx format. Activate when the user mentions PowerPoint,
  slides, presentations, decks, or .pptx files in the context of Markdown conversion.
license: MIT
compatibility: Requires .NET 10.0 SDK. Works on Windows, macOS, and Linux.
metadata:
  author: spec-works
  version: "1.0"
  repository: https://github.com/spec-works/MarkMyDeck
---

# MarkMyDeck CLI — Markdown to PowerPoint

Convert CommonMark/GitHub Flavored Markdown to Microsoft PowerPoint (.pptx) presentations.

## Prerequisites — Installing .NET

The MarkMyDeck CLI requires .NET 10.0. Install the SDK for your platform:

### Windows

```powershell
# Using winget (recommended)
winget install Microsoft.DotNet.SDK.10

# Or download from https://dotnet.microsoft.com/download/dotnet/10.0
```

### macOS

```bash
# Using Homebrew (recommended)
brew install dotnet-sdk

# Or download from https://dotnet.microsoft.com/download/dotnet/10.0
```

### Linux (Ubuntu/Debian)

```bash
# Add Microsoft package repository
wget https://packages.microsoft.com/config/ubuntu/24.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

# Install .NET SDK
sudo apt-get update
sudo apt-get install -y dotnet-sdk-10.0
```

### Linux (Fedora/RHEL)

```bash
sudo dnf install dotnet-sdk-10.0
```

### Verify Installation

```bash
dotnet --version
# Should show 10.x.x
```

## Installing the CLI Tool

```bash
dotnet tool install --global SpecWorks.MarkMyDeck.CLI
```

Verify it works:

```bash
markmydeck version
```

## Basic Usage

```bash
# Convert a Markdown file to PowerPoint
markmydeck convert -i presentation.md

# Specify output file
markmydeck convert -i input.md -o slides.pptx

# Custom font and size
markmydeck convert -i input.md --font "Arial" --font-size 20

# With title metadata
markmydeck convert -i input.md --title "My Presentation"

# Verbose output
markmydeck convert -i input.md -v

# Force overwrite existing file
markmydeck convert -i input.md --force
```

### CLI Options

| Option | Alias | Description |
|--------|-------|-------------|
| `--input` | `-i` | Input markdown file path (required) |
| `--output` | `-o` | Output file path (default: same name with .pptx) |
| `--verbose` | `-v` | Enable verbose output |
| `--force` | — | Overwrite output file if it exists |
| `--font` | `-f` | Default font name |
| `--font-size` | `-s` | Default font size (6–72 points) |
| `--title` | `-t` | Presentation title metadata |

## Structuring Markdown for Best Slide Output

Understanding how MarkMyDeck maps Markdown to slides is critical for good results.

### Slide Boundaries

New slides are created by:

- **`# H1` headings** — creates a new slide with a title
- **`## H2` headings** — creates a new slide with a title
- **`---` thematic breaks** — forces a new slide (no title)

Everything between slide boundaries renders on the current slide.

### Heading Hierarchy

- `# H1` and `## H2` → **slide title** (creates new slide)
- `### H3` through `###### H6` → **styled text** on the current slide (does NOT create a new slide)

### Recommended Slide Structure

```markdown
# Slide Title

Content for this slide: paragraphs, lists, tables, code blocks.
Keep content concise — slides aren't documents.

---

## Another Slide

- Bullet points work well on slides
- Keep to 4-6 bullets per slide
- Use **bold** for emphasis

---

# Code Example

```json
{
  "key": "value"
}
```

---

## Data Table

| Column A | Column B |
|----------|----------|
| Data 1   | Data 2   |
```

### Best Practices

1. **One idea per slide** — use `---` or headings to break content into digestible slides
2. **Short bullet lists** — 4–6 items maximum per slide
3. **Use H3–H6 for sub-sections** within a slide, not for new slides
4. **Code blocks stay compact** — long code blocks will overflow; keep to ~15 lines
5. **Tables should be narrow** — 3–5 columns work best; wide tables may be clipped
6. **Images are supported** — `![alt](path)` for local files or URLs

## Supported Markdown Elements

### Block Elements

| Element | Syntax | Slide Behavior |
|---------|--------|----------------|
| Heading 1–2 | `# H1`, `## H2` | Creates new slide with title |
| Heading 3–6 | `### H3` – `###### H6` | Styled text on current slide |
| Paragraph | Plain text | Body text on current slide |
| Code block | ` ```lang ` | Styled code with syntax highlighting |
| Block quote | `> text` | Styled quote block |
| Thematic break | `---`, `***`, `___` | Forces new slide |
| Unordered list | `- item` | Bullet list |
| Ordered list | `1. item` | Numbered list |
| Nested lists | Indented items | Proper indentation preserved |
| Table | GFM table syntax | Formatted table with headers |

### Inline Elements

| Element | Syntax | Result |
|---------|--------|--------|
| Bold | `**text**` | Bold run |
| Italic | `*text*` | Italic run |
| Bold+Italic | `***text***` | Bold italic run |
| Inline code | `` `code` `` | Monospace font |
| Link | `[text](url)` | Clickable hyperlink |
| Image | `![alt](url)` | Embedded image |
| Hard line break | Two spaces or `\` at EOL | Line break within slide |

### Syntax Highlighting

Code blocks with language identifiers get syntax highlighting:

- **JSON** — property names, strings, numbers, booleans
- **TypeSpec** — keywords, types, decorators, comments
- **Bash/Shell** — commands, keywords, variables, strings

## Limitations — What MarkMyDeck Cannot Do

Be aware of these limitations when authoring Markdown for slides:

1. **No slide templates/themes** — output uses a default blank theme; you cannot apply custom PowerPoint templates
2. **No speaker notes** — Markdown has no syntax for speaker notes; they are not generated
3. **No slide transitions or animations** — static slides only
4. **No custom slide layouts** — all slides use a single content layout
5. **No embedded video or audio** — media embeds are not supported
6. **No math/LaTeX rendering** — math formulas are rendered as plain text
7. **No Mermaid diagrams** — diagram code blocks render as plain code
8. **No task lists** — `- [ ]` checkboxes render as plain text
9. **No footnotes** — footnote syntax is not converted
10. **No strikethrough** — `~~text~~` is not rendered with strikethrough styling
11. **Limited image sizing** — images are embedded at natural size; no resize controls
12. **Long content overflow** — content that exceeds slide dimensions is clipped, not auto-paginated
13. **Limited syntax highlighting languages** — only JSON, TypeSpec, and Bash are highlighted; other languages render as plain monospace

## Installing This Skill

### GitHub Copilot CLI (personal)

```bash
# Clone and copy
git clone https://github.com/spec-works/MarkMyDeck.git /tmp/MarkMyDeck
mkdir -p ~/.copilot/skills/markmydeck-cli
cp -r /tmp/MarkMyDeck/skills/markmydeck-cli/* ~/.copilot/skills/markmydeck-cli/
```

### GitHub Copilot CLI (project)

```bash
# Add to your repository
mkdir -p .github/skills/markmydeck-cli
cp -r /path/to/MarkMyDeck/skills/markmydeck-cli/* .github/skills/markmydeck-cli/
git add .github/skills/
git commit -m "Add MarkMyDeck CLI agent skill"
```

### Claude Code (personal)

```bash
mkdir -p ~/.claude/skills/markmydeck-cli
cp -r /path/to/MarkMyDeck/skills/markmydeck-cli/* ~/.claude/skills/markmydeck-cli/
```

### Claude Code (project)

```bash
mkdir -p .claude/skills/markmydeck-cli
cp -r /path/to/MarkMyDeck/skills/markmydeck-cli/* .claude/skills/markmydeck-cli/
```

### VS Code / Cursor (project)

```bash
mkdir -p .github/skills/markmydeck-cli
cp -r /path/to/MarkMyDeck/skills/markmydeck-cli/* .github/skills/markmydeck-cli/
```

After installing, restart your agent session to pick up the new skill.

## Further Reference

See [references/REFERENCE.md](references/REFERENCE.md) for detailed examples, the full architecture overview, and advanced styling options.
