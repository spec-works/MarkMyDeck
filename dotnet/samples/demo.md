# MarkMyDeck Demo

Welcome to MarkMyDeck — a Markdown to PowerPoint converter.

---

## Features

- Convert Markdown to `.pptx` presentations
- Syntax highlighting for code blocks
- Tables, lists, links, and more
- Built with the same libraries as **MarkMyWord**

![Portrait Photo](portrait-image.jpg)

---

## Code Example

Here's a JSON snippet with syntax highlighting:

```json
{
  "name": "MarkMyDeck",
  "version": "0.1.0",
  "features": ["slides", "code", "tables"],
  "enabled": true
}
```

---

## Bash Example

```bash
#!/bin/bash
echo "Converting markdown to slides..."
for file in *.md; do
    markmydeck convert -i "$file"
done
```

---

## Tables

| Feature | Status |
|---------|--------|
| Headings | ✅ Supported |
| Bold/Italic | ✅ Supported |
| Code Blocks | ✅ Supported |
| Tables | ✅ Supported |
| Lists | ✅ Supported |
| Links | ✅ Supported |

---

## Lists

### Unordered

- First item
- Second item
- Third item with **bold** text

### Ordered

1. Step one
2. Step two
3. Step three

---

## Links and Formatting

Visit [MarkMyDeck on GitHub](https://github.com/spec-works/MarkMyDeck) for more info.

This text has **bold**, *italic*, and `inline code` formatting.

> This is a block quote — great for callouts and notes.

---

# Thank You!

Built with Markdig, DocumentFormat.OpenXml, and ColorCode.Core.

---

# Images

Here's an embedded image:

![Sample Photo](test-image.jpg)
