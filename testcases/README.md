# MarkMyDeck Test Cases

Test case data files for verifying Markdown-to-PowerPoint conversion.

Each `.md` file contains Markdown input, and the corresponding `.json` file
describes the expected conversion results.

## Structure

```
testcases/
├── basic_heading.md              # Single heading → single slide
├── basic_heading.json
├── multiple_slides.md            # Multiple H1s → multiple slides
├── multiple_slides.json
├── bold_text.md                  # Bold emphasis → bold run properties
├── bold_text.json
├── italic_text.md                # Italic emphasis → italic run properties
├── italic_text.json
├── inline_code.md                # Inline code → monospace font run
├── inline_code.json
├── code_block.md                 # Fenced code block → styled shape
├── code_block.json
├── hyperlink.md                  # Link → hyperlink run
├── hyperlink.json
├── unordered_list.md             # Bullet list → bullet items
├── unordered_list.json
├── ordered_list.md               # Numbered list → numbered items
├── ordered_list.json
├── table.md                      # Table → OOXML table
├── table.json
├── thematic_break.md             # --- → slide separator
├── thematic_break.json
├── block_quote.md                # Block quote → styled quote
├── block_quote.json
├── mixed_content.md              # Combined elements
├── mixed_content.json
├── negative/
│   ├── empty_input.json          # Empty string → ArgumentException
│   └── null_stream.json          # Null stream → ArgumentNullException
└── README.md
```

## Test Case JSON Schema

```json
{
  "description": "Human-readable description of the test case",
  "inputFile": "test_case_name.md",
  "expectedSlideCount": 1,
  "expectations": [
    {
      "type": "slide_count",
      "value": 2
    },
    {
      "type": "has_bold_runs",
      "slideIndex": 0
    },
    {
      "type": "has_table",
      "slideIndex": 0,
      "tableCount": 1
    }
  ]
}
```
