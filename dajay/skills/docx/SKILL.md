# docx — Word Document Generation Skill

Generate `.docx` (Word Open XML) documents programmatically using `python-docx`.

## Capabilities

| Feature | Description |
|---------|-------------|
| **Headers / Footers** | Add title, section headers (H1–H3), and page footers |
| **Paragraphs** | Body text with configurable font, size, bold, italic, color, alignment |
| **Tables** | Create tables with headers, merged cells, borders, and styling |
| **Lists** | Bulleted and numbered lists |
| **Images** | Embed images (local or URLs) with size and alignment control |
| **Page Layout** | Page breaks, margins, orientation (portrait/landscape), page numbers |
| **Styles** | Apply built-in or custom styles; consistent formatting throughout |

## Parameters

When invoking this skill, provide a JSON specification with this structure:

```json
{
  "output_path": "path/to/output.docx",
  "document": {
    "orientation": "portrait|landscape",
    "margins": { "top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0 },
    "sections": [
      {
        "type": "heading|paragraph|table|list|image|page_break",
        "content": "...",
        "formatting": { "bold": false, "italic": false, "size": 12, "color": "000000", "alignment": "left|center|right" }
      }
    ]
  }
}
```

## Dependencies

- Python 3.8+
- `python-docx` library (`pip install python-docx`)
- `docx` CLI tool at `dajay/skills/docx/generate.py`

## Usage

```
python dajay/skills/docx/generate.py --spec spec.json
```