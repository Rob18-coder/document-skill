# docx — Word Document Generation Skill (Node.js)

Generate `.docx` (Word Open XML) documents programmatically using the [`docx`](https://www.npmjs.com/package/docx) npm package.

## Capabilities

| Feature | Description |
|---------|-------------|
| **Headings** | H1–H6 with configurable formatting |
| **Paragraphs** | Body text with font, size, bold, italic, color, alignment, line spacing |
| **Tables** | Create tables with headers, cell styling, percentage widths |
| **Lists** | Bulleted and numbered lists with indentation |
| **Page Layout** | Page breaks, margins (inches), orientation (portrait/landscape) |
| **Numbering** | Auto-numbered headings and list styles |

## Parameters

Provide a JSON specification file:

```json
{
  "output_path": "path/to/output.docx",
  "document": {
    "orientation": "portrait|landscape",
    "margins": { "top": 1.0, "bottom": 1.0, "left": 1.0, "right": 1.0 },
    "sections": [
      { "type": "heading", "content": "Title", "level": 1, "formatting": { "bold": true, "size": 24, "color": "2F5496", "alignment": "center" } },
      { "type": "paragraph", "content": "Body text here.", "formatting": { "size": 11, "alignment": "justify" } },
      { "type": "table", "data": [["Name", "Value"], ["A", "1"], ["B", "2"]], "header": true },
      { "type": "list", "items": ["Item one", "Item two"], "ordered": false },
      { "type": "page_break" }
    ]
  }
}
```

## Dependencies

- Node.js 18+
- npm packages: `docx`, `yargs`

Install:
```bash
cd dajay/skills/docx && npm install
```

## Usage

```bash
node dajay/skills/docx/generate.js --spec spec.json
```

Or via npm script:
```bash
cd dajay/skills/docx && npm run generate -- --spec spec.json
```