import argparse
import json
import sys

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn


def _parse_alignment(val):
    mapping = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    return mapping.get(val, WD_ALIGN_PARAGRAPH.LEFT)


def _apply_run_format(run, fmt):
    if fmt.get("bold"):
        run.bold = True
    if fmt.get("italic"):
        run.italic = True
    if fmt.get("size"):
        run.font.size = Pt(fmt["size"])
    if fmt.get("color"):
        run.font.color.rgb = RGBColor.from_string(fmt["color"])
    if fmt.get("name"):
        run.font.name = fmt["name"]


def _add_heading(doc, section):
    level = section.get("level", 1)
    heading = doc.add_heading(section["content"], level=level)
    fmt = section.get("formatting", {})
    if fmt:
        for run in heading.runs:
            _apply_run_format(run, fmt)


def _add_paragraph(doc, section):
    p = doc.add_paragraph()
    fmt = section.get("formatting", {})
    if fmt.get("alignment"):
        p.alignment = _parse_alignment(fmt["alignment"])
    run = p.add_run(section["content"])
    _apply_run_format(run, fmt)


def _add_table(doc, section):
    data = section["data"]
    rows = len(data)
    cols = max(len(row) for row in data) if data else 0
    table = doc.add_table(rows=rows, cols=cols)
    table.style = "Table Grid"
    align = section.get("formatting", {}).get("alignment")
    if align:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER if align == "center" else WD_TABLE_ALIGNMENT.LEFT
    for i, row_data in enumerate(data):
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i, j)
            cell.text = str(cell_text)
            if i == 0 and section.get("header", True):
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True


def _add_list(doc, section):
    ordered = section.get("ordered", False)
    for item in section["items"]:
        p = doc.add_paragraph(style="List Number" if ordered else "List Bullet")
        run = p.add_run(item)
        fmt = section.get("formatting", {})
        if fmt:
            _apply_run_format(run, fmt)


def _add_image(doc, section):
    path = section["path"]
    width = section.get("width")
    height = section.get("height")
    kwargs = {}
    if width:
        kwargs["width"] = Inches(width)
    if height:
        kwargs["height"] = Inches(height)
    doc.add_picture(path, **kwargs)
    if section.get("alignment"):
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = _parse_alignment(section["alignment"])


def build_document(spec):
    doc = Document()
    doc_config = spec.get("document", {})
    margins = doc_config.get("margins", {})
    orientation = doc_config.get("orientation", "portrait")
    section = doc.sections[0]
    section.top_margin = Inches(margins.get("top", 1.0))
    section.bottom_margin = Inches(margins.get("bottom", 1.0))
    section.left_margin = Inches(margins.get("left", 1.0))
    section.right_margin = Inches(margins.get("right", 1.0))
    if orientation == "landscape":
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    for s in doc_config.get("sections", []):
        t = s["type"]
        if t == "heading":
            _add_heading(doc, s)
        elif t == "paragraph":
            _add_paragraph(doc, s)
        elif t == "table":
            _add_table(doc, s)
        elif t == "list":
            _add_list(doc, s)
        elif t == "image":
            _add_image(doc, s)
        elif t == "page_break":
            doc.add_page_break()
    output_path = spec["output_path"]
    doc.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Generate a .docx document from a JSON spec")
    parser.add_argument("--spec", required=True, help="Path to JSON specification file")
    args = parser.parse_args()
    with open(args.spec, "r", encoding="utf-8") as f:
        spec = json.load(f)
    path = build_document(spec)
    print(f"Document generated: {path}")


if __name__ == "__main__":
    main()
