import argparse
import json
import sys
import os
import tempfile
import urllib.request

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls


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


def _apply_paragraph_format(p, fmt):
    if fmt.get("alignment"):
        p.alignment = _parse_alignment(fmt["alignment"])
    if fmt.get("space_before") is not None:
        p.paragraph_format.space_before = Pt(fmt["space_before"])
    if fmt.get("space_after") is not None:
        p.paragraph_format.space_after = Pt(fmt["space_after"])
    if fmt.get("line_spacing") is not None:
        p.paragraph_format.line_spacing = fmt["line_spacing"]


def set_cell_background(cell, fill_hex):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))


def _add_heading(doc, section):
    level = section.get("level", 1)
    heading = doc.add_heading(section["content"], level=level)
    fmt = section.get("formatting", {})
    if fmt:
        _apply_paragraph_format(heading, fmt)
        for run in heading.runs:
            _apply_run_format(run, fmt)


def _add_paragraph(doc, section):
    p = doc.add_paragraph()
    fmt = section.get("formatting", {})
    if fmt:
        _apply_paragraph_format(p, fmt)
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
    
    # Populate cell contents
    for i, row_data in enumerate(data):
        for j, cell_text in enumerate(row_data):
            cell = table.cell(i, j)
            cell.text = str(cell_text)
            if i == 0 and section.get("header", True):
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True

    # Apply merges
    merges = section.get("merges", [])
    for merge_info in merges:
        start_row = merge_info.get("start_row")
        start_col = merge_info.get("start_col")
        end_row = merge_info.get("end_row")
        end_col = merge_info.get("end_col")
        if start_row is not None and start_col is not None and end_row is not None and end_col is not None:
            try:
                cell_start = table.cell(start_row, start_col)
                cell_end = table.cell(end_row, end_col)
                cell_start.merge(cell_end)
            except Exception as e:
                print(f"Warning: Cell merge failed for {merge_info}: {e}", file=sys.stderr)

    # Apply cell styling
    cell_styles = section.get("cell_styles", [])
    for cs in cell_styles:
        row_idx = cs.get("row")
        col_idx = cs.get("col")
        if row_idx is not None and col_idx is not None:
            try:
                cell = table.cell(row_idx, col_idx)
                bg = cs.get("background_color")
                if bg:
                    set_cell_background(cell, bg)
                
                align = cs.get("alignment")
                for p in cell.paragraphs:
                    if align:
                        p.alignment = _parse_alignment(align)
                    # Support applying formatting to runs in the cell
                    for run in p.runs:
                        _apply_run_format(run, cs)
            except Exception as e:
                print(f"Warning: Applying cell styles failed for {cs}: {e}", file=sys.stderr)


def _add_list(doc, section):
    ordered = section.get("ordered", False)
    for item in section["items"]:
        p = doc.add_paragraph(style="List Number" if ordered else "List Bullet")
        run = p.add_run(item)
        fmt = section.get("formatting", {})
        if fmt:
            _apply_paragraph_format(p, fmt)
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
    
    temp_file = None
    if path.startswith("http://") or path.startswith("https://"):
        try:
            with urllib.request.urlopen(path) as response:
                content = response.read()
            suffix = os.path.splitext(path.split('?')[0])[1]
            if not suffix or len(suffix) > 5:
                suffix = ".jpg"
            fd, temp_path = tempfile.mkstemp(suffix=suffix)
            with os.fdopen(fd, 'wb') as tmp:
                tmp.write(content)
            path = temp_path
            temp_file = temp_path
        except Exception as e:
            print(f"Error downloading image from URL {section['path']}: {e}", file=sys.stderr)

    try:
        doc.add_picture(path, **kwargs)
        if section.get("alignment"):
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = _parse_alignment(section["alignment"])
    finally:
        if temp_file and os.path.exists(temp_file):
            try:
                os.remove(temp_file)
            except OSError:
                pass


def _setup_headers_footers(doc, doc_config):
    section = doc.sections[0]
    
    # Header
    header_config = doc_config.get("header")
    if header_config:
        header = section.header
        p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
        p.text = header_config.get("text", "")
        align = header_config.get("alignment")
        if align:
            p.alignment = _parse_alignment(align)
        fmt = header_config.get("formatting", {})
        if fmt:
            _apply_paragraph_format(p, fmt)
            for run in p.runs:
                _apply_run_format(run, fmt)

    # Footer
    footer_config = doc_config.get("footer")
    if footer_config:
        footer = section.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        text = footer_config.get("text", "")
        p.text = text
        align = footer_config.get("alignment")
        if align:
            p.alignment = _parse_alignment(align)
        fmt = footer_config.get("formatting", {})
        if fmt:
            _apply_paragraph_format(p, fmt)
            for run in p.runs:
                _apply_run_format(run, fmt)
        
        if footer_config.get("show_page_number", False):
            sep = " - Page " if text else "Page "
            run_sep = p.add_run(sep)
            if fmt:
                _apply_run_format(run_sep, fmt)
            run_num = p.add_run()
            if fmt:
                _apply_run_format(run_num, fmt)
            
            fldChar1 = OxmlElement('w:fldChar')
            fldChar1.set(qn('w:fldCharType'), 'begin')
            instrText = OxmlElement('w:instrText')
            instrText.set(qn('xml:space'), 'preserve')
            instrText.text = "PAGE"
            fldChar2 = OxmlElement('w:fldChar')
            fldChar2.set(qn('w:fldCharType'), 'separate')
            fldChar3 = OxmlElement('w:fldChar')
            fldChar3.set(qn('w:fldCharType'), 'end')
            
            run_num._r.append(fldChar1)
            run_num._r.append(instrText)
            run_num._r.append(fldChar2)
            run_num._r.append(fldChar3)


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
    
    _setup_headers_footers(doc, doc_config)
    
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

