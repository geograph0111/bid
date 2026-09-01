from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
BLUE = "2F80ED"
LIGHT_BLUE = "EAF3FF"
TEXT = "243447"
MUTED = "667085"
BORDER = "D9E5F2"
WHITE = "FFFFFF"


def shade(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def cell_border(cell, color=BORDER, size="6"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), size)
        element.set(qn("w:color"), color)


def set_cell_margins(cell, top=100, start=120, bottom=100, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn("w:" + m))
        if node is None:
            node = OxmlElement("w:" + m)
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def add_text(paragraph, text, size=10.5, bold=False, color=TEXT):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    return run


def add_page_title(doc, page: dict[str, Any]):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    add_text(p, page.get("code", ""), 9, True, BLUE)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    add_text(p, page.get("title", ""), 22, True, TEXT)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(12)
    add_text(p, page.get("summary", ""), 10.5, False, MUTED)


def add_screenshot_placeholder(doc, path: str | None, label: str):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Cm(16.2)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    shade(cell, "F7FAFE")
    cell_border(cell)
    set_cell_margins(cell, 500, 200, 500, 200)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if path and Path(path).exists():
        run = p.add_run()
        run.add_picture(path, width=Cm(15.4))
    else:
        add_text(p, f"截图占位：{label}\n将实际 PC / 手机界面截图放入 assets/screenshots 后即可生成。", 10, False, MUTED)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_feature_table(doc, features: list[dict[str, str]]):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)
    add_text(p, "功能说明", 13, True, TEXT)

    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.columns[0].width = Cm(4.0)
    table.columns[1].width = Cm(12.2)
    hdr = table.rows[0].cells
    for i, text in enumerate(("功能", "说明")):
        shade(hdr[i], BLUE)
        set_cell_margins(hdr[i])
        p = hdr[i].paragraphs[0]
        add_text(p, text, 10, True, WHITE)
    set_repeat_table_header(table.rows[0])

    for feature in features:
        cells = table.add_row().cells
        for c in cells:
            cell_border(c)
            set_cell_margins(c)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(cells[0], LIGHT_BLUE)
        add_text(cells[0].paragraphs[0], feature.get("name", ""), 10, True, BLUE)
        add_text(cells[1].paragraphs[0], feature.get("description", ""), 9.5, False, TEXT)


def add_values(doc, values: list[str]):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    add_text(p, "应用价值", 13, True, TEXT)
    table = doc.add_table(rows=1, cols=len(values) or 1)
    table.autofit = False
    for i, value in enumerate(values):
        cell = table.cell(0, i)
        shade(cell, LIGHT_BLUE)
        cell_border(cell, LIGHT_BLUE)
        set_cell_margins(cell, 120, 150, 120, 150)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, "● " + value, 9.5, True, BLUE)


def add_mobile_gallery(doc, page):
    screenshots = page.get("screenshots", [])
    if not screenshots:
        add_screenshot_placeholder(doc, None, "手机端界面组")
        return
    table = doc.add_table(rows=1, cols=min(3, len(screenshots)))
    table.autofit = False
    for i, path in enumerate(screenshots[:3]):
        cell = table.cell(0, i)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(cell, "F7FAFE")
        cell_border(cell)
        set_cell_margins(cell, 180, 120, 180, 120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if Path(path).exists():
            p.add_run().add_picture(path, height=Cm(10.5))
        else:
            add_text(p, "手机截图占位", 9, False, MUTED)


def add_standard_page(doc, page):
    add_page_title(doc, page)
    add_screenshot_placeholder(doc, page.get("screenshot"), page.get("title", "产品界面"))
    add_feature_table(doc, page.get("features", []))
    add_values(doc, page.get("values", []))


def add_page(doc, page):
    kind = page.get("type", "standard")
    if kind == "mobile_gallery":
        add_page_title(doc, page)
        add_mobile_gallery(doc, page)
        add_feature_table(doc, page.get("features", []))
        add_values(doc, page.get("values", []))
    else:
        add_standard_page(doc, page)


def configure_document(doc: Document):
    sec = doc.sections[0]
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.7)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.8)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor.from_string(TEXT)
    normal.paragraph_format.space_after = Pt(5)

    header = sec.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_text(header, "在途商旅 · 产品功能手册", 8, False, MUTED)

    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(footer, "在途商旅", 8, False, MUTED)


def generate(data_path: Path, output_path: Path):
    data = json.loads(data_path.read_text(encoding="utf-8"))
    doc = Document()
    configure_document(doc)

    # Cover
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(100)
    add_text(p, data.get("title", "产品功能手册"), 28, True, TEXT)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(p, data.get("subtitle", ""), 13, False, MUTED)
    doc.add_page_break()

    for s_idx, section in enumerate(data.get("sections", [])):
        # Section divider
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(80)
        add_text(p, section.get("code", ""), 12, True, BLUE)
        p = doc.add_paragraph()
        add_text(p, section.get("title", ""), 24, True, TEXT)
        doc.add_page_break()

        for p_idx, page in enumerate(section.get("pages", [])):
            add_page(doc, page)
            if not (s_idx == len(data["sections"]) - 1 and p_idx == len(section.get("pages", [])) - 1):
                doc.add_page_break()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default=str(ROOT / "data" / "sample_manual.json"))
    parser.add_argument("--output", default=str(ROOT / "output" / "zaitu_product_manual_demo.docx"))
    args = parser.parse_args()
    generate(Path(args.data), Path(args.output))
    print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
