from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]

STYLES: dict[str, dict[str, str]] = {
    "clean_blue": {
        "label": "A｜专业蓝白",
        "primary": "2F80ED", "soft": "EAF3FF", "soft2": "F7FAFE",
        "text": "243447", "muted": "667085", "border": "D9E5F2", "accent": "6FA8DC",
    },
    "teal_grid": {
        "label": "B｜现代青绿",
        "primary": "147D7A", "soft": "EAF7F6", "soft2": "F6FBFB",
        "text": "20363A", "muted": "66797C", "border": "D7E9E8", "accent": "57A8A3",
    },
    "indigo_report": {
        "label": "C｜稳重靛蓝",
        "primary": "4B4E9B", "soft": "F0F0FB", "soft2": "F8F8FC",
        "text": "292B3D", "muted": "6D7085", "border": "DDDDF0", "accent": "7D80C6",
    },
    "amber_standard": {
        "label": "D｜商务金棕",
        "primary": "A87520", "soft": "FBF4E7", "soft2": "FCFAF5",
        "text": "342D24", "muted": "776F63", "border": "E9DFC9", "accent": "C49A55",
    },
    "green_governance": {
        "label": "E｜治理绿白",
        "primary": "3C7D3B", "soft": "EEF7ED", "soft2": "F8FCF7",
        "text": "263529", "muted": "6B786C", "border": "D9E8D7", "accent": "76A874",
    },
}


def theme(style: str) -> dict[str, str]:
    if style not in STYLES:
        raise ValueError(f"Unknown style: {style}. Choose from: {', '.join(STYLES)}")
    return STYLES[style]


def shade(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def cell_border(cell, color: str, size: str = "6"):
    tc_pr = cell._tc.get_or_add_tcPr()
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
    tc_pr = cell._tc.get_or_add_tcPr()
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


def add_text(paragraph, text, size=10.5, bold=False, color="243447"):
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    return run


def set_cell_widths(table, widths_cm: list[float]):
    table.autofit = False
    for i, width in enumerate(widths_cm):
        if i < len(table.columns):
            table.columns[i].width = Cm(width)


def add_section_label(doc, text: str, t: dict[str, str], style: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(4)
    if style == "clean_blue":
        add_text(p, text, 12.5, True, t["primary"])
    elif style == "teal_grid":
        add_text(p, "▌ " + text, 12.5, True, t["primary"])
    elif style == "indigo_report":
        add_text(p, text.upper(), 11, True, t["primary"])
    elif style == "amber_standard":
        add_text(p, "01  " + text, 11.5, True, t["primary"])
    else:
        add_text(p, text, 12.5, True, t["primary"])


def add_page_title(doc, page: dict[str, Any], t: dict[str, str], style: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    add_text(p, page.get("code", ""), 8.5, True, t["primary"])

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    title_size = 21 if style in ("clean_blue", "teal_grid") else 20
    add_text(p, page.get("title", ""), title_size, True, t["text"])

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(9)
    add_text(p, page.get("summary", ""), 10.2, False, t["muted"])


def add_screenshot(doc, path: str | None, label: str, t: dict[str, str], width_cm: float = 15.4):
    table = doc.add_table(rows=1, cols=1)
    table.autofit = False
    table.columns[0].width = Cm(width_cm + 0.8)
    cell = table.cell(0, 0)
    cell.width = Cm(width_cm + 0.8)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    shade(cell, t["soft2"])
    cell_border(cell, t["border"])
    set_cell_margins(cell, 260, 180, 260, 180)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if path and Path(path).exists():
        p.add_run().add_picture(path, width=Cm(width_cm))
    else:
        add_text(p, f"截图占位：{label}\n将实际 PC / 手机界面截图放入 assets/screenshots 后即可生成。", 9.5, False, t["muted"])
    doc.add_paragraph().paragraph_format.space_after = Pt(1)


def add_feature_table(doc, features: list[dict[str, str]], t: dict[str, str], style: str):
    add_section_label(doc, "功能说明", t, style)
    table = doc.add_table(rows=1, cols=2)
    set_cell_widths(table, [4.0, 12.2])
    hdr = table.rows[0].cells
    for i, text in enumerate(("功能", "说明")):
        shade(hdr[i], t["primary"])
        set_cell_margins(hdr[i], 100, 120, 100, 120)
        add_text(hdr[i].paragraphs[0], text, 9.5, True, "FFFFFF")
    for feature in features:
        cells = table.add_row().cells
        for c in cells:
            cell_border(c, t["border"])
            set_cell_margins(c, 85, 110, 85, 110)
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(cells[0], t["soft"])
        add_text(cells[0].paragraphs[0], feature.get("name", ""), 9.5, True, t["primary"])
        add_text(cells[1].paragraphs[0], feature.get("description", ""), 9.1, False, t["text"])


def add_values(doc, values: list[str], t: dict[str, str], style: str):
    if not values:
        return
    add_section_label(doc, "应用价值", t, style)
    table = doc.add_table(rows=1, cols=len(values))
    set_cell_widths(table, [16.2 / max(len(values), 1)] * len(values))
    for i, value in enumerate(values):
        cell = table.cell(0, i)
        shade(cell, t["soft"])
        cell_border(cell, t["soft"])
        set_cell_margins(cell, 105, 120, 105, 120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, "● " + value, 8.8, True, t["primary"])


def add_mobile_gallery(doc, page: dict[str, Any], t: dict[str, str], style: str):
    screenshots = page.get("screenshots", [])
    count = min(3, len(screenshots)) or 3
    table = doc.add_table(rows=1, cols=count)
    set_cell_widths(table, [16.2 / count] * count)
    for i in range(count):
        path = screenshots[i] if i < len(screenshots) else None
        cell = table.cell(0, i)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(cell, t["soft2"])
        cell_border(cell, t["border"])
        set_cell_margins(cell, 130, 90, 130, 90)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if path and Path(path).exists():
            p.add_run().add_picture(path, height=Cm(9.5))
        else:
            add_text(p, "手机截图\n占位", 8.5, False, t["muted"])
        label = page.get("mobile_labels", [])[i] if i < len(page.get("mobile_labels", [])) else "移动端"
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p2, label, 8.5, True, t["primary"])


def add_pc_mobile(doc, page: dict[str, Any], t: dict[str, str], style: str):
    table = doc.add_table(rows=1, cols=2)
    set_cell_widths(table, [10.8, 5.0])
    pc = page.get("pc_screenshot") or page.get("screenshot")
    mobile = page.get("mobile_screenshot")
    for i, (path, label) in enumerate(((pc, "管理端"), (mobile, "员工端"))):
        cell = table.cell(0, i)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(cell, t["soft2"])
        cell_border(cell, t["border"])
        set_cell_margins(cell, 140, 100, 140, 100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if path and Path(path).exists():
            if i == 0:
                p.add_run().add_picture(path, width=Cm(10.0))
            else:
                p.add_run().add_picture(path, height=Cm(8.8))
        else:
            add_text(p, label + "截图占位", 9, False, t["muted"])
        lp = cell.add_paragraph()
        lp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(lp, label, 9, True, t["primary"])


def add_process(doc, page: dict[str, Any], t: dict[str, str], style: str):
    add_section_label(doc, "业务流程", t, style)
    steps = page.get("steps", [])
    table = doc.add_table(rows=1, cols=max(1, min(5, len(steps))))
    count = len(table.columns)
    set_cell_widths(table, [16.2 / count] * count)
    for i in range(count):
        step = steps[i]
        cell = table.cell(0, i)
        shade(cell, t["soft2"])
        cell_border(cell, t["border"])
        set_cell_margins(cell, 150, 100, 150, 100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, f"{i + 1:02d}", 17, True, t["primary"])
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p2, step.get("title", "步骤"), 9.5, True, t["text"])
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p3, step.get("description", ""), 8.5, False, t["muted"])


def add_dashboard(doc, page: dict[str, Any], t: dict[str, str], style: str):
    add_screenshot(doc, page.get("screenshot"), page.get("title", "数据看板"), t, 15.0)
    metrics = page.get("metrics", [])
    if not metrics:
        return
    table = doc.add_table(rows=1, cols=min(4, len(metrics)))
    set_cell_widths(table, [16.2 / len(table.columns)] * len(table.columns))
    for i, metric in enumerate(metrics[:4]):
        cell = table.cell(0, i)
        shade(cell, t["soft"])
        cell_border(cell, t["border"])
        set_cell_margins(cell, 130, 110, 130, 110)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p, metric.get("value", "-"), 16, True, t["primary"])
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        add_text(p2, metric.get("name", "指标"), 8.5, False, t["muted"])


def add_overview(doc, page: dict[str, Any], t: dict[str, str], style: str):
    add_section_label(doc, page.get("title", "能力总览"), t, style)
    cards = page.get("items", [])
    table = doc.add_table(rows=(len(cards) + 1) // 2, cols=2)
    set_cell_widths(table, [8.0, 8.0])
    for idx, item in enumerate(cards):
        cell = table.cell(idx // 2, idx % 2)
        shade(cell, t["soft2"])
        cell_border(cell, t["border"])
        set_cell_margins(cell, 160, 150, 160, 150)
        p = cell.paragraphs[0]
        add_text(p, item.get("title", "能力模块"), 11, True, t["primary"])
        p2 = cell.add_paragraph()
        add_text(p2, item.get("description", ""), 9, False, t["text"])


def add_standard_page(doc, page: dict[str, Any], t: dict[str, str], style: str):
    add_page_title(doc, page, t, style)

    # Five styles use the same information architecture but different spatial rhythm.
    if style == "clean_blue":
        add_screenshot(doc, page.get("screenshot"), page.get("title", "产品界面"), t, 15.4)
        add_feature_table(doc, page.get("features", []), t, style)
        add_values(doc, page.get("values", []), t, style)
    elif style == "teal_grid":
        table = doc.add_table(rows=1, cols=2)
        set_cell_widths(table, [10.6, 5.6])
        left = table.cell(0, 0)
        right = table.cell(0, 1)
        for c in (left, right):
            shade(c, t["soft2"])
            cell_border(c, t["border"])
            set_cell_margins(c, 140, 120, 140, 120)
        p = left.paragraphs[0]
        if page.get("screenshot") and Path(page["screenshot"]).exists():
            p.add_run().add_picture(page["screenshot"], width=Cm(10.0))
        else:
            add_text(p, "PC截图占位", 9, False, t["muted"])
        rp = right.paragraphs[0]
        add_text(rp, "核心能力", 11, True, t["primary"])
        for f in page.get("features", [])[:5]:
            p = right.add_paragraph(style=None)
            add_text(p, "• " + f.get("name", ""), 9, True, t["text"])
            add_text(p, "\n" + f.get("description", ""), 8.5, False, t["muted"])
        doc.add_paragraph().paragraph_format.space_after = Pt(1)
        add_values(doc, page.get("values", []), t, style)
    elif style == "indigo_report":
        add_feature_table(doc, page.get("features", []), t, style)
        add_screenshot(doc, page.get("screenshot"), page.get("title", "产品界面"), t, 14.8)
        add_values(doc, page.get("values", []), t, style)
    elif style == "amber_standard":
        # Report-style: narrow title block, screenshot, then two-column value summary.
        add_screenshot(doc, page.get("screenshot"), page.get("title", "产品界面"), t, 15.0)
        add_feature_table(doc, page.get("features", []), t, style)
        vals = page.get("values", [])
        if vals:
            table = doc.add_table(rows=1, cols=2)
            set_cell_widths(table, [8.0, 8.0])
            for i in range(2):
                cell = table.cell(0, i)
                shade(cell, t["soft"])
                cell_border(cell, t["border"])
                set_cell_margins(cell, 120, 130, 120, 130)
                text = vals[i] if i < len(vals) else ""
                add_text(cell.paragraphs[0], text, 9.2, True, t["primary"])
    else:
        # Governance style: capability cards first, screenshot as evidence.
        features = page.get("features", [])[:6]
        table = doc.add_table(rows=(len(features) + 2) // 3, cols=3)
        set_cell_widths(table, [5.3, 5.3, 5.3])
        for idx, f in enumerate(features):
            cell = table.cell(idx // 3, idx % 3)
            shade(cell, t["soft"])
            cell_border(cell, t["border"])
            set_cell_margins(cell, 120, 110, 120, 110)
            add_text(cell.paragraphs[0], f.get("name", "能力"), 9.3, True, t["primary"])
            p = cell.add_paragraph()
            add_text(p, f.get("description", ""), 8.4, False, t["text"])
        doc.add_paragraph().paragraph_format.space_after = Pt(1)
        add_screenshot(doc, page.get("screenshot"), page.get("title", "产品界面"), t, 14.8)
        add_values(doc, page.get("values", []), t, style)


def add_page(doc, page: dict[str, Any], style: str):
    t = theme(style)
    kind = page.get("type", "standard")
    if kind == "mobile_gallery":
        add_page_title(doc, page, t, style)
        add_mobile_gallery(doc, page, t, style)
        add_feature_table(doc, page.get("features", []), t, style)
        add_values(doc, page.get("values", []), t, style)
    elif kind == "pc_mobile":
        add_page_title(doc, page, t, style)
        add_pc_mobile(doc, page, t, style)
        add_feature_table(doc, page.get("features", []), t, style)
        add_values(doc, page.get("values", []), t, style)
    elif kind == "process" or kind == "steps":
        add_page_title(doc, page, t, style)
        add_process(doc, page, t, style)
        if page.get("screenshot"):
            add_screenshot(doc, page.get("screenshot"), page.get("title", "流程界面"), t, 14.5)
        add_values(doc, page.get("values", []), t, style)
    elif kind == "dashboard":
        add_page_title(doc, page, t, style)
        add_dashboard(doc, page, t, style)
        add_feature_table(doc, page.get("features", []), t, style)
    elif kind == "overview":
        add_page_title(doc, page, t, style)
        add_overview(doc, page, t, style)
    else:
        add_standard_page(doc, page, t, style)


def configure_document(doc: Document, style: str):
    t = theme(style)
    sec = doc.sections[0]
    sec.top_margin = Cm(1.8)
    sec.bottom_margin = Cm(1.7)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.8)

    normal = doc.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(10.2)
    normal.font.color.rgb = RGBColor.from_string(t["text"])
    normal.paragraph_format.space_after = Pt(4)

    header = sec.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_text(header, f"在途商旅 · {t['label']}", 7.8, False, t["muted"])

    footer = sec.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(footer, "在途商旅", 7.8, False, t["muted"])


def add_cover(doc: Document, data: dict[str, Any], t: dict[str, str], style: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(95)
    add_text(p, data.get("title", "产品功能手册"), 27, True, t["text"])
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(p, data.get("subtitle", ""), 12.5, False, t["muted"])
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(18)
    add_text(p, t["label"], 9.5, True, t["primary"])
    doc.add_page_break()


def generate(data_path: Path, output_path: Path, style: str):
    data = json.loads(data_path.read_text(encoding="utf-8"))
    doc = Document()
    configure_document(doc, style)
    t = theme(style)
    add_cover(doc, data, t, style)

    sections = data.get("sections", [])
    for s_idx, section in enumerate(sections):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(80)
        add_text(p, section.get("code", ""), 11.5, True, t["primary"])
        p = doc.add_paragraph()
        add_text(p, section.get("title", ""), 23, True, t["text"])
        if section.get("summary"):
            p = doc.add_paragraph()
            add_text(p, section["summary"], 10, False, t["muted"])
        doc.add_page_break()

        pages = section.get("pages", [])
        for p_idx, page in enumerate(pages):
            add_page(doc, page, style)
            if not (s_idx == len(sections) - 1 and p_idx == len(pages) - 1):
                doc.add_page_break()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Generate the 在途商旅 Word product manual")
    parser.add_argument("--data", default=str(ROOT / "data" / "sample_manual.json"))
    parser.add_argument("--output", default=str(ROOT / "output" / "zaitu_product_manual_demo.docx"))
    parser.add_argument("--style", choices=sorted(STYLES), default="clean_blue")
    parser.add_argument("--all-styles", action="store_true", help="Generate one demo DOCX for each of the five visual styles")
    args = parser.parse_args()

    if args.all_styles:
        for style in STYLES:
            out = Path(args.output).with_name(f"zaitu_product_manual_{style}.docx")
            generate(Path(args.data), out, style)
            print(f"Generated: {out}")
    else:
        generate(Path(args.data), Path(args.output), args.style)
        print(f"Generated: {args.output}")


if __name__ == "__main__":
    main()
