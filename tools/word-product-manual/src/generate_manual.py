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
CONTENT_W = 16.2

STYLES = {
    "clean_blue": {"label": "A｜专业蓝白·证据型", "primary": "2F80ED", "soft": "EAF3FF", "paper": "F7FAFE", "text": "243447", "muted": "667085", "border": "D9E5F2"},
    "teal_grid": {"label": "B｜青绿·双栏型", "primary": "147D7A", "soft": "EAF7F6", "paper": "F6FBFB", "text": "20363A", "muted": "66797C", "border": "D7E9E8"},
    "indigo_report": {"label": "C｜靛蓝·报告型", "primary": "4B4E9B", "soft": "F0F0FB", "paper": "F8F8FC", "text": "292B3D", "muted": "6D7085", "border": "DDDDF0"},
    "amber_standard": {"label": "D｜金棕·杂志型", "primary": "A87520", "soft": "FBF4E7", "paper": "FCFAF5", "text": "342D24", "muted": "776F63", "border": "E9DFC9"},
    "green_governance": {"label": "E｜治理绿白·卡片型", "primary": "3C7D3B", "soft": "EEF7ED", "paper": "F8FCF7", "text": "263529", "muted": "6B786C", "border": "D9E8D7"},
}


def T(style: str):
    if style not in STYLES:
        raise ValueError(f"unknown style: {style}")
    return STYLES[style]


def shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = tcpr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tcpr.append(shd)
    shd.set(qn("w:fill"), fill)


def border(cell, color, size="6"):
    tcpr = cell._tc.get_or_add_tcPr()
    borders = tcpr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tcpr.append(borders)
    for edge in ("top", "left", "bottom", "right"):
        tag = qn("w:" + edge)
        node = borders.find(tag)
        if node is None:
            node = OxmlElement("w:" + edge)
            borders.append(node)
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:color"), color)


def margins(cell, top=100, side=120, bottom=100):
    tcpr = cell._tc.get_or_add_tcPr()
    mar = tcpr.first_child_found_in("w:tcMar")
    if mar is None:
        mar = OxmlElement("w:tcMar")
        tcpr.append(mar)
    for name, value in (("top", top), ("start", side), ("bottom", bottom), ("end", side)):
        node = mar.find(qn("w:" + name))
        if node is None:
            node = OxmlElement("w:" + name)
            mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def txt(p, value, size=10, bold=False, color="243447"):
    r = p.add_run(value)
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(size)
    r.font.color.rgb = RGBColor.from_string(color)
    r.bold = bold
    return r


def widths(table, values):
    table.autofit = False
    for i, v in enumerate(values):
        if i < len(table.columns):
            table.columns[i].width = Cm(v)


def title(doc, page, t, style):
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(1)
    txt(p, page.get("code", ""), 8.5, True, t["primary"])
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(3)
    txt(p, page.get("title", ""), 21, True, t["text"])
    p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(8)
    txt(p, page.get("summary", ""), 10, False, t["muted"])


def label(doc, value, t, prefix=""):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(5); p.paragraph_format.space_after = Pt(4)
    txt(p, prefix + value, 11.5, True, t["primary"])


def image_cell(cell, path, t, width=None, height=None, label_text="截图占位"):
    shade(cell, t["paper"]); border(cell, t["border"]); margins(cell, 150, 100, 150)
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if path and Path(path).exists():
        kwargs = {"width": Cm(width)} if width else {"height": Cm(height)}
        p.add_run().add_picture(path, **kwargs)
    else:
        txt(p, label_text, 9, False, t["muted"])


def feature_table(doc, features, t, style, compact=False):
    label(doc, "功能说明", t, "▌ " if style == "teal_grid" else "")
    table = doc.add_table(rows=1, cols=2); widths(table, [4.0, 12.2])
    for i, h in enumerate(("功能", "说明")):
        c = table.rows[0].cells[i]; shade(c, t["primary"]); margins(c, 80, 110, 80); txt(c.paragraphs[0], h, 9, True, "FFFFFF")
    for f in features:
        cs = table.add_row().cells
        for c in cs: border(c, t["border"]); margins(c, 70, 100, 70); c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        shade(cs[0], t["soft"]); txt(cs[0].paragraphs[0], f.get("name", ""), 9.2, True, t["primary"])
        txt(cs[1].paragraphs[0], f.get("description", ""), 8.9 if compact else 9.2, False, t["text"])
    return table


def value_strip(doc, values, t):
    if not values: return
    table = doc.add_table(rows=1, cols=len(values)); widths(table, [CONTENT_W / len(values)] * len(values))
    for i, v in enumerate(values):
        c = table.cell(0, i); shade(c, t["soft"]); border(c, t["soft"]); margins(c, 90, 90, 90)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, v, 8.7, True, t["primary"])


def cards(doc, features, t, cols=2):
    rows = (len(features) + cols - 1) // cols
    table = doc.add_table(rows=rows, cols=cols); widths(table, [CONTENT_W / cols] * cols)
    for i, f in enumerate(features):
        c = table.cell(i // cols, i % cols); shade(c, t["paper"]); border(c, t["border"]); margins(c, 130, 130, 130)
        txt(c.paragraphs[0], f.get("name", "能力"), 9.5, True, t["primary"])
        p = c.add_paragraph(); txt(p, f.get("description", ""), 8.8, False, t["text"])
    return table


def standard_A(doc, page, t):
    # A: screenshot is the hero evidence, followed by a conventional tender-friendly evidence table.
    table = doc.add_table(rows=1, cols=1); widths(table, [CONTENT_W])
    image_cell(table.cell(0, 0), page.get("screenshot"), t, width=15.25, label_text="PC 产品界面截图")
    label(doc, "功能说明", t)
    feature_table(doc, page.get("features", []), t, "clean_blue", compact=True)
    label(doc, "应用价值", t)
    value_strip(doc, page.get("values", []), t)


def standard_B(doc, page, t):
    # B: management evidence + capability narrative in a strong 60/40 split.
    table = doc.add_table(rows=1, cols=2); widths(table, [10.5, 5.7])
    image_cell(table.cell(0, 0), page.get("screenshot"), t, width=9.85, label_text="PC 管理端截图")
    c = table.cell(0, 1); shade(c, t["soft"]); border(c, t["border"]); margins(c, 160, 130, 160)
    txt(c.paragraphs[0], "核心能力", 11.5, True, t["primary"])
    for i, f in enumerate(page.get("features", [])[:5], 1):
        p = c.add_paragraph(); txt(p, f"0{i}  {f.get('name', '')}", 9.2, True, t["text"])
        p = c.add_paragraph(); txt(p, f.get("description", ""), 8.3, False, t["muted"])
    label(doc, "应用价值", t, "▌ ")
    value_strip(doc, page.get("values", []), t)


def standard_C(doc, page, t):
    # C: report layout. Function matrix precedes visual evidence, reversing the normal reading order.
    table = doc.add_table(rows=1, cols=3); widths(table, [5.4, 5.4, 5.4])
    feats = page.get("features", [])
    for i in range(3):
        f = feats[i] if i < len(feats) else {"name": "", "description": ""}
        c = table.cell(0, i); shade(c, t["soft"]); border(c, t["border"]); margins(c, 120, 110, 120)
        txt(c.paragraphs[0], f.get("name", ""), 9.3, True, t["primary"])
        p = c.add_paragraph(); txt(p, f.get("description", ""), 8.3, False, t["text"])
    label(doc, "产品界面证据", t, "01  ")
    table2 = doc.add_table(rows=1, cols=1); widths(table2, [CONTENT_W])
    image_cell(table2.cell(0, 0), page.get("screenshot"), t, width=14.7, label_text="PC 产品界面截图")
    if len(feats) > 3:
        feature_table(doc, feats[3:], t, "indigo_report", compact=True)
    label(doc, "管理价值", t, "02  ")
    value_strip(doc, page.get("values", []), t)


def standard_D(doc, page, t):
    # D: editorial/magazine rhythm. Large left evidence, numbered right notes, then a quiet value footer.
    table = doc.add_table(rows=1, cols=2); widths(table, [10.9, 5.3])
    image_cell(table.cell(0, 0), page.get("screenshot"), t, width=10.15, label_text="产品界面截图")
    c = table.cell(0, 1); shade(c, "FFFFFF"); border(c, t["border"]); margins(c, 100, 110, 100)
    for i, f in enumerate(page.get("features", [])[:5], 1):
        p = c.paragraphs[0] if i == 1 else c.add_paragraph()
        txt(p, f"{i:02d}", 16, True, t["primary"])
        txt(p, f"  {f.get('name', '')}", 9.4, True, t["text"])
        p = c.add_paragraph(); txt(p, f.get("description", ""), 8.4, False, t["muted"])
    label(doc, "应用价值", t, "01  ")
    value_strip(doc, page.get("values", []), t)


def standard_E(doc, page, t):
    # E: governance card layout. The screenshot is evidence below a six-capability matrix, not the page hero.
    feats = page.get("features", [])
    cards(doc, feats[:6], t, cols=3)
    label(doc, "界面证据", t)
    table = doc.add_table(rows=1, cols=1); widths(table, [CONTENT_W])
    image_cell(table.cell(0, 0), page.get("screenshot"), t, width=14.65, label_text="产品界面截图")
    label(doc, "管理价值", t)
    value_strip(doc, page.get("values", []), t)


def mobile_page(doc, page, t, style):
    title(doc, page, t, style)
    shots = page.get("screenshots", [])[:3]; n = len(shots) or 3
    table = doc.add_table(rows=1, cols=n); widths(table, [CONTENT_W / n] * n)
    labels = page.get("mobile_labels", [])
    for i in range(n):
        c = table.cell(0, i); image_cell(c, shots[i] if i < len(shots) else None, t, height=9.2, label_text="手机端截图")
        p = c.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, labels[i] if i < len(labels) else "移动端", 8.5, True, t["primary"])
    feature_table(doc, page.get("features", []), t, style, compact=True)
    value_strip(doc, page.get("values", []), t)


def pc_mobile_page(doc, page, t, style):
    title(doc, page, t, style)
    table = doc.add_table(rows=1, cols=2); widths(table, [11.0, 5.2])
    image_cell(table.cell(0, 0), page.get("pc_screenshot") or page.get("screenshot"), t, width=10.2, label_text="管理端截图")
    image_cell(table.cell(0, 1), page.get("mobile_screenshot"), t, height=8.6, label_text="员工端截图")
    feature_table(doc, page.get("features", []), t, style, compact=True)
    value_strip(doc, page.get("values", []), t)


def process_page(doc, page, t, style):
    title(doc, page, t, style)
    steps = page.get("steps", [])[:5]; n = len(steps) or 1
    table = doc.add_table(rows=1, cols=n); widths(table, [CONTENT_W / n] * n)
    for i, s in enumerate(steps):
        c = table.cell(0, i); shade(c, t["soft"]); border(c, t["border"]); margins(c, 130, 90, 130)
        p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, f"{i+1:02d}", 17, True, t["primary"])
        p = c.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, s.get("title", "步骤"), 9.2, True, t["text"])
        p = c.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, s.get("description", ""), 8.2, False, t["muted"])
    if page.get("screenshot"):
        table2 = doc.add_table(rows=1, cols=1); widths(table2, [CONTENT_W]); image_cell(table2.cell(0, 0), page["screenshot"], t, width=14.4, label_text="流程界面截图")
    value_strip(doc, page.get("values", []), t)


def dashboard_page(doc, page, t, style):
    title(doc, page, t, style)
    table = doc.add_table(rows=1, cols=1); widths(table, [CONTENT_W]); image_cell(table.cell(0, 0), page.get("screenshot"), t, width=14.8, label_text="BI / 数据看板截图")
    metrics = page.get("metrics", [])[:4]
    if metrics:
        table = doc.add_table(rows=1, cols=len(metrics)); widths(table, [CONTENT_W / len(metrics)] * len(metrics))
        for i, m in enumerate(metrics):
            c = table.cell(0, i); shade(c, t["soft"]); border(c, t["border"]); margins(c, 120, 100, 120)
            p = c.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, m.get("value", "-"), 15, True, t["primary"])
            p = c.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, m.get("name", "指标"), 8.3, False, t["muted"])
    feature_table(doc, page.get("features", []), t, style, compact=True)


def overview_page(doc, page, t, style):
    title(doc, page, t, style)
    cards(doc, page.get("items", []), t, cols=2)


def configure(doc, style):
    t = T(style); sec = doc.sections[0]
    sec.top_margin = Cm(1.8); sec.bottom_margin = Cm(1.7); sec.left_margin = Cm(2); sec.right_margin = Cm(2)
    sec.header_distance = Cm(.8); sec.footer_distance = Cm(.8)
    normal = doc.styles["Normal"]; normal.font.name = "Microsoft YaHei"; normal.font.size = Pt(10); normal.font.color.rgb = RGBColor.from_string(t["text"]); normal.paragraph_format.space_after = Pt(4)
    p = sec.header.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.RIGHT; txt(p, f"在途商旅 · {t['label']}", 7.5, False, t["muted"])
    p = sec.footer.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, "在途商旅", 7.5, False, t["muted"])


def cover(doc, data, t, style):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(105); txt(p, data.get("title", "产品功能手册"), 27, True, t["text"])
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; txt(p, data.get("subtitle", ""), 12, False, t["muted"])
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_before = Pt(18); txt(p, t["label"], 9, True, t["primary"])
    doc.add_page_break()


def section_page(doc, section, t):
    p = doc.add_paragraph(); p.paragraph_format.space_before = Pt(82); txt(p, section.get("code", ""), 12, True, t["primary"])
    p = doc.add_paragraph(); txt(p, section.get("title", ""), 23, True, t["text"])
    if section.get("summary"):
        p = doc.add_paragraph(); txt(p, section["summary"], 10, False, t["muted"])
    doc.add_page_break()


def add_page(doc, page, style):
    t = T(style); kind = page.get("type", "standard")
    if kind == "mobile_gallery": mobile_page(doc, page, t, style)
    elif kind == "pc_mobile": pc_mobile_page(doc, page, t, style)
    elif kind in ("process", "steps"): process_page(doc, page, t, style)
    elif kind == "dashboard": dashboard_page(doc, page, t, style)
    elif kind == "overview": overview_page(doc, page, t, style)
    else:
        title(doc, page, t, style)
        {"clean_blue": standard_A, "teal_grid": standard_B, "indigo_report": standard_C, "amber_standard": standard_D, "green_governance": standard_E}[style](doc, page, t)


def generate(data_path: Path, output_path: Path, style: str):
    data = json.loads(data_path.read_text(encoding="utf-8")); doc = Document(); configure(doc, style); t = T(style); cover(doc, data, t, style)
    sections = data.get("sections", [])
    for si, section in enumerate(sections):
        section_page(doc, section, t)
        pages = section.get("pages", [])
        for pi, page in enumerate(pages):
            add_page(doc, page, style)
            if not (si == len(sections)-1 and pi == len(pages)-1): doc.add_page_break()
    output_path.parent.mkdir(parents=True, exist_ok=True); doc.save(output_path)


def main():
    parser = argparse.ArgumentParser(); parser.add_argument("--data", default=str(ROOT / "data" / "sample_manual.json")); parser.add_argument("--output", default=str(ROOT / "output" / "zaitu_product_manual_demo.docx")); parser.add_argument("--style", choices=sorted(STYLES), default="clean_blue"); parser.add_argument("--all-styles", action="store_true")
    a = parser.parse_args()
    if a.all_styles:
        for style in STYLES:
            out = Path(a.output).with_name(f"zaitu_product_manual_{style}.docx"); generate(Path(a.data), out, style); print(f"Generated: {out}")
    else:
        generate(Path(a.data), Path(a.output), a.style); print(f"Generated: {a.output}")


if __name__ == "__main__":
    main()
