"""
محرّك التحويل — PDF والصور إلى Word / Excel / PowerPoint
=========================================================
يعتمد على مكتبات محلية بالكامل (بدون إنترنت ولا خدمات خارجية):
  • pdf2docx      : PDF -> DOCX مع الحفاظ على التنسيق
  • pdfplumber    : استخراج الجداول والنصوص من PDF
  • python-docx   : بناء ملفات Word
  • openpyxl      : بناء ملفات Excel
  • python-pptx   : بناء ملفات PowerPoint
  • PyMuPDF       : تصيير صفحات PDF إلى صور
"""

from __future__ import annotations

import logging
import os
import re
from pathlib import Path
from typing import Iterable, Sequence

import pdfplumber
from PIL import Image as PILImage
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.util import Inches as PInches

import ocr_engine

for _noisy in ("pdf2docx", "pdfminer", "fontTools", "PIL"):
    _lg = logging.getLogger(_noisy)
    _lg.setLevel(logging.CRITICAL)
    _lg.propagate = False
    _lg.disabled = True

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif"}
PDF_EXTS = {".pdf"}
MAX_PDF_PAGES = 300

# ألوان الهوية البصرية للملفات الناتجة
ACCENT = "4F46E5"
HEADER_FILL = PatternFill("solid", fgColor="EEF2FF")
HEADER_FONT = Font(bold=True, color="1E1B4B", size=11)
THIN = Side(style="thin", color="D8DCF0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


# --------------------------------------------------------------------------
# أدوات مساعدة
# --------------------------------------------------------------------------
def is_pdf(path: Path) -> bool:
    return path.suffix.lower() in PDF_EXTS


def is_image(path: Path) -> bool:
    return path.suffix.lower() in IMAGE_EXTS


def _safe_sheet_name(raw: str, used: set[str]) -> str:
    """اسم ورقة صالح لـ Excel (≤31 حرفًا، بلا رموز ممنوعة، غير مكرر)."""
    name = re.sub(r"[\\/*?:\[\]]", " ", raw).strip() or "ورقة"
    name = name[:31]
    base, i = name, 2
    while name.lower() in used:
        suffix = f" ({i})"
        name = base[: 31 - len(suffix)] + suffix
        i += 1
    used.add(name.lower())
    return name


def _autosize_columns(ws) -> None:
    widths: dict[int, int] = {}
    for row in ws.iter_rows(values_only=True):
        for idx, value in enumerate(row, start=1):
            if value is None:
                continue
            text = str(value)
            # النص العربي يبدو أعرض بصريًا
            length = max(len(line) for line in text.splitlines()) + text.count(" ") * 0.4
            widths[idx] = max(widths.get(idx, 8), min(int(length) + 3, 60))
    for idx, width in widths.items():
        ws.column_dimensions[get_column_letter(idx)].width = width


def _fit_size(px_w: int, px_h: int, max_w: float, max_h: float) -> tuple[float, float]:
    """يرجّع أبعادًا تحفظ نسبة الصورة داخل حدّ أقصى (إنش)."""
    if px_w <= 0 or px_h <= 0:
        return max_w, max_h
    ratio = min(max_w / px_w, max_h / px_h)
    return px_w * ratio, px_h * ratio


def _to_png(path: Path, out_dir: Path) -> Path:
    """يوحّد صيغ الصور إلى PNG (يدعم webp/bmp/tiff الشفافية)."""
    if path.suffix.lower() == ".png":
        return path
    target = out_dir / (path.stem + ".png")
    with PILImage.open(path) as img:
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGB")
        img.save(target, "PNG")
    return target


# --------------------------------------------------------------------------
# PDF  ->  Word
# --------------------------------------------------------------------------
def pdf_to_docx(src: Path, dst: Path) -> None:
    from pdf2docx import Converter

    cv = Converter(str(src))
    try:
        cv.convert(str(dst))
    finally:
        cv.close()


# --------------------------------------------------------------------------
# PDF  ->  Excel
# --------------------------------------------------------------------------
def _write_table(ws, table: list[list], start_row: int = 1) -> int:
    # إزالة الصفوف الفارغة في النهاية
    rows = [list(r) for r in table]
    while rows and not any((c or "").strip() if isinstance(c, str) else c for c in rows[-1]):
        rows.pop()
    row_idx = start_row
    for r, row in enumerate(rows):
        for c, cell in enumerate(row, start=1):
            value = "" if cell is None else str(cell).strip()
            target = ws.cell(row=row_idx, column=c, value=value)
            target.alignment = Alignment(vertical="center", wrap_text=True)
            target.border = BORDER
            if r == 0:
                target.font = HEADER_FONT
                target.fill = HEADER_FILL
        row_idx += 1
    return row_idx


def _write_text_fallback(ws, text: str, title: str) -> None:
    ws.cell(row=1, column=1, value=title).font = Font(bold=True, size=12, color="1E1B4B")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
    lines = [ln for ln in (text or "").splitlines()]
    row_idx = 3
    for line in lines:
        if not line.strip():
            continue
        parts = [p.strip() for p in re.split(r"\t|\s{3,}|\s*\|\s*", line) if p.strip()]
        for c, part in enumerate(parts or [line], start=1):
            cell = ws.cell(row=row_idx, column=c, value=part)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = BORDER
        row_idx += 1
    if row_idx == 3:
        ws.cell(
            row=3,
            column=1,
            value="لم يُعثر على نص قابل للاستخراج — يبدو الملف صورة ممسوحة (Scan).",
        ).font = Font(color="B45309", italic=True)


def pdf_to_xlsx(src: Path, dst: Path) -> dict:
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    tables_found = 0
    text_pages = 0

    with pdfplumber.open(str(src)) as pdf:
        pages = pdf.pages[:MAX_PDF_PAGES]
        for p_no, page in enumerate(pages, start=1):
            try:
                tables = page.extract_tables() or []
            except Exception:
                tables = []
            tables = [t for t in tables if t and any(any(c for c in row) for row in t)]

            if tables:
                for t_no, table in enumerate(tables, start=1):
                    ws = wb.create_sheet(_safe_sheet_name(f"ص{p_no} جدول{t_no}", used))
                    _write_table(ws, table)
                    _autosize_columns(ws)
                    tables_found += 1
            else:
                text = page.extract_text() or ""
                ws = wb.create_sheet(_safe_sheet_name(f"صفحة {p_no}", used))
                _write_text_fallback(ws, text, f"صفحة {p_no}")
                _autosize_columns(ws)
                text_pages += 1

    if not wb.sheetnames:
        ws = wb.create_sheet("النتيجة")
        ws["A1"] = "لم يتم العثور على أي محتوى في ملف PDF."
    wb.save(str(dst))
    return {"tables": tables_found, "text_pages": text_pages}


# --------------------------------------------------------------------------
# PDF  ->  PowerPoint
# --------------------------------------------------------------------------
def _new_presentation() -> Presentation:
    prs = Presentation()
    prs.slide_width = PInches(13.333)
    prs.slide_height = PInches(7.5)
    return prs


def _add_picture_slide(prs: Presentation, image_path: Path) -> None:
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    with PILImage.open(image_path) as img:
        w_px, h_px = img.size
    max_w, max_h = prs.slide_width, prs.slide_height
    ratio = min(max_w / w_px, max_h / h_px)
    w, h = int(w_px * ratio), int(h_px * ratio)
    slide.shapes.add_picture(
        str(image_path), int((max_w - w) / 2), int((max_h - h) / 2), width=w, height=h
    )


def pdf_to_pptx(src: Path, dst: Path, work_dir: Path) -> dict:
    import pymupdf

    prs = _new_presentation()
    doc = pymupdf.open(str(src))
    count = 0
    try:
        for idx in range(min(doc.page_count, MAX_PDF_PAGES)):
            page = doc.load_page(idx)
            pix = page.get_pixmap(dpi=150)
            img_path = work_dir / f"_p{idx + 1}.png"
            pix.save(str(img_path))
            _add_picture_slide(prs, img_path)
            img_path.unlink(missing_ok=True)
            count += 1
    finally:
        doc.close()
    prs.save(str(dst))
    return {"slides": count}


# --------------------------------------------------------------------------
# الصور  ->  Word
# --------------------------------------------------------------------------
def images_to_docx(images: Sequence[Path], dst: Path) -> dict:
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    max_w = (section.page_width - section.left_margin - section.right_margin) / 914400
    max_h = (section.page_height - section.top_margin - section.bottom_margin) / 914400 - 0.4

    for i, img_path in enumerate(images):
        with PILImage.open(img_path) as img:
            w_px, h_px = img.size
        w_in, h_in = _fit_size(w_px, h_px, max_w, max_h)
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.add_run().add_picture(str(img_path), width=Inches(w_in), height=Inches(h_in))
        if i < len(images) - 1:
            doc.add_page_break()

    doc.save(str(dst))
    return {"images": len(images)}


# --------------------------------------------------------------------------
# الصور  ->  Excel
# --------------------------------------------------------------------------
def images_to_xlsx(images: Sequence[Path], dst: Path, work_dir: Path, per_row: int = 2) -> dict:
    wb = Workbook()
    ws = wb.active
    ws.title = "الصور"
    ws.sheet_view.rightToLeft = True

    ws["A1"] = "أرشيف الصور"
    ws["A1"].font = Font(bold=True, size=14, color="1E1B4B")
    ws["A2"] = "كل صورة مُدرجة داخل خليّة — استخدمها كأرشيف مرئي أو مرجع سريع."
    ws["A2"].font = Font(italic=True, size=9, color="6B7280")

    cell_px = 200
    gap_rows = 12
    start_row = 4
    for i, img_path in enumerate(images):
        png = _to_png(img_path, work_dir)
        with PILImage.open(png) as img:
            w_px, h_px = img.size
        scale = cell_px / max(h_px, 1)
        w_px, h_px = int(w_px * scale), int(h_px * scale)

        col_block, row_block = divmod(i, per_row)
        anchor_col = 1 + col_block * 8
        anchor_row = start_row + row_block * gap_rows

        label = ws.cell(row=anchor_row, column=anchor_col, value=png.stem)
        label.font = Font(bold=True, size=10, color="374151")

        xl_img = XLImage(str(png))
        xl_img.width, xl_img.height = w_px, h_px
        ws.add_image(xl_img, f"{get_column_letter(anchor_col)}{anchor_row + 1}")

        for r in range(anchor_row + 1, anchor_row + 1 + gap_rows):
            ws.row_dimensions[r].height = max(h_px / gap_rows * 0.75, 14)
        ws.column_dimensions[get_column_letter(anchor_col)].width = 30

    wb.save(str(dst))
    return {"images": len(images)}


# --------------------------------------------------------------------------
# الصور  ->  PowerPoint
# --------------------------------------------------------------------------
def images_to_pptx(images: Sequence[Path], dst: Path, work_dir: Path) -> dict:
    prs = _new_presentation()
    for img_path in images:
        _add_picture_slide(prs, _to_png(img_path, work_dir))
    prs.save(str(dst))
    return {"slides": len(images)}


# --------------------------------------------------------------------------
# OCR  —  تحويل الصور وملفات PDF الممسوحة إلى نص حقيقي
# --------------------------------------------------------------------------
ARABIC_FONT = "Arial"

# ترتيب العناصر مطلوب بحسب معيار OOXML، وإلا تجاهلها Word
_AFTER_BIDI = (
    "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)
_AFTER_RTL = ("w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath")
_AFTER_SECT_BIDI = ("w:rtlGutter", "w:docGrid", "w:printerSettings", "w:sectPrChange")


def _insert_ordered(parent, element, successors: tuple[str, ...]) -> None:
    for tag in successors:
        found = parent.find(qn(tag))
        if found is not None:
            found.addprevious(element)
            return
    parent.append(element)


def _has_arabic(text: str) -> bool:
    return any("\u0600" <= ch <= "\u06FF" for ch in text or "")


def _rtl_paragraph(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    if pPr.find(qn("w:bidi")) is None:
        _insert_ordered(pPr, OxmlElement("w:bidi"), _AFTER_BIDI)


def _rtl_run(run, size: float, bold: bool = False, gray: bool = False) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = ARABIC_FONT
    if gray:
        run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:cs"), ARABIC_FONT)
    rFonts.set(qn("w:ascii"), ARABIC_FONT)
    rFonts.set(qn("w:hAnsi"), ARABIC_FONT)

    size_cs = OxmlElement("w:szCs")
    size_cs.set(qn("w:val"), str(int(size * 2)))
    if rPr.find(qn("w:szCs")) is None:
        _insert_ordered(rPr, size_cs, _AFTER_RTL)

    # الأرقام والمقاطع اللاتينية تبقى باتجاهها الطبيعي
    if _has_arabic(run.text) and rPr.find(qn("w:rtl")) is None:
        _insert_ordered(rPr, OxmlElement("w:rtl"), _AFTER_RTL)


def _add_arabic_paragraph(
    doc, text: str, size: float = 11, bold: bool = False, gray: bool = False, align: str | None = None
):
    paragraph = doc.add_paragraph()
    _rtl_paragraph(paragraph)
    if align == "center":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "left":
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0.5)
    paragraph.paragraph_format.line_spacing = 0.92
    run = paragraph.add_run(text)
    _rtl_run(run, size, bold=bold, gray=gray)
    return paragraph


def _set_cell_margins(table, top: int = 15, bottom: int = 15, left: int = 50, right: int = 50) -> None:
    margins = OxmlElement("w:tblCellMar")
    for tag, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        node = OxmlElement(f"w:{tag}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    _insert_ordered(
        table._tbl.tblPr, margins, ("w:tblLook", "w:tblCaption", "w:tblDescription")
    )


def _add_ocr_table(doc, grid, pt_per_px: float, usable_width_pt: float, fidelity: bool = True) -> None:
    """
    يبني جدول Word بقياسات الصفحة الأصلية:
    عرض كل عمود وارتفاع كل صف وحجم خط الخلية مأخوذة من الصورة نفسها.
    """
    rows = grid.rows
    n_cols = grid.n_cols
    table = doc.add_table(rows=len(rows), cols=n_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _insert_ordered(
        table._tbl.tblPr,
        OxmlElement("w:bidiVisual"),
        ("w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
         "w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription"),
    )
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    _insert_ordered(
        table._tbl.tblPr, layout, ("w:tblCellMar", "w:tblLook", "w:tblCaption", "w:tblDescription")
    )
    _set_cell_margins(table)

    # عرض الأعمدة من حدود الجدول الأصلية، مع تقليص متناسب إن تجاوز عرض الصفحة
    scale = 1.0
    widths: list[float] | None = None
    if fidelity and len(grid.col_bounds) == n_cols:
        raw = [max(10.0, (b[1] - b[0]) * pt_per_px) for b in grid.col_bounds]
        total = sum(raw)
        if total > 0 and usable_width_pt > 0:
            # نملأ عرض الصفحة كما في الأصل، بسقف يمنع التمدد المبالغ
            scale = min(1.35, usable_width_pt / total)
        widths = [w * scale for w in raw]
        for index, width in enumerate(widths):
            table.columns[index].width = Pt(width)

    header_rows = grid.header_rows if fidelity else 1

    for r, cells_text in enumerate(rows):
        row = table.rows[r]
        if fidelity and r < len(grid.row_bounds):
            # ارتفاع الصف من خطوط الجدول الأصلية — لا يُقاس بمعامل العرض
            height = (grid.row_bounds[r][1] - grid.row_bounds[r][0]) * pt_per_px
            row.height = Pt(max(12.0, height))
            row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        if r < header_rows:
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))

        for c in range(n_cols):
            cell = row.cells[c]
            text = cells_text[c] if c < len(cells_text) else ""
            if widths:
                cell.width = Pt(widths[c])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            paragraph = cell.paragraphs[0]
            _rtl_paragraph(paragraph)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0

            size = 8.5
            if fidelity and r < len(grid.cell_heights):
                heights = grid.cell_heights[r]
                if c < len(heights) and heights[c]:
                    size = max(6.5, min(13.0, heights[c] * pt_per_px * 0.88))

            if r < header_rows or (text and _is_numeric(text)):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.add_run(text)
            _rtl_run(run, size, bold=r < header_rows)


def _is_numeric(text: str) -> bool:
    allowed = set("0123456789٠١٢٣٤٥٦٧٨٩.,،/- ")
    return bool(text) and all(ch in allowed for ch in text)


def _emit_layout(
    doc,
    layout,
    pt_per_px: float,
    usable_width_pt: float,
    first: bool = False,
    label: str | None = None,
    fidelity: bool = True,
) -> None:
    anchor = _page_label_text(doc, label) if label else None

    body_size = 11.0
    if fidelity and layout.text_height:
        body_size = max(7.5, min(16.0, layout.text_height * pt_per_px * 0.86))

    for text, align in layout.before:
        paragraph = _add_arabic_paragraph(doc, text, size=body_size, align=align)
        if anchor is None:
            anchor = paragraph

    # إعادة الفراغ الرأسي الذي كان يفصل النص عن الجدول في الأصل
    if fidelity and layout.gap_before > 0:
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_before = Pt(0)
        spacer.paragraph_format.space_after = Pt(0)
        spacer.paragraph_format.line_spacing = Pt(min(150.0, max(4.0, layout.gap_before * pt_per_px)))
        spacer.add_run("").font.size = Pt(1)

    if layout.grid and layout.grid.rows:
        if anchor is None:
            anchor = doc.add_paragraph()
            anchor.paragraph_format.space_after = Pt(0)
            anchor.add_run("").font.size = Pt(1)
        _add_ocr_table(doc, layout.grid, pt_per_px, usable_width_pt, fidelity=fidelity)
    for text, align in layout.after:
        _add_arabic_paragraph(
            doc, text, size=max(7.0, body_size - 1.5), gray=True, align=align
        )

    if not first and anchor is not None:
        anchor.paragraph_format.page_break_before = True


MARGIN_PT = 18.0


def _new_ocr_document(
    width_pt: float | None = None, height_pt: float | None = None, margin_pt: float = MARGIN_PT
) -> Document:
    doc = Document()
    section = doc.sections[0]
    if width_pt and height_pt:
        section.page_width = Pt(width_pt)
        section.page_height = Pt(height_pt)
    margin = Pt(margin_pt)
    section.top_margin = margin
    section.bottom_margin = margin
    section.left_margin = margin
    section.right_margin = margin
    # اتجاه المقطع من اليمين إلى اليسار
    sectPr = section._sectPr
    if sectPr.find(qn("w:bidi")) is None:
        _insert_ordered(sectPr, OxmlElement("w:bidi"), _AFTER_SECT_BIDI)
    style = doc.styles["Normal"]
    style.font.name = ARABIC_FONT
    style.font.size = Pt(11)
    return doc


def _page_label_text(doc, label: str):
    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.space_after = Pt(4)
    run = paragraph.add_run(label)
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    run.font.name = ARABIC_FONT
    return paragraph


def _page_size_pt(src: Path) -> tuple[float, float]:
    import pymupdf

    doc = pymupdf.open(str(src))
    try:
        rect = doc.load_page(0).rect
        return float(rect.width), float(rect.height)
    finally:
        doc.close()


def ocr_pdf_to_docx(src: Path, dst: Path, dpi: int = 200, fidelity: bool = True) -> dict:
    if fidelity:
        width_pt, height_pt = _page_size_pt(src)
    else:
        width_pt, height_pt = 595.0, 842.0  # A4 للتنسيق البسيط
    pages = ocr_engine.ocr_pdf(src, dpi=dpi)
    doc = _new_ocr_document(width_pt, height_pt)
    pt_per_px = 72.0 / dpi
    usable = width_pt - 2 * MARGIN_PT
    for index, page in enumerate(pages, start=1):
        layout = ocr_engine.analyze_page(page)
        label = f"— صفحة {index} من {len(pages)} —" if len(pages) > 1 else None
        _emit_layout(
            doc, layout, pt_per_px, usable, first=index == 1, label=label, fidelity=fidelity
        )
    doc.save(str(dst))
    return {"pages": len(pages), "ocr": True, "fidelity": fidelity}


def ocr_pdf_to_xlsx(src: Path, dst: Path, dpi: int = 200) -> dict:
    pages = ocr_engine.ocr_pdf(src, dpi=dpi)
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    tables = 0

    for index, page in enumerate(pages, start=1):
        layout = ocr_engine.analyze_page(page)
        ws = wb.create_sheet(_safe_sheet_name(f"صفحة {index}", used))
        ws.sheet_view.rightToLeft = True
        row_cursor = 1
        for text, _align in layout.before:
            ws.cell(row=row_cursor, column=1, value=text).alignment = Alignment(wrap_text=True)
            row_cursor += 1
        if layout.grid and layout.grid.rows:
            for r, cells in enumerate(layout.grid.rows):
                for c, value in enumerate(cells, start=1):
                    cell = ws.cell(row=row_cursor, column=c, value=value or None)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.border = BORDER
                    if r == 0:
                        cell.font = HEADER_FONT
                        cell.fill = HEADER_FILL
                row_cursor += 1
            tables += 1
        else:
            for text in layout.paragraphs:
                ws.cell(row=row_cursor, column=1, value=text)
                row_cursor += 1
        _autosize_columns(ws)

    if not wb.sheetnames:
        wb.create_sheet("النتيجة")["A1"] = "لم يُعثر على نص."
    wb.save(str(dst))
    return {"pages": len(pages), "tables": tables, "ocr": True}


def ocr_images_to_docx(images: Sequence[Path], dst: Path, fidelity: bool = True) -> dict:
    first = PILImage.open(images[0])
    width_px, height_px = first.size
    # نطابق عرض الصورة مع عرض A4 ليبقى التنسيق بنفس نسب الأصل
    a4_width_pt = 595.0
    pt_per_px = a4_width_pt / max(1, width_px)
    page_h_pt = height_px * pt_per_px

    doc = _new_ocr_document(a4_width_pt, page_h_pt)
    pages = 0
    text_found = 0
    usable = a4_width_pt - 2 * MARGIN_PT
    for index, img_path in enumerate(images):
        page = ocr_engine.ocr_image(PILImage.open(img_path))
        layout = ocr_engine.analyze_page(page)
        label = f"— صورة {index + 1} من {len(images)} —" if len(images) > 1 else None
        _emit_layout(
            doc, layout, pt_per_px, usable, first=index == 0, label=label, fidelity=fidelity
        )
        pages += 1
        if layout.grid or layout.before or layout.after or layout.paragraphs:
            text_found += 1
    doc.save(str(dst))
    return {"pages": pages, "with_text": text_found, "ocr": True}


def ocr_images_to_xlsx(images: Sequence[Path], dst: Path) -> dict:
    wb = Workbook()
    wb.remove(wb.active)
    used: set[str] = set()
    tables = 0
    for index, img_path in enumerate(images, start=1):
        page = ocr_engine.ocr_image(PILImage.open(img_path))
        layout = ocr_engine.analyze_page(page)
        ws = wb.create_sheet(_safe_sheet_name(f"صورة {index}", used))
        ws.sheet_view.rightToLeft = True
        row_cursor = 1
        for text, _align in layout.before:
            ws.cell(row=row_cursor, column=1, value=text).alignment = Alignment(wrap_text=True)
            row_cursor += 1
        if layout.grid and layout.grid.rows:
            for r, cells in enumerate(layout.grid.rows):
                for c, value in enumerate(cells, start=1):
                    cell = ws.cell(row=row_cursor, column=c, value=value or None)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.border = BORDER
                    if r == 0:
                        cell.font = HEADER_FONT
                        cell.fill = HEADER_FILL
                row_cursor += 1
            tables += 1
        else:
            for text in layout.paragraphs:
                ws.cell(row=row_cursor, column=1, value=text)
                row_cursor += 1
        _autosize_columns(ws)
    if not wb.sheetnames:
        wb.create_sheet("النتيجة")["A1"] = "لم يُعثر على نص."
    wb.save(str(dst))
    return {"pages": len(images), "tables": tables, "ocr": True}


def needs_ocr(src: Path, sample_dpi: int = 150) -> bool:
    """
    يحدّد إن كانت طبقة النص في PDF غير موثوقة (ملف ممسوح أو نص مشوّه).
    يقارن عيّنة OCR لصفحة واحدة مع طبقة النص: تطابق ضعيف = يلزم OCR.
    """
    import pymupdf

    try:
        doc = pymupdf.open(str(src))
    except Exception:
        return False
    try:
        page = doc.load_page(0)
        text = (page.get_text() or "").strip()
        page_area = abs(page.rect.width * page.rect.height) or 1.0
        covered = 0.0
        for item in page.get_images(full=True):
            try:
                for rect in page.get_image_rects(item[0]):
                    covered += abs(rect.width * rect.height)
            except Exception:
                continue
        has_full_image = covered / page_area > 0.5
        if not has_full_image:
            return False
        if len(text) < 80:
            return True
        if not ocr_engine.available():
            return False
        sample = ocr_engine.ocr_pdf_page(src, 0, dpi=sample_dpi).text
        return ocr_engine.similarity(text, sample) < 0.35
    except Exception:
        return False
    finally:
        doc.close()


# --------------------------------------------------------------------------
# الموزّع الرئيسي
# --------------------------------------------------------------------------
TARGET_EXT = {"docx": ".docx", "xlsx": ".xlsx", "pptx": ".pptx"}
MERGED_TARGETS = {"docx", "xlsx", "pptx"}


def convert(
    srcs: Sequence[Path],
    target: str,
    out_dir: Path,
    work_dir: Path,
    use_ocr: bool = True,
    fidelity: bool = True,
) -> tuple[Path, dict]:
    """يحوّل ملفًا واحدًا أو مجموعة ملفات متجانسة، ويرجّع (مسار الناتج، تفاصيل)."""
    if not srcs:
        raise ValueError("لم يتم إرسال أي ملف.")
    target = target.lower()
    if target not in TARGET_EXT:
        raise ValueError("صيغة هدف غير مدعومة.")

    pdfs = [p for p in srcs if is_pdf(p)]
    images = [p for p in srcs if is_image(p)]

    if len(pdfs) + len(images) != len(srcs):
        raise ValueError("نوع ملف غير مدعوم.")

    stem = re.sub(r'[\\/:*?"<>|]+', "_", srcs[0].stem)[:60] or "ملف"
    use_ocr = bool(use_ocr) and ocr_engine.available()

    # 1) صور متعددة + هدف يقبل الدمج  ->  ملف واحد
    if images and not pdfs and len(images) > 1:
        dst = out_dir / f"{stem}_و{len(images) - 1}{TARGET_EXT[target]}"
        if target == "docx":
            if use_ocr:
                meta = ocr_images_to_docx(images, dst, fidelity=fidelity)
                if meta.get("with_text"):
                    return dst, meta
            return dst, images_to_docx(images, dst)
        if target == "xlsx":
            if use_ocr:
                meta = ocr_images_to_xlsx(images, dst)
                if meta.get("tables") or meta.get("pages"):
                    return dst, meta
            return dst, images_to_xlsx(images, dst, work_dir)
        return dst, images_to_pptx(images, dst, work_dir)

    # 2) ملف واحد
    if len(srcs) == 1:
        src = srcs[0]
        dst = out_dir / f"{stem}{TARGET_EXT[target]}"
        if is_pdf(src):
            scanned = use_ocr and needs_ocr(src)
            if target == "docx":
                if scanned:
                    return dst, ocr_pdf_to_docx(src, dst, fidelity=fidelity)
                pdf_to_docx(src, dst)
                return dst, {"pages": None}
            if target == "xlsx":
                if scanned:
                    return dst, ocr_pdf_to_xlsx(src, dst)
                return dst, pdf_to_xlsx(src, dst)
            return dst, pdf_to_pptx(src, dst, work_dir)
        # صورة واحدة
        if target == "docx":
            if use_ocr:
                meta = ocr_images_to_docx([src], dst, fidelity=fidelity)
                if meta.get("with_text"):
                    return dst, meta
            return dst, images_to_docx([src], dst)
        if target == "xlsx":
            if use_ocr:
                meta = ocr_images_to_xlsx([src], dst)
                if meta.get("tables") or meta.get("pages"):
                    return dst, meta
            return dst, images_to_xlsx([src], dst, work_dir)
        return dst, images_to_pptx([src], dst, work_dir)

    # 3) حالات مختلطة تُدار في طبقة الخدمة (ملف لكل مدخل ثم ZIP)
    raise ValueError("لا يمكن تحويل عدة ملفات مختلفة النوع في طلب واحد.")
