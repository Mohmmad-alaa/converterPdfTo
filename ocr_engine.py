"""
محرّك OCR — يعتمد على Windows.Media.Ocr المدمج في ويندوز
=========================================================
لا يحتاج أي تنزيل خارجي ولا اتصال بالإنترنت، ويستخدم حزم اللغة المثبّتة
في النظام (العربية ar-SA والإنجليزية en-US على هذا الجهاز).

المخرجات: أسطر مع إحداثيات الكلمات، لبناء نص أو جداول حقيقية.
"""

from __future__ import annotations

import io
import threading
from dataclasses import dataclass, field
from typing import Iterable

from PIL import Image

_thread_local = threading.local()
_ARABIC_RANGE = range(0x0600, 0x0700)


@dataclass
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass
class Line:
    text: str
    words: list[Word] = field(default_factory=list)
    x0: float = 0.0
    y0: float = 0.0
    x1: float = 0.0
    y1: float = 0.0

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2

    @property
    def height(self) -> float:
        return self.y1 - self.y0


@dataclass
class OcrPage:
    lines: list[Line] = field(default_factory=list)
    width: int = 0
    height: int = 0
    v_lines: list[int] = field(default_factory=list)
    h_lines: list[int] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(line.text for line in self.lines if line.text.strip())


def available() -> bool:
    try:
        import winrt.windows.media.ocr  # noqa: F401
    except Exception:
        return False
    return _engine() is not None


def languages() -> list[str]:
    try:
        from winrt.windows.media.ocr import OcrEngine
    except Exception:
        return []
    return [lang.language_tag for lang in OcrEngine.available_recognizer_languages]


def _engine():
    """محرّك لكل خيط (Windows OCR غير مضمون للمشاركة بين الخيوط)."""
    engine = getattr(_thread_local, "engine", None)
    if engine is not None:
        return engine
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine
    except Exception:
        return None

    for tag in ("ar-SA", "ar", "en-US"):
        try:
            engine = OcrEngine.try_create_from_language(Language(tag))
        except Exception:
            engine = None
        if engine is not None:
            break
    if engine is None:
        engine = OcrEngine.try_create_from_user_profile_languages()

    _thread_local.engine = engine
    return engine


def _arabic_ratio(text: str) -> float:
    letters = [c for c in text if c.isalpha()]
    if not letters:
        return 0.0
    arabic = sum(1 for c in letters if ord(c) in _ARABIC_RANGE)
    return arabic / len(letters)


def _order_words(words: list[Word]) -> list[Word]:
    """Windows OCR يُرجع كلمات السطر من اليسار لليمين؛ نعيد ترتيبها للعربية."""
    joined = " ".join(w.text for w in words)
    if _arabic_ratio(joined) > 0.3:
        return sorted(words, key=lambda w: w.cx, reverse=True)
    return sorted(words, key=lambda w: w.cx)


def _group_positions(values: list[int], tolerance: int = 8) -> list[int]:
    out: list[list[int]] = []
    for value in values:
        if out and value - out[-1][-1] <= tolerance:
            out[-1].append(value)
        else:
            out.append([value])
    return [int(sum(group) / len(group)) for group in out]


def detect_grid_lines(img: Image.Image) -> tuple[list[int], list[int]]:
    """
    يستخرج مواضع خطوط الجدول (أفقية ورأسية) من الصورة عبر OpenCV.
    يعيد المحاولة بعتبات أخفض للصفحات الباهتة أو المائلة قليلًا.
    """
    try:
        import cv2
        import numpy as np
    except Exception:
        return [], []

    try:
        gray = np.array(img.convert("L"))
        height, width = gray.shape
        binary = cv2.adaptiveThreshold(
            ~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2
        )
        horizontal = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (max(10, width // 40), 1)),
        )
        vertical = cv2.morphologyEx(
            binary,
            cv2.MORPH_OPEN,
            cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(10, height // 40))),
        )
        h_proj = (horizontal > 0).sum(axis=1)
        v_proj = (vertical > 0).sum(axis=0)

        candidates: list[tuple[list[int], list[int]]] = []
        for threshold in (0.15, 0.08, 0.06):
            ys = _group_positions([i for i, c in enumerate(h_proj) if c > width * threshold])
            xs = _group_positions([i for i, c in enumerate(v_proj) if c > height * threshold])
            candidates.append((xs, ys))
            # شبكة مكتملة: أعمدة وصفوف كافية — لا حاجة لمحاولة أخفض
            if len(xs) >= 8 and len(ys) >= 20:
                return xs, ys

        def score(item: tuple[list[int], list[int]]) -> int:
            xs_i, ys_i = item
            if len(xs_i) > 25 or len(ys_i) > 90:  # ضجيج
                return -1
            return len(xs_i) + len(ys_i)

        return max(candidates, key=score)
    except Exception:
        return [], []


def ocr_image(img: Image.Image) -> OcrPage:
    """يشغّل OCR على صورة PIL ويرجّع الأسطر بإحداثياتها."""
    engine = _engine()
    if engine is None:
        raise RuntimeError("لا يوجد محرّك OCR متاح على هذا النظام.")

    from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapPixelFormat, SoftwareBitmap
    from winrt.windows.security.cryptography import CryptographicBuffer

    rgba = img.convert("RGBA")
    r, g, b, a = rgba.split()
    bgra = Image.merge("RGBA", (b, g, r, a)).tobytes()
    buffer = CryptographicBuffer.create_from_byte_array(bgra)
    bitmap = SoftwareBitmap.create_copy_from_buffer(
        buffer, BitmapPixelFormat.BGRA8, rgba.width, rgba.height, BitmapAlphaMode.PREMULTIPLIED
    )

    result = engine.recognize_async(bitmap).get()
    page = OcrPage(width=rgba.width, height=rgba.height)
    page.v_lines, page.h_lines = detect_grid_lines(rgba)

    for raw_line in list(result.lines):
        words = []
        for w in list(raw_line.words):
            rect = w.bounding_rect
            words.append(
                Word(
                    text=w.text,
                    x0=float(rect.x),
                    y0=float(rect.y),
                    x1=float(rect.x + rect.width),
                    y1=float(rect.y + rect.height),
                )
            )
        if not words:
            continue
        words = _order_words(words)
        page.lines.append(
            Line(
                text=_join_words(words),
                words=words,
                x0=min(w.x0 for w in words),
                y0=min(w.y0 for w in words),
                x1=max(w.x1 for w in words),
                y1=max(w.y1 for w in words),
            )
        )
    return page


def ocr_pdf_page(path, index: int, dpi: int = 250) -> OcrPage:
    """يصيّر صفحة PDF إلى صورة ثم يشغّل عليها OCR."""
    import pymupdf

    doc = pymupdf.open(str(path))
    try:
        pix = doc.load_page(index).get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        return ocr_image(img)
    finally:
        doc.close()


def ocr_pdf(path, dpi: int = 250, progress=None) -> list[OcrPage]:
    """يشغّل OCR على كل صفحات PDF بالتوازي."""
    import pymupdf
    from concurrent.futures import ThreadPoolExecutor

    doc = pymupdf.open(str(path))
    try:
        count = doc.page_count
        pixmaps = []
        for i in range(count):
            pix = doc.load_page(i).get_pixmap(dpi=dpi)
            pixmaps.append(pix.tobytes("png"))
    finally:
        doc.close()

    def work(data: bytes) -> OcrPage:
        return ocr_image(Image.open(io.BytesIO(data)))

    workers = min(4, max(1, (len(pixmaps) or 1)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        pages = list(pool.map(work, pixmaps))

    if progress:
        for i, _ in enumerate(pages, start=1):
            progress(i, len(pages))
    return pages


def word_sets(text: str) -> set[str]:
    """كلمات مطبَّعة للمقارنة بين طبقة نص ونتيجة OCR."""
    out: set[str] = set()
    for token in text.split():
        clean = "".join(c for c in token if c.isalnum())
        if len(clean) >= 3:
            out.add(clean.lower())
    return out


def similarity(a: str, b: str) -> float:
    """معامل تشابه (Jaccard) بين كلمات نصين — لكشف طبقة النص المشوّهة."""
    sa, sb = word_sets(a), word_sets(b)
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


# --------------------------------------------------------------------------
# تحليل التخطيط: فصل النص الحر عن الجداول وبناء شبكة خلايا
# --------------------------------------------------------------------------
@dataclass
class Grid:
    rows: list[list[str]] = field(default_factory=list)
    n_cols: int = 0
    # قياسات هندسية بالبكسل (بنفس ترتيب الأعمدة/الصفوف) لبناء تنسيق مطابق للأصل
    col_bounds: list[tuple[float, float]] = field(default_factory=list)
    row_bounds: list[tuple[float, float]] = field(default_factory=list)
    cell_heights: list[list[float]] = field(default_factory=list)
    header_rows: int = 1

    @property
    def header(self) -> list[str]:
        return self.rows[0] if self.rows else []


@dataclass
class PageLayout:
    paragraphs: list[str] = field(default_factory=list)
    grid: Grid | None = None
    before: list[tuple[str, str]] = field(default_factory=list)
    after: list[tuple[str, str]] = field(default_factory=list)
    text_height: float = 0.0
    page_width: float = 0.0
    page_height: float = 0.0
    gap_before: float = 0.0


def _median(values: list[float]) -> float:
    if not values:
        return 0.0
    ordered = sorted(values)
    return ordered[len(ordered) // 2]


def _cluster(values: list[float], tolerance: float) -> list[float]:
    """تجميع قيم متقاربة وإرجاع مراكز المجموعات."""
    if not values:
        return []
    ordered = sorted(values)
    groups: list[list[float]] = [[ordered[0]]]
    for value in ordered[1:]:
        if value - groups[-1][-1] <= tolerance:
            groups[-1].append(value)
        else:
            groups.append([value])
    return [sum(g) / len(g) for g in groups]


def _column_bands(words: list[Word], width: float, n_rows: int = 1) -> list[tuple[float, float]]:
    """
    يكتشف أعمدة الجدول عبر إسقاط الكلمات على المحور الأفقي:
    الموضع الذي تغطّيه كلمات قليلة جدًا (سطر ترويسة عريض مثلًا) لا يُعدّ عمودًا،
    والفراغ المستمر بين الأعمدة يُعدّ فاصلًا.
    """
    if not words:
        return []
    lo = min(w.x0 for w in words)
    hi = max(w.x1 for w in words)
    span = max(1.0, hi - lo)
    buckets = 1400
    scale = buckets / span
    occupied = [0] * (buckets + 2)
    for w in words:
        a = int((w.x0 - lo) * scale)
        b = int((w.x1 - lo) * scale)
        for i in range(max(0, a), min(buckets + 1, b + 1)):
            occupied[i] += 1

    solid = max(2, int(n_rows * 0.10))
    min_gap = max(8, int(span * 0.005 * scale))

    bands: list[tuple[float, float]] = []
    start = None
    gap = 0
    for i, count in enumerate(occupied):
        if count >= solid:
            if start is None:
                start = i
            gap = 0
        else:
            if start is not None:
                gap += 1
                if gap >= min_gap:
                    bands.append((lo + start / scale, lo + (i - gap + 1) / scale))
                    start = None
                    gap = 0
    if start is not None:
        bands.append((lo + start / scale, lo + buckets / scale))

    counts = [sum(1 for w in words if a - 3 <= w.cx <= b + 3) for a, b in bands]
    return [band for band, c in zip(bands, counts) if c >= 3 and band[1] - band[0] >= 8]


def analyze_page(page: OcrPage) -> PageLayout:
    """
    يبني تخطيط الصفحة: فقرات نص حر + جدول إن وُجد.
    يعتمد على خطوط الجدول المرسومة إن أمكن، وإلا على إسقاط الإحداثيات.
    """
    if page.v_lines and page.h_lines:
        result = _analyze_with_lines(page)
        if result is not None:
            return result
    return _analyze_with_projection(page)


def _join_words(words: list[Word]) -> str:
    """
    يجمع كلمات سطر واحد بترتيب قراءة عربي (يمين → يسار)،
    مع إبقاء المقاطع اللاتينية بترتيبها الطبيعي (يسار → يمين).
    """
    if not words:
        return ""
    joined = " ".join(w.text for w in words)
    if _arabic_ratio(joined) <= 0.3:
        return " ".join(w.text for w in sorted(words, key=lambda w: w.cx))

    ordered = sorted(words, key=lambda w: w.cx, reverse=True)
    out: list[str] = []
    run: list[Word] = []
    for word in ordered:
        if _arabic_ratio(word.text) <= 0.3 and any(c.isalpha() for c in word.text):
            run.append(word)
        else:
            if run:
                out.extend(w.text for w in reversed(run))
                run = []
            out.append(word.text)
    if run:
        out.extend(w.text for w in reversed(run))
    return " ".join(out)


def _words_to_paragraphs(
    words: list[Word], line_height: float, page_width: float
) -> list[tuple[str, str]]:
    """يعيد بناء أسطر النص الحر مع محاذاتها (وسط أو يمين)."""
    if not words:
        return []
    ordered = sorted(words, key=lambda w: w.cy)
    groups: list[list[Word]] = [[ordered[0]]]
    for w in ordered[1:]:
        if abs(w.cy - groups[-1][-1].cy) <= max(6.0, line_height * 0.55):
            groups[-1].append(w)
        else:
            groups.append([w])

    out: list[tuple[str, str]] = []
    for group in groups:
        # قد يحتوي السطر على كتل منفصلة (يمين ووسط ويسار) — نفصلها بفراغ كبير
        by_x = sorted(group, key=lambda w: w.cx)
        segments: list[list[Word]] = [[by_x[0]]]
        for w in by_x[1:]:
            if w.x0 - segments[-1][-1].x1 > page_width * 0.05:
                segments.append([w])
            else:
                segments[-1].append(w)

        for segment in reversed(segments):  # من اليمين إلى اليسار
            text = _join_words(segment)
            if not text.strip():
                continue
            if len(text.strip()) <= 2 and not any(c.isdigit() for c in text):
                continue
            x0 = min(w.x0 for w in segment)
            x1 = max(w.x1 for w in segment)
            center = (x0 + x1) / 2
            if abs(center - page_width / 2) < page_width * 0.08:
                align = "center"
            elif center > page_width / 2:
                align = "right"
            else:
                align = "left"
            out.append((text, align))
    return out


def _analyze_with_lines(page: OcrPage) -> PageLayout | None:
    xs, ys = page.v_lines, page.h_lines
    col_bands = [(xs[i], xs[i + 1]) for i in range(len(xs) - 1) if xs[i + 1] - xs[i] >= 20]
    row_bands = [(ys[i], ys[i + 1]) for i in range(len(ys) - 1) if ys[i + 1] - ys[i] >= 15]
    if len(col_bands) < 2 or len(row_bands) < 3:
        return None

    top, bottom = ys[0], ys[-1]
    left, right = xs[0], xs[-1]
    words = [w for ln in page.lines for w in ln.words]
    if not words:
        return None

    inside = [w for w in words if top - 15 <= w.cy <= bottom + 15 and left - 15 <= w.cx <= right + 15]
    inside_ids = {id(w) for w in inside}
    outside = [w for w in words if id(w) not in inside_ids]
    if not inside:
        return None

    def pick(value: float, bands: list[tuple[float, float]]) -> int:
        for i, (a, b) in enumerate(bands):
            if a - 6 <= value <= b + 6:
                return i
        return min(range(len(bands)), key=lambda i: abs((bands[i][0] + bands[i][1]) / 2 - value))

    matrix: dict[tuple[int, int], list[Word]] = {}
    for w in inside:
        matrix.setdefault((pick(w.cy, row_bands), pick(w.cx, col_bands)), []).append(w)

    col_order = sorted(range(len(col_bands)), key=lambda i: -col_bands[i][0])

    def cell_text(items: list[Word]) -> str:
        return _join_words(items).strip()

    # حذف الأعمدة الوهمية فقط: خطوط مزدوجة رقيقة أو رموز عابرة.
    # الأعمدة الفارغة الحقيقية تبقى حتى يطابق الجدول أبعاد الأصل.
    keep = []
    n_rows = len(row_bands)
    for c in col_order:
        left_x, right_x = col_bands[c]
        if right_x - left_x < 25:
            continue
        cells = [cell_text(matrix.get((r, c), [])) for r in range(n_rows)]
        filled = [text for text in cells if text]
        if filled and len(filled) < n_rows * 0.2 and max(len(text) for text in filled) <= 2:
            continue
        keep.append(c)
    if len(keep) < 2:
        return None

    # الأعمدة الفارغة تمامًا (خطوط عابرة أو فراغ داخل خلية) تُحذف
    def _has_data(column: int) -> bool:
        return any(cell_text(matrix.get((r, column), [])) for r in range(n_rows))

    keep = [c for c in keep if _has_data(c)]
    if len(keep) < 2:
        return None

    rows: list[list[str]] = []
    row_heights: list[list[float]] = []
    row_indices: list[int] = []
    for r in range(len(row_bands)):
        cells = [cell_text(matrix.get((r, c), [])) for c in keep]
        if any(cells):
            rows.append(cells)
            row_indices.append(r)
            row_heights.append(
                [
                    _median([w.height for w in matrix.get((r, c), [])]) or 0.0
                    for c in keep
                ]
            )
    if len(rows) < 3:
        return None

    header_rows = 1
    merged_header = False
    # دمج صفّي الترويسة إن كانا متكاملين (مثل: "رقم" + "القطعة")
    if len(rows) >= 2:
        first, second = rows[0], rows[1]
        n = len(first)
        filled_first = sum(1 for c in first if c)
        filled_second = sum(1 for c in second if c)
        has_digits = any(ch.isdigit() for cell in second for ch in cell)
        if (
            filled_first >= n * 0.5
            and filled_second <= n * 0.6
            and not has_digits
            and filled_first + filled_second > max(filled_first, filled_second)
        ):
            merged: list[str] = []
            complementary = False
            for a, b in zip(first, second):
                if a and b:
                    merged.append(a if a == b else f"{a} {b}")
                elif a or b:
                    merged.append(a or b)
                    complementary = True
                else:
                    merged.append("")
            if complementary:
                rows = [merged] + rows[2:]
                row_heights = [
                    [max(h1, h2) for h1, h2 in zip(row_heights[0], row_heights[1])]
                ] + row_heights[2:]
                row_indices = [row_indices[0]] + row_indices[2:]
                merged_header = True

    # حدود الصفوف بالبكسل (بنفس ترتيب rows)
    row_bounds: list[tuple[float, float]] = []
    for i, r in enumerate(row_indices):
        top_y, bottom_y = row_bands[r]
        if i == 0 and merged_header:
            bottom_y = row_bands[r + 1][1]
        row_bounds.append((top_y, bottom_y))

    layout = PageLayout()
    layout.page_width = float(page.width)
    layout.page_height = float(page.height)
    layout.grid = Grid(
        rows=rows,
        n_cols=len(keep),
        col_bounds=[col_bands[c] for c in keep],
        row_bounds=row_bounds,
        cell_heights=row_heights,
        header_rows=header_rows,
    )
    line_height = _median([w.height for w in words]) or 20
    above = [w for w in outside if w.cy < top]
    layout.text_height = _median([w.height for w in outside]) or line_height
    layout.gap_before = max(0.0, top - max((w.y1 for w in above), default=top))
    layout.before = _words_to_paragraphs(above, line_height, float(page.width))
    layout.after = _words_to_paragraphs(
        [w for w in outside if w.cy > bottom], line_height, float(page.width)
    )
    return layout


def _analyze_with_projection(page: OcrPage) -> PageLayout:
    width = page.width or 1
    layout = PageLayout()
    flow_lines: list[Line] = []
    table_lines: list[Line] = []

    for line in page.lines:
        span = line.x1 - line.x0
        if span > 0.42 * width and len(line.words) >= 4:
            flow_lines.append(line)
        else:
            table_lines.append(line)

    if len(table_lines) < 5:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    words = [w for ln in table_lines for w in ln.words]
    if not words:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    line_height = _median([w.height for w in words]) or 20
    centers = _cluster([w.cy for w in words], tolerance=max(6.0, line_height * 0.55))
    if len(centers) < 3:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    bands = _column_bands(words, width, n_rows=len(centers))
    if len(bands) < 2:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    def band_index(x: float) -> int:
        for i, (start, end) in enumerate(bands):
            if start - 6 <= x <= end + 6:
                return i
        return min(range(len(bands)), key=lambda i: abs((bands[i][0] + bands[i][1]) / 2 - x))

    def row_index(y: float) -> int:
        return min(range(len(centers)), key=lambda i: abs(centers[i] - y))

    matrix: dict[tuple[int, int], list[Word]] = {}
    for w in words:
        matrix.setdefault((row_index(w.cy), band_index(w.cx)), []).append(w)

    # ترتيب الأعمدة من اليمين إلى اليسار (اتجاه القراءة العربي)
    col_order = sorted(range(len(bands)), key=lambda i: -bands[i][0])

    def cell_text(items: list[Word]) -> str:
        return _join_words(items).strip()

    rows: list[tuple[float, list[str]]] = []
    for r, center in enumerate(centers):
        cells = [cell_text(matrix.get((r, c), [])) for c in col_order]
        if any(cells):
            rows.append((center, cells))

    if len(rows) < 3:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    # كتلة الجدول = أطول سلسلة صفوف متتالية تحوي خلايا كافية
    min_cells = 3 if len(col_order) >= 5 else 2
    flags = [sum(1 for cell in cells if cell) >= min_cells for _, cells in rows]
    best_start, best_len, current = 0, 0, None
    for i, flag in enumerate(flags):
        if flag:
            if current is None:
                current = i
            if i - current + 1 > best_len:
                best_start, best_len = current, i - current + 1
        else:
            current = None
    if best_len < 3:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    block = [cells for _, cells in rows[best_start : best_start + best_len]]
    filled = sum(1 for row in block for cell in row if cell)
    density = filled / max(1, len(block) * len(col_order))
    if density < 0.15:
        layout.paragraphs = [ln.text for ln in page.lines if ln.text.strip()]
        return layout

    # دمج صفّي الترويسة إن كانا متكاملين (مثل: "رقم" + "القطعة")
    if len(block) >= 2:
        first, second = block[0], block[1]
        merged: list[str] = []
        complementary = False
        for a, b in zip(first, second):
            if a and b:
                merged.append(a if a == b else f"{a} {b}")
            elif a or b:
                merged.append(a or b)
                complementary = True
            else:
                merged.append("")
        if complementary:
            block = [merged] + block[2:]

    layout.grid = Grid(rows=block, n_cols=len(col_order))

    table_top = rows[best_start][0]
    table_bottom = rows[best_start + best_len - 1][0]
    extras_before = [
        (" ".join(c for c in cells if c), "right")
        for y, cells in rows[:best_start]
        if y < table_top
    ]
    extras_after = [
        (" ".join(c for c in cells if c), "right")
        for y, cells in rows[best_start + best_len :]
        if y > table_bottom
    ]

    layout.page_width = float(page.width)
    layout.page_height = float(page.height)
    layout.text_height = line_height
    layout.before = [
        (ln.text, "right") for ln in flow_lines if ln.y1 <= table_top and ln.text.strip()
    ]
    layout.before.extend(extras_before)
    layout.after = [
        (ln.text, "right") for ln in flow_lines if ln.y0 >= table_bottom and ln.text.strip()
    ]
    layout.after.extend(extras_after)
    return layout
