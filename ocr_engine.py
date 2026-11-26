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
    image: Image.Image | None = field(default=None, repr=False, compare=False)
    native_words: list[Word] = field(default_factory=list)

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


def _english_engine():
    """محرّك إنجليزي منفصل لقراءة الأرقام داخل خلايا الجداول بدقة أعلى."""
    engine = getattr(_thread_local, "english_engine", None)
    if engine is not None:
        return engine
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine
        engine = OcrEngine.try_create_from_language(Language("en-US"))
    except Exception:
        engine = None
    _thread_local.english_engine = engine
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
        edge_margin = max(2, int(height * 0.005))
        for threshold in (0.15, 0.12, 0.10, 0.08, 0.06):
            ys = _group_positions([i for i, c in enumerate(h_proj) if c > width * threshold])
            # تجاهل حدود الصفحة نفسها (خطوط أعلى/أسفل الصورة) لأنها ليست صفوف جدول
            ys = [y for y in ys if edge_margin < y < height - edge_margin]
            xs = _group_positions([i for i, c in enumerate(v_proj) if c > height * threshold])
            candidates.append((xs, ys))

        def score(item: tuple[list[int], list[int]]) -> int:
            xs_i, ys_i = item
            if len(xs_i) < 2 or len(ys_i) < 3:
                return -1
            if len(xs_i) > 25 or len(ys_i) > 90:  # ضجيج
                return -1
            return len(xs_i) + len(ys_i)

        # لا نتوقف عند أول عتبة: العتبة المرتفعة قد تلتقط أعلى الجدول فقط
        # وتفوّت الصفوف السفلية الباهتة، لذلك نختار أفضل مرشح من كل العتبات.
        return max(candidates, key=score)
    except Exception:
        return [], []


def _enhance_for_ocr(img: Image.Image) -> Image.Image:
    """تحسين تباين الصورة لإبراز الحروف الخافتة ورفع دقة التعرف."""
    try:
        import cv2
        import numpy as np
        cv_img = np.array(img.convert("RGB"))
        gray = cv2.cvtColor(cv_img, cv2.COLOR_RGB2GRAY)
        clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        return Image.fromarray(enhanced).convert("RGBA")
    except Exception:
        return img.convert("RGBA")


def ocr_image(img: Image.Image) -> OcrPage:
    """يشغّل OCR على صورة PIL ويرجّع الأسطر بإحداثياتها."""
    engine = _engine()
    if engine is None:
        raise RuntimeError("لا يوجد محرّك OCR متاح على هذا النظام.")

    from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapPixelFormat, SoftwareBitmap
    from winrt.windows.security.cryptography import CryptographicBuffer

    rgba = _enhance_for_ocr(img)
    r, g, b, a = rgba.split()
    bgra = Image.merge("RGBA", (b, g, r, a)).tobytes()
    buffer = CryptographicBuffer.create_from_byte_array(bgra)
    bitmap = SoftwareBitmap.create_copy_from_buffer(
        buffer, BitmapPixelFormat.BGRA8, rgba.width, rgba.height, BitmapAlphaMode.PREMULTIPLIED
    )

    result = engine.recognize_async(bitmap).get()
    page = OcrPage(width=rgba.width, height=rgba.height, image=img.convert("RGB"))
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


def _attach_native_words(page: OcrPage, words, rect) -> None:
    """يحوّل كلمات طبقة نص PDF إلى إحداثيات الصورة — احتياط للأرقام التي يفوّتها OCR."""
    if not words or page.width <= 0 or page.height <= 0:
        return
    sx = page.width / max(1.0, rect.width)
    sy = page.height / max(1.0, rect.height)
    for item in words:
        if len(item) < 5:
            continue
        text = str(item[4]).strip()
        if not text or not any(ch.isdigit() for ch in text):
            continue
        page.native_words.append(
            Word(
                text=text,
                x0=float(item[0]) * sx,
                y0=float(item[1]) * sy,
                x1=float(item[2]) * sx,
                y1=float(item[3]) * sy,
            )
        )


def ocr_pdf_page(path, index: int, dpi: int = 250) -> OcrPage:
    """يصيّر صفحة PDF إلى صورة ثم يشغّل عليها OCR."""
    import pymupdf

    doc = pymupdf.open(str(path))
    try:
        pdf_page = doc.load_page(index)
        pix = pdf_page.get_pixmap(dpi=dpi)
        img = Image.open(io.BytesIO(pix.tobytes("png")))
        page = ocr_image(img)
        _attach_native_words(page, pdf_page.get_text("words"), pdf_page.rect)
        return page
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
        native = []
        rects = []
        for i in range(count):
            pdf_page = doc.load_page(i)
            pix = pdf_page.get_pixmap(dpi=dpi)
            pixmaps.append(pix.tobytes("png"))
            native.append(pdf_page.get_text("words"))
            rects.append(pdf_page.rect)
    finally:
        doc.close()

    def work(data: bytes) -> OcrPage:
        return ocr_image(Image.open(io.BytesIO(data)))

    workers = min(4, max(1, (len(pixmaps) or 1)))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        pages = list(pool.map(work, pixmaps))

    for page, words, rect in zip(pages, native, rects):
        _attach_native_words(page, words, rect)

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
    # خلايا الترويسة المدمجة أفقيًا: (بداية العمود، نهايته، النص)
    header_spans: list[tuple[int, int, str]] = field(default_factory=list)

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


import re

# الحروف اللاتينية التي يخلطها OCR مع الأرقام داخل الخلايا الرقمية
_DIGIT_CONFUSABLE = str.maketrans(
    {
        "I": "1", "l": "1", "|": "1", "!": "1",
        "O": "0", "o": "0", "Q": "0",
        "S": "5", "s": "5",
        "Z": "2", "z": "2",
        "B": "8", "G": "6",
    }
)
_NUMERIC_TOKEN_RE = re.compile(r"^[0-9٠-٩.,:/\-+%()\sIl|!OoQqSsZzBbGg]+$")


def _normalize_digit_token(token: str, require_digit: bool = True) -> str:
    """يصحح خلط الأرقام بالحروف فقط داخل رمز رقمي، ولا يمسّ الكلمات العربية."""
    if not token:
        return token
    if require_digit and not any(ch.isdigit() for ch in token):
        return token
    if not _NUMERIC_TOKEN_RE.fullmatch(token):
        return token
    return token.translate(_DIGIT_CONFUSABLE)


def _normalize_numeric_text(text: str) -> str:
    """يطبّق تصحيح الأرقام على خلية رقمية كاملة (حتى لو كانت رمزًا واحدًا مثل l)."""
    return " ".join(_normalize_digit_token(tok, require_digit=False) for tok in text.split())


def _is_numeric_like(text: str) -> bool:
    text = text.strip()
    return bool(text) and bool(_NUMERIC_TOKEN_RE.fullmatch(text))


def _ocr_row_numeric_words(img: Image.Image) -> list[tuple[float, float, str]]:
    """يقرأ صفًا كاملًا بمحرّك إنجليزي ويعيد (cx, cy, النص) للأرقام فقط."""
    engine = _english_engine()
    if engine is None or img.width < 8 or img.height < 8:
        return []
    try:
        from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapPixelFormat, SoftwareBitmap
        from winrt.windows.security.cryptography import CryptographicBuffer

        scale = 2
        big = img.resize((img.width * scale, img.height * scale), Image.LANCZOS).convert("RGBA")
        r, g, b, a = big.split()
        bgra = Image.merge("RGBA", (b, g, r, a)).tobytes()
        buffer = CryptographicBuffer.create_from_byte_array(bgra)
        bitmap = SoftwareBitmap.create_copy_from_buffer(
            buffer, BitmapPixelFormat.BGRA8, big.width, big.height, BitmapAlphaMode.PREMULTIPLIED
        )
        result = engine.recognize_async(bitmap).get()
        out: list[tuple[float, float, str]] = []
        for line in list(result.lines):
            for word in list(line.words):
                text = _normalize_numeric_text(word.text.strip())
                if not _is_numeric_like(text):
                    continue
                rect = word.bounding_rect
                out.append(
                    (
                        (rect.x + rect.width / 2) / scale,
                        (rect.y + rect.height / 2) / scale,
                        text,
                    )
                )
        return out
    except Exception:
        return []


def _ocr_numeric_strip(img: Image.Image) -> list[tuple[float, str]]:
    """يقرأ عمودًا رقميًا كاملًا بمحرّك إنجليزي مكبّر، ويعيد (y, النص) لكل سطر."""
    engine = _english_engine()
    if engine is None or img.width < 8 or img.height < 8:
        return []
    try:
        from winrt.windows.graphics.imaging import BitmapAlphaMode, BitmapPixelFormat, SoftwareBitmap
        from winrt.windows.security.cryptography import CryptographicBuffer

        scale = 3
        big = img.resize((img.width * scale, img.height * scale), Image.LANCZOS).convert("RGBA")
        r, g, b, a = big.split()
        bgra = Image.merge("RGBA", (b, g, r, a)).tobytes()
        buffer = CryptographicBuffer.create_from_byte_array(bgra)
        bitmap = SoftwareBitmap.create_copy_from_buffer(
            buffer, BitmapPixelFormat.BGRA8, big.width, big.height, BitmapAlphaMode.PREMULTIPLIED
        )
        result = engine.recognize_async(bitmap).get()
        out: list[tuple[float, str]] = []
        for line in list(result.lines):
            text = _normalize_numeric_text(line.text.strip())
            if not _is_numeric_like(text):
                continue
            words = list(line.words)
            if not words:
                continue
            cy = sum(w.bounding_rect.y + w.bounding_rect.height / 2 for w in words) / len(words)
            out.append((cy / scale, text))
        return out
    except Exception:
        return []


def clean_arabic_ocr(text: str) -> str:
    """تصحيح أخطاء OCR الشائعة في العربية (الحروف المقطوعة والأرقام المشوهة)."""
    if not text:
        return ""
    # إزالة محارف التنسيق المخفية
    text = re.sub(r"[\u200b-\u200f\ufeff]", "", text)
    # دمج الحروف والكلمات المقطوعة بسبب مسافات OCR الزائدة
    text = re.sub(r"\bر\s+قم\b", "رقم", text)
    text = re.sub(r"\bار\s+ض\b", "ارض", text)
    text = re.sub(r"\bمير\s+ي\b", "ميري", text)
    text = re.sub(r"\bميد\s*ي\b", "ميري", text)
    text = re.sub(r"\bنوج\s+الارض\b", "نوع الارض", text)
    text = re.sub(r"\bنوج\b", "نوع", text)
    text = re.sub(r"\bار\s+ض\b", "ارض", text)
    text = re.sub(r"^ي$", "ميري", text.strip())
    text = re.sub(r"\bقط\s+عة\b", "قطعة", text)
    text = re.sub(r"\bالحموق\b", "الحقوق", text)
    text = re.sub(r"\bبممتضى\b", "بمقتضى", text)
    text = re.sub(r"\bالفمره\b", "الفقرة", text)
    text = re.sub(r"\bلانون\b", "قانون", text)
    text = re.sub(r"\bتسو\s*ٌ?\s*ية\b", "تسوية", text)
    text = re.sub(r"\bالم\s*ٌ?\s*اه\b", "المياه", text)
    text = re.sub(r"\bمم\s*ٌ?\s*اس\b", "مقياس", text)
    text = re.sub(r"\bلف\s*ٌ?\s*ن\b", "قفين", text)
    text = re.sub(r"\bفلسط\s*ٌ?\s*ن\b", "فلسطين", text)
    text = re.sub(r"\bI/IOOO\b", "1/1000", text)
    text = re.sub(r"\bll\b", "11", text)
    text = re.sub(r"٠\s*٤", "40", text)
    # إصلاح الأرقام المختلطة بحروف لاتينية: 4lO -> 410، 94O -> 940 ...
    text = " ".join(_normalize_digit_token(tok) for tok in text.split())
    return text


def _join_words(words: list[Word]) -> str:
    """
    يجمع كلمات سطر واحد بترتيب قراءة عربي (يمين → يسار)،
    مع إبقاء المقاطع اللاتينية بترتيبها الطبيعي (يسار → يمين).
    """
    if not words:
        return ""
    joined = " ".join(w.text for w in words)
    if _arabic_ratio(joined) <= 0.3:
        raw = " ".join(w.text for w in sorted(words, key=lambda w: w.cx))
        return clean_arabic_ocr(raw)

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
    return clean_arabic_ocr(" ".join(out))


def _join_cell_words(words: list[Word]) -> str:
    """يجمع نصوص خلية الجدول بترتيب أسطر رأسي أولاً ثم أفقي، لمنع تشابك الكلمات."""
    if not words:
        return ""
    if len(words) == 1:
        return clean_arabic_ocr(words[0].text.strip())
    line_h = _median([w.height for w in words]) or 15.0
    ordered = sorted(words, key=lambda w: w.cy)
    lines: list[list[Word]] = [[ordered[0]]]
    for w in ordered[1:]:
        if abs(w.cy - lines[-1][-1].cy) <= max(6.0, line_h * 0.55):
            lines[-1].append(w)
        else:
            lines.append([w])
    return clean_arabic_ocr(" ".join(_join_words(ln) for ln in lines).strip())


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

    def cell_text(items: list[Word]) -> str:
        return _join_cell_words(items).strip()

    def raw(r: int, c: int) -> str:
        return cell_text(matrix.get((r, c), []))

    n_rows = len(row_bands)
    # كل الأعمدة ذات العرض الحقيقي — لا نحذف عموداً لأنه فارغ فقط
    real = [c for c in sorted(range(len(col_bands))) if col_bands[c][1] - col_bands[c][0] >= 25]

    # دمج الشقوق الوهمية: عمود عريض بلا أي نص ترويسة ليس عموداً مستقلاً
    # (مثل الخط الوهمي داخل خلية «رقم الهوية / المالك» أو «رقم القطعة»)
    groups: list[list[int]] = []
    for c in real:
        width = col_bands[c][1] - col_bands[c][0]
        empty_header = not raw(0, c) and not raw(1, c)
        if groups and width >= 120 and empty_header:
            groups[-1].append(c)
        else:
            groups.append([c])

    def group_text(r: int, group: list[int]) -> str:
        parts = [raw(r, c) for c in group]
        return " ".join(p for p in parts if p).strip()

    def group_width(group: list[int]) -> float:
        return col_bands[group[-1]][1] - col_bands[group[0]][0]

    # فحص ما إذا كان الصف 1 ترويسة فرعية حقيقية (مثل د / ف / م2) أم أنه صف بيانات
    row1_raw = [group_text(1, g) for g in groups]
    has_large_number = any(re.search(r"\d{3,}", c) for c in row1_raw)
    has_long_name = any(len(c.split()) >= 3 for c in row1_raw)
    is_real_sub_header = not (has_large_number or has_long_name) and any(
        c in {"د", "ف", "م2", "م٢", "سم", "فلس", "دينار", "المجموع"} or len(c) <= 4
        for c in row1_raw if c
    )

    header_top = [group_text(0, g) for g in groups]
    header_bottom = [group_text(1, g) for g in groups] if is_real_sub_header else ["" for _ in groups]
    header_spans: list[tuple[int, int, str]] = []

    if is_real_sub_header:
        span_members: set[int] = set()
        for i, group in enumerate(groups):
            if i in span_members or header_top[i] or group_width(group) >= 120:
                continue
            candidates = [
                j
                for j in (i + 1, i - 1)
                if 0 <= j < len(groups) and header_top[j]
            ]
            parent = next(
                (
                    j
                    for j in candidates
                    if "قيم" in header_top[j] or "مساح" in header_top[j]
                ),
                None,
            )
            if parent is None:
                continue
            parent_text = header_top[parent]
            if "قيم" in parent_text:
                parent_label, child_label = "د", "ف"
            else:  # مساحة
                parent_label, child_label = "م2", "د"
            header_bottom[parent] = parent_label
            header_bottom[i] = child_label
            start, end = (parent, i) if parent < i else (i, parent)
            header_spans.append((start, end, parent_text))
            for k in range(start, end + 1):
                header_top[k] = parent_text if k == start else ""
            span_members.update(range(start, end + 1))

        # بقية الأعمدة: دمج سطري الترويسة في عنوان واحد (نوع + الأرض = نوع الأرض)
        for i, group in enumerate(groups):
            if any(start <= i <= end for start, end, _ in header_spans):
                continue
            if header_bottom[i]:
                header_top[i] = f"{header_top[i]} {header_bottom[i]}".strip()
                header_bottom[i] = ""

    rows: list[list[str]] = []
    row_heights: list[list[float]] = []
    row_indices: list[int] = []
    for r in range(n_rows):
        cells = [group_text(r, g) for g in groups]
        if r == 0:
            cells = list(header_top)
        elif r == 1 and is_real_sub_header:
            cells = list(header_bottom)
        if any(cells) or r < 2:
            rows.append(cells)
            row_indices.append(r)
            row_heights.append(
                [
                    max(
                        (_median([w.height for w in matrix.get((r, c), [])]) or 0.0)
                        for c in group
                    )
                    for group in groups
                ]
            )
    if len(rows) < 3:
        return None

    # أعمدة رقمية: صحّح فيها الرموز المفردة التي قرأها OCR حروفًا (l -> 1، O -> 0)
    data_start = 2 if is_real_sub_header else 1
    if len(rows) > data_start:
        for c in range(len(rows[0])):
            group = groups[c]
            values = [rows[r][c] for r in range(data_start, len(rows)) if rows[r][c].strip()]
            if not values:
                continue
            numeric = sum(1 for value in values if _is_numeric_like(value))
            if numeric >= max(2, int(len(values) * 0.6)):
                for r in range(data_start, len(rows)):
                    rows[r][c] = _normalize_numeric_text(rows[r][c])

                # احتياط أول: طبقة نص PDF قد تحمل الرقم الصحيح رغم تشوّه العربي.
                # قد يكون الرقم مقطوعًا إلى كلمتين (1 + 1936) بسبب الختم، لذا نجمع
                # كلمات الخلية المتقاربة قبل المقارنة.
                if page.native_words:
                    nx0 = col_bands[group[0]][0]
                    nx1 = col_bands[group[-1]][1]
                    by_row: dict[int, list[Word]] = {}
                    for nw in page.native_words:
                        if not (nx0 <= nw.cx <= nx1):
                            continue
                        target = next(
                            (
                                r
                                for r in range(data_start, len(rows))
                                if row_bands[row_indices[r]][0] <= nw.cy <= row_bands[row_indices[r]][1]
                            ),
                            None,
                        )
                        if target is not None:
                            by_row.setdefault(target, []).append(nw)

                    for target, words in by_row.items():
                        words.sort(key=lambda w: w.x0)
                        combined = words[0].text
                        for prev, nw in zip(words, words[1:]):
                            gap = nw.x0 - prev.x1
                            combined += ("" if gap <= 24 else " ") + nw.text
                        fresh = _normalize_numeric_text(combined)
                        if " " in fresh or not _is_numeric_like(fresh):
                            continue
                        cell = rows[target][c].strip()
                        old_digits = sum(ch.isdigit() for ch in cell)
                        new_digits = sum(ch.isdigit() for ch in fresh)
                        if new_digits > old_digits:
                            rows[target][c] = fresh

                # إعادة قراءة العمود الرقمي كاملًا بمحرّك إنجليزي مكبّر:
                # أسرع من قراءة كل خلية على حدة، وأدق للأرقام الصغيرة/المقطوعة.
                if page.image is not None:
                    x0 = int(col_bands[group[0]][0]) + 3
                    x1 = int(col_bands[group[-1]][1]) - 3
                    y0 = int(row_bands[row_indices[data_start]][0]) + 3
                    y1 = int(row_bands[row_indices[-1]][1]) - 3
                    if x1 - x0 >= 8 and y1 - y0 >= 8:
                        strip = page.image.crop(
                            (max(0, x0), max(0, y0), min(page.width, x1), min(page.height, y1))
                        )
                        for rel_y, fresh in _ocr_numeric_strip(strip):
                            abs_y = y0 + rel_y
                            target = next(
                                (
                                    r
                                    for r in range(data_start, len(rows))
                                    if row_bands[row_indices[r]][0] <= abs_y <= row_bands[row_indices[r]][1]
                                ),
                                None,
                            )
                            if target is None:
                                continue
                            cell = rows[target][c].strip()
                            if cell and not re.search(r"\d\s+\d", cell):
                                continue
                            old_digits = sum(ch.isdigit() for ch in cell)
                            new_digits = sum(ch.isdigit() for ch in fresh)
                            if new_digits > old_digits:
                                rows[target][c] = fresh

    # أرقام القطع المفردة: OCR يتجاهلها أحيانًا إذا قرأ الخلية وحدها،
    # فنقرأ الصف كاملًا (سياق حقيقي) ونأخذ الرقم الواقع داخل عمود القطعة فقط.
    plot_col = None
    for c in range(len(rows[0])):
        header = rows[0][c] + (" " + rows[1][c] if is_real_sub_header and len(rows) > 1 else "")
        if "قطعة" in header and ("رقم" in header or "القطعة" in header):
            plot_col = c
            break
    if plot_col is not None and page.image is not None:
        px0 = col_bands[groups[plot_col][0]][0]
        px1 = col_bands[groups[plot_col][-1]][1]
        for r in range(data_start, len(rows)):
            if rows[r][plot_col].strip():
                continue
            if not any(rows[r][j].strip() for j in range(len(rows[r])) if j != plot_col):
                continue
            y0 = int(row_bands[row_indices[r]][0]) + 2
            y1 = int(row_bands[row_indices[r]][1]) - 2
            if y1 - y0 < 8:
                continue
            row_img = page.image.crop((0, max(0, y0), page.width, min(page.height, y1)))
            candidates = [
                text for cx, _cy, text in _ocr_row_numeric_words(row_img) if px0 <= cx <= px1
            ]
            if candidates:
                rows[r][plot_col] = max(candidates, key=lambda text: sum(ch.isdigit() for ch in text))

    # إصلاح ترقيم التسلسل المفقود في الخلية الأولى (الرقم 1)
    if len(rows) > 2:
        for c in range(len(rows[0])):
            hdr = rows[0][c]
            if "رقم" in hdr or "قطعة" in hdr:
                if not rows[1][c].strip() and rows[2][c].strip() in {"2", "٢"}:
                    rows[1][c] = "1"

    header_rows = 2 if is_real_sub_header and (any(header_bottom) or header_spans) else 1

    # حدود الصفوف بالبكسل (بنفس ترتيب rows)
    row_bounds = [row_bands[r] for r in row_indices]

    layout = PageLayout()
    layout.page_width = float(page.width)
    layout.page_height = float(page.height)
    layout.grid = Grid(
        rows=rows,
        n_cols=len(groups),
        col_bounds=[
            (col_bands[group[0]][0], col_bands[group[-1]][1]) for group in groups
        ],
        row_bounds=row_bounds,
        cell_heights=row_heights,
        header_rows=header_rows,
        header_spans=header_spans,
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
        return _join_cell_words(items).strip()

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
