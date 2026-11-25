"""
محرّك الذكاء الاصطناعي — التحويل فائق الدقة عبر Google Gemini Vision
====================================================================
يحلل صفحات PDF والصور بدقة 100% ويستخرج الجداول والنصوص والهوامش
بدون أي تشويه للحروف أو تداخل في الأعمدة.
"""

from __future__ import annotations

import base64
import io
import json
import logging
import os
import re
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from PIL import Image

import ocr_engine
from ocr_engine import Grid, PageLayout

logger = logging.getLogger("muhawwil.ai")


class AiEngineError(RuntimeError):
    """خطأ من محرك Gemini يمكن عرضه مباشرةً للمستخدم."""


GEMINI_MODELS = [
    "gemini-2.0-flash",
    "gemini-1.5-flash",
]

PROMPT = """You are an expert document OCR and layout preservation engine.
Analyze this document image with extreme precision.
Transcribe all text accurately (preserving Arabic, English, numbers, symbols, and dates).

Return a valid JSON object matching this schema:
{
  "headers": [
    [{"text": "...", "align": "right|center|left"}]
  ],
  "table": {
    "headers": [["col1", "col2", ...]],
    "rows": [["val1", "val2", ...]]
  },
  "paragraphs": ["..."],
  "footers": [
    [{"text": "...", "align": "right|center|left"}]
  ]
}

Strict Rules:
1. For Arabic / RTL documents: the columns MUST be ordered from Right to Left (the right-most column in the image is index 0 in the JSON list).
2. NEVER merge data rows into headers! The headers array must only contain actual column titles (e.g. رقم القطعة, المساحة, اسماء اصحاب الحقوق).
3. Transcribe national ID numbers, area values, and amounts with 100% exact digits.
4. Do NOT drop empty cells in a data row; represent empty cells as empty string "".
5. In headers, group horizontal elements on the same row into the same inner list.
"""


def get_api_key(custom_key: str | None = None) -> str | None:
    """يستخرج مفتاح Gemini من المعامل أو متغير البيئة أو ملف .env"""
    if custom_key and custom_key.strip():
        return custom_key.strip()

    # فحص متغير البيئة
    env_key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if env_key and env_key.strip():
        return env_key.strip()

    # فحص ملف .env في مسار المشروع
    env_file = Path(__file__).resolve().parent / ".env"
    if env_file.exists():
        try:
            for line in env_file.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if line.startswith("GEMINI_API_KEY=") or line.startswith("GOOGLE_API_KEY="):
                    val = line.split("=", 1)[1].strip().strip("\"'")
                    if val:
                        return val
        except Exception:
            pass
    return None


def is_available(custom_key: str | None = None) -> bool:
    return bool(get_api_key(custom_key))


def _parse_gemini_json(raw_text: str) -> dict[str, Any]:
    """يتحمل استجابة JSON خام أو ملفوفة داخل Markdown code fence."""
    text = (raw_text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start < 0 or end <= start:
            raise
        data = json.loads(text[start : end + 1])
    if not isinstance(data, dict):
        raise ValueError("استجابة Gemini ليست كائن JSON صالحًا.")
    return data


def _call_gemini_vision(img_bytes: bytes, api_key: str) -> dict[str, Any]:
    """يطلب تحليل الصورة من Gemini API ويسترجع كائن JSON المهيكل."""
    b64_data = base64.b64encode(img_bytes).decode("utf-8")

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": PROMPT},
                    {
                        "inline_data": {
                            "mime_type": "image/png",
                            "data": b64_data,
                        }
                    },
                ]
            }
        ],
        "generationConfig": {
            "response_mime_type": "application/json",
            "temperature": 0.05,
        },
    }

    req_data = json.dumps(payload).encode("utf-8")
    last_err = None

    for model in GEMINI_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        req = urllib.request.Request(
            url,
            data=req_data,
            headers={
                "Content-Type": "application/json",
                "User-Agent": "Muhawwil-AI/1.0",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=45) as resp:
                resp_json = json.loads(resp.read().decode("utf-8"))
                candidates = resp_json.get("candidates", [])
                if not candidates:
                    continue
                content_parts = candidates[0].get("content", {}).get("parts", [])
                if not content_parts:
                    continue
                raw_text = content_parts[0].get("text", "").strip()
                return _parse_gemini_json(raw_text)
        except urllib.error.HTTPError as err:
            err_msg = err.read().decode("utf-8", errors="ignore")
            logger.warning("Gemini error (%s) on model %s: %s", err.code, model, err_msg)
            if err.code in (400, 401, 403) and ("api key" in err_msg.lower() or "api_key" in err_msg.lower()):
                raise AiEngineError("مفتاح Gemini API غير صالح. تحقق من المفتاح أو أنشئ مفتاحًا جديدًا.")
            if err.code == 429:
                raise AiEngineError("تم تجاوز حد استخدام Gemini المجاني. حاول لاحقًا.")
            last_err = AiEngineError(f"خطأ من خدمة Gemini ({err.code}). حاول لاحقًا.")
        except AiEngineError:
            raise
        except Exception as exc:
            logger.warning("Gemini request failed on model %s: %s", model, exc)
            last_err = AiEngineError(f"تعذّر الاتصال بخدمة الذكاء الاصطناعي: {exc}")

    raise last_err or AiEngineError("تعذّر الاتصال بخدمة الذكاء الاصطناعي.")


def analyze_page_ai(img: Image.Image, api_key: str | None = None) -> PageLayout:
    """يحول صورة الصفحة عبر Gemini Vision إلى PageLayout بنفس صيغة المحرك المحلي."""
    key = get_api_key(api_key)
    if not key:
        raise ValueError("مفتاح Gemini API غير متوفر. الرجاء إدخال المفتاح في الإعدادات.")

    # تصدير الصورة إلى بايتات PNG
    buf = io.BytesIO()
    if img.mode not in ("RGB", "L"):
        img = img.convert("RGB")
    img.save(buf, format="PNG")
    img_bytes = buf.getvalue()

    data = _call_gemini_vision(img_bytes, key)

    layout = PageLayout()
    layout.page_width = float(img.width)
    layout.page_height = float(img.height)

    # 1. الترويسة (Headers)
    raw_headers = data.get("headers") or []
    for line in raw_headers:
        if isinstance(line, list):
            for item in line:
                if isinstance(item, dict):
                    t = item.get("text", "").strip()
                    a = item.get("align", "right").lower()
                    if a not in {"right", "center", "left"}:
                        a = "right"
                    if t:
                        layout.before.append((t, a))
        elif isinstance(line, dict):
            t = line.get("text", "").strip()
            a = line.get("align", "right").lower()
            if t:
                layout.before.append((t, a))

    # 2. الجداول (Tables)
    table_data = data.get("table") or {}
    table_headers = table_data.get("headers") or []
    table_rows = table_data.get("rows") or []

    all_rows: list[list[str]] = []
    for h in table_headers:
        if isinstance(h, list) and any(str(c).strip() for c in h):
            all_rows.append([str(c).strip() for c in h])
    for r in table_rows:
        if isinstance(r, list) and any(str(c).strip() for c in r):
            all_rows.append([str(c).strip() for c in r])

    if all_rows:
        n_cols = max(len(r) for r in all_rows)
        # توحيد أطوال الصفوف
        padded_rows = [r + [""] * (n_cols - len(r)) for r in all_rows]
        header_count = max(1, len(table_headers)) if table_headers else 1

        layout.grid = Grid(
            rows=padded_rows,
            n_cols=n_cols,
            header_rows=header_count,
        )

    # 3. الفقرات الحرة (Paragraphs)
    layout.paragraphs = [str(p).strip() for p in (data.get("paragraphs") or []) if str(p).strip()]

    # 4. التذييل (Footers)
    raw_footers = data.get("footers") or []
    for item in raw_footers:
        if isinstance(item, list):
            for f in item:
                if isinstance(f, dict):
                    t = f.get("text", "").strip()
                    a = f.get("align", "center").lower()
                    if t:
                        layout.after.append((t, a))
        elif isinstance(item, dict):
            t = item.get("text", "").strip()
            a = item.get("align", "center").lower()
            if t:
                layout.after.append((t, a))

    return layout
