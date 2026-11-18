"""
محوّل — خدمة الويب (FastAPI)
============================
• GET  /                واجهة التطبيق
• POST /api/convert     استقبال الملفات + الصيغة الهدف وإرجاع الناتج
• GET  /api/health      فحص الجاهزية
"""

from __future__ import annotations

import re
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask
from starlette.concurrency import run_in_threadpool

import converters

BASE = Path(__file__).resolve().parent
STATIC = BASE / "static"
WORK_ROOT = BASE / "work"
WORK_ROOT.mkdir(exist_ok=True)

MAX_TOTAL_BYTES = 120 * 1024 * 1024  # 120 ميجابايت إجمالًا
MAX_FILES = 30
ALLOWED_EXTS = converters.PDF_EXTS | converters.IMAGE_EXTS

app = FastAPI(title="محوّل — PDF & صور إلى Office", docs_url=None, redoc_url=None)


@app.get("/api/health")
async def health():
    return {"status": "ok"}


def _safe_stem(name: str) -> str:
    stem = Path(name or "ملف").stem
    cleaned = "".join(ch for ch in stem if ch.isalnum() or ch in " _-()[]").strip()
    return (cleaned or "ملف")[:60]


def _strip_prefix(path: Path) -> Path:
    """يزيل بادئة الترتيب (00_) التي أُضيفت لاسم الملف المؤقت."""
    clean = re.sub(r"^\d{2}_", "", path.name)
    if clean and clean != path.name:
        target = path.with_name(clean)
        try:
            path.rename(target)
            return target
        except OSError:
            return path
    return path


def _run(paths: list[Path], target: str, job_dir: Path, use_ocr: bool = True):
    """ينفّذ التحويل داخل خيط منفصل (عمليات معالجة ثقيلة)."""
    out_dir = job_dir / "out"
    out_dir.mkdir(exist_ok=True)
    work_dir = job_dir / "tmp"
    work_dir.mkdir(exist_ok=True)

    images = [p for p in paths if converters.is_image(p)]

    # عدة صور فقط -> تُدمج في ملف واحد (كل صورة صفحة/شريحة/صف)
    if len(paths) > 1 and len(images) == len(paths):
        dst, meta = converters.convert(paths, target, out_dir, work_dir, use_ocr=use_ocr)
        return _strip_prefix(dst), {**meta, "files": len(paths), "zipped": False}

    # ملف واحد
    if len(paths) == 1:
        dst, meta = converters.convert(paths, target, out_dir, work_dir, use_ocr=use_ocr)
        return _strip_prefix(dst), {**meta, "files": 1, "zipped": False}

    # خلط PDF وصور (أو عدة PDF) -> ملف لكل مدخل ثم ZIP
    produced: list[Path] = []
    for src in paths:
        dst, _ = converters.convert([src], target, out_dir, work_dir, use_ocr=use_ocr)
        produced.append(dst)

    zip_path = out_dir / f"محوّل_{target}_{len(produced)}_ملفات.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in produced:
            zf.write(item, arcname=item.name)
            item.unlink(missing_ok=True)
    return zip_path, {"files": len(produced), "zipped": True}


@app.post("/api/convert")
async def convert(
    target: str = Form(...),
    files: list[UploadFile] = File(...),
    ocr: str = Form("1"),
):
    target = (target or "").lower().strip()
    use_ocr = str(ocr).lower() not in {"0", "false", "off", "no"}
    if target not in converters.TARGET_EXT:
        raise HTTPException(400, "صيغة الهدف غير مدعومة.")
    if not files:
        raise HTTPException(400, "لم يتم اختيار أي ملف.")
    if len(files) > MAX_FILES:
        raise HTTPException(400, f"الحد الأقصى {MAX_FILES} ملفًا في المرة الواحدة.")

    job_dir = Path(tempfile.mkdtemp(prefix="muhawwil_", dir=str(WORK_ROOT)))
    cleanup = BackgroundTask(shutil.rmtree, str(job_dir), ignore_errors=True)

    saved: list[Path] = []
    total = 0
    try:
        for index, upload in enumerate(files):
            ext = Path(upload.filename or "").suffix.lower()
            if ext not in ALLOWED_EXTS:
                raise HTTPException(400, f"صيغة غير مدعومة: {ext or upload.filename}")
            data = await upload.read()
            total += len(data)
            if total > MAX_TOTAL_BYTES:
                raise HTTPException(413, "الحجم الإجمالي يتجاوز 120 ميجابايت.")
            if not data:
                raise HTTPException(400, f"الملف فارغ: {upload.filename}")

            stem = _safe_stem(upload.filename or f"file{index}")
            dest = job_dir / f"{index:02d}_{stem}{ext}"
            dest.write_bytes(data)
            saved.append(dest)

        try:
            out_path, meta = await run_in_threadpool(_run, saved, target, job_dir, use_ocr)
        except HTTPException:
            raise
        except ValueError as exc:
            raise HTTPException(400, str(exc))
        except Exception as exc:  # noqa: BLE001
            raise HTTPException(
                500, f"تعذّر إكمال التحويل ({type(exc).__name__}). تأكد أن الملف سليم وغير محمي بكلمة مرور."
            )

        if not out_path.exists() or out_path.stat().st_size == 0:
            raise HTTPException(500, "الناتج فارغ — قد يكون الملف تالفًا أو محميًا.")

        headers = {
            "X-Output-Name": quote(out_path.name),
            "X-Meta": str(meta.get("files", 1)),
            "Cache-Control": "no-store",
        }
        return FileResponse(
            path=str(out_path),
            filename=out_path.name,
            media_type="application/octet-stream",
            headers=headers,
            background=cleanup,
        )
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise


@app.exception_handler(HTTPException)
async def http_exc_handler(request, exc: HTTPException):
    return JSONResponse({"detail": exc.detail}, status_code=exc.status_code)


app.mount("/", StaticFiles(directory=str(STATIC), html=True), name="static")


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(app, host="127.0.0.1", port=8000, log_level="info")
