"""
DocPass Backend API (FastAPI)
=============================
Accepts file conversion requests from the frontend, actually converts files
(Office -> PDF, Image -> PDF/Image, Markdown/Text -> PDF, PDF -> Images/PDF),
and provides two quick, portal-ready features:
  1. Portal Presets & Preflight Pass/Fail Validation Report
  2. Portal-Safe Filename Sanitization, A4 Normalization & Target-Size Compression

Run:
    uvicorn docpass_backend:app --host 0.0.0.0 --port 8000 --reload
"""

import asyncio
import io
import os
import re
import shutil
import subprocess
import tempfile
import uuid
import zipfile
from contextlib import asynccontextmanager
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, List, Literal

from docx import Document as DocxDocument
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from PIL import Image
from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

IST = timezone(timedelta(hours=5, minutes=30))
STORAGE_DIR = Path(os.getenv("DOCPASS_STORAGE_DIR", tempfile.gettempdir())) / "docpass_jobs"
STORAGE_DIR.mkdir(parents=True, exist_ok=True)

FILE_TTL_MINUTES = int(os.getenv("DOCPASS_TTL_MINUTES", "60"))
FREE_DAILY_LIMIT = int(os.getenv("DOCPASS_DAILY_LIMIT", "10"))
FREE_MAX_FILES = 3
FREE_MAX_MB = 50
A4_WIDTH_PT, A4_HEIGHT_PT = 595.28, 841.89

SAFE_FILENAME_RE = re.compile(r"^[A-Za-z0-9_-]+\.[A-Za-z0-9]+$")

PORTAL_PRESETS: dict[str, dict[str, Any]] = {
    "none": {
        "id": "none",
        "name": "None (Plain conversion)",
        "max_kb": None,
        "require_a4": False,
        "require_searchable_text": False,
        "sanitize_filename": False,
    },
    "govt_portal_200kb": {
        "id": "govt_portal_200kb",
        "name": "Govt Portal Upload (Under 200 KB)",
        "max_kb": 200,
        "require_a4": True,
        "require_searchable_text": True,
        "sanitize_filename": True,
    },
    "visa_500kb": {
        "id": "visa_500kb",
        "name": "Visa / Immigration (Under 500 KB)",
        "max_kb": 500,
        "require_a4": True,
        "require_searchable_text": True,
        "sanitize_filename": True,
    },
    "tender_2mb": {
        "id": "tender_2mb",
        "name": "e-Tender / Procurement (Under 2 MB)",
        "max_kb": 2048,
        "require_a4": True,
        "require_searchable_text": True,
        "sanitize_filename": True,
    },
}

JOBS: dict[str, dict[str, Any]] = {}
DAILY_COUNTERS: dict[str, dict[str, Any]] = {}


# ---------------------------------------------------------------------------
# Background Sweeper (deletes expired job files every 5 minutes)
# ---------------------------------------------------------------------------

def cleanup_expired_jobs() -> int:
    now = datetime.now(IST)
    expired_ids = [
        job_id
        for job_id, job in list(JOBS.items())
        if datetime.fromisoformat(job["expires_at"]) <= now
    ]
    for job_id in expired_ids:
        job_dir = STORAGE_DIR / job_id
        if job_dir.exists():
            shutil.rmtree(job_dir, ignore_errors=True)
        JOBS.pop(job_id, None)
    return len(expired_ids)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    async def _sweeper_loop():
        while True:
            await asyncio.sleep(300)
            cleanup_expired_jobs()

    task = asyncio.create_task(_sweeper_loop())
    try:
        yield
    finally:
        task.cancel()


app = FastAPI(
    title="DocPass Backend API",
    version="1.0.0",
    description="File conversion & portal-ready preflight API for DocPass.",
    lifespan=lifespan,
)

ALLOWED_ORIGINS = os.getenv("DOCPASS_CORS_ORIGINS", "*").split(",")
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)


# ---------------------------------------------------------------------------
# Helpers: Rate Limiting, Filename Sanitization & LibreOffice Discovery
# ---------------------------------------------------------------------------

def _today_ist_str() -> str:
    return datetime.now(IST).strftime("%Y-%m-%d")


def _next_midnight_ist_iso() -> str:
    now = datetime.now(IST)
    tomorrow = (now + timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    return tomorrow.isoformat()


def check_and_increment_daily_limit(client_key: str) -> dict[str, Any]:
    today = _today_ist_str()
    rec = DAILY_COUNTERS.get(client_key)
    if not rec or rec["date"] != today:
        rec = {"date": today, "used": 0}
        DAILY_COUNTERS[client_key] = rec

    if rec["used"] >= FREE_DAILY_LIMIT:
        raise HTTPException(
            status_code=429,
            detail={
                "error": "Daily free conversion limit reached (10/day, resets at midnight IST).",
                "used": rec["used"],
                "limit": FREE_DAILY_LIMIT,
                "resets_at": _next_midnight_ist_iso(),
            },
        )
    rec["used"] += 1
    return {
        "used": rec["used"],
        "remaining": FREE_DAILY_LIMIT - rec["used"],
        "limit": FREE_DAILY_LIMIT,
        "resets_at": _next_midnight_ist_iso(),
    }


def sanitize_portal_filename(original_name: str, target_ext: str) -> str:
    """Keep only letters, digits, _ and - in the stem."""
    stem = Path(original_name).stem
    stem = re.sub(r"[^A-Za-z0-9_-]+", "_", stem).strip("_-")
    if not stem:
        stem = "document"
    ext = target_ext.lstrip(".").lower()
    return f"{stem}.{ext}"


def _find_libreoffice() -> str | None:
    """Find LibreOffice / soffice binary across Linux, macOS, and Windows."""
    for name in ("libreoffice", "soffice"):
        found = shutil.which(name)
        if found:
            return found
    for win_path in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ):
        if Path(win_path).exists():
            return win_path
    return None


def _docx_to_pdf_fallback(input_path: Path, out_path: Path) -> None:
    """Pure-Python fallback for .docx -> PDF when LibreOffice is not installed."""
    doc = DocxDocument(str(input_path))
    styles = getSampleStyleSheet()
    pdf = SimpleDocTemplate(str(out_path), pagesize=A4)
    story: list[Any] = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            story.append(Spacer(1, 8))
            continue
        safe = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        style_name = para.style.name if para.style else ""
        if style_name.startswith("Heading 1"):
            story.append(Paragraph(safe, styles["Heading1"]))
        elif style_name.startswith("Heading 2"):
            story.append(Paragraph(safe, styles["Heading2"]))
        elif style_name.startswith("Heading"):
            story.append(Paragraph(safe, styles["Heading3"]))
        else:
            story.append(Paragraph(safe, styles["Normal"]))
        story.append(Spacer(1, 4))

    for tbl in doc.tables:
        data = [[cell.text.strip() for cell in row.cells] for row in tbl.rows]
        if data:
            t = Table(data)
            t.setStyle(
                TableStyle([
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                    ("FONTSIZE", (0, 0), (-1, -1), 9),
                ])
            )
            story.append(t)
            story.append(Spacer(1, 10))

    if not story:
        story.append(Paragraph(" ", styles["Normal"]))
    pdf.build(story)


# ---------------------------------------------------------------------------
# Feature 2: A4 Normalization & Target-Size Compression
# ---------------------------------------------------------------------------

def normalize_pdf_to_a4(pdf_path: Path) -> None:
    """Scale every page in-place to A4 (595.28 x 841.89 pt) while preserving aspect ratio."""
    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        if abs(w - A4_WIDTH_PT) <= 5 and abs(h - A4_HEIGHT_PT) <= 5:
            writer.add_page(page)
            continue
        scale = min(A4_WIDTH_PT / w, A4_HEIGHT_PT / h)
        tx = (A4_WIDTH_PT - w * scale) / 2
        ty = (A4_HEIGHT_PT - h * scale) / 2
        new_page = writer.add_blank_page(width=A4_WIDTH_PT, height=A4_HEIGHT_PT)
        new_page.merge_transformed_page(
            page, Transformation().scale(scale, scale).translate(tx, ty)
        )
    with open(pdf_path, "wb") as f:
        writer.write(f)


def compress_pdf_to_target_kb(pdf_path: Path, max_kb: int) -> None:
    """
    Attempt lossless stream compression first; if still over max_kb,
    re-render pages as compressed JPEG-backed PDF via pdftoppm + Pillow.
    """
    if pdf_path.stat().st_size <= max_kb * 1024:
        return

    reader = PdfReader(str(pdf_path))
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
        writer.pages[-1].compress_content_streams()
    with open(pdf_path, "wb") as f:
        writer.write(f)

    if pdf_path.stat().st_size <= max_kb * 1024:
        return

    if shutil.which("pdftoppm"):
        with tempfile.TemporaryDirectory(prefix="docpass_cmp_") as tmp:
            prefix = Path(tmp) / "page"
            subprocess.run(
                ["pdftoppm", "-jpeg", "-r", "110", str(pdf_path), str(prefix)],
                check=False,
                timeout=30,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            jpgs = sorted(Path(tmp).glob("page-*.jpg"))
            if jpgs:
                for quality in (65, 45, 30):
                    imgs: list[Image.Image] = []
                    for j in jpgs:
                        im = Image.open(j).convert("RGB")
                        buf = io.BytesIO()
                        im.save(buf, format="JPEG", quality=quality, optimize=True)
                        buf.seek(0)
                        imgs.append(Image.open(buf))
                    out_buf = io.BytesIO()
                    imgs[0].save(
                        out_buf,
                        format="PDF",
                        save_all=True,
                        append_images=imgs[1:],
                        resolution=100.0,
                    )
                    if len(out_buf.getvalue()) <= max_kb * 1024 or quality == 30:
                        pdf_path.write_bytes(out_buf.getvalue())
                        break


# ---------------------------------------------------------------------------
# Feature 1: Portal Preset Preflight Validation Report
# ---------------------------------------------------------------------------

def build_portal_report(output_path: Path, preset_id: str) -> dict[str, Any]:
    preset = PORTAL_PRESETS.get(preset_id, PORTAL_PRESETS["none"])
    size_bytes = output_path.stat().st_size
    size_kb = round(size_bytes / 1024, 1)
    filename = output_path.name
    checks: list[dict[str, Any]] = []

    max_kb = preset["max_kb"]
    if max_kb is not None:
        passed = size_kb <= max_kb
        checks.append({
            "check": "file_size",
            "label": "Size",
            "detail": f"{size_kb} KB of {max_kb} KB max",
            "status": "Pass" if passed else "Fail",
            "passed": passed,
        })
    else:
        checks.append({
            "check": "file_size",
            "label": "Size",
            "detail": f"{size_kb} KB",
            "status": "Pass",
            "passed": True,
        })

    if output_path.suffix.lower() == ".pdf":
        reader = PdfReader(str(output_path))
        total_pages = len(reader.pages)
        a4_ok = True
        pages_with_text = 0
        for p in reader.pages:
            w, h = float(p.mediabox.width), float(p.mediabox.height)
            if not (abs(w - A4_WIDTH_PT) <= 8 and abs(h - A4_HEIGHT_PT) <= 8):
                a4_ok = False
            if (p.extract_text() or "").strip():
                pages_with_text += 1

        if preset["require_a4"]:
            checks.append({
                "check": "page_size",
                "label": "Page size",
                "detail": "A4" if a4_ok else "Non-A4",
                "status": "Pass" if a4_ok else "Fail",
                "passed": a4_ok,
            })

        if preset["require_searchable_text"]:
            text_ok = total_pages > 0 and pages_with_text == total_pages
            checks.append({
                "check": "searchable_text",
                "label": "Searchable text",
                "detail": (
                    f"Found on all {total_pages} pages"
                    if text_ok
                    else f"Found on {pages_with_text} of {total_pages} pages"
                ),
                "status": "Pass" if text_ok else "Fail",
                "passed": text_ok,
            })

    name_ok = bool(SAFE_FILENAME_RE.match(filename))
    checks.append({
        "check": "filename",
        "label": "File name",
        "detail": filename,
        "status": "Pass" if (name_ok or not preset["sanitize_filename"]) else "Fail",
        "passed": name_ok or not preset["sanitize_filename"],
    })

    ready = all(c["passed"] for c in checks)
    return {
        "preset_id": preset["id"],
        "preset_name": preset["name"],
        "ready_for_portal": ready,
        "checks": checks,
    }


# ---------------------------------------------------------------------------
# Core Conversion Engine
# ---------------------------------------------------------------------------

OFFICE_EXTS = {".doc", ".docx", ".odt", ".rtf", ".xls", ".xlsx", ".ods", ".ppt", ".pptx", ".odp"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff", ".tif"}
TEXT_EXTS = {".txt", ".md", ".csv", ".log"}
HTML_EXTS = {".html", ".htm"}


class _VisibleTextHTMLParser(HTMLParser):
    SKIP_TAGS = {"script", "style", "noscript", "svg", "template"}
    BLOCK_TAGS = {
        "p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6",
        "section", "article", "header", "footer", "blockquote", "pre", "hr",
    }

    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.skip_depth = 0
        self.in_title = False
        self.title_parts: list[str] = []
        self.chunks: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        t = tag.lower()
        if t in self.SKIP_TAGS:
            self.skip_depth += 1
        elif t == "title":
            self.in_title = True
        elif t in self.BLOCK_TAGS and self.skip_depth == 0:
            self.chunks.append("\n")

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in self.SKIP_TAGS and self.skip_depth > 0:
            self.skip_depth -= 1
        elif t == "title":
            self.in_title = False
        elif t in self.BLOCK_TAGS and self.skip_depth == 0:
            self.chunks.append("\n")

    def handle_data(self, data: str) -> None:
        if self.in_title:
            self.title_parts.append(data)
        elif self.skip_depth == 0:
            self.chunks.append(data)


def _xml_safe_text(text: str) -> str:
    # Break very long unbroken tokens so ReportLab Paragraph never hits LayoutError
    parts: list[str] = []
    for tok in text.split(" "):
        if len(tok) > 90:
            parts.extend(tok[i:i + 90] for i in range(0, len(tok), 90))
        else:
            parts.append(tok)
    joined = " ".join(parts)
    return joined.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _html_to_pdf(input_path: Path, out_path: Path) -> None:
    raw_html = input_path.read_text(encoding="utf-8", errors="replace")
    parser = _VisibleTextHTMLParser()
    parser.feed(raw_html)
    parser.close()

    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(str(out_path), pagesize=A4)
    story: list[Any] = []

    title = " ".join("".join(parser.title_parts).split()).strip()
    if title:
        story.append(Paragraph(_xml_safe_text(title), styles["Heading1"]))
        story.append(Spacer(1, 10))

    raw_text = "".join(parser.chunks)
    lines = [" ".join(line.split()) for line in raw_text.splitlines()]
    lines = [line for line in lines if line]

    for line in lines[:1500]:
        story.append(Paragraph(_xml_safe_text(line), styles["Normal"]))
        story.append(Spacer(1, 4))

    if not story:
        story.append(Paragraph(_xml_safe_text(f"HTML Document: {input_path.name}"), styles["Heading2"]))
        story.append(Spacer(1, 6))
        story.append(
            Paragraph(
                "No static visible text was found (content may be rendered dynamically by JavaScript).",
                styles["Normal"],
            )
        )
    doc.build(story)



def convert_single_file(
    input_path: Path,
    target_format: str,
    output_dir: Path,
    preset_id: str,
) -> Path:
    preset = PORTAL_PRESETS.get(preset_id, PORTAL_PRESETS["none"])
    src_ext = input_path.suffix.lower()
    tgt_ext = target_format.lower().lstrip(".")

    out_name = (
        sanitize_portal_filename(input_path.name, tgt_ext)
        if preset["sanitize_filename"]
        else f"{input_path.stem}.{tgt_ext}"
    )
    out_path = output_dir / out_name

    # --- Case 1: Office -> PDF (LibreOffice headless, or python-docx fallback for .docx) ---
    if src_ext in OFFICE_EXTS and tgt_ext == "pdf":
        lo_bin = _find_libreoffice()
        if lo_bin:
            tmp_root = tempfile.gettempdir()
            with tempfile.TemporaryDirectory(prefix="lo_profile_") as lo_profile:
                env = {**os.environ}
                if os.name != "nt":
                    env["HOME"] = tmp_root
                    env["TMPDIR"] = tmp_root
                res = subprocess.run(
                    [
                        lo_bin,
                        "--headless",
                        f"-env:UserInstallation={Path(lo_profile).as_uri()}",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        str(output_dir),
                        str(input_path),
                    ],
                    env=env,
                    timeout=45,
                    capture_output=True,
                    text=True,
                )
                generated = output_dir / f"{input_path.stem}.pdf"
                if res.returncode != 0 or not generated.exists() or generated.stat().st_size == 0:
                    raise RuntimeError(f"LibreOffice conversion failed: {res.stderr or res.stdout}")
                if generated != out_path:
                    generated.rename(out_path)
        elif src_ext == ".docx":
            _docx_to_pdf_fallback(input_path, out_path)
        else:
            raise HTTPException(
                status_code=500,
                detail=f"LibreOffice is required to convert '{src_ext}' files to PDF.",
            )

    # --- Case 2: Image -> PDF or Image -> Image ---
    elif src_ext in IMAGE_EXTS:
        img = Image.open(input_path)
        if tgt_ext == "pdf":
            rgb = img.convert("RGB")
            rgb.save(out_path, format="PDF", resolution=150.0)
        elif tgt_ext in {"png", "jpg", "jpeg", "webp"}:
            fmt = "JPEG" if tgt_ext in {"jpg", "jpeg"} else tgt_ext.upper()
            save_img = img.convert("RGB") if fmt == "JPEG" else img
            save_img.save(out_path, format=fmt, quality=90)
        else:
            raise HTTPException(status_code=400, detail=f"Unsupported target format '{tgt_ext}' for image.")

    # --- Case 3: Text / Markdown / CSV -> PDF ---
    elif src_ext in TEXT_EXTS and tgt_ext == "pdf":
        raw_text = input_path.read_text(encoding="utf-8", errors="replace")
        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(str(out_path), pagesize=A4)
        story: list[Any] = []
        for line in raw_text.splitlines():
            clean = line.strip()
            if not clean:
                story.append(Spacer(1, 8))
            else:
                story.append(Paragraph(_xml_safe_text(clean), styles["Normal"]))
        if not story:
            story.append(Paragraph(" ", styles["Normal"]))
        doc.build(story)

    # --- Case 3b: HTML -> PDF ---
    elif src_ext in HTML_EXTS and tgt_ext == "pdf":
        _html_to_pdf(input_path, out_path)

    # --- Case 4: PDF -> PDF (compress / A4 normalize / sanitize) ---
    elif src_ext == ".pdf" and tgt_ext == "pdf":
        reader = PdfReader(str(input_path))
        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)
            writer.pages[-1].compress_content_streams()
        with open(out_path, "wb") as f:
            writer.write(f)

    # --- Case 5: PDF -> PNG/JPEG/WEBP (ZIP of pages or single image) ---
    elif src_ext == ".pdf" and tgt_ext in {"png", "jpg", "jpeg", "webp"}:
        if not shutil.which("pdftoppm"):
            raise HTTPException(status_code=500, detail="pdftoppm (poppler) is not installed on server.")
        with tempfile.TemporaryDirectory(prefix="docpass_img_") as tmp:
            flag = "-png" if tgt_ext in {"png", "webp"} else "-jpeg"
            subprocess.run(
                ["pdftoppm", flag, "-r", "150", str(input_path), str(Path(tmp) / "page")],
                check=True,
                timeout=45,
            )
            pages = sorted(Path(tmp).glob("page-*"))
            if not pages:
                raise RuntimeError("No pages rendered from PDF.")
            if tgt_ext == "webp":
                webp_pages: list[Path] = []
                for pg in pages:
                    wp = pg.with_suffix(".webp")
                    Image.open(pg).convert("RGB").save(wp, format="WEBP", quality=85)
                    webp_pages.append(wp)
                pages = webp_pages
            if len(pages) == 1:
                shutil.copy2(pages[0], out_path)
            else:
                out_path = out_path.with_suffix(".zip")
                with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
                    for p in pages:
                        zf.write(p, arcname=p.name)
    else:
        raise HTTPException(
            status_code=400,
            detail=f"Conversion from '{src_ext}' to '{tgt_ext}' is not supported.",
        )

    if out_path.suffix.lower() == ".pdf":
        if preset["require_a4"]:
            normalize_pdf_to_a4(out_path)
        if preset["max_kb"]:
            compress_pdf_to_target_kb(out_path, preset["max_kb"])

    return out_path


# ---------------------------------------------------------------------------
# Built-in Test Frontend (GET /) & API Endpoints
# ---------------------------------------------------------------------------

TEST_UI_HTML = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>DocPass — Convert & Portal Preflight</title>
  <style>
    body{font-family:system-ui,-apple-system,sans-serif;max-width:680px;margin:40px auto;padding:0 16px;background:#faf9f5;color:#1f1e1d}
    .card{background:#fff;border:1px solid #e5e4df;border-radius:12px;padding:24px;box-shadow:0 1px 3px rgba(0,0,0,.06)}
    h1{margin:0 0 6px;font-size:22px} p.sub{margin:0 0 20px;color:#6b6a68;font-size:14px}
    label{display:block;font-size:13px;font-weight:600;margin:14px 0 6px}
    input[type=file],select{width:100%;padding:10px;border:1px solid #d5d4cf;border-radius:8px;background:#fff;box-sizing:border-box}
    button{margin-top:18px;width:100%;padding:12px;border:0;border-radius:8px;background:#1f1e1d;color:#fff;font-size:15px;font-weight:600;cursor:pointer}
    button:disabled{opacity:.5;cursor:wait}
    .res{margin-top:20px;display:flex;flex-direction:column;gap:12px}
    .item{border:1px solid #e5e4df;border-radius:10px;padding:14px;background:#fcfcfa}
    .badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:12px;font-weight:600}
    .pass{background:#e6f4ea;color:#137333} .fail{background:#fce8e6;color:#c5221f}
    ul{margin:8px 0 12px;padding-left:18px;font-size:13px}
    a.dl{display:inline-block;padding:8px 14px;background:#0b57d0;color:#fff;text-decoration:none;border-radius:6px;font-size:13px;font-weight:600}
  </style>
</head>
<body>
  <div class="card">
    <h1>DocPass Converter</h1>
    <p class="sub">Upload up to 3 files (.docx, .png, .jpg, .md, .txt, .csv, .html, .pdf) and test portal-ready conversion.</p>
    <label>Choose files (max 3)</label>
    <input id="files" type="file" multiple />
    <label>Target format</label>
    <select id="fmt">
      <option value="pdf">PDF</option>
      <option value="png">PNG</option>
      <option value="jpg">JPG</option>
      <option value="webp">WEBP</option>
    </select>
    <label>Portal preset</label>
    <select id="preset">
      <option value="govt_portal_200kb">Govt Portal Upload (Under 200 KB, A4, Searchable, Safe Name)</option>
      <option value="visa_500kb">Visa / Immigration (Under 500 KB)</option>
      <option value="tender_2mb">e-Tender / Procurement (Under 2 MB)</option>
      <option value="none">None (Plain conversion)</option>
    </select>
    <button id="btn">Convert</button>
    <div id="out" class="res"></div>
  </div>
  <script>
    const btn = document.getElementById('btn'), out = document.getElementById('out');
    btn.onclick = async () => {
      const files = document.getElementById('files').files;
      if (!files.length) return alert('Pick at least 1 file');
      btn.disabled = true; btn.textContent = 'Converting...'; out.innerHTML = '';
      const fd = new FormData();
      for (const f of files) fd.append('files', f);
      fd.append('target_format', document.getElementById('fmt').value);
      fd.append('preset', document.getElementById('preset').value);
      try {
        const r = await fetch('/api/convert', { method: 'POST', body: fd });
        const data = await r.json();
        if (!r.ok) throw new Error(JSON.stringify(data.detail || data));
        out.innerHTML = data.files.map(f => {
          const ok = f.portal_report.ready_for_portal;
          const checks = f.portal_report.checks.map(c =>
            `<li><b>${c.label}:</b> ${c.detail} <span class="badge ${c.passed?'pass':'fail'}">${c.status}</span></li>`
          ).join('');
          return `<div class="item">
            <div><b>${f.output_filename}</b> (${f.size_kb} KB)
              <span class="badge ${ok?'pass':'fail'}">${ok ? 'Ready for Portal' : 'Needs Review'}</span>
            </div>
            <ul>${checks}</ul>
            <a class="dl" href="${f.download_url}" download>Download ${f.output_filename}</a>
          </div>`;
        }).join('');
      } catch (e) {
        out.innerHTML = `<div class="item" style="color:#c5221f">Error: ${e.message}</div>`;
      } finally {
        btn.disabled = false; btn.textContent = 'Convert';
      }
    };
  </script>
</body>
</html>"""


APP_DIR = Path(__file__).resolve().parent


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    idx = APP_DIR / "index.html"
    if idx.exists():
        return HTMLResponse(idx.read_text(encoding="utf-8"))
    return HTMLResponse(TEST_UI_HTML)


@app.get("/style.css")
def serve_css() -> FileResponse:
    css = APP_DIR / "style.css"
    if not css.exists():
        raise HTTPException(status_code=404, detail="style.css not found")
    return FileResponse(css, media_type="text/css")


@app.get("/script.js")
def serve_js() -> FileResponse:
    js = APP_DIR / "script.js"
    if not js.exists():
        raise HTTPException(status_code=404, detail="script.js not found")
    return FileResponse(js, media_type="application/javascript")


@app.get("/health")
def health() -> dict[str, Any]:
    return {
        "status": "ok",
        "service": "docpass-backend",
        "libreoffice_found": bool(_find_libreoffice()),
        "time_ist": datetime.now(IST).isoformat(),
    }


@app.get("/api/presets")
def list_presets() -> dict[str, Any]:
    return {"presets": list(PORTAL_PRESETS.values())}


@app.get("/api/limits")
def get_limits(request: Request) -> dict[str, Any]:
    client_key = request.client.host if request.client else "anonymous"
    today = _today_ist_str()
    rec = DAILY_COUNTERS.get(client_key)
    used = rec["used"] if rec and rec["date"] == today else 0
    return {
        "plan": "free",
        "max_files_per_run": FREE_MAX_FILES,
        "max_file_mb": FREE_MAX_MB,
        "daily_limit": FREE_DAILY_LIMIT,
        "used_today": used,
        "remaining_today": max(0, FREE_DAILY_LIMIT - used),
        "resets_at": _next_midnight_ist_iso(),
    }


@app.post("/api/convert")
async def convert_files(
    request: Request,
    files: List[UploadFile] = File(...),
    target_format: Literal["pdf", "png", "jpg", "jpeg", "webp"] = Form("pdf"),
    preset: str = Form("none"),
) -> JSONResponse:
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded.")
    if len(files) > FREE_MAX_FILES:
        raise HTTPException(
            status_code=400,
            detail=f"Free plan allows up to {FREE_MAX_FILES} files per run.",
        )
    if preset not in PORTAL_PRESETS:
        raise HTTPException(status_code=400, detail=f"Unknown portal preset '{preset}'.")

    client_key = request.client.host if request.client else "anonymous"
    quota = check_and_increment_daily_limit(client_key)

    job_id = uuid.uuid4().hex[:12]
    job_dir = STORAGE_DIR / job_id
    in_dir = job_dir / "input"
    out_dir = job_dir / "output"
    in_dir.mkdir(parents=True, exist_ok=True)
    out_dir.mkdir(parents=True, exist_ok=True)

    now = datetime.now(IST)
    expires_at = now + timedelta(minutes=FILE_TTL_MINUTES)

    converted_items: list[dict[str, Any]] = []
    try:
        for idx, upload in enumerate(files):
            raw_bytes = await upload.read()
            if len(raw_bytes) > FREE_MAX_MB * 1024 * 1024:
                raise HTTPException(
                    status_code=413,
                    detail=f"File '{upload.filename}' exceeds {FREE_MAX_MB} MB limit.",
                )
            safe_in_name = Path(upload.filename or f"file_{idx}").name
            in_path = in_dir / safe_in_name
            in_path.write_bytes(raw_bytes)

            out_path = convert_single_file(in_path, target_format, out_dir, preset)
            report = build_portal_report(out_path, preset)

            converted_items.append({
                "index": idx,
                "original_filename": safe_in_name,
                "output_filename": out_path.name,
                "size_kb": round(out_path.stat().st_size / 1024, 1),
                "download_url": f"/api/jobs/{job_id}/download/{idx}",
                "output_path": str(out_path),
                "portal_report": report,
            })
    except HTTPException:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(job_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail=f"Conversion error: {exc}") from exc

    job_record = {
        "job_id": job_id,
        "status": "completed",
        "preset": preset,
        "target_format": target_format,
        "created_at": now.isoformat(),
        "expires_at": expires_at.isoformat(),
        "ttl_minutes": FILE_TTL_MINUTES,
        "quota": quota,
        "files": [
            {k: v for k, v in item.items() if k != "output_path"}
            for item in converted_items
        ],
        "_paths": {item["index"]: item["output_path"] for item in converted_items},
    }
    JOBS[job_id] = job_record

    return JSONResponse(
        status_code=200,
        content={k: v for k, v in job_record.items() if k != "_paths"},
    )


@app.get("/api/jobs/{job_id}")
def get_job_status(job_id: str) -> dict[str, Any]:
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found or expired.")
    return {k: v for k, v in job.items() if k != "_paths"}


@app.get("/api/jobs/{job_id}/download/{file_index}")
def download_converted_file(job_id: str, file_index: int) -> FileResponse:
    job = JOBS.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail="Job not found or expired.")
    path_str = job["_paths"].get(file_index)
    if not path_str or not Path(path_str).exists():
        raise HTTPException(status_code=404, detail="Converted file not found.")
    file_meta = job["files"][file_index]
    return FileResponse(
        path=path_str,
        filename=file_meta["output_filename"],
        media_type="application/octet-stream",
    )