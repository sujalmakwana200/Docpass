"""
End-to-end verification tests for DocPass Backend & Frontend routes.
Works with both `python test_docpass_backend.py` and `pytest`.
"""

import io
import sys
import tempfile
from pathlib import Path

from docx import Document
from fastapi.testclient import TestClient
from PIL import Image
from pypdf import PdfReader
from reportlab.pdfgen import canvas

sys.path.insert(0, str(Path(__file__).parent))
import docpass_backend as backend  # noqa: E402

client = TestClient(backend.app)


def make_sample_docx(path: Path) -> None:
    doc = Document()
    doc.add_heading("Invoice - March", level=1)
    doc.add_paragraph("Prepared for DocPass Portal Preset verification.")
    doc.add_paragraph("Total Amount: Rs. 12,500")
    doc.save(str(path))


def make_sample_png(path: Path) -> None:
    img = Image.new("RGB", (800, 1000), color=(240, 245, 250))
    img.save(str(path), format="PNG")


def make_sample_pdf_bytes() -> bytes:
    buf = io.BytesIO()
    c = canvas.Canvas(buf)
    c.drawString(72, 720, "DocPass PDF verification page")
    c.save()
    return buf.getvalue()


def test_cors_and_presets() -> None:
    cors_resp = client.options(
        "/api/convert",
        headers={
            "Origin": "http://localhost:3000",
            "Access-Control-Request-Method": "POST",
        },
    )
    assert cors_resp.status_code == 200, cors_resp.text
    assert cors_resp.headers.get("access-control-allow-origin") in ("*", "http://localhost:3000")

    presets_resp = client.get("/api/presets")
    assert presets_resp.status_code == 200
    preset_ids = {p["id"] for p in presets_resp.json()["presets"]}
    assert {"none", "govt_portal_200kb", "visa_500kb", "tender_2mb"}.issubset(preset_ids)
    print("[PASS] CORS preflight & portal presets verified:", sorted(preset_ids))


def test_frontend_static_routes() -> None:
    for path, ctype in [("/", "text/html"), ("/style.css", "text/css"), ("/script.js", "javascript")]:
        resp = client.get(path)
        assert resp.status_code == 200, f"{path} returned {resp.status_code}"
        assert ctype in resp.headers.get("content-type", "")
    print("[PASS] Frontend routes (/, /style.css, /script.js) returned 200 OK")


def test_docx_and_png_conversion_with_preset() -> None:
    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        docx_path = tmp_dir / "invoice (march)!.docx"
        png_path = tmp_dir / "scan_id.png"
        make_sample_docx(docx_path)
        make_sample_png(png_path)

        with open(docx_path, "rb") as f1, open(png_path, "rb") as f2:
            resp = client.post(
                "/api/convert",
                data={"target_format": "pdf", "preset": "govt_portal_200kb"},
                files=[
                    ("files", (docx_path.name, f1, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
                    ("files", (png_path.name, f2, "image/png")),
                ],
            )
        assert resp.status_code == 200, resp.text
        body = resp.json()
        assert body["status"] == "completed"
        assert len(body["files"]) == 2

        docx_out = body["files"][0]
        assert docx_out["output_filename"] == "invoice_march.pdf", docx_out
        assert docx_out["portal_report"]["ready_for_portal"] is True, docx_out["portal_report"]

        png_out = body["files"][1]
        assert png_out["output_filename"] == "scan_id.pdf"
        checks_by_name = {c["check"]: c for c in png_out["portal_report"]["checks"]}
        assert checks_by_name["page_size"]["passed"] is True
        assert checks_by_name["file_size"]["passed"] is True
        assert checks_by_name["searchable_text"]["passed"] is False

        dl = client.get(docx_out["download_url"])
        assert dl.status_code == 200
        reader = PdfReader(io.BytesIO(dl.content))
        assert len(reader.pages) >= 1
        assert "Invoice" in (reader.pages[0].extract_text() or "")
    print("[PASS] DOCX & PNG -> PDF conversion + preflight + download verified")


def test_html_and_csv_to_pdf() -> None:
    html_bytes = b"<!doctype html><html><head><title>Portal Notice</title></head><body><h1>Hello DocPass</h1><p>Test paragraph.</p></body></html>"
    csv_bytes = b"name,amount\nInvoice A,1200\nInvoice B,3400\n"

    resp = client.post(
        "/api/convert",
        data={"target_format": "pdf", "preset": "govt_portal_200kb"},
        files=[
            ("files", ("notice.html", io.BytesIO(html_bytes), "text/html")),
            ("files", ("report.csv", io.BytesIO(csv_bytes), "text/csv")),
        ],
    )
    assert resp.status_code == 200, resp.text
    files = resp.json()["files"]
    assert len(files) == 2
    for item in files:
        assert item["portal_report"]["ready_for_portal"] is True
        dl = client.get(item["download_url"])
        assert dl.status_code == 200 and dl.content.startswith(b"%PDF")
    print("[PASS] HTML & CSV -> PDF conversion + preflight verified")


def test_pdf_all_target_formats() -> None:
    pdf_bytes = make_sample_pdf_bytes()
    for fmt in ("pdf", "png", "jpg", "webp"):
        resp = client.post(
            "/api/convert",
            data={"target_format": fmt, "preset": "govt_portal_200kb"},
            files={"files": ("sample.pdf", io.BytesIO(pdf_bytes), "application/pdf")},
        )
        assert resp.status_code == 200, f"PDF -> {fmt} failed: {resp.text}"
    print("[PASS] PDF -> pdf/png/jpg/webp all returned 200 OK")


def run_tests() -> None:
    test_cors_and_presets()
    test_frontend_static_routes()
    test_docx_and_png_conversion_with_preset()
    test_html_and_csv_to_pdf()
    test_pdf_all_target_formats()


if __name__ == "__main__":
    run_tests()