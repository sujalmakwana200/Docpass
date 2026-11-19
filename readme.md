DocPass

Convert once. Upload once. Accepted.

DocPass combines an in-browser PDF workbench (pdf-lib) with a FastAPI conversion and portal-preflight backend.

Features

In-browser tools (no upload): Image to PDF, Merge PDFs, Split PDF, Extract/Delete pages, Rotate pages, Watermark, Page numbers.

FastAPI backend tools: Word/Excel/PowerPoint/Text/HTML/CSV → PDF, PDF → JPG/PNG/WEBP, and PDF compression.

Portal-ready preflight: Checks file size limits (govt_portal_200kb, visa_500kb, tender_2mb), A4 page dimensions, searchable text, and portal-safe filenames.

Run Locally (venv)

python -m venv venv
# Windows:
.\venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate

pip install -r requirements.txt
uvicorn docpass_backend:app --host 0.0.0.0 --port 8000 --reload


Open http://localhost:8000 in your browser.

Run with Docker

docker compose up --build


Run Tests

python test_docpass_backend.py
