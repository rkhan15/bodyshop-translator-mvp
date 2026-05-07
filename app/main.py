from datetime import datetime
from pathlib import Path
from urllib.parse import quote
import io
import re

from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, PlainTextResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from .translate import extract_workorder_from_pdf_bytes
from .pdf_render import render_translation_pdf_bytes

app = FastAPI()
app.mount("/static", StaticFiles(directory="app/static"), name="static")


@app.get("/", response_class=HTMLResponse)
def home():
    with open("app/static/index.html", "r", encoding="utf-8") as f:
        return f.read()


def _safe_download_name(uploaded_filename: str) -> str:
    stem = Path(uploaded_filename or "work_order.pdf").stem
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem).strip("_")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"translated_{timestamp}_{stem}.pdf"


@app.post("/translate")
async def translate(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        return PlainTextResponse("Please upload a PDF file.", status_code=400)

    try:
        pdf_bytes = await file.read()
        header, df = extract_workorder_from_pdf_bytes(pdf_bytes)
        out_pdf_bytes = render_translation_pdf_bytes(header, df)

        download_name = _safe_download_name(file.filename)
        encoded_name = quote(download_name)

        return StreamingResponse(
            io.BytesIO(out_pdf_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_name}"
            },
        )
    except Exception as e:
        return PlainTextResponse(f"Failed to process PDF: {e}", status_code=500)
