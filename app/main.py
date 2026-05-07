from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, PlainTextResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import io
from datetime import datetime
import os

from .translate import extract_workorder_from_pdf_bytes
from .pdf_render import render_translation_pdf_bytes

app = FastAPI()

app.mount("/static", StaticFiles(directory="app/static"), name="static")


@app.get("/", response_class=HTMLResponse)
def home():
    with open("app/static/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/translate")
async def translate(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        return PlainTextResponse("Please upload a PDF file.", status_code=400)

    try:
        pdf_bytes = await file.read()

        # Extract header + line items
        header_lines, df = extract_workorder_from_pdf_bytes(pdf_bytes)

        # Render to PDF bytes
        out_pdf_bytes = render_translation_pdf_bytes(header_lines, df)

        # Generate dynamic filename with timestamp and original filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        original_name = os.path.splitext(file.filename)[0]
        output_filename = f"translated_{timestamp}_{original_name}.pdf"

        return StreamingResponse(
            io.BytesIO(out_pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": f'attachment; filename="{output_filename}"'},
        )
    except Exception as e:
        return PlainTextResponse(f"Failed to process PDF: {e}", status_code=500)
