from fastapi import FastAPI, UploadFile, File
from fastapi.responses import HTMLResponse, PlainTextResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import io

from .translate import extract_rows_and_translate_bytes
from .pdf_render import render_translation_pdf_bytes

app = FastAPI()

# Serve the single-page UI
app.mount("/static", StaticFiles(directory="app/static"), name="static")


@app.get("/", response_class=HTMLResponse)
def home():
    with open("app/static/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/translate")
async def translate(file: UploadFile = File(...)):
    # Basic validation
    if not file.filename.lower().endswith(".pdf"):
        return PlainTextResponse("Please upload a PDF file.", status_code=400)

    try:
        pdf_bytes = await file.read()

        # Extract + translate
        df = extract_rows_and_translate_bytes(pdf_bytes)

        # Render translated PDF into bytes (no temp files)
        out_pdf_bytes = render_translation_pdf_bytes(df)

        return StreamingResponse(
            io.BytesIO(out_pdf_bytes),
            media_type="application/pdf",
            headers={"Content-Disposition": 'attachment; filename="translated_work_order.pdf"'},
        )
    except Exception as e:
        return PlainTextResponse(f"Failed to process PDF: {e}", status_code=500)

