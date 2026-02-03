import io
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


def render_translation_pdf(df: pd.DataFrame, out_path: str) -> None:
    """
    File-path variant (kept for convenience).
    """
    pdf_bytes = render_translation_pdf_bytes(df)
    with open(out_path, "wb") as f:
        f.write(pdf_bytes)


def render_translation_pdf_bytes(df: pd.DataFrame) -> bytes:
    """
    Render a translated work order table into PDF bytes.
    Safer for web responses (no temp-file issues).
    """
    buf = io.BytesIO()

    cols = ["Line", "Qty", "Operation", "Description", "Hours", "Plain English", "Spanish"]
    df2 = df.copy()
    for c in cols:
        if c not in df2.columns:
            df2[c] = ""

    doc = SimpleDocTemplate(buf, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Translated Work Order Table (English + Spanish)", styles["Title"]),
        Spacer(1, 12),
    ]

    data = [cols] + df2[cols].astype(str).values.tolist()

    table = Table(
        data,
        colWidths=[30, 30, 80, 140, 40, 170, 170],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("FONTSIZE", (0, 1), (-1, -1), 8),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
            ]
        )
    )

    story.append(table)
    doc.build(story)

    return buf.getvalue()

