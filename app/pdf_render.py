import io
from typing import List

import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, KeepTogether
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def render_translation_pdf_bytes(header_lines: List[str], df: pd.DataFrame) -> bytes:
    """
    Render:
      1) Header/top-box lines from the original PDF
      2) A landscape table with wrapping columns for Plain English + Spanish
    """
    buf = io.BytesIO()

    # Landscape gives us much more width, preventing overlap
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=28,
        rightMargin=28,
        topMargin=28,
        bottomMargin=28,
    )

    styles = getSampleStyleSheet()
    title_style = styles["Title"]

    # Smaller, wrapped cell styles
    cell_style = ParagraphStyle(
        "cell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        spaceAfter=0,
        spaceBefore=0,
    )

    header_style = ParagraphStyle(
        "header",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=11,
    )

    story = []
    story.append(Paragraph("Translated Work Order (English + Spanish)", title_style))
    story.append(Spacer(1, 10))

    # --- Header box content (top of original PDF) ---
    if header_lines:
        story.append(Paragraph("<b>Original Work Order Details</b>", styles["Heading3"]))
        for line in header_lines:
            story.append(Paragraph(line, header_style))
        story.append(Spacer(1, 12))

    # --- Table ---
    cols = ["Line", "Qty", "Operation", "Description", "Hours", "Plain English", "Spanish"]
    df2 = df.copy()
    for c in cols:
        if c not in df2.columns:
            df2[c] = ""

    # Build table data with Paragraph-wrapped cells
    data = [
        [Paragraph(f"<b>{c}</b>", cell_style) for c in cols]
    ]

    for _, row in df2[cols].iterrows():
        out_row = []
        for c in cols:
            val = "" if pd.isna(row[c]) else str(row[c])
            # Preserve ALL-CAPS section headers visually
            if c == "Description" and (row["Operation"] == "" or pd.isna(row["Operation"])) and str(row["Description"]).isupper():
                val = f"<b>{val}</b>"
            out_row.append(Paragraph(val, cell_style))
        data.append(out_row)

    # Column widths tuned to keep Line visible and give English/Spanish room
    # Total width in landscape letter minus margins is ~ 11in*72=792 minus 56 = ~736
    col_widths = [38, 32, 90, 170, 46, 180, 180]

    table = Table(
        data,
        colWidths=col_widths,
        repeatRows=1,   # keep header row on each page
        hAlign="LEFT",
    )

    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.whitesmoke, colors.white]),
            ]
        )
    )

    story.append(KeepTogether([table]))
    doc.build(story)

    return buf.getvalue()
