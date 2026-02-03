import io
from typing import Dict

import pandas as pd
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
)
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle


def render_translation_pdf_bytes(header_kv: Dict[str, str], df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()

    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=28,
        rightMargin=28,
        topMargin=28,
        bottomMargin=28,
    )

    styles = getSampleStyleSheet()
    cell = ParagraphStyle("cell", fontSize=8, leading=10)
    label = ParagraphStyle("label", fontSize=8, leading=10, fontName="Helvetica-Bold")

    story = []
    story.append(Paragraph("Translated Work Order (English + Spanish)", styles["Title"]))
    story.append(Spacer(1, 10))

    # -------- Header box --------
    if header_kv:
        header_rows = []
        items = list(header_kv.items())

        # Two-column box: (label, value) x 2 per row
        for i in range(0, len(items), 2):
            left = items[i]
            right = items[i + 1] if i + 1 < len(items) else ("", "")
            header_rows.append(
                [
                    Paragraph(left[0], label),
                    Paragraph(left[1], cell),
                    Paragraph(right[0], label),
                    Paragraph(right[1], cell),
                ]
            )

        header_table = Table(
            header_rows,
            colWidths=[90, 180, 90, 180],
            hAlign="LEFT",
        )
        header_table.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )

        story.append(header_table)
        story.append(Spacer(1, 14))

    # -------- Line item table --------
    cols = ["Line", "Qty", "Operation", "Description", "Hours", "Plain English", "Spanish"]
    data = [[Paragraph(f"<b>{c}</b>", cell) for c in cols]]

    for _, r in df.iterrows():
        data.append([Paragraph(str(r[c]), cell) for c in cols])

    table = Table(
        data,
        colWidths=[38, 32, 90, 170, 46, 180, 180],
        repeatRows=1,
    )
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )

    story.append(table)
    doc.build(story)
    return buf.getvalue()
