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


def render_translation_pdf_bytes(header: Dict[str, str], df: pd.DataFrame) -> bytes:
    """
    Renders:
    1) Fixed-grid header box (Label | Value | Label | Value)
    2) Landscape translated line-item table with wrapping text
       - ALL-CAPS section headers rendered bold
    """
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=28,
        rightMargin=28,
        topMargin=28,
        bottomMargin=28,
    )

    styles = getSampleStyleSheet()

    label_style = ParagraphStyle(
        "label",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
    )

    value_style = ParagraphStyle(
        "value",
        fontName="Helvetica",
        fontSize=8,
        leading=10,
    )

    title_style = styles["Title"]

    story = []
    story.append(Paragraph("Translated Work Order (English + Spanish)", title_style))
    story.append(Spacer(1, 10))

    # ------------------------------------------------------------------
    # HEADER BOX — FIXED GRID
    # ------------------------------------------------------------------
    header_rows = [
        ["RO Number", header.get("RO Number", ""), "Owner", header.get("Owner", "")],
        ["Year", header.get("Year", ""), "Exterior Color", header.get("Exterior Color", "")],
        ["Make", header.get("Make", ""), "Vehicle In", header.get("Vehicle In", "")],
        ["Model", header.get("Model", ""), "Vehicle Out", header.get("Vehicle Out", "")],
        ["Mileage In", header.get("Mileage In", ""), "Estimator", header.get("Estimator", "")],
        ["Body Style", header.get("Body Style", ""), "Insurance", header.get("Insurance", "")],
        ["VIN", header.get("VIN", ""), "Job Number", header.get("Job Number", "")],
    ]

    header_table_data = []
    for row in header_rows:
        rendered_row = []
        for i, cell in enumerate(row):
            style = label_style if i % 2 == 0 else value_style
            rendered_row.append(Paragraph(str(cell), style))
        header_table_data.append(rendered_row)

    header_table = Table(
        header_table_data,
        colWidths=[90, 200, 90, 220],
        hAlign="LEFT",
    )

    header_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
                ("BACKGROUND", (0, 0), (-1, -1), colors.whitesmoke),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )

    story.append(header_table)
    story.append(Spacer(1, 16))

    # ------------------------------------------------------------------
    # TRANSLATED LINE ITEM TABLE
    # ------------------------------------------------------------------
    table_columns = ["Line", "Qty", "Operation", "Description", "Hours", "Plain English", "Spanish"]

    # header row
    table_data = [[Paragraph(f"<b>{c}</b>", value_style) for c in table_columns]]

    for _, row in df.iterrows():
        op = str(row.get("Operation", "") or "")
        desc = str(row.get("Description", "") or "")

        is_section_header = (op.strip() == "") and desc.isupper()

        rendered_cells = []
        for col in table_columns:
            val = str(row.get(col, "") or "")

            if is_section_header and col == "Description":
                rendered_cells.append(Paragraph(f"<b>{val}</b>", value_style))
            else:
                rendered_cells.append(Paragraph(val, value_style))

        table_data.append(rendered_cells)

    line_item_table = Table(
        table_data,
        colWidths=[38, 32, 90, 170, 46, 180, 180],
        repeatRows=1,
        hAlign="LEFT",
    )

    line_item_table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )

    story.append(line_item_table)
    doc.build(story)
    return buffer.getvalue()
