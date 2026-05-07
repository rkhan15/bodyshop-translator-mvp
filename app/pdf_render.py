import io
import html
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


def _p(text, style, bold: bool = False):
    safe = html.escape(str(text or ""))
    if bold:
        safe = f"<b>{safe}</b>"
    return Paragraph(safe, style)


def _is_section_or_total_row(row) -> bool:
    op = str(row.get("Operation", "") or "").strip()
    desc = str(row.get("Description", "") or "").strip()

    if op:
        return False

    if desc.isupper():
        return True

    if desc.upper().startswith("TOTAL") and "HOURS" in desc.upper():
        return True

    return False


def render_translation_pdf_bytes(header: Dict[str, str], df: pd.DataFrame) -> bytes:
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
    # HEADER BOX — ONLY ESSENTIAL VEHICLE INFO
    # ------------------------------------------------------------------
    header_rows = [
        ["RO Number", header.get("RO Number", ""), "Year", header.get("Year", "")],
        ["Make", header.get("Make", ""), "Model", header.get("Model", "")],
        ["Exterior Color", header.get("Exterior Color", ""), "", ""],
    ]

    header_table_data = []
    for row in header_rows:
        rendered_row = []
        for i, cell in enumerate(row):
            style = label_style if i % 2 == 0 else value_style
            rendered_row.append(_p(cell, style))
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
    include_part_number = (
        "Part Number" in df.columns
        and df["Part Number"].fillna("").astype(str).str.strip().ne("").any()
    )

    table_columns = ["Line", "Qty"]
    if include_part_number:
        table_columns.append("Part Number")

    table_columns += ["Operation", "Description", "Hours", "Plain English", "Spanish"]

    table_data = [[_p(c, value_style, bold=True) for c in table_columns]]

    row_is_bold = []

    for _, row in df.iterrows():
        bold_row = _is_section_or_total_row(row)
        row_is_bold.append(bold_row)

        rendered_cells = []
        for col in table_columns:
            val = str(row.get(col, "") or "")
            rendered_cells.append(_p(val, value_style, bold=bold_row))

        table_data.append(rendered_cells)

    if include_part_number:
        col_widths = [32, 28, 92, 78, 145, 40, 170, 170]
    else:
        col_widths = [38, 32, 90, 170, 46, 180, 180]

    line_item_table = Table(
        table_data,
        colWidths=col_widths,
        repeatRows=1,
        hAlign="LEFT",
    )

    style_cmds = [
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]

    # Light gray background for section/total rows.
    for idx, bold_row in enumerate(row_is_bold, start=1):
        if bold_row:
            style_cmds.append(("BACKGROUND", (0, idx), (-1, idx), colors.whitesmoke))

    line_item_table.setStyle(TableStyle(style_cmds))

    story.append(line_item_table)
    doc.build(story)
    return buffer.getvalue()
