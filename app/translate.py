import io
import re
from typing import List, Tuple

import pandas as pd
import pdfplumber

SPANISH_GLOSSARY = {
    "belt molding": "moldura de la ventana",
    "side molding": "moldura lateral",
    "run channel": "canal de la ventana",
    "trim panel": "panel interior",
    "door shell": "estructura de la puerta",
    "door glass": "vidrio de la puerta",
    "mirror": "espejo lateral",
    "weatherstrip": "sello de la puerta",
    "applique": "moldura decorativa",
    "door assembly": "ensamble de la puerta",
    "aperture panel": "panel de apertura",
}


def _plain_english(op: str, desc: str) -> str:
    if not op and desc.isupper():
        return f"Section: {desc.title()}"

    d = desc.replace("LT ", "Left ").replace("RT ", "Right ")
    d = d.replace("w'strip", "weatherstrip").replace("assy", "assembly")

    op_l = (op or "").lower()
    if "repair" in op_l:
        return f"Repair the {d.lower()}."
    if "remove" in op_l and "replace" in op_l:
        return f"Remove and replace the {d.lower()}."
    if "remove" in op_l and "install" in op_l:
        return f"Remove and reinstall the {d.lower()}."
    return d


def _spanish(op: str, desc: str) -> str:
    if not op and desc.isupper():
        return "Sección: " + desc.title().replace(" & ", " y ")

    side = None
    d = desc
    if d.startswith("LT "):
        side = "izquierd"
        d = d[3:]
    elif d.startswith("RT "):
        side = "derech"
        d = d[3:]

    d_low = d.lower().replace("w'strip", "weatherstrip").replace("assy", "door assembly")

    base = None
    for k, v in SPANISH_GLOSSARY.items():
        if k in d_low:
            base = v
            break
    if base is None:
        base = d  # fallback

    if side:
        # simple gender guess for adjectives
        fem = any(w in base for w in ["puerta", "moldura", "estructura"])
        adj = (side + "a") if fem else (side + "o")
        base = f"{base} {adj}"

    op_l = (op or "").lower()
    if "repair" in op_l:
        return f"Reparar {base}."
    if "remove" in op_l and "replace" in op_l:
        return f"Retirar y reemplazar {base}."
    if "remove" in op_l and "install" in op_l:
        return f"Retirar y reinstalar {base}."
    return base


def _extract_header_lines(page_text: str) -> List[str]:
    """
    Pull the 'top box' info from the PDF text (everything above the table header).
    We keep it as a list of lines to render on the translated PDF.
    """
    lines = [l.strip() for l in page_text.splitlines() if l.strip()]

    # Stop when we reach the table header or "Work Order - ..."
    stop_idx = len(lines)
    for i, l in enumerate(lines):
        if l.startswith("Work Order") or (l.startswith("Line") and "Assigned" in l):
            stop_idx = i
            break

    header_lines = lines[:stop_idx]

    # Light cleanup: remove trailing "Page 1" note if it sneaks in, keep useful metadata
    cleaned = []
    for l in header_lines:
        if re.search(r"\bPage\s+\d+\b", l):
            continue
        cleaned.append(l)

    return cleaned


def _parse_rows_from_text(full_text: str) -> pd.DataFrame:
    """
    Parse the table section into rows.
    Supports:
      - ALL CAPS section headers like: "2 PILLARS, ROCKER & FLOOR"
      - Op rows like: "8 Remove / Install 0 DP5Z... LT Belt molding Body 0.3"
    """
    rows = []
    lines = [l.strip() for l in full_text.splitlines() if l.strip()]

    # Find start of table
    start_idx = 0
    for i, l in enumerate(lines):
        if l.startswith("Line") and "Assigned" in l:
            start_idx = i + 1
            break

    for l in lines[start_idx:]:
        if l.startswith("Subtotals") or l.startswith("Grand Total"):
            break

        # ALL CAPS section header line
        m_header = re.match(r"^(\d+)\s+([A-Z0-9 ,&'/.-]+)$", l)
        if m_header and ("Repair" not in l) and ("Remove" not in l):
            rows.append(
                {
                    "Line": int(m_header.group(1)),
                    "Qty": "",
                    "Operation": "",
                    "Description": m_header.group(2).strip(),
                    "Hours": "",
                }
            )
            continue

        # Typical row:
        m = re.match(r"^(\d+)\s+([A-Za-z ]+(?:/ [A-Za-z]+)?)\s+(\d+)\s+([A-Z0-9]+)\s+(.*)$", l)
        if not m:
            continue

        line_no = int(m.group(1))
        op = m.group(2).strip()
        qty = int(m.group(3))
        rest = m.group(5)

        # Hours at end
        mh = re.search(r"(\d+\.\d+)\s*$", rest)
        hours = float(mh.group(1)) if mh else ""

        rest2 = rest[: mh.start()].strip() if mh else rest

        # Strip common trailing tokens
        if " Body " in f" {rest2} ":
            rest2 = rest2.split(" Body ")[0].strip()
        rest2 = re.sub(r"\bOEM\b\s*$", "", rest2).strip()

        rows.append(
            {
                "Line": line_no,
                "Qty": qty,
                "Operation": op,
                "Description": rest2,
                "Hours": hours,
            }
        )

    df = pd.DataFrame(rows).sort_values("Line").reset_index(drop=True)
    df["Plain English"] = [_plain_english(o, d) for o, d in zip(df["Operation"], df["Description"])]
    df["Spanish"] = [_spanish(o, d) for o, d in zip(df["Operation"], df["Description"])]
    return df


def extract_workorder_from_pdf_bytes(pdf_bytes: bytes) -> Tuple[List[str], pd.DataFrame]:
    """
    Returns:
      - header_lines: top-of-PDF details (RO#, owner, vehicle, etc.)
      - df: parsed line items with translations
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        # MVP assumes first page has the header + table (common for many work orders).
        page0 = pdf.pages[0]
        page_text = page0.extract_text() or ""

        # If multi-page, we still parse table rows across all pages (helps future-proofing).
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    header_lines = _extract_header_lines(page_text)
    df = _parse_rows_from_text(full_text)
    return header_lines, df
