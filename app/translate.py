import io
import re
from typing import Dict, Tuple

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


# ----------------------------
# Translation helpers
# ----------------------------
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
        base = d

    if side:
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


# ----------------------------
# Header (top box) extraction
# ----------------------------
def _extract_header_kv(page_text: str) -> Dict[str, str]:
    """
    Extracts key/value pairs from the top box of the work order.
    Output is structured so it can be rendered as a box.
    """
    lines = [l.strip() for l in page_text.splitlines() if l.strip()]

    header = {}

    for l in lines:
        if l.startswith("Work Order") or (l.startswith("Line") and "Assigned" in l):
            break

        # Common patterns in body shop estimates
        patterns = {
            "RO Number": r"RO Number:\s*(.+)",
            "Owner": r"Owner:\s*(.+)",
            "Year": r"Year:\s*(\d{4})",
            "Make": r"Make:\s*([A-Za-z]+)",
            "Model": r"Model:\s*(.+)",
            "Exterior Color": r"Exterior Color:\s*(.+)",
            "Mileage In": r"Mileage In:\s*(\d+)",
            "Vehicle In": r"Vehicle In:\s*(.+)",
            "Vehicle Out": r"Vehi?cle Out:\s*(.+)",
            "Estimator": r"Estimator:\s*(.+)",
            "Insurance": r"Insurance:\s*(.+)",
            "VIN": r"VIN:\s*([A-Z0-9]+)",
            "Body Style": r"Body Style:\s*(.+)",
        }

        for key, pat in patterns.items():
            m = re.search(pat, l)
            if m and key not in header:
                header[key] = m.group(1)

    return header


# ----------------------------
# Table parsing
# ----------------------------
def _parse_rows(full_text: str) -> pd.DataFrame:
    rows = []
    lines = [l.strip() for l in full_text.splitlines() if l.strip()]

    start_idx = 0
    for i, l in enumerate(lines):
        if l.startswith("Line") and "Assigned" in l:
            start_idx = i + 1
            break

    for l in lines[start_idx:]:
        if l.startswith("Subtotals") or l.startswith("Grand Total"):
            break

        m_header = re.match(r"^(\d+)\s+([A-Z0-9 ,&'/.-]+)$", l)
        if m_header and ("Repair" not in l) and ("Remove" not in l):
            rows.append(
                {"Line": int(m_header.group(1)), "Qty": "", "Operation": "", "Description": m_header.group(2), "Hours": ""}
            )
            continue

        m = re.match(r"^(\d+)\s+([A-Za-z ]+(?:/ [A-Za-z]+)?)\s+(\d+)\s+[A-Z0-9]+\s+(.*)$", l)
        if not m:
            continue

        line_no, op, qty, rest = int(m.group(1)), m.group(2), int(m.group(3)), m.group(4)

        mh = re.search(r"(\d+\.\d+)\s*$", rest)
        hours = float(mh.group(1)) if mh else ""
        desc = rest[: mh.start()].strip() if mh else rest
        desc = desc.split(" Body ")[0].replace(" OEM", "").strip()

        rows.append(
            {"Line": line_no, "Qty": qty, "Operation": op, "Description": desc, "Hours": hours}
        )

    df = pd.DataFrame(rows).sort_values("Line").reset_index(drop=True)
    df["Plain English"] = [_plain_english(o, d) for o, d in zip(df["Operation"], df["Description"])]
    df["Spanish"] = [_spanish(o, d) for o, d in zip(df["Operation"], df["Description"])]
    return df


def extract_workorder_from_pdf_bytes(pdf_bytes: bytes):
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page0 = pdf.pages[0]
        text = page0.extract_text() or ""
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    def grab(pattern):
        m = re.search(pattern, text)
        return m.group(1).strip() if m else ""

    header = {
        "RO Number": grab(r"RO Number:\s*(\d+)"),
        "Owner": grab(r"Owner:\s*([A-Z ,]+)"),
        "Year": grab(r"Year:\s*(\d{4})"),
        "Exterior Color": grab(r"Exterior Color:\s*([A-Z]+)"),
        "Make": grab(r"Make:\s*([A-Z]+)"),
        "Vehicle In": grab(r"Vehicle In:\s*([\d/]+)"),
        "Model": grab(r"Model:\s*([A-Z0-9 ]+)"),
        "Vehicle Out": grab(r"Vehicle Out:\s*([\d/]+)"),
        "Mileage In": grab(r"Mileage In:\s*(\d+)"),
        "Estimator": grab(r"Estimator:\s*([A-Za-z .]+)"),
        "Body Style": grab(r"Body Style:\s*([A-Z0-9 ]+)"),
        "Insurance": grab(r"Insurance:\s*([A-Z0-9 ]+)"),
        "VIN": grab(r"VIN:\s*([A-Z0-9]+)"),
        "Job Number": grab(r"Job Number:\s*(\S+)"),
    }

    df = _parse_rows(full_text)
    return header, df

