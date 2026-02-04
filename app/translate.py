import io
import re
from typing import Dict, Tuple, List, Any

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
# Header parsing (bold-aware)
# ----------------------------
def _is_bold_word(w: Dict[str, Any]) -> bool:
    """
    pdfplumber returns fontname; bold faces typically include 'Bold' in font name.
    This is not perfect for every PDF, but works well for most estimate templates.
    """
    fontname = (w.get("fontname") or "").lower()
    return "bold" in fontname or "demi" in fontname or "black" in fontname


def _group_words_into_lines(words: List[Dict[str, Any]], y_tol: float = 2.0) -> List[List[Dict[str, Any]]]:
    """
    Group words into lines by their 'top' coordinate.
    """
    if not words:
        return []

    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: List[List[Dict[str, Any]]] = []
    current: List[Dict[str, Any]] = []
    current_top = None

    for w in words:
        if current_top is None:
            current_top = w["top"]
            current = [w]
            continue
        if abs(w["top"] - current_top) <= y_tol:
            current.append(w)
        else:
            lines.append(sorted(current, key=lambda ww: ww["x0"]))
            current_top = w["top"]
            current = [w]

    if current:
        lines.append(sorted(current, key=lambda ww: ww["x0"]))

    return lines


def _extract_header_kv_by_bold(page0: pdfplumber.page.Page) -> Dict[str, str]:
    """
    Extract header key/value pairs using font boldness:
      - Key text is NOT bold, typically ends with ':'
      - Value text is bold, continues until next non-bold key ending ':'

    Also supports multiple key/value pairs on the same line.
    """
    # Grab words with font info
    words = page0.extract_words(
        extra_attrs=["fontname", "size"],
        use_text_flow=True,
        keep_blank_chars=False,
    )

    # Limit to header area (everything above the table header row).
    # We detect the Y position of the 'Line Assigned' header if available.
    header_bottom = None
    for w in words:
        if w["text"] == "Line":
            header_bottom = w["top"]
            break
    # If we didn't find it, just use a reasonable cutoff near the top
    if header_bottom is None:
        header_bottom = 220

    header_words = [w for w in words if w["top"] < header_bottom]
    lines = _group_words_into_lines(header_words)

    header: Dict[str, str] = {}

    for line in lines:
        # Walk left-to-right and segment into key/value pairs
        current_key_parts: List[str] = []
        current_val_parts: List[str] = []
        in_value = False

        def flush():
            nonlocal current_key_parts, current_val_parts, in_value
            if current_key_parts:
                key = " ".join(current_key_parts).strip().rstrip(":")
                val = " ".join(current_val_parts).strip()
                if key and val and key not in header:
                    header[key] = val
            current_key_parts = []
            current_val_parts = []
            in_value = False

        for w in line:
            text = w["text"].strip()
            if not text:
                continue

            bold = _is_bold_word(w)

            # A new key usually appears as non-bold text that ends with ':'
            if (not bold) and text.endswith(":"):
                # flush previous pair if any
                flush()
                current_key_parts = [text.rstrip(":")]
                in_value = True  # value expected next
                continue

            # Sometimes keys are multiple words before ':' (e.g., "Exterior Color:")
            # We handle that by accumulating non-bold key parts until we see a token ending with ':'
            if (not bold) and (not in_value):
                current_key_parts.append(text)
                continue

            # Value tokens are bold (your requirement)
            if in_value and bold:
                current_val_parts.append(text)
                continue

            # If we are in a value and we hit a non-bold token that looks like part of the next key,
            # we keep it in a buffer until we see a ':' token. A simpler rule:
            # ignore non-bold tokens while in_value unless they end with ':' (handled above)
            # This prevents "Year: 2020 Exterior Color:" from being pulled into the value.
            if in_value and (not bold):
                # ignore
                continue

        flush()

    return header


def _extract_header_fallback_regex(page_text: str) -> Dict[str, str]:
    """
    Fallback for PDFs where font info isn't usable.
    """
    def grab(pat):
        m = re.search(pat, page_text)
        return m.group(1).strip() if m else ""

    return {
        "RO Number": grab(r"RO Number:\s*(\d+)"),
        "Owner": grab(r"Owner:\s*(.+)"),
        "Year": grab(r"Year:\s*(\d{4})"),
        "Exterior Color": grab(r"Exterior Color:\s*(.+)"),
        "Make": grab(r"Make:\s*(.+)"),
        "Vehicle In": grab(r"Vehicle In:\s*(.+)"),
        "Vehicle Out": grab(r"Vehicle Out:\s*(.+)"),
        "Model": grab(r"Model:\s*(.+)"),
        "Mileage In": grab(r"Mileage In:\s*(.+)"),
        "Estimator": grab(r"Estimator:\s*(.+)"),
        "Body Style": grab(r"Body Style:\s*(.+)"),
        "Insurance": grab(r"Insurance:\s*(.+)"),
        "VIN": grab(r"VIN:\s*([A-Z0-9]+)"),
        "Job Number": grab(r"Job Number:\s*(.+)"),
    }


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

        # ALL CAPS section header line: "2 PILLARS, ROCKER & FLOOR"
        m_header = re.match(r"^(\d+)\s+([A-Z0-9 ,&'/.-]+)$", l)
        if m_header and ("Repair" not in l) and ("Remove" not in l):
            rows.append({"Line": int(m_header.group(1)), "Qty": "", "Operation": "", "Description": m_header.group(2), "Hours": ""})
            continue

        # Typical row:
        m = re.match(r"^(\d+)\s+([A-Za-z ]+(?:/ [A-Za-z]+)?)\s+(\d+)\s+[A-Z0-9]+\s+(.*)$", l)
        if not m:
            continue

        line_no, op, qty, rest = int(m.group(1)), m.group(2).strip(), int(m.group(3)), m.group(4)

        mh = re.search(r"(\d+\.\d+)\s*$", rest)
        hours = float(mh.group(1)) if mh else ""
        desc = rest[: mh.start()].strip() if mh else rest

        # Trim trailing tokens like "Body" or "OEM"
        if " Body " in f" {desc} ":
            desc = desc.split(" Body ")[0].strip()
        desc = re.sub(r"\bOEM\b\s*$", "", desc).strip()

        rows.append({"Line": line_no, "Qty": qty, "Operation": op, "Description": desc, "Hours": hours})

    df = pd.DataFrame(rows).sort_values("Line").reset_index(drop=True)
    df["Plain English"] = [_plain_english(o, d) for o, d in zip(df["Operation"], df["Description"])]
    df["Spanish"] = [_spanish(o, d) for o, d in zip(df["Operation"], df["Description"])]
    return df


def extract_workorder_from_pdf_bytes(pdf_bytes: bytes) -> Tuple[Dict[str, str], pd.DataFrame]:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page0 = pdf.pages[0]
        page_text = page0.extract_text() or ""
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    # Bold-aware extraction first; fallback if it returns nothing
    header = _extract_header_kv_by_bold(page0)
    if not header:
        header = _extract_header_fallback_regex(page_text)

    df = _parse_rows(full_text)
    return header, df
