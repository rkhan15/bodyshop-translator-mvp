import io
import re
from typing import Dict, Tuple, List, Any, Optional

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
# Font/style helpers
# ----------------------------
def _is_bold_word(w: Dict[str, Any]) -> bool:
    fontname = (w.get("fontname") or "").lower()
    return ("bold" in fontname) or ("demi" in fontname) or ("black" in fontname)


def _clean_key(k: str) -> str:
    return re.sub(r"\s+", " ", (k or "")).strip().rstrip(":")


def _words_in_bbox(page0, bbox) -> List[Dict[str, Any]]:
    """
    bbox = (x0, top, x1, bottom)
    """
    x0, top, x1, bottom = bbox
    words = page0.within_bbox((x0, top, x1, bottom)).extract_words(
        extra_attrs=["fontname", "size"],
        use_text_flow=True,
        keep_blank_chars=False,
    )
    # reading order
    return sorted(words, key=lambda w: (w["top"], w["x0"]))


def _extract_inline_pairs_from_cell(words: List[Dict[str, Any]]) -> Tuple[str, Dict[str, str]]:
    """
    Given words from a value cell, return:
      - main_value: bold text up to the first detected inline key (e.g., "EDSON, TERRY")
      - extras: inline key/value pairs where key is non-bold ending in ':' and value is bold
               e.g., {"Year": "2020", "Exterior Color": "SILVER"}
    """
    # Build a token stream preserving bold flags
    tokens = [(w["text"].strip(), _is_bold_word(w)) for w in words if w.get("text", "").strip()]
    if not tokens:
        return "", {}

    # Detect inline keys that look like "Year:" "Exterior" "Color:" (multi-word keys)
    # We'll accumulate non-bold tokens until one ends with ":" => that becomes the key.
    main_value_parts: List[str] = []
    extras: Dict[str, str] = {}

    i = 0
    in_main = True

    while i < len(tokens):
        text, bold = tokens[i]

        # Identify a key: sequence of non-bold tokens ending with ":"
        if (not bold):
            # try to build a key phrase ending with ':'
            key_parts = []
            j = i
            found_key = False
            while j < len(tokens):
                t, b = tokens[j]
                if b:
                    break
                key_parts.append(t)
                if t.endswith(":"):
                    found_key = True
                    break
                j += 1

            if found_key:
                # Switch from main value (if we were in it)
                in_main = False
                key = _clean_key(" ".join(key_parts))

                # Now consume bold value tokens after this key
                k = j + 1
                val_parts = []
                while k < len(tokens):
                    t2, b2 = tokens[k]
                    if not b2:
                        # stop when next non-bold token (likely next key)
                        break
                    val_parts.append(t2)
                    k += 1

                val = " ".join(val_parts).strip()
                if key and val:
                    extras[key] = val

                i = k
                continue

        # If we're still in main section, collect bold tokens only
        if in_main and bold:
            main_value_parts.append(text)

        i += 1

    main_value = " ".join(main_value_parts).strip()
    return main_value, extras


def _header_table_from_page(page0) -> Optional[Any]:
    """
    Find the header "top box" table on the first page using line strategies.
    Returns the pdfplumber Table object if found.
    """
    settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 5,
        "snap_tolerance": 3,
        "join_tolerance": 3,
        "edge_min_length": 10,
        "min_words_vertical": 1,
        "min_words_horizontal": 1,
    }

    tables = page0.find_tables(settings)
    if not tables:
        return None

    # Prefer the table that contains "RO Number" in its extracted text
    for t in tables:
        data = t.extract()
        flat = " ".join([" ".join([c or "" for c in row]) for row in data]).lower()
        if "ro number" in flat:
            return t

    # fallback: first table
    return tables[0]


def _extract_header_from_top_box(page0) -> Dict[str, str]:
    """
    Extract header key/value pairs from the top box while:
      - Keeping ONLY bold tokens as the "value" for the cell's main key
      - Splitting inline key/value pairs inside the same value cell (unbold key + bold value)
    """
    header: Dict[str, str] = {}

    t = _header_table_from_page(page0)
    if t is None:
        return header

    # Table cells in reading order; each cell has bbox
    # We expect a 7x4 grid (Label, Value, Label, Value), but templates vary slightly.
    # We'll just read each row's 4 cells if available.
    cells = t.cells
    # Group cells by row using their top coordinate
    cells_sorted = sorted(cells, key=lambda c: (c[1], c[0]))  # (x0, top) but c is (x0, top, x1, bottom)

    # Build rows by y coordinate clustering
    rows: List[List[tuple]] = []
    y_tol = 3.0
    current_row = []
    current_top = None
    for bbox in cells_sorted:
        top = bbox[1]
        if current_top is None or abs(top - current_top) <= y_tol:
            current_row.append(bbox)
            current_top = top if current_top is None else current_top
        else:
            rows.append(sorted(current_row, key=lambda b: b[0]))
            current_row = [bbox]
            current_top = top
    if current_row:
        rows.append(sorted(current_row, key=lambda b: b[0]))

    # For each row, we try to map 4 cells: (k1, v1, k2, v2)
    for row_bboxes in rows:
        if len(row_bboxes) < 2:
            continue

        # Extract texts from each cell bbox
        cell_texts = []
        cell_words = []
        for bbox in row_bboxes[:4]:
            w = _words_in_bbox(page0, bbox)
            cell_words.append(w)
            cell_texts.append(" ".join([ww["text"] for ww in w]).strip())

        # Basic mapping:
        # cell 0 label -> cell 1 value
        if len(cell_texts) >= 2:
            k1 = _clean_key(cell_texts[0])
            if k1:
                main_val, extras = _extract_inline_pairs_from_cell(cell_words[1])
                if main_val:
                    header[k1] = main_val
                # also capture extras found inside that value cell
                for ek, ev in extras.items():
                    header[ek] = ev

        # cell 2 label -> cell 3 value (if present)
        if len(cell_texts) >= 4:
            k2 = _clean_key(cell_texts[2])
            if k2:
                main_val, extras = _extract_inline_pairs_from_cell(cell_words[3])
                if main_val:
                    header[k2] = main_val
                for ek, ev in extras.items():
                    header[ek] = ev

    return header


# ----------------------------
# Line item table parsing
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
            rows.append({"Line": int(m_header.group(1)), "Qty": "", "Operation": "", "Description": m_header.group(2), "Hours": ""})
            continue

        m = re.match(r"^(\d+)\s+([A-Za-z ]+(?:/ [A-Za-z]+)?)\s+(\d+)\s+[A-Z0-9]+\s+(.*)$", l)
        if not m:
            continue

        line_no, op, qty, rest = int(m.group(1)), m.group(2).strip(), int(m.group(3)), m.group(4)

        mh = re.search(r"(\d+\.\d+)\s*$", rest)
        hours = float(mh.group(1)) if mh else ""
        desc = rest[: mh.start()].strip() if mh else rest

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
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    header = _extract_header_from_top_box(page0)
    df = _parse_rows(full_text)
    return header, df
