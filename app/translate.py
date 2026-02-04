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
    return re.sub(r"\s+", " ", (k or "")).strip().rstrip(":").strip()


def _words_in_bbox(page0, bbox, pad: float = 2.0) -> List[Dict[str, Any]]:
    """
    bbox = (x0, top, x1, bottom)
    We expand it slightly to avoid missing text that sits near borders.
    """
    x0, top, x1, bottom = bbox
    x0p = max(0, x0 - pad)
    topp = max(0, top - pad)
    x1p = x1 + pad
    botp = bottom + pad

    crop = page0.within_bbox((x0p, topp, x1p, botp))
    words = crop.extract_words(
        extra_attrs=["fontname", "size"],
        use_text_flow=True,
        keep_blank_chars=False,
    )
    return sorted(words, key=lambda w: (w["top"], w["x0"]))


def _group_words_into_lines(words: List[Dict[str, Any]], y_tol: float = 2.0) -> List[List[Dict[str, Any]]]:
    if not words:
        return []
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines: List[List[Dict[str, Any]]] = []
    cur: List[Dict[str, Any]] = []
    cur_top = None
    for w in words:
        if cur_top is None or abs(w["top"] - cur_top) <= y_tol:
            cur.append(w)
            cur_top = w["top"] if cur_top is None else cur_top
        else:
            lines.append(sorted(cur, key=lambda ww: ww["x0"]))
            cur = [w]
            cur_top = w["top"]
    if cur:
        lines.append(sorted(cur, key=lambda ww: ww["x0"]))
    return lines


def _extract_inline_pairs_from_tokens(tokens: List[tuple]) -> Tuple[str, Dict[str, str]]:
    """
    tokens: List[(text, is_bold)]
    Returns:
      - main_value: bold tokens until first inline key (or non-bold key pattern)
      - extras: inline key/value pairs on same line/cell
    Rules:
      - inline key: non-bold phrase that ends with ':'
      - inline value: preferably bold tokens immediately after key
      - if bold is not available, fall back to non-bold tokens until next key
    """
    if not tokens:
        return "", {}

    main_value_parts: List[str] = []
    extras: Dict[str, str] = {}

    i = 0
    in_main = True

    while i < len(tokens):
        text, bold = tokens[i]
        if not text:
            i += 1
            continue

        # detect key phrase ending with ':'
        if (not bold):
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
                in_main = False
                key = _clean_key(" ".join(key_parts))

                # collect value tokens after key
                k = j + 1
                val_parts: List[str] = []

                # Prefer bold
                while k < len(tokens) and tokens[k][1]:
                    val_parts.append(tokens[k][0])
                    k += 1

                # If no bold value, fallback to non-bold tokens until next key
                if not val_parts:
                    while k < len(tokens):
                        t2, b2 = tokens[k]
                        if (not b2) and t2.endswith(":"):
                            break
                        if b2:
                            # if bold appears later, treat it as value too
                            val_parts.append(t2)
                        else:
                            val_parts.append(t2)
                        k += 1

                val = " ".join(val_parts).strip()
                if key and val:
                    extras[key] = val

                i = k
                continue

        # main value rule:
        # - if bold exists, collect bold tokens
        # - if no bold exists at all, collect tokens until first inline key
        if in_main:
            if bold:
                main_value_parts.append(text)
            else:
                # only collect non-bold main value if we never see bold anywhere
                # handled below after loop
                pass

        i += 1

    main_value = " ".join(main_value_parts).strip()

    # If main_value ended empty (e.g., no bold fonts), fallback to tokens until first key
    if not main_value:
        fallback_parts = []
        for t, b in tokens:
            if (not b) and t.endswith(":"):
                break
            # ignore obvious label-like bits
            fallback_parts.append(t)
        main_value = " ".join(fallback_parts).strip()

    return main_value, extras


# ----------------------------
# Header extraction strategies
# ----------------------------
def _header_table_from_page(page0) -> Optional[Any]:
    settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "lines",
        "intersection_tolerance": 6,
        "snap_tolerance": 4,
        "join_tolerance": 4,
        "edge_min_length": 12,
    }
    tables = page0.find_tables(settings)
    if not tables:
        return None

    # Prefer the table that contains "RO Number" somewhere
    for t in tables:
        try:
            data = t.extract()
            flat = " ".join([" ".join([c or "" for c in row]) for row in data]).lower()
            if "ro number" in flat:
                return t
        except Exception:
            continue

    return tables[0]


def _extract_header_from_top_box_table(page0) -> Dict[str, str]:
    """
    Best attempt: parse header from detected table cells, with bbox padding.
    """
    header: Dict[str, str] = {}
    t = _header_table_from_page(page0)
    if t is None:
        return header

    # t.cells are bboxes; cluster into rows by 'top'
    cells = sorted(t.cells, key=lambda b: (b[1], b[0]))
    rows: List[List[tuple]] = []
    y_tol = 4.0
    cur: List[tuple] = []
    cur_top = None
    for bbox in cells:
        top = bbox[1]
        if cur_top is None or abs(top - cur_top) <= y_tol:
            cur.append(bbox)
            cur_top = top if cur_top is None else cur_top
        else:
            rows.append(sorted(cur, key=lambda bb: bb[0]))
            cur = [bbox]
            cur_top = top
    if cur:
        rows.append(sorted(cur, key=lambda bb: bb[0]))

    # Map each row to (k1,v1,k2,v2) if possible
    for row_bboxes in rows:
        if len(row_bboxes) < 2:
            continue

        # Label 1
        w_k1 = _words_in_bbox(page0, row_bboxes[0], pad=3.5)
        k1 = _clean_key(" ".join([w["text"] for w in w_k1]))
        # Value 1
        v1 = ""
        extras1: Dict[str, str] = {}
        if len(row_bboxes) >= 2:
            w_v1 = _words_in_bbox(page0, row_bboxes[1], pad=3.5)
            tokens = [(w["text"], _is_bold_word(w)) for w in w_v1 if w.get("text", "").strip()]
            v1, extras1 = _extract_inline_pairs_from_tokens(tokens)

        if k1 and v1:
            header[k1] = v1
        for ek, ev in extras1.items():
            header[ek] = ev

        # Label 2 / Value 2 (if present)
        if len(row_bboxes) >= 4:
            w_k2 = _words_in_bbox(page0, row_bboxes[2], pad=3.5)
            k2 = _clean_key(" ".join([w["text"] for w in w_k2]))

            w_v2 = _words_in_bbox(page0, row_bboxes[3], pad=3.5)
            tokens2 = [(w["text"], _is_bold_word(w)) for w in w_v2 if w.get("text", "").strip()]
            v2, extras2 = _extract_inline_pairs_from_tokens(tokens2)

            if k2 and v2:
                header[k2] = v2
            for ek, ev in extras2.items():
                header[ek] = ev

    return header


def _extract_header_from_bold_lines(page0) -> Dict[str, str]:
    """
    Fallback: parse using lines of words and bold/non-bold transitions.
    """
    words = page0.extract_words(
        extra_attrs=["fontname", "size"],
        use_text_flow=True,
        keep_blank_chars=False,
    )
    if not words:
        return {}

    # stop at table header if present
    header_bottom = None
    for w in words:
        if w.get("text") == "Line":
            header_bottom = w["top"]
            break
    if header_bottom is None:
        header_bottom = 220

    words = [w for w in words if w["top"] < header_bottom]
    lines = _group_words_into_lines(words)

    header: Dict[str, str] = {}
    for line in lines:
        tokens = [(w["text"].strip(), _is_bold_word(w)) for w in line if w.get("text", "").strip()]
        if not tokens:
            continue

        # Parse potentially multiple key/values in one line by "Key:" markers
        i = 0
        while i < len(tokens):
            # find next key end ":"
            if (not tokens[i][1]) and tokens[i][0].endswith(":"):
                key = _clean_key(tokens[i][0])
                # value = bold run after key
                j = i + 1
                val_parts = []
                while j < len(tokens) and tokens[j][1]:
                    val_parts.append(tokens[j][0])
                    j += 1
                if not val_parts:
                    # fallback: collect non-bold until next key
                    while j < len(tokens):
                        if (not tokens[j][1]) and tokens[j][0].endswith(":"):
                            break
                        val_parts.append(tokens[j][0])
                        j += 1

                val = " ".join(val_parts).strip()
                if key and val and key not in header:
                    header[key] = val

                i = j
            else:
                i += 1

    return header


def _extract_header_regex(page_text: str) -> Dict[str, str]:
    """
    Last-resort fallback.
    """
    def grab(pat):
        m = re.search(pat, page_text)
        return m.group(1).strip() if m else ""

    return {
        "RO Number": grab(r"RO Number:\s*(\d+)"),
        "Owner": grab(r"Owner:\s*([A-Z ,]+)"),
        "Year": grab(r"Year:\s*(\d{4})"),
        "Exterior Color": grab(r"Exterior Color:\s*([A-Z]+)"),
        "Make": grab(r"Make:\s*([A-Z]+)"),
        "Vehicle In": grab(r"Vehicle In:\s*([\d/]+)"),
        "Vehicle Out": grab(r"Vehicle Out:\s*([\d/]+)"),
        "Model": grab(r"Model:\s*(.+?)(?:Vehicle Out:|$)"),
        "Mileage In": grab(r"Mileage In:\s*(\d+)"),
        "Estimator": grab(r"Estimator:\s*(.+?)(?:Insurance:|$)"),
        "Body Style": grab(r"Body Style:\s*(.+?)(?:VIN:|$)"),
        "Insurance": grab(r"Insurance:\s*(.+)"),
        "VIN": grab(r"VIN:\s*([A-Z0-9]+)"),
        "Job Number": grab(r"Job Number:\s*(.+)"),
    }


def _normalize_header_keys(header: Dict[str, str]) -> Dict[str, str]:
    """
    Map common variants into the keys your renderer expects.
    """
    out = dict(header)

    # Common key variants
    mapping = {
        "RO#": "RO Number",
        "RO": "RO Number",
        "Exterior": "Exterior Color",
        "ExteriorColor": "Exterior Color",
        "JobNumber": "Job Number",
        "VehicleIn": "Vehicle In",
        "VehicleOut": "Vehicle Out",
        "Mileage": "Mileage In",
    }

    for k, v in list(out.items()):
        kk = _clean_key(k)
        if kk in mapping:
            out[mapping[kk]] = v
            del out[k]
        elif kk != k:
            out[kk] = v
            del out[k]

    return out


# ----------------------------
# Line item parsing
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

        # ALL CAPS section heading
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
        page_text = page0.extract_text() or ""
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    # Strategy 1: table-based (with bbox padding + bold fallback)
    header = _extract_header_from_top_box_table(page0)

    # Strategy 2: bold-line parsing fallback
    if not header:
        header = _extract_header_from_bold_lines(page0)

    # Strategy 3: regex fallback
    if not header:
        header = _extract_header_regex(page_text)

    header = _normalize_header_keys(header)

    df = _parse_rows(full_text)
    return header, df
