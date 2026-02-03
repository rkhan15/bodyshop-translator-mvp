import io
import re
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
        # crude gender rules for adjectives
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


def _parse_rows_from_text(text: str) -> pd.DataFrame:
    """
    Parse 'table-ish' estimate text into rows.
    MVP heuristic:
      - Looks for the header line starting with 'Line' and containing 'Assigned'
      - Supports ALL-CAPS section headers like: '2 PILLARS, ROCKER & FLOOR'
      - Supports op rows like: '8 Remove / Install 0 DP5Z... LT Belt molding Body 0.3'
    """
    rows = []
    lines = [l.strip() for l in text.splitlines() if l.strip()]

    # Find start of table (heuristic)
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

        # Typical row pattern:
        # line, operation, qty, part_number, rest...
        m = re.match(r"^(\d+)\s+([A-Za-z ]+(?:/ [A-Za-z]+)?)\s+(\d+)\s+([A-Z0-9]+)\s+(.*)$", l)
        if not m:
            continue

        line_no = int(m.group(1))
        op = m.group(2).strip()
        qty = int(m.group(3))
        rest = m.group(5)

        # Hours at the end
        mh = re.search(r"(\d+\.\d+)\s*$", rest)
        hours = float(mh.group(1)) if mh else ""

        rest2 = rest[: mh.start()].strip() if mh else rest

        # Strip trailing labor/part tokens commonly seen:
        # labor often includes "Body"; part may include "OEM"
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


def extract_rows_and_translate(pdf_path: str) -> pd.DataFrame:
    """
    Path-based variant (kept for convenience).
    """
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    return _parse_rows_from_text(text)


def extract_rows_and_translate_bytes(pdf_bytes: bytes) -> pd.DataFrame:
    """
    Bytes-based variant (used by FastAPI upload endpoint).
    Avoids temp files and fixes FileResponse temp-path issues.
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        text = "\n".join((p.extract_text() or "") for p in pdf.pages)
    return _parse_rows_from_text(text)

