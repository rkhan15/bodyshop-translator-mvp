import io
import re
from typing import Dict, Tuple, List

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
    "quarter panel": "panel de cuarto",
    "outer panel": "panel exterior",
    "bumper cover": "cubierta del parachoques",
    "mud flap accessory kit": "kit de accesorios de guardabarros",
    "add for clear coat": "agregar por capa transparente",
    "add for inside": "agregar por el interior",
    "overlap major non-adj. panel": "traslape de panel mayor no adyacente",
    "nib removal & polish": "remoción de imperfecciones y pulido",
}


HEADER_LAYOUT = [
    ("RO Number", "Owner"),
    ("Year", "Exterior Color"),
    ("Make", "Vehicle In"),
    ("Model", "Vehicle Out"),
    ("Mileage In", "Estimator"),
    ("Body Style", "Insurance"),
    ("VIN", "Job Number"),
    ("Customer", "License"),
    ("State", "Claim Number"),
    ("Production Date", "Condition"),
]


def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _strip_blanks(s: str) -> str:
    s = re.sub(r"_+", " ", s or "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _plain_english(op: str, desc: str) -> str:
    op = op or ""
    desc = desc or ""

    if not op and desc.isupper():
        return f"Section: {desc.title()}"

    d = desc.replace("LT ", "Left ").replace("RT ", "Right ")
    d = d.replace("w'strip", "weatherstrip").replace("assy", "assembly")
    d = d.replace("NIB", "nib")

    op_l = op.lower()

    if "repair" in op_l:
        return f"Repair the {d.lower()}."
    if "replace" in op_l or op_l == "r&r":
        return f"Remove and replace the {d.lower()}."
    if "remove" in op_l and "install" in op_l:
        return f"Remove and reinstall the {d.lower()}."
    if "blend" in op_l:
        return f"Blend refinish into the {d.lower()}."
    if "add" in op_l:
        return f"Add labor for {d.lower()}."
    if desc:
        return d

    return ""


def _spanish(op: str, desc: str) -> str:
    op = op or ""
    desc = desc or ""

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

    d_low = (
        d.lower()
        .replace("w'strip", "weatherstrip")
        .replace("assy", "door assembly")
        .replace("(hss)", "")
        .strip()
    )

    base = None
    for k, v in SPANISH_GLOSSARY.items():
        if k in d_low:
            base = v
            break

    if base is None:
        base = d

    if side:
        fem = any(w in base for w in ["puerta", "moldura", "estructura", "cubierta", "capa"])
        adj = side + ("a" if fem else "o")
        base = f"{base} {adj}"

    op_l = op.lower()

    if "repair" in op_l:
        return f"Reparar {base}."
    if "replace" in op_l or op_l == "r&r":
        return f"Retirar y reemplazar {base}."
    if "remove" in op_l and "install" in op_l:
        return f"Retirar y reinstalar {base}."
    if "blend" in op_l:
        return f"Mezclar/acoplar la pintura en {base}."
    if "add" in op_l:
        return f"Agregar labor por {base}."
    if desc:
        return base

    return ""


def _extract_header_maaco_priority2(text: str) -> Dict[str, str]:
    """
    Handles priority2_body_copy / priority2_paint_copy format.

    Important: This intentionally parses each known header line into multiple
    key/value fields instead of letting one value absorb the rest of the line.
    """
    header: Dict[str, str] = {}
    lines = [_clean_spaces(l) for l in text.splitlines() if _clean_spaces(l)]

    if lines and lines[0].startswith("MAACO"):
        header["Shop"] = lines[0]

    for line in lines:
        if line.startswith("Work Order"):
            header["Report Type"] = line
            break

        m = re.match(r"RO Number:\s*(.*)$", line)
        if m:
            header["RO Number"] = m.group(1).strip()
            continue

        m = re.match(
            r"Owner:\s*(.*?)(?:\s+Year:\s*(.*?))?(?:\s+Exterior Color:\s*(.*))?$",
            line,
        )
        if m:
            header["Owner"] = _clean_spaces(m.group(1))
            if m.group(2):
                header["Year"] = _clean_spaces(m.group(2))
            if m.group(3):
                header["Exterior Color"] = _clean_spaces(m.group(3))
            continue

        m = re.match(
            r"Vehicle In:\s*(.*?)(?:\s+Make:\s*(.*?))?(?:\s+Paint Code:\s*(.*))?$",
            line,
        )
        if m:
            header["Vehicle In"] = _clean_spaces(m.group(1))
            if m.group(2):
                header["Make"] = _clean_spaces(m.group(2))
            if m.group(3):
                header["Paint Code"] = _clean_spaces(m.group(3))
            continue

        m = re.match(
            r"Vehicle Out:\s*(.*?)(?:\s+Model:\s*(.*?))?(?:\s+License:\s*(.*))?$",
            line,
        )
        if m:
            header["Vehicle Out"] = _clean_spaces(m.group(1))
            if m.group(2):
                header["Model"] = _clean_spaces(m.group(2))
            if m.group(3):
                header["License"] = _clean_spaces(m.group(3))
            continue

        m = re.match(
            r"Estimator:\s*(.*?)(?:\s+Body Style:\s*(.*?))?(?:\s+Mileage In:\s*(.*))?$",
            line,
        )
        if m:
            header["Estimator"] = _clean_spaces(m.group(1))
            if m.group(2):
                header["Body Style"] = _clean_spaces(m.group(2))
            if m.group(3):
                header["Mileage In"] = _clean_spaces(m.group(3))
            continue

        m = re.match(
            r"Insurance:\s*(.*?)(?:\s+VIN:\s*([A-Z0-9]+))?(?:\s+Job Number:\s*(.*))?$",
            line,
        )
        if m:
            header["Insurance"] = _clean_spaces(m.group(1))
            if m.group(2):
                header["VIN"] = _clean_spaces(m.group(2))
            if m.group(3):
                header["Job Number"] = _clean_spaces(m.group(3))
            continue

    return {k: v for k, v in header.items() if v}


def _extract_header_priority1(text: str) -> Dict[str, str]:
    """
    Handles priority1_work_order blank report format.
    """
    header: Dict[str, str] = {}
    lines = [_clean_spaces(l) for l in text.splitlines() if _clean_spaces(l)]

    if lines and lines[0].startswith("Work Order"):
        header["Report Type"] = lines[0]

    for i, line in enumerate(lines[:20]):
        if line.startswith("Job Number:"):
            m = re.match(r"Job Number:\s*(.*?)\s+Customer:\s*(.*)$", line)
            if m:
                if m.group(1):
                    header["Job Number"] = _clean_spaces(m.group(1))
                if m.group(2):
                    header["Customer"] = _clean_spaces(m.group(2))
            continue

        if re.match(r"^\d{4}\s+", line):
            parts = line.split()
            header["Year"] = parts[0]
            header["Make"] = parts[1] if len(parts) > 1 else ""
            if "VIN:" not in line:
                header["Model"] = _clean_spaces(" ".join(parts[2:]))
            continue

        patterns = [
            ("VIN", r"VIN:\s*([A-Z0-9]+)"),
            ("Mileage In", r"Mileage In:\s*([\d,]+)"),
            ("Vehicle Out", r"Vehicle Out:\s*(.*?)(?:\s+License:|\s+State:|\s+Insurance Company:|$)"),
            ("License", r"License:\s*(.*?)(?:\s+Exterior Color:|\s+Mileage Out:|$)"),
            ("Exterior Color", r"Exterior Color:\s*(.*?)(?:\s+Mileage Out:|$)"),
            ("State", r"State:\s*(.*?)(?:\s+Production Date:|$)"),
            ("Production Date", r"Production Date:\s*(.*?)(?:\s+Condition:|$)"),
            ("Condition", r"Condition:\s*(.*?)(?:\s+Job #:|$)"),
            ("Insurance", r"Insurance Company:\s*(.*?)(?:\s+MAACO Paint Code:|$)"),
            ("Claim Number", r"Claim Number:\s*(.*?)(?:\s+Estimator:|$)"),
            ("Estimator", r"Estimator:\s*(.*)$"),
            ("Adjuster", r"Adjuster:\s*(.*)$"),
        ]

        for key, pat in patterns:
            m = re.search(pat, line)
            if m and _clean_spaces(m.group(1)):
                header[key] = _clean_spaces(m.group(1))

    return {k: v for k, v in header.items() if v}


def _extract_header(text: str) -> Dict[str, str]:
    if "MAACO UNION" in text and "RO Number:" in text:
        return _extract_header_maaco_priority2(text)
    return _extract_header_priority1(text)


def _is_priority1_blank_report(text: str) -> bool:
    return "BODY LABOR" in text and "REFINISH LABOR" in text and "Line # Operation Description" in text


def _parse_priority1_rows(text: str) -> pd.DataFrame:
    """
    Handles the common blank-report work order:
    BODY LABOR and REFINISH LABOR sections with repeated line numbers.
    """
    rows: List[Dict[str, object]] = []
    current_labor = None

    raw_lines = [_strip_blanks(l) for l in text.splitlines()]
    raw_lines = [l for l in raw_lines if l]

    skip_prefixes = (
        "Work Order",
        "Job Number:",
        "VIN:",
        "License:",
        "State:",
        "Insurance Company:",
        "Claim Number:",
        "Adjuster:",
        "Line #",
        "Hours",
        "Actual",
        "Date",
        "Completed",
        "Technician",
        "Total Body Labor",
        "Total Refinish Labor",
    )

    pending = None

    def flush_pending():
        nonlocal pending
        if pending:
            rows.append(pending)
            pending = None

    for line in raw_lines:
        if any(line.startswith(p) for p in skip_prefixes):
            continue
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}", line):
            continue

        if line == "BODY LABOR":
            flush_pending()
            current_labor = "BODY LABOR"
            rows.append({"Line": "", "Qty": "", "Operation": "", "Description": "BODY LABOR", "Hours": ""})
            continue

        if line == "REFINISH LABOR":
            flush_pending()
            current_labor = "REFINISH LABOR"
            rows.append({"Line": "", "Qty": "", "Operation": "", "Description": "REFINISH LABOR", "Hours": ""})
            continue

        # Section heading like "1 REAR DOOR"
        m_section = re.match(r"^(\d+)\s+([A-Z0-9 ,&'/.-]+)$", line)
        if m_section and not any(w in line for w in ["Repair", "Replace", "Blend", "Add", "Overlap", "NIB"]):
            flush_pending()
            rows.append({
                "Line": int(m_section.group(1)),
                "Qty": "",
                "Operation": "",
                "Description": m_section.group(2).strip(),
                "Hours": "",
            })
            continue

        # Task row like "4 Repair LT Quarter panel"
        m_task = re.match(
            r"^(\d+)\s+(Repair|Replace|Blend|Add|Overlap|NIB removal & Polish)\s+(.+)$",
            line,
            flags=re.IGNORECASE,
        )
        if m_task:
            flush_pending()
            line_no = int(m_task.group(1))
            op = m_task.group(2).strip()
            desc = _clean_spaces(m_task.group(3))
            pending = {
                "Line": line_no,
                "Qty": "",
                "Operation": op,
                "Description": desc,
                "Hours": "",
            }
            continue

        # Continuation line for wrapped descriptions, e.g. "side)"
        if pending and not re.match(r"^\d+\s+", line):
            pending["Description"] = _clean_spaces(str(pending["Description"]) + " " + line)

    flush_pending()

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=["Line", "Qty", "Operation", "Description", "Hours"])

    df["Plain English"] = [_plain_english(o, d) for o, d in zip(df["Operation"], df["Description"])]
    df["Spanish"] = [_spanish(o, d) for o, d in zip(df["Operation"], df["Description"])]
    return df


def _parse_priority2_rows(text: str) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []
    lines = [_clean_spaces(l) for l in text.splitlines() if _clean_spaces(l)]

    start_idx = 0
    for i, l in enumerate(lines):
        if l.startswith("Line") and "Assigned" in l:
            start_idx = i + 1
            break

    i = start_idx
    while i < len(lines):
        l = lines[i]

        if l.startswith("Subtotals") or l.startswith("Grand Total"):
            break

        m_header = re.match(r"^(\d+)\s+([A-Z0-9 ,&'/.-]+)$", l)
        if m_header and ("Repair" not in l) and ("Remove" not in l) and ("Replace" not in l):
            rows.append({
                "Line": int(m_header.group(1)),
                "Qty": "",
                "Operation": "",
                "Description": m_header.group(2).strip(),
                "Hours": "",
            })
            i += 1
            continue

        # Format with explicit operation and part number
        m = re.match(
            r"^(\d+)\s+([A-Za-z ]+(?:/ [A-Za-z]+)?)\s+(\d+)\s+([A-Z0-9]+)\s+(.*)$",
            l,
        )

        # Paint-copy rows like "4 0 Add for Clear Coat Refinish / Paint 1.7"
        m_no_part = None
        if not m:
            m_no_part = re.match(r"^(\d+)\s+(\d+)\s+(.+)$", l)

        if m:
            line_no = int(m.group(1))
            op = m.group(2).strip()
            qty = int(m.group(3))
            rest = m.group(5)
        elif m_no_part:
            line_no = int(m_no_part.group(1))
            qty = int(m_no_part.group(2))
            rest = m_no_part.group(3)
            op = ""
        else:
            i += 1
            continue

        # Pull wrapped part-number suffix lines out of the description if present.
        while i + 1 < len(lines) and re.fullmatch(r"[A-Z]{1,3}", lines[i + 1]):
            i += 1

        mh = re.search(r"(-?\d+\.\d+)\s*$", rest)
        hours = float(mh.group(1)) if mh else ""
        desc = rest[: mh.start()].strip() if mh else rest

        # Remove labor type / part type tokens.
        desc = re.split(r"\s+(Body|Refinish / Paint)\s+", desc)[0].strip()
        desc = re.sub(r"\bOEM\b\s*$", "", desc).strip()

        # Infer operation for add/overlap rows in paint copy.
        if not op:
            if desc.lower().startswith("add for"):
                op = "Add"
            elif desc.lower().startswith("overlap"):
                op = "Overlap"
            else:
                op = ""

        rows.append({
            "Line": line_no,
            "Qty": qty,
            "Operation": op,
            "Description": desc,
            "Hours": hours,
        })

        i += 1

    df = pd.DataFrame(rows)
    if df.empty:
        df = pd.DataFrame(columns=["Line", "Qty", "Operation", "Description", "Hours"])

    df = df.sort_values("Line", kind="stable").reset_index(drop=True)
    df["Plain English"] = [_plain_english(o, d) for o, d in zip(df["Operation"], df["Description"])]
    df["Spanish"] = [_spanish(o, d) for o, d in zip(df["Operation"], df["Description"])]
    return df


def _parse_rows(text: str) -> pd.DataFrame:
    if _is_priority1_blank_report(text):
        return _parse_priority1_rows(text)
    return _parse_priority2_rows(text)


def extract_workorder_from_pdf_bytes(pdf_bytes: bytes) -> Tuple[Dict[str, str], pd.DataFrame]:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = "\n".join((p.extract_text() or "") for p in pdf.pages)

    header = _extract_header(full_text)
    df = _parse_rows(full_text)
    return header, df
