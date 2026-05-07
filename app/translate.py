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
    "clear coat": "capa transparente",
    "inside": "interior",
    "overlap major non-adj. panel": "traslape de panel mayor no adyacente",
    "nib removal & polish": "remoción de imperfecciones y pulido",
}


def _clean_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "")).strip()


def _strip_blanks(s: str) -> str:
    s = (s or "").replace("(cid:9)", " ")
    s = re.sub(r"_+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _normalize_add_desc(desc: str) -> str:
    d = _clean_spaces(desc)
    if d.lower().startswith("add for "):
        return _clean_spaces(d[8:])
    if d.lower().startswith("for "):
        return _clean_spaces(d[4:])
    return d


def _extract_labeled_values(line: str, labels: List[str]) -> Dict[str, str]:
    matches = []
    for label in labels:
        m = re.search(re.escape(label) + r":", line)
        if m:
            matches.append((m.start(), m.end(), label))

    matches.sort()
    out = {}

    for idx, (_, end, label) in enumerate(matches):
        next_start = matches[idx + 1][0] if idx + 1 < len(matches) else len(line)
        val = _clean_spaces(line[end:next_start])
        if val:
            out[label] = val

    return out


def _plain_english(op: str, desc: str) -> str:
    op = op or ""
    desc = desc or ""

    if not op and desc.isupper():
        return f"Section: {desc.title()}"

    d = desc.replace("LT ", "Left ").replace("RT ", "Right ")
    d = d.replace("w'strip", "weatherstrip").replace("assy", "assembly")
    d = d.replace("NIB", "nib")

    op_l = op.lower()
    d_l = d.lower()

    if "total" in d_l and "hours" in d_l:
        return "Total labor hours."
    if "nib removal" in op_l or "nib removal" in d_l:
        return "Remove paint nibs/imperfections and polish the finish."
    if "repair" in op_l:
        return f"Repair the {d.lower()}."
    if "replace" in op_l or op_l == "r&r":
        return f"Remove and replace the {d.lower()}."
    if "remove" in op_l and "install" in op_l:
        return f"Remove and reinstall the {d.lower()}."
    if "blend" in op_l:
        return f"Blend refinish into the {d.lower()}."
    if "add" in op_l:
        return f"Add labor for {_normalize_add_desc(d).lower()}."
    if "overlap" in op_l or d_l.startswith("overlap"):
        return f"Apply overlap labor for {d.lower()}."
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

    op_l = op.lower()

    if "total" in d_low and "hours" in d_low:
        return "Total de horas de labor."
    if "nib removal" in op_l or "nib removal" in d_low:
        base = "remoción de imperfecciones y pulido"
    else:
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

    if "nib removal" in op_l or "nib removal" in d_low:
        return "Remover imperfecciones de pintura y pulir el acabado."
    if "repair" in op_l:
        return f"Reparar {base}."
    if "replace" in op_l or op_l == "r&r":
        return f"Retirar y reemplazar {base}."
    if "remove" in op_l and "install" in op_l:
        return f"Retirar y reinstalar {base}."
    if "blend" in op_l:
        return f"Mezclar/acoplar la pintura en {base}."
    if "add" in op_l:
        return f"Agregar labor por {_normalize_add_desc(base)}."
    if "overlap" in op_l or d_low.startswith("overlap"):
        return f"Aplicar labor de traslape para {base}."
    if desc:
        return base

    return ""


def _extract_header_maaco_priority2(text: str) -> Dict[str, str]:
    header: Dict[str, str] = {}
    lines = [_clean_spaces(l) for l in text.splitlines() if _clean_spaces(l)]

    if lines and lines[0].startswith("MAACO"):
        header["Shop"] = lines[0]

    for line in lines:
        if line.startswith("Work Order"):
            header["Report Type"] = line
            break

        if line.startswith("RO Number:"):
            header.update(_extract_labeled_values(line, ["RO Number"]))
            continue

        for labels in [
            ["Owner", "Year", "Exterior Color"],
            ["Vehicle In", "Make", "Paint Code"],
            ["Vehicle Out", "Model", "License"],
            ["Estimator", "Body Style", "Mileage In"],
            ["Insurance", "VIN", "Job Number"],
        ]:
            vals = _extract_labeled_values(line, labels)
            if vals:
                header.update(vals)
                break

    return {k: v for k, v in header.items() if v}


def _extract_header_priority1(text: str) -> Dict[str, str]:
    header: Dict[str, str] = {}
    lines = [_clean_spaces(l) for l in text.splitlines() if _clean_spaces(l)]

    if lines and lines[0].startswith("Work Order"):
        header["Report Type"] = lines[0]

    for line in lines[:25]:
        if line.startswith("Job Number:"):
            header.update(_extract_labeled_values(line, ["Job Number", "Customer"]))
            continue

        if re.match(r"^\d{4}\s+", line):
            parts = line.split()
            header["Year"] = parts[0]
            header["Make"] = parts[1] if len(parts) > 1 else ""
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
            if m:
                val = _clean_spaces(m.group(1))
                if val and not val.endswith(":") and not re.fullmatch(r"[A-Za-z ]+:\s*", val):
                    header[key] = val

    return {k: v for k, v in header.items() if v}


def _extract_header(text: str) -> Dict[str, str]:
    if "MAACO UNION" in text and "RO Number:" in text:
        return _extract_header_maaco_priority2(text)
    return _extract_header_priority1(text)


def _is_priority1_blank_report(text: str) -> bool:
    return "BODY LABOR" in text and "REFINISH LABOR" in text and "Line # Operation Description" in text


def _append_translations(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        df = pd.DataFrame(columns=["Line", "Qty", "Part Number", "Operation", "Description", "Hours"])

    for col in ["Line", "Qty", "Part Number", "Operation", "Description", "Hours"]:
        if col not in df.columns:
            df[col] = ""

    df["Plain English"] = [_plain_english(o, d) for o, d in zip(df["Operation"], df["Description"])]
    df["Spanish"] = [_spanish(o, d) for o, d in zip(df["Operation"], df["Description"])]
    return df


def _append_total_row(rows: List[Dict[str, object]], label: str, hours: object = "") -> None:
    rows.append({
        "Line": "",
        "Qty": "",
        "Part Number": "",
        "Operation": "",
        "Description": label,
        "Hours": hours,
    })


def _extract_total_hours(text: str) -> str:
    m = re.search(r"Grand Total:\s*(-?\d+(?:\.\d+)?)", text)
    if m:
        return m.group(1)

    m = re.search(r"Subtotals:\s+[A-Za-z /]+\s+(-?\d+(?:\.\d+)?)", text)
    if m:
        return m.group(1)

    m = re.search(r"Total Body Labor\s+(-?\d+(?:\.\d+)?)", text)
    if m:
        return m.group(1)

    return ""


def _parse_priority1_rows(text: str) -> pd.DataFrame:
    rows: List[Dict[str, object]] = []
    raw = _strip_blanks(text)

    section_specs = [
        ("BODY LABOR", "Total Body Labor", "TOTAL BODY HOURS"),
        ("REFINISH LABOR", "Total Refinish Labor", "TOTAL REFINISH HOURS"),
    ]

    for section_name, total_marker, total_label in section_specs:
        if section_name not in raw:
            continue

        start = raw.find(section_name)
        end = raw.find(total_marker, start)
        if end == -1:
            end = len(raw)

        section_text = raw[start + len(section_name):end]

        rows.append({
            "Line": "",
            "Qty": "",
            "Part Number": "",
            "Operation": "",
            "Description": section_name,
            "Hours": "",
        })

        # Strip column header text.
        section_text = re.sub(
            r"Line # Operation Description Assigned Hours Actual Hours Date Completed Technician",
            " ",
            section_text,
            flags=re.IGNORECASE,
        )

        # Candidate starts we expect in priority1.
        start_re = re.compile(
            r"(?<!\d)(\d+)\s+(?=(REAR DOOR|QUARTER PANEL|REAR BUMPER|FRONT DOOR|Repair|Replace|Blend|Add|Overlap|NIB removal))",
            flags=re.IGNORECASE,
        )

        matches = list(start_re.finditer(section_text))

        for idx, match in enumerate(matches):
            line_no = int(match.group(1))
            record_start = match.end()
            record_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(section_text)
            content = _clean_spaces(section_text[record_start:record_end])

            if not content:
                continue

            # Section heading row.
            if content.isupper() and not re.search(r"\b(Repair|Replace|Blend|Add|Overlap|NIB)\b", content, re.IGNORECASE):
                rows.append({
                    "Line": line_no,
                    "Qty": "",
                    "Part Number": "",
                    "Operation": "",
                    "Description": content,
                    "Hours": "",
                })
                continue

            op = ""
            desc = ""

            task_patterns = [
                (r"^(NIB removal & Polish)\b\s*(.*)$", "NIB removal & Polish"),
                (r"^(Overlap Major Non-Adj\. Panel)\b\s*(.*)$", "Overlap"),
                (r"^(Repair|Replace|Blend|Add|Overlap)\b\s*(.*)$", None),
            ]

            for pat, forced_op in task_patterns:
                mt = re.match(pat, content, flags=re.IGNORECASE)
                if mt:
                    op = forced_op or _clean_spaces(mt.group(1))
                    desc = _clean_spaces(mt.group(2) if len(mt.groups()) > 1 else "")
                    break

            if op.lower() == "nib removal & polish" and not desc:
                desc = "NIB removal & Polish"
            elif op.lower() == "add":
                desc = _normalize_add_desc(desc)
            elif op.lower() == "overlap" and not desc:
                desc = "Overlap Major Non-Adj. Panel"

            rows.append({
                "Line": line_no,
                "Qty": "",
                "Part Number": "",
                "Operation": op,
                "Description": desc,
                "Hours": "",
            })

        # Priority1 often has blank assigned hours. Keep the total row anyway.
        total_match = re.search(re.escape(total_marker) + r"\s+(-?\d+(?:\.\d+)?)", raw)
        total_hours = total_match.group(1) if total_match else ""
        _append_total_row(rows, total_label, total_hours)

    df = pd.DataFrame(rows)
    return _append_translations(df)


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
        if m_header and not re.search(r"\b(Repair|Remove|Replace|Blend|Add|Overlap|NIB)\b", l, re.IGNORECASE):
            rows.append({
                "Line": int(m_header.group(1)),
                "Qty": "",
                "Part Number": "",
                "Operation": "",
                "Description": m_header.group(2).strip(),
                "Hours": "",
            })
            i += 1
            continue

        # Merge wrapped lines until next numbered row/section or subtotal.
        row_text = l
        j = i + 1
        while j < len(lines):
            nxt = lines[j]
            if nxt.startswith("Subtotals") or nxt.startswith("Grand Total"):
                break
            if re.match(r"^\d+\s+", nxt):
                break
            row_text += " " + nxt
            j += 1

        # Body/paint row with operation and part number.
        m = re.match(
            r"^(\d+)\s+(Repair|Remove / Install|Remove / Replace|Replace|Blend|Add|Overlap)\s+(\d+)\s+(.+)$",
            row_text,
            flags=re.IGNORECASE,
        )

        # Paint rows without explicit operation/part number:
        # 4 0 Add for Clear Coat Refinish / Paint 1.7
        m_no_op = None
        if not m:
            m_no_op = re.match(r"^(\d+)\s+(\d+)\s+(.+)$", row_text)

        if m:
            line_no = int(m.group(1))
            op = _clean_spaces(m.group(2))
            qty = int(m.group(3))
            rest = _clean_spaces(m.group(4))

            mh = re.search(r"(-?\d+\.\d+)\s*$", rest)
            hours = float(mh.group(1)) if mh else ""
            before_hours = rest[:mh.start()].strip() if mh else rest

            before_hours = re.split(r"\s+(Body|Refinish / Paint)\s+", before_hours)[0].strip()
            before_hours = re.sub(r"\bOEM\b\s*$", "", before_hours).strip()

            # Part number is the first alphanumeric token, optionally followed by a short suffix.
            pm = re.match(r"^([A-Z0-9]+(?:\s+[A-Z]{1,3})?)\s+(.+)$", before_hours)
            if pm:
                part_number = _clean_spaces(pm.group(1))
                desc = _clean_spaces(pm.group(2))
            else:
                part_number = ""
                desc = before_hours

        elif m_no_op:
            line_no = int(m_no_op.group(1))
            qty = int(m_no_op.group(2))
            rest = _clean_spaces(m_no_op.group(3))
            part_number = ""

            mh = re.search(r"(-?\d+\.\d+)\s*$", rest)
            hours = float(mh.group(1)) if mh else ""
            desc = rest[:mh.start()].strip() if mh else rest
            desc = re.split(r"\s+(Body|Refinish / Paint)\s+", desc)[0].strip()
            desc = re.sub(r"\bOEM\b\s*$", "", desc).strip()

            if desc.lower().startswith("add for"):
                op = "Add"
                desc = _normalize_add_desc(desc)
            elif desc.lower().startswith("overlap"):
                op = "Overlap"
            else:
                op = ""

        else:
            i += 1
            continue

        rows.append({
            "Line": line_no,
            "Qty": qty,
            "Part Number": part_number,
            "Operation": op,
            "Description": desc,
            "Hours": hours,
        })

        i = j

    total_hours = _extract_total_hours(text)
    _append_total_row(rows, "TOTAL LABOR HOURS", total_hours)

    df = pd.DataFrame(rows)
    if not df.empty:
        df = df.sort_values(
            by=["Line"],
            key=lambda s: pd.to_numeric(s, errors="coerce").fillna(10**9),
            kind="stable",
        ).reset_index(drop=True)

    return _append_translations(df)


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
