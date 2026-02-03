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
            "Exterior Color": r"Exterior Color
