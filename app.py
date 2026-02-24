import re
import math
import io
import zipfile
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

import pandas as pd
import streamlit as st

from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm


# =========================
# BASIC HELPERS
# =========================

def norm(x) -> str:
    if x is None or pd.isna(x):
        return ""
    s = str(x)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return s.strip()

def norm_lower(x) -> str:
    return norm(x).lower()

def to_number(x):
    if x is None or pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = norm(x)
    if s == "":
        return None
    s = s.replace(",", "")
    try:
        return float(s)
    except:
        return None

def safe_filename(text):
    text = norm(text)
    text = re.sub(r"[^\w\-]+", "_", text)
    return text or "file"

def is_meta_row(ship_to: str) -> bool:
    s = norm_lower(ship_to)
    if s in ["", "total", "grand total", "subtotal"]:
        return True
    if s.startswith("total"):
        return True
    if "packing size" in s:
        return True
    if "cost center" in s:
        return True
    return False


# =========================
# LABEL STRUCTURE
# =========================

@dataclass
class LabelLine:
    sender: str
    ship_to: str
    contact: str
    program: str
    product: str
    total_qty: int
    carton_i: int
    carton_n: int
    qty_in_carton: int
    packing_size: int


# =========================
# PDF RENDER
# =========================

def render_pdf(labels: List[LabelLine]) -> bytes:
    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=20*mm,
        rightMargin=20*mm,
        topMargin=20*mm,
        bottomMargin=20*mm
    )

    styles = getSampleStyleSheet()
    big = ParagraphStyle("big", parent=styles["Normal"], fontSize=22, leading=28)
    mid = ParagraphStyle("mid", parent=styles["Normal"], fontSize=18, leading=24)
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=14, leading=18)

    story = []

    for i, lab in enumerate(labels):
        story.append(Paragraph(f"<b>Sender:</b> {lab.sender}", big))
        story.append(Spacer(1, 10))

        story.append(Paragraph(f"<b>Ship to:</b> {lab.ship_to}", mid))
        story.append(Spacer(1, 6))

        contact_html = lab.contact.replace("\n", "<br/>") if lab.contact else "(blank)"
        story.append(Paragraph("<b>Contact & Address:</b>", mid))
        story.append(Paragraph(contact_html, small))
        story.append(Spacer(1, 12))

        title = " - ".join([x for x in [lab.program, lab.product] if x])
        story.append(Paragraph(f"<b>{title}</b> - Total Qty: {lab.total_qty}", mid))
        story.append(Spacer(1, 12))

        story.append(Paragraph(f"<b>Carton:</b> {lab.carton_i} / {lab.carton_n}", big))
        story.append(Paragraph(f"<b>Qty in Carton:</b> {lab.qty_in_carton}", big))
        story.append(Spacer(1, 12))

        story.append(Paragraph(f"Packing size: {lab.packing_size} pcs/carton", mid))

        if i < len(labels) - 1:
            story.append(PageBreak())

    doc.build(story)
    return buffer.getvalue()


# =========================
# EXCEL PARSE
# =========================

def find_position(matrix, condition):
    for r, row in enumerate(matrix):
        for c, val in enumerate(row):
            if condition(norm_lower(val)):
                return r, c
    return None

def parse_excel(file_bytes) -> List[LabelLine]:
    df = pd.read_excel(io.BytesIO(file_bytes), header=None)
    matrix = df.values.tolist()

    header_pos = find_position(matrix, lambda s: s == "country / service")
    if not header_pos:
        raise ValueError("Cannot find 'COUNTRY / SERVICE'")

    header_row, cs_col = header_pos

    packing_pos = find_position(matrix, lambda s: "packing size" in s)
    if not packing_pos:
        raise ValueError("Cannot find 'Packing size' row")
    packing_row, _ = packing_pos

    contact_pos = find_position(matrix, lambda s: "shipping" in s and "contact" in s)
    contact_col = contact_pos[1] if contact_pos else None

    # product columns = only numeric packing size
    product_cols: Dict[int, int] = {}
    for col, val in enumerate(matrix[packing_row]):
        num = to_number(val)
        if num and num > 0:
            product_cols[col] = int(num)

    labels: List[LabelLine] = []

    for r in range(packing_row + 1, len(matrix)):
        if cs_col >= len(matrix[r]):
            continue

        ship_to = norm(matrix[r][cs_col])
        if is_meta_row(ship_to):
            continue

        contact = ""
        if contact_col is not None and contact_col < len(matrix[r]):
            contact = norm(matrix[r][contact_col])

        for col, pack_size in product_cols.items():
            if col >= len(matrix[r]):
                continue

            qty = to_number(matrix[r][col])
            if not qty or qty <= 0:
                continue

            qty = int(qty)
            carton_n = math.ceil(qty / pack_size)

            program = norm(matrix[packing_row - 2][col]) if packing_row >= 2 else ""
            product = norm(matrix[packing_row - 1][col]) if packing_row >= 1 else ""

            if norm_lower(program) == "nan":
                program = ""
            if norm_lower(product) == "nan":
                product = ""

            for i in range(1, carton_n + 1):
                if i < carton_n:
                    qty_carton = pack_size
                else:
                    qty_carton = qty - pack_size * (carton_n - 1)

                labels.append(LabelLine(
                    sender="PM STUDIO",
                    ship_to=ship_to,
                    contact=contact,
                    program=program,
                    product=product,
                    total_qty=qty,
                    carton_i=i,
                    carton_n=carton_n,
                    qty_in_carton=qty_carton,
                    packing_size=pack_size
                ))

    if not labels:
        raise ValueError("No labels generated")
    return labels


# =========================
# GROUP -> MULTI PDF -> ZIP
# =========================

def group_by_program_product(labels: List[LabelLine]) -> Dict[Tuple[str, str], List[LabelLine]]:
    groups: Dict[Tuple[str, str], List[LabelLine]] = {}
    for lab in labels:
        key = (lab.program, lab.product)
        groups.setdefault(key, []).append(lab)
    return groups

def build_zip_of_pdfs(groups: Dict[Tuple[str, str], List[LabelLine]]) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for (program, product), items in groups.items():
            # keep order stable
            items = sorted(items, key=lambda x: (x.ship_to, x.carton_i))
            pdf_bytes = render_pdf(items)

            name_part = " - ".join([x for x in [program, product] if x])
            filename = safe_filename(name_part) + ".pdf"
            zf.writestr(filename, pdf_bytes)
    return out.getvalue()


# =========================
# STREAMLIT UI
# =========================

st.title("📦 Carton Label PDF Generator (1 product = 1 PDF)")

uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        labels = parse_excel(uploaded.read())
        groups = group_by_program_product(labels)

        st.success(f"Generated {len(labels)} labels → {len(groups)} PDFs (one per Program-Product).")

        zip_bytes = build_zip_of_pdfs(groups)

        st.download_button(
            "Download ZIP (PDFs by Product)",
            zip_bytes,
            file_name="carton_labels_by_product.zip",
            mime="application/zip"
        )

    except Exception as e:
        st.error(str(e))
