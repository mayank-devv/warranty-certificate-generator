import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls


# ============================================================
# XML ALIGNMENT HELPERS (bulletproof)
# ============================================================
def align_xml(obj, val):
    obj._element.get_or_add_pPr().append(
        parse_xml(
            f'<w:jc w:val="{val}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'  # noqa
        )
    )


def cell_align(cell, val):
    for p in cell.paragraphs:
        align_xml(p, val)


def color_blue(cell):
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = RGBColor(0, 112, 192)
            r.font.name = "Calibri"
            r.font.size = Pt(12)


# ============================================================
# STREAMLIT UI
# ============================================================
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")

raw_text = st.text_area("Paste Extracted Details (Option A Block):", height=250)
template_file = st.file_uploader("Upload Warranty Template (.docx)", type=["docx"])


# ============================================================
# MAIN PROCESS
# ============================================================
if st.button("Generate Certificate"):

    if not template_file:
        st.error("Upload a template DOCX.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste extracted block.")
        st.stop()

    today = datetime.now().strftime("%d-%m-%Y")

    # --------------------
    # Parse block
    # --------------------
    data = {}
    for line in raw_text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            data[k.strip()] = v.strip()

    # --------------------
    # Clean Address
    # --------------------
    raw_addr = data.get("Address", "")
    raw_addr = " ".join(raw_addr.split())  # remove weird spacing
    parts = raw_addr.replace(",", ", ").split(",")

    addr_lines = []
    buffer = ""
    for seg in parts:
        seg = seg.strip()
        if len(seg) > 40:
            addr_lines.append(seg)
        else:
            if not buffer:
                buffer = seg
            else:
                buffer += ", " + seg
    if buffer:
        addr_lines.append(buffer)

    # --------------------
    # Load Template
    # --------------------
    doc = Document(template_file)

    # Narrow margins
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # Replace simple placeholders that already exist
    mapping = {
        "{Company}": data.get("Company", ""),
        "{CustomerName}": data.get("Customer Name", ""),
        "{GEMContractNo}": data.get("GEM Contract No", ""),
        "{Date}": today,
    }

    for p in doc.paragraphs:
        text = p.text
        for k, v in mapping.items():
            text = text.replace(k, v)
        p.text = text

    # ============================================================
    # REMOVE OLD CUSTOMER BLOCK (4 lines)
    # ============================================================
    c_index = None
    for i, p in enumerate(doc.paragraphs):
        if "Customer" in p.text:
            c_index = i
            break

    if c_index is None:
        c_index = 7  # fallback

    # Delete 4 old paragraphs
    for _ in range(4):
        if c_index < len(doc.paragraphs):
            el = doc.paragraphs[c_index]._p
            parent = el.getparent()
            parent.remove(el)

    # ============================================================
    # CREATE NEW CUSTOMER BLOCK (TABLE-ONLY)
    # ============================================================

    # Insert table BEFORE c_index
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    tbl = table._tbl
    doc.paragraphs[c_index]._p.addprevious(tbl)

    # Row 1: Customer (left) | Date (right)
    left_cell = table.rows[0].cells[0]
    right_cell = table.rows[0].cells[1]

    left_cell.text = f"Customer: {data.get('Customer Name', '')}"
    right_cell.text = f"Date: {today}"

    cell_align(left_cell, "left")
    cell_align(right_cell, "right")
    color_blue(left_cell)
    color_blue(right_cell)

    # --------------------
    # Row 2: Organisation
    # --------------------
    row2 = table.add_row()
    cell_org = row2.cells[0]
    cell_org.merge(row2.cells[1])
    cell_org.text = data.get("Organisation", "")
    cell_align(cell_org, "left")
    color_blue(cell_org)

    # --------------------
    # Row 3+: Address lines
    # --------------------
    for line in addr_lines:
        row = table.add_row()
        c = row.cells[0]
        c.merge(row.cells[1])
        c.text = line
        cell_align(c, "left")
        color_blue(c)

    # ============================================================
    # FIX LETTERHEAD (CENTER TOP 5 PARAGRAPHS)
    # ============================================================
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5, len(non_empty))):
        p = non_empty[i]
        align_xml(p, "center")
        for r in p.runs:
            r.font.color.rgb = RGBColor(0,112,192)
            r.font.name = "Calibri"
            r.font.size = Pt(22 if i==0 else 12)
            r.font.bold = (i == 0)

    # ============================================================
    # WARRANTY CERTIFICATE TITLE
    # ============================================================
    for p in doc.paragraphs:
        if p.text.strip().upper() == "WARRANTY CERTIFICATE":
            align_xml(p, "center")
            for r in p.runs:
                r.font.bold = True
                r.font.underline = True
                r.font.size = Pt(16)
                r.font.color.rgb = RGBColor(0,112,192)

    # ============================================================
    # OUTPUT
    # ============================================================
    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    st.success("Warranty Certificate Generated Successfully!")
    st.download_button("⬇ Download DOCX", output, f"Warranty_{today}.docx")
