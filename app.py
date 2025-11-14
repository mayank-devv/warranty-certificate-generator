import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# -------------------------------
# Streamlit Setup
# -------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Paste extracted details + upload DOCX template → auto-generate formatted warranty certificate.")

# -------------------------------
# Inputs
# -------------------------------
raw_text = st.text_area("Paste Extracted Details (Option A format)", height=350)
template_file = st.file_uploader("Upload Warranty Certificate DOCX Template", type=["docx"])

# -------------------------------
# Helpers
# -------------------------------
BLUE = RGBColor(0, 112, 192)

def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    xml = r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>' % nsdecls("w")
    pPr.append(parse_xml(xml))

def parse_block(text):
    data = {}
    for line in text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            data[k.strip()] = v.strip()
    return data

def style_default(p):
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(12)
        run.font.color.rgb = BLUE

def style_label_value(p):
    full = p.text
    if ":" in full:
        label, _, value = full.partition(":")
        p.clear()

        r1 = p.add_run(label.strip() + ":")
        r1.font.bold = True
        r1.font.name = "Calibri"
        r1.font.size = Pt(12)
        r1.font.color.rgb = BLUE

        r2 = p.add_run(" " + value.strip())
        r2.font.bold = False
        r2.font.name = "Calibri"
        r2.font.size = Pt(12)
        r2.font.color.rgb = BLUE
    else:
        style_default(p)

# -------------------------------
# MAIN
# -------------------------------
if st.button("Generate Certificate"):

    if not template_file:
        st.error("Upload a DOCX template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste the details block.")
        st.stop()

    details = parse_block(raw_text)

    # FORCE TODAY'S DATE ALWAYS
    today_str = datetime.now().strftime("%d-%m-%Y")

    mapping = {
        "{Company}": details.get("Company", ""),
        "{Brand}": details.get("Brand", ""),
        "{Make}": details.get("Brand", ""),
        "{Category}": details.get("Category", ""),
        "{ProductName}": details.get("Product Name", ""),
        "{Model}": details.get("Model", ""),
        "{Quantity}": details.get("Quantity", ""),
        "{SerialNumber}": details.get("Serial Number", ""),
        "{Warranty}": details.get("Warranty", ""),
        "{WarrantyOnCompressor}": details.get("Warranty on Compressor", ""),
        "{CustomerName}": details.get("Customer Name", ""),
        "{Organisation}": details.get("Organisation", ""),
        "{Address}": details.get("Address", ""),
        "{GEMContractNo}": details.get("GEM Contract No", ""),
        "{Date}": today_str   # ALWAYS TODAY
    }

    doc = Document(template_file)

    # Narrow margins
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # Replace placeholders (EVERYWHERE)
    for p in doc.paragraphs:
        t = p.text
        for k, v in mapping.items():
            t = t.replace(k, v)
        p.text = t

    # Re-style all non-header paragraphs
    for i, p in enumerate(doc.paragraphs):
        if i >= 7:
            style_label_value(p)

    # --------------------------
    # LETTERHEAD CENTER → ALWAYS
    # --------------------------
    for i in range(0, 7):
        if i < len(doc.paragraphs):
            p = doc.paragraphs[i]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.color.rgb = BLUE
                run.font.name = "Calibri"
                run.font.bold = True if i == 0 else False
                run.font.size = Pt(22) if i == 0 else Pt(12)

    # --------------------------
    # CENTER “WARRANTY CERTIFICATE”
    # --------------------------
    for p in doc.paragraphs:
        if "WARRANTY CERTIFICATE" in p.text.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.underline = True
                run.font.color.rgb = BLUE

    # --------------------------
    # BLUE LINES
    # --------------------------
    # Under letterhead
    for i, p in enumerate(doc.paragraphs):
        if "@" in p.text or "Email" in p.text:
            np = doc.paragraphs[i+1].insert_paragraph_before("")
            add_horizontal_line(np)
            break

    # Under GEM Contract No
    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            np = doc.paragraphs[i+1].insert_paragraph_before("")
            add_horizontal_line(np)
            break

    # --------------------------
    # Output
    # --------------------------
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    fname = details.get("Customer Name", "Customer").replace(" ", "_")
    gem = details.get("GEM Contract No", "GEM").replace(" ", "_")
    filename = f"Warranty_{fname}_{gem}.docx"

    st.success("✅ Certificate generated!")
    st.download_button("⬇️ Download", buffer, filename)
