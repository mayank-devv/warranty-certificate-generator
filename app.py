import io
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# -------------------------------
# Streamlit Page Setup
# -------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Paste extracted details + upload DOCX template → auto-generate formatted certificate.")

# -------------------------------
# Input (Paste the block)
# -------------------------------
raw_text = st.text_area(
    "Paste Extracted Details (Option A format)",
    height=350
)

# -------------------------------
# Upload DOCX Template
# -------------------------------
template_file = st.file_uploader("Upload Warranty Certificate DOCX Template", type=["docx"])

# -------------------------------
# Utility functions
# -------------------------------
BLUE = RGBColor(0, 112, 192)

def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = parse_xml(
        r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>'
        % nsdecls("w")
    )
    pPr.append(pBdr)

def render_styled_text(p):
    """Apply blue Calibri formatting WITHOUT changing alignment"""
    for run in p.runs:
        run.font.name = "Calibri"
        run.font.size = Pt(12)
        run.font.color.rgb = BLUE

def render_label_value(p):
    """Bold label: value styling"""
    full_text = p.text
    if ":" in full_text:
        label, _, value = full_text.partition(":")
        p.clear()
        r1 = p.add_run(label.strip() + ":")
        r1.font.bold = True
        r1.font.color.rgb = BLUE
        r1.font.name = "Calibri"

        r2 = p.add_run(" " + value.strip())
        r2.font.bold = False
        r2.font.color.rgb = BLUE
        r2.font.name = "Calibri"
    else:
        render_styled_text(p)

def parse_block(text):
    data = {}
    for line in text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            data[k.strip()] = v.strip()
    return data

# -------------------------------
# MAIN
# -------------------------------
if st.button("Generate Certificate"):
    if not template_file:
        st.error("Upload a DOCX template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste extracted block.")
        st.stop()

    details = parse_block(raw_text)

    # Map placeholders
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
        "{Date}": details.get("Date", "")
    }

    doc = Document(template_file)

    # Narrow margins
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # -------------------------------
    # REPLACE PLACEHOLDERS EVERYWHERE
    # -------------------------------
    for p in doc.paragraphs:
        text = p.text
        for k, v in mapping.items():
            text = text.replace(k, v)
        p.text = text  # direct replace without styling

    # -------------------------------
    # NOW re-apply styling to non-header paragraphs
    # -------------------------------
    for idx, p in enumerate(doc.paragraphs):
        if idx < 7:
            continue  # header formatting done later
        render_label_value(p)

    # -------------------------------
    # FIX HEADER (CENTER + COLOR + SIZE)
    # -------------------------------
    for i in range(0, 7):
        if i < len(doc.paragraphs):
            para = doc.paragraphs[i]
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.font.color.rgb = BLUE
                run.font.name = "Calibri"
                run.font.size = Pt(22 if i == 0 else 12)
                run.font.bold = True if i == 0 else False

    # -------------------------------
    # WARRANTY CERTIFICATE TITLE
    # -------------------------------
    for p in doc.paragraphs:
        if "WARRANTY CERTIFICATE" in p.text.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.underline = True
                run.font.color.rgb = BLUE

    # -------------------------------
    # BLUE LINES
    # -------------------------------
    # Line under letterhead
    for i, p in enumerate(doc.paragraphs):
        if "Email" in p.text or "@" in p.text:
            newp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_horizontal_line(newp)
            break

    # Line below GEM Contract No
    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            newp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_horizontal_line(newp)
            break

    # -------------------------------
    # SAVE OUTPUT
    # -------------------------------
    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)

    fname = details.get("Customer Name", "Customer").replace(" ", "_")
    gem = details.get("GEM Contract No", "GEM").replace(" ", "_")
    fname_dl = f"Warranty_{fname}_{gem}.docx"

    st.success("✅ Certificate generated successfully!")
    st.download_button(
        "⬇️ Download Certificate",
        data=out_buf,
        file_name=fname_dl,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
