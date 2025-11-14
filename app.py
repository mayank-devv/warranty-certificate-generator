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
st.caption("Paste extracted details + upload template → auto-generate warranty certificate.")

# -------------------------------
# Inputs
# -------------------------------
raw_text = st.text_area("Paste Extracted Details (Option A format)", height=350)
template_file = st.file_uploader("Upload Warranty Template (.docx)", type=["docx"])

# -------------------------------
# Helpers
# -------------------------------
BLUE = RGBColor(0, 112, 192)

def center_align(para):
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in para.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = BLUE

def add_line(p):
    pPr = p._p.get_or_add_pPr()
    xml = r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>' % nsdecls("w")
    pPr.append(parse_xml(xml))

def parse_block(text):
    data = {}
    for line in text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            data[k.strip()] = v.strip()
    return data


# -------------------------------
# ADDRESS CLEANER (NEW)
# -------------------------------
def clean_address(addr):
    if not addr:
        return ""

    parts = [p.strip() for p in addr.replace("\n", ",").split(",") if p.strip()]

    clean_lines = []
    for part in parts:
        # Keep acronyms like CWC uppercase
        if part.isupper():
            clean_lines.append(part)
        else:
            clean_lines.append(part.title())

    return "\n".join(clean_lines)


# -------------------------------
# MAIN
# -------------------------------
if st.button("Generate Certificate"):
    if not template_file:
        st.error("Upload a DOCX template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste extracted details.")
        st.stop()

    details = parse_block(raw_text)

    # Clean address to multi-line
    cleaned_address = clean_address(details.get("Address", ""))

    # ALWAYS TODAY
    today = datetime.now().strftime("%d-%m-%Y")

    # Placeholder mapping (FULL + CASE-INSENSITIVE)
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
        "{Warranty on Compressor}": details.get("Warranty on Compressor", ""),
        "{warranty on compressor}": details.get("Warranty on Compressor", ""),
        "{CustomerName}": details.get("Customer Name", ""),
        "{Organisation}": details.get("Organisation", ""),
        "{Address}": cleaned_address,  # <-- CLEANED MULTI-LINE ADDRESS
        "{GEMContractNo}": details.get("GEM Contract No", ""),
        "{Date}": today,
    }

    doc = Document(template_file)

    # Narrow margins
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # REPLACE placeholders everywhere
    for p in doc.paragraphs:
        newt = p.text
        for k, v in mapping.items():
            newt = newt.replace(k, v)
        p.text = newt

    # AFTER REPLACEMENT — apply basic styling
    for p in doc.paragraphs:
        for run in p.runs:
            run.font.name = "Calibri"
            run.font.size = Pt(12)
            run.font.color.rgb = BLUE

    # --------------------------
    # FIX ADDRESS MULTI-LINE
    # --------------------------
    for p in doc.paragraphs:
        if cleaned_address in p.text:
            p.clear()
            for line in cleaned_address.split("\n"):
                r = p.add_run(line)
                r.font.name = "Calibri"
                r.font.size = Pt(12)
                r.font.color.rgb = BLUE
                p.add_run("\n")
            break

    # -------------------------------
    # CENTER LETTERHEAD (top 5 non-empty)
    # -------------------------------
    count = 0
    for p in doc.paragraphs:
        if p.text.strip():
            center_align(p)
            count += 1
            if count == 5:
                break

    # -------------------------------
    # CENTER WARRANTY CERTIFICATE
    # -------------------------------
    for p in doc.paragraphs:
        if "WARRANTY CERTIFICATE" in p.text.upper():
            center_align(p)
            for run in p.runs:
                run.font.size = Pt(16)
                run.font.bold = True
                run.font.underline = True

    # -------------------------------
    # RIGHT ALIGN DATE
    # -------------------------------
    for p in doc.paragraphs:
        if "Date:" in p.text:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            break

    # -------------------------------
    # BLUE LINES
    # -------------------------------
    for i, p in enumerate(doc.paragraphs):
        if "Email" in p.text or "@" in p.text:
            NP = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(NP)
            break

    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            NP = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(NP)
            break

    # -------------------------------
    # OUTPUT
    # -------------------------------
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)

    fname = f"Warranty_{details.get('Customer Name','Customer').replace(' ','_')}_{details.get('GEM Contract No','GEM')}.docx"

    st.success("✅ Certificate Generated Successfully!")
    st.download_button("⬇ Download Certificate", buf, fname)
