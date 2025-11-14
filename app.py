import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------
# STREAMLIT PAGE SETUP
# ---------------------------------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Upload DOCX template + paste extracted GEMC details → auto-generate final certificate.")

# ---------------------------------------------------------
# USER INPUTS
# ---------------------------------------------------------
raw_text = st.text_area("Paste Extracted Details (Option A format)", height=350)
template_file = st.file_uploader("Upload Warranty Certificate Template (.docx)", type=["docx"])


# ---------------------------------------------------------
# HELPER FUNCTIONS
# ---------------------------------------------------------
BLUE = RGBColor(0, 112, 192)

def add_line(paragraph):
    """Adds a blue horizontal line under a paragraph"""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    xml = r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>' % nsdecls("w")
    pPr.append(parse_xml(xml))

def center_align(para):
    """Center align + apply header-style formatting"""
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in para.runs:
        run.font.name = "Calibri"
        run.font.color.rgb = BLUE

def parse_block(text):
    """Convert 'Label: Value' lines into dict"""
    data = {}
    for line in text.split("\n"):
        if ":" in line:
            key, _, value = line.partition(":")
            data[key.strip()] = value.strip()
    return data

def style_label_value(p):
    """Bold label: value blue formatting"""
    text = p.text
    if ":" in text:
        p.clear()
        label, _, value = text.partition(":")
        r1 = p.add_run(label.strip() + ":")
        r1.font.bold = True
        r1.font.name = "Calibri"
        r1.font.color.rgb = BLUE
        r1.font.size = Pt(12)

        r2 = p.add_run(" " + value.strip())
        r2.font.bold = False
        r2.font.name = "Calibri"
        r2.font.color.rgb = BLUE
        r2.font.size = Pt(12)
    else:
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.color.rgb = BLUE
            r.font.size = Pt(12)


# ---------------------------------------------------------
# MAIN GENERATION
# ---------------------------------------------------------
if st.button("Generate Certificate"):

    # --- VALIDATIONS ---
    if not template_file:
        st.error("Please upload a DOCX template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste the extracted details block.")
        st.stop()

    # Parse pasted text
    details = parse_block(raw_text)

    # ALWAYS USE TODAY'S DATE
    today = datetime.now().strftime("%d-%m-%Y")

    # ---------------------------------------------------------
    # AUTO-CLEAN ADDRESS → MULTI-LINE
    # ---------------------------------------------------------
    raw_address = details.get("Address", "").strip()
    addr = " ".join(raw_address.split())   # normalize spaces

    break_keywords = [
        "DIVISION", "CWC", "OPP.", "NEAR", "BEHIND", "OFFICE", "BUILDING",
        "FLOOR", "ROAD", "MARG", "AREA", "COLONY"
    ]

    parts = addr.replace(",", ", ").split(",")
    clean_lines = []
    temp_line = ""

    for part in parts:
        seg = part.strip()

        # If segment starts with keyword, split immediately
        if any(seg.upper().startswith(k) for k in break_keywords):
            clean_lines.append(seg)
            continue

        # If long segment (>30 chars), put as single line
        if len(seg) > 30:
            clean_lines.append(seg)
            continue

        # Otherwise accumulate shorter segments
        if not temp_line:
            temp_line = seg
        else:
            temp_line += ", " + seg

    if temp_line:
        clean_lines.append(temp_line)

    final_address = "\n".join(clean_lines)

    # ---------------------------------------------------------
    # PLACEHOLDER MAPPING
    # ---------------------------------------------------------
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

        # Warranty on compressor (case-insensitive fixes)
        "{WarrantyOnCompressor}": details.get("Warranty on Compressor", ""),
        "{Warranty on Compressor}": details.get("Warranty on Compressor", ""),
        "{warranty on compressor}": details.get("Warranty on Compressor", ""),

        "{CustomerName}": details.get("Customer Name", ""),
        "{Organisation}": details.get("Organisation", ""),

        # Multiline Address FIX
        "{Address}": final_address,

        "{GEMContractNo}": details.get("GEM Contract No", ""),
        "{Date}": today,
    }

    # ---------------------------------------------------------
    # PROCESS DOCX
    # ---------------------------------------------------------
    doc = Document(template_file)

    # Narrow margins
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Replace placeholders everywhere
    for p in doc.paragraphs:
        t = p.text
        for k, v in mapping.items():
            t = t.replace(k, v)
        p.text = t

    # Style content paragraphs (except letterhead)
    for i, p in enumerate(doc.paragraphs):
        if i >= 7:     # letterhead is first 5–7 lines
            style_label_value(p)

    # ---------------------------------------------------------
    # FIX LETTERHEAD CENTER ALIGNMENT
    # ---------------------------------------------------------
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5, len(non_empty))):
        center_align(non_empty[i])
        # First line company name bigger
        if i == 0:
            for r in non_empty[i].runs:
                r.font.size = Pt(22)
                r.font.bold = True

    # ---------------------------------------------------------
    # FIX WARRANTY CERTIFICATE TITLE CENTER
    # ---------------------------------------------------------
    for p in doc.paragraphs:
        if "WARRANTY CERTIFICATE" in p.text.upper():
            center_align(p)
            for r in p.runs:
                r.font.size = Pt(16)
                r.font.bold = True
                r.font.underline = True

    # ---------------------------------------------------------
    # ALIGN DATE → RIGHT ALWAYS
    # ---------------------------------------------------------
    for p in doc.paragraphs:
        if "Date:" in p.text or "DATE:" in p.text:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            break

    # ---------------------------------------------------------
    # ADD BLUE LINES
    # ---------------------------------------------------------
    # Line below letterhead
    for i, p in enumerate(doc.paragraphs):
        if "@" in p.text or "Email" in p.text:
            ln = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(ln)
            break

    # Line below GEM Contract
    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            ln = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(ln)
            break

    # ---------------------------------------------------------
    # OUTPUT FILE
    # ---------------------------------------------------------
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    fname = details.get("Customer Name", "Customer").replace(" ", "_")
    g = details.get("GEM Contract No", "GEM").replace(" ", "_")
    filename = f"Warranty_{fname}_{g}.docx"

    st.success("✅ Warranty Certificate Generated Successfully!")
    st.download_button("⬇ Download Certificate", out, filename)
