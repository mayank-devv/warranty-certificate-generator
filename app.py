import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------
# STREAMLIT CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Upload template + paste extracted details → auto-generate warranty certificate.")

# ---------------------------------------------------------
# INPUTS
# ---------------------------------------------------------
raw_text = st.text_area("Paste Extracted Details (Option A format)", height=350)
template_file = st.file_uploader("Upload Warranty Certificate Template (.docx)", type=["docx"])

# ---------------------------------------------------------
# HELPERS
# ---------------------------------------------------------
BLUE = RGBColor(0, 112, 192)

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

def style_label_value(p):
    t = p.text
    if ":" in t:
        p.clear()
        label, _, value = t.partition(":")
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
            r.font.color.rgb = BLUE
            r.font.name = "Calibri"
            r.font.size = Pt(12)

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if st.button("Generate Certificate"):

    if not template_file:
        st.error("Upload a DOCX template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste the extracted details block.")
        st.stop()

    details = parse_block(raw_text)

    # Always today's date
    today = datetime.now().strftime("%d-%m-%Y")

    # -------------------------------
    # MULTILINE ADDRESS CLEANING
    # -------------------------------
    raw_address = details.get("Address", "").strip()
    addr = " ".join(raw_address.split())

    break_keywords = [
        "DIVISION","CWC","OPP.","NEAR","BEHIND","OFFICE","BUILDING",
        "FLOOR","ROAD","MARG","AREA","COLONY"
    ]

    parts = addr.replace(",", ", ").split(",")
    lines = []
    temp = ""

    for p in parts:
        seg = p.strip()
        if any(seg.upper().startswith(k) for k in break_keywords):
            lines.append(seg)
            continue
        if len(seg) > 30:
            lines.append(seg)
            continue
        if not temp:
            temp = seg
        else:
            temp += ", " + seg

    if temp:
        lines.append(temp)

    final_address = "\n".join(lines)

    # -------------------------------
    # MAPPING FOR NEW TEMPLATE
    # -------------------------------
    mapping = {
        "{Company}": details.get("Company", ""),

        "{Brand}": details.get("Brand", ""),
        "{Make}": details.get("Brand", ""),

        "{Category}": details.get("Category", ""),
        "{ProductName}": details.get("Product Name", ""),

        "{Model}": "",    # TEMPLATE HAS FIXED TEXT
        "{SerialNumber}": "",  # TEMPLATE HAS FIXED TEXT

        "{Quantity}": details.get("Quantity", ""),

        "{Warranty}": details.get("Warranty", ""),
        "{WarrantyOnCompressor}": details.get("Warranty on Compressor", ""),
        "{Warranty on Compressor}": details.get("Warranty on Compressor", ""),
        "{warranty on compressor}": details.get("Warranty on Compressor", ""),

        "{CustomerName}": details.get("Customer Name", ""),

        # CRITICAL FIX → Always add a newline between Org + Address
        "{Organisation}": details.get("Organisation", "") + "\n",
        "{Address}": final_address,

        "{GEMContractNo}": details.get("GEM Contract No", ""),
        "{Date}": today,
    }

    # -------------------------------
    # PROCESS DOCUMENT
    # -------------------------------
    doc = Document(template_file)

    # Narrow margins
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # Replace placeholders
    for p in doc.paragraphs:
        text = p.text
        for k, v in mapping.items():
            text = text.replace(k, v)
        p.text = text

    # Apply label styling to non-header paragraphs
    for i, p in enumerate(doc.paragraphs):
        if i >= 7:
            style_label_value(p)

    # -------------------------------
    # MAKE LETTERHEAD CENTERED
    # -------------------------------
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5, len(non_empty))):
        pe = non_empty[i]
        pe.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in pe.runs:
            r.font.color.rgb = BLUE
            r.font.name = "Calibri"
            r.font.size = Pt(22 if i == 0 else 12)
            r.font.bold = True if i == 0 else False

    # -------------------------------
    # CENTER WARRANTY CERTIFICATE
    # -------------------------------
    for p in doc.paragraphs:
        if "WARRANTY CERTIFICATE" in p.text.upper():
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.size = Pt(16)
                r.font.bold = True
                r.font.underline = True
                r.font.color.rgb = BLUE

    # -------------------------------
    # DATE MUST BE RIGHT-ALIGNED
    # -------------------------------
    for p in doc.paragraphs:
        if "Date:" in p.text:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            break

    # -------------------------------
    # BLUE LINES
    # -------------------------------
    # Line under header
    for i, p in enumerate(doc.paragraphs):
        if "Email" in p.text or "@" in p.text:
            pp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(pp)
            break

    # Line under GEM Contract
    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            pp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(pp)
            break

    # -------------------------------
    # OUTPUT FILE
    # -------------------------------
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    fname = details.get("Customer Name", "Customer").replace(" ", "_")
    g = details.get("GEM Contract No", "GEM").replace(" ", "_")
    fn = f"Warranty_{fname}_{g}.docx"

    st.success("✅ Warranty Certificate Generated Successfully!")
    st.download_button("⬇ Download Certificate", out, fn)
