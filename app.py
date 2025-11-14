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

def style_blue(p):
    for r in p.runs:
        r.font.name = "Calibri"
        r.font.size = Pt(12)
        r.font.color.rgb = BLUE

def style_label_value(p):
    t = p.text
    if ":" in t:
        p.clear()
        label, _, value = t.partition(":")
        r1 = p.add_run(label.strip() + ":")
        r1.font.bold = True
        r1.font.color.rgb = BLUE
        r1.font.name = "Calibri"
        r1.font.size = Pt(12)

        r2 = p.add_run(" " + value.strip())
        r2.font.bold = False
        r2.font.color.rgb = BLUE
        r2.font.name = "Calibri"
        r2.font.size = Pt(12)
    else:
        style_blue(p)

def center_header(p, big=False):
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in p.runs:
        r.font.name = "Calibri"
        r.font.color.rgb = BLUE
        if big:
            r.font.size = Pt(22)
            r.font.bold = True
        else:
            r.font.size = Pt(12)

# ---------------------------------------------------------
# MAIN PROCESS
# ---------------------------------------------------------
if st.button("Generate Certificate"):

    if not template_file:
        st.error("Upload a DOCX template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste the extracted details block.")
        st.stop()

    details = parse_block(raw_text)

    # ALWAYS TODAY DATE
    today = datetime.now().strftime("%d-%m-%Y")

    # -------------------------------
    # CLEAN MULTILINE ADDRESS
    # -------------------------------
    raw_address = details.get("Address", "").strip()
    addr = " ".join(raw_address.split())

    break_keywords = ["DIVISION","CWC","OPP.","NEAR","BEHIND","OFFICE","BUILDING","FLOOR","ROAD","MARG","AREA","COLONY"]

    parts = addr.replace(",", ", ").split(",")
    clean = []
    temp = ""

    for p in parts:
        seg = p.strip()
        if any(seg.upper().startswith(k) for k in break_keywords):
            clean.append(seg)
            continue
        if len(seg) > 30:
            clean.append(seg)
            continue
        if not temp:
            temp = seg
        else:
            temp += ", " + seg

    if temp:
        clean.append(temp)

    final_address = "\n".join(clean)

    # -------------------------------
    # MAPPING
    # -------------------------------
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

        # 🔥 FIX: ALWAYS add line break after Organisation
        "{Organisation}": details.get("Organisation", "") + "\n",
        "{Address}": final_address,

        "{GEMContractNo}": details.get("GEM Contract No", ""),
        "{Date}": today,
    }

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
        t = p.text
        for k, v in mapping.items():
            t = t.replace(k, v)
        p.text = t

    # Apply styling
    for i, p in enumerate(doc.paragraphs):
        if i >= 7:
            style_label_value(p)

    # -------------------------------
    # FIX LETTERHEAD CENTER
    # -------------------------------
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5, len(non_empty))):
        center_header(non_empty[i], big=(i == 0))

    # -------------------------------
    # FIX WARRANTY CERTIFICATE HEADING
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
    # DATE → ALWAYS RIGHT
    # -------------------------------
    for p in doc.paragraphs:
        if "Date:" in p.text:
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            break

    # -------------------------------
    # ADD BLUE LINES
    # -------------------------------
    # Under header
    for i, p in enumerate(doc.paragraphs):
        if "@" in p.text or "Email" in p.text:
            n = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(n)
            break

    # Under GEM Contract
    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            n = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(n)
            break

    # -------------------------------
    # OUTPUT FILE
    # -------------------------------
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    fname = details.get("Customer Name", "Customer").replace(" ", "_")
    g = details.get("GEM Contract No", "GEM").replace(" ", "_")
    file = f"Warranty_{fname}_{g}.docx"

    st.success("✅ Warranty Certificate Generated Successfully!")
    st.download_button("⬇ Download Certificate", out, file)
