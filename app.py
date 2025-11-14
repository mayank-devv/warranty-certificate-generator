import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ---------------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Upload template + paste extracted details → auto-generate certificate.")

# ---------------------------------------------------------
# INPUT
# ---------------------------------------------------------
raw_text = st.text_area("Paste Extracted Details (Option A format)", height=350)
template_file = st.file_uploader("Upload Warranty Certificate Template (.docx)", type=["docx"])

# ---------------------------------------------------------
# HELPER FUNCS
# ---------------------------------------------------------
BLUE = RGBColor(0, 112, 192)

def add_line(p):
    pPr = p._p.get_or_add_pPr()
    xml = r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>' % nsdecls("w")
    pPr.append(parse_xml(xml))

def parse_block(text):
    out = {}
    for line in text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            out[k.strip()] = v.strip()
    return out

def blue(p):
    for r in p.runs:
        r.font.color.rgb = BLUE
        r.font.name = "Calibri"
        r.font.size = Pt(12)

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if st.button("Generate Certificate"):

    if not template_file:
        st.error("Upload template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste extracted block.")
        st.stop()

    d = parse_block(raw_text)

    today = datetime.now().strftime("%d-%m-%Y")

    # Warranty block (1 paragraph with line breaks)
    warranty_block = (
        "Warranty can be checked anytime by contacting OEM customer care.\n"
        "Warranty is taken care of by OEM as per their terms & conditions. "
        "Original Warranty certificate is to be taken by above if needed."
    )

    # -----------------------------------
    # ADDRESS CLEANING
    # -----------------------------------
    raw_addr = d.get("Address", "")
    addr = " ".join(raw_addr.split())

    break_keys = [
        "DIVISION","CWC","OPP.","NEAR","BEHIND","OFFICE","BUILDING",
        "FLOOR","ROAD","MARG","AREA","COLONY"
    ]

    parts = addr.replace(",", ", ").split(",")
    lines = []
    buf = ""

    for p in parts:
        seg = p.strip()
        if any(seg.upper().startswith(k) for k in break_keys):
            lines.append(seg)
            continue
        if len(seg) > 30:
            lines.append(seg)
            continue
        if not buf:
            buf = seg
        else:
            buf += ", " + seg

    if buf:
        lines.append(buf)

    final_address = lines

    # -----------------------------------
    # MAPPING
    # -----------------------------------
    mapping = {
        "{Company}": d.get("Company", ""),
        "{Brand}": d.get("Brand", ""),
        "{Make}": d.get("Brand", ""),
        "{Category}": d.get("Category", ""),
        "{ProductName}": d.get("Product Name", ""),
        "{Model}": "",
        "{SerialNumber}": "",
        "{Quantity}": d.get("Quantity", ""),
        "{Warranty}": d.get("Warranty", ""),
        "{WarrantyOnCompressor}": d.get("Warranty on Compressor", ""),
        "{warranty on compressor}": d.get("Warranty on Compressor", ""),
        "{CustomerName}": d.get("Customer Name", ""),
        "{Organisation}": "",  # will rebuild manually
        "{Address}": "",       # will rebuild manually
        "{WarrantyBlock}": warranty_block,
        "{GEMContractNo}": d.get("GEM Contract No", ""),
        "{Date}": today,
    }

    # -----------------------------------
    # PROCESS DOCX
    # -----------------------------------
    doc = Document(template_file)

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

    # -----------------------------------
    # REBUILD CUSTOMER BLOCK (v5)
    # -----------------------------------
    # 1. Locate placeholder line OR block start
    insert_index = None
    for i, p in enumerate(doc.paragraphs):
        if "Customer" in p.text:
            insert_index = i
            break

    # Clear old customer block paragraphs completely
    for _ in range(5):
        if insert_index < len(doc.paragraphs):
            para = doc.paragraphs[insert_index]
            p_element = para._p
            p_element.getparent().remove(p_element)
        else:
            break

    # Insert new clean block
    cust_name = d.get("Customer Name", "")
    organisation = d.get("Organisation", "")
    date_text = today

    # --- Paragraph 1: Customer + Date (long single line)
    p1 = doc.paragraphs.insert(insert_index, f"Customer: {cust_name}\tDate: {date_text}")
    p1.style = doc.styles["Normal"]
    p1.paragraph_format.tab_stops.add_tab_stop(Inches(6.0))
    p1.alignment = WD_ALIGN_PARAGRAPH.LEFT
    blue(p1)

    # --- Paragraph 2: Organisation
    p2 = doc.paragraphs.insert(insert_index + 1, organisation)
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    blue(p2)

    # --- Paragraphs 3+: Address lines
    base = insert_index + 2
    for j, line in enumerate(final_address):
        p = doc.paragraphs.insert(base + j, line)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        blue(p)

    # -----------------------------------
    # FIX LETTERHEAD & TITLE
    # -----------------------------------
    # Center top letterhead (first 5 non-empty paragraphs)
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5, len(non_empty))):
        p = non_empty[i]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs:
            r.font.color.rgb = BLUE
            r.font.name = "Calibri"
            r.font.size = Pt(22 if i == 0 else 12)
            r.font.bold = True if i == 0 else False

    # Center the heading ONLY
    for p in doc.paragraphs:
        if p.text.strip().upper() == "WARRANTY CERTIFICATE":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.underline = True
                r.font.size = Pt(16)
                r.font.color.rgb = BLUE

    # -----------------------------------
    # FIX WARRANTY BLOCK (left aligned)
    # -----------------------------------
    for p in doc.paragraphs:
        if warranty_block.split("\n")[0] in p.text:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = None
            p.paragraph_format.right_indent = None
            blue(p)
            break

    # -----------------------------------
    # DRAW LINES
    # -----------------------------------
    for i, p in enumerate(doc.paragraphs):
        if "@" in p.text:
            newp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(newp)
            break

    for i, p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            newp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(newp)
            break

    # -----------------------------------
    # OUTPUT
    # -----------------------------------
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    fname = d.get("Customer Name", "Customer").replace(" ", "_")
    g = d.get("GEM Contract No", "GEM").replace(" ", "_")
    fn = f"Warranty_{fname}_{g}.docx"

    st.success("✅ Certificate Generated Successfully!")
    st.download_button("⬇ Download Certificate", out, fn)
