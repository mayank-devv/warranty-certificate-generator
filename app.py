import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
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
# HELPERS
# ---------------------------------------------------------
BLUE = RGBColor(0, 112, 192)

def add_line(p):
    pPr = p._p.get_or_add_pPr()
    xml = r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>' % nsdecls("w")
    pPr.append(parse_xml(xml))

def align_left(p):
    p._p.get_or_add_pPr().append(
        parse_xml(r'<w:jc w:val="left" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    )

def align_center(p):
    p._p.get_or_add_pPr().append(
        parse_xml(r'<w:jc w:val="center" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    )

def align_right(p):
    p._p.get_or_add_pPr().append(
        parse_xml(r'<w:jc w:val="right" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    )

def parse_block(text):
    out = {}
    for line in text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            out[k.strip()] = v.strip()
    return out

def apply_blue(p):
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

    warranty_block = (
        "Warranty can be checked anytime by contacting OEM customer care.\n"
        "Warranty is taken care of by OEM as per their terms & conditions. "
        "Original Warranty certificate is to be taken by above if needed."
    )

    # ---------------------------------------------------------
    # CLEAN ADDRESS
    # ---------------------------------------------------------
    raw_addr = d.get("Address", "")
    addr = " ".join(raw_addr.split())
    parts = addr.replace(",", ", ").split(",")

    lines = []
    buf = ""

    for seg in parts:
        seg = seg.strip()
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

    # ---------------------------------------------------------
    # PLACEHOLDER MAPPING
    # ---------------------------------------------------------
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
        "{CustomerName}": d.get("Customer Name", ""),
        "{Organisation}": "",
        "{Address}": "",
        "{WarrantyBlock}": warranty_block,
        "{GEMContractNo}": d.get("GEM Contract No", ""),
        "{Date}": today,
    }

    # ---------------------------------------------------------
    # LOAD DOCX
    # ---------------------------------------------------------
    doc = Document(template_file)

    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # Replace placeholders
    for p in doc.paragraphs:
        txt = p.text
        for k, v in mapping.items():
            txt = txt.replace(k, v)
        p.text = txt

    # ---------------------------------------------------------
    # REBUILD CUSTOMER BLOCK
    # ---------------------------------------------------------
    insert_index = None
    for i, p in enumerate(doc.paragraphs):
        if "Customer" in p.text:
            insert_index = i
            break

    # remove 4–5 old paragraphs
    for _ in range(5):
        if insert_index < len(doc.paragraphs):
            para = doc.paragraphs[insert_index]
            parent = para._p.getparent()
            parent.remove(para._p)

    cust_name = d.get("Customer Name", "")
    organisation = d.get("Organisation", "")

    line1 = f"Customer: {cust_name}{' '*40}Date: {today}"

    # Paragraph 1 — LEFT
    p1 = doc.paragraphs.insert(insert_index, line1)
    apply_blue(p1)
    align_left(p1)

    # Paragraph 2 — Organisation
    p2 = doc.paragraphs.insert(insert_index + 1, organisation)
    apply_blue(p2)
    align_left(p2)

    # Address lines
    base = insert_index + 2
    for j, line in enumerate(final_address):
        p = doc.paragraphs.insert(base + j, line)
        apply_blue(p)
        align_left(p)

    # ---------------------------------------------------------
    # LETTERHEAD CENTER
    # ---------------------------------------------------------
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5, len(non_empty))):
        p = non_empty[i]
        align_center(p)
        for r in p.runs:
            r.font.color.rgb = BLUE
            r.font.name = "Calibri"
            r.font.size = Pt(22 if i == 0 else 12)
            r.font.bold = (i == 0)

    # ---------------------------------------------------------
    # WARRANTY CERTIFICATE HEADING
    # ---------------------------------------------------------
    for p in doc.paragraphs:
        if p.text.strip().upper() == "WARRANTY CERTIFICATE":
            align_center(p)
            for r in p.runs:
                r.font.bold = True
                r.font.underline = True
                r.font.size = Pt(16)
                r.font.color.rgb = BLUE

    # ---------------------------------------------------------
    # WARRANTY BLOCK LEFT
    # ---------------------------------------------------------
    for p in doc.paragraphs:
        if warranty_block.split("\n")[0] in p.text:
            align_left(p)
            apply_blue(p)
            break

    # ---------------------------------------------------------
    # BLUE LINES
    # ---------------------------------------------------------
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

    # ---------------------------------------------------------
    # OUTPUT
    # ---------------------------------------------------------
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    fname = d.get("Customer Name", "Customer").replace(" ", "_")
    g = d.get("GEM Contract No", "GEM").replace(" ", "_")
    fn = f"Warranty_{fname}_{g}.docx"

    st.success("✅ Certificate Generated Successfully!")
    st.download_button("⬇ Download Certificate", out, fn)
