import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

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

def apply_blue_to_runs(p):
    for r in p.runs:
        r.font.color.rgb = BLUE
        r.font.name = "Calibri"
        r.font.size = Pt(12)

def safe_add_paragraph(doc, index, text, align="left"):
    p = doc.paragraphs.insert(index, "")  # always empty first
    r = p.add_run(text)                   # safe run created

    # apply blue font
    r.font.color.rgb = BLUE
    r.font.name = "Calibri"
    r.font.size = Pt(12)

    # xml alignment
    if align == "left":
        align_left(p)
    elif align == "center":
        align_center(p)

    return p

# ---------------------------------------------------------
# STREAMLIT PAGE SETUP
# ---------------------------------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Upload your DOCX template and paste extracted details.")

raw_text = st.text_area("Paste Extracted Details:", height=280)
template_file = st.file_uploader("Upload Template (.docx)", type=["docx"])

# ---------------------------------------------------------
# PROCESS
# ---------------------------------------------------------
if st.button("Generate"):

    if not template_file:
        st.error("Upload template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste details.")
        st.stop()

    today = datetime.now().strftime("%d-%m-%Y")

    # parse key:value
    d = {}
    for line in raw_text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            d[k.strip()] = v.strip()

    # warranty block
    warranty_block = (
        "Warranty can be checked anytime by contacting OEM customer care.\n"
        "Warranty is taken care of by OEM as per their terms & conditions. "
        "Original Warranty certificate is to be taken by above if needed."
    )

    # address cleaning
    addr_text = d.get("Address", "")
    addr_text = " ".join(addr_text.split())
    parts = [p.strip() for p in addr_text.replace(",", ", ").split(",")]

    address_lines = []
    buf = ""
    for seg in parts:
        if len(seg) > 30:
            address_lines.append(seg)
        else:
            if not buf:
                buf = seg
            else:
                buf += ", " + seg
    if buf:
        address_lines.append(buf)

    # open document
    doc = Document(template_file)

    # narrow margins
    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # Replace placeholders
    mapping = {
        "{Company}": d.get("Company", ""),
        "{CustomerName}": d.get("Customer Name",""),
        "{WarrantyBlock}": warranty_block,
        "{GEMContractNo}": d.get("GEM Contract No", ""),
        "{Date}": today,
    }
    for p in doc.paragraphs:
        txt = p.text
        for k,v in mapping.items():
            txt = txt.replace(k,v)
        p.text = txt

    # FIND CUSTOMER BLOCK LOCATION
    insert_index = None
    for i,p in enumerate(doc.paragraphs):
        if "Customer" in p.text:
            insert_index = i
            break

    # remove 4 old paragraphs
    for _ in range(4):
        if insert_index < len(doc.paragraphs):
            para = doc.paragraphs[insert_index]
            parent = para._p.getparent()
            parent.remove(para._p)

    # BUILD NEW CUSTOMER BLOCK (bulletproof)

    line1 = f"Customer: {d.get('Customer Name','')}{' '*40}Date: {today}"
    safe_add_paragraph(doc, insert_index, line1, align="left")

    org = d.get("Organisation", "")
    safe_add_paragraph(doc, insert_index+1, org, align="left")

    base = insert_index+2
    for j, line in enumerate(address_lines):
        safe_add_paragraph(doc, base+j, line, align="left")

    # FIX HEADER (top 5 non-empty center)
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5,len(non_empty))):
        p = non_empty[i]
        align_center(p)
        for r in p.runs:
            r.font.name="Calibri"
            r.font.color.rgb=BLUE
            r.font.size=Pt(22 if i==0 else 12)
            r.font.bold = (i==0)

    # WARRANTY CERTIFICATE TITLE
    for p in doc.paragraphs:
        if p.text.strip().upper() == "WARRANTY CERTIFICATE":
            align_center(p)
            for r in p.runs:
                r.font.bold=True
                r.font.underline=True
                r.font.size=Pt(16)
                r.font.color.rgb=BLUE

    # WARRANTY BLOCK ALIGN
    for p in doc.paragraphs:
        if warranty_block.split("\n")[0] in p.text:
            align_left(p)
            apply_blue_to_runs(p)
            break

    # BLUE LINES
    for i,p in enumerate(doc.paragraphs):
        if "@" in p.text:
            newp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(newp)
            break

    for i,p in enumerate(doc.paragraphs):
        if "GEM Contract No" in p.text:
            newp = doc.paragraphs[i+1].insert_paragraph_before("")
            add_line(newp)
            break

    # save
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    st.download_button("DOWNLOAD DOCX", out, f"Warranty_{today}.docx")
