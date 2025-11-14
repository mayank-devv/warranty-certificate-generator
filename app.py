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

def align_xml(p, val):
    p._p.get_or_add_pPr().append(
        parse_xml(f'<w:jc w:val="{val}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    )

def colorize(cell):
    """Safe color + font update for ALL runs in cell"""
    for p in cell.paragraphs:
        for r in p.runs:
            r.font.color.rgb = BLUE
            r.font.name = "Calibri"
            r.font.size = Pt(12)

def colorize_para(p):
    for r in p.runs:
        r.font.color.rgb = BLUE
        r.font.name = "Calibri"
        r.font.size = Pt(12)

def add_line(paragraph):
    pPr = paragraph._p.get_or_add_pPr()
    xml = (r'<w:pBdr %s>'
           r'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/>'
           r'</w:pBdr>') % nsdecls("w")
    pPr.append(parse_xml(xml))

# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------
st.set_page_config(page_title="Warranty Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")

raw_text = st.text_area("Paste Extracted Details", height=280)
template_file = st.file_uploader("Upload Template (.docx)", type=["docx"])

# ---------------------------------------------------------
# MAIN
# ---------------------------------------------------------
if st.button("Generate"):

    if not template_file:
        st.error("Upload template.")
        st.stop()

    if not raw_text.strip():
        st.error("Paste extracted data.")
        st.stop()

    today = datetime.now().strftime("%d-%m-%Y")

    # Parse block into dict
    data = {}
    for line in raw_text.split("\n"):
        if ":" in line:
            k, _, v = line.partition(":")
            data[k.strip()] = v.strip()

    warranty_block = (
        "Warranty can be checked anytime by contacting OEM customer care.\n"
        "Warranty is taken care of by OEM as per their terms & conditions. "
        "Original Warranty certificate is to be taken by above if needed."
    )

    # Address split
    addr = " ".join(data.get("Address","").split())
    parts = addr.replace(",", ", ").split(",")

    address_lines = []
    buffer = ""
    for seg in parts:
        seg = seg.strip()
        if len(seg) > 30:
            address_lines.append(seg)
        else:
            if not buffer:
                buffer = seg
            else:
                buffer += ", " + seg
    if buffer:
        address_lines.append(buffer)

    # Load template
    doc = Document(template_file)

    for s in doc.sections:
        s.top_margin = Inches(0.5)
        s.bottom_margin = Inches(0.5)
        s.left_margin = Inches(0.5)
        s.right_margin = Inches(0.5)

    # Replace placeholders first
    mapping = {
        "{Company}": data.get("Company",""),
        "{CustomerName}": data.get("Customer Name",""),
        "{WarrantyBlock}": warranty_block,
        "{GEMContractNo}": data.get("GEM Contract No",""),
        "{Date}": today
    }

    for p in doc.paragraphs:
        txt = p.text
        for k,v in mapping.items():
            txt = txt.replace(k,v)
        p.text = txt

    # ---------------------------------------------------------
    # FIND OLD CUSTOMER BLOCK & DELETE IT
    # ---------------------------------------------------------
    c_index = None
    for i,p in enumerate(doc.paragraphs):
        if "Customer" in p.text:
            c_index = i
            break

    if c_index is None:
        c_index = 6

    # delete 4 old lines
    for _ in range(4):
        if c_index < len(doc.paragraphs):
            el = doc.paragraphs[c_index]._p
            parent = el.getparent()
            parent.remove(el)

    # ---------------------------------------------------------
    # INSERT CUSTOMER BLOCK USING TABLE (SAFE)
    # ---------------------------------------------------------
    # Insert table BEFORE c_index
    table = doc.add_table(rows=1, cols=2)
    table.alignment = 0  # left alignment

    # Move table to correct location
    tbl_elem = table._tbl
    doc.paragraphs[c_index]._p.addprevious(tbl_elem)

    # Fill table cells
    left_cell  = table.rows[0].cells[0]
    right_cell = table.rows[0].cells[1]

    left_cell.text  = f"Customer: {data.get('Customer Name','')}"
    right_cell.text = f"Date: {today}"

    colorize(left_cell)
    colorize(right_cell)

    # ---------------------------------------------------------
    # ORGANISATION BELOW TABLE
    # ---------------------------------------------------------
    org_para = doc.paragraphs.insert(c_index+1, "")
    org_para.add_run(data.get("Organisation",""))
    colorize_para(org_para)
    align_xml(org_para, "left")

    # ---------------------------------------------------------
    # ADDRESS LINES
    # ---------------------------------------------------------
    pos = c_index+2
    for line in address_lines:
        p = doc.paragraphs.insert(pos, "")
        p.add_run(line)
        colorize_para(p)
        align_xml(p, "left")
        pos += 1

    # ---------------------------------------------------------
    # FIX HEADER (CENTER TOP 5)
    # ---------------------------------------------------------
    non_empty = [p for p in doc.paragraphs if p.text.strip()]
    for i in range(min(5,len(non_empty))):
        p = non_empty[i]
        align_xml(p, "center")
        for r in p.runs:
            r.font.name="Calibri"
            r.font.color.rgb=BLUE
            r.font.size=Pt(22 if i==0 else 12)
            r.font.bold = (i==0)

    # ---------------------------------------------------------
    # WARRANTY CERTIFICATE HEADING
    # ---------------------------------------------------------
    for p in doc.paragraphs:
        if p.text.strip().upper() == "WARRANTY CERTIFICATE":
            align_xml(p, "center")
            for r in p.runs:
                r.font.bold=True
                r.font.underline=True
                r.font.size=Pt(16)
                r.font.color.rgb=BLUE

    # ---------------------------------------------------------
    # WARRANTY BLOCK (LEFT)
    # ---------------------------------------------------------
    for p in doc.paragraphs:
        if warranty_block.split("\n")[0] in p.text:
            colorize_para(p)
            align_xml(p, "left")
            break

    # ---------------------------------------------------------
    # BLUE LINES
    # ---------------------------------------------------------
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

    # ---------------------------------------------------------
    # SAVE
    # ---------------------------------------------------------
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)

    st.success("Generated Successfully!")
    st.download_button("Download DOCX", out, f"Warranty_{today}.docx")
