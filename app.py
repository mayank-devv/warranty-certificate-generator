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
# Text Input (Paste the ChatGPT Extracted Block)
# -------------------------------
raw_text = st.text_area(
    "Paste Extracted Details (as provided by ChatGPT, Option A format)",
    height=350,
    placeholder="Paste the block like:\n\nCompany: Shrii Salez Corporation\nBrand: Godrej\nCategory: AC\nProduct Name: ...\n..."
)

# -------------------------------
# Upload DOCX Template
# -------------------------------
template_file = st.file_uploader("Upload Warranty Certificate DOCX Template", type=["docx"])

# -------------------------------
# Utility Functions
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

def render_labeled_paragraph(p, text):
    for i in range(len(p.runs) - 1, -1, -1):
        p._element.remove(p.runs[i]._element)

    if ":" in text:
        label, _, value = text.partition(":")
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
        r = p.add_run(text)
        r.font.name = "Calibri"
        r.font.size = Pt(12)
        r.font.color.rgb = BLUE

    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

def merge_and_replace(doc, mapping):
    def _process(container):
        for p in container.paragraphs:
            original = "".join(r.text for r in p.runs)
            replaced = original
            for k, v in mapping.items():
                replaced = replaced.replace(k, v)
            render_labeled_paragraph(p, replaced)

        for table in getattr(container, "tables", []):
            for row in table.rows:
                for cell in row.cells:
                    _process(cell)
    _process(doc)

# -------------------------------
# Extract Key-Value Pairs from Pasted Block
# -------------------------------
def parse_text_block(text):
    data = {}
    for line in text.split("\n"):
        if ":" in line:
            key, _, value = line.partition(":")
            data[key.strip()] = value.strip()
    return data

# -------------------------------
# MAIN BUTTON
# -------------------------------
if st.button("Generate Certificate"):
    if not template_file:
        st.error("Please upload a DOCX template.")
    elif not raw_text.strip():
        st.error("Please paste the extracted details.")
    else:
        details = parse_text_block(raw_text)

        # Mapping placeholders from template → values
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

        # Set narrow margins
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        # Replace placeholders
        merge_and_replace(doc, mapping)

        # Letterhead formatting + title formatting
        if doc.paragraphs:
            header = doc.paragraphs[0]
            for run in header.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(22)
                run.font.bold = True
                run.font.color.rgb = BLUE
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Warranty certificate title formatting
        for p in doc.paragraphs:
            if "WARRANTY CERTIFICATE" in p.text.upper():
                for run in p.runs:
                    run.font.size = Pt(16)
                    run.font.bold = True
                    run.font.underline = True
                    run.font.color.rgb = BLUE
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Output file
        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_buf.seek(0)

        fname = details.get("Customer Name", "Customer").replace(" ", "_")
        gem = details.get("GEM Contract No", "GEM").replace(" ", "_")
        file_name = f"Warranty_{fname}_{gem}.docx"

        st.success("✅ Certificate generated successfully!")
        st.download_button(
            "⬇️ Download Certificate",
            data=out_buf,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
