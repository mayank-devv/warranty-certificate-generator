import io
import streamlit as st
from datetime import datetime
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# -------------------------------
# Streamlit Page Setup
# -------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator — Simplified Mode")
st.caption("Upload your DOCX template and paste the raw data block. The app will auto-parse and generate the final certificate.")

# -------------------------------
# User Inputs
# -------------------------------
template_file = st.file_uploader("📄 Upload DOCX Template", type=["docx"])

raw_data = st.text_area(
    "📌 Paste Data Block (as text)",
    height=300,
    placeholder="Example:\nProduct Name: Godrej AC\nModel: DSS 12...\nQuantity: 1 Unit\nSerial: XXXXX\n..."
)

submitted = st.button("Generate Warranty Certificate")

# -------------------------------
# Formatting Utilities
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
            original = "".join(run.text for run in p.runs)
            new_text = original
            for k, v in mapping.items():
                new_text = new_text.replace(k, v)
            render_labeled_paragraph(p, new_text)

        for table in getattr(container, "tables", []):
            for row in table.rows:
                for cell in row.cells:
                    _process(cell)

    _process(doc)

# -------------------------------
# Parsing Raw Block
# -------------------------------
def parse_raw_block(raw):
    mapping = {}
    for line in raw.split("\n"):
        if ":" in line:
            key, _, value = line.partition(":")
            k = "{" + key.strip() + "}"
            v = value.strip()
            mapping[k] = v
    return mapping

# -------------------------------
# Main Generation
# -------------------------------
if submitted:
    if not template_file:
        st.error("❌ Please upload a DOCX template.")
    elif not raw_data.strip():
        st.error("❌ Please paste the data block.")
    else:
        # Parse raw data block into mapping
        mapping = parse_raw_block(raw_data)

        # Add date placeholder automatically
        mapping["{Date}"] = datetime.now().strftime("%d-%m-%Y")

        # Load template
        doc = Document(template_file)

        # Narrow margins
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        # Replace placeholders
        merge_and_replace(doc, mapping)

        # Save Output
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)

        st.success("✅ Warranty Certificate Generated Successfully!")

        st.download_button(
            "⬇️ Download Certificate (DOCX)",
            data=out,
            file_name="Warranty_Certificate.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
