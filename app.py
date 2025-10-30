import io
import tempfile
import os
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm

# -------------------------------
# Streamlit Setup
# -------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Upload DOCX template → Fill details → Get formatted Warranty Certificate as PDF")

# -------------------------------
# Options
# -------------------------------
companies = ["Mathuralal Balkishan India", "Shrii Salez Corporation"]
categories = ["AC", "Refrigerator", "Appliances", "Display Panel", "Other"]
brands = [
    "Godrej", "Whirlpool", "LG", "Samsung", "Llyod", "Blue Star", "Uniline", "Numeric",
    "Epson", "Viewsonic", "Acer", "Exide", "Amaron", "Okaya", "Microtek", "Other"
]

# -------------------------------
# Form Inputs
# -------------------------------
with st.form("wc_form"):
    col1, col2 = st.columns(2)
    with col1:
        company = st.selectbox("Company", companies)
        category = st.selectbox("Category", categories)
        brand = st.selectbox("Brand (Make)", brands)
    with col2:
        product_name = st.text_input("Product Name")
        model = st.text_input("Model")
        quantity = st.text_input("Quantity", "1 Unit")
        serial_no = st.text_input("Serial Number")

    brand_custom = st.text_input("Enter Brand (if Other)") if brand == "Other" else ""
    gem_no = st.text_input("GEM Contract No")
    warranty = st.text_input("Warranty (e.g., 5 Years Overall)")
    warranty_compressor = st.text_input("Warranty on Compressor (e.g., 10 Years)")
    customer_name = st.text_input("Customer Name / Dept")
    organisation = st.text_input("Organisation")
    address = st.text_area("Address (use commas or new lines)")

    today_str = datetime.now().strftime("%d-%m-%Y")
    st.info(f"Certificate Date will be automatically set to: **{today_str}**")

    template_file = st.file_uploader("Upload DOCX Template", type=["docx"])
    submitted = st.form_submit_button("Generate Certificate")

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
# Main Process
# -------------------------------
if submitted:
    if not template_file:
        st.error("Please upload a DOCX template first.")
    else:
        clean_address = address.replace("\n", ", ").replace(",,", ",").strip().strip(",")
        final_brand = brand_custom.strip() if (brand == "Other" and brand_custom.strip()) else brand

        mapping = {
            "{Company}": company,
            "{Category}": category,
            "{Brand}": final_brand,
            "{Make}": final_brand,
            "{ProductName}": product_name,
            "{Model}": model,
            "{Quantity}": quantity,
            "{SerialNumber}": serial_no,
            "{GEMContractNo}": gem_no,
            "{Warranty}": warranty,
            "{CustomerName}": customer_name,
            "{Organisation}": organisation,
            "{Address}": clean_address,
            "{Date}": today_str,
            "{WarrantyOnCompressor}": warranty_compressor,
            "{Warranty on Compressor}": warranty_compressor,
            "{warranty on compressor}": warranty_compressor,
        }

        doc = Document(template_file)

        # --- Narrow Layout ---
        for section in doc.sections:
            section.top_margin = Inches(0.5)
            section.bottom_margin = Inches(0.5)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        merge_and_replace(doc, mapping)

        # Formatting (titles, lines, alignment)
        if doc.paragraphs:
            header = doc.paragraphs[0]
            for run in header.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(22)
                run.font.bold = True
                run.font.color.rgb = BLUE
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for i, p in enumerate(doc.paragraphs):
            if "Email" in p.text or "@" in p.text:
                new_p = doc.paragraphs[i + 1].insert_paragraph_before("")
                add_horizontal_line(new_p)
                break

        for p in doc.paragraphs:
            if "WARRANTY CERTIFICATE" in p.text.upper():
                for run in p.runs:
                    run.font.size = Pt(16)
                    run.font.bold = True
                    run.font.underline = True
                    run.font.color.rgb = BLUE
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # -------------------------------
        # Save DOCX and generate PDF via ReportLab
        # -------------------------------
        fname_customer = (customer_name or "Customer").replace(" ", "_").strip("_")
        fname_gem = (gem_no or "GEM").replace(" ", "_").strip("_")
        out_name_docx = f"Warranty_{fname_customer}_{fname_gem}.docx"
        out_name_pdf = f"Warranty_{fname_customer}_{fname_gem}.pdf"

        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = os.path.join(tmpdir, out_name_docx)
            pdf_path = os.path.join(tmpdir, out_name_pdf)
            doc.save(docx_path)

            pdf = canvas.Canvas(pdf_path, pagesize=A4)
            width, height = A4
            y = height - 2 * cm

            pdf.setFont("Helvetica-Bold", 14)
            pdf.drawString(2 * cm, y, company)
            y -= 1 * cm

            pdf.setFont("Helvetica", 11)
            pdf.drawString(2 * cm, y, f"Customer: {customer_name}")
            y -= 0.7 * cm
            pdf.drawString(2 * cm, y, f"Organisation: {organisation}")
            y -= 0.7 * cm
            pdf.drawString(2 * cm, y, f"Address: {clean_address}")
            y -= 0.7 * cm
            pdf.drawString(2 * cm, y, f"GEM Contract No: {gem_no}")
            y -= 0.7 * cm
            pdf.drawString(2 * cm, y, f"Product: {product_name} | Model: {model} | Qty: {quantity}")
            y -= 0.7 * cm
            pdf.drawString(2 * cm, y, f"Warranty: {warranty}")
            y -= 0.7 * cm
            pdf.drawString(2 * cm, y, f"On Compressor: {warranty_compressor}")
            y -= 1 * cm
            pdf.drawString(2 * cm, y, f"Date: {today_str}")
            y -= 2 * cm

            pdf.setFont("Helvetica", 10)
            pdf.drawString(2 * cm, y, "This is to certify that the supplied goods are new and of first quality as per GEM contract.")
            y -= 1 * cm
            pdf.line(2 * cm, y, width - 2 * cm, y)
            y -= 1 * cm
            pdf.setFont("Helvetica-Bold", 11)
            pdf.drawString(2 * cm, y, company)
            y -= 0.5 * cm
            pdf.setFont("Helvetica", 10)
            pdf.drawString(2 * cm, y, "Authorized Signatory")

            pdf.save()

            with open(pdf_path, "rb") as f:
                pdf_data = f.read()

            st.download_button(
                "⬇️ Download Certificate (PDF)",
                data=pdf_data,
                file_name=out_name_pdf,
                mime="application/pdf"
            )

        st.success("✅ Certificate generated successfully as PDF!")
