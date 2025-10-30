import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")

st.caption("Upload your DOCX template, fill details, and generate a perfectly formatted certificate.")

# Dropdown options
companies = ["Mathuralal Balkishan India", "Shrii Salez Corporation"]
categories = ["AC", "Refrigerator", "Appliances", "Display Panel", "Other"]
brands = [
    "Godrej", "Whirlpool", "LG", "Samsung", "Llyod", "Blue Star", "Uniline", "Numeric",
    "Epson", "Viewsonic", "Acer", "Exide", "Amaron", "Okaya", "Microtek", "Other"
]

# Form
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
    warranty = st.text_input("Warranty (e.g., 5 Years)")
    customer_name = st.text_input("Customer Name / Dept")
    organisation = st.text_input("Organisation")
    address = st.text_area("Address")
    ministry = st.text_input("Ministry")

    today_str = datetime.now().strftime("%d-%m-%Y")
    st.info(f"Certificate Date will be automatically set to: **{today_str}**")

    template_file = st.file_uploader("Upload DOCX Template", type=["docx"])
    submitted = st.form_submit_button("Generate Certificate")

# Helper functions
def style_run(run, size=12, bold=False):
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    # Dark blue from your sample LG 214 file
    run.font.color.rgb = RGBColor(31, 73, 125)

def replace_in_paragraph(paragraph, mapping):
    for run in paragraph.runs:
        for key, val in mapping.items():
            if key in run.text:
                run.text = run.text.replace(key, val)
        style_run(run)

def replace_in_table(table, mapping):
    for row in table.rows:
        for cell in row.cells:
            for p in cell.paragraphs:
                replace_in_paragraph(p, mapping)

def apply_mapping(doc, mapping):
    for p in doc.paragraphs:
        replace_in_paragraph(p, mapping)
    for tbl in doc.tables:
        replace_in_table(tbl, mapping)

if submitted:
    if not template_file:
        st.error("Please upload a DOCX template first.")
    else:
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
            "{Address}": address,
            "{Ministry}": ministry,
            "{Date}": today_str,
        }

        doc = Document(template_file)

        # If top doesn't already have company name, insert it
        if not doc.paragraphs or "{Company}" not in doc.paragraphs[0].text:
            doc.paragraphs.insert(0, doc.add_paragraph("{Company}"))

        # Replace placeholders
        apply_mapping(doc, mapping)

        # Apply formatting globally
        for i, p in enumerate(doc.paragraphs):
            for run in p.runs:
                style_run(run, size=12)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        # Style top header (first paragraph)
        if doc.paragraphs:
            header = doc.paragraphs[0]
            for run in header.runs:
                style_run(run, size=22, bold=True)
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Style "WARRANTY CERTIFICATE"
        for p in doc.paragraphs:
            if "WARRANTY CERTIFICATE" in p.text.upper():
                for run in p.runs:
                    style_run(run, size=16, bold=True)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Save and download
        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_buf.seek(0)

        fname_customer = (customer_name or "Customer").replace(" ", "_").strip("_")
        fname_gem = (gem_no or "GEM").replace(" ", "_").strip("_")
        out_name = f"Warranty_{fname_customer}_{fname_gem}.docx"

        st.success("Warranty Certificate generated with correct Calibri dark blue style ✅")
        st.download_button("⬇️ Download DOCX",
                           data=out_buf,
                           file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
