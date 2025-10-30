import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")

st.caption("Upload your DOCX template, fill details, and generate a perfectly formatted certificate.")

# Dropdowns
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

# --- Bulletproof replacement ---
def merge_and_replace(doc, mapping):
    # paragraphs
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs)
        for key, val in mapping.items():
            if key in full_text:
                full_text = full_text.replace(key, val)
        # remove old runs
        for i in range(len(p.runs)-1, -1, -1):
            p._element.remove(p.runs[i]._element)
        # rebuild one run with consistent styling
        new_run = p.add_run(full_text)
        new_run.font.name = "Calibri"
        new_run.font.size = Pt(12)
        new_run.font.color.rgb = RGBColor(31, 73, 125)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                merge_and_replace(cell, mapping)

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

        # Perform replacements
        merge_and_replace(doc, mapping)

        # Style company heading (first para)
        if doc.paragraphs:
            header = doc.paragraphs[0]
            for run in header.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(22)
                run.font.bold = True
                run.font.color.rgb = RGBColor(31, 73, 125)
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Style "WARRANTY CERTIFICATE"
        for p in doc.paragraphs:
            if "WARRANTY CERTIFICATE" in p.text.upper():
                for run in p.runs:
                    run.font.size = Pt(16)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(31, 73, 125)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Save and download
        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_buf.seek(0)

        fname_customer = (customer_name or "Customer").replace(" ", "_").strip("_")
        fname_gem = (gem_no or "GEM").replace(" ", "_").strip("_")
        out_name = f"Warranty_{fname_customer}_{fname_gem}.docx"

        st.success("Warranty Certificate generated with full replacements and Calibri Dark Blue formatting ✅")
        st.download_button("⬇️ Download DOCX",
                           data=out_buf,
                           file_name=out_name,
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
