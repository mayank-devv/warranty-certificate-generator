import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")

st.caption("Upload your DOCX template, fill details, and generate a styled warranty certificate automatically.")

# --- Dropdown Options ---
companies = ["Mathuralal Balkishan India", "Shrii Salez Corporation"]
categories = ["AC", "Refrigerator", "Appliances", "Display Panel", "Other"]
brands = [
    "Godrej", "Whirlpool", "LG", "Samsung", "Llyod", "Blue Star", "Uniline", "Numeric",
    "Epson", "Viewsonic", "Acer", "Exide", "Amaron", "Okaya", "Microtek", "Other"
]

# --- Form ---
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

# --- Utility: Add horizontal blue line ---
def add_horizontal_line(paragraph):
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = parse_xml(
        r'<w:pBdr %s><w:bottom w:val="single" w:sz="6" w:space="1" w:color="0070C0"/></w:pBdr>'
        % nsdecls("w")
    )
    pPr.append(pBdr)

# --- Merge placeholders with formatting ---
def merge_and_replace(doc, mapping):
    for p in doc.paragraphs:
        full_text = "".join(run.text for run in p.runs)
        for key, val in mapping.items():
            if key in full_text:
                if ":" in full_text:
                    label, _, _ = full_text.partition(":")
                    full_text = label + ": " + val
                else:
                    full_text = full_text.replace(key, val)

        for i in range(len(p.runs) - 1, -1, -1):
            p._element.remove(p.runs[i]._element)

        if ":" in full_text:
            parts = full_text.split(":")
            run_label = p.add_run(parts[0] + ":")
            run_label.font.bold = True
            run_label.font.name = "Calibri"
            run_label.font.size = Pt(12)
            run_label.font.color.rgb = RGBColor(0, 112, 192)

            if len(parts) > 1:
                run_value = p.add_run(parts[1])
                run_value.font.name = "Calibri"
                run_value.font.size = Pt(12)
                run_value.font.color.rgb = RGBColor(0, 112, 192)
        else:
            run_text = p.add_run(full_text)
            run_text.font.name = "Calibri"
            run_text.font.size = Pt(12)
            run_text.font.color.rgb = RGBColor(0, 112, 192)

        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                merge_and_replace(cell, mapping)

# --- Main Process ---
if submitted:
    if not template_file:
        st.error("Please upload a DOCX template first.")
    else:
        final_brand = brand_custom.strip() if (brand == "Other" and brand_custom.strip()) else brand
        clean_address = address.replace("\n", ", ").replace(",,", ",")

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
            "{WarrantyOnCompressor}": warranty_compressor,
            "{CustomerName}": customer_name,
            "{Organisation}": organisation,
            "{Address}": clean_address,
            "{Date}": today_str,
        }

        doc = Document(template_file)
        merge_and_replace(doc, mapping)

        # --- STYLING ---
        # Company heading
        if doc.paragraphs:
            header = doc.paragraphs[0]
            for run in header.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(22)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 112, 192)
            header.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Center the letterhead
        for i in range(1, 7):
            if i < len(doc.paragraphs):
                p = doc.paragraphs[i]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = "Calibri"
                    run.font.size = Pt(12)
                    run.font.color.rgb = RGBColor(0, 112, 192)

        # --- Controlled Blue Lines ---

        # 1️⃣ Line below letterhead
        for i, p in enumerate(doc.paragraphs):
            if "Email" in p.text or "@" in p.text:
                new_p = doc.paragraphs[i+1].insert_paragraph_before("")
                add_horizontal_line(new_p)
                break

        # 2️⃣ Line below Customer / GEM Contract No block
        customer_index = None
        gem_index = None
        for i, p in enumerate(doc.paragraphs):
            if "Customer" in p.text and customer_index is None:
                customer_index = i
            if "GEM Contract No" in p.text:
                gem_index = i
        if gem_index is not None:
            new_p = doc.paragraphs[gem_index+1].insert_paragraph_before("")
            add_horizontal_line(new_p)

        # 3️⃣ Line above Supplied Product Details
        for i, p in enumerate(doc.paragraphs):
            if "Supplied Product Details" in p.text:
                new_p = doc.paragraphs[i].insert_paragraph_before("")
                add_horizontal_line(new_p)
                break

        # Warranty title formatting
        for p in doc.paragraphs:
            if "WARRANTY CERTIFICATE" in p.text.upper():
                for run in p.runs:
                    run.font.size = Pt(16)
                    run.font.bold = True
                    run.font.underline = True
                    run.font.color.rgb = RGBColor(0, 112, 192)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # --- Save & Download ---
        out_buf = io.BytesIO()
        doc.save(out_buf)
        out_buf.seek(0)

        fname_customer = (customer_name or "Customer").replace(" ", "_").strip("_")
        fname_gem = (gem_no or "GEM").replace(" ", "_").strip("_")
        out_name = f"Warranty_{fname_customer}_{fname_gem}.docx"

        st.success("✅ Warranty Certificate formatted perfectly.")
        st.download_button(
            "⬇️ Download Certificate (DOCX)",
            data=out_buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
