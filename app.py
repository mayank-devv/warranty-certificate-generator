import io
from datetime import datetime
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# -------------------------------------------
# STREAMLIT PAGE SETUP
# -------------------------------------------
st.set_page_config(page_title="Warranty Certificate Generator", page_icon="🧾", layout="centered")
st.title("🧾 Warranty Certificate Generator")
st.caption("Auto-create formatted warranty certificates — with or without uploading a DOCX template.")

# -------------------------------------------
# CONSTANTS
# -------------------------------------------
BLUE = RGBColor(0, 112, 192)

# -------------------------------------------
# OPTIONS
# -------------------------------------------
companies = ["Mathuralal Balkishan India", "Shrii Salez Corporation"]
categories = ["AC", "Refrigerator", "Appliances", "Display Panel", "Other"]
brands = [
    "Godrej", "Whirlpool", "LG", "Samsung", "Llyod", "Blue Star", "Uniline", "Numeric",
    "Epson", "Viewsonic", "Acer", "Exide", "Amaron", "Okaya", "Microtek", "Other"
]

# -------------------------------------------
# FORM
# -------------------------------------------
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
    warranty = st.text_input("On-Site OEM Warranty (e.g., 5 Years)")
    warranty_compressor = st.text_input("Warranty on Compressor (e.g., 10 Years)")
    customer_name = st.text_input("Customer Name / Department")
    organisation = st.text_input("Organisation")
    address = st.text_area("Address (use commas or new lines)")

    today_str = datetime.now().strftime("%d-%m-%Y")
    st.info(f"Certificate Date will be automatically set to: **{today_str}**")

    template_file = st.file_uploader("Upload DOCX Template (optional)", type=["docx"])
    submitted = st.form_submit_button("Generate Certificate")

# -------------------------------------------
# UTILITIES
# -------------------------------------------
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
        r1.font.bold = True; r1.font.name = "Calibri"; r1.font.size = Pt(12); r1.font.color.rgb = BLUE

        r2 = p.add_run(" " + value.strip())
        r2.font.bold = False; r2.font.name = "Calibri"; r2.font.size = Pt(12); r2.font.color.rgb = BLUE
    else:
        r = p.add_run(text)
        r.font.name = "Calibri"; r.font.size = Pt(12); r.font.color.rgb = BLUE

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

# -------------------------------------------
# MAIN GENERATION
# -------------------------------------------
if submitted:
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

    # -------------------------------------------
    # USE TEMPLATE OR CREATE BUILT-IN
    # -------------------------------------------
    if template_file:
        doc = Document(template_file)
    else:
        st.warning("No template uploaded — using built-in layout.")
        doc = Document()

        # Company header
        p = doc.add_paragraph(company)
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(22)
            r.font.bold = True
            r.font.color.rgb = BLUE
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph("20D, Kibe Compound, RNT Marg Corner, Indore")
        doc.add_paragraph("Ph. (O) 07314285302 (M) +919713323666 M) +919713356161")
        doc.add_paragraph("Email: shrishagrawaal@gmail.com, siddarthagrawaal@gmail.com")
        add_horizontal_line(doc.add_paragraph())

        p = doc.add_paragraph("WARRANTY CERTIFICATE")
        for r in p.runs:
            r.font.name = "Calibri"
            r.font.size = Pt(16)
            r.font.bold = True
            r.font.underline = True
            r.font.color.rgb = BLUE
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph(f"Customer: {customer_name}")
        doc.add_paragraph(f"Organisation: {organisation}")
        doc.add_paragraph(f"Address: {clean_address}")
        doc.add_paragraph(f"Date: {today_str}")
        doc.add_paragraph(f"GEM Contract No: {gem_no}")
        add_horizontal_line(doc.add_paragraph())
        doc.add_paragraph(
            f"This is to certify that the supplied {brand} {product_name} ({model}) "
            f"of quantity {quantity} and serial number {serial_no} is covered under "
            f"OEM warranty for {warranty}. On Compressor: {warranty_compressor}."
        )

    # -------------------------------------------
    # MERGE PLACEHOLDERS
    # -------------------------------------------
    merge_and_replace(doc, mapping)

    # -------------------------------------------
    # STYLING AND LINE PLACEMENT
    # -------------------------------------------
    for i, p in enumerate(doc.paragraphs):
        if "Email" in p.text or "@" in p.text:
            new_p = doc.paragraphs[i+1].insert_paragraph_before("")
            add_horizontal_line(new_p)
            break

    for p in doc.paragraphs:
        if "WARRANTY CERTIFICATE" in p.text.upper():
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(16)
                r.font.underline = True
                r.font.color.rgb = BLUE
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    cust_idx = None
    date_idx = None
    gem_idx = None
    for i, p in enumerate(doc.paragraphs):
        if cust_idx is None and "Customer" in p.text:
            cust_idx = i
        if date_idx is None and "Date" in p.text:
            date_idx = i
        if gem_idx is None and "GEM Contract No" in p.text:
            gem_idx = i

    if cust_idx is not None:
        doc.paragraphs[cust_idx].alignment = WD_ALIGN_PARAGRAPH.LEFT
    if date_idx is not None:
        doc.paragraphs[date_idx].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if cust_idx is not None and gem_idx is not None:
        for k in range(cust_idx + 1, gem_idx):
            if k != date_idx:
                doc.paragraphs[k].alignment = WD_ALIGN_PARAGRAPH.LEFT

    if gem_idx is not None:
        new_p = doc.paragraphs[gem_idx+1].insert_paragraph_before("")
        add_horizontal_line(new_p)

    for i, p in enumerate(doc.paragraphs):
        if "Supplied Product Details" in p.text:
            new_p = doc.paragraphs[i].insert_paragraph_before("")
            add_horizontal_line(new_p)
            break

    # -------------------------------------------
    # SAVE OUTPUT
    # -------------------------------------------
    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)

    fname_customer = (customer_name or "Customer").replace(" ", "_").strip("_")
    fname_gem = (gem_no or "GEM").replace(" ", "_").strip("_")
    out_name = f"Warranty_{fname_customer}_{fname_gem}.docx"

    st.success("✅ Warranty Certificate Ready — including Auto-Date, Compressor Warranty & Lines.")
    st.download_button(
        "⬇️ Download Certificate (DOCX)",
        data=out_buf,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
