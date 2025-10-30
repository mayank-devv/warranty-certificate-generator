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

st.caption("Upload your DOCX template, fill the form, and generate a professionally formatted certificate.")

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
        product_name = st.text_input("Pro
