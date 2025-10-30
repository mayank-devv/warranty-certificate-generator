if not template_file:
    st.warning("No DOCX uploaded — using built-in warranty format.")
    from docx import Document
    doc = Document()
    p = doc.add_paragraph()
    run = p.add_run(company)
    run.font.name = "Calibri"
    run.font.bold = True
    run.font.size = Pt(22)
    run.font.color.rgb = RGBColor(0,112,192)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"20D, Kibe Compound, RNT Marg Corner, Indore", style=None)
    doc.add_paragraph("Ph. (O) 07314285302 (M) +919713323666 M) +919713356161", style=None)
    doc.add_paragraph("Email: shrishagrawaal@gmail.com, siddarthagrawaal@gmail.com", style=None)
    add_horizontal_line(doc.add_paragraph())

    p = doc.add_paragraph("WARRANTY CERTIFICATE")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in p.runs:
        r.font.bold = True
        r.font.size = Pt(16)
        r.font.underline = True
        r.font.color.rgb = RGBColor(0,112,192)

    doc.add_paragraph(f"Customer: {customer_name}")
    doc.add_paragraph(f"Organisation: {organisation}")
    doc.add_paragraph(f"Address: {address}")
    doc.add_paragraph(f"Date: {today_str}")
    doc.add_paragraph(f"GEM Contract No: {gem_no}")
    add_horizontal_line(doc.add_paragraph())

    doc.add_paragraph(
        f"This is to certify that the supplied {brand} {product_name} ({model}) "
        f"of quantity {quantity} and serial number {serial_no} "
        f"is covered under OEM warranty for {warranty}.\n"
        f"On Compressor: {warranty_compressor}."
    )
else:
    doc = Document(template_file)
