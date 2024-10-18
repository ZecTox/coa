import os
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

# Function to generate the PDF based on COA structure
def generate_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=10, bottomMargin=10)

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title_style', fontSize=12, spaceAfter=1, alignment=1, fontName='Helvetica-Bold')
    normal_style = styles['BodyText']
    normal_style.alignment = 1  # Center alignment for paragraphs

    elements = []

    # Set paths for images
    logo_path = os.path.join(os.getcwd(), "images", "tru_herb_logo.png")
    footer_path = os.path.join(os.getcwd(), "images", "footer.png")

    # Check if images exist
    if not os.path.exists(logo_path) or not os.path.exists(footer_path):
        st.error("One or more images are missing.")
        return None

    # Add logo and title at the top
    elements.append(Image(logo_path, width=100, height=40))
    elements.append(Spacer(1, 3))
    elements.append(Paragraph("CERTIFICATE OF ANALYSIS", title_style))
    elements.append(Paragraph(data.get('product_name', ''), title_style))  # Add product name
    elements.append(Spacer(1, 3))

    # Product Information Table (skipping empty fields)
    product_info = [
        ["Product Name", data.get('product_name', '')],
        ["Chemical Name", data.get('chemical_name', '')],  # New row
        ["CAS No.", data.get('cas_no', '')],  # New row
        ["Product Code", data.get('product_code', '')],
        ["Batch No.", data.get('batch_no', '')],
        ["Date of Manufacturing", data.get('manufacturing_date', '')],
        ["Date of Reanalysis", data.get('reanalysis_date', '')],
        ["Quantity (in Kgs)", data.get('quantity', '')],
        ["Source", data.get('source', '')],
        ["Country of Origin", data.get('origin', '')],
        ["Plant Parts", data.get('plant_part', '')],
        ["Extraction Ratio", data.get('extraction_ratio', '')],
        ["Extraction Solvents", data.get('solvent', '')],
        ["Botanical Name", data.get('botanical_name', '')],
    ]

    # Only include non-empty rows
    product_info = [row for row in product_info if row[1]]

    if product_info:
        product_table = Table(product_info, colWidths=[150, 350])
        product_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ]))
        elements.append(product_table)
        elements.append(Spacer(1, 0))

    # Specifications Table
    spec_headers = ["Parameter", "Specification", "Result", "Method"]

    # Define sections
    sections = {
        "Physical": [
            ("Identification", data['identification_spec'], data['identification_result'], data['identification_method']),
            ("Description", data['description_spec'], data['description_result'], data['description_method']),
            ("Moisture/Loss of Drying", data['moisture_spec'], data['moisture_result'], data['moisture_method']),
            ("Particle Size", data['particle_size_spec'], data['particle_size_result'], data['particle_size_method']),
            ("Bulk Density", data['bulk_density_spec'], data['bulk_density_result'], data['bulk_density_method']),
            ("Tapped Density", data['tapped_density_spec'], data['tapped_density_result'], data['tapped_density_method']),
            ("Ash Contents", data['ash_contents_spec'], data['ash_contents_result'], data['ash_contents_method']),
            ("pH", data['ph_spec'], data['ph_result'], data['ph_method']),
            ("Fats", data['fats_spec'], data['fats_result'], data['fats_method']),
            ("Protein", data['protein_spec'], data['protein_result'], data['protein_method']),
            ("Solubility", data['solubility_spec'], data['solubility_result'], data['solubility_method']),
            ("Limit of Oxalic Acid", data['oxalic_acid_spec'], data['oxalic_acid_result'], data['oxalic_acid_method']),
            ("Limit of NaCl", data['nacl_spec'], data['nacl_result'], data['nacl_method']),
            ("Sulphates", data['sulphates_spec'], data['sulphates_result'], data['sulphates_method']),
            ("Chloride", data['chloride_spec'], data['chloride_result'], data['chloride_method']),
        ],
        "Others": [
            ("Heavy Metals", data['heavy_metals_spec'], data['heavy_metals_result'], data['heavy_metals_method']),
            ("Lead", data['lead_spec'], data['lead_result'], data['lead_method']),
            ("Cadmium", data['cadmium_spec'], data['cadmium_result'], data['cadmium_method']),
            ("Arsenic", data['arsenic_spec'], data['arsenic_result'], data['arsenic_method']),
            ("Mercury", data['mercury_spec'], data['mercury_result'], data['mercury_method']),
        ],
        "Chemicals": [
            ("Assays", data['assays_spec'], data['assays_result'], data['assays_method']),
            ("Extraction", data['extraction_spec'], data['extraction_result'], data['extraction_method']),
        ],
        "Pesticides": [
            ("Pesticide", data['pesticide_spec'], data['pesticide_result'], data['pesticide_method']),
        ],
        "Microbiological Profile": [
            ("Total Plate Count", data['total_plate_count_spec'], data['total_plate_count_result'], data['total_plate_count_method']),
            ("Yeasts & Mould Count", data['yeasts_mould_spec'], data['yeasts_mould_result'], data['yeasts_mould_method']),
            ("E.coli", data['e_coli_spec'], data['e_coli_result'], data['e_coli_method']),
            ("Salmonella", data['salmonella_spec'], data['salmonella_result'], data['salmonella_method']),
            ("Coliforms", data['coliforms_spec'], data['coliforms_result'], data['coliforms_method']),
        ]
    }

    spec_data = [spec_headers]

    # Add sections to spec_data
    for section, params in sections.items():
        if any(param[1:] for param in params):
            spec_data.append([Paragraph(f"<b>{section}</b>", styles['Normal']), "", "", ""])
            spec_data.extend(params)

    # Define a bold and centered style for the end text
    bold_center_style = ParagraphStyle(
        'bold_center',
        parent=styles['Normal'],
        fontName='Helvetica-Bold',
        alignment=1  # Center alignment
    )

    # Remarks Section - merging rows for Remarks and Final Remark
    remarks_text = "Since the product is derived from natural origin, there is likely to be minor color variation because of the geographical and seasonal variations of the raw material"
    end_text = "REMARKS: COMPLIES WITH IN HOUSE SPECIFICATIONS"

    spec_data.append([Paragraph(remarks_text, styles['Normal']), "", "", ""])
    spec_data.append([Paragraph(end_text, bold_center_style), "", "", ""])

    # Build the table
    spec_table = Table(spec_data, colWidths=[120, 140, 120, 120])
    spec_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('SPAN', (0, 1), (-1, 1)),  # Span the Physical section header
        ('SPAN', (0, len(sections["Physical"])+2), (-1, len(sections["Physical"])+2)),  # Span the Others section header
        ('SPAN', (0, len(sections["Physical"])+len(sections["Others"])+3), (-1, len(sections["Physical"])+len(sections["Others"])+3)),  # Span the Chemicals section header
        ('SPAN', (0, len(sections["Physical"])+len(sections["Others"])+len(sections["Chemicals"])+4), (-1, len(sections["Physical"])+len(sections["Others"])+len(sections["Chemicals"])+4)),  # Span the Pesticides section header
        ('SPAN', (0, len(sections["Physical"])+len(sections["Others"])+len(sections["Chemicals"])+len(sections["Pesticides"])+5), (-1, len(sections["Physical"])+len(sections["Others"])+len(sections["Chemicals"])+len(sections["Pesticides"])+5)),  # Span the Microbiological Profile section header
        ('SPAN', (0, len(spec_data)-2), (-1, len(spec_data)-2)),  # Span the remarks row
        ('SPAN', (0, len(spec_data)-1), (-1, len(spec_data)-1)),  # Span the final remark row
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(spec_table)
    elements.append(Spacer(1, 5))

    # Declaration
    elements.append(Paragraph("Declaration:", title_style))
    declaration_text = "The said material conforms to the above specifications."
    elements.append(Paragraph(declaration_text, normal_style))

    # Footer Image
    elements.append(Spacer(1, 1))
    elements.append(Image(footer_path, width=500, height=60))

    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title("Tru Herb COA PDF Generator")

# Form for user input
with st.form("coa_form"):
    st.header("Product Information")
    product_name = st.text_input("Product Name")
    botanical_name = st.text_input("Botanical Name")
    chemical_name = st.text_input("Chemical Name")
    cas_no = st.text_input("CAS No.")
    product_code = st.text_input("Product Code")
    batch_no = st.text_input("Batch No.")
    manufacturing_date = st.text_input("Date of Manufacturing")  # Changed to text input
    reanalysis_date = st.text_input("Date of Reanalysis")  # Changed to text input
    quantity = st.text_input("Quantity (in Kgs)")
    source = st.text_input("Source")
    origin = st.text_input("Country of Origin")
    plant_part = st.text_input("Plant Parts")
    extraction_ratio = st.text_input("Extraction Ratio")
    solvent = st.text_input("Extraction Solvents")

    st.header("Specifications")
    st.subheader("Physical")
    identification_spec = st.text_input("Specification for Identification")
    identification_result = st.text_input("Result for Identification")
    identification_method = st.text_input("Method for Identification")
    description_spec = st.text_input("Specification for Description")
    description_result = st.text_input("Result for Description")
    description_method = st.text_input("Method for Description")
    moisture_spec = st.text_input("Specification for Moisture/Loss of Drying")
    moisture_result = st.text_input("Result for Moisture/Loss of Drying")
    moisture_method = st.text_input("Method for Moisture/Loss of Drying")
    particle_size_spec = st.text_input("Specification for Particle Size")
    particle_size_result = st.text_input("Result for Particle Size")
    particle_size_method = st.text_input("Method for Particle Size")
    bulk_density_spec = st.text_input("Specification for Bulk Density")
    bulk_density_result = st.text_input("Result for Bulk Density")
    bulk_density_method = st.text_input("Method for Bulk Density")
    tapped_density_spec = st.text_input("Specification for Tapped Density")
    tapped_density_result = st.text_input("Result for Tapped Density")
    tapped_density_method = st.text_input("Method for Tapped Density")
    ash_contents_spec = st.text_input("Specification for Ash Contents")
    ash_contents_result = st.text_input("Result for Ash Contents")
    ash_contents_method = st.text_input("Method for Ash Contents")
    ph_spec = st.text_input("Specification for pH")
    ph_result = st.text_input("Result for pH")
    ph_method = st.text_input("Method for pH")
    fats_spec = st.text_input("Specification for Fats")
    fats_result = st.text_input("Result for Fats")
    fats_method = st.text_input("Method for Fats")
    protein_spec = st.text_input("Specification for Protein")
    protein_result = st.text_input("Result for Protein")
    protein_method = st.text_input("Method for Protein")
    solubility_spec = st.text_input("Specification for Solubility")
    solubility_result = st.text_input("Result for Solubility")
    solubility_method = st.text_input("Method for Solubility")
    oxalic_acid_spec = st.text_input("Specification for Limit of Oxalic Acid")
    oxalic_acid_result = st.text_input("Result for Limit of Oxalic Acid")
    oxalic_acid_method = st.text_input("Method for Limit of Oxalic Acid")
    nacl_spec = st.text_input("Specification for Limit of NaCl")
    nacl_result = st.text_input("Result for Limit of NaCl")
    nacl_method = st.text_input("Method for Limit of NaCl")
    sulphates_spec = st.text_input("Specification for Sulphates")
    sulphates_result = st.text_input("Result for Sulphates")
    sulphates_method = st.text_input("Method for Sulphates")
    chloride_spec = st.text_input("Specification for Chloride")
    chloride_result = st.text_input("Result for Chloride")
    chloride_method = st.text_input("Method for Chloride")

    st.subheader("Others")
    heavy_metals_spec = st.text_input("Specification for Heavy Metals")
    heavy_metals_result = st.text_input("Result for Heavy Metals")
    heavy_metals_method = st.text_input("Method for Heavy Metals")
    lead_spec = st.text_input("Specification for Lead")
    lead_result = st.text_input("Result for Lead")
    lead_method = st.text_input("Method for Lead")
    cadmium_spec = st.text_input("Specification for Cadmium")
    cadmium_result = st.text_input("Result for Cadmium")
    cadmium_method = st.text_input("Method for Cadmium")
    arsenic_spec = st.text_input("Specification for Arsenic")
    arsenic_result = st.text_input("Result for Arsenic")
    arsenic_method = st.text_input("Method for Arsenic")
    mercury_spec = st.text_input("Specification for Mercury")
    mercury_result = st.text_input("Result for Mercury")
    mercury_method = st.text_input("Method for Mercury")

    st.subheader("Chemicals")
    assays_spec = st.text_input("Specification for Assays")
    assays_result = st.text_input("Result for Assays")
    assays_method = st.text_input("Method for Assays")
    extraction_spec = st.text_input("Specification for Extraction Ratio")
    extraction_result = st.text_input("Result for Extraction Ratio")
    extraction_method = st.text_input("Method for Extraction Ratio")

    st.subheader("Pesticides")
    pesticide_spec = st.text_input("Specification for Pesticide")
    pesticide_result = st.text_input("Result for Pesticide")
    pesticide_method = st.text_input("Method for Pesticide")

    st.subheader("Microbiological Profile")
    total_plate_count_spec = st.text_input("Specification for Total Plate Count")
    total_plate_count_result = st.text_input("Result for Total Plate Count")
    total_plate_count_method = st.text_input("Method for Total Plate Count")
    yeasts_mould_spec = st.text_input("Specification for Yeasts & Mould Count")
    yeasts_mould_result = st.text_input("Result for Yeasts & Mould Count")
    yeasts_mould_method = st.text_input("Method for Yeasts & Mould Count")
    e_coli_spec = st.text_input("Specification for E.coli")
    e_coli_result = st.text_input("Result for E.coli")
    e_coli_method = st.text_input("Method for E.coli")
    salmonella_spec = st.text_input("Specification for Salmonella")
    salmonella_result = st.text_input("Result for Salmonella")
    salmonella_method = st.text_input("Method for Salmonella")
    coliforms_spec = st.text_input("Specification for Coliforms")
    coliforms_result = st.text_input("Result for Coliforms")
    coliforms_method = st.text_input("Method for Coliforms")

    # Input for the name to save the PDF
    pdf_filename = st.text_input("Enter the filename for the PDF (without extension):", "COA")

    submitted = st.form_submit_button("Generate and Download PDF")

if submitted:
    # Ensure required fields are filled
    required_fields = [product_name, batch_no, manufacturing_date, reanalysis_date]
    if not all(required_fields):
        st.error("Please fill in all required fields.")
    else:
        data = {
            "product_name": product_name,
            "botanical_name": botanical_name,
            "chemical_name": chemical_name,
            "cas_no": cas_no,
            "product_code": product_code,
            "batch_no": batch_no,
            "manufacturing_date": manufacturing_date,  # No conversion needed
            "reanalysis_date": reanalysis_date,  # No conversion needed
            "quantity": quantity,
            "source": source,
            "origin": origin,
            "plant_part": plant_part,
            "extraction_ratio": extraction_ratio,
            "solvent": solvent,

            "identification_spec": identification_spec,
            "identification_result": identification_result,
            "identification_method": identification_method,

            "description_spec": description_spec,
            "description_result": description_result,
            "description_method": description_method,

            "moisture_spec": moisture_spec,
            "moisture_result": moisture_result,
            "moisture_method": moisture_method,

            "particle_size_spec": particle_size_spec,
            "particle_size_result": particle_size_result,
            "particle_size_method": particle_size_method,

            "bulk_density_spec": bulk_density_spec,
            "bulk_density_result": bulk_density_result,
            "bulk_density_method": bulk_density_method,

            "tapped_density_spec": tapped_density_spec,
            "tapped_density_result": tapped_density_result,
            "tapped_density_method": tapped_density_method,

            "ash_contents_spec": ash_contents_spec,
            "ash_contents_result": ash_contents_result,
            "ash_contents_method": ash_contents_method,

            "ph_spec": ph_spec,
            "ph_result": ph_result,
            "ph_method": ph_method,

            "fats_spec": fats_spec,
            "fats_result": fats_result,
            "fats_method": fats_method,

            "protein_spec": protein_spec,
            "protein_result": protein_result,
            "protein_method": protein_method,

            "solubility_spec": solubility_spec,
            "solubility_result": solubility_result,
            "solubility_method": solubility_method,

            "oxalic_acid_spec": oxalic_acid_spec,
            "oxalic_acid_result": oxalic_acid_result,
            "oxalic_acid_method": oxalic_acid_method,

            "nacl_spec": nacl_spec,
            "nacl_result": nacl_result,
            "nacl_method": nacl_method,

            "sulphates_spec": sulphates_spec,
            "sulphates_result": sulphates_result,
            "sulphates_method": sulphates_method,

            "chloride_spec": chloride_spec,
            "chloride_result": chloride_result,
            "chloride_method": chloride_method,

            "heavy_metals_spec": heavy_metals_spec,
            "heavy_metals_result": heavy_metals_result,
            "heavy_metals_method": heavy_metals_method,

            "lead_spec": lead_spec,
            "lead_result": lead_result,
            "lead_method": lead_method,

            "cadmium_spec": cadmium_spec,
            "cadmium_result": cadmium_result,
            "cadmium_method": cadmium_method,

            "arsenic_spec": arsenic_spec,
            "arsenic_result": arsenic_result,
            "arsenic_method": arsenic_method,

            "mercury_spec": mercury_spec,
            "mercury_result": mercury_result,
            "mercury_method": mercury_method,

            "assays_spec": assays_spec,
            "assays_result": assays_result,
            "assays_method": assays_method,

            "extraction_spec": extraction_spec,
            "extraction_result": extraction_result,
            "extraction_method": extraction_method,

            "pesticide_spec": pesticide_spec,
            "pesticide_result": pesticide_result,
            "pesticide_method": pesticide_method,

            "total_plate_count_spec": total_plate_count_spec,
            "total_plate_count_result": total_plate_count_result,
            "total_plate_count_method": total_plate_count_method,

            "yeasts_mould_spec": yeasts_mould_spec,
            "yeasts_mould_result": yeasts_mould_result,
            "yeasts_mould_method": yeasts_mould_method,

            "e_coli_spec": e_coli_spec,
            "e_coli_result": e_coli_result,
            "e_coli_method": e_coli_method,

            "salmonella_spec": salmonella_spec,
            "salmonella_result": salmonella_result,
            "salmonella_method": salmonella_method,

            "coliforms_spec": coliforms_spec,
            "coliforms_result": coliforms_result,
            "coliforms_method": coliforms_method,
        }

        pdf_buffer = generate_pdf(data)

        if pdf_buffer:
            # Provide download button with user-defined filename
            st.download_button(
                label="Download COA PDF",
                data=pdf_buffer,
                file_name=f"{pdf_filename}.pdf",
                mime="application/pdf"
            )

            # Success message
            st.success("COA PDF generated successfully!")
