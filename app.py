import os
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
import base64
import fitz  # PyMuPDF
import configparser

# Initialize the configparser
config = configparser.ConfigParser()

# Read the config file
config.read('.streamlit/config.toml')

# Access configuration values
# For example, if you have a section [settings] with a key 'theme'
theme = config.get('settings', 'theme', fallback='default')

# Use the configuration values in your app
st.set_page_config(page_title="Tru Herb COA PDF Generator", layout="wide")

# Function to generate the PDF based on COA structure
def generate_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=10, bottomMargin=10)

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title_style', fontSize=13, spaceAfter=1, alignment=1, fontName='Helvetica-Bold')
    title_style1 = ParagraphStyle('title_style1', fontSize=10, spaceAfter=0, alignment=1, fontName='Helvetica-Bold')
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
        ["Product Name", Paragraph(data.get('product_name', ''), normal_style)],
        ["Product Code", Paragraph(data.get('product_code', ''), normal_style)],
        ["Batch No.", Paragraph(data.get('batch_no', ''), normal_style)],
        ["Date of Manufacturing", Paragraph(data.get('manufacturing_date', ''), normal_style)],
        ["Date of Reanalysis", Paragraph(data.get('reanalysis_date', ''), normal_style)],
        ["Botanical Name", Paragraph(data.get('botanical_name', ''), normal_style)],
        ["Extraction Ratio", Paragraph(data.get('extraction_ratio', ''), normal_style)],
        ["Extraction Solvents", Paragraph(data.get('solvent', ''), normal_style)],
        ["Plant Parts", Paragraph(data.get('plant_part', ''), normal_style)],
        ["CAS No.", Paragraph(data.get('cas_no', ''), normal_style)],
        ["Chemical Name", Paragraph(data.get('chemical_name', ''), normal_style)],
        ["Quantity (in Kgs)", Paragraph(data.get('quantity', ''), normal_style)],
        ["Country of Origin", Paragraph(data.get('origin', ''), normal_style)],
    ]

    # Only include non-empty rows
    product_info = [row for row in product_info if row[1].text]

    if product_info:
        product_table = Table(product_info, colWidths=[140, 360])
        product_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Ensure text is top-aligne   d
            ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),  # Enable word wrapping
        ]))
        elements.append(product_table)
        elements.append(Spacer(1, 0))

    # Specifications Table
    spec_headers = ["Parameter", "Specification", "Result", "Method"]

    # Define sections with filtering to exclude empty entries
    sections = {
        "Physical": [
            ("Identification", data['identification_spec'], data['identification_result'], data['identification_method']) if data['identification_spec'] and data['identification_result'] and data['identification_method'] else None,
            ("Description", data['description_spec'], data['description_result'], data['description_method']) if data['description_spec'] and data['description_result'] and data['description_method'] else None,
            ("Moisture/Loss of Drying", data['moisture_spec'], data['moisture_result'], data['moisture_method']) if data['moisture_spec'] and data['moisture_result'] and data['moisture_method'] else None,
            ("Particle Size", data['particle_size_spec'], data['particle_size_result'], data['particle_size_method']) if data['particle_size_spec'] and data['particle_size_result'] and data['particle_size_method'] else None,
            ("Bulk Density", data['bulk_density_spec'], data['bulk_density_result'], data['bulk_density_method']) if data['bulk_density_spec'] and data['bulk_density_result'] and data['bulk_density_method'] else None,
            ("Tapped Density", data['tapped_density_spec'], data['tapped_density_result'], data['tapped_density_method']) if data['tapped_density_spec'] and data['tapped_density_result'] and data['tapped_density_method'] else None,
            ("Ash Contents", data['ash_contents_spec'], data['ash_contents_result'], data['ash_contents_method']) if data['ash_contents_spec'] and data['ash_contents_result'] and data['ash_contents_method'] else None,
            ("pH", data['ph_spec'], data['ph_result'], data['ph_method']) if data['ph_spec'] and data['ph_result'] and data['ph_method'] else None,
            ("Fats", data['fats_spec'], data['fats_result'], data['fats_method']) if data['fats_spec'] and data['fats_result'] and data['fats_method'] else None,
            ("Protein", data['protein_spec'], data['protein_result'], data['protein_method']) if data['protein_spec'] and data['protein_result'] and data['protein_method'] else None,
            ("Solubility", data['solubility_spec'], data['solubility_result'], data['solubility_method']) if data['solubility_spec'] and data['solubility_result'] and data['solubility_method'] else None,
            ("Limit of Oxalic Acid", data['oxalic_acid_spec'], data['oxalic_acid_result'], data['oxalic_acid_method']) if data['oxalic_acid_spec'] and data['oxalic_acid_result'] and data['oxalic_acid_method'] else None,
            ("Limit of NaCl", data['nacl_spec'], data['nacl_result'], data['nacl_method']) if data['nacl_spec'] and data['nacl_result'] and data['nacl_method'] else None,
            ("Sulphates", data['sulphates_spec'], data['sulphates_result'], data['sulphates_method']) if data['sulphates_spec'] and data['sulphates_result'] and data['sulphates_method'] else None,
            ("Chloride", data['chloride_spec'], data['chloride_result'], data['chloride_method']) if data['chloride_spec'] and data['chloride_result'] and data['chloride_method'] else None,
        ],
        "Others": [
            ("Heavy Metals", data['heavy_metals_spec'], data['heavy_metals_result'], data['heavy_metals_method']) if data['heavy_metals_spec'] and data['heavy_metals_result'] and data['heavy_metals_method'] else None,
            ("Lead", data['lead_spec'], data['lead_result'], data['lead_method']) if data['lead_spec'] and data['lead_result'] and data['lead_method'] else None,
            ("Cadmium", data['cadmium_spec'], data['cadmium_result'], data['cadmium_method']) if data['cadmium_spec'] and data['cadmium_result'] and data['cadmium_method'] else None,
            ("Arsenic", data['arsenic_spec'], data['arsenic_result'], data['arsenic_method']) if data['arsenic_spec'] and data['arsenic_result'] and data['arsenic_method'] else None,
            ("Mercury", data['mercury_spec'], data['mercury_result'], data['mercury_method']) if data['mercury_spec'] and data['mercury_result'] and data['mercury_method'] else None,
        ],
        "Chemicals": [
            ("Assays", data['assays_spec'], data['assays_result'], data['assays_method']) if data['assays_spec'] and data['assays_result'] and data['assays_method'] else None,
            ("Extraction", data['extraction_spec'], data['extraction_result'], data['extraction_method']) if data['extraction_spec'] and data['extraction_result'] and data['extraction_method'] else None,
        ],
        "Pesticides": [
            ("Pesticide", data['pesticide_spec'], data['pesticide_result'], data['pesticide_method']) if data['pesticide_spec'] and data['pesticide_result'] and data['pesticide_method'] else None,
        ],
        "Microbiological Profile": [
            ("Total Plate Count", data['total_plate_count_spec'], data['total_plate_count_result'], data['total_plate_count_method']) if data['total_plate_count_spec'] and data['total_plate_count_result'] and data['total_plate_count_method'] else None,
            ("Yeasts & Mould Count", data['yeasts_mould_spec'], data['yeasts_mould_result'], data['yeasts_mould_method']) if data['yeasts_mould_spec'] and data['yeasts_mould_result'] and data['yeasts_mould_method'] else None,
            ("E.coli", data['e_coli_spec'], data['e_coli_result'], data['e_coli_method']) if data['e_coli_spec'] and data['e_coli_result'] and data['e_coli_method'] else None,
            ("Salmonella", data['salmonella_spec'], data['salmonella_result'], data['salmonella_method']) if data['salmonella_spec'] and data['salmonella_result'] and data['salmonella_method'] else None,
            ("Coliforms", data['coliforms_spec'], data['coliforms_result'], data['coliforms_method']) if data['coliforms_spec'] and data['coliforms_result'] and data['coliforms_method'] else None,
        ]
    }

    spec_data = [spec_headers]

    # Add sections to spec_data with filtering
    for section, params in sections.items():
        filtered_params = [param for param in params if param is not None]
        if filtered_params:
            spec_data.append([Paragraph(f"<b>{section}</b>", styles['Normal']), "", "", ""])
            spec_data.extend([
                [Paragraph(str(item), normal_style) for item in param] for param in filtered_params
            ])

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
    total_width = 500  # Total width of all columns combined
    num_columns = len(spec_data[0]) if spec_data else 4  # Get number of columns from data, default to 4
    # Calculate widths - first column slightly smaller, second slightly larger, rest equal
    col_widths = [
        total_width * 0.24,  # 24% for first column
        total_width * 0.28,  # 28% for second column
        total_width * 0.24,  # 24% for third column
        total_width * 0.24   # 24% for fourth column
    ]
    spec_table = Table(spec_data, colWidths=col_widths)
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
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Ensure text is top-aligned
        ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),  # Enable word wrapping
    ]))
    elements.append(spec_table)
    elements.append(Spacer(1, 2))

    # Declaration
    elements.append(Paragraph("Declaration", title_style1))

    # Declaration Table
    declaration_data = [
        ["GMO Status:", Paragraph("Free from GMO", normal_style), "", "Allergen statement:", Paragraph("Free from allergen", normal_style)],
        ["Irradiation status:", Paragraph("Non – Irradiated", normal_style), "", "Storage condition:", Paragraph("At room temperature", normal_style)],
        ["Prepared by", Paragraph("Executive – QC", normal_style), "", "Approved by", Paragraph("Head-QC/QA", normal_style)]
    ]

    declaration_table = Table(declaration_data, colWidths=[80, 150, 75, 100, 95])
    declaration_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (1, -1), 'LEFT'),
        ('ALIGN', (3, 0), (4, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Ensure text is top-aligned
        ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),  # Enable word wrapping
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('SPAN', (2, 0), (2, 2)),  # Span the middle empty column for spacing
    ]))

    elements.append(declaration_table)

    # Reduce the number of spacers between the specifications table and footer image
    elements.append(Spacer(1, 3))  # Only one spacer as requested

    # Footer Image
    elements.append(Image(footer_path, width=500, height=80))

    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title("Tru Herb COA PDF Generator")

# Create two columns for layout
col1, col2 = st.columns(2)

# Form for user input in the left column
with col1.form("coa_form"):
    st.header("Product Information")
    product_name = st.text_input("Product Name", value="Product Name")
    product_code = st.text_input("Product Code")
    batch_no = st.text_input("Batch No.")
    manufacturing_date = st.text_input("Date of Manufacturing")  # Changed to text input
    reanalysis_date = st.text_input("Date of Reanalysis")  # Changed to text input
    botanical_name = st.text_input("Botanical Name")
    extraction_ratio = st.text_input("Extraction Ratio")
    solvent = st.text_input("Extraction Solvents")
    plant_part = st.text_input("Plant Parts")
    cas_no = st.text_input("CAS No.")
    chemical_name = st.text_input("Chemical Name")
    quantity = st.text_input("Quantity (in Kgs)")
    origin = st.text_input("Country of Origin")
    # source = st.text_input("Source")

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

    # Preview button
    preview_button = st.form_submit_button("Preview")

    # Generate and download button
    download_button = st.form_submit_button("Download PDF")

# Handle preview in the right column
if preview_button:
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
        # "source": source,
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
        # Use PyMuPDF to convert PDF to images
        doc = fitz.open(stream=pdf_buffer, filetype="pdf")
        # Display each page as an image in the right column
        with col2:
            for page in doc:
                pix = page.get_pixmap()
                # Adjust the width and height as needed
                st.image(pix.tobytes(), caption=f"Page {page.number + 1}", width=900, use_column_width=False)
            st.success("Preview generated successfully!")

# Handle PDF download
if download_button:
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
        # "source": source,
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
        st.success("COA PDF generated and ready for download!")
        
# st.markdown(
#     """
#     <style>
#     .input-row {
#         display: flex;
#         align-items: center;
#         margin-bottom: 10px;
#     }
#     .input-row label {
#         margin-right: 10px;
#         font-weight: bold;
#     }
#     .input-row input {
#         flex: 1;
#     }
#     </style>
#     """,
#     unsafe_allow_html=True
# )

# # Example usage in Streamlit
# st.markdown('<div class="input-row"><label for="product_name">Product Name</label><input type="text" id="product_name" name="product_name"></div>', unsafe_allow_html=True)
