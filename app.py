import os
import streamlit as st
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

# Function to generate the PDF based on COA structure
def generate_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=10, bottomMargin=10)  # Adjusted margins

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title_style', fontSize=13, spaceAfter=2, alignment=1, fontName='Helvetica-Bold')
    normal_style = styles['BodyText']
    normal_style.alignment = 1  # Center alignment for paragraphs
    
    elements = []
    
    # Set paths for images
    logo_path = os.path.join(os.getcwd(), "images", "tru_herb_logo.png")
    footer_path = os.path.join(os.getcwd(), "images", "footer.png")
    
    # Check if images exist
    if not os.path.exists(logo_path) or not os.path.exists(footer_path):
        st.error("One or more images are missing.")
        return
    
    # Add logo and title at the top
    elements.append(Image(logo_path, width=100, height=40))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("CERTIFICATE OF ANALYSIS", title_style))
    elements.append(Paragraph(data.get('product_name', ''), title_style))  # Add product name
    elements.append(Spacer(1, 20))
    
    # Product Information Table
    product_info = [
    ["Product Name", data.get('product_name', '')],
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
]

    product_table = Table(product_info, colWidths=[150, 400])
    product_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(product_table)
    elements.append(Spacer(1, 20))
    
    # Specifications Table
    # Define the table headers
    spec_headers = ["Parameter", "Specification", "Result", "Method"]
    
    # Physical Section
    physical_params = [
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
    ]
    
    # Others Section
    others_params = [
        ("Lead", data['lead_spec'], data['lead_result'], data['lead_method']),
        ("Cadmium", data['cadmium_spec'], data['cadmium_result'], data['cadmium_method']),
        ("Arsenic", data['arsenic_spec'], data['arsenic_result'], data['arsenic_method']),
        ("Mercury", data['mercury_spec'], data['mercury_result'], data['mercury_method']),
    ]
    
    # Chemicals Section
    chemicals_params = [
        ("Extraction", data['extraction_spec'], data['extraction_result'], data['extraction_method']),
    ]
    
    # Pesticides Section
    pesticides_params = [
        ("Pesticide", data['pesticide_spec'], data['pesticide_result'], data['pesticide_method']),
    ]
    
    # Microbiological Profile Section
    micro_params = [
        ("Total Plate Count", data['total_plate_count_spec'], data['total_plate_count_result'], data['total_plate_count_method']),
        ("Yeasts & Mould Count", data['yeasts_mould_spec'], data['yeasts_mould_result'], data['yeasts_mould_method']),
        ("E.coli", data['e_coli_spec'], data['e_coli_result'], data['e_coli_method']),
        ("Salmonella", data['salmonella_spec'], data['salmonella_result'], data['salmonella_method']),
        ("Coliforms", data['coliforms_spec'], data['coliforms_result'], data['coliforms_method']),
    ]
    
    # Build the specifications table
    spec_data = []
    
    # Physical Section Header
    spec_data.append([Paragraph("<b>Physical</b>", styles['Normal'])])
    spec_data[-1].extend(["", "", ""])  # Empty cells to fill the row
    # Physical Parameters
    for param in physical_params:
        spec_data.append(list(param))
    
    # Others Section Header
    spec_data.append([Paragraph("<b>Others</b>", styles['Normal'])])
    spec_data[-1].extend(["", "", ""])
    # Others Parameters
    for param in others_params:
        spec_data.append(list(param))
    
    # Chemicals Section Header
    spec_data.append([Paragraph("<b>Chemicals</b>", styles['Normal'])])
    spec_data[-1].extend(["", "", ""])
    # Chemicals Parameters
    for param in chemicals_params:
        spec_data.append(list(param))
    
    # Pesticides Section Header
    spec_data.append([Paragraph("<b>Pesticides</b>", styles['Normal'])])
    spec_data[-1].extend(["", "", ""])
    # Pesticides Parameters
    for param in pesticides_params:
        spec_data.append(list(param))
    
    # Microbiological Profile Section Header
    spec_data.append([Paragraph("<b>Microbiological Profile</b>", styles['Normal'])])
    spec_data[-1].extend(["", "", ""])
    # Microbiological Parameters
    for param in micro_params:
        spec_data.append(list(param))
    
    # Build the table
    spec_table = Table([spec_headers] + spec_data, colWidths=[120, 100, 100, 100]) #keep this dynamic?
    spec_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('SPAN', (0,1), (-1,1)),  # Span across for section headers
        ('SPAN', (0,12), (-1,12)),
        ('SPAN', (0,17), (-1,17)),
        ('SPAN', (0,19), (-1,19)),
        ('SPAN', (0,21), (-1,21)),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(spec_table)
    elements.append(Spacer(1, 20))
    
    # Remarks Section
    remarks_text = "Since the product is derived from natural origin, there is likely to be minor color variation because of the geographical and seasonal variations of the raw material"
    remarks = Paragraph(remarks_text, styles['Normal'])
    elements.append(remarks)
    elements.append(Spacer(1, 10))
    
    # Final Remark
    final_remark = Paragraph("<b>REMARKS: COMPLIES WITH IN HOUSE SPECIFICATIONS</b>", styles['Normal'])
    elements.append(final_remark)
    elements.append(Spacer(1, 20))
    
    # Declaration Section
    declaration_data = [
        [
            Paragraph("<b>Declaration:-</b><br/>"
                      "GMO status: Free from GMO<br/>"
                      "Irradiation status: Non–Irradiated<br/><br/>"
                      "Prepared by<br/>"
                      "Executive – QC", styles['Normal']),
            Paragraph("Allergen statement: Free from allergen<br/>"
                      "Storage condition: At room temperature<br/><br/>"
                      "Approved by<br/>"
                      "Head-QC/QA", styles['Normal'])
        ]
    ]
    
    declaration_table = Table(declaration_data, colWidths=[300, 300])
    declaration_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0, colors.white),  # No grid lines
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(declaration_table)
    elements.append(Spacer(1, 20))
    
    # Footer image with increased height
    elements.append(Image(footer_path, width=500, height=100))  # Doubled height for footer
    
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
    product_code = st.text_input("Product Code")
    batch_no = st.text_input("Batch No.")
    manufacturing_date = st.text_input("Date of Manufacturing")
    reanalysis_date = st.text_input("Date of Reanalysis")
    quantity = st.text_input("Quantity (in Kgs)")
    source = st.text_input("Source")
    origin = st.text_input("Country of Origin")
    plant_part = st.text_input("Plant Parts")
    extraction_ratio = st.text_input("Extraction Ratio")
    solvent = st.text_input("Extraction Solvents")
    
    st.header("Specifications")
    st.subheader("Physical")
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
    
    st.subheader("Others")
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
    extraction_spec = st.text_input("Specification for Extraction")
    extraction_result = st.text_input("Result for Extraction")
    extraction_method = st.text_input("Method for Extraction")
    
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
    if not product_name or not batch_no or not manufacturing_date or not reanalysis_date:
        st.error("Please fill in all required fields.")
    else:
        data = {
            "product_name": product_name,
            "product_code": product_code,
            "batch_no": batch_no,
            "manufacturing_date": manufacturing_date,
            "reanalysis_date": reanalysis_date,
            "quantity": quantity,
            "source": source,
            "origin": origin,
            "plant_part": plant_part,
            "extraction_ratio": extraction_ratio,
            "solvent": solvent,
            
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
        
        # Provide download button with user-defined filename
        st.download_button("Download COA PDF", data=pdf_buffer, file_name=f"{pdf_filename}.pdf", mime="application/pdf")
        
        # Success message
        st.success("COA PDF generated successfully!")
