import os
import streamlit as st
from reportlab.lib.pagesizes import A4  # Import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io

# Function to generate the PDF based on COA structure
def generate_pdf(data, specifications):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0)  # Set pagesize to A4 and topMargin to 0

    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title_style', fontSize=16, spaceAfter=3, alignment=1, bold=True)
    normal_style = styles['BodyText']

    elements = []

    # Set paths for images
    logo_path = os.path.join(os.getcwd(), "images", "tru_herb_logo.png")
    footer_path = os.path.join(os.getcwd(), "images", "footer.png")

    # Check if images exist
    if not os.path.exists(logo_path) or not os.path.exists(footer_path):
        st.error("One or more images are missing.")
        return

    # Add logo and title
    elements.append(Image(logo_path, width=100, height=40))
    elements.append(Spacer(1, 5))  # Reduced space for logo position
    elements.append(Paragraph("CERTIFICATE OF ANALYSIS", title_style))
    elements.append(Paragraph(data.get('product_name', ''), title_style))  # Add product name
    elements.append(Spacer(1, 10))

    # Product Information Table
    product_info = [
        ["Product", data.get('product_name')],
        ["Product Code", data.get('product_code')],
        ["Batch No", data.get('batch_no')],
        ["Date of Manufacturing", data.get('manufacturing_date')],
        ["Date of Re-Analysis", data.get('re_analysis_date')],
        ["Quantity (Kg)", f"{data.get('quantity', 0)} Kg" if data.get('quantity') else ''],
        ["Source", data.get('source')],
        ["Country of Origin", data.get('origin')],
        ["Plant Parts", data.get('plant_part')],
        ["Extraction Ratio", data.get('extraction_ratio')],
        ["Extraction Solvent", data.get('solvent')],
    ]
    
    product_info = [row for row in product_info if row[1]]  # Filter out empty fields

    product_table = Table(product_info, colWidths=[200, 300])
    product_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(product_table)
    elements.append(Spacer(1, 20))

    # Specifications Table
    spec_headers = ["Parameter", "Specification", "Result", "Method"]
    specifications_data = [
        ["Physical", "", "", ""],
        ["Description", "", "", ""],
        ["Moisture/Loss of Drying", "", "", ""],
        ["Particle Size", "", "", ""],
        ["Bulk Density", "", "", ""],
        ["Tapped Density", "", "", ""],
        ["Ash Content", "", "", ""],
        ["pH", "", "", ""],
        ["Fats", "", "", ""],
        ["Protein", "", "", ""],
        ["Solubility", "", "", ""],
        ["Others", "", "", ""],
        ["Cadmium", "", "", ""],
        ["Arsenic", "", "", ""],
        ["Mercury", "", "", ""],
        ["Chemicals", "", "", ""],
        ["Extraction", "", "", ""],
        ["Pesticides", "", "", ""],
        ["Pesticide Residue", "", "", ""],
        ["Microbiological Profile", "", "", ""],
        ["Total Plate Count", "", "", ""],
        ["Yeasts and Mould Count", "", "", ""],
        ["E. coli", "", "", ""],
        ["Salmonella", "", "", ""],
        ["Coliforms", "", "", ""]
    ]

    # Filtering out empty specification rows
    specifications_data = [row for row in specifications_data if row[1] or row[2] or row[3]]

    spec_table = Table([spec_headers] + specifications_data, colWidths=[150, 150, 100, 200])
    spec_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (0, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT')
    ]))
    elements.append(spec_table)
    elements.append(Spacer(1, 20))

    # Remarks Section
    remarks = """
    <p style="text-align: center; font-weight: bold;">
        Since the product is derived from natural origin, there is likely to be minor color variation 
        due to geographical and seasonal variations of the raw material.
    </p>
    <p style="text-align: center; font-weight: bold;">REMARKS: COMPLIES WITH IN HOUSE SPECIFICATIONS</p>
    """
    elements.append(Paragraph(remarks, normal_style))
    elements.append(Spacer(1, 20))

    # Declaration Section
    
    declaration_data = [
        [
            Paragraph("<b>Declaration:</b><br/>"
                    "<b>GMO Status:</b> Free from GMO<br/>"
                    "<b>Irradiation Status:</b> Non-irradiated<br/>"
                    "<b>Prepared by:</b><br/>"
                    "Executive – QC", normal_style),
            Paragraph("<b>Allergen Statement:</b> Free from allergens<br/>"
                    "<b>Storage Condition:</b> At room temperature<br/>"
                    "<b>Approved by:</b><br/>"
                    "Head-QC/QA", normal_style)
        ]
    ]

    # Create an invisible table for the declaration
    declaration_table = Table(declaration_data, colWidths=[300, 300])
    declaration_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0, colors.white),  # Invisible borders
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),  # Align text to the left
        ('LEFTPADDING', (0, 0), (-1, -1), 10),  # Add left padding
        ('RIGHTPADDING', (0, 0), (-1, -1), 10)   # Add right padding
    ]))

    elements.append(declaration_table)
    elements.append(Spacer(1, 20))



    # Footer image with increased height
    elements.append(Image(footer_path, width=500, height=100))  # Increased height for footer

    # Build PDF
    doc.build(elements)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title("Tru Herb COA PDF Generator")

# Form for user input
with st.form("coa_form"):
    product_name = st.text_input("Product Name")
    product_code = st.text_input("Product Code")
    batch_no = st.text_input("Batch No")
    manufacturing_date = st.date_input("Manufacturing Date")
    re_analysis_date = st.date_input("Re-Analysis Date")
    quantity = st.number_input("Quantity (Kg)", min_value=1)
    source = st.text_input("Source")
    origin = st.text_input("Country of Origin")
    plant_part = st.text_input("Plant Parts")
    extraction_ratio = st.text_input("Extraction Ratio")
    solvent = st.text_input("Extraction Solvent")

    # Specifications input
    st.subheader("Specifications")
    specifications = []
    for i in range(5):  # Allow users to input specifications
        param_name = st.text_input(f"Parameter Name {i+1}", key=f"param_name_{i}")
        specification = st.text_input(f"Specification for {param_name}", key=f"specification_{i}")
        result = st.text_input(f"Result for {param_name}", key=f"result_{i}")
        method_protocol = st.text_input(f"Method/Protocols for {param_name}", key=f"method_protocol_{i}")

        if param_name and specification:  # Only add if parameter name and specification are filled
            specifications.append([param_name, specification, result, method_protocol])
    
    # Input for the name to save the PDF
    pdf_filename = st.text_input("Enter the filename for the PDF (without extension):", "COA")
    
    submitted = st.form_submit_button("Generate and Download PDF")

if submitted:
    # Ensure required fields are filled
    if not product_name or not batch_no or not manufacturing_date or not re_analysis_date:
        st.error("Please fill in all required fields.")
    else:
        data = {
            "product_name": product_name,
            "product_code": product_code,
            "batch_no": batch_no,
            "manufacturing_date": manufacturing_date.strftime("%B %Y"),
            "re_analysis_date": re_analysis_date.strftime("%B %Y"),
            "quantity": quantity,
            "source": source,
            "origin": origin,
            "plant_part": plant_part,
            "extraction_ratio": extraction_ratio,
            "solvent": solvent,
        }

        pdf_buffer = generate_pdf(data, specifications)

        # Provide download button with user-defined filename
        st.download_button("Download COA PDF", data=pdf_buffer, file_name=f"{pdf_filename}.pdf", mime="application/pdf")

        # Success message
        st.success("COA PDF generated successfully!")
