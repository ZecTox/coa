import os
from typing import Container
import streamlit as st

# ReportLab imports
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

import io
import fitz  # PyMuPDF
import configparser

# PIL used for combining pages into a single image
from PIL import Image

# Initialize the configparser
config = configparser.ConfigParser()
config.read('.streamlit/config.toml')
theme = config.get('settings', 'theme', fallback='default')

st.set_page_config(page_title="Tru Herb COA PDF Generator", layout="wide")

#####################################
# Page Header/Footer in PDF
#####################################
def header_footer(canvas, doc):
    canvas.saveState()
    logo_path = os.path.join(os.getcwd(), "images", "tru_herb_logo.png")
    footer_path = os.path.join(os.getcwd(), "images", "footer.png")

    if os.path.exists(logo_path):
        canvas.drawImage(logo_path, x=250, y=A4[1] - 65, width=100, height=50)
    if os.path.exists(footer_path):
        canvas.drawImage(footer_path, x=50, y=20, width=500, height=70)
    canvas.restoreState()

#####################################
# Main PDF Generation
#####################################
def generate_pdf(data):
    """
    Builds the PDF using ReportLab. If the PDF ends up with multiple pages,
    we combine them into a single, scaled page.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=60, bottomMargin=80)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('title_style', fontSize=12, spaceAfter=1, alignment=1, fontName='Times-Bold')
    title_style1 = ParagraphStyle('title_style1', fontSize=10, spaceAfter=0, alignment=1, fontName='Times-Bold')
    normal_style = styles['BodyText']
    normal_style.fontName = 'Times-Roman'
    normal_style.alignment = 0
    style_for_sections = styles["Normal"]

    elements = []
    elements.append(Spacer(1, 3))
    elements.append(Paragraph("CERTIFICATE OF ANALYSIS", title_style))
    elements.append(Paragraph(data.get('product_name', ''), title_style))
    elements.append(Spacer(1, 3))

    # ----------------------------------------------------------------
    # Build Product Info table, skipping truly empty fields
    # ----------------------------------------------------------------
    product_info = []

    def maybe_add_product_row(label, value, italic=False):
        """Helper to add a product info row only if `value` is non-empty."""
        text_str = value.strip() if value else ""
        if text_str:
            if italic:
                text_str = f"<i>{text_str}</i>"
            # Create the Paragraph only if there's real content
            product_info.append([label, Paragraph(text_str, normal_style)])

    maybe_add_product_row("Product Name", data.get('product_name', ''))
    maybe_add_product_row("Product Code", data.get('product_code', ''))
    maybe_add_product_row("Batch No.", data.get('batch_no', ''))
    maybe_add_product_row("Date of Manufacturing", data.get('manufacturing_date', ''))
    maybe_add_product_row("Date of Reanalysis", data.get('reanalysis_date', ''))
    maybe_add_product_row("Botanical Name", data.get('botanical_name', ''), italic=True)
    maybe_add_product_row("Extraction Ratio", data.get('extraction_ratio', ''))
    maybe_add_product_row("Extraction Solvents", data.get('solvent', ''))
    maybe_add_product_row("Plant Parts", data.get('plant_part', ''))
    maybe_add_product_row("CAS No.", data.get('cas_no', ''))
    maybe_add_product_row("Chemical Name", data.get('chemical_name', ''))
    maybe_add_product_row("Quantity", data.get('quantity', ''))
    maybe_add_product_row("Country of Origin", data.get('origin', ''))

    if product_info:
        product_info.insert(0, ["Parameter", "Value"])
        product_table = Table(product_info, colWidths=[140, 360])
        product_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),
        ]))
        elements.append(product_table)
        elements.append(Spacer(1, 0))

    # ----------------------------------------------------------------
    # Prepare the SPECIFICATIONS table data
    # ----------------------------------------------------------------
    spec_headers = ["Parameter", "Specification", "Result", "Method"]
    spec_data = [spec_headers]  # The first row is the table header

    heading_indices = []

    # Define the sections dictionary
    sections = {
        "Physical": [
            ("Description", data['description_spec'], data['description_result'], data['description_method']) 
                if data['description_spec'] and data['description_result'] and data['description_method'] else None,
            ("Identification", data['identification_spec'], data['identification_result'], data['identification_method']) 
                if data['identification_spec'] and data['identification_result'] and data['identification_method'] else None,
            ("Loss on Drying", data['loss_on_drying_spec'], data['loss_on_drying_result'], data['loss_on_drying_method']) 
                if data['loss_on_drying_spec'] and data['loss_on_drying_result'] and data['loss_on_drying_method'] else None,
            ("Moisture", data['moisture_spec'], data['moisture_result'], data['moisture_method']) 
                if data['moisture_spec'] and data['moisture_result'] and data['moisture_method'] else None,
            ("Particle Size", data['particle_size_spec'], data['particle_size_result'], data['particle_size_method']) 
                if data['particle_size_spec'] and data['particle_size_result'] and data['particle_size_method'] else None,
            ("Ash Contents", data['ash_contents_spec'], data['ash_contents_result'], data['ash_contents_method']) 
                if data['ash_contents_spec'] and data['ash_contents_result'] and data['ash_contents_method'] else None,
            ("Residue on Ignition", data['residue_on_ignition_spec'], data['residue_on_ignition_result'], data['residue_on_ignition_method']) 
                if data['residue_on_ignition_spec'] and data['residue_on_ignition_result'] and data['residue_on_ignition_method'] else None,
            ("Bulk Density", data['bulk_density_spec'], data['bulk_density_result'], data['bulk_density_method']) 
                if data['bulk_density_spec'] and data['bulk_density_result'] and data['bulk_density_method'] else None,
            ("Tapped Density", data['tapped_density_spec'], data['tapped_density_result'], data['tapped_density_method']) 
                if data['tapped_density_spec'] and data['tapped_density_result'] and data['tapped_density_method'] else None,
            ("Solubility", data['solubility_spec'], data['solubility_result'], data['solubility_method']) 
                if data['solubility_spec'] and data['solubility_result'] and data['solubility_method'] else None,
            ("pH", data['ph_spec'], data['ph_result'], data['ph_method']) 
                if data['ph_spec'] and data['ph_result'] and data['ph_method'] else None,
            ("Chlorides of NaCl", data['chlorides_nacl_spec'], data['chlorides_nacl_result'], data['chlorides_nacl_method']) 
                if data['chlorides_nacl_spec'] and data['chlorides_nacl_result'] and data['chlorides_nacl_method'] else None,
            ("Sulphates", data['sulphates_spec'], data['sulphates_result'], data['sulphates_method']) 
                if data['sulphates_spec'] and data['sulphates_result'] and data['sulphates_method'] else None,
            ("Fats", data['fats_spec'], data['fats_result'], data['fats_method']) 
                if data['fats_spec'] and data['fats_result'] and data['fats_method'] else None,
            ("Protein", data['protein_spec'], data['protein_result'], data['protein_method']) 
                if data['protein_spec'] and data['protein_result'] and data['protein_method'] else None,
            ("Total IgG", data['total_ig_g_spec'], data['total_ig_g_result'], data['total_ig_g_method']) 
                if data['total_ig_g_spec'] and data['total_ig_g_result'] and data['total_ig_g_method'] else None,
            ("Sodium", data['sodium_spec'], data['sodium_result'], data['sodium_method']) 
                if data['sodium_spec'] and data['sodium_result'] and data['sodium_method'] else None,
            ("Gluten", data['gluten_spec'], data['gluten_result'], data['gluten_method']) 
                if data['gluten_spec'] and data['gluten_result'] and data['gluten_method'] else None,
        ],
        "Others": [
            ("Lead", data['lead_spec'], data['lead_result'], data['lead_method']) 
                if data['lead_spec'] and data['lead_result'] and data['lead_method'] else None,
            ("Cadmium", data['cadmium_spec'], data['cadmium_result'], data['cadmium_method']) 
                if data['cadmium_spec'] and data['cadmium_result'] and data['cadmium_method'] else None,
            ("Arsenic", data['arsenic_spec'], data['arsenic_result'], data['arsenic_method']) 
                if data['arsenic_spec'] and data['arsenic_result'] and data['arsenic_method'] else None,
            ("Mercury", data['mercury_spec'], data['mercury_result'], data['mercury_method']) 
                if data['mercury_spec'] and data['mercury_result'] and data['mercury_method'] else None,
        ],
        "Assays": [
            ("Assays", data['assays_spec'], data['assays_result'], data['assays_method']) 
                if data['assays_spec'] and data['assays_result'] and data['assays_method'] else None,
        ],
        "Pesticides": [
            ("Pesticide", data['pesticide_spec'], data['pesticide_result'], data['pesticide_method']) 
                if data['pesticide_spec'] and data['pesticide_result'] and data['pesticide_method'] else None,
        ],
        "Residual Solvent": [
            ("Residual Solvent", data['residual_solvent_spec'], data['residual_solvent_result'], data['residual_solvent_method']) 
                if data['residual_solvent_spec'] and data['residual_solvent_result'] and data['residual_solvent_method'] else None,
        ],
        "Microbiological Profile": [
            ("Total Plate Count", data['total_plate_count_spec'], data['total_plate_count_result'], data['total_plate_count_method']) 
                if data['total_plate_count_spec'] and data['total_plate_count_result'] and data['total_plate_count_method'] else None,
            ("Yeasts & Mould Count", data['yeasts_mould_spec'], data['yeasts_mould_result'], data['yeasts_mould_method']) 
                if data['yeasts_mould_spec'] and data['yeasts_mould_result'] and data['yeasts_mould_method'] else None,
            ("Salmonella", data['salmonella_spec'], data['salmonella_result'], data['salmonella_method']) 
                if data['salmonella_spec'] and data['salmonella_result'] and data['salmonella_method'] else None,
            ("Escherichia coli", data['e_coli_spec'], data['e_coli_result'], data['e_coli_method']) 
                if data['e_coli_spec'] and data['e_coli_result'] and data['e_coli_method'] else None,
            ("Coliforms", data['coliforms_spec'], data['coliforms_result'], data['coliforms_method']) 
                if data['coliforms_spec'] and data['coliforms_result'] and data['coliforms_method'] else None,
        ]
    }

    # Append dynamic rows from session_state data
    if "physical_dynamic_rows" in data:
        for row in data["physical_dynamic_rows"]:
            if row["parameter"] and row["spec"] and row["result"] and row["method"]:
                sections["Physical"].append((row["parameter"], row["spec"], row["result"], row["method"]))

    if "others_dynamic_rows" in data:
        for row in data["others_dynamic_rows"]:
            if row["parameter"] and row["spec"] and row["result"] and row["method"]:
                sections["Others"].append((row["parameter"], row["spec"], row["result"], row["method"]))

    if "assays_dynamic_rows" in data:
        for row in data["assays_dynamic_rows"]:
            if row["parameter"] and row["spec"] and row["result"] and row["method"]:
                sections["Assays"].append((row["parameter"], row["spec"], row["result"], row["method"]))

    if "pesticides_dynamic_rows" in data:
        for row in data["pesticides_dynamic_rows"]:
            if row["parameter"] and row["spec"] and row["result"] and row["method"]:
                sections["Pesticides"].append((row["parameter"], row["spec"], row["result"], row["method"]))

    if "residual_solvent_dynamic_rows" in data:
        for row in data["residual_solvent_dynamic_rows"]:
            if row["parameter"] and row["spec"] and row["result"] and row["method"]:
                sections["Residual Solvent"].append((row["parameter"], row["spec"], row["result"], row["method"]))

    if "microbio_dynamic_rows" in data:
        for row in data["microbio_dynamic_rows"]:
            if row["parameter"] and row["spec"] and row["result"] and row["method"]:
                sections["Microbiological Profile"].append((row["parameter"], row["spec"], row["result"], row["method"]))

    # Add each section to spec_data
    for section_name, rows in sections.items():
        filtered_rows = [r for r in rows if r]
        if filtered_rows:
            spec_data.append([Paragraph(f"<b>{section_name}</b>", style_for_sections), "", "", ""])
            heading_indices.append(len(spec_data) - 1)
            for param_tuple in filtered_rows:
                row_cells = [Paragraph(str(cell), normal_style) for cell in param_tuple]
                spec_data.append(row_cells)

    # Add remarks
    remarks_text = ("Since the product is derived from natural origin, there is likely to be minor color "
                    "variation because of the geographical and seasonal variations of the raw material")
    end_text = "REMARKS: COMPLIES WITH IN HOUSE SPECIFICATIONS"
    spec_data.append([Paragraph(remarks_text, normal_style), "", "", ""])
    last_remarks_row = len(spec_data) - 1
    spec_data.append([Paragraph(end_text, ParagraphStyle('bold_center', parent=styles['Normal'],
                                                        fontName='Helvetica-Bold', alignment=1)),
                      "", "", ""])
    final_remark_row = len(spec_data) - 1

    total_width = 500
    col_widths = [total_width * 0.23, total_width * 0.39, total_width * 0.18, total_width * 0.20]
    spec_table = Table(spec_data, colWidths=col_widths)
    spec_table_style = [
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),
    ]
    for heading_row in heading_indices:
        spec_table_style.append(('SPAN', (0, heading_row), (-1, heading_row)))
    spec_table_style.append(('SPAN', (0, last_remarks_row), (-1, last_remarks_row)))
    spec_table_style.append(('SPAN', (0, final_remark_row), (-1, final_remark_row)))
    spec_table.setStyle(TableStyle(spec_table_style))

    elements.append(spec_table)
    elements.append(Spacer(1, 2))

    # Declaration
    elements.append(Paragraph("Declaration", title_style1))
    declaration_data = [
        [
            "GMO Status:",
            Paragraph("Free from GMO", normal_style),
            "",
            "Allergen statement:",
            Paragraph(f"{data.get('allergen_statement','Free from allergen')}", normal_style)
        ],
        [
            "Irradiation status:",
            Paragraph("Non – Irradiated", normal_style),
            "",
            "Storage condition:",
            Paragraph("At room temperature", normal_style)
        ],
        [
            "Prepared by",
            Paragraph("Executive – QC", normal_style),
            "",
            "Approved by",
            Paragraph("Head-QC/QA", normal_style)
        ]
    ]
    declaration_table = Table(declaration_data, colWidths=[80, 150, 75, 100, 95])
    declaration_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('ALIGN', (0, 0), (1, -1), 'LEFT'),
        ('ALIGN', (3, 0), (4, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),
        ('LEFTPADDING', (0, 0), (-1, -1), 0),
        ('RIGHTPADDING', (0, 0), (-1, -1), 0),
        ('TOPPADDING', (0, 0), (-1, -1), 0),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ('SPAN', (2, 0), (2, 2)),
    ]))
    elements.append(declaration_table)
    elements.append(Spacer(1, 3))

    # Build normal multi-page PDF
    doc.build(elements, onFirstPage=header_footer, onLaterPages=header_footer)
    buffer.seek(0)

    # ----------------------------------------------------------------
    # Force single page if there's more than 1 page
    # ----------------------------------------------------------------
    pdf_in = fitz.open(stream=buffer.getvalue(), filetype="pdf")
    if len(pdf_in) > 1:
        # Combine all pages into a single tall image, then scale to A4
        images = []
        for page_index in range(len(pdf_in)):
            # Increase dpi for better resolution if desired
            pix = pdf_in[page_index].get_pixmap(dpi=144)
            # Convert PyMuPDF Pixmap to PIL Image
            img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
            images.append(img)

        # Stack vertically
        total_height = sum(im.height for im in images)
        max_width = max(im.width for im in images)
        combined_img = Image.new("RGBA", (max_width, total_height), (255, 255, 255, 0))
        current_y = 0
        for im in images:
            combined_img.paste(im, (0, current_y))
            current_y += im.height

        # Now scale to fit standard A4 (595 x 842 points)
        target_w, target_h = 595, 842
        scale_factor = min(target_w / combined_img.width, target_h / combined_img.height)
        new_w = int(combined_img.width * scale_factor)
        new_h = int(combined_img.height * scale_factor)
        combined_img = combined_img.resize((new_w, new_h), Image.LANCZOS)

        # Create a new single-page PDF
        new_pdf = fitz.open()
        page = new_pdf.new_page(width=595, height=842)
        # Convert the combined_img back into bytes
        img_bytes = io.BytesIO()
        combined_img.save(img_bytes, format="PNG")
        img_bytes.seek(0)
        rect = fitz.Rect(0, 0, new_w, new_h)
        page.insert_image(rect, stream=img_bytes.getvalue())
        final_buffer = io.BytesIO(new_pdf.tobytes())
        return final_buffer

    # If only 1 page, return as-is
    return buffer


##############################################
# STREAMLIT APP
##############################################

# Initialize session_state lists for dynamic rows
if "physical_rows" not in st.session_state:
    st.session_state["physical_rows"] = []
if "others_rows" not in st.session_state:
    st.session_state["others_rows"] = []
if "assays_rows" not in st.session_state:
    st.session_state["assays_rows"] = []
if "pesticides_rows" not in st.session_state:
    st.session_state["pesticides_rows"] = []
if "residual_solvent_rows" not in st.session_state:
    st.session_state["residual_solvent_rows"] = []
if "microbio_rows" not in st.session_state:
    st.session_state["microbio_rows"] = []

col1, col2 = st.columns(2)

##############################################
# BUTTONS TO ADD DYNAMIC ROWS (OUTSIDE FORM)
##############################################

st.markdown("### Add More Rows (outside the main form)")

# Physical
if st.button("Add new Physical Row"):
    st.session_state["physical_rows"].append({"parameter": "", "spec": "", "result": "", "method": ""})

# Others
if st.button("Add new Others Row"):
    st.session_state["others_rows"].append({"parameter": "", "spec": "", "result": "", "method": ""})

# Assays
if st.button("Add new Assays Row"):
    st.session_state["assays_rows"].append({"parameter": "", "spec": "", "result": "", "method": ""})

# Pesticides
if st.button("Add new Pesticides Row"):
    st.session_state["pesticides_rows"].append({"parameter": "", "spec": "", "result": "", "method": ""})

# Residual Solvent
if st.button("Add new Residual Solvent Row"):
    st.session_state["residual_solvent_rows"].append({"parameter": "", "spec": "", "result": "", "method": ""})

# Microbiological
if st.button("Add new Microbiological Row"):
    st.session_state["microbio_rows"].append({"parameter": "", "spec": "", "result": "", "method": ""})

##############################################
# MAIN FORM
##############################################
with col1.form("coa_form"):
    st.title("Tru Herb COA PDF Generator")
    st.header("Product Information")

    # 2-col for Product Info
    row1_col1, row1_col2 = st.columns(2)
    product_name = row1_col1.text_input("Product Name", placeholder="X")
    product_code = row1_col2.text_input("Product Code", placeholder="X")

    row2_col1, row2_col2 = st.columns(2)
    batch_no = row2_col1.text_input("Batch No.", placeholder="X")
    manufacturing_date = row2_col2.text_input("Date of Manufacturing", placeholder="X")

    row3_col1, row3_col2 = st.columns(2)
    reanalysis_date = row3_col1.text_input("Date of Reanalysis", placeholder="X")
    botanical_name = row3_col2.text_input("Botanical Name", placeholder="X")

    row4_col1, row4_col2 = st.columns(2)
    extraction_ratio = row4_col1.text_input("Extraction Ratio", placeholder="X")
    solvent = row4_col2.text_input("Extraction Solvents", placeholder="X")

    row5_col1, row5_col2 = st.columns(2)
    plant_part = row5_col1.text_input("Plant Parts", placeholder="X")
    cas_no = row5_col2.text_input("CAS No.", placeholder="X")

    row6_col1, row6_col2 = st.columns(2)
    chemical_name = row6_col1.text_input("Chemical Name", placeholder="X")
    quantity = row6_col2.text_input("Quantity", placeholder="X")

    origin = st.text_input("Country of Origin", value="India")

    st.header("Specifications")
    st.subheader("Physical")

    # Existing Physical fields
    phys_1col1, phys_1col2, phys_1col3 = st.columns(3)
    description_spec = phys_1col1.text_input("Spec for Description", value="X with Characteristic taste and odour")
    description_result = phys_1col2.text_input("Result for Description", value="Compiles")
    description_method = phys_1col3.text_input("Method for Description", value="Physical")

    phys_2col1, phys_2col2, phys_2col3 = st.columns(3)
    identification_spec = phys_2col1.text_input("Spec for Identification", value="To comply by TLC")
    identification_result = phys_2col2.text_input("Result for Identification", value="Compiles")
    identification_method = phys_2col3.text_input("Method for Identification", value="TLC")

    phys_3col1, phys_3col2, phys_3col3 = st.columns(3)
    loss_on_drying_spec = phys_3col1.text_input("Spec for Loss on Drying", value="Not more than X")
    loss_on_drying_result = phys_3col2.text_input("Result for Loss on Drying", placeholder="X")
    loss_on_drying_method = phys_3col3.text_input("Method for Loss on Drying", value="USP<731>")

    phys_4col1, phys_4col2, phys_4col3 = st.columns(3)
    moisture_spec = phys_4col1.text_input("Spec for Moisture", value="Not more than X")
    moisture_result = phys_4col2.text_input("Result for Moisture", placeholder="X")
    moisture_method = phys_4col3.text_input("Method for Moisture", value="USP<921>")

    phys_5col1, phys_5col2, phys_5col3 = st.columns(3)
    particle_size_spec = phys_5col1.text_input("Spec for Particle Size", placeholder="X")
    particle_size_result = phys_5col2.text_input("Result for Particle Size", placeholder="X")
    particle_size_method = phys_5col3.text_input("Method for Particle Size", value="USP<786>")

    phys_6col1, phys_6col2, phys_6col3 = st.columns(3)
    ash_contents_spec = phys_6col1.text_input("Spec for Ash Contents", value="Not more than X")
    ash_contents_result = phys_6col2.text_input("Result for Ash Contents", placeholder="X")
    ash_contents_method = phys_6col3.text_input("Method for Ash Contents", value="USP<561>")

    phys_7col1, phys_7col2, phys_7col3 = st.columns(3)
    residue_on_ignition_spec = phys_7col1.text_input("Spec for Residue on Ignition", value="Not more than X")
    residue_on_ignition_result = phys_7col2.text_input("Result for Residue on Ignition", placeholder="X")
    residue_on_ignition_method = phys_7col3.text_input("Method for Residue on Ignition", value="USP<281>")

    phys_8col1, phys_8col2, phys_8col3 = st.columns(3)
    bulk_density_spec = phys_8col1.text_input("Spec for Bulk Density", value="Between 0.3g/ml to 0.6g/ml")
    bulk_density_result = phys_8col2.text_input("Result for Bulk Density", placeholder="X")
    bulk_density_method = phys_8col3.text_input("Method for Bulk Density", value="USP<616>")

    phys_9col1, phys_9col2, phys_9col3 = st.columns(3)
    tapped_density_spec = phys_9col1.text_input("Spec for Tapped Density", value="Between 0.4g/ml to 0.8g/ml")
    tapped_density_result = phys_9col2.text_input("Result for Tapped Density", placeholder="X")
    tapped_density_method = phys_9col3.text_input("Method for Tapped Density", value="USP<616>")

    phys_10col1, phys_10col2, phys_10col3 = st.columns(3)
    solubility_spec = phys_10col1.text_input("Spec for Solubility", placeholder="X")
    solubility_result = phys_10col2.text_input("Result for Solubility", placeholder="X")
    solubility_method = phys_10col3.text_input("Method for Solubility", value="USP<1236>")

    phys_11col1, phys_11col2, phys_11col3 = st.columns(3)
    ph_spec = phys_11col1.text_input("Spec for pH", placeholder="X")
    ph_result = phys_11col2.text_input("Result for pH", placeholder="X")
    ph_method = phys_11col3.text_input("Method for pH", value="USP<791>")

    phys_12col1, phys_12col2, phys_12col3 = st.columns(3)
    chlorides_nacl_spec = phys_12col1.text_input("Spec for Chlorides of NaCl", placeholder="X")
    chlorides_nacl_result = phys_12col2.text_input("Result for Chlorides of NaCl", placeholder="X")
    chlorides_nacl_method = phys_12col3.text_input("Method for Chlorides of NaCl", value="USP<221>")

    phys_13col1, phys_13col2, phys_13col3 = st.columns(3)
    sulphates_spec = phys_13col1.text_input("Spec for Sulphates", placeholder="X")
    sulphates_result = phys_13col2.text_input("Result for Sulphates", placeholder="X")
    sulphates_method = phys_13col3.text_input("Method for Sulphates", value="USP<221>")

    phys_14col1, phys_14col2, phys_14col3 = st.columns(3)
    fats_spec = phys_14col1.text_input("Spec for Fats", placeholder="X")
    fats_result = phys_14col2.text_input("Result for Fats", placeholder="X")
    fats_method = phys_14col3.text_input("Method for Fats", value="USP<731>")

    phys_15col1, phys_15col2, phys_15col3 = st.columns(3)
    protein_spec = phys_15col1.text_input("Spec for Protein", placeholder="X")
    protein_result = phys_15col2.text_input("Result for Protein", placeholder="X")
    protein_method = phys_15col3.text_input("Method for Protein", value="Kjeldahl")

    phys_16col1, phys_16col2, phys_16col3 = st.columns(3)
    total_ig_g_spec = phys_16col1.text_input("Spec for Total IgG", placeholder="X")
    total_ig_g_result = phys_16col2.text_input("Result for Total IgG", placeholder="X")
    total_ig_g_method = phys_16col3.text_input("Method for Total IgG", value="HPLC")

    phys_17col1, phys_17col2, phys_17col3 = st.columns(3)
    sodium_spec = phys_17col1.text_input("Spec for Sodium", placeholder="X")
    sodium_result = phys_17col2.text_input("Result for Sodium", placeholder="X")
    sodium_method = phys_17col3.text_input("Method for Sodium", value="ICP-MS")

    phys_18col1, phys_18col2, phys_18col3 = st.columns(3)
    gluten_spec = phys_18col1.text_input("Spec for Gluten", value="NMT X")
    gluten_result = phys_18col2.text_input("Result for Gluten", placeholder="X")
    gluten_method = phys_18col3.text_input("Method for Gluten", value="ELISA")

    # Additional Physical Rows (display only)
    st.markdown("#### Additional Physical Rows")
    for i, row in enumerate(st.session_state["physical_rows"]):
        c1, c2, c3, c4 = st.columns(4)
        row["parameter"] = c1.text_input(f"Physical Param {i+1}", value=row["parameter"], key=f"phys_param_{i}")
        row["spec"] = c2.text_input(f"Physical Spec {i+1}", value=row["spec"], key=f"phys_spec_{i}")
        row["result"] = c3.text_input(f"Physical Result {i+1}", value=row["result"], key=f"phys_result_{i}")
        row["method"] = c4.text_input(f"Physical Method {i+1}", value=row["method"], key=f"phys_method_{i}")

    # -----------------------------
    # Others
    # -----------------------------
    st.subheader("Others")
    others_1col1, others_1col2, others_1col3 = st.columns(3)
    lead_spec = others_1col1.text_input("Spec for Lead", value="Not more than X ppm")
    lead_result = others_1col2.text_input("Result for Lead", placeholder="X")
    lead_method = others_1col3.text_input("Method for Lead", value="ICP-MS")

    others_2col1, others_2col2, others_2col3 = st.columns(3)
    cadmium_spec = others_2col1.text_input("Spec for Cadmium", value="Not more than X ppm")
    cadmium_result = others_2col2.text_input("Result for Cadmium", placeholder="X")
    cadmium_method = others_2col3.text_input("Method for Cadmium", value="ICP-MS")

    others_3col1, others_3col2, others_3col3 = st.columns(3)
    arsenic_spec = others_3col1.text_input("Spec for Arsenic", value="Not more than X ppm")
    arsenic_result = others_3col2.text_input("Result for Arsenic", placeholder="X")
    arsenic_method = others_3col3.text_input("Method for Arsenic", value="ICP-MS")

    others_4col1, others_4col2, others_4col3 = st.columns(3)
    mercury_spec = others_4col1.text_input("Spec for Mercury", value="Not more than X ppm")
    mercury_result = others_4col2.text_input("Result for Mercury", placeholder="X")
    mercury_method = others_4col3.text_input("Method for Mercury", value="ICP-MS")

    st.markdown("#### Additional Others Rows")
    for i, row in enumerate(st.session_state["others_rows"]):
        c1, c2, c3, c4 = st.columns(4)
        row["parameter"] = c1.text_input(f"Others Param {i+1}", value=row["parameter"], key=f"others_param_{i}")
        row["spec"] = c2.text_input(f"Others Spec {i+1}", value=row["spec"], key=f"others_spec_{i}")
        row["result"] = c3.text_input(f"Others Result {i+1}", value=row["result"], key=f"others_result_{i}")
        row["method"] = c4.text_input(f"Others Method {i+1}", value=row["method"], key=f"others_method_{i}")

    # -----------------------------
    # Assays
    # -----------------------------
    st.subheader("Assays")
    assays_col1, assays_col2, assays_col3 = st.columns(3)
    assays_spec = assays_col1.text_input("Specification for Assays", placeholder="X")
    assays_result = assays_col2.text_input("Result for Assays", placeholder="X")
    assays_method = assays_col3.text_input("Method for Assays", placeholder="X")

    st.markdown("#### Additional Assays Rows")
    for i, row in enumerate(st.session_state["assays_rows"]):
        c1, c2, c3, c4 = st.columns(4)
        row["parameter"] = c1.text_input(f"Assays Param {i+1}", value=row["parameter"], key=f"assays_param_{i}")
        row["spec"] = c2.text_input(f"Assays Spec {i+1}", value=row["spec"], key=f"assays_spec_{i}")
        row["result"] = c3.text_input(f"Assays Result {i+1}", value=row["result"], key=f"assays_result_{i}")
        row["method"] = c4.text_input(f"Assays Method {i+1}", value=row["method"], key=f"assays_method_{i}")

    # -----------------------------
    # Pesticides
    # -----------------------------
    st.subheader("Pesticides")
    pest_col1, pest_col2, pest_col3 = st.columns(3)
    pesticide_spec = pest_col1.text_input("Specification for Pesticide", value="Meet USP<561>")
    pesticide_result = pest_col2.text_input("Result for Pesticide", value="Compiles")
    pesticide_method = pest_col3.text_input("Method for Pesticide", value="USP<561>")

    st.markdown("#### Additional Pesticides Rows")
    for i, row in enumerate(st.session_state["pesticides_rows"]):
        c1, c2, c3, c4 = st.columns(4)
        row["parameter"] = c1.text_input(f"Pesticides Param {i+1}", value=row["parameter"], key=f"pest_param_{i}")
        row["spec"] = c2.text_input(f"Pesticides Spec {i+1}", value=row["spec"], key=f"pest_spec_{i}")
        row["result"] = c3.text_input(f"Pesticides Result {i+1}", value=row["result"], key=f"pest_result_{i}")
        row["method"] = c4.text_input(f"Pesticides Method {i+1}", value=row["method"], key=f"pest_method_{i}")

    # -----------------------------
    # Residual Solvent
    # -----------------------------
    st.subheader("Residual Solvent")
    rs_col1, rs_col2, rs_col3 = st.columns(3)
    residual_solvent_spec = rs_col1.text_input("Specification for Residual Solvent", placeholder="X")
    residual_solvent_result = rs_col2.text_input("Result for Residual Solvent", value="Compiles")
    residual_solvent_method = rs_col3.text_input("Method for Residual Solvent", placeholder="X")

    st.markdown("#### Additional Residual Solvent Rows")
    for i, row in enumerate(st.session_state["residual_solvent_rows"]):
        c1, c2, c3, c4 = st.columns(4)
        row["parameter"] = c1.text_input(f"Residual Param {i+1}", value=row["parameter"], key=f"resid_param_{i}")
        row["spec"] = c2.text_input(f"Residual Spec {i+1}", value=row["spec"], key=f"resid_spec_{i}")
        row["result"] = c3.text_input(f"Residual Result {i+1}", value=row["result"], key=f"resid_result_{i}")
        row["method"] = c4.text_input(f"Residual Method {i+1}", value=row["method"], key=f"resid_method_{i}")

    # -----------------------------
    # Microbiological Profile
    # -----------------------------
    st.subheader("Microbiological Profile")
    micro_1col1, micro_1col2, micro_1col3 = st.columns(3)
    total_plate_count_spec = micro_1col1.text_input("Spec for Total Plate Count", value="Not more than X cfu/g")
    total_plate_count_result = micro_1col2.text_input("Result for Total Plate Count", value="X cfu/g")
    total_plate_count_method = micro_1col3.text_input("Method for Total Plate Count", value="USP<61>")

    micro_2col1, micro_2col2, micro_2col3 = st.columns(3)
    yeasts_mould_spec = micro_2col1.text_input("Spec for Yeasts & Mould Count", value="Not more than X cfu/g")
    yeasts_mould_result = micro_2col2.text_input("Result for Yeasts & Mould Count", value="X cfu/g")
    yeasts_mould_method = micro_2col3.text_input("Method for Yeasts & Mould Count", value="USP<61>")

    micro_3col1, micro_3col2, micro_3col3 = st.columns(3)
    salmonella_spec = micro_3col1.text_input("Spec for Salmonella", value="Absent/25g")
    salmonella_result = micro_3col2.text_input("Result for Salmonella", value="Absent")
    salmonella_method = micro_3col3.text_input("Method for Salmonella", value="USP<62>")

    micro_4col1, micro_4col2, micro_4col3 = st.columns(3)
    e_coli_spec = micro_4col1.text_input("Spec for Escherichia coli", value="Absent/10g")
    e_coli_result = micro_4col2.text_input("Result for Escherichia coli", value="Absent")
    e_coli_method = micro_4col3.text_input("Method for Escherichia coli", value="USP<62>")

    micro_5col1, micro_5col2, micro_5col3 = st.columns(3)
    coliforms_spec = micro_5col1.text_input("Spec for Coliforms", value="NMT X cfu/g")
    coliforms_result = micro_5col2.text_input("Result for Coliforms", placeholder="X")
    coliforms_method = micro_5col3.text_input("Method for Coliforms", value="USP<62>")

    st.markdown("#### Additional Microbiological Rows")
    for i, row in enumerate(st.session_state["microbio_rows"]):
        c1, c2, c3, c4 = st.columns(4)
        row["parameter"] = c1.text_input(f"Microbio Param {i+1}", value=row["parameter"], key=f"microbio_param_{i}")
        row["spec"] = c2.text_input(f"Microbio Spec {i+1}", value=row["spec"], key=f"microbio_spec_{i}")
        row["result"] = c3.text_input(f"Microbio Result {i+1}", value=row["result"], key=f"microbio_result_{i}")
        row["method"] = c4.text_input(f"Microbio Method {i+1}", value=row["method"], key=f"microbio_method_{i}")

    st.subheader("Declaration - Allergen Statement")
    allergen_statement = st.selectbox("Allergen Statement", options=["Free from allergen", "Contains Allergen"])

    # Two submit buttons inside the form
    preview_button = st.form_submit_button("Preview")
    download_button = st.form_submit_button("Compile and Generate PDF")


######################
# Handle Preview
######################
if preview_button:
    data = {
        "product_name": product_name,
        "botanical_name": botanical_name,
        "chemical_name": chemical_name,
        "cas_no": cas_no,
        "product_code": product_code,
        "batch_no": batch_no,
        "manufacturing_date": manufacturing_date,
        "reanalysis_date": reanalysis_date,
        "quantity": quantity,
        "origin": origin,
        "plant_part": plant_part,
        "extraction_ratio": extraction_ratio,
        "solvent": solvent,

        "description_spec": description_spec,
        "description_result": description_result,
        "description_method": description_method,
        "identification_spec": identification_spec,
        "identification_result": identification_result,
        "identification_method": identification_method,
        "loss_on_drying_spec": loss_on_drying_spec,
        "loss_on_drying_result": loss_on_drying_result,
        "loss_on_drying_method": loss_on_drying_method,
        "moisture_spec": moisture_spec,
        "moisture_result": moisture_result,
        "moisture_method": moisture_method,
        "particle_size_spec": particle_size_spec,
        "particle_size_result": particle_size_result,
        "particle_size_method": particle_size_method,
        "ash_contents_spec": ash_contents_spec,
        "ash_contents_result": ash_contents_result,
        "ash_contents_method": ash_contents_method,
        "residue_on_ignition_spec": residue_on_ignition_spec,
        "residue_on_ignition_result": residue_on_ignition_result,
        "residue_on_ignition_method": residue_on_ignition_method,
        "bulk_density_spec": bulk_density_spec,
        "bulk_density_result": bulk_density_result,
        "bulk_density_method": bulk_density_method,
        "tapped_density_spec": tapped_density_spec,
        "tapped_density_result": tapped_density_result,
        "tapped_density_method": tapped_density_method,
        "solubility_spec": solubility_spec,
        "solubility_result": solubility_result,
        "solubility_method": solubility_method,
        "ph_spec": ph_spec,
        "ph_result": ph_result,
        "ph_method": ph_method,
        "chlorides_nacl_spec": chlorides_nacl_spec,
        "chlorides_nacl_result": chlorides_nacl_result,
        "chlorides_nacl_method": chlorides_nacl_method,
        "sulphates_spec": sulphates_spec,
        "sulphates_result": sulphates_result,
        "sulphates_method": sulphates_method,
        "fats_spec": fats_spec,
        "fats_result": fats_result,
        "fats_method": fats_method,
        "protein_spec": protein_spec,
        "protein_result": protein_result,
        "protein_method": protein_method,
        "total_ig_g_spec": total_ig_g_spec,
        "total_ig_g_result": total_ig_g_result,
        "total_ig_g_method": total_ig_g_method,
        "sodium_spec": sodium_spec,
        "sodium_result": sodium_result,
        "sodium_method": sodium_method,
        "gluten_spec": gluten_spec,
        "gluten_result": gluten_result,
        "gluten_method": gluten_method,
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
        "pesticide_spec": pesticide_spec,
        "pesticide_result": pesticide_result,
        "pesticide_method": pesticide_method,
        "residual_solvent_spec": residual_solvent_spec,
        "residual_solvent_result": residual_solvent_result,
        "residual_solvent_method": residual_solvent_method,
        "total_plate_count_spec": total_plate_count_spec,
        "total_plate_count_result": total_plate_count_result,
        "total_plate_count_method": total_plate_count_method,
        "yeasts_mould_spec": yeasts_mould_spec,
        "yeasts_mould_result": yeasts_mould_result,
        "yeasts_mould_method": yeasts_mould_method,
        "salmonella_spec": salmonella_spec,
        "salmonella_result": salmonella_result,
        "salmonella_method": salmonella_method,
        "e_coli_spec": e_coli_spec,
        "e_coli_result": e_coli_result,
        "e_coli_method": e_coli_method,
        "coliforms_spec": coliforms_spec,
        "coliforms_result": coliforms_result,
        "coliforms_method": coliforms_method,
        "allergen_statement": allergen_statement,

        # add dynamic rows
        "physical_dynamic_rows": st.session_state["physical_rows"],
        "others_dynamic_rows": st.session_state["others_rows"],
        "assays_dynamic_rows": st.session_state["assays_rows"],
        "pesticides_dynamic_rows": st.session_state["pesticides_rows"],
        "residual_solvent_dynamic_rows": st.session_state["residual_solvent_rows"],
        "microbio_dynamic_rows": st.session_state["microbio_rows"],
    }

    pdf_buffer = generate_pdf(data)
    if pdf_buffer:
        doc_preview = fitz.open(stream=pdf_buffer, filetype="pdf")
        with col2:
            for page in doc_preview:
                pix = page.get_pixmap()
                st.image(pix.tobytes(), caption=f"Page {page.number + 1}", use_container_width=True)
        with col1:
            st.success("Preview generated successfully!")

######################
# Handle Download
######################
if download_button:
    data = {
        "product_name": product_name,
        "botanical_name": botanical_name,
        "chemical_name": chemical_name,
        "cas_no": cas_no,
        "product_code": product_code,
        "batch_no": batch_no,
        "manufacturing_date": manufacturing_date,
        "reanalysis_date": reanalysis_date,
        "quantity": quantity,
        "origin": origin,
        "plant_part": plant_part,
        "extraction_ratio": extraction_ratio,
        "solvent": solvent,
        "description_spec": description_spec,
        "description_result": description_result,
        "description_method": description_method,
        "identification_spec": identification_spec,
        "identification_result": identification_result,
        "identification_method": identification_method,
        "loss_on_drying_spec": loss_on_drying_spec,
        "loss_on_drying_result": loss_on_drying_result,
        "loss_on_drying_method": loss_on_drying_method,
        "moisture_spec": moisture_spec,
        "moisture_result": moisture_result,
        "moisture_method": moisture_method,
        "particle_size_spec": particle_size_spec,
        "particle_size_result": particle_size_result,
        "particle_size_method": particle_size_method,
        "ash_contents_spec": ash_contents_spec,
        "ash_contents_result": ash_contents_result,
        "ash_contents_method": ash_contents_method,
        "residue_on_ignition_spec": residue_on_ignition_spec,
        "residue_on_ignition_result": residue_on_ignition_result,
        "residue_on_ignition_method": residue_on_ignition_method,
        "bulk_density_spec": bulk_density_spec,
        "bulk_density_result": bulk_density_result,
        "bulk_density_method": bulk_density_method,
        "tapped_density_spec": tapped_density_spec,
        "tapped_density_result": tapped_density_result,
        "tapped_density_method": tapped_density_method,
        "solubility_spec": solubility_spec,
        "solubility_result": solubility_result,
        "solubility_method": solubility_method,
        "ph_spec": ph_spec,
        "ph_result": ph_result,
        "ph_method": ph_method,
        "chlorides_nacl_spec": chlorides_nacl_spec,
        "chlorides_nacl_result": chlorides_nacl_result,
        "chlorides_nacl_method": chlorides_nacl_method,
        "sulphates_spec": sulphates_spec,
        "sulphates_result": sulphates_result,
        "sulphates_method": sulphates_method,
        "fats_spec": fats_spec,
        "fats_result": fats_result,
        "fats_method": fats_method,
        "protein_spec": protein_spec,
        "protein_result": protein_result,
        "protein_method": protein_method,
        "total_ig_g_spec": total_ig_g_spec,
        "total_ig_g_result": total_ig_g_result,
        "total_ig_g_method": total_ig_g_method,
        "sodium_spec": sodium_spec,
        "sodium_result": sodium_result,
        "sodium_method": sodium_method,
        "gluten_spec": gluten_spec,
        "gluten_result": gluten_result,
        "gluten_method": gluten_method,
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
        "pesticide_spec": pesticide_spec,
        "pesticide_result": pesticide_result,
        "pesticide_method": pesticide_method,
        "residual_solvent_spec": residual_solvent_spec,
        "residual_solvent_result": residual_solvent_result,
        "residual_solvent_method": residual_solvent_method,
        "total_plate_count_spec": total_plate_count_spec,
        "total_plate_count_result": total_plate_count_result,
        "total_plate_count_method": total_plate_count_method,
        "yeasts_mould_spec": yeasts_mould_spec,
        "yeasts_mould_result": yeasts_mould_result,
        "yeasts_mould_method": yeasts_mould_method,
        "salmonella_spec": salmonella_spec,
        "salmonella_result": salmonella_result,
        "salmonella_method": salmonella_method,
        "e_coli_spec": e_coli_spec,
        "e_coli_result": e_coli_result,
        "e_coli_method": e_coli_method,
        "coliforms_spec": coliforms_spec,
        "coliforms_result": coliforms_result,
        "coliforms_method": coliforms_method,
        "allergen_statement": allergen_statement,

        "physical_dynamic_rows": st.session_state["physical_rows"],
        "others_dynamic_rows": st.session_state["others_rows"],
        "assays_dynamic_rows": st.session_state["assays_rows"],
        "pesticides_dynamic_rows": st.session_state["pesticides_rows"],
        "residual_solvent_dynamic_rows": st.session_state["residual_solvent_rows"],
        "microbio_dynamic_rows": st.session_state["microbio_rows"],
    }

    pdf_buffer = generate_pdf(data)
    if pdf_buffer:
        st.download_button(
            label="Download COA PDF",
            data=pdf_buffer,
            file_name=(product_name or "COA") + ".pdf",
            mime="application/pdf"
        )
        st.success("COA PDF generated and ready for download!")
