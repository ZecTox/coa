import os
from typing import Container
import streamlit as st

# ReportLab imports
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    KeepInFrame
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

import io
import fitz  # PyMuPDF
import configparser

# -----------------------------
# INITIALIZE SESSION STATE
# -----------------------------
if "Physical_rows" not in st.session_state:
    st.session_state["Physical_rows"] = []
if "Others_rows" not in st.session_state:
    st.session_state["Others_rows"] = []
if "Assays_rows" not in st.session_state:
    st.session_state["Assays_rows"] = []
if "Pesticides_rows" not in st.session_state:
    st.session_state["Pesticides_rows"] = []
if "ResidualSolvent_rows" not in st.session_state:
    st.session_state["ResidualSolvent_rows"] = []
if "MicrobiologicalProfile_rows" not in st.session_state:
    st.session_state["MicrobiologicalProfile_rows"] = []

# Initialize the configparser
config = configparser.ConfigParser()
config.read('.streamlit/config.toml')
theme = config.get('settings', 'theme', fallback='default')

st.set_page_config(page_title="Tru Herb COA PDF Generator", layout="wide")


def header_footer(canvas, doc):
    canvas.saveState()
    logo_path = os.path.join(os.getcwd(), "images", "tru_herb_logo.png")
    footer_path = os.path.join(os.getcwd(), "images", "footer.png")

    if os.path.exists(logo_path):
        canvas.drawImage(logo_path, x=250, y=A4[1] - 55, width=100, height=50)
    if os.path.exists(footer_path):
        canvas.drawImage(footer_path, x=50, y=5, width=500, height=80)
    canvas.restoreState()


def generate_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=50,
        bottomMargin=80
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle('title_style', fontSize=12, spaceAfter=1,
                                 alignment=1, fontName='Times-Bold')
    title_style1 = ParagraphStyle('title_style1', fontSize=10, spaceAfter=0,
                                  alignment=1, fontName='Times-Bold')
    normal_style = styles['BodyText']
    normal_style.fontName = 'Times-Roman'
    normal_style.alignment = 0
    style_for_sections = styles["Normal"]

    elements = []
    elements.append(Spacer(1, 3))
    elements.append(Paragraph("CERTIFICATE OF ANALYSIS", title_style))
    elements.append(Paragraph(data.get('product_name', '').upper(), title_style)) #Added the Product Name here in CapsLock(All in Uppercase)
    elements.append(Spacer(1, 3))

    # ----------------------------------------------------------------
    # Build Product Info table, skipping truly empty fields
    # ----------------------------------------------------------------
    product_info = []

    def maybe_add_product_row(label, value, italic=False):
        text_str = value.strip() if value else ""
        if text_str:
            if italic:
                text_str = f"<i>{text_str}</i>"
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
        product_table = Table(product_info, colWidths=[140, 360])
        product_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),
        ]))
        elements.append(product_table)
        elements.append(Spacer(1, 0))

    # ----------------------------------------------------------------
    # SPECIFICATIONS TABLE
    # ----------------------------------------------------------------
    spec_headers = ["Parameter", "Specification", "Result", "Method"]
    spec_data = [spec_headers]
    heading_rows = []
    current_row_index = 1

    def combine_section(section_key, base_rows):
        extra_rows = data.get(section_key, [])
        return [row for row in base_rows if row] + [r for r in extra_rows if r]

    physical_base = [
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
        ("Residue on Ignition", data['residue_on_ignition_spec'], data['residue_on_ignition_result'],
         data['residue_on_ignition_method'])
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
    ]
    others_base = [
        ("Lead", data['lead_spec'], data['lead_result'], data['lead_method'])
            if data['lead_spec'] and data['lead_result'] and data['lead_method'] else None,
        ("Cadmium", data['cadmium_spec'], data['cadmium_result'], data['cadmium_method'])
            if data['cadmium_spec'] and data['cadmium_result'] and data['cadmium_method'] else None,
        ("Arsenic", data['arsenic_spec'], data['arsenic_result'], data['arsenic_method'])
            if data['arsenic_spec'] and data['arsenic_result'] and data['arsenic_method'] else None,
        ("Mercury", data['mercury_spec'], data['mercury_result'], data['mercury_method'])
            if data['mercury_spec'] and data['mercury_result'] and data['mercury_method'] else None,
    ]
    assays_base = [
        ("Assays", data['assays_spec'], data['assays_result'], data['assays_method'])
            if data['assays_spec'] and data['assays_result'] and data['assays_method'] else None,
    ]
    pesticides_base = [
        ("Pesticide", data['pesticide_spec'], data['pesticide_result'], data['pesticide_method'])
            if data['pesticide_spec'] and data['pesticide_result'] and data['pesticide_method'] else None,
    ]
    residual_solvent_base = [
        ("Residual Solvent", data['residual_solvent_spec'], data['residual_solvent_result'], data['residual_solvent_method'])
            if data['residual_solvent_spec'] and data['residual_solvent_result'] and data['residual_solvent_method'] else None,
    ]
    microbio_base = [
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

    sections = {
        "Physical": combine_section("physical_extra_rows", physical_base),
        "Others": combine_section("others_extra_rows", others_base),
        "Assays": combine_section("assays_extra_rows", assays_base),
        "Pesticides": combine_section("pesticides_extra_rows", pesticides_base),
        "Residual Solvent": combine_section("residual_solvent_extra_rows", residual_solvent_base),
        "Microbiological Profile": combine_section("microbio_extra_rows", microbio_base),
    }

    for section_name, rows in sections.items():
        if rows:
            spec_data.append([Paragraph(f"<b>{section_name}</b>", style_for_sections), "", "", ""])
            heading_rows.append(len(spec_data) - 1)
            for param_tuple in rows:
                row_cells = [Paragraph(str(cell), normal_style) for cell in param_tuple]
                spec_data.append(row_cells)

    # Remarks
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
    col_widths = [total_width * 0.23,
                  total_width * 0.39,
                  total_width * 0.18,
                  total_width * 0.20]
    spec_table = Table(spec_data, colWidths=col_widths)

    spec_table_style = [
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, -1), 'Times-Roman'),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('WORDWRAP', (0, 0), (-1, -1), 'LTR'),
    ]
    for heading_row in heading_rows:
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

    # Force single page
    kiframe = KeepInFrame(
        maxWidth=A4[0] - doc.leftMargin - doc.rightMargin,
        maxHeight=A4[1] - doc.topMargin - doc.bottomMargin,
        content=elements,
        mode='shrink'
    )
    elements = [kiframe]

    doc.build(elements, onFirstPage=header_footer, onLaterPages=header_footer)
    buffer.seek(0)
    return buffer

# ----------------------------------------------------------------------------
# HELPER to initialize a session_state key if not present
# ----------------------------------------------------------------------------
def init_ss(key, default):
    if key not in st.session_state:
        st.session_state[key] = default

# ----------------------------------------------------------------------------
# STREAMLIT UI
# ----------------------------------------------------------------------------

col1, col2 = st.columns(2)
with col1:
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

    # ---------- SPECIFICATIONS -----------
    st.header("Specifications")

    # -----------------------------------------------------------------
    # PHYSICAL
    # -----------------------------------------------------------------
    st.subheader("Physical")

    # For each "base" physical row, we'll store in session_state so we can clear it
    init_ss("description_spec", "X with Characteristic taste and odour")
    init_ss("description_result", "Compiles")
    init_ss("description_method", "Physical")

    phys1_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["description_spec"] = phys1_cols[0].text_area(
    "Spec for Description",
    value=st.session_state["description_spec"],
    height = 68  # Adjust as needed
    )
    st.session_state["description_result"] = phys1_cols[1].text_input("Result for Description",
        value=st.session_state["description_result"])
    st.session_state["description_method"] = phys1_cols[2].text_input("Method for Description",
        value=st.session_state["description_method"])
    if phys1_cols[3].button("Delete", key="del_desc"):
        st.session_state["description_spec"] = ""
        st.session_state["description_result"] = ""
        st.session_state["description_method"] = ""
        st.rerun()

    init_ss("identification_spec", "To comply by TLC")
    init_ss("identification_result", "Compiles")
    init_ss("identification_method", "TLC")

    phys2_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["identification_spec"] = phys2_cols[0].text_input("Spec for Identification",
        value=st.session_state["identification_spec"])
    st.session_state["identification_result"] = phys2_cols[1].text_input("Result for Identification",
        value=st.session_state["identification_result"])
    st.session_state["identification_method"] = phys2_cols[2].text_input("Method for Identification",
        value=st.session_state["identification_method"])
    if phys2_cols[3].button("Delete", key="del_ident"):
        st.session_state["identification_spec"] = ""
        st.session_state["identification_result"] = ""
        st.session_state["identification_method"] = ""
        st.rerun()

    init_ss("loss_on_drying_spec", "Not more than X")
    init_ss("loss_on_drying_result", "")
    init_ss("loss_on_drying_method", "USP<731>")

    phys3_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["loss_on_drying_spec"] = phys3_cols[0].text_input("Spec for Loss on Drying",
        value=st.session_state["loss_on_drying_spec"])
    st.session_state["loss_on_drying_result"] = phys3_cols[1].text_input("Result for Loss on Drying",
        value=st.session_state["loss_on_drying_result"], placeholder="X")
    st.session_state["loss_on_drying_method"] = phys3_cols[2].text_input("Method for Loss on Drying",
        value=st.session_state["loss_on_drying_method"])
    if phys3_cols[3].button("Delete", key="del_lod"):
        st.session_state["loss_on_drying_spec"] = ""
        st.session_state["loss_on_drying_result"] = ""
        st.session_state["loss_on_drying_method"] = ""
        st.rerun()

    init_ss("moisture_spec", "Not more than X")
    init_ss("moisture_result", "")
    init_ss("moisture_method", "USP<921>")

    phys4_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["moisture_spec"] = phys4_cols[0].text_input("Spec for Moisture",
        value=st.session_state["moisture_spec"])
    st.session_state["moisture_result"] = phys4_cols[1].text_input("Result for Moisture",
        value=st.session_state["moisture_result"], placeholder="X")
    st.session_state["moisture_method"] = phys4_cols[2].text_input("Method for Moisture",
        value=st.session_state["moisture_method"])
    if phys4_cols[3].button("Delete", key="del_moist"):
        st.session_state["moisture_spec"] = ""
        st.session_state["moisture_result"] = ""
        st.session_state["moisture_method"] = ""
        st.rerun()

    init_ss("particle_size_spec", "")
    init_ss("particle_size_result", "")
    init_ss("particle_size_method", "USP<786>")

    phys5_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["particle_size_spec"] = phys5_cols[0].text_input("Spec for Particle Size",
        value=st.session_state["particle_size_spec"], placeholder="X")
    st.session_state["particle_size_result"] = phys5_cols[1].text_input("Result for Particle Size",
        value=st.session_state["particle_size_result"], placeholder="X")
    st.session_state["particle_size_method"] = phys5_cols[2].text_input("Method for Particle Size",
        value=st.session_state["particle_size_method"])
    if phys5_cols[3].button("Delete", key="del_partsize"):
        st.session_state["particle_size_spec"] = ""
        st.session_state["particle_size_result"] = ""
        st.session_state["particle_size_method"] = ""
        st.rerun()

    init_ss("ash_contents_spec", "Not more than X")
    init_ss("ash_contents_result", "")
    init_ss("ash_contents_method", "USP<561>")

    phys6_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["ash_contents_spec"] = phys6_cols[0].text_input("Spec for Ash Contents",
        value=st.session_state["ash_contents_spec"])
    st.session_state["ash_contents_result"] = phys6_cols[1].text_input("Result for Ash Contents",
        value=st.session_state["ash_contents_result"], placeholder="X")
    st.session_state["ash_contents_method"] = phys6_cols[2].text_input("Method for Ash Contents",
        value=st.session_state["ash_contents_method"])
    if phys6_cols[3].button("Delete", key="del_ash"):
        st.session_state["ash_contents_spec"] = ""
        st.session_state["ash_contents_result"] = ""
        st.session_state["ash_contents_method"] = ""
        st.rerun()

    init_ss("residue_on_ignition_spec", "Not more than X")
    init_ss("residue_on_ignition_result", "")
    init_ss("residue_on_ignition_method", "USP<281>")

    phys7_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["residue_on_ignition_spec"] = phys7_cols[0].text_input("Spec for Residue on Ignition",
        value=st.session_state["residue_on_ignition_spec"])
    st.session_state["residue_on_ignition_result"] = phys7_cols[1].text_input("Result for Residue on Ignition",
        value=st.session_state["residue_on_ignition_result"], placeholder="X")
    st.session_state["residue_on_ignition_method"] = phys7_cols[2].text_input("Method for Residue on Ignition",
        value=st.session_state["residue_on_ignition_method"])
    if phys7_cols[3].button("Delete", key="del_resign"):
        st.session_state["residue_on_ignition_spec"] = ""
        st.session_state["residue_on_ignition_result"] = ""
        st.session_state["residue_on_ignition_method"] = ""
        st.rerun()

    init_ss("bulk_density_spec", "Between 0.3g/ml to 0.6g/ml")
    init_ss("bulk_density_result", "")
    init_ss("bulk_density_method", "USP<616>")

    phys8_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["bulk_density_spec"] = phys8_cols[0].text_input("Spec for Bulk Density",
        value=st.session_state["bulk_density_spec"])
    st.session_state["bulk_density_result"] = phys8_cols[1].text_input("Result for Bulk Density",
        value=st.session_state["bulk_density_result"], placeholder="X")
    st.session_state["bulk_density_method"] = phys8_cols[2].text_input("Method for Bulk Density",
        value=st.session_state["bulk_density_method"])
    if phys8_cols[3].button("Delete", key="del_bulk"):
        st.session_state["bulk_density_spec"] = ""
        st.session_state["bulk_density_result"] = ""
        st.session_state["bulk_density_method"] = ""
        st.rerun()

    init_ss("tapped_density_spec", "Between 0.4g/ml to 0.8g/ml")
    init_ss("tapped_density_result", "")
    init_ss("tapped_density_method", "USP<616>")

    phys9_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["tapped_density_spec"] = phys9_cols[0].text_input("Spec for Tapped Density",
        value=st.session_state["tapped_density_spec"])
    st.session_state["tapped_density_result"] = phys9_cols[1].text_input("Result for Tapped Density",
        value=st.session_state["tapped_density_result"], placeholder="X")
    st.session_state["tapped_density_method"] = phys9_cols[2].text_input("Method for Tapped Density",
        value=st.session_state["tapped_density_method"])
    if phys9_cols[3].button("Delete", key="del_tapped"):
        st.session_state["tapped_density_spec"] = ""
        st.session_state["tapped_density_result"] = ""
        st.session_state["tapped_density_method"] = ""
        st.rerun()

    init_ss("solubility_spec", "")
    init_ss("solubility_result", "")
    init_ss("solubility_method", "USP<1236>")

    phys10_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["solubility_spec"] = phys10_cols[0].text_input("Spec for Solubility",
        value=st.session_state["solubility_spec"], placeholder="X")
    st.session_state["solubility_result"] = phys10_cols[1].text_input("Result for Solubility",
        value=st.session_state["solubility_result"], placeholder="X")
    st.session_state["solubility_method"] = phys10_cols[2].text_input("Method for Solubility",
        value=st.session_state["solubility_method"])
    if phys10_cols[3].button("Delete", key="del_solub"):
        st.session_state["solubility_spec"] = ""
        st.session_state["solubility_result"] = ""
        st.session_state["solubility_method"] = ""
        st.rerun()

    init_ss("ph_spec", "")
    init_ss("ph_result", "")
    init_ss("ph_method", "USP<791>")

    phys11_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["ph_spec"] = phys11_cols[0].text_input("Spec for pH",
        value=st.session_state["ph_spec"], placeholder="X")
    st.session_state["ph_result"] = phys11_cols[1].text_input("Result for pH",
        value=st.session_state["ph_result"], placeholder="X")
    st.session_state["ph_method"] = phys11_cols[2].text_input("Method for pH",
        value=st.session_state["ph_method"])
    if phys11_cols[3].button("Delete", key="del_ph"):
        st.session_state["ph_spec"] = ""
        st.session_state["ph_result"] = ""
        st.session_state["ph_method"] = ""
        st.rerun()

    init_ss("chlorides_nacl_spec", "")
    init_ss("chlorides_nacl_result", "")
    init_ss("chlorides_nacl_method", "USP<221>")

    phys12_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["chlorides_nacl_spec"] = phys12_cols[0].text_input("Spec for Chlorides of NaCl",
        value=st.session_state["chlorides_nacl_spec"], placeholder="X")
    st.session_state["chlorides_nacl_result"] = phys12_cols[1].text_input("Result for Chlorides of NaCl",
        value=st.session_state["chlorides_nacl_result"], placeholder="X")
    st.session_state["chlorides_nacl_method"] = phys12_cols[2].text_input("Method for Chlorides of NaCl",
        value=st.session_state["chlorides_nacl_method"])
    if phys12_cols[3].button("Delete", key="del_chlorides"):
        st.session_state["chlorides_nacl_spec"] = ""
        st.session_state["chlorides_nacl_result"] = ""
        st.session_state["chlorides_nacl_method"] = ""
        st.rerun()

    init_ss("sulphates_spec", "")
    init_ss("sulphates_result", "")
    init_ss("sulphates_method", "USP<221>")

    phys13_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["sulphates_spec"] = phys13_cols[0].text_input("Spec for Sulphates",
        value=st.session_state["sulphates_spec"], placeholder="X")
    st.session_state["sulphates_result"] = phys13_cols[1].text_input("Result for Sulphates",
        value=st.session_state["sulphates_result"], placeholder="X")
    st.session_state["sulphates_method"] = phys13_cols[2].text_input("Method for Sulphates",
        value=st.session_state["sulphates_method"])
    if phys13_cols[3].button("Delete", key="del_sulphates"):
        st.session_state["sulphates_spec"] = ""
        st.session_state["sulphates_result"] = ""
        st.session_state["sulphates_method"] = ""
        st.rerun()

    init_ss("fats_spec", "")
    init_ss("fats_result", "")
    init_ss("fats_method", "USP<731>")

    phys14_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["fats_spec"] = phys14_cols[0].text_input("Spec for Fats",
        value=st.session_state["fats_spec"], placeholder="X")
    st.session_state["fats_result"] = phys14_cols[1].text_input("Result for Fats",
        value=st.session_state["fats_result"], placeholder="X")
    st.session_state["fats_method"] = phys14_cols[2].text_input("Method for Fats",
        value=st.session_state["fats_method"])
    if phys14_cols[3].button("Delete", key="del_fats"):
        st.session_state["fats_spec"] = ""
        st.session_state["fats_result"] = ""
        st.session_state["fats_method"] = ""
        st.rerun()

    init_ss("protein_spec", "")
    init_ss("protein_result", "")
    init_ss("protein_method", "Kjeldahl")

    phys15_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["protein_spec"] = phys15_cols[0].text_input("Spec for Protein",
        value=st.session_state["protein_spec"], placeholder="X")
    st.session_state["protein_result"] = phys15_cols[1].text_input("Result for Protein",
        value=st.session_state["protein_result"], placeholder="X")
    st.session_state["protein_method"] = phys15_cols[2].text_input("Method for Protein",
        value=st.session_state["protein_method"])
    if phys15_cols[3].button("Delete", key="del_protein"):
        st.session_state["protein_spec"] = ""
        st.session_state["protein_result"] = ""
        st.session_state["protein_method"] = ""
        st.rerun()

    init_ss("total_ig_g_spec", "")
    init_ss("total_ig_g_result", "")
    init_ss("total_ig_g_method", "HPLC")

    phys16_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["total_ig_g_spec"] = phys16_cols[0].text_input("Spec for Total IgG",
        value=st.session_state["total_ig_g_spec"], placeholder="X")
    st.session_state["total_ig_g_result"] = phys16_cols[1].text_input("Result for Total IgG",
        value=st.session_state["total_ig_g_result"], placeholder="X")
    st.session_state["total_ig_g_method"] = phys16_cols[2].text_input("Method for Total IgG",
        value=st.session_state["total_ig_g_method"])
    if phys16_cols[3].button("Delete", key="del_igg"):
        st.session_state["total_ig_g_spec"] = ""
        st.session_state["total_ig_g_result"] = ""
        st.session_state["total_ig_g_method"] = ""
        st.rerun()

    init_ss("sodium_spec", "")
    init_ss("sodium_result", "")
    init_ss("sodium_method", "ICP-MS")

    phys17_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["sodium_spec"] = phys17_cols[0].text_input("Spec for Sodium",
        value=st.session_state["sodium_spec"], placeholder="X")
    st.session_state["sodium_result"] = phys17_cols[1].text_input("Result for Sodium",
        value=st.session_state["sodium_result"], placeholder="X")
    st.session_state["sodium_method"] = phys17_cols[2].text_input("Method for Sodium",
        value=st.session_state["sodium_method"])
    if phys17_cols[3].button("Delete", key="del_sodium"):
        st.session_state["sodium_spec"] = ""
        st.session_state["sodium_result"] = ""
        st.session_state["sodium_method"] = ""
        st.rerun()

    init_ss("gluten_spec", "NMT X")
    init_ss("gluten_result", "")
    init_ss("gluten_method", "ELISA")

    phys18_cols = st.columns([3, 2.5, 2.5, 2])
    st.session_state["gluten_spec"] = phys18_cols[0].text_input("Spec for Gluten",
        value=st.session_state["gluten_spec"])
    st.session_state["gluten_result"] = phys18_cols[1].text_input("Result for Gluten",
        value=st.session_state["gluten_result"], placeholder="X")
    st.session_state["gluten_method"] = phys18_cols[2].text_input("Method for Gluten",
        value=st.session_state["gluten_method"])
    if phys18_cols[3].button("Delete", key="del_gluten"):
        st.session_state["gluten_spec"] = ""
        st.session_state["gluten_result"] = ""
        st.session_state["gluten_method"] = ""
        st.rerun()

    st.markdown("#### Add Additional Physical Rows")
    for i, row_data in enumerate(st.session_state["Physical_rows"]):
        c1, c2, c3, c4, del_col = st.columns([3, 2.5, 2.5, 2.5, 2])
        st.session_state["Physical_rows"][i]["param"] = c1.text_input(
            f"Physical Parameter {i+1}", row_data["param"], key=f"PhysicalParam_{i}"
        )
        st.session_state["Physical_rows"][i]["spec"] = c2.text_input(
            f"Physical Spec {i+1}", row_data["spec"], key=f"PhysicalSpec_{i}"
        )
        st.session_state["Physical_rows"][i]["result"] = c3.text_input(
            f"Physical Result {i+1}", row_data["result"], key=f"PhysicalResult_{i}"
        )
        st.session_state["Physical_rows"][i]["method"] = c4.text_input(
            f"Physical Method {i+1}", row_data["method"], key=f"PhysicalMethod_{i}"
        )
        if del_col.button("Delete", key=f"del_physical_{i}"):
            st.session_state["Physical_rows"].pop(i)
            st.rerun()

    if st.button("Add New Physical Row"):
        st.session_state["Physical_rows"].append({"param": "", "spec": "", "result": "", "method": ""})
        st.rerun()

    # ----------------------------------------------------------------------
    # OTHERS
    # ----------------------------------------------------------------------
    st.subheader("Others")
    init_ss("lead_spec", "Not more than X ppm")
    init_ss("lead_result", "")
    init_ss("lead_method", "ICP-MS")

    others_1 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["lead_spec"] = others_1[0].text_input("Spec for Lead", value=st.session_state["lead_spec"])
    st.session_state["lead_result"] = others_1[1].text_input("Result for Lead", value=st.session_state["lead_result"], placeholder="X ppm")
    st.session_state["lead_method"] = others_1[2].text_input("Method for Lead", value=st.session_state["lead_method"])
    if others_1[3].button("Delete", key="del_lead"):
        st.session_state["lead_spec"] = ""
        st.session_state["lead_result"] = ""
        st.session_state["lead_method"] = ""
        st.rerun()

    init_ss("cadmium_spec", "Not more than X ppm")
    init_ss("cadmium_result", "")
    init_ss("cadmium_method", "ICP-MS")

    others_2 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["cadmium_spec"] = others_2[0].text_input("Spec for Cadmium", value=st.session_state["cadmium_spec"])
    st.session_state["cadmium_result"] = others_2[1].text_input("Result for Cadmium", value=st.session_state["cadmium_result"], placeholder="X ppm")
    st.session_state["cadmium_method"] = others_2[2].text_input("Method for Cadmium", value=st.session_state["cadmium_method"])
    if others_2[3].button("Delete", key="del_cadmium"):
        st.session_state["cadmium_spec"] = ""
        st.session_state["cadmium_result"] = ""
        st.session_state["cadmium_method"] = ""
        st.rerun()

    init_ss("arsenic_spec", "Not more than X ppm")
    init_ss("arsenic_result", "")
    init_ss("arsenic_method", "ICP-MS")

    others_3 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["arsenic_spec"] = others_3[0].text_input("Spec for Arsenic", value=st.session_state["arsenic_spec"])
    st.session_state["arsenic_result"] = others_3[1].text_input("Result for Arsenic", value=st.session_state["arsenic_result"], placeholder="X ppm")
    st.session_state["arsenic_method"] = others_3[2].text_input("Method for Arsenic", value=st.session_state["arsenic_method"])
    if others_3[3].button("Delete", key="del_arsenic"):
        st.session_state["arsenic_spec"] = ""
        st.session_state["arsenic_result"] = ""
        st.session_state["arsenic_method"] = ""
        st.rerun()

    init_ss("mercury_spec", "Not more than X ppm")
    init_ss("mercury_result", "")
    init_ss("mercury_method", "ICP-MS")

    others_4 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["mercury_spec"] = others_4[0].text_input("Spec for Mercury", value=st.session_state["mercury_spec"])
    st.session_state["mercury_result"] = others_4[1].text_input("Result for Mercury", value=st.session_state["mercury_result"], placeholder="X ppm")
    st.session_state["mercury_method"] = others_4[2].text_input("Method for Mercury", value=st.session_state["mercury_method"])
    if others_4[3].button("Delete", key="del_mercury"):
        st.session_state["mercury_spec"] = ""
        st.session_state["mercury_result"] = ""
        st.session_state["mercury_method"] = ""
        st.rerun()

    st.markdown("#### Add Additional Others Rows")
    for i, row_data in enumerate(st.session_state["Others_rows"]):
        c1, c2, c3, c4, del_col = st.columns([3, 2.5, 2.5, 2.5, 2])
        st.session_state["Others_rows"][i]["param"] = c1.text_input(
            f"Others Parameter {i+1}", row_data.get("param",""), key=f"OthersParam_{i}"
        )
        st.session_state["Others_rows"][i]["spec"] = c2.text_input(
            f"Others Spec {i+1}", row_data.get("spec",""), key=f"OthersSpec_{i}"
        )
        st.session_state["Others_rows"][i]["result"] = c3.text_input(
            f"Others Result {i+1}", row_data.get("result",""), key=f"OthersResult_{i}"
        )
        st.session_state["Others_rows"][i]["method"] = c4.text_input(
            f"Others Method {i+1}", row_data.get("method",""), key=f"OthersMethod_{i}"
        )
        if del_col.button("Delete", key=f"del_others_{i}"):
            st.session_state["Others_rows"].pop(i)
            st.rerun()

    if st.button("Add New Others Row"):
        st.session_state["Others_rows"].append({"param": "", "spec": "", "result": "", "method": ""})
        st.rerun()

    # ASSAYS
    st.subheader("Assays")
    init_ss("assays_spec", "")
    init_ss("assays_result", "")
    init_ss("assays_method", "")

    assays_1 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["assays_spec"] = assays_1[0].text_input("Specification for Assays",
        value=st.session_state["assays_spec"], placeholder="X")
    st.session_state["assays_result"] = assays_1[1].text_input("Result for Assays",
        value=st.session_state["assays_result"], placeholder="X")
    st.session_state["assays_method"] = assays_1[2].text_input("Method for Assays",
        value=st.session_state["assays_method"], placeholder="X")
    if assays_1[3].button("Delete", key="del_assays_base"):
        st.session_state["assays_spec"] = ""
        st.session_state["assays_result"] = ""
        st.session_state["assays_method"] = ""
        st.rerun()

    st.markdown("#### Add Additional Assays Rows")
    for i, row_data in enumerate(st.session_state["Assays_rows"]):
        c1, c2, c3, c4, del_col = st.columns([3, 2.5, 2.5, 2.5, 2])
        st.session_state["Assays_rows"][i]["param"] = c1.text_input(
            f"Assays Parameter {i+1}", row_data.get("param",""), key=f"AssaysParam_{i}"
        )
        st.session_state["Assays_rows"][i]["spec"] = c2.text_input(
            f"Assays Spec {i+1}", row_data.get("spec",""), key=f"AssaysSpec_{i}"
        )
        st.session_state["Assays_rows"][i]["result"] = c3.text_input(
            f"Assays Result {i+1}", row_data.get("result",""), key=f"AssaysResult_{i}"
        )
        st.session_state["Assays_rows"][i]["method"] = c4.text_input(
            f"Assays Method {i+1}", row_data.get("method",""), key=f"AssaysMethod_{i}"
        )
        if del_col.button("Delete", key=f"del_assays_{i}"):
            st.session_state["Assays_rows"].pop(i)
            st.rerun()

    if st.button("Add New Assays Row"):
        st.session_state["Assays_rows"].append({"param": "", "spec": "", "result": "", "method": ""})
        st.rerun()

    # PESTICIDES
    st.subheader("Pesticides")
    init_ss("pesticide_spec", "Meet USP<561>")
    init_ss("pesticide_result", "Compiles")
    init_ss("pesticide_method", "USP<561>")

    pest_1 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["pesticide_spec"] = pest_1[0].text_input("Specification for Pesticide",
        value=st.session_state["pesticide_spec"])
    st.session_state["pesticide_result"] = pest_1[1].text_input("Result for Pesticide",
        value=st.session_state["pesticide_result"])
    st.session_state["pesticide_method"] = pest_1[2].text_input("Method for Pesticide",
        value=st.session_state["pesticide_method"])
    if pest_1[3].button("Delete", key="del_pesticide_base"):
        st.session_state["pesticide_spec"] = ""
        st.session_state["pesticide_result"] = ""
        st.session_state["pesticide_method"] = ""
        st.rerun()

    st.markdown("#### Add Additional Pesticides Rows")
    for i, row_data in enumerate(st.session_state["Pesticides_rows"]):
        c1, c2, c3, c4, del_col = st.columns([3, 2.5, 2.5, 2.5, 2])
        st.session_state["Pesticides_rows"][i]["param"] = c1.text_input(
            f"Pesticides Parameter {i+1}", row_data.get("param",""), key=f"PesticidesParam_{i}"
        )
        st.session_state["Pesticides_rows"][i]["spec"] = c2.text_input(
            f"Pesticides Spec {i+1}", row_data.get("spec",""), key=f"PesticidesSpec_{i}"
        )
        st.session_state["Pesticides_rows"][i]["result"] = c3.text_input(
            f"Pesticides Result {i+1}", row_data.get("result",""), key=f"PesticidesResult_{i}"
        )
        st.session_state["Pesticides_rows"][i]["method"] = c4.text_input(
            f"Pesticides Method {i+1}", row_data.get("method",""), key=f"PesticidesMethod_{i}"
        )
        if del_col.button("Delete", key=f"del_pesticides_{i}"):
            st.session_state["Pesticides_rows"].pop(i)
            st.rerun()

    if st.button("Add New Pesticides Row"):
        st.session_state["Pesticides_rows"].append({"param": "", "spec": "", "result": "", "method": ""})
        st.rerun()

    # RESIDUAL SOLVENT
    st.subheader("Residual Solvent")
    init_ss("residual_solvent_spec", "")
    init_ss("residual_solvent_result", "Compiles")
    init_ss("residual_solvent_method", "")

    rs_1 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["residual_solvent_spec"] = rs_1[0].text_input("Specification for Residual Solvent",
        value=st.session_state["residual_solvent_spec"], placeholder="X")
    st.session_state["residual_solvent_result"] = rs_1[1].text_input("Result for Residual Solvent",
        value=st.session_state["residual_solvent_result"])
    st.session_state["residual_solvent_method"] = rs_1[2].text_input("Method for Residual Solvent",
        value=st.session_state["residual_solvent_method"], placeholder="X")
    if rs_1[3].button("Delete", key="del_resid_base"):
        st.session_state["residual_solvent_spec"] = ""
        st.session_state["residual_solvent_result"] = ""
        st.session_state["residual_solvent_method"] = ""
        st.rerun()

    st.markdown("#### Add Additional Residual Solvent Rows")
    for i, row_data in enumerate(st.session_state["ResidualSolvent_rows"]):
        c1, c2, c3, c4, del_col = st.columns([3, 2.5, 2.5, 2.5, 2])
        st.session_state["ResidualSolvent_rows"][i]["param"] = c1.text_input(
            f"Residual Solvent Parameter {i+1}", row_data.get("param",""), key=f"ResidualSolventParam_{i}"
        )
        st.session_state["ResidualSolvent_rows"][i]["spec"] = c2.text_input(
            f"Residual Solvent Spec {i+1}", row_data.get("spec",""), key=f"ResidualSolventSpec_{i}"
        )
        st.session_state["ResidualSolvent_rows"][i]["result"] = c3.text_input(
            f"Residual Solvent Result {i+1}", row_data.get("result",""), key=f"ResidualSolventResult_{i}"
        )
        st.session_state["ResidualSolvent_rows"][i]["method"] = c4.text_input(
            f"Residual Solvent Method {i+1}", row_data.get("method",""), key=f"ResidualSolventMethod_{i}"
        )
        if del_col.button("Delete", key=f"del_residual_{i}"):
            st.session_state["ResidualSolvent_rows"].pop(i)
            st.rerun()

    if st.button("Add New Residual Solvent Row"):
        st.session_state["ResidualSolvent_rows"].append({"param": "", "spec": "", "result": "", "method": ""})
        st.rerun()

    # MICROBIOLOGICAL
    st.subheader("Microbiological Profile")
    init_ss("total_plate_count_spec", "Not more than X cfu/g")
    init_ss("total_plate_count_result", "X cfu/g")
    init_ss("total_plate_count_method", "USP<61>")

    micro_1 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["total_plate_count_spec"] = micro_1[0].text_input("Spec for Total Plate Count",
        value=st.session_state["total_plate_count_spec"])
    st.session_state["total_plate_count_result"] = micro_1[1].text_input("Result for Total Plate Count",
        value=st.session_state["total_plate_count_result"])
    st.session_state["total_plate_count_method"] = micro_1[2].text_input("Method for Total Plate Count",
        value=st.session_state["total_plate_count_method"])
    if micro_1[3].button("Delete", key="del_tpc"):
        st.session_state["total_plate_count_spec"] = ""
        st.session_state["total_plate_count_result"] = ""
        st.session_state["total_plate_count_method"] = ""
        st.rerun()

    init_ss("yeasts_mould_spec", "Not more than X cfu/g")
    init_ss("yeasts_mould_result", "X cfu/g")
    init_ss("yeasts_mould_method", "USP<61>")

    micro_2 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["yeasts_mould_spec"] = micro_2[0].text_input("Spec for Yeasts & Mould Count",
        value=st.session_state["yeasts_mould_spec"])
    st.session_state["yeasts_mould_result"] = micro_2[1].text_input("Result for Yeasts & Mould Count",
        value=st.session_state["yeasts_mould_result"])
    st.session_state["yeasts_mould_method"] = micro_2[2].text_input("Method for Yeasts & Mould Count",
        value=st.session_state["yeasts_mould_method"])
    if micro_2[3].button("Delete", key="del_ym"):
        st.session_state["yeasts_mould_spec"] = ""
        st.session_state["yeasts_mould_result"] = ""
        st.session_state["yeasts_mould_method"] = ""
        st.rerun()

    init_ss("salmonella_spec", "Absent/25g")
    init_ss("salmonella_result", "Absent")
    init_ss("salmonella_method", "USP<62>")

    micro_3 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["salmonella_spec"] = micro_3[0].text_input("Spec for Salmonella",
        value=st.session_state["salmonella_spec"])
    st.session_state["salmonella_result"] = micro_3[1].text_input("Result for Salmonella",
        value=st.session_state["salmonella_result"])
    st.session_state["salmonella_method"] = micro_3[2].text_input("Method for Salmonella",
        value=st.session_state["salmonella_method"])
    if micro_3[3].button("Delete", key="del_salmonella"):
        st.session_state["salmonella_spec"] = ""
        st.session_state["salmonella_result"] = ""
        st.session_state["salmonella_method"] = ""
        st.rerun()

    init_ss("e_coli_spec", "Absent/10g")
    init_ss("e_coli_result", "Absent")
    init_ss("e_coli_method", "USP<62>")

    micro_4 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["e_coli_spec"] = micro_4[0].text_input("Spec for Escherichia coli",
        value=st.session_state["e_coli_spec"])
    st.session_state["e_coli_result"] = micro_4[1].text_input("Result for Escherichia coli",
        value=st.session_state["e_coli_result"])
    st.session_state["e_coli_method"] = micro_4[2].text_input("Method for Escherichia coli",
        value=st.session_state["e_coli_method"])
    if micro_4[3].button("Delete", key="del_ecoli"):
        st.session_state["e_coli_spec"] = ""
        st.session_state["e_coli_result"] = ""
        st.session_state["e_coli_method"] = ""
        st.rerun()

    init_ss("coliforms_spec", "NMT X cfu/g")
    init_ss("coliforms_result", "")
    init_ss("coliforms_method", "USP<62>")

    micro_5 = st.columns([3, 2.5, 2.5, 2])
    st.session_state["coliforms_spec"] = micro_5[0].text_input("Spec for Coliforms",
        value=st.session_state["coliforms_spec"])
    st.session_state["coliforms_result"] = micro_5[1].text_input("Result for Coliforms",
        value=st.session_state["coliforms_result"], placeholder="X")
    st.session_state["coliforms_method"] = micro_5[2].text_input("Method for Coliforms",
        value=st.session_state["coliforms_method"])
    if micro_5[3].button("Delete", key="del_coliforms"):
        st.session_state["coliforms_spec"] = ""
        st.session_state["coliforms_result"] = ""
        st.session_state["coliforms_method"] = ""
        st.rerun()

    st.markdown("#### Add Additional Microbiological Profile Rows")
    for i, row_data in enumerate(st.session_state["MicrobiologicalProfile_rows"]):
        c1, c2, c3, c4, del_col = st.columns([3, 2.5, 2.5, 2.5, 2])
        st.session_state["MicrobiologicalProfile_rows"][i]["param"] = c1.text_input(
            f"Microbio Parameter {i+1}", row_data.get("param",""), key=f"MicrobioParam_{i}"
        )
        st.session_state["MicrobiologicalProfile_rows"][i]["spec"] = c2.text_input(
            f"Microbio Spec {i+1}", row_data.get("spec",""), key=f"MicrobioSpec_{i}"
        )
        st.session_state["MicrobiologicalProfile_rows"][i]["result"] = c3.text_input(
            f"Microbio Result {i+1}", row_data.get("result",""), key=f"MicrobioResult_{i}"
        )
        st.session_state["MicrobiologicalProfile_rows"][i]["method"] = c4.text_input(
            f"Microbio Method {i+1}", row_data.get("method",""), key=f"MicrobioMethod_{i}"
        )
        if del_col.button("Delete", key=f"del_micro_{i}"):
            st.session_state["MicrobiologicalProfile_rows"].pop(i)
            st.rerun()

    if st.button("Add New Microbiological Row"):
        st.session_state["MicrobiologicalProfile_rows"].append({"param": "", "spec": "", "result": "", "method": ""})
        st.rerun()

    # Declaration
    st.subheader("Declaration - Allergen Statement")
    allergen_statement = st.selectbox("Allergen Statement", options=["Free from allergen", "Contains Allergen"])

    # ----------- PREVIEW & COMPILE BUTTONS -----------
    st.write("---")
    if st.button("Preview"):
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

            "description_spec": st.session_state["description_spec"],
            "description_result": st.session_state["description_result"],
            "description_method": st.session_state["description_method"],

            "identification_spec": st.session_state["identification_spec"],
            "identification_result": st.session_state["identification_result"],
            "identification_method": st.session_state["identification_method"],

            "loss_on_drying_spec": st.session_state["loss_on_drying_spec"],
            "loss_on_drying_result": st.session_state["loss_on_drying_result"],
            "loss_on_drying_method": st.session_state["loss_on_drying_method"],

            "moisture_spec": st.session_state["moisture_spec"],
            "moisture_result": st.session_state["moisture_result"],
            "moisture_method": st.session_state["moisture_method"],

            "particle_size_spec": st.session_state["particle_size_spec"],
            "particle_size_result": st.session_state["particle_size_result"],
            "particle_size_method": st.session_state["particle_size_method"],

            "ash_contents_spec": st.session_state["ash_contents_spec"],
            "ash_contents_result": st.session_state["ash_contents_result"],
            "ash_contents_method": st.session_state["ash_contents_method"],

            "residue_on_ignition_spec": st.session_state["residue_on_ignition_spec"],
            "residue_on_ignition_result": st.session_state["residue_on_ignition_result"],
            "residue_on_ignition_method": st.session_state["residue_on_ignition_method"],

            "bulk_density_spec": st.session_state["bulk_density_spec"],
            "bulk_density_result": st.session_state["bulk_density_result"],
            "bulk_density_method": st.session_state["bulk_density_method"],

            "tapped_density_spec": st.session_state["tapped_density_spec"],
            "tapped_density_result": st.session_state["tapped_density_result"],
            "tapped_density_method": st.session_state["tapped_density_method"],

            "solubility_spec": st.session_state["solubility_spec"],
            "solubility_result": st.session_state["solubility_result"],
            "solubility_method": st.session_state["solubility_method"],

            "ph_spec": st.session_state["ph_spec"],
            "ph_result": st.session_state["ph_result"],
            "ph_method": st.session_state["ph_method"],

            "chlorides_nacl_spec": st.session_state["chlorides_nacl_spec"],
            "chlorides_nacl_result": st.session_state["chlorides_nacl_result"],
            "chlorides_nacl_method": st.session_state["chlorides_nacl_method"],

            "sulphates_spec": st.session_state["sulphates_spec"],
            "sulphates_result": st.session_state["sulphates_result"],
            "sulphates_method": st.session_state["sulphates_method"],

            "fats_spec": st.session_state["fats_spec"],
            "fats_result": st.session_state["fats_result"],
            "fats_method": st.session_state["fats_method"],

            "protein_spec": st.session_state["protein_spec"],
            "protein_result": st.session_state["protein_result"],
            "protein_method": st.session_state["protein_method"],

            "total_ig_g_spec": st.session_state["total_ig_g_spec"],
            "total_ig_g_result": st.session_state["total_ig_g_result"],
            "total_ig_g_method": st.session_state["total_ig_g_method"],

            "sodium_spec": st.session_state["sodium_spec"],
            "sodium_result": st.session_state["sodium_result"],
            "sodium_method": st.session_state["sodium_method"],

            "gluten_spec": st.session_state["gluten_spec"],
            "gluten_result": st.session_state["gluten_result"],
            "gluten_method": st.session_state["gluten_method"],

            "lead_spec": st.session_state["lead_spec"],
            "lead_result": st.session_state["lead_result"],
            "lead_method": st.session_state["lead_method"],

            "cadmium_spec": st.session_state["cadmium_spec"],
            "cadmium_result": st.session_state["cadmium_result"],
            "cadmium_method": st.session_state["cadmium_method"],

            "arsenic_spec": st.session_state["arsenic_spec"],
            "arsenic_result": st.session_state["arsenic_result"],
            "arsenic_method": st.session_state["arsenic_method"],

            "mercury_spec": st.session_state["mercury_spec"],
            "mercury_result": st.session_state["mercury_result"],
            "mercury_method": st.session_state["mercury_method"],

            "assays_spec": st.session_state["assays_spec"],
            "assays_result": st.session_state["assays_result"],
            "assays_method": st.session_state["assays_method"],

            "pesticide_spec": st.session_state["pesticide_spec"],
            "pesticide_result": st.session_state["pesticide_result"],
            "pesticide_method": st.session_state["pesticide_method"],

            "residual_solvent_spec": st.session_state["residual_solvent_spec"],
            "residual_solvent_result": st.session_state["residual_solvent_result"],
            "residual_solvent_method": st.session_state["residual_solvent_method"],

            "total_plate_count_spec": st.session_state["total_plate_count_spec"],
            "total_plate_count_result": st.session_state["total_plate_count_result"],
            "total_plate_count_method": st.session_state["total_plate_count_method"],

            "yeasts_mould_spec": st.session_state["yeasts_mould_spec"],
            "yeasts_mould_result": st.session_state["yeasts_mould_result"],
            "yeasts_mould_method": st.session_state["yeasts_mould_method"],

            "salmonella_spec": st.session_state["salmonella_spec"],
            "salmonella_result": st.session_state["salmonella_result"],
            "salmonella_method": st.session_state["salmonella_method"],

            "e_coli_spec": st.session_state["e_coli_spec"],
            "e_coli_result": st.session_state["e_coli_result"],
            "e_coli_method": st.session_state["e_coli_method"],

            "coliforms_spec": st.session_state["coliforms_spec"],
            "coliforms_result": st.session_state["coliforms_result"],
            "coliforms_method": st.session_state["coliforms_method"],

            "allergen_statement": allergen_statement,

            "physical_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Physical_rows"]
            ],
            "others_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Others_rows"]
            ],
            "assays_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Assays_rows"]
            ],
            "pesticides_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Pesticides_rows"]
            ],
            "residual_solvent_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["ResidualSolvent_rows"]
            ],
            "microbio_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["MicrobiologicalProfile_rows"]
            ],
        }

        pdf_buffer = generate_pdf(data)
        if pdf_buffer:
            doc_preview = fitz.open(stream=pdf_buffer, filetype="pdf")
            with col2:
                for page in doc_preview:
                    pix = page.get_pixmap()
                    st.image(pix.tobytes(), caption=f"Page {page.number + 1}", use_container_width=True)
            st.success("Preview generated successfully!")

    if st.button("Compile and Generate PDF"):
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

            "description_spec": st.session_state["description_spec"],
            "description_result": st.session_state["description_result"],
            "description_method": st.session_state["description_method"],
            "identification_spec": st.session_state["identification_spec"],
            "identification_result": st.session_state["identification_result"],
            "identification_method": st.session_state["identification_method"],
            "loss_on_drying_spec": st.session_state["loss_on_drying_spec"],
            "loss_on_drying_result": st.session_state["loss_on_drying_result"],
            "loss_on_drying_method": st.session_state["loss_on_drying_method"],
            "moisture_spec": st.session_state["moisture_spec"],
            "moisture_result": st.session_state["moisture_result"],
            "moisture_method": st.session_state["moisture_method"],
            "particle_size_spec": st.session_state["particle_size_spec"],
            "particle_size_result": st.session_state["particle_size_result"],
            "particle_size_method": st.session_state["particle_size_method"],
            "ash_contents_spec": st.session_state["ash_contents_spec"],
            "ash_contents_result": st.session_state["ash_contents_result"],
            "ash_contents_method": st.session_state["ash_contents_method"],
            "residue_on_ignition_spec": st.session_state["residue_on_ignition_spec"],
            "residue_on_ignition_result": st.session_state["residue_on_ignition_result"],
            "residue_on_ignition_method": st.session_state["residue_on_ignition_method"],
            "bulk_density_spec": st.session_state["bulk_density_spec"],
            "bulk_density_result": st.session_state["bulk_density_result"],
            "bulk_density_method": st.session_state["bulk_density_method"],
            "tapped_density_spec": st.session_state["tapped_density_spec"],
            "tapped_density_result": st.session_state["tapped_density_result"],
            "tapped_density_method": st.session_state["tapped_density_method"],
            "solubility_spec": st.session_state["solubility_spec"],
            "solubility_result": st.session_state["solubility_result"],
            "solubility_method": st.session_state["solubility_method"],
            "ph_spec": st.session_state["ph_spec"],
            "ph_result": st.session_state["ph_result"],
            "ph_method": st.session_state["ph_method"],
            "chlorides_nacl_spec": st.session_state["chlorides_nacl_spec"],
            "chlorides_nacl_result": st.session_state["chlorides_nacl_result"],
            "chlorides_nacl_method": st.session_state["chlorides_nacl_method"],
            "sulphates_spec": st.session_state["sulphates_spec"],
            "sulphates_result": st.session_state["sulphates_result"],
            "sulphates_method": st.session_state["sulphates_method"],
            "fats_spec": st.session_state["fats_spec"],
            "fats_result": st.session_state["fats_result"],
            "fats_method": st.session_state["fats_method"],
            "protein_spec": st.session_state["protein_spec"],
            "protein_result": st.session_state["protein_result"],
            "protein_method": st.session_state["protein_method"],
            "total_ig_g_spec": st.session_state["total_ig_g_spec"],
            "total_ig_g_result": st.session_state["total_ig_g_result"],
            "total_ig_g_method": st.session_state["total_ig_g_method"],
            "sodium_spec": st.session_state["sodium_spec"],
            "sodium_result": st.session_state["sodium_result"],
            "sodium_method": st.session_state["sodium_method"],
            "gluten_spec": st.session_state["gluten_spec"],
            "gluten_result": st.session_state["gluten_result"],
            "gluten_method": st.session_state["gluten_method"],

            "lead_spec": st.session_state["lead_spec"],
            "lead_result": st.session_state["lead_result"],
            "lead_method": st.session_state["lead_method"],
            "cadmium_spec": st.session_state["cadmium_spec"],
            "cadmium_result": st.session_state["cadmium_result"],
            "cadmium_method": st.session_state["cadmium_method"],
            "arsenic_spec": st.session_state["arsenic_spec"],
            "arsenic_result": st.session_state["arsenic_result"],
            "arsenic_method": st.session_state["arsenic_method"],
            "mercury_spec": st.session_state["mercury_spec"],
            "mercury_result": st.session_state["mercury_result"],
            "mercury_method": st.session_state["mercury_method"],

            "assays_spec": st.session_state["assays_spec"],
            "assays_result": st.session_state["assays_result"],
            "assays_method": st.session_state["assays_method"],
            "pesticide_spec": st.session_state["pesticide_spec"],
            "pesticide_result": st.session_state["pesticide_result"],
            "pesticide_method": st.session_state["pesticide_method"],
            "residual_solvent_spec": st.session_state["residual_solvent_spec"],
            "residual_solvent_result": st.session_state["residual_solvent_result"],
            "residual_solvent_method": st.session_state["residual_solvent_method"],
            "total_plate_count_spec": st.session_state["total_plate_count_spec"],
            "total_plate_count_result": st.session_state["total_plate_count_result"],
            "total_plate_count_method": st.session_state["total_plate_count_method"],
            "yeasts_mould_spec": st.session_state["yeasts_mould_spec"],
            "yeasts_mould_result": st.session_state["yeasts_mould_result"],
            "yeasts_mould_method": st.session_state["yeasts_mould_method"],
            "salmonella_spec": st.session_state["salmonella_spec"],
            "salmonella_result": st.session_state["salmonella_result"],
            "salmonella_method": st.session_state["salmonella_method"],
            "e_coli_spec": st.session_state["e_coli_spec"],
            "e_coli_result": st.session_state["e_coli_result"],
            "e_coli_method": st.session_state["e_coli_method"],
            "coliforms_spec": st.session_state["coliforms_spec"],
            "coliforms_result": st.session_state["coliforms_result"],
            "coliforms_method": st.session_state["coliforms_method"],

            "allergen_statement": allergen_statement,

            "physical_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Physical_rows"]
            ],
            "others_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Others_rows"]
            ],
            "assays_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Assays_rows"]
            ],
            "pesticides_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["Pesticides_rows"]
            ],
            "residual_solvent_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["ResidualSolvent_rows"]
            ],
            "microbio_extra_rows": [
                (row["param"], row["spec"], row["result"], row["method"])
                for row in st.session_state["MicrobiologicalProfile_rows"]
            ],
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
