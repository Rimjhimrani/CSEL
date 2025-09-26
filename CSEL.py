import streamlit as st
import pandas as pd
import os
from reportlab.lib.pagesizes import landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, PageBreak, Paragraph
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
import sys
import subprocess
import tempfile

# --- MANUAL ROW HEIGHT ADJUSTMENT ---
ROW_HEIGHTS = [1.6*cm, 1.6*cm, 1.8*cm]

# --- STICKER AND CONTENT DIMENSIONS ---
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)
CONTENT_BOX_WIDTH = 9.8 * cm

# --- PARAGRAPH STYLES ---
header_style = ParagraphStyle(
    name='Header',
    fontName='Helvetica-Bold',
    fontSize=10,
    alignment=TA_LEFT,
    leading=11
)
value_style = ParagraphStyle(
    name='Value',
    fontName='Helvetica',
    fontSize=13,
    alignment=TA_CENTER,
    leading=14
)
bold_value_style = ParagraphStyle(
    name='BoldValue',
    fontName='Helvetica-Bold', # Use bold font
    fontSize=16,
    alignment=TA_CENTER,
    leading=17
)
bold_centered_value_style = ParagraphStyle(
    name='BoldCenteredValue',
    fontName='Helvetica-Bold',
    fontSize=16,
    alignment=TA_CENTER,
    leading=20
)

def format_description_v1(desc):
    """Format description text with dynamic font sizing and alignment."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    
    desc_length = len(desc)
    
    if desc_length <= 90:
        font_size = 9
    else:
        font_size = 8
        desc = desc[:100] + "..." if len(desc) > 100 else desc
    
    desc_style_v1 = ParagraphStyle(
        name='Description_v1',
        fontName='Helvetica',
        fontSize=font_size,
        alignment=TA_CENTER,
        leading=font_size + 2,
        spaceBefore=1,
        spaceAfter=1
    )
    
    return Paragraph(desc, desc_style_v1)

def find_column(df, keywords):
    """Find a column in the DataFrame that matches any of the keywords (case-insensitive)"""
    cols = df.columns.tolist()
    for keyword in keywords:
        for col in cols:
            if isinstance(col, str) and keyword.upper() in col.upper():
                return col
    return None

def generate_final_labels(df, progress_bar=None, status_container=None):
    """Generate final hybrid labels using a nested table structure for flexible column widths."""
    
    # Identify columns from the uploaded file
    model_col = find_column(df, ['MODEL'])
    structure_col = find_column(df, ['STRUCTURE'])
    station_no_col = find_column(df, ['STATION NO', 'STATION_NO', 'STATION'])
    fixture_location_col = find_column(df, ['FIXTURE LOCATION', 'FIXTURE_LOCATION', 'LOCATION'])
    part_no_col = find_column(df, ['PART NO', 'PARTNO', 'PART_NO', 'PART#'])
    qty_veh_col = find_column(df, ['QTY/VEH', 'QTY_VEH', 'QTY/BIN', 'QTY'])
    desc_col = find_column(df, ['PART DESC', 'PART_DESCRIPTION', 'DESC', 'DESCRIPTION', 'PART NAME'])

    if status_container:
        status_container.write("**Attempting to map columns from your file...**")
        # (Status messages remain the same)

    # Setup PDF document
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_path = temp_file.name
    temp_file.close()
    doc = SimpleDocTemplate(temp_path, pagesize=STICKER_PAGESIZE,
                          topMargin=1*cm, bottomMargin=1*cm,
                          leftMargin=1*cm, rightMargin=1*cm)

    all_elements = []
    total_rows = len(df)

    for index, row in df.iterrows():
        if progress_bar: progress_bar.progress((index + 1) / total_rows)
        if status_container: status_container.write(f"Creating label {index + 1} of {total_rows}")

        # Extract data from the current row
        model = str(row.get(model_col, ""))
        structure = str(row.get(structure_col, ""))
        station_no = str(row.get(station_no_col, ""))
        fixture_location = str(row.get(fixture_location_col, ""))
        part_no = str(row.get(part_no_col, ""))
        qty_veh = str(row.get(qty_veh_col, ""))
        part_desc = str(row.get(desc_col, ""))
        
        # --- MODIFICATION START: Updated Table Styling ---

        # --- TABLE FOR ROW 1 (Model, Structure, etc.) ---
        data_r1 = [[Paragraph(model, bold_centered_value_style), Paragraph(structure, bold_centered_value_style), Paragraph(station_no, bold_centered_value_style), Paragraph(fixture_location, bold_centered_value_style)]]
        col_widths_r1 = [CONTENT_BOX_WIDTH * 0.25, CONTENT_BOX_WIDTH * 0.25, CONTENT_BOX_WIDTH * 0.25, CONTENT_BOX_WIDTH * 0.25]
        table_r1 = Table(data_r1, colWidths=col_widths_r1, rowHeights=ROW_HEIGHTS[0])
        # Use 'GRID' to draw all lines for this table
        table_r1.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))

        # --- TABLE FOR ROW 2 (Part No, Qty/Veh) ---
        data_r2 = [[Paragraph('<b>PART NO</b>', header_style), Paragraph(part_no, bold_value_style), Paragraph('<b>QTY/\nVEH</b>', header_style), Paragraph(qty_veh, value_style)]]
        col_widths_r2 = [CONTENT_BOX_WIDTH * 0.20, CONTENT_BOX_WIDTH * 0.43, CONTENT_BOX_WIDTH * 0.15, CONTENT_BOX_WIDTH * 0.20]
        table_r2 = Table(data_r2, colWidths=col_widths_r2, rowHeights=ROW_HEIGHTS[1])
        # Use 'GRID' to draw all lines for this table
        table_r2.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))

        # --- TABLE FOR ROW 3 (Part Name) ---
        data_r3 = [[Paragraph('<b>PART NAME</b>', header_style), format_description_v1(part_desc)]]
        col_widths_r3 = [CONTENT_BOX_WIDTH * 0.20, CONTENT_BOX_WIDTH * 0.80] 
        table_r3 = Table(data_r3, colWidths=col_widths_r3, rowHeights=ROW_HEIGHTS[2])
        # Use 'GRID' to draw all lines for this table
        table_r3.setStyle(TableStyle([('GRID', (0, 0), (-1, -1), 1, colors.black), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))

        # --- CONTAINER TABLE (to hold the 3 rows together) ---
        container_data = [[table_r1], [table_r2], [table_r3]]
        container_table = Table(container_data, colWidths=[CONTENT_BOX_WIDTH])
        # Remove all padding from the container so the inner tables touch
        container_table.setStyle(TableStyle([
            ('LEFTPADDING', (0,0), (-1,-1), 0),
            ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ]))
        
        all_elements.append(container_table)
        
        # --- MODIFICATION END ---

        if index < total_rows - 1:
            all_elements.append(PageBreak())

    # Build the final PDF document
    try:
        doc.build(all_elements)
        if status_container: status_container.success("PDF generated successfully!")
        return temp_path
    except Exception as e:
        if status_container: status_container.error(f"Error building PDF: {e}")
        return None

def main():
    st.set_page_config(page_title="CESL Label Generator", page_icon="🏷️", layout="wide")
    
    st.title("🏷️ CESL Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    st.info("""
    **Label Logic:**
    - Each row now has an independent column layout for maximum flexibility.
    - **PART NAME** value is center-aligned with dynamic font size.
    - **PART NO** value is bold with a larger font size for emphasis.
    """)
    
    st.subheader("📋 Reference Data Format")
    sample_data = {
        'MODEL': ['3WC', '3WM'],
        'STRUCTURE': ['S-A', 'S-B'],
        'STATION NO': ['STN-1', 'STN-2'],
        'FIXTURE LOCATION': ['9M CSEL', '8L BSEAT'],
        'PART NO': ['08-DRA-14-02', 'P0012124-07'],
        'QTY/VEH': [2, 1],
        'PART DESCRIPTION': ['BELLOW ASSY. WITH RETAINING CLIP', 'GUARD RING (hirkesh)']
    }
    st.dataframe(pd.DataFrame(sample_data), use_container_width=True)
    st.markdown("---")
        
    with st.sidebar:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader("Choose Excel or CSV file", type=['xlsx', 'xls', 'csv'])
        if uploaded_file: st.success(f"File uploaded: {uploaded_file.name}")
    
    if uploaded_file is not None:
        try:
            df = pd.read_csv(uploaded_file) if uploaded_file.name.lower().endswith('.csv') else pd.read_excel(uploaded_file)
            st.subheader("📊 Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            if st.button("🚀 Generate Labels", type="primary", use_container_width=True):
                with st.spinner("Generating your labels..."):
                    progress_bar = st.progress(0)
                    status_container = st.empty()
                    pdf_path = generate_final_labels(df, progress_bar, status_container)
                    
                    if pdf_path:
                        with open(pdf_path, 'rb') as f:
                            pdf_data = f.read()
                        os.unlink(pdf_path)
                        st.download_button(
                            label="📥 Download PDF",
                            data=pdf_data,
                            file_name=f"{os.path.splitext(uploaded_file.name)[0]}_labels.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                        st.success("✅ Labels generated successfully!")
                    else:
                        st.error("❌ Failed to generate labels.")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
