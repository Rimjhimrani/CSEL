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
# Adjust the values in this list to change the height of each row in the PDF.
# - ROW_HEIGHTS[0]: Height for the 'Fixture Location / Model' row.
# - ROW_HEIGHTS[1]: Height for the 'Part No / Qty/Veh' row.
# - ROW_HEIGHTS[2]: Height for the 'Part Name' row.
ROW_HEIGHTS = [0.8*cm, 0.8*cm, 1.4*cm]

# --- STICKER AND CONTENT DIMENSIONS ---
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)
CONTENT_BOX_WIDTH = 8 * cm

# --- PARAGRAPH STYLES ---
header_style = ParagraphStyle(
    name='Header',
    fontName='Helvetica-Bold',
    fontSize=9,
    alignment=TA_LEFT,
    leading=11
)
value_style = ParagraphStyle(
    name='Value',
    fontName='Helvetica',
    fontSize=9,
    alignment=TA_LEFT,
    leading=11
)
centered_value_style = ParagraphStyle(
    name='CenteredValue',
    fontName='Helvetica',
    fontSize=9,
    alignment=TA_CENTER,
    leading=11
)

def find_column(df, keywords):
    """Find a column in the DataFrame that matches any of the keywords (case-insensitive)"""
    cols = df.columns.tolist()
    for keyword in keywords:
        for col in cols:
            if isinstance(col, str) and keyword.upper() in col.upper():
                return col
    return None

def generate_final_labels(df, progress_bar=None, status_container=None):
    """Generate final hybrid labels with adjustable row heights and side-by-side layout."""
    
    # Identify columns from the uploaded file
    fixture_location_col = find_column(df, ['FIXTURE LOCATION', 'FIXTURE_LOCATION', 'LOCATION'])
    model_col = find_column(df, ['MODEL'])
    part_no_col = find_column(df, ['PART NO', 'PARTNO', 'PART_NO', 'PART#'])
    qty_veh_col = find_column(df, ['QTY/VEH', 'QTY_VEH', 'QTY/BIN', 'QTY'])
    desc_col = find_column(df, ['PART DESC', 'PART_DESCRIPTION', 'DESC', 'DESCRIPTION', 'PART NAME'])

    if status_container:
        status_container.write("**Attempting to map columns from your file:**")
        status_container.write(f"- For Fixture Location: `{fixture_location_col if fixture_location_col else 'Not Found'}`")
        status_container.write(f"- For Model: `{model_col if model_col else 'Not Found'}`")
        status_container.write(f"- For Part No: `{part_no_col if part_no_col else 'Not Found'}`")
        status_container.write(f"- For Qty/Veh: `{qty_veh_col if qty_veh_col else 'Not Found'}`")
        status_container.write(f"- For Part Name (looks for 'Part Description'): `{desc_col if desc_col else 'Not Found'}`")

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
        fixture_location = str(row[fixture_location_col]) if fixture_location_col and fixture_location_col in row and pd.notna(row[fixture_location_col]) else ""
        model = str(row[model_col]) if model_col and model_col in row and pd.notna(row[model_col]) else ""
        part_no = str(row[part_no_col]) if part_no_col and part_no_col in row and pd.notna(row[part_no_col]) else ""
        qty_veh = str(row[qty_veh_col]) if qty_veh_col and qty_veh_col in row and pd.notna(row[qty_veh_col]) else ""
        part_desc = str(row[desc_col]) if desc_col and desc_col in row and pd.notna(row[desc_col]) else ""

        # Structure the data for the table
        data = [
            # Row 1: Values only for Fixture Location and Model
            [Paragraph(fixture_location, centered_value_style), '', Paragraph(model, centered_value_style), ''],
            # Row 2: Headers and values for Part No and Qty/Veh
            [Paragraph('<b>PART NO</b>', header_style), Paragraph(part_no, value_style), Paragraph('<b>QTY/VEH</b>', header_style), Paragraph(qty_veh, value_style)],
            # Row 3: Header and value for Part Name
            [Paragraph('<b>PART NAME</b>', header_style), Paragraph(part_desc, value_style), '', '']
        ]
        
        # Define column widths
        col_widths = [CONTENT_BOX_WIDTH * 0.25, CONTENT_BOX_WIDTH * 0.25, CONTENT_BOX_WIDTH * 0.25, CONTENT_BOX_WIDTH * 0.25]
        # Create the table with manually adjustable row heights
        table = Table(data, colWidths=col_widths, rowHeights=ROW_HEIGHTS)

        # Apply styles for grid, merged cells, and alignment
        style = TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            
            # --- Cell Merging (SPAN) ---
            # Row 1 (Fixture Location and Model values)
            ('SPAN', (0, 0), (1, 0)),
            ('SPAN', (2, 0), (3, 0)),
            # Row 3 (Part Name value)
            ('SPAN', (1, 2), (3, 2)),
        ])
        
        table.setStyle(style)
        all_elements.append(table)

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
    st.set_page_config(page_title="Final Label Generator", page_icon="🏷️", layout="wide")
    
    st.title("🏷️ Final Label Generator")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )
    st.markdown("---")

    st.info("""
    **Label Logic:**
    - **PART NO**, **QTY/VEH**, and **PART NAME** will show a header and a value.
    - **Fixture Location** and **MODEL** will show only their values.
    - The header for 'Part Description' from your file will be displayed as **'PART NAME'**.
    """)
    
    st.subheader("📋 Reference Data Format")
    st.markdown("Your file should contain columns with headers like these (case-insensitive):")
    sample_data = {
        'FIXTURE LOCATION': ['9M CSEL', '8L BSEAT'],
        'MODEL': ['3WC', '3WM'],
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
