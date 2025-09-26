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
# --- MODIFICATION START ---
# Added a new style specifically for the bold Part No value
bold_value_style = ParagraphStyle(
    name='BoldValue',
    fontName='Helvetica-Bold', # Use bold font
    fontSize=14,
    alignment=TA_CENTER,
    leading=14
)
# --- MODIFICATION END ---
bold_centered_value_style = ParagraphStyle(
    name='BoldCenteredValue',
    fontName='Helvetica-Bold',
    fontSize=16,
    alignment=TA_CENTER,
    leading=20
)

# --- NEW FUNCTION PROVIDED BY USER ---
def format_description_v1(desc):
    """Format description text with dynamic font sizing based on length for v1."""
    if not desc or not isinstance(desc, str):
        desc = str(desc)
    
    # Dynamic font sizing based on description length
    desc_length = len(desc)
    
    if desc_length <= 30:
        font_size = 9
    elif desc_length <= 50:
        font_size = 9
    elif desc_length <= 70:
        font_size = 9
    elif desc_length <= 90:
        font_size = 9
    else:
        font_size = 8
        # Truncate very long descriptions to prevent overflow
        desc = desc[:100] + "..." if len(desc) > 100 else desc
    
    # Create a custom style for this description
    desc_style_v1 = ParagraphStyle(
        name='Description_v1',
        fontName='Helvetica',
        fontSize=font_size,
        alignment=TA_LEFT,
        leading=font_size + 2, # Adjust leading based on font size
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
    """Generate final hybrid labels with dynamic description formatting."""
    
    # Identify columns from the uploaded file
    model_col = find_column(df, ['MODEL'])
    structure_col = find_column(df, ['STRUCTURE'])
    station_no_col = find_column(df, ['STATION NO', 'STATION_NO', 'STATION'])
    fixture_location_col = find_column(df, ['FIXTURE LOCATION', 'FIXTURE_LOCATION', 'LOCATION'])
    part_no_col = find_column(df, ['PART NO', 'PARTNO', 'PART_NO', 'PART#'])
    qty_veh_col = find_column(df, ['QTY/VEH', 'QTY_VEH', 'QTY/BIN', 'QTY'])
    desc_col = find_column(df, ['PART DESC', 'PART_DESCRIPTION', 'DESC', 'DESCRIPTION', 'PART NAME'])

    if status_container:
        status_container.write("**Attempting to map columns from your file:**")
        status_container.write(f"- For Model: `{model_col if model_col else 'Not Found'}`")
        status_container.write(f"- For Structure: `{structure_col if structure_col else 'Not Found'}`")
        status_container.write(f"- For Station No: `{station_no_col if station_no_col else 'Not Found'}`")
        status_container.write(f"- For Fixture Location: `{fixture_location_col if fixture_location_col else 'Not Found'}`")
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
        model = str(row.get(model_col, ""))
        structure = str(row.get(structure_col, ""))
        station_no = str(row.get(station_no_col, ""))
        fixture_location = str(row.get(fixture_location_col, ""))
        part_no = str(row.get(part_no_col, ""))
        qty_veh = str(row.get(qty_veh_col, ""))
        part_desc = str(row.get(desc_col, ""))
        
        # Structure the data for the table
        data = [
            # Row 1: Four columns for Model, Structure, Station No, and Fixture Location
            [Paragraph(model, bold_centered_value_style), Paragraph(structure, bold_centered_value_style), Paragraph(station_no, bold_centered_value_style), Paragraph(fixture_location, bold_centered_value_style)],
            # --- MODIFICATION START ---
            # Row 2: Headers and values. Part No value now uses the new bold style.
            [Paragraph('<b>PART NO</b>', header_style), Paragraph(part_no, bold_value_style), Paragraph('<b>QTY/VEH</b>', header_style), Paragraph(qty_veh, value_style)],
            # --- MODIFICATION END ---
            # Row 3: Header and value for Part Name
            [Paragraph('<b>PART NAME</b>', header_style), format_description_v1(part_desc), '', '']
        ]
        
        # Adjusted column widths for a 4-column top row
        col_widths = [CONTENT_BOX_WIDTH * 0.20, CONTENT_BOX_WIDTH * 0.43, CONTENT_BOX_WIDTH * 0.15, CONTENT_BOX_WIDTH * 0.20]
        table = Table(data, colWidths=col_widths, rowHeights=ROW_HEIGHTS)

        # Apply styles for grid, merged cells, and alignment
        style = TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            
            # Cell Merging
            # Merge for the Part Name value (spans 3 columns)
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
    - The top row now displays **Model**, **Structure**, **Station No**, and **Fixture Location**. All are bold with a larger font size.
    - **PART NAME** value continues to have dynamic font size and automatic text wrapping.
    - **PART NO** value is now bold for emphasis.
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
