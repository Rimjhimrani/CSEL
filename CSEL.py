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

# Define sticker dimensions
STICKER_WIDTH = 10 * cm
STICKER_HEIGHT = 15 * cm
STICKER_PAGESIZE = (STICKER_WIDTH, STICKER_HEIGHT)

# Define content box dimensions
CONTENT_BOX_WIDTH = 8 * cm
CONTENT_BOX_HEIGHT = 3 * cm

# Define paragraph styles for headers and values
header_style = ParagraphStyle(
    name='Header',
    fontName='Helvetica-Bold', # Bold for headers
    fontSize=9,
    alignment=TA_LEFT,
    leading=11
)
value_style = ParagraphStyle(
    name='Value',
    fontName='Helvetica', # Regular for values
    fontSize=9,
    alignment=TA_LEFT,
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

def generate_labels_with_headers(df, progress_bar=None, status_container=None):
    """Generate sticker labels with both headers and values, based on the final image layout"""
    
    # Identify columns from the uploaded file (case-insensitive)
    fixture_location_col = find_column(df, ['FIXTURE LOCATION', 'FIXTURE_LOCATION', 'LOCATION'])
    model_col = find_column(df, ['MODEL'])
    part_no_col = find_column(df, ['PART NO', 'PARTNO', 'PART_NO', 'PART#'])
    qty_veh_col = find_column(df, ['QTY/VEH', 'QTY_VEH', 'QTY/BIN', 'QTY'])
    desc_col = find_column(df, ['PART DESC', 'PART_DESCRIPTION', 'DESC', 'DESCRIPTION'])

    if status_container:
        status_container.write("**Attempting to map these columns:**")
        status_container.write(f"- Fixture Location: `{fixture_location_col if fixture_location_col else 'Not Found'}`")
        status_container.write(f"- Model: `{model_col if model_col else 'Not Found'}`")
        status_container.write(f"- Part No: `{part_no_col if part_no_col else 'Not Found'}`")
        status_container.write(f"- Qty/Veh: `{qty_veh_col if qty_veh_col else 'Not Found'}`")
        status_container.write(f"- Part Description: `{desc_col if desc_col else 'Not Found'}`")

    # Create a temporary file to save the PDF
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
    temp_path = temp_file.name
    temp_file.close()

    # Create the PDF document
    doc = SimpleDocTemplate(temp_path, pagesize=STICKER_PAGESIZE,
                          topMargin=1*cm, bottomMargin=1*cm,
                          leftMargin=1*cm, rightMargin=1*cm)

    all_elements = []
    total_rows = len(df)

    # Process each row in the file to create one label
    for index, row in df.iterrows():
        if progress_bar:
            progress_bar.progress((index + 1) / total_rows)
        if status_container:
            status_container.write(f"Creating label {index + 1} of {total_rows}")

        # Extract data, using a blank space if a column isn't found
        fixture_location = str(row[fixture_location_col]) if fixture_location_col and fixture_location_col in row and pd.notna(row[fixture_location_col]) else ""
        model = str(row[model_col]) if model_col and model_col in row and pd.notna(row[model_col]) else ""
        part_no = str(row[part_no_col]) if part_no_col and part_no_col in row and pd.notna(row[part_no_col]) else ""
        qty_veh = str(row[qty_veh_col]) if qty_veh_col and qty_veh_col in row and pd.notna(row[qty_veh_col]) else ""
        part_desc = str(row[desc_col]) if desc_col and desc_col in row and pd.notna(row[desc_col]) else ""

        # Structure the data for the table with headers and values
        # We use a nested table approach for clean separation
        data = [
            # Row 1: Fixture Location and Model
            [Paragraph('<b>Fixture Location</b>', header_style), Paragraph(fixture_location, value_style), Paragraph('<b>MODEL</b>', header_style), Paragraph(model, value_style)],
            # Row 2: Part No and Qty/Veh
            [Paragraph('<b>PART NO</b>', header_style), Paragraph(part_no, value_style), Paragraph('<b>QTY/VEH</b>', header_style), Paragraph(qty_veh, value_style)],
            # Row 3: Part Description Header
            [Paragraph('<b>PART DESCRIPTION</b>', header_style), '', '', ''],
            # Row 4: Part Description Value
            [Paragraph(part_desc, value_style), '', '', '']
        ]
        
        # Define column widths to match the layout
        col_widths = [CONTENT_BOX_WIDTH * 0.28, CONTENT_BOX_WIDTH * 0.22, CONTENT_BOX_WIDTH * 0.22, CONTENT_BOX_WIDTH * 0.28]
        table = Table(data, colWidths=col_widths)

        # Apply styles for grid, merged cells, and padding
        style = TableStyle([
            # Grid lines for all cells
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            
            # Merge cells for the Part Description header and value
            ('SPAN', (0, 2), (3, 2)),  # Span for "PART DESCRIPTION" header
            ('SPAN', (0, 3), (3, 3)),  # Span for the description value
            
            # Alignment and padding
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ])
        
        table.setStyle(style)
        all_elements.append(table)

        # Add a page break after each label (except the last one)
        if index < total_rows - 1:
            all_elements.append(PageBreak())

    # Build the final PDF
    try:
        doc.build(all_elements)
        if status_container:
            status_container.success("PDF generated successfully!")
        return temp_path
    except Exception as e:
        if status_container:
            status_container.error(f"Error building PDF: {e}")
        return None

def main():
    st.set_page_config(
        page_title="Label Generator",
        page_icon="🏷️",
        layout="wide"
    )
    
    st.title("🏷️ Label Generator (With Headers)")
    st.markdown(
        "<p style='font-size:18px; font-style:italic; margin-top:-10px; text-align:left;'>"
        "Designed and Developed by Agilomatrix</p>",
        unsafe_allow_html=True
    )

    st.markdown("---")

    st.info("👈 Upload an Excel or CSV file to generate labels with headers and values.")
    
    # Show the expected data format
    st.subheader("📋 Reference Data Format")
    st.markdown("Your file should contain columns with headers like these (case-insensitive):")
    sample_data = {
        'FIXTURE LOCATION': ['9M CSEL', '8L BSEAT', '7K FRAME'],
        'MODEL': ['3WC', '3WM', '3WS'],
        'PART NO': ['08-DRA-14-02', 'P0012124-07', 'P0012126-07'],
        'QTY/VEH': [2, 1, 4],
        'PART DESCRIPTION': ['BELLOW ASSY. WITH RETAINING CLIP', 'GUARD RING (hirkesh)', 'GUARD RING SEAL (hirkesh)']
    }
    
    sample_df = pd.DataFrame(sample_data)
    st.dataframe(sample_df, use_container_width=True)

    st.markdown("---")
        
    # Sidebar for file upload
    with st.sidebar:
        st.header("📁 File Upload")
        uploaded_file = st.file_uploader(
            "Choose Excel or CSV file",
            type=['xlsx', 'xls', 'csv'],
            help="Upload the file containing data for your labels."
        )
        
        if uploaded_file:
            st.success(f"File uploaded: {uploaded_file.name}")
    
    # Main content area
    if uploaded_file is not None:
        try:
            if uploaded_file.name.lower().endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.subheader("📊 Data Preview")
            st.dataframe(df.head(10), use_container_width=True)
            
            if st.button("🚀 Generate Labels", type="primary", use_container_width=True):
                with st.spinner("Generating your labels..."):
                    progress_bar = st.progress(0)
                    status_container = st.empty()
                    
                    pdf_path = generate_labels_with_headers(df, progress_bar, status_container)
                    
                    if pdf_path:
                        with open(pdf_path, 'rb') as pdf_file:
                            pdf_data = pdf_file.read()
                        
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
            st.error(f"An error occurred while processing the file: {str(e)}")

if __name__ == "__main__":
    main()
