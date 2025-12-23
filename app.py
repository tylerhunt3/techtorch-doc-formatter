import streamlit as st
from docx import Document as ReadDocument
from docx.shared import Pt, Inches, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import re
from datetime import datetime

# --- CONFIGURATION ---
# TechTorch Color Palette (RGB values)
COLORS = {
    "HEADING1": RGBColor(0x1F, 0x4E, 0x79),      # Dark blue
    "HEADING2": RGBColor(0x2E, 0x75, 0xB6),      # Medium blue
    "HEADING3": RGBColor(0x40, 0x40, 0x40),      # Dark gray
    "BODY": RGBColor(0x00, 0x00, 0x00),          # Black
    "SECONDARY": RGBColor(0x66, 0x66, 0x66),     # Gray
    "CODE_TEXT": RGBColor(0x2E, 0x2E, 0x2E),     # Near-black
    "WHITE": RGBColor(0xFF, 0xFF, 0xFF),         # White
}

# Font sizes
SIZES = {
    "TITLE": Pt(24),
    "SUBTITLE": Pt(14),
    "HEADING1": Pt(12),
    "HEADING2": Pt(11),
    "HEADING3": Pt(10),
    "BODY": Pt(9),
    "CODE": Pt(8),
    "HEADER_FOOTER": Pt(8),
}

# --- HELPER FUNCTIONS ---

def extract_content_from_docx(uploaded_file):
    """Extract text content from uploaded Word document."""
    doc = ReadDocument(uploaded_file)
    content = []
    
    for para in doc.paragraphs:
        text = para.text.strip()
        if text:
            # Try to detect heading level based on style or formatting
            style_name = para.style.name.lower() if para.style else ""
            
            if "heading 1" in style_name or (para.runs and para.runs[0].bold and len(text) < 100):
                # Check if it looks like a section header
                if re.match(r'^\d+\.?\s+\w+', text) or len(text) < 80:
                    content.append({"type": "heading1", "text": text})
                else:
                    content.append({"type": "paragraph", "text": text})
            elif "heading 2" in style_name:
                content.append({"type": "heading2", "text": text})
            elif "heading 3" in style_name:
                content.append({"type": "heading3", "text": text})
            elif "title" in style_name:
                content.append({"type": "title", "text": text})
            else:
                # Check for bullet points
                if text.startswith(("-", "•", "*")) or para.style.name == "List Paragraph":
                    clean_text = re.sub(r'^[-•*]\s*', '', text)
                    content.append({"type": "bullet", "text": clean_text})
                # Check for numbered items
                elif re.match(r'^\d+\.\s+', text) and len(text) < 100:
                    content.append({"type": "heading1", "text": text})
                elif re.match(r'^\d+\.\d+\s+', text):
                    content.append({"type": "heading2", "text": text})
                elif re.match(r'^\d+\.\d+\.\d+\s+', text):
                    content.append({"type": "heading3", "text": text})
                else:
                    content.append({"type": "paragraph", "text": text})
    
    # Also extract tables
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
        if table_data:
            content.append({"type": "table", "data": table_data})
    
    return content


def set_cell_shading(cell, color_hex):
    """Set background color for a table cell."""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color_hex)
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_borders(cell, color="CCCCCC", size="4"):
    """Set borders for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    
    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), size)
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    
    tcPr.append(tcBorders)


def create_formatted_document(content, doc_title, organization="TechTorch Inc."):
    """Create a new formatted Word document."""
    doc = ReadDocument()
    
    # Set default font for document
    style = doc.styles['Normal']
    style.font.name = 'Aptos'
    style.font.size = SIZES["BODY"]
    style.font.color.rgb = COLORS["BODY"]
    
    # Set up Heading styles
    for i, (style_name, size, color) in enumerate([
        ('Heading 1', SIZES["HEADING1"], COLORS["HEADING1"]),
        ('Heading 2', SIZES["HEADING2"], COLORS["HEADING2"]),
        ('Heading 3', SIZES["HEADING3"], COLORS["HEADING3"]),
    ]):
        heading_style = doc.styles[style_name]
        heading_style.font.name = 'Aptos'
        heading_style.font.size = size
        heading_style.font.color.rgb = color
        heading_style.font.bold = True
    
    # --- TITLE PAGE ---
    # Add spacing at top
    for _ in range(4):
        doc.add_paragraph()
    
    # Document title
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(doc_title)
    title_run.font.name = 'Aptos'
    title_run.font.size = SIZES["TITLE"]
    title_run.font.bold = True
    title_run.font.color.rgb = COLORS["BODY"]
    
    # Organization
    org_para = doc.add_paragraph()
    org_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    org_run = org_para.add_run(organization)
    org_run.font.name = 'Aptos'
    org_run.font.size = Pt(12)
    org_run.font.italic = True
    org_run.font.color.rgb = COLORS["BODY"]
    
    # Date
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_para.add_run(f"As of {datetime.now().strftime('%B %Y')}")
    date_run.font.name = 'Aptos'
    date_run.font.size = Pt(10)
    date_run.font.color.rgb = COLORS["BODY"]
    
    # Version
    version_para = doc.add_paragraph()
    version_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    version_run = version_para.add_run("Version 1.0")
    version_run.font.name = 'Aptos'
    version_run.font.size = SIZES["BODY"]
    version_run.font.color.rgb = COLORS["SECONDARY"]
    
    # Page break after title page
    doc.add_page_break()
    
    # --- MAIN CONTENT ---
    for item in content:
        if item["type"] == "title":
            # Skip title in body since we added it to title page
            continue
            
        elif item["type"] == "heading1":
            para = doc.add_paragraph(item["text"], style='Heading 1')
            para.paragraph_format.space_before = Twips(300)
            para.paragraph_format.space_after = Twips(100)
            
        elif item["type"] == "heading2":
            para = doc.add_paragraph(item["text"], style='Heading 2')
            para.paragraph_format.space_before = Twips(200)
            para.paragraph_format.space_after = Twips(80)
            
        elif item["type"] == "heading3":
            para = doc.add_paragraph(item["text"], style='Heading 3')
            para.paragraph_format.space_before = Twips(160)
            para.paragraph_format.space_after = Twips(60)
            
        elif item["type"] == "paragraph":
            para = doc.add_paragraph()
            para.paragraph_format.space_after = Twips(160)
            run = para.add_run(item["text"])
            run.font.name = 'Aptos'
            run.font.size = SIZES["BODY"]
            run.font.color.rgb = COLORS["BODY"]
            
        elif item["type"] == "bullet":
            para = doc.add_paragraph(style='List Bullet')
            para.paragraph_format.left_indent = Twips(720)
            run = para.add_run(item["text"])
            run.font.name = 'Aptos'
            run.font.size = SIZES["BODY"]
            run.font.color.rgb = COLORS["BODY"]
            
        elif item["type"] == "table":
            table_data = item["data"]
            if len(table_data) > 0:
                num_cols = len(table_data[0])
                table = doc.add_table(rows=len(table_data), cols=num_cols)
                table.alignment = WD_TABLE_ALIGNMENT.CENTER
                
                for row_idx, row_data in enumerate(table_data):
                    row = table.rows[row_idx]
                    for col_idx, cell_text in enumerate(row_data):
                        cell = row.cells[col_idx]
                        cell.text = cell_text
                        
                        # Format cell
                        for para in cell.paragraphs:
                            for run in para.runs:
                                run.font.name = 'Aptos'
                                run.font.size = SIZES["BODY"]
                        
                        # Set borders
                        set_cell_borders(cell)
                        
                        # Header row styling
                        if row_idx == 0:
                            set_cell_shading(cell, "1F4E79")
                            for para in cell.paragraphs:
                                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                for run in para.runs:
                                    run.font.bold = True
                                    run.font.color.rgb = COLORS["WHITE"]
                
                doc.add_paragraph()  # Space after table
    
    # --- END OF DOCUMENT ---
    end_para = doc.add_paragraph()
    end_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    end_para.paragraph_format.space_before = Twips(400)
    end_run = end_para.add_run("End of Document")
    end_run.font.name = 'Aptos'
    end_run.font.size = SIZES["BODY"]
    end_run.font.italic = True
    end_run.font.color.rgb = COLORS["SECONDARY"]
    
    return doc


def save_doc_to_bytes(doc):
    """Save document to bytes for download."""
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


# --- STREAMLIT APP ---

st.set_page_config(
    page_title="TechTorch Document Formatter",
    page_icon="📄",
    layout="centered"
)

# Header
st.title("📄 TechTorch Document Formatter")
st.markdown("Upload a Word document to automatically apply TechTorch formatting standards.")

st.divider()

# File upload
uploaded_file = st.file_uploader(
    "Upload your Word document",
    type=["docx"],
    help="Upload a .docx file to format"
)

# Document title input
doc_title = st.text_input(
    "Document Title",
    placeholder="Enter the document title...",
    help="This will appear on the title page"
)

# Organization input
organization = st.text_input(
    "Organization",
    value="TechTorch Inc.",
    help="Organization name for the title page"
)

st.divider()

# Process button
if uploaded_file is not None and doc_title:
    if st.button("Format Document", type="primary"):
        with st.spinner("Processing your document..."):
            try:
                # Extract content
                content = extract_content_from_docx(uploaded_file)
                
                # Create formatted document
                formatted_doc = create_formatted_document(
                    content, 
                    doc_title, 
                    organization
                )
                
                # Save to bytes
                doc_bytes = save_doc_to_bytes(formatted_doc)
                
                # Success message
                st.success("Document formatted successfully!")
                
                # Download button
                clean_title = re.sub(r'[^\w\s-]', '', doc_title).replace(' ', '_')
                filename = f"{clean_title}_Formatted.docx"
                
                st.download_button(
                    label="Download Formatted Document",
                    data=doc_bytes,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Error processing document: {str(e)}")

elif uploaded_file is None:
    st.info("Please upload a Word document to get started.")
elif not doc_title:
    st.info("Please enter a document title.")

# Footer
st.divider()
st.markdown(
    "<div style='text-align: center; color: #666666; font-size: 12px;'>"
    "TechTorch Documentation Formatting Tool v1.0"
    "</div>",
    unsafe_allow_html=True
)
