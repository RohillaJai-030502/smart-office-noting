from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
import shutil
from datetime import datetime

OUTPUT_DIR = "generated_notices"

def replace_in_paragraph(para, placeholders):
    for key, value in placeholders.items():
        if key in para.text:
            para.text = para.text.replace(key, str(value))

def set_table_borders(table):
    """Utility function to add borders to the dynamic word table"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def generate_tbrl_noting(data):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    master_path = os.path.join("masters", "TBRL_Noting_Master.docx")
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    group = data.get('group_name', 'General').replace("/", "-")
    filename = f"TBRL_Noting_{group}_{timestamp}.docx"
    output_path = os.path.join(OUTPUT_DIR, filename)
    
    # Copy the master template to start working on it
    shutil.copy2(master_path, output_path)

    doc = Document(output_path)
    
    # 1. Date Math calculation
    start_date = datetime.strptime(data['start_date'], '%Y-%m-%d')
    end_date = datetime.strptime(data['end_date'], '%Y-%m-%d')
    days_count = (end_date - start_date).days + 1
    
    # 2. Grammar Engine (Hindi Singular/Plural)
    num_nominees = len(data.get('nominees', []))
    if num_nominees <= 1:
        off_word, ka_ke, hua_hue, nam_word = "अधिकारी", "का", "हुआ", "नामांकन"
    else:
        off_word, ka_ke, hua_hue, nam_word = "अधिकारियों", "के", "हुए", "नामांकनों"

    # 3. Placeholders Dictionary
    placeholders = {
        "{{REF_NO}}": data.get("ref_no", ""),
        "{{CURRENT_YEAR}}": str(datetime.now().year),
        "{{SUBJECT_HINDI}}": data.get("subject_hindi", ""),
        "{{SUBJECT_ENGLISH}}": data.get("subject_english", ""),
        "{{COURSE_TYPE}}": data.get("course_type", ""),
        "{{REFERENCE_TEXT}}": data.get("reference_text", ""),
        "{{REF_DATE}}": data.get("ref_date", ""),
        "{{START_DATE}}": start_date.strftime('%d %B %Y'),
        "{{END_DATE}}": end_date.strftime('%d %B %Y'),
        "{{DAYS_COUNT}}": str(days_count),
        "{{ORG_INSTITUTE}}": data.get("org_institute", ""),
        "{{GROUP_NAME}}": f"{data.get('group_name', '')} Group",
        
        # Injecting Grammar
        "{{OFFICIAL_WORD}}": off_word,
        "{{KA_KE}}": ka_ke,
        "{{HUA_HUE}}": hua_hue,
        "{{NAMANKAN_WORD}}": nam_word,
        
        # Signatories
        "{{SIGNATORY_1_NAME}}": data.get("sig1_name", ""),
        "{{SIGNATORY_1_DESIG}}": data.get("sig1_desig", ""),
        "{{SIGNATORY_2_NAME}}": data.get("sig2_name", ""),
        "{{SIGNATORY_2_DESIG}}": data.get("sig2_desig", ""),
    }

    # Process all paragraphs
    table_index = None
    for i, para in enumerate(doc.paragraphs):
        if "{{COMPLEX_DYNAMIC_TABLE}}" in para.text:
            table_index = i
            para.text = "" # Clear the placeholder text
        else:
            replace_in_paragraph(para, placeholders)

    # 4. Generate Dynamic Table
    if table_index is not None and data.get('columns') and data.get('nominees'):
        columns = data['columns']
        nominees = data['nominees']
        
        table = doc.add_table(rows=len(nominees) + 1, cols=len(columns))
        set_table_borders(table)
        
        # Add Headers
        for col_idx, col_name in enumerate(columns):
            cell = table.cell(0, col_idx)
            cell.text = col_name
            cell.paragraphs[0].runs[0].font.bold = True
            
        # Add Data Rows
        for row_idx, nominee_data in enumerate(nominees, start=1):
            for col_idx, col_name in enumerate(columns):
                cell = table.cell(row_idx, col_idx)
                # Auto-shrink font if table has many columns to fit A4 page
                font_size = 9 if len(columns) > 5 else 10
                
                cell_value = str(nominee_data.get(col_name, ""))
                run = cell.paragraphs[0].add_run(cell_value)
                run.font.size = Pt(font_size)

        # Move table to correct position in document XML (exactly where placeholder was)
        doc.paragraphs[table_index]._element.addnext(table._element)

    doc.save(output_path)
    return output_path, filename