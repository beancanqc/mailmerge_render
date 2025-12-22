"""
Mail Merge SaaS - Main Application
Flask web server for Render deployment - Word documents only
"""

import os
import tempfile
import zipfile
import shutil
from datetime import datetime
from pathlib import Path
import uuid

from flask import Flask, request, jsonify, send_file, session
from werkzeug.utils import secure_filename
from docx import Document
from docx.enum.text import WD_BREAK
import openpyxl
import re
from typing import List, Dict, Any, Optional
import mammoth

from jinja2 import Template

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Configure session management for production
app.secret_key = os.environ.get('SECRET_KEY', 'dev-secret-key-' + str(uuid.uuid4()))
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_PERMANENT'] = False
app.config['SESSION_USE_SIGNER'] = True
app.config['SESSION_KEY_PREFIX'] = 'mailmerge:'

# Configure folders
UPLOAD_FOLDER = tempfile.mkdtemp()
OUTPUT_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_TEMPLATE_EXTENSIONS = {'docx'}
ALLOWED_DATA_EXTENSIONS = {'xlsx'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

class MailMergeProcessor:
    def __init__(self, session_id=None):
        self.session_id = session_id or str(uuid.uuid4())
        self.template_path: Optional[str] = None
        self.data_path: Optional[str] = None
        self.data: List[Dict[str, Any]] = []
        self.first_column_header: Optional[str] = None
        
    def cleanup(self):
        """Clean up temporary files"""
        try:
            if self.template_path and os.path.exists(self.template_path):
                os.remove(self.template_path)
                print(f"Cleaned up template: {self.template_path}")
            
            if self.data_path and os.path.exists(self.data_path):
                os.remove(self.data_path)
                print(f"Cleaned up data file: {self.data_path}")
                
        except Exception as e:
            print(f"Cleanup error: {str(e)}")
        
        # Reset state
        self.template_path = None
        self.data_path = None
        self.data = []
        self.first_column_header = None
        
    def load_template(self, template_path: str) -> bool:
        """Load and validate Word template file"""
        try:
            # Clean up previous template
            if self.template_path and os.path.exists(self.template_path):
                os.remove(self.template_path)
            
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {template_path}")
            
            if not template_path.lower().endswith('.docx'):
                raise ValueError("Template must be a Word document (.docx)")
            
            # Test if file can be opened
            doc = Document(template_path)
            doc = None  # Close the document
            
            self.template_path = template_path
            print(f"Template loaded successfully: {template_path}")
            return True
            
        except Exception as e:
            print(f"Error loading template: {str(e)}")
            return False
    
    def load_data(self, data_path: str) -> bool:
        """Load and validate Excel data file"""
        try:
            # Clean up previous data file
            if self.data_path and os.path.exists(self.data_path):
                os.remove(self.data_path)
            
            if not os.path.exists(data_path):
                raise FileNotFoundError(f"Data file not found: {data_path}")
            
            if not data_path.lower().endswith('.xlsx'):
                raise ValueError("Data file must be an Excel file (.xlsx)")
            
            # Load Excel data using openpyxl
            workbook = openpyxl.load_workbook(data_path, data_only=True)
            sheet = workbook.active
            
            if sheet.max_row <= 1:
                workbook.close()
                raise ValueError("Excel file appears to be empty or has no data")
            
            # Convert to list of dictionaries
            headers = []
            for cell in sheet[1]:
                headers.append(str(cell.value) if cell.value is not None else "")
            
            # Store first column header for filename generation
            self.first_column_header = headers[0] if headers else None
            
            self.data = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                row_data = {}
                for i, value in enumerate(row):
                    if i < len(headers):
                        row_data[headers[i]] = str(value) if value is not None else ""
                self.data.append(row_data)
            
            workbook.close()
            
            if not self.data:
                raise ValueError("No data rows found in Excel file")
            
            self.data_path = data_path
            print(f"Data loaded successfully: {len(self.data)} records from {data_path}")
            return True
            
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            return False

    def _find_run_for_position(self, paragraph, text_position):
        """Find which run contains the text at the given position and return its formatting"""
        current_pos = 0
        
        for run in paragraph.runs:
            run_len = len(run.text)
            
            if current_pos <= text_position < current_pos + run_len:
                # Extract formatting from this run
                formatting = {
                    'bold': run.bold,
                    'italic': run.italic,
                    'underline': run.underline,
                    'font_name': run.font.name,
                    'font_size': run.font.size,
                    'font_color': run.font.color.rgb if run.font.color.rgb else None
                }
                return {'formatting': formatting}
            
            current_pos += run_len
        
        # If position not found, return formatting from first run
        if paragraph.runs:
            first_run = paragraph.runs[0]
            formatting = {
                'bold': first_run.bold,
                'italic': first_run.italic,
                'underline': first_run.underline,
                'font_name': first_run.font.name,
                'font_size': first_run.font.size,
                'font_color': first_run.font.color.rgb if first_run.font.color.rgb else None
            }
            return {'formatting': formatting}
        
        return None

    def _apply_formatting(self, run, formatting):
        """Apply formatting to a run"""
        try:
            if formatting.get('bold') is not None:
                run.bold = formatting['bold']
            if formatting.get('italic') is not None:
                run.italic = formatting['italic']
            if formatting.get('underline') is not None:
                run.underline = formatting['underline']
            if formatting.get('font_name'):
                run.font.name = formatting['font_name']
            if formatting.get('font_size'):
                run.font.size = formatting['font_size']
            if formatting.get('font_color'):
                run.font.color.rgb = formatting['font_color']
        except Exception as e:
            print(f"Error applying formatting: {e}")
            # Continue without formatting if there's an error

    def replace_merge_fields_advanced(self, paragraph, data_row: Dict[str, Any]):
        """Advanced merge field replacement that preserves individual character formatting"""
        full_text = paragraph.text
        
        # Find all merge fields
        merge_fields = re.finditer(r'\{\{(\w+)\}\}', full_text)
        merge_list = list(merge_fields)
        
        if not merge_list:
            return
        
        # Create a list to store new runs
        new_runs_data = []
        current_pos = 0
        
        for match in merge_list:
            field_name = match.group(1)
            start_pos = match.start()
            end_pos = match.end()
            replacement_text = str(data_row.get(field_name, ""))
            
            # Add text before the merge field
            if start_pos > current_pos:
                before_text = full_text[current_pos:start_pos]
                if before_text:
                    # Find the run that contains this text and its formatting
                    run_info = self._find_run_for_position(paragraph, current_pos)
                    new_runs_data.append({
                        'text': before_text,
                        'formatting': run_info['formatting'] if run_info else None
                    })
            
            # Add the replacement text with the formatting of the merge field location
            if replacement_text:
                run_info = self._find_run_for_position(paragraph, start_pos)
                new_runs_data.append({
                    'text': replacement_text,
                    'formatting': run_info['formatting'] if run_info else None
                })
            
            current_pos = end_pos
        
        # Add remaining text after the last merge field
        if current_pos < len(full_text):
            remaining_text = full_text[current_pos:]
            if remaining_text:
                run_info = self._find_run_for_position(paragraph, current_pos)
                new_runs_data.append({
                    'text': remaining_text,
                    'formatting': run_info['formatting'] if run_info else None
                })
        
        # Clear existing runs
        for run in paragraph.runs:
            run.clear()
        
        # Create new runs with preserved formatting
        for run_data in new_runs_data:
            if run_data['text']:
                new_run = paragraph.add_run(run_data['text'])
                if run_data['formatting']:
                    self._apply_formatting(new_run, run_data['formatting'])
    
    def replace_merge_fields(self, doc: Document, data_row: Dict[str, Any]) -> Document:
        """Replace merge fields with actual data while preserving formatting"""
        
        # Replace in paragraphs with advanced formatting preservation
        for paragraph in doc.paragraphs:
            if '{{' in paragraph.text and '}}' in paragraph.text:
                self.replace_merge_fields_advanced(paragraph, data_row)
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{' in paragraph.text and '}}' in paragraph.text:
                            self.replace_merge_fields_advanced(paragraph, data_row)
        
        # Replace in headers and footers
        for section in doc.sections:
            # Header
            if section.header:
                for paragraph in section.header.paragraphs:
                    if '{{' in paragraph.text and '}}' in paragraph.text:
                        self.replace_merge_fields_advanced(paragraph, data_row)
            
            # Footer
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    if '{{' in paragraph.text and '}}' in paragraph.text:
                        self.replace_merge_fields_advanced(paragraph, data_row)
        
        return doc
    
    def generate_single_word(self, output_path: str) -> bool:
        """Generate a single Word document using SECTION BREAKS - Most reliable method"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            print(f"Creating single Word document with {len(self.data)} records using SECTION BREAKS...")
            
            # Start with first record - load fresh template and process
            final_doc = Document(self.template_path)
            final_doc = self.replace_merge_fields(final_doc, self.data[0])
            print(f"Added record 1 of {len(self.data)}")
            
            # Add remaining records using section breaks
            for i, row_data in enumerate(self.data[1:], 1):
                print(f"Adding record {i+1} of {len(self.data)} with section break...")
                
                # Load and process template for this record
                template_doc = Document(self.template_path)
                processed_doc = self.replace_merge_fields(template_doc, row_data)
                
                # Add new section with NEW_PAGE start (more reliable than page breaks)
                from docx.enum.section import WD_SECTION_START
                new_section = final_doc.add_section(WD_SECTION_START.NEW_PAGE)
                
                # Copy all elements from processed document in their original order
                # This preserves the document structure with tables in correct positions
                for element in processed_doc.element.body:
                    if element.tag.endswith('p'):  # Paragraph
                        # Find corresponding paragraph in processed_doc
                        for para in processed_doc.paragraphs:
                            if para._element == element:
                                new_para = final_doc.add_paragraph()
                                
                                # Copy paragraph-level formatting
                                try:
                                    new_para.style = para.style
                                    new_para.alignment = para.alignment
                                except:
                                    pass
                                
                                # Copy all runs with their formatting
                                for run in para.runs:
                                    new_run = new_para.add_run(run.text)
                                    
                                    # Copy comprehensive formatting
                                    try:
                                        if run.bold is not None:
                                            new_run.bold = run.bold
                                        if run.italic is not None:
                                            new_run.italic = run.italic
                                        if run.underline is not None:
                                            new_run.underline = run.underline
                                        if run.font.size:
                                            new_run.font.size = run.font.size
                                        if run.font.name:
                                            new_run.font.name = run.font.name
                                        if run.font.color.rgb:
                                            new_run.font.color.rgb = run.font.color.rgb
                                    except:
                                        pass
                                break
                    
                    elif element.tag.endswith('tbl'):  # Table
                        # Find corresponding table in processed_doc
                        for table in processed_doc.tables:
                            if table._element == element:
                                # Create new table with same dimensions
                                new_table = final_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                                
                                # Copy table-level formatting
                                try:
                                    if table.style:
                                        new_table.style = table.style
                                    
                                    # Copy table alignment
                                    if hasattr(table, 'alignment'):
                                        new_table.alignment = table.alignment
                                        
                                except Exception as table_style_error:
                                    print(f"   ⚠️  Table style copy failed: {table_style_error}")
                                
                                # Copy column widths to maintain table structure
                                try:
                                    for col_idx in range(len(table.columns)):
                                        if col_idx < len(new_table.columns):
                                            original_width = table.columns[col_idx].width
                                            if original_width:
                                                new_table.columns[col_idx].width = original_width
                                except Exception as width_error:
                                    print(f"   ⚠️  Column width copy failed: {width_error}")
                                
                                # Copy row heights and content
                                for row_idx, row in enumerate(table.rows):
                                    new_row = new_table.rows[row_idx]
                                    
                                    # Copy row height if available
                                    try:
                                        if hasattr(row, 'height') and row.height:
                                            new_row.height = row.height
                                    except:
                                        pass
                                    
                                    # Copy cell content and formatting
                                    for col_idx, cell in enumerate(row.cells):
                                        new_cell = new_row.cells[col_idx]
                                        
                                        # Copy cell background color and borders
                                        try:
                                            # Direct approach to copy cell shading (background color)
                                            from docx.oxml import OxmlElement, ns
                                            from docx.oxml.ns import qn
                                            
                                            # Get original cell properties
                                            original_tc_pr = cell._element.tcPr
                                            if original_tc_pr is not None:
                                                # Get or create cell properties for new cell
                                                new_tc_pr = new_cell._element.tcPr
                                                if new_tc_pr is None:
                                                    new_tc_pr = OxmlElement('w:tcPr')
                                                    new_cell._element.insert(0, new_tc_pr)
                                                
                                                # Copy shading (background color)
                                                original_shd = original_tc_pr.find(qn('w:shd'))
                                                if original_shd is not None:
                                                    # Remove existing shading if any
                                                    existing_shd = new_tc_pr.find(qn('w:shd'))
                                                    if existing_shd is not None:
                                                        new_tc_pr.remove(existing_shd)
                                                    
                                                    # Create new shading element
                                                    new_shd = OxmlElement('w:shd')
                                                    # Copy all shading attributes
                                                    for attr_name, attr_value in original_shd.attrib.items():
                                                        new_shd.set(attr_name, attr_value)
                                                    new_tc_pr.append(new_shd)
                                                    print(f"     🎨 Copied cell shading: {original_shd.attrib}")
                                                
                                                # Copy table cell borders
                                                original_borders = original_tc_pr.find(qn('w:tcBorders'))
                                                if original_borders is not None:
                                                    # Remove existing borders if any
                                                    existing_borders = new_tc_pr.find(qn('w:tcBorders'))
                                                    if existing_borders is not None:
                                                        new_tc_pr.remove(existing_borders)
                                                    
                                                    # Create new borders element
                                                    new_borders = OxmlElement('w:tcBorders')
                                                    # Copy all border elements
                                                    for border_element in original_borders:
                                                        new_border = OxmlElement(border_element.tag)
                                                        for attr_name, attr_value in border_element.attrib.items():
                                                            new_border.set(attr_name, attr_value)
                                                        new_borders.append(new_border)
                                                    new_tc_pr.append(new_borders)
                                                
                                                # Copy vertical alignment
                                                original_valign = original_tc_pr.find(qn('w:vAlign'))
                                                if original_valign is not None:
                                                    existing_valign = new_tc_pr.find(qn('w:vAlign'))
                                                    if existing_valign is not None:
                                                        new_tc_pr.remove(existing_valign)
                                                    
                                                    new_valign = OxmlElement('w:vAlign')
                                                    for attr_name, attr_value in original_valign.attrib.items():
                                                        new_valign.set(attr_name, attr_value)
                                                    new_tc_pr.append(new_valign)
                                        
                                        except Exception as cell_format_error:
                                            print(f"   ⚠️  Cell formatting copy failed: {cell_format_error}")
                                            # Fallback to basic copy
                                            try:
                                                if hasattr(cell._element, 'tcPr') and cell._element.tcPr is not None:
                                                    import copy
                                                    new_cell._element.tcPr = copy.deepcopy(cell._element.tcPr)
                                            except:
                                                pass
                                        
                                        # Clear default content and copy actual content
                                        new_cell.text = ""
                                        
                                        # Copy paragraphs
                                        for para_idx, para in enumerate(cell.paragraphs):
                                            if para.text.strip() or len(para.runs) > 0:
                                                if para_idx == 0:
                                                    # Use the existing first paragraph
                                                    new_para = new_cell.paragraphs[0]
                                                else:
                                                    # Add additional paragraphs
                                                    new_para = new_cell.add_paragraph()
                                                
                                                # Copy paragraph formatting
                                                try:
                                                    new_para.alignment = para.alignment
                                                    if hasattr(para, 'style') and para.style:
                                                        new_para.style = para.style
                                                except:
                                                    pass
                                                
                                                # Copy runs with all formatting
                                                for run in para.runs:
                                                    new_run = new_para.add_run(run.text)
                                                    try:
                                                        # Copy all text formatting
                                                        if run.bold is not None:
                                                            new_run.bold = run.bold
                                                        if run.italic is not None:
                                                            new_run.italic = run.italic
                                                        if run.underline is not None:
                                                            new_run.underline = run.underline
                                                        if run.font.size:
                                                            new_run.font.size = run.font.size
                                                        if run.font.name:
                                                            new_run.font.name = run.font.name
                                                        if run.font.color.rgb:
                                                            new_run.font.color.rgb = run.font.color.rgb
                                                    except Exception as run_format_error:
                                                        pass  # Continue even if some formatting fails
                                
                                print(f"   ✅ Copied table with enhanced formatting preservation")
                                break
            
            # Save the final document
            final_doc.save(output_path)
            print(f"✅ Successfully created single Word document with proper table positioning")
            return True
            
        except Exception as e:
            print(f"❌ Error creating single Word document: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Fallback to traditional approach if section breaks fail
            print("🔄 Trying fallback approach with traditional page breaks...")
            return self.generate_single_word_fallback(output_path)

    def generate_single_word_fallback(self, output_path: str) -> bool:
        """Fallback method using traditional page breaks with XML manipulation"""
        try:
            print("Using fallback method with XML-level page break insertion...")
            
            # Process all records first
            all_processed_docs = []
            for i, row_data in enumerate(self.data):
                template_doc = Document(self.template_path)
                processed_doc = self.replace_merge_fields(template_doc, row_data)
                all_processed_docs.append(processed_doc)
                print(f"Processed record {i+1} of {len(self.data)}")
            
            # Start with first document
            final_doc = all_processed_docs[0]
            
            # Add remaining documents with XML page breaks
            for i, doc in enumerate(all_processed_docs[1:], 1):
                print(f"Merging record {i+1} with XML page break...")
                
                # Insert page break at XML level (more reliable)
                body = final_doc._body._body
                
                # Create page break paragraph
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls, qn
                
                # Add page break using XML
                page_break_xml = f'''
                <w:p {nsdecls('w')}>
                    <w:r>
                        <w:br w:type="page"/>
                    </w:r>
                </w:p>
                '''
                page_break_p = parse_xml(page_break_xml)
                body.append(page_break_p)
                
                # Copy all content from the document
                for para in doc.paragraphs:
                    new_para = final_doc.add_paragraph()
                    
                    try:
                        new_para.style = para.style
                        new_para.alignment = para.alignment
                    except:
                        pass
                    
                    for run in para.runs:
                        new_run = new_para.add_run(run.text)
                        try:
                            if run.bold is not None:
                                new_run.bold = run.bold
                            if run.italic is not None:
                                new_run.italic = run.italic
                            if run.underline is not None:
                                new_run.underline = run.underline
                        except:
                            pass
                
                # Copy tables
                for table in doc.tables:
                    new_table = final_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                    for row_idx, row in enumerate(table.rows):
                        for col_idx, cell in enumerate(row.cells):
                            new_table.cell(row_idx, col_idx).text = cell.text
            
            final_doc.save(output_path)
            print("✅ Fallback method successful")
            return True
            
        except Exception as e:
            print(f"❌ Fallback method also failed: {str(e)}")
            return False
    
    def generate_multiple_word(self, output_dir: str) -> bool:
        """Generate multiple Word documents (one per record)"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            print(f"📁 Creating directory: {output_dir}")
            os.makedirs(output_dir, exist_ok=True)
            
            print(f"📄 Generating {len(self.data)} individual Word documents...")
            
            generated_files = []
            for index, row_data in enumerate(self.data):
                print(f"   Processing record {index+1}/{len(self.data)}...")
                
                # Load fresh template for each document
                doc = Document(self.template_path)
                
                # Replace merge fields
                processed_doc = self.replace_merge_fields(doc, row_data)
                
                # Generate filename (use first column value or index)
                if self.first_column_header and self.first_column_header in row_data:
                    first_value = row_data[self.first_column_header]
                else:
                    first_value = list(row_data.values())[0] if row_data else f"record_{index+1}"
                # Clean filename - remove invalid characters
                safe_filename = re.sub(r'[<>:"/\\|?*]', '_', str(first_value))
                safe_filename = safe_filename.strip()[:50]  # Limit length
                if not safe_filename:
                    safe_filename = f"record_{index+1}"
                
                output_path = os.path.join(output_dir, f"{safe_filename}.docx")
                
                # Handle duplicate filenames
                counter = 1
                original_path = output_path
                while os.path.exists(output_path):
                    base_name = os.path.splitext(original_path)[0]
                    output_path = f"{base_name}_{counter}.docx"
                    counter += 1
                
                # Save the document
                processed_doc.save(output_path)
                generated_files.append(output_path)
                print(f"   ✅ Saved: {os.path.basename(output_path)}")
            
            print(f"🎉 Successfully generated {len(generated_files)} Word documents")
            return True
            
        except Exception as e:
            print(f"❌ Error creating multiple Word files: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def convert_word_to_pdf_simple(self, word_path: str, pdf_path: str) -> bool:
        """Linux-optimized Word to PDF conversion with enhanced formatting preservation"""
        try:
            print(f"🔄 Converting {os.path.basename(word_path)} to PDF...")
            print(f"   Input: {word_path} ({os.path.getsize(word_path):,} bytes)")
            print(f"   Output: {pdf_path}")
            
            # Detect platform for optimal conversion strategy
            import platform
            system = platform.system()
            print(f"   🖥️  Platform: {system}")
            
            if system == "Linux":
                # Linux-optimized conversion order (prioritizing methods that work well on headless servers)
                print(f"   🐧 Using Linux-optimized conversion strategy...")
                
                # 1. LibreOffice headless (most reliable on Linux with proper spacing)
                if self.convert_docx_to_pdf_libreoffice_enhanced(word_path, pdf_path):
                    print(f"   ✅ Enhanced LibreOffice conversion successful")
                    return True
                
                # 2. Pandoc with specific Linux options
                if self.convert_docx_to_pdf_pandoc_linux(word_path, pdf_path):
                    print(f"   ✅ Pandoc Linux conversion successful")
                    return True
                
                # 3. Enhanced HTML conversion with better CSS for Linux
                if self.convert_docx_to_pdf_html_linux_enhanced(word_path, pdf_path):
                    print(f"   ✅ Enhanced HTML conversion successful")
                    return True
                
                # 4. Basic fallback
                if self.create_basic_text_pdf_from_docx_enhanced(word_path, pdf_path):
                    print(f"   ⚠️  Enhanced text PDF created")
                    return True
                
            else:
                # Windows/Mac conversion order (original priority)
                print(f"   🪟 Using Windows/Mac conversion strategy...")
                
                # 1. Pandoc first (highest quality formatting preservation)
                if self.convert_docx_to_pdf_pandoc(word_path, pdf_path):
                    print(f"   ✅ Pandoc conversion successful")
                    return True
                
                # 2. Word COM (Windows only, native Microsoft Word quality)
                if self.convert_docx_to_pdf_with_word(word_path, pdf_path):
                    print(f"   ✅ Word COM conversion successful")
                    return True
                    
                # 3. LibreOffice (excellent cross-platform formatting preservation)
                if self.convert_docx_to_pdf_libreoffice(word_path, pdf_path):
                    print(f"   ✅ LibreOffice conversion successful")
                    return True
                
                # 4. docx2pdf (known formatting issues but works on Windows)
                if self.convert_docx_to_pdf_direct(word_path, pdf_path):
                    print(f"   ⚠️  docx2pdf conversion successful (may have formatting issues)")
                    return True
                
                # 5. HTML fallback
                if self.convert_docx_to_pdf_html_fallback(word_path, pdf_path):
                    print(f"   ⚠️  HTML fallback conversion successful")
                    return True
                
                # 6. Basic text PDF
                if self.create_basic_text_pdf_from_docx(word_path, pdf_path):
                    print(f"   ⚠️  Basic text PDF created")
                    return True
                
            print(f"   ❌ ALL conversion methods failed on {system}")
            return False
            
        except Exception as e:
            print(f"❌ Error in PDF conversion: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def convert_docx_to_pdf_pandoc(self, docx_path: str, pdf_path: str) -> bool:
        """Convert DOCX to PDF using Pandoc (highest quality formatting preservation)"""
        try:
            import subprocess
            import shutil
            
            print(f"🔄 Converting DOCX to PDF with Pandoc (highest quality)")
            print(f"   Input: {docx_path} ({os.path.getsize(docx_path):,} bytes)")
            print(f"   Output: {pdf_path}")
            
            # Check if Pandoc is available
            pandoc_cmd = shutil.which('pandoc')
            if not pandoc_cmd:
                print("❌ Pandoc not found in PATH")
                print("💡 Install Pandoc from: https://pandoc.org/installing.html")
                return False
            
            # Verify input file
            if not os.path.exists(docx_path):
                print(f"❌ Input DOCX file not found: {docx_path}")
                return False
            
            # Ensure output directory exists
            output_dir = os.path.dirname(pdf_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # Build Pandoc command with options for better formatting
            cmd = [
                pandoc_cmd,
                docx_path,
                '-o', pdf_path,
                '--pdf-engine=wkhtmltopdf',  # Use wkhtmltopdf for better formatting
                '-V', 'geometry:margin=1in',  # Set margins
                '-V', 'fontsize=11pt',  # Set font size
                '--standalone'  # Create standalone document
            ]
            
            # Try with wkhtmltopdf first, fallback to default PDF engine
            try:
                print("🔄 Running Pandoc with wkhtmltopdf...")
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                
                if result.returncode != 0:
                    # Try with default PDF engine (LaTeX)
                    print("⚠️  wkhtmltopdf failed, trying default PDF engine...")
                    cmd_fallback = [
                        pandoc_cmd,
                        docx_path,
                        '-o', pdf_path,
                        '-V', 'geometry:margin=1in',
                        '-V', 'fontsize=11pt',
                        '--standalone'
                    ]
                    
                    result = subprocess.run(cmd_fallback, capture_output=True, text=True, timeout=120)
                
                if result.returncode == 0:
                    # Verify PDF was created
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                        print(f"✅ Pandoc conversion successful: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
                        return True
                    else:
                        print("❌ Pandoc completed but no PDF was created")
                        return False
                else:
                    print(f"❌ Pandoc conversion failed:")
                    print(f"   Return code: {result.returncode}")
                    if result.stderr:
                        print(f"   Error: {result.stderr}")
                    return False
                    
            except subprocess.TimeoutExpired:
                print("❌ Pandoc conversion timed out after 2 minutes")
                return False
            except FileNotFoundError:
                print("❌ Pandoc executable not found")
                return False
            
        except Exception as e:
            print(f"❌ Error in Pandoc conversion: {str(e)}")
            return False

    def convert_docx_to_pdf_with_word(self, docx_path: str, pdf_path: str) -> bool:
        """Convert DOCX to PDF using Microsoft Word COM automation (Windows only)"""
        try:
            import win32com.client
            
            print(f"Converting {docx_path} to PDF using Microsoft Word...")
            
            # Convert paths to absolute paths
            docx_path = os.path.abspath(docx_path)
            pdf_path = os.path.abspath(pdf_path)
            
            # Start Word application
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False  # Run Word in background
            
            # Open the document
            doc = word.Documents.Open(docx_path)
            
            # Save as PDF (wdFormatPDF = 17)
            doc.SaveAs2(pdf_path, FileFormat=17)
            
            # Close document and quit Word
            doc.Close()
            word.Quit()
            
            print(f"✅ Successfully converted to PDF: {pdf_path}")
            return True
            
        except ImportError:
            print("ℹ️  Microsoft Word not available (non-Windows environment). Using fallback PDF conversion...")
            return False
        except Exception as e:
            print(f"❌ Error converting to PDF with Word: {str(e)}")
            try:
                # Try to clean up if something went wrong
                if 'doc' in locals():
                    doc.Close()
                if 'word' in locals():
                    word.Quit()
            except:
                pass
            return False

    def convert_docx_to_pdf_direct(self, docx_path: str, pdf_path: str) -> bool:
        """Direct DOCX to PDF conversion using docx2pdf (Windows-optimized)"""
        try:
            print(f"🔄 Converting DOCX to PDF with docx2pdf (Windows-optimized)")
            print(f"   Input: {docx_path} ({os.path.getsize(docx_path):,} bytes)")
            print(f"   Output: {pdf_path}")
            
            # Verify input file
            if not os.path.exists(docx_path):
                print(f"❌ Input DOCX file not found: {docx_path}")
                return False
            
            if os.path.getsize(docx_path) == 0:
                print("❌ Input DOCX file is empty")
                return False
            
            # Test if DOCX file can be opened first
            try:
                test_doc = Document(docx_path)
                print(f"✅ DOCX file validation passed ({len(test_doc.paragraphs)} paragraphs)")
            except Exception as docx_error:
                print(f"❌ DOCX file is corrupted: {docx_error}")
                return False
            
            # Try importing docx2pdf
            try:
                from docx2pdf import convert
                print("✅ docx2pdf library is available")
            except ImportError as e:
                print(f"❌ docx2pdf not available: {e}")
                print("💡 To install: pip install docx2pdf")
                return False
            
            # Ensure output directory exists
            output_dir = os.path.dirname(pdf_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # Convert with timeout and comprehensive error handling
            try:
                print("🔄 Converting with docx2pdf...")
                
                # Add timeout to prevent hanging
                import signal
                import threading
                
                conversion_success = [False]
                conversion_error = [None]
                
                def conversion_thread():
                    try:
                        convert(docx_path, pdf_path)
                        conversion_success[0] = True
                    except Exception as e:
                        conversion_error[0] = e
                
                thread = threading.Thread(target=conversion_thread)
                thread.daemon = True
                thread.start()
                thread.join(timeout=30)  # 30 second timeout
                
                if thread.is_alive():
                    print("❌ docx2pdf conversion timed out (30s)")
                    return False
                
                if conversion_error[0]:
                    print(f"❌ docx2pdf conversion error: {conversion_error[0]}")
                    return False
                
                if not conversion_success[0]:
                    print("❌ docx2pdf conversion failed silently")
                    return False
                
                # Verify output
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                    pdf_size = os.path.getsize(pdf_path)
                    print(f"✅ docx2pdf SUCCESS: {pdf_path} ({pdf_size:,} bytes)")
                    print("⚠️  NOTE: docx2pdf may have formatting issues (line breaks, spacing)")
                    print("   If PDF formatting is poor, try Pandoc or LibreOffice for better results")
                    
                    # Quick validation - check if PDF has content
                    if pdf_size > 1024:  # At least 1KB
                        print("✅ PDF appears to have substantial content")
                        return True
                    else:
                        print("⚠️  PDF is very small, might be empty")
                        # Don't fail completely, but warn
                        return True
                else:
                    print("❌ docx2pdf failed: No output or empty file")
                    if os.path.exists(pdf_path):
                        try:
                            os.remove(pdf_path)
                        except:
                            pass
                    return False
                    
            except Exception as conversion_error:
                print(f"❌ docx2pdf conversion error: {conversion_error}")
                # Clean up any partial file
                if os.path.exists(pdf_path):
                    try:
                        os.remove(pdf_path)
                    except:
                        pass
                return False
                
        except Exception as e:
            print(f"❌ docx2pdf method failed: {str(e)}")
            return False
    
    def convert_docx_to_pdf_libreoffice_enhanced(self, docx_path: str, pdf_path: str) -> bool:
        """Enhanced LibreOffice conversion specifically optimized for Linux servers with proper text spacing"""
        try:
            import subprocess
            import shutil
            import os
            
            # Check if LibreOffice is available
            libreoffice_cmd = None
            for cmd in ['libreoffice', 'libreoffice7.0', 'libreoffice6.4', '/usr/bin/libreoffice']:
                if shutil.which(cmd):
                    libreoffice_cmd = cmd
                    break
            
            if not libreoffice_cmd:
                print(f"   LibreOffice not found")
                return False
            
            print(f"   Using LibreOffice: {libreoffice_cmd}")
            output_dir = os.path.dirname(pdf_path)
            
            # Enhanced LibreOffice command with specific options for better formatting
            cmd = [
                libreoffice_cmd,
                '--headless',                    # Run without GUI
                '--convert-to', 'pdf',           # Convert to PDF
                '--writer',                      # Use Writer for processing
                '--calc',                        # Ensure calc is available for complex documents
                '--outdir', output_dir,          # Output directory
                docx_path
            ]
            
            print(f"   🔄 Running: {' '.join(cmd)}")
            
            # Set environment variables for better text rendering on Linux
            env = os.environ.copy()
            env.update({
                'SAL_USE_VCLPLUGIN': 'svp',      # Use server headless backend
                'DISPLAY': ':99',                # Virtual display if needed
                'LANG': 'en_US.UTF-8',          # Proper locale
                'LC_ALL': 'en_US.UTF-8'         # Proper locale
            })
            
            # Run with timeout and proper error handling
            result = subprocess.run(
                cmd, 
                capture_output=True, 
                text=True, 
                timeout=120,
                env=env
            )
            
            if result.returncode != 0:
                print(f"   LibreOffice error: {result.stderr}")
                return False
            
            # Check if the PDF was created (LibreOffice may change the filename)
            expected_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
            
            if os.path.exists(expected_pdf):
                if expected_pdf != pdf_path:
                    # Move to the desired location
                    shutil.move(expected_pdf, pdf_path)
                
                # Verify PDF has content and proper spacing
                if os.path.getsize(pdf_path) > 1000:  # Reasonable minimum size
                    print(f"   ✅ Enhanced LibreOffice conversion completed: {pdf_path}")
                    return True
                else:
                    print(f"   ❌ Generated PDF is too small (likely corrupted)")
                    return False
            else:
                print(f"   ❌ Expected PDF not found: {expected_pdf}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"   ❌ LibreOffice conversion timed out")
            return False
        except Exception as e:
            print(f"   ❌ Enhanced LibreOffice conversion error: {str(e)}")
            return False

    def convert_docx_to_pdf_libreoffice(self, docx_path: str, pdf_path: str) -> bool:
        """Standard LibreOffice conversion (original method for Windows/Mac)"""
        try:
            import subprocess
            import shutil
            import os
            
            # Check if LibreOffice is available
            libreoffice_cmd = None
            for cmd in ['libreoffice', '/usr/bin/libreoffice', 'soffice']:
                if shutil.which(cmd):
                    libreoffice_cmd = cmd
                    break
            
            if not libreoffice_cmd:
                print(f"   LibreOffice not found")
                return False
            
            output_dir = os.path.dirname(pdf_path)
            
            cmd = [
                libreoffice_cmd,
                '--headless',
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                docx_path
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
            
            if result.returncode == 0:
                # LibreOffice creates PDF with same base name
                expected_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_path))[0] + '.pdf')
                
                if os.path.exists(expected_pdf) and expected_pdf != pdf_path:
                    shutil.move(expected_pdf, pdf_path)
                
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 500:
                    return True
                    
            return False
            
        except Exception as e:
            print(f"   ❌ LibreOffice conversion error: {str(e)}")
            return False

    def convert_docx_to_pdf_pandoc_linux(self, docx_path: str, pdf_path: str) -> bool:
        """Enhanced Pandoc conversion with Linux-specific optimizations for proper text spacing"""
        try:
            import subprocess
            import shutil
            import tempfile
            import os
            
            # Check if pandoc is available
            if not shutil.which('pandoc'):
                print(f"   Pandoc not found")
                return False
            
            print(f"   Using Pandoc for Linux-optimized conversion")
            
            # Enhanced pandoc command with specific options for better spacing on Linux
            cmd = [
                'pandoc',
                docx_path,
                '-o', pdf_path,
                '--pdf-engine=wkhtmltopdf',      # Use wkhtmltopdf for better rendering
                '--pdf-engine-opt=--page-size', 'A4',
                '--pdf-engine-opt=--margin-top', '1in',
                '--pdf-engine-opt=--margin-bottom', '1in',
                '--pdf-engine-opt=--margin-left', '1in',
                '--pdf-engine-opt=--margin-right', '1in',
                '--pdf-engine-opt=--encoding', 'UTF-8',
                '--pdf-engine-opt=--print-media-type',  # Better text rendering
                '--verbose'
            ]
            
            # Fallback to other engines if wkhtmltopdf not available
            if not shutil.which('wkhtmltopdf'):
                # Try with LaTeX engine
                cmd = [
                    'pandoc',
                    docx_path,
                    '-o', pdf_path,
                    '--pdf-engine=xelatex',
                    '--verbose'
                ]
                
                if not shutil.which('xelatex'):
                    # Basic pandoc without specific engine
                    cmd = [
                        'pandoc',
                        docx_path,
                        '-o', pdf_path,
                        '--verbose'
                    ]
            
            print(f"   🔄 Running: {' '.join(cmd)}")
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
            
            if result.returncode == 0 and os.path.exists(pdf_path):
                if os.path.getsize(pdf_path) > 1000:  # Reasonable minimum size
                    print(f"   ✅ Pandoc Linux conversion completed: {pdf_path}")
                    return True
                else:
                    print(f"   ❌ Generated PDF is too small")
                    return False
            else:
                print(f"   ❌ Pandoc conversion failed: {result.stderr}")
                return False
                
        except subprocess.TimeoutExpired:
            print(f"   ❌ Pandoc conversion timed out")
            return False
        except Exception as e:
            print(f"   ❌ Pandoc Linux conversion error: {str(e)}")
            return False

    def convert_docx_to_pdf_html_linux_enhanced(self, docx_path: str, pdf_path: str) -> bool:
        """Enhanced HTML to PDF conversion specifically optimized for Linux with proper text spacing"""
        try:
            # Enhanced HTML conversion with better CSS for text spacing
            html_content = self.convert_docx_to_html_enhanced(docx_path)
            if not html_content:
                return False
            
            # Enhanced CSS specifically for Linux rendering with proper spacing
            enhanced_css = """
            <style>
            body {
                font-family: 'DejaVu Sans', 'Liberation Sans', Arial, sans-serif;
                font-size: 11pt;
                line-height: 1.6;
                margin: 1in;
                color: #000;
                text-align: left;
                word-spacing: normal;
                letter-spacing: 0.02em;
            }
            p {
                margin: 0 0 12pt 0;
                padding: 0;
                white-space: pre-line;
                word-wrap: break-word;
                text-justify: inter-word;
            }
            .paragraph {
                margin-bottom: 12pt;
                line-height: 1.6;
                white-space: pre-wrap;
            }
            h1, h2, h3, h4, h5, h6 {
                margin: 18pt 0 12pt 0;
                font-weight: bold;
                line-height: 1.4;
            }
            table {
                border-collapse: collapse;
                margin: 12pt 0;
                width: 100%;
            }
            td, th {
                padding: 6pt;
                border: 1px solid #ccc;
                text-align: left;
                vertical-align: top;
            }
            .bold { font-weight: bold; }
            .italic { font-style: italic; }
            .underline { text-decoration: underline; }
            
            /* Ensure proper word spacing */
            * {
                word-spacing: normal !important;
                letter-spacing: normal !important;
            }
            
            /* Page break settings for better printing */
            @media print {
                body { 
                    margin: 0.5in;
                    font-size: 10pt;
                }
                p {
                    page-break-inside: avoid;
                    orphans: 3;
                    widows: 3;
                }
            }
            </style>
            """
            
            # Combine CSS with HTML
            full_html = f"""<!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Document</title>
                {enhanced_css}
            </head>
            <body>
                {html_content}
            </body>
            </html>"""
            
            return self.convert_html_to_pdf_enhanced(full_html, pdf_path)
            
        except Exception as e:
            print(f"   ❌ Enhanced HTML Linux conversion error: {str(e)}")
            return False
        """Convert DOCX to PDF using LibreOffice with enhanced formatting options"""
        try:
            import subprocess
            import shutil
            
            print(f"🔄 Converting DOCX to PDF with LibreOffice (excellent formatting)")
            print(f"   Input: {docx_path} ({os.path.getsize(docx_path):,} bytes)")
            print(f"   Output: {pdf_path}")
            
            # Check if LibreOffice is available
            libreoffice_cmd = shutil.which('libreoffice') or shutil.which('soffice')
            if not libreoffice_cmd:
                print("❌ LibreOffice not found in PATH")
                print("💡 Install LibreOffice or use 'apt-get install libreoffice-headless' on Linux")
                return False
            
            # Verify input file
            if not os.path.exists(docx_path):
                print(f"❌ Input DOCX file not found: {docx_path}")
                return False
            
            # Ensure output directory exists
            output_dir = os.path.dirname(pdf_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # Use enhanced LibreOffice command with formatting options
            cmd = [
                libreoffice_cmd,
                '--headless',  # Run without GUI
                '--convert-to', 'pdf',
                '--outdir', output_dir,
                docx_path
            ]
            
            try:
                print(f"🔄 Running LibreOffice conversion...")
                print(f"   Command: {' '.join(cmd)}")
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
                
                if result.returncode == 0:
                    # LibreOffice creates PDF with same base name as input
                    expected_pdf = os.path.join(
                        output_dir,
                        os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
                    )
                    
                    print(f"   Expected output: {expected_pdf}")
                    
                    # Move to desired location if different
                    if expected_pdf != pdf_path and os.path.exists(expected_pdf):
                        print(f"   Moving {expected_pdf} to {pdf_path}")
                        shutil.move(expected_pdf, pdf_path)
                    
                    # Verify PDF was created successfully
                    if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                        print(f"✅ LibreOffice conversion successful: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
                        return True
                    else:
                        print("❌ LibreOffice conversion failed - no valid output file created")
                        if result.stdout:
                            print(f"   Stdout: {result.stdout}")
                        if result.stderr:
                            print(f"   Stderr: {result.stderr}")
                        return False
                else:
                    print(f"❌ LibreOffice conversion failed:")
                    print(f"   Return code: {result.returncode}")
                    if result.stdout:
                        print(f"   Stdout: {result.stdout}")
                    if result.stderr:
                        print(f"   Stderr: {result.stderr}")
                    return False
                    
            except subprocess.TimeoutExpired:
                print("❌ LibreOffice conversion timed out after 2 minutes")
                return False
            except FileNotFoundError:
                print("❌ LibreOffice executable not found")
                return False
            
        except Exception as e:
            print(f"❌ Error in LibreOffice conversion: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def convert_html_to_pdf(self, html_content: str, output_path: str) -> bool:
        """Convert HTML content to PDF using weasyprint with enhanced formatting for line break preservation"""
        try:
            print("🔄 Converting HTML to PDF using WeasyPrint (enhanced formatting)...")
            
            # Check if WeasyPrint is available and import it properly to avoid naming conflicts
            try:
                import weasyprint as wp
                print("✅ WeasyPrint is available")
            except ImportError as e:
                print(f"❌ WeasyPrint not available: {e}")
                print("💡 Try installing: pip install weasyprint")
                return self.convert_html_to_pdf_alternative(html_content, output_path)
            
            # Enhanced CSS with better line break and spacing preservation
            html_with_css = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    @page {{
                        size: A4;
                        margin: 0.8in 1in;
                        @top-center {{
                            content: "";
                        }}
                    }}
                    
                    body {{
                        font-family: 'Calibri', 'DejaVu Sans', Arial, sans-serif;
                        font-size: 11pt;
                        line-height: 1.3;
                        color: #000000;
                        margin: 0;
                        padding: 0;
                        background: white;
                        word-wrap: break-word;
                        overflow-wrap: break-word;
                    }}
                    
                    /* Preserve line breaks and spacing */
                    p {{
                        margin: 0 0 8pt 0;
                        text-align: left;
                        white-space: pre-wrap; /* Preserve whitespace and line breaks */
                        word-wrap: break-word;
                        orphans: 2;
                        widows: 2;
                    }}
                    
                    /* Empty paragraphs should create vertical space */
                    p:empty {{
                        height: 8pt;
                        margin: 0;
                    }}
                    
                    /* Preserve line breaks in text */
                    br {{
                        margin: 0;
                        padding: 0;
                        line-height: 1.3;
                    }}
                    
                    /* Headings with proper spacing */
                    h1, h2, h3, h4, h5, h6 {{
                        color: #1f497d;
                        margin-top: 16pt;
                        margin-bottom: 8pt;
                        font-weight: bold;
                        page-break-after: avoid;
                        orphans: 3;
                        widows: 3;
                    }}
                    h1 {{ font-size: 16pt; line-height: 1.2; }}
                    h2 {{ font-size: 14pt; line-height: 1.2; }}
                    h3 {{ font-size: 12pt; line-height: 1.2; }}
                    h4, h5, h6 {{ font-size: 11pt; line-height: 1.2; }}
                    
                    /* Text formatting that preserves document structure */
                    strong, b {{
                        font-weight: bold;
                        color: inherit;
                    }}
                    
                    em, i {{
                        font-style: italic;
                        color: inherit;
                    }}
                    
                    /* Tables with better formatting */
                    table {{
                        border-collapse: collapse;
                        width: 100%;
                        margin: 8pt 0 16pt 0;
                        page-break-inside: avoid;
                        font-size: 10pt;
                    }}
                    
                    th, td {{
                        border: 1px solid #000000;
                        padding: 6pt 8pt;
                        text-align: left;
                        vertical-align: top;
                        word-wrap: break-word;
                    }}
                    
                    th {{
                        background-color: #d9d9d9;
                        font-weight: bold;
                    }}
                    
                    /* Page breaks */
                    .page-break {{
                        page-break-before: always;
                        margin: 0;
                        padding: 0;
                    }}
                    
                    /* Ensure content flows properly */
                    div {{
                        margin: 0;
                        padding: 0;
                    }}
                    
                    /* Preserve original Word spacing */
                    .merge-field-replacement {{
                        display: inline;
                        white-space: pre-wrap;
                    }}
                    
                    /* Address blocks and contact info */
                    .address-block {{
                        margin-bottom: 16pt;
                        line-height: 1.2;
                    }}
                    
                    /* Document sections */
                    .document-section {{
                        margin-bottom: 16pt;
                        page-break-inside: avoid;
                    }}
                </style>
            </head>
            <body>
                {html_content}
            </body>
            </html>
            """
            
            # Convert to PDF with enhanced error handling
            try:
                print("🔄 Generating PDF from HTML with enhanced formatting...")
                
                # Create HTML document object with base URL for relative resources
                html_doc = wp.HTML(string=html_with_css, base_url='.')
                
                # Write PDF to file with optimized settings
                html_doc.write_pdf(
                    output_path,
                    optimize_images=True,
                    presentational_hints=True  # Preserve HTML presentational attributes
                )
                
                # Verify PDF was created successfully
                if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                    print(f"✅ Enhanced PDF created successfully: {output_path} ({os.path.getsize(output_path):,} bytes)")
                    return True
                else:
                    print("❌ PDF file was not created or is empty")
                    return False
                    return False
                    
            except Exception as pdf_error:
                print(f"❌ WeasyPrint PDF generation error: {pdf_error}")
                import traceback
                traceback.print_exc()
                print("🔄 Trying alternative PDF generation method...")
                return self.convert_html_to_pdf_alternative(html_content, output_path)
                
        except Exception as e:
            print(f"❌ HTML to PDF conversion failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
                
    def convert_docx_to_pdf_html_fallback(self, docx_path: str, pdf_path: str) -> bool:
        """HTML fallback PDF conversion (formatting may be lost) - use as last resort"""
        try:
            print("🚨 ========================================")
            print("🚨 WARNING: Using HTML fallback method!")
            print("🚨 Word formatting will be LOST!")
            print("🚨 Install LibreOffice for better quality")
            print("🚨 ========================================")
            
            # First convert DOCX to HTML
            html_content = self.convert_docx_to_html(docx_path)
            if not html_content:
                print("❌ Failed to convert DOCX to HTML")
                return False
            
            print(f"✅ Successfully converted DOCX to HTML ({len(html_content)} characters)")
            
            # Then convert HTML to PDF (try multiple methods)
            if self.convert_html_to_pdf(html_content, pdf_path):
                print(f"⚠️  HTML-to-PDF conversion completed: {pdf_path}")
                print("⚠️  WARNING: Output PDF has basic formatting only!")
                return True
            elif self.convert_html_to_pdf_alternative(html_content, pdf_path):
                print(f"⚠️  Alternative HTML-to-PDF conversion completed: {pdf_path}")
                print("⚠️  WARNING: Output PDF has very basic formatting!")
                return True
            else:
                print("❌ All HTML-to-PDF conversion methods failed")
                # As absolute last resort, create a simple text PDF
                return self.create_basic_text_pdf(html_content, pdf_path)
                
        except Exception as e:
            print(f"❌ HTML fallback PDF conversion failed: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def convert_html_to_pdf_alternative(self, html_content: str, output_path: str) -> bool:
        """Alternative PDF generation using reportlab as fallback"""
        try:
            print("🔄 Using alternative PDF generation with reportlab...")
            
            # Try importing reportlab
            try:
                from reportlab.pdfgen import canvas
                from reportlab.lib.pagesizes import letter, A4
                from reportlab.lib.styles import getSampleStyleSheet
                from reportlab.platypus import SimpleDocTemplate, Paragraph
                from io import StringIO
                import html
                print("✅ Reportlab is available")
            except ImportError:
                print("❌ Reportlab not available. Installing basic text-only PDF fallback...")
                return self.create_basic_text_pdf(html_content, output_path)
            
            # Create PDF with reportlab
            doc = SimpleDocTemplate(output_path, pagesize=A4)
            styles = getSampleStyleSheet()
            story = []
            
            # Convert HTML to simple text and create paragraphs
            # Remove HTML tags for basic text conversion
            import re
            text_content = re.sub(r'<[^>]+>', ' ', html_content)
            text_content = html.unescape(text_content)
            
            # Split into paragraphs and add to story
            paragraphs = text_content.split('\n\n')
            for para_text in paragraphs:
                if para_text.strip():
                    para = Paragraph(para_text.strip(), styles['Normal'])
                    story.append(para)
            
            # Build PDF
            doc.build(story)
            
            # Verify PDF was created
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print(f"✅ Alternative PDF created successfully: {output_path} ({os.path.getsize(output_path)} bytes)")
                return True
            else:
                print("❌ Alternative PDF creation failed")
                return False
                
        except Exception as e:
            print(f"❌ Alternative PDF generation failed: {str(e)}")
            return self.create_basic_text_pdf(html_content, output_path)
    
    def create_basic_text_pdf(self, html_content: str, output_path: str) -> bool:
        """Last resort: create a basic text-only PDF"""
        try:
            print("🔄 Creating basic text-only PDF as last resort...")
            
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            import re
            import html
            
            # Convert HTML to plain text
            text_content = re.sub(r'<[^>]+>', ' ', html_content)
            text_content = html.unescape(text_content)
            text_content = ' '.join(text_content.split())  # Clean up whitespace
            
            # Create PDF
            c = canvas.Canvas(output_path, pagesize=letter)
            width, height = letter
            
            # Set up text
            c.setFont("Helvetica", 12)
            y_position = height - 72  # Start 1 inch from top
            line_height = 14
            
            # Split text into lines that fit the page width
            words = text_content.split()
            lines = []
            current_line = ""
            
            for word in words:
                test_line = current_line + " " + word if current_line else word
                if c.stringWidth(test_line, "Helvetica", 12) < (width - 144):  # Leave 1 inch margins
                    current_line = test_line
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = word
            
            if current_line:
                lines.append(current_line)
            
            # Write lines to PDF
            for line in lines:
                if y_position < 72:  # Start new page if needed
                    c.showPage()
                    c.setFont("Helvetica", 12)
                    y_position = height - 72
                
                c.drawString(72, y_position, line)
                y_position -= line_height
            
            c.save()
            
            # Verify PDF was created
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print(f"✅ Basic PDF created successfully: {output_path} ({os.path.getsize(output_path)} bytes)")
                return True
            else:
                print("❌ Basic PDF creation failed")
                return False
                
        except Exception as e:
            print(f"❌ Basic PDF creation failed: {str(e)}")
            return False
    
    def create_basic_text_pdf_from_docx(self, docx_path: str, output_path: str) -> bool:
        """Create basic text PDF directly from DOCX file with proper formatting"""
        try:
            print("🔄 Creating formatted text PDF from DOCX...")
            
            # Read text from DOCX with proper structure preservation
            doc = Document(docx_path)
            
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.lib.styles import getSampleStyleSheet
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
            from reportlab.lib.units import inch
            
            # Create PDF document
            pdf_doc = SimpleDocTemplate(output_path, pagesize=A4,
                                      rightMargin=72, leftMargin=72,
                                      topMargin=72, bottomMargin=18)
            
            # Container for the 'Flowable' objects
            story = []
            styles = getSampleStyleSheet()
            
            # Track if this is a multi-record document (check for section breaks)
            has_multiple_records = len(self.data) > 1
            record_count = 0
            
            # Process paragraphs
            for i, paragraph in enumerate(doc.paragraphs):
                para_text = paragraph.text.strip()
                
                if para_text:
                    # Detect if this might be a new record (simple heuristic)
                    # Look for patterns like "Invoice", "Dear", or repeated content
                    is_new_record = False
                    if has_multiple_records and record_count > 0:
                        # Common mail merge starting patterns
                        record_starters = ['Invoice', 'Dear', 'Hello', 'Hi', 'Welcome']
                        for starter in record_starters:
                            if para_text.startswith(starter):
                                is_new_record = True
                                break
                    
                    # Add page break before new record (except first)
                    if is_new_record and record_count > 0:
                        story.append(PageBreak())
                        print(f"   📄 Page break added before record {record_count + 1}")
                    
                    # Determine style based on content
                    if any(heading in para_text.lower() for heading in ['invoice', 'title', 'heading']):
                        # This looks like a heading
                        para = Paragraph(para_text, styles['Title'])
                        story.append(para)
                        story.append(Spacer(1, 12))
                        if is_new_record:
                            record_count += 1
                    else:
                        # Regular paragraph
                        para = Paragraph(para_text, styles['Normal'])
                        story.append(para)
                        story.append(Spacer(1, 6))
                        
                        # Count records based on content patterns
                        if not is_new_record and has_multiple_records:
                            # Look for end-of-record patterns
                            end_patterns = ['regards', 'sincerely', 'best', 'thank you', 'company']
                            if any(end in para_text.lower() for end in end_patterns):
                                if record_count == 0:  # First record ending
                                    record_count = 1
                else:
                    # Empty paragraph - add some space
                    story.append(Spacer(1, 12))
            
            # Add table content if any
            for table in doc.tables:
                # Add spacing before table
                story.append(Spacer(1, 12))
                
                # Convert table to simple text representation
                table_text = "\\n--- Table ---\\n"
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        if cell.text.strip():
                            row_text.append(cell.text.strip())
                    if row_text:
                        table_text += " | ".join(row_text) + "\\n"
                
                table_para = Paragraph(table_text.replace('\\n', '<br/>'), styles['Normal'])
                story.append(table_para)
                story.append(Spacer(1, 12))
            
            # Build PDF
            pdf_doc.build(story)
            
            # Verify PDF was created
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                pdf_size = os.path.getsize(output_path)
                print(f"✅ Formatted PDF created: {output_path} ({pdf_size:,} bytes)")
                if has_multiple_records:
                    print(f"   📄 Multi-page PDF with {record_count + 1} records")
                return True
            else:
                print("❌ Formatted PDF creation failed")
                return False
                
        except Exception as e:
            print(f"❌ Formatted PDF creation failed: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Fallback to simple canvas method
            print("🔄 Falling back to simple canvas PDF...")
            return self.create_simple_canvas_pdf(docx_path, output_path)
    
    def create_simple_canvas_pdf(self, docx_path: str, output_path: str) -> bool:
        """Simple canvas-based PDF creation with manual formatting"""
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import A4
            
            doc = Document(docx_path)
            
            c = canvas.Canvas(output_path, pagesize=A4)
            width, height = A4
            
            # Set up fonts and spacing
            c.setFont("Helvetica", 11)
            y_position = height - 50  # Start near top
            line_height = 16
            page_number = 1
            
            # Track records for page breaks
            record_count = 0
            lines_since_break = 0
            
            def add_page_break():
                nonlocal y_position, page_number
                c.showPage()
                c.setFont("Helvetica", 11)
                y_position = height - 50
                page_number += 1
                print(f"   📄 Started page {page_number}")
            
            def write_line(text, bold=False):
                nonlocal y_position, lines_since_break
                
                if y_position < 50:  # Near bottom of page
                    add_page_break()
                
                if text.strip():
                    font_name = "Helvetica-Bold" if bold else "Helvetica"
                    c.setFont(font_name, 11)
                    
                    # Handle long lines by wrapping
                    max_width = width - 100  # Leave margins
                    words = text.split()
                    current_line = ""
                    
                    for word in words:
                        test_line = current_line + " " + word if current_line else word
                        if c.stringWidth(test_line, font_name, 11) < max_width:
                            current_line = test_line
                        else:
                            if current_line:
                                c.drawString(50, y_position, current_line)
                                y_position -= line_height
                                lines_since_break += 1
                                if y_position < 50:
                                    add_page_break()
                            current_line = word
                    
                    if current_line:
                        c.drawString(50, y_position, current_line)
                        y_position -= line_height
                        lines_since_break += 1
                else:
                    # Empty line
                    y_position -= line_height
                    lines_since_break += 1
            
            # Process document content
            for paragraph in doc.paragraphs:
                para_text = paragraph.text.strip()
                
                if para_text:
                    # Detect potential new records
                    record_starters = ['Invoice', 'Dear', 'Hello', 'Hi', 'Welcome']
                    is_new_record = any(para_text.startswith(starter) for starter in record_starters)
                    
                    # Add page break for new records (except first)
                    if is_new_record and record_count > 0 and lines_since_break > 5:
                        add_page_break()
                        lines_since_break = 0
                    
                    # Determine if this should be bold (headings)
                    is_heading = any(heading in para_text.lower() for heading in ['invoice', 'title', 'heading'])
                    
                    write_line(para_text, bold=is_heading)
                    
                    if is_new_record:
                        record_count += 1
                    
                    # Add extra space after certain types of content
                    if any(end in para_text.lower() for end in ['regards', 'sincerely', 'company']):
                        write_line("")  # Extra spacing
                else:
                    write_line("")  # Preserve empty paragraphs as spacing
            
            # Add table content
            for table in doc.tables:
                write_line("--- Table ---", bold=True)
                for row in table.rows:
                    row_data = []
                    for cell in row.cells:
                        if cell.text.strip():
                            row_data.append(cell.text.strip())
                    if row_data:
                        write_line(" | ".join(row_data))
                write_line("")  # Space after table
            
            c.save()
            
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                print(f"✅ Simple PDF created: {output_path} ({os.path.getsize(output_path):,} bytes)")
                print(f"   📄 {page_number} pages, {record_count} records detected")
                return True
            else:
                print("❌ Simple PDF creation failed")
                return False
                
        except Exception as e:
            print(f"❌ Simple PDF creation failed: {str(e)}")
            return False
    
    def convert_docx_to_html(self, docx_path: str) -> str:
        """Convert a single docx file to HTML using mammoth with enhanced styling and line break preservation"""
        try:
            with open(docx_path, "rb") as docx_file:
                # Enhanced style map to preserve more Word formatting
                style_map = """
                p[style-name='Normal'] => p.normal:fresh
                p[style-name='Heading 1'] => h1:fresh
                p[style-name='Heading 2'] => h2:fresh
                p[style-name='Heading 3'] => h3:fresh
                p[style-name='Heading 4'] => h4:fresh
                p[style-name='Title'] => h1.title:fresh
                p[style-name='Subtitle'] => h2.subtitle:fresh
                r[style-name='Strong'] => strong
                r[style-name='Emphasis'] => em
                table => table.word-table
                """
                
                # Convert with enhanced options
                result = mammoth.convert_to_html(
                    docx_file,
                    style_map=style_map,
                    include_default_style_map=True,
                    convert_image=mammoth.images.img_element,  # Preserve images
                    include_embedded_style_map=True  # Include Word's embedded styles
                )
                
                html_content = result.value
                
                # Post-process HTML to improve line break preservation
                # Replace multiple spaces with non-breaking spaces where appropriate
                import re
                
                # Preserve multiple spaces (common in addresses, forms, etc.)
                html_content = re.sub(r'  +', lambda m: '&nbsp;' * len(m.group()), html_content)
                
                # Ensure empty paragraphs are preserved
                html_content = re.sub(r'<p></p>', '<p>&nbsp;</p>', html_content)
                html_content = re.sub(r'<p>\s*</p>', '<p>&nbsp;</p>', html_content)
                
                # Add enhanced CSS styling that matches Word formatting better
                enhanced_html = f"""
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset="utf-8">
                    <style>
                        body {{
                            font-family: 'Calibri', 'Arial', sans-serif;
                            font-size: 11pt;
                            line-height: 1.15;  /* Match Word's default line spacing */
                            color: #000000;
                            margin: 72pt;
                            background: white;
                            text-rendering: optimizeLegibility;
                        }}
                        
                        /* Paragraph styles that preserve Word spacing */
                        p, p.normal {{
                            margin: 0 0 8pt 0;  /* Match Word's default paragraph spacing */
                            text-align: left;
                            white-space: pre-wrap;  /* Preserve whitespace and line breaks */
                            word-wrap: break-word;
                        }}
                        
                        /* Empty paragraphs create proper spacing */
                        p:empty, p.normal:empty {{
                            height: 8pt;
                            margin: 0;
                        }}
                        
                        /* Headings with Word-like spacing */
                        h1, h2, h3, h4, h5, h6 {{
                            color: #1f497d;
                            font-weight: bold;
                            margin-top: 16pt;
                            margin-bottom: 8pt;
                        }}
                        
                        h1, h1.title {{
                            font-size: 18pt;
                            line-height: 1.15;
                        }}
                        
                        h2, h2.subtitle {{
                            font-size: 14pt;
                            line-height: 1.15;
                        }}
                        
                        h3 {{
                            font-size: 12pt;
                            line-height: 1.15;
                        }}
                        
                        h4, h5, h6 {{
                            font-size: 11pt;
                            line-height: 1.15;
                        }}
                        
                        /* Text formatting */
                        strong, b {{
                            font-weight: bold;
                            color: inherit;
                        }}
                        
                        em, i {{
                            font-style: italic;
                            color: inherit;
                        }}
                        
                        /* Table formatting to match Word tables */
                        table, table.word-table {{
                            border-collapse: collapse;
                            width: 100%;
                            margin: 8pt 0;
                            font-size: inherit;
                            page-break-inside: avoid;
                        }}
                        
                        td, th {{
                            border: 1px solid #000000;
                            padding: 6pt 8pt;
                            text-align: left;
                            vertical-align: top;
                            white-space: pre-wrap;  /* Preserve formatting in cells */
                        }}
                        
                        th {{
                            background-color: #d9d9d9;
                            font-weight: bold;
                        }}
                        
                        /* Page breaks */
                        .page-break {{
                            page-break-before: always;
                        }}
                        
                        /* Address blocks and structured content */
                        .address-block {{
                            margin-bottom: 16pt;
                        }}
                        
                        /* Document sections */
                        .document-section {{
                            margin-bottom: 16pt;
                        }}
                        
                        /* Preserve any inline styles from Word */
                        span[style], p[style], div[style] {{
                            /* Inline styles from Word will be preserved */
                        }}
                        
                        /* Non-breaking spaces and preserved whitespace */
                        .preserve-whitespace {{
                            white-space: pre-wrap;
                        }}
                    </style>
                </head>
                <body>
                    {html_content}
                </body>
                </html>
                """
                
                # Log any conversion messages for debugging
                if result.messages:
                    print("⚠️  Conversion messages:")
                    for message in result.messages[:5]:  # Limit to first 5 messages
                        print(f"   {message}")
                
                return enhanced_html
                
        except Exception as e:
            print(f"❌ Error converting docx to HTML: {str(e)}")
            import traceback
            traceback.print_exc()
            return ""

    def convert_docx_to_html_enhanced(self, docx_path: str) -> str:
        """Enhanced DOCX to HTML conversion with better Linux compatibility and spacing preservation"""
        try:
            with open(docx_path, "rb") as docx_file:
                # Enhanced style map for better formatting preservation
                style_map = """
                p[style-name='Normal'] => p.paragraph:fresh
                p[style-name='Heading 1'] => h1:fresh
                p[style-name='Heading 2'] => h2:fresh
                p[style-name='Heading 3'] => h3:fresh
                p[style-name='Heading 4'] => h4:fresh
                p[style-name='Title'] => h1.title:fresh
                p[style-name='Subtitle'] => h2.subtitle:fresh
                r[style-name='Strong'] => span.bold
                r[style-name='Emphasis'] => span.italic
                r[style-name='Intense Emphasis'] => span.bold.italic
                table => table
                """
                
                # Convert with enhanced options
                result = mammoth.convert_to_html(
                    docx_file,
                    style_map=style_map,
                    include_default_style_map=True,
                    convert_image=mammoth.images.img_element,
                    include_embedded_style_map=True
                )
                
                html_content = result.value
                
                # Enhanced post-processing for Linux compatibility
                import re
                
                # Preserve spaces and line breaks more aggressively
                html_content = re.sub(r'\s{2,}', lambda m: '&nbsp;' * (len(m.group()) - 1) + ' ', html_content)
                
                # Ensure paragraphs have content (prevent collapse)
                html_content = re.sub(r'<p>\s*</p>', '<p>&nbsp;</p>', html_content)
                
                # Add word spacing to text nodes to prevent running together
                html_content = re.sub(r'>([^<]+)<', lambda m: f'>{m.group(1).replace(" ", "&nbsp;")}<' if m.group(1).count(" ") == 1 else f'>{m.group(1)}<', html_content)
                
                # Wrap in paragraphs with proper spacing class
                if not html_content.strip().startswith('<'):
                    html_content = f'<div class="paragraph">{html_content}</div>'
                
                return html_content
                
        except Exception as e:
            print(f"❌ Error in enhanced DOCX to HTML conversion: {str(e)}")
            return ""

    def convert_html_to_pdf_enhanced(self, html_content: str, pdf_path: str) -> bool:
        """Enhanced HTML to PDF conversion with better Linux compatibility"""
        try:
            from weasyprint import HTML, CSS
            from weasyprint.text.fonts import FontConfiguration
            import tempfile
            import os
            
            # Create font configuration for better rendering on Linux
            font_config = FontConfiguration()
            
            # Enhanced CSS for Linux rendering
            css_content = CSS(string="""
            @page {
                size: A4;
                margin: 1in;
            }
            
            body {
                font-family: 'DejaVu Sans', 'Liberation Sans', Arial, sans-serif;
                font-size: 11pt;
                line-height: 1.6;
                color: black;
                text-rendering: optimizeLegibility;
                -webkit-font-smoothing: antialiased;
                word-spacing: 0.25em;
            }
            
            p, .paragraph {
                margin: 0 0 12pt 0;
                padding: 0;
                white-space: pre-line;
                word-break: normal;
                overflow-wrap: break-word;
            }
            
            .bold { font-weight: bold; }
            .italic { font-style: italic; }
            .underline { text-decoration: underline; }
            
            h1, h2, h3, h4, h5, h6 {
                margin: 16pt 0 8pt 0;
                font-weight: bold;
                page-break-after: avoid;
            }
            
            table {
                border-collapse: collapse;
                width: 100%;
                margin: 8pt 0;
                page-break-inside: avoid;
            }
            
            td, th {
                border: 1px solid #000;
                padding: 4pt 6pt;
                text-align: left;
                vertical-align: top;
                word-break: normal;
            }
            """, font_config=font_config)
            
            # Write HTML to temporary file for processing
            with tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8') as temp_html:
                temp_html.write(html_content)
                temp_html_path = temp_html.name
            
            try:
                # Generate PDF with enhanced settings
                html_doc = HTML(filename=temp_html_path, encoding='utf-8')
                document = html_doc.render(stylesheets=[css_content], font_config=font_config)
                document.write_pdf(pdf_path)
                
                # Verify the PDF was created successfully
                if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
                    print(f"   ✅ Enhanced HTML to PDF conversion successful")
                    return True
                else:
                    print(f"   ❌ Generated PDF is too small or missing")
                    return False
                    
            finally:
                # Clean up temporary HTML file
                try:
                    os.unlink(temp_html_path)
                except:
                    pass
                    
        except ImportError:
            print(f"   ❌ WeasyPrint not available")
            return False
        except Exception as e:
            print(f"   ❌ Enhanced HTML to PDF conversion error: {str(e)}")
            return False

    def create_basic_text_pdf_from_docx_enhanced(self, docx_path: str, pdf_path: str) -> bool:
        """Enhanced basic text PDF creation with better formatting and spacing"""
        try:
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
            from reportlab.lib.units import inch
            from reportlab.lib import colors
            import docx
            
            # Read the docx file
            doc = docx.Document(docx_path)
            
            # Create PDF with enhanced formatting
            pdf_doc = SimpleDocTemplate(
                pdf_path,
                pagesize=letter,
                rightMargin=72,
                leftMargin=72,
                topMargin=72,
                bottomMargin=18,
                title="Generated Document"
            )
            
            # Create enhanced styles
            styles = getSampleStyleSheet()
            
            # Custom paragraph style with proper spacing
            normal_style = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=11,
                leading=14,  # 1.27 line spacing
                spaceAfter=12,
                spaceBefore=0,
                wordWrap='LTR',
                alignment=0,  # Left alignment
                fontName='Helvetica'
            )
            
            heading_style = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading1'],
                fontSize=14,
                leading=18,
                spaceAfter=12,
                spaceBefore=18,
                fontName='Helvetica-Bold'
            )
            
            # Process document content
            story = []
            
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if not text:
                    # Add space for empty paragraphs
                    story.append(Spacer(1, 6))
                    continue
                
                # Simple formatting detection
                is_heading = any(run.font.size and run.font.size.pt > 12 for run in paragraph.runs if run.font.size)
                is_bold = any(run.bold for run in paragraph.runs if run.bold)
                
                # Apply appropriate style
                if is_heading or is_bold:
                    para = Paragraph(text, heading_style)
                else:
                    para = Paragraph(text, normal_style)
                
                story.append(para)
            
            # Handle tables
            for table in doc.tables:
                story.append(Spacer(1, 12))
                for row in table.rows:
                    row_text = " | ".join(cell.text.strip() for cell in row.cells)
                    if row_text.strip():
                        para = Paragraph(row_text, normal_style)
                        story.append(para)
                story.append(Spacer(1, 12))
            
            # Build the PDF
            pdf_doc.build(story)
            
            # Verify creation
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
                print(f"   ✅ Enhanced basic PDF created successfully")
                return True
            else:
                print(f"   ❌ Failed to create enhanced basic PDF")
                return False
                
        except ImportError as ie:
            print(f"   ❌ Missing required libraries: {ie}")
            return False
        except Exception as e:
            print(f"   ❌ Enhanced basic PDF creation error: {str(e)}")
            return False

    def generate_single_pdf(self, output_path: str) -> bool:
        """Generate a single PDF by creating a Word file then converting it to PDF"""
        try:
            print(f"📄 Creating single PDF document with {len(self.data)} records...")
            
            # Step 1: Use the proven "1 Word file" method
            print("🔄 Step 1: Creating Word document (using proven '1 Word file' method)...")
            temp_word_file = tempfile.NamedTemporaryFile(suffix='.docx', delete=False).name
            
            if not self.generate_single_word(temp_word_file):
                print("❌ Failed to create Word document")
                try:
                    os.unlink(temp_word_file)
                except:
                    pass
                return False
            
            print(f"✅ Step 1 Complete: Word document created ({os.path.getsize(temp_word_file):,} bytes)")
            
            # Step 2: Convert to PDF using simplest available method
            print("🔄 Step 2: Converting to PDF...")
            success = self.convert_word_to_pdf_simple(temp_word_file, output_path)
            
            # Clean up temp file
            try:
                os.unlink(temp_word_file)
            except:
                pass
            
            if success:
                print(f"✅ SUCCESS: Single PDF created at {output_path}")
            else:
                print("❌ Failed to convert Word document to PDF")
            
            return success
            
        except Exception as e:
            print(f"❌ Error in single PDF generation: {str(e)}")
            return False
    def convert_docx_to_pdf_preserve_formatting(self, docx_path: str, pdf_path: str) -> bool:
        """Convert Word document to PDF while preserving formatting using python-docx + reportlab"""
        try:
            print("🔄 Converting Word to PDF with preserved formatting...")
            
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import letter, A4
            from reportlab.lib.colors import black, blue, red
            from reportlab.lib.units import inch
            from docx import Document
            from docx.shared import RGBColor
            
            # Read the Word document
            doc = Document(docx_path)
            
            # Create PDF
            c = canvas.Canvas(pdf_path, pagesize=A4)
            width, height = A4
            
            # Starting position
            x_margin = 72  # 1 inch margin
            y_position = height - 72  # Start 1 inch from top
            line_height = 14
            
            print(f"   Processing {len(doc.paragraphs)} paragraphs...")
            
            for para in doc.paragraphs:
                # Check for page break
                if y_position < 72:  # Less than 1 inch from bottom
                    c.showPage()
                    y_position = height - 72
                
                # Handle empty paragraphs (line breaks)
                if not para.text.strip():
                    y_position -= line_height
                    continue
                
                # Process paragraph with formatting
                para_text = ""
                x_position = x_margin
                
                # Check if paragraph has runs with different formatting
                if len(para.runs) > 0:
                    for run in para.runs:
                        if run.text:
                            # Set font properties based on run formatting
                            font_name = "Helvetica"
                            font_size = 11
                            
                            # Handle bold
                            if run.bold:
                                font_name = "Helvetica-Bold"
                            
                            # Handle italic
                            if run.italic:
                                if run.bold:
                                    font_name = "Helvetica-BoldOblique"
                                else:
                                    font_name = "Helvetica-Oblique"
                            
                            # Set font color
                            text_color = black
                            if run.font.color and run.font.color.rgb:
                                rgb = run.font.color.rgb
                                text_color = (rgb.red/255.0, rgb.green/255.0, rgb.blue/255.0)
                            
                            # Apply formatting and draw text
                            c.setFont(font_name, font_size)
                            c.setFillColor(text_color)
                            c.drawString(x_position, y_position, run.text)
                            
                            # Update x position for next run
                            text_width = c.stringWidth(run.text, font_name, font_size)
                            x_position += text_width
                
                else:
                    # Simple paragraph without runs
                    c.setFont("Helvetica", 11)
                    c.setFillColor(black)
                    c.drawString(x_margin, y_position, para.text)
                
                y_position -= line_height
            
            # Process tables
            for table in doc.tables:
                if y_position < 200:  # Need space for table
                    c.showPage()
                    y_position = height - 72
                
                print(f"   Processing table with {len(table.rows)} rows...")
                
                # Calculate column widths
                available_width = width - (2 * x_margin)
                col_width = available_width / len(table.columns) if len(table.columns) > 0 else 100
                
                # Draw table
                table_y = y_position
                for row_idx, row in enumerate(table.rows):
                    row_x = x_margin
                    
                    for col_idx, cell in enumerate(row.cells):
                        # Draw cell border
                        c.rect(row_x, table_y - 20, col_width, 20, stroke=1, fill=0)
                        
                        # Draw cell text
                        if cell.text:
                            c.setFont("Helvetica", 9)
                            c.setFillColor(black)
                            # Truncate text if too long
                            text = cell.text[:int(col_width/6)] + "..." if len(cell.text) > col_width/6 else cell.text
                            c.drawString(row_x + 2, table_y - 15, text)
                        
                        row_x += col_width
                    
                    table_y -= 20
                    
                    # Check if we need a new page
                    if table_y < 72:
                        c.showPage()
                        table_y = height - 72
                
                y_position = table_y - 20
            
            # Save the PDF
            c.save()
            
            # Verify PDF was created
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                print(f"✅ PDF created with preserved formatting: {pdf_path} ({os.path.getsize(pdf_path):,} bytes)")
                return True
            else:
                print("❌ Failed to create PDF")
                return False
                
        except Exception as e:
            print(f"❌ Error in formatted PDF conversion: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
            
            # Clean up temp file
            try:
                os.unlink(temp_word_file)
            except:
                pass
            
            return success
            
        except Exception as e:
            print(f"❌ Error in single PDF generation: {str(e)}")
            return False
    
    def generate_multiple_pdf(self, output_dir: str) -> bool:
        """Generate multiple PDFs by first creating Word files, then converting each to PDF"""
        try:
            print(f"📑 Creating multiple PDF documents for {len(self.data)} records...")
            
            # Step 1: Create Word files using the PROVEN working method
            print("🔄 Step 1: Creating Word files (same as 'Multiple Word files' option)...")
            temp_word_dir = tempfile.mkdtemp(prefix='mailmerge_word_')
            
            # Use the exact same method that works for "Multiple Word files"
            if not self.generate_multiple_word(temp_word_dir):
                print("❌ Failed to create Word files")
                try:
                    import shutil
                    shutil.rmtree(temp_word_dir, ignore_errors=True)
                except:
                    pass
                return False
            
            print(f"✅ Step 1 Complete: Word files created successfully")
            
            # Step 2: Convert each Word file to PDF
            print("🔄 Step 2: Converting each Word file to PDF...")
            os.makedirs(output_dir, exist_ok=True)
            
            # Get all Word files
            word_files = [f for f in os.listdir(temp_word_dir) if f.endswith('.docx')]
            successful_conversions = 0
            
            for word_file in word_files:
                word_path = os.path.join(temp_word_dir, word_file)
                pdf_file = word_file.replace('.docx', '.pdf')
                pdf_path = os.path.join(output_dir, pdf_file)
                
                print(f"   Converting: {word_file} → {pdf_file}")
                
                if self.convert_word_to_pdf_simple(word_path, pdf_path):
                    print(f"   ✅ SUCCESS: Converted {word_file}")
                    successful_conversions += 1
                else:
                    print(f"   ❌ Failed to convert {word_file}")
            
            # Clean up temp Word files
            try:
                import shutil
                shutil.rmtree(temp_word_dir, ignore_errors=True)
                print("🗑️  Cleaned up temporary Word files")
            except:
                pass
            
            if successful_conversions > 0:
                print(f"🎉 SUCCESS: {successful_conversions}/{len(word_files)} PDFs created")
                return True
            else:
                print("❌ No PDF files were created")
                return False
                
        except Exception as e:
            print(f"❌ Error in multiple PDF generation: {str(e)}")
            return False
    
    def process_merge(self, output_format: str, output_path: str) -> bool:
        """Main processing function"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Both template and data files must be loaded")
            
            # Process based on output format
            if output_format == "single-word":
                return self.generate_single_word(output_path)
            elif output_format == "multiple-word":
                return self.generate_multiple_word(output_path)
            elif output_format == "single-pdf":
                return self.generate_single_pdf(output_path)
            elif output_format == "multiple-pdf":
                return self.generate_multiple_pdf(output_path)
            else:
                raise ValueError(f"Unsupported output format: {output_format}")
                
        except Exception as e:
            print(f"Error processing mail merge: {str(e)}")
            return False

# Store processors per session
processors = {}

def get_processor():
    """Get or create processor for current session"""
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
        print(f"🆕 Created new session: {session['session_id']}")
    
    session_id = session['session_id']
    print(f"🔄 Using session: {session_id}")
    
    if session_id not in processors:
        processors[session_id] = MailMergeProcessor(session_id)
        print(f"🆕 Created new processor for session: {session_id}")
    else:
        print(f"♻️  Reusing existing processor for session: {session_id}")
    
    print(f"📊 Total active processors: {len(processors)}")
    return processors[session_id]

def cleanup_old_processors():
    """Clean up old processors (simple cleanup)"""
    if len(processors) > 50:  # Clean up if too many processors
        old_sessions = list(processors.keys())[:25]
        for session_id in old_sessions:
            processors[session_id].cleanup()
            del processors[session_id]

# Flask Routes
@app.route('/')
def index():
    """Serve the main page"""
    try:
        with open('index.html', 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return "<h1>Mail Merge SaaS</h1><p>Main page not found. Please upload index.html</p>"

@app.route('/mailmerge')
def mailmerge():
    """Serve the mail merge page"""
    try:
        with open('mailmerge.html', 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return "<h1>Mail Merge</h1><p>Mail merge page not found. Please upload mailmerge.html</p>"

@app.route('/style.css')
def serve_css():
    """Serve CSS file"""
    try:
        with open('style.css', 'r', encoding='utf-8') as f:
            css_content = f.read()
        response = app.response_class(
            response=css_content,
            status=200,
            mimetype='text/css'
        )
        return response
    except FileNotFoundError:
        return "/* CSS file not found */", 404

@app.route('/mailmerge.js')
def serve_js():
    """Serve JavaScript file"""
    try:
        with open('mailmerge.js', 'r', encoding='utf-8') as f:
            js_content = f.read()
        response = app.response_class(
            response=js_content,
            status=200,
            mimetype='application/javascript'
        )
        return response
    except FileNotFoundError:
        return "/* JavaScript file not found */", 404

@app.route('/upload_template', methods=['POST'])
def upload_template():
    """Handle template file upload"""
    try:
        print("Template upload request received")
        cleanup_old_processors()
        
        processor = get_processor()
        
        if 'file' not in request.files:
            print("No file in request")
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        file = request.files['file']
        if file.filename == '':
            print("Empty filename")
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        print(f"Template file: {file.filename}")
        
        if not allowed_file(file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
            print(f"Invalid file type: {file.filename}")
            return jsonify({'success': False, 'error': 'Invalid file type. Please upload a .docx file'}), 400
        
        # Create unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"template_{processor.session_id}_{timestamp}_{secure_filename(file.filename)}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        print(f"Template saved to: {filepath}")
        
        # Load template
        if processor.load_template(filepath):
            print(f"✅ Template loaded successfully for session {processor.session_id}")
            print(f"   Template path set to: {processor.template_path}")
            print(f"   File exists: {os.path.exists(processor.template_path)}")
            return jsonify({
                'success': True,
                'message': f'Template uploaded successfully: {file.filename}',
                'filepath': filepath,
                'filename': file.filename
            })
        else:
            print(f"❌ Failed to load template for session {processor.session_id}")
            if os.path.exists(filepath):
                os.remove(filepath)
            return jsonify({'success': False, 'error': 'Invalid template file'}), 400
            
    except Exception as e:
        print(f"Template upload error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/upload_data', methods=['POST'])
def upload_data():
    """Handle data file upload"""
    try:
        print("Data upload request received")
        cleanup_old_processors()
        
        processor = get_processor()
        
        if 'file' not in request.files:
            print("No file in request")
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        file = request.files['file']
        if file.filename == '':
            print("Empty filename")
            return jsonify({'success': False, 'error': 'No file selected'}), 400
        
        print(f"Data file: {file.filename}")
        
        if not allowed_file(file.filename, ALLOWED_DATA_EXTENSIONS):
            print(f"Invalid file type: {file.filename}")
            return jsonify({'success': False, 'error': 'Invalid file type. Please upload an Excel file (.xlsx)'}), 400
        
        # Create unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        filename = f"data_{processor.session_id}_{timestamp}_{secure_filename(file.filename)}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        print(f"Data saved to: {filepath}")
        
        # Load data
        if processor.load_data(filepath):
            # Return preview of data
            preview_data = processor.data[:3]  # First 3 rows
            columns = list(processor.data[0].keys()) if processor.data else []
            total_rows = len(processor.data)
            
            return jsonify({
                'success': True,
                'message': f'Data uploaded successfully: {file.filename}',
                'filepath': filepath,
                'filename': file.filename,
                'preview': preview_data,
                'columns': columns,
                'total_rows': total_rows
            })
        else:
            if os.path.exists(filepath):
                os.remove(filepath)
            return jsonify({'success': False, 'error': 'Invalid data file'}), 400
            
    except Exception as e:
        print(f"Data upload error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/check_status', methods=['GET'])
def check_status():
    """Check current upload status - with fallback file checking"""
    try:
        processor = get_processor()
        
        # Debug logging
        print(f"🔍 Status check for session: {processor.session_id}")
        print(f"   Template path: {processor.template_path}")
        print(f"   Template exists: {processor.template_path and os.path.exists(processor.template_path) if processor.template_path else False}")
        print(f"   Data loaded: {len(processor.data) if processor.data else 0} records")
        
        # Fallback: If processor doesn't have files, check for recent uploads
        template_loaded = processor.template_path is not None and os.path.exists(processor.template_path) if processor.template_path else False
        data_loaded = processor.data_path is not None and len(processor.data) > 0
        
        # If files not found in processor, check for recent uploads in the directory
        if not template_loaded or not data_loaded:
            print("🔍 Checking for recent uploads as fallback...")
            if os.path.exists(app.config['UPLOAD_FOLDER']):
                files = os.listdir(app.config['UPLOAD_FOLDER'])
                print(f"   Found files: {files}")
                
                # Look for recent template files
                if not template_loaded:
                    for file in files:
                        if file.startswith(f'template_{processor.session_id}') and file.endswith('.docx'):
                            template_path = os.path.join(app.config['UPLOAD_FOLDER'], file)
                            if processor.load_template(template_path):
                                template_loaded = True
                                print(f"✅ Recovered template: {file}")
                                break
                
                # Look for recent data files  
                if not data_loaded:
                    for file in files:
                        if file.startswith(f'data_{processor.session_id}') and file.endswith('.xlsx'):
                            data_path = os.path.join(app.config['UPLOAD_FOLDER'], file)
                            if processor.load_data(data_path):
                                data_loaded = True
                                print(f"✅ Recovered data: {file}")
                                break
        
        result = {
            'template_loaded': template_loaded,
            'data_loaded': data_loaded,
            'template_path': processor.template_path,
            'data_records': len(processor.data) if processor.data else 0,
            'session_id': processor.session_id  # Add for debugging
        }
        
        print(f"   Returning status: {result}")
        return jsonify(result)
        
    except Exception as e:
        print(f"❌ Status check error: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/process_merge', methods=['POST'])
def process_merge():
    """Process the mail merge - Support both Word and PDF formats"""
    try:
        print("Process merge request received")
        
        processor = get_processor()
        data = request.get_json()
        output_format = data.get('format', 'single-word')
        
        print(f"Output format: {output_format}")
        print(f"Template loaded: {processor.template_path is not None}")
        print(f"Data loaded: {len(processor.data) if processor.data else 0} records")
        
        if not processor.template_path or not processor.data:
            error_msg = f"Missing files - Template: {processor.template_path is not None}, Data: {len(processor.data) if processor.data else 0} records"
            print(error_msg)
            return jsonify({'success': False, 'error': 'Please upload both template and data files first'}), 400
        
        # Generate unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        
        # Determine file extension and format type
        is_pdf = 'pdf' in output_format
        is_single = 'single' in output_format
        file_ext = '.pdf' if is_pdf else '.docx'
        format_name = 'PDF' if is_pdf else 'Word'
        
        if is_single:
            # Single file output
            output_filename = f"mailmerge_result_{processor.session_id}_{timestamp}{file_ext}"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            
            if processor.process_merge(output_format, output_path):
                return jsonify({
                    'success': True,
                    'message': f'Mail merge completed successfully! {format_name} formatting preserved.',
                    'download_url': f'/download/{output_filename}',
                    'filename': output_filename
                })
            else:
                return jsonify({'success': False, 'error': 'Failed to process mail merge'}), 500
                
        else:  # multiple files
            # Multiple files - create ZIP
            output_dir = os.path.join(OUTPUT_FOLDER, f"mailmerge_{processor.session_id}_{timestamp}")
            zip_filename = f"mailmerge_results_{processor.session_id}_{timestamp}.zip"
            zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
            
            if processor.process_merge(output_format, output_dir):
                # Create ZIP file
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for root, dirs, files in os.walk(output_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, output_dir)
                            zipf.write(file_path, arcname)
                
                # Clean up individual files
                shutil.rmtree(output_dir)
                
                return jsonify({
                    'success': True,
                    'message': f'Mail merge completed! Generated {len(processor.data)} {format_name} documents with preserved formatting.',
                    'download_url': f'/download/{zip_filename}',
                    'filename': zip_filename
                })
            else:
                return jsonify({'success': False, 'error': 'Failed to process mail merge'}), 500
                
    except Exception as e:
        print(f"Process merge error: {str(e)}")
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download processed file"""
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health_check():
    """Health check endpoint"""
    return jsonify({
        'status': 'healthy', 
        'service': 'Mail Merge SaaS - Word & PDF Support',
        'active_sessions': len(processors),
        'upload_folder': app.config['UPLOAD_FOLDER'],
        'output_folder': OUTPUT_FOLDER
    })

@app.route('/debug')
def debug_info():
    """Debug endpoint to check application state"""
    try:
        processor = get_processor()
        
        # List files in upload directory
        upload_files = []
        if os.path.exists(app.config['UPLOAD_FOLDER']):
            upload_files = os.listdir(app.config['UPLOAD_FOLDER'])
        
        return jsonify({
            'session_id': processor.session_id,
            'session_data': dict(session),
            'processors_count': len(processors),
            'processor_template_path': processor.template_path,
            'processor_data_records': len(processor.data) if processor.data else 0,
            'upload_folder': app.config['UPLOAD_FOLDER'],
            'upload_files': upload_files,
            'template_file_exists': os.path.exists(processor.template_path) if processor.template_path else False
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    # Development server
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)