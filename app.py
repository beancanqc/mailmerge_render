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
import openpyxl
import re
from typing import List, Dict, Any, Optional
import html
from io import BytesIO
import html

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
app.secret_key = 'your-secret-key-change-this'  # Add secret key for sessions

# Configure folders
UPLOAD_FOLDER = tempfile.mkdtemp()
OUTPUT_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_TEMPLATE_EXTENSIONS = {'docx'}
ALLOWED_DATA_EXTENSIONS = {'xlsx'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

def check_pdf_conversion_capabilities():
    """Check what PDF conversion methods are available"""
    capabilities = []
    
    # Check operating system
    import platform
    print(f"Operating System: {platform.system()}")
    
    # Check for docx2pdf (Word)
    try:
        import docx2pdf
        capabilities.append("Word native (docx2pdf)")
        print("Word native conversion available")
    except ImportError:
        print("Word native conversion not available")
    
    # Check for LibreOffice
    try:
        import subprocess
        
        if platform.system() == "Windows":
            # Try common Windows LibreOffice paths
            libreoffice_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe", 
                "soffice.exe",
                "libreoffice"
            ]
            
            libre_found = False
            for path in libreoffice_paths:
                try:
                    result = subprocess.run([path, '--version'], 
                                          capture_output=True, text=True, timeout=5)
                    if result.returncode == 0:
                        capabilities.append("LibreOffice")
                        print(f"LibreOffice available at: {path}")
                        libre_found = True
                        break
                except:
                    continue
                    
            if not libre_found:
                print("LibreOffice not found on Windows")
        else:
            # Linux/macOS
            try:
                result = subprocess.run(['libreoffice', '--version'], 
                                      capture_output=True, text=True, timeout=5)
                if result.returncode == 0:
                    capabilities.append("LibreOffice")
                    print("LibreOffice available")
                    print(f"   Version: {result.stdout.strip()}")
                else:
                    print(f"LibreOffice command failed: {result.stderr}")
                    
                    # Try alternative commands for Render
                    for alt_cmd in ['/usr/bin/libreoffice', '/usr/local/bin/libreoffice', 'soffice']:
                        try:
                            result = subprocess.run([alt_cmd, '--version'], 
                                                  capture_output=True, text=True, timeout=5)
                            if result.returncode == 0:
                                capabilities.append("LibreOffice")
                                print(f"LibreOffice available at: {alt_cmd}")
                                print(f"   Version: {result.stdout.strip()}")
                                break
                        except:
                            continue
                    else:
                        print("LibreOffice not found with any command")
            except FileNotFoundError:
                print("LibreOffice command not found")
            except Exception as e:
                print(f"LibreOffice check error: {e}")
    except:
        print("LibreOffice not available")
    
    # ReportLab is always available
    capabilities.append("ReportLab (basic)")
    print("ReportLab fallback available")
    
    print(f"PDF conversion capabilities: {', '.join(capabilities)}")
    return capabilities

class MailMergeProcessor:
    def __init__(self, session_id=None):
        self.session_id = session_id or str(uuid.uuid4())
        self.template_path: Optional[str] = None
        self.data_path: Optional[str] = None
        self.data: List[Dict[str, Any]] = []
        
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
        """Advanced merge field replacement that preserves both Word styles and character formatting"""
        try:
            full_text = paragraph.text
            
            # Find all merge fields
            merge_fields = re.finditer(r'\{\{(\w+)\}\}', full_text)
            merge_list = list(merge_fields)
            
            if not merge_list:
                return
            
            # PRESERVE the original paragraph style - this is crucial!
            original_style = paragraph.style
            original_alignment = paragraph.alignment
            
            # Try the advanced formatting preservation approach
            try:
                # Store all runs with their formatting before we modify anything
                original_runs = []
                for run in paragraph.runs:
                    try:
                        original_runs.append({
                            'text': run.text,
                            'bold': run.bold,
                            'italic': run.italic,
                            'underline': run.underline,
                            'font_name': run.font.name,
                            'font_size': run.font.size,
                            'font_color': run.font.color.rgb if run.font.color.rgb else None
                        })
                    except Exception as e:
                        # If we can't read formatting, store basic info
                        original_runs.append({
                            'text': run.text,
                            'bold': False,
                            'italic': False,
                            'underline': False,
                            'font_name': None,
                            'font_size': None,
                            'font_color': None
                        })
                
                # Build character position to formatting map
                char_to_formatting = {}
                char_pos = 0
                for run_info in original_runs:
                    run_text = run_info['text']
                    for i in range(len(run_text)):
                        char_to_formatting[char_pos + i] = run_info
                    char_pos += len(run_text)
                
                # Process merge fields and rebuild with formatting
                success = self._rebuild_paragraph_with_formatting(paragraph, full_text, merge_list, data_row, char_to_formatting)
                
                if success:
                    # Restore paragraph-level formatting (Word styles)
                    paragraph.style = original_style
                    if original_alignment is not None:
                        paragraph.alignment = original_alignment
                    return
                    
            except Exception as e:
                print(f"Advanced formatting failed, using simple approach: {e}")
            
            # Fallback to simple approach if advanced formatting fails
            print("Using simple merge approach...")
            self._simple_merge_fields(paragraph, full_text, merge_list, data_row, original_style, original_alignment)
            
        except Exception as e:
            print(f"Error in replace_merge_fields_advanced: {e}")
            # Ultimate fallback - just do basic text replacement
            try:
                full_text = paragraph.text
                for match in reversed(list(re.finditer(r'\{\{(\w+)\}\}', full_text))):
                    field_name = match.group(1)
                    replacement_text = str(data_row.get(field_name, ""))
                    full_text = full_text[:match.start()] + replacement_text + full_text[match.end():]
                
                paragraph.clear()
                if full_text.strip():
                    paragraph.add_run(full_text)
            except Exception as final_error:
                print(f"Even basic replacement failed: {final_error}")
    
    def _simple_merge_fields(self, paragraph, full_text, merge_list, data_row, original_style, original_alignment):
        """Simple merge field replacement that preserves paragraph styles"""
        try:
            # Simple replacement that maintains style integrity
            new_text = full_text
            for match in reversed(merge_list):  # Process in reverse to maintain positions
                field_name = match.group(1)
                replacement_text = str(data_row.get(field_name, ""))
                new_text = new_text[:match.start()] + replacement_text + new_text[match.end():]
            
            # Clear the paragraph but keep its style
            paragraph.clear()
            
            # Add the new text as a single run to maintain style integrity
            if new_text.strip():
                run = paragraph.add_run(new_text)
                # The paragraph style should automatically apply to the run
                
            # Restore paragraph-level formatting
            paragraph.style = original_style
            if original_alignment is not None:
                paragraph.alignment = original_alignment
                
        except Exception as e:
            print(f"Simple merge failed: {e}")
            raise
    
    def _rebuild_paragraph_with_formatting(self, paragraph, full_text, merge_list, data_row, char_to_formatting):
        """Rebuild paragraph with detailed formatting preservation"""
        try:
            # Process merge fields in reverse order to maintain positions
            new_text = full_text
            formatting_changes = []  # Track where replacements happen
            
            for match in reversed(merge_list):
                field_name = match.group(1)
                replacement_text = str(data_row.get(field_name, ""))
                start_pos = match.start()
                end_pos = match.end()
                
                # Store the formatting that should apply to the replacement
                if start_pos in char_to_formatting:
                    replacement_formatting = char_to_formatting[start_pos]
                else:
                    # Find the closest formatting
                    replacement_formatting = None
                    for pos in range(start_pos, -1, -1):
                        if pos in char_to_formatting:
                            replacement_formatting = char_to_formatting[pos]
                            break
                
                formatting_changes.append({
                    'start': start_pos,
                    'end': start_pos + len(replacement_text),
                    'formatting': replacement_formatting,
                    'text': replacement_text
                })
                
                # Replace in text
                new_text = new_text[:start_pos] + replacement_text + new_text[end_pos:]
            
            # Clear the paragraph
            paragraph.clear()
            
            # Rebuild the paragraph with preserved formatting
            if new_text.strip():
                current_pos = 0
                
                # Sort formatting changes by position
                formatting_changes.sort(key=lambda x: x['start'])
                
                for change in formatting_changes:
                    # Add text before this replacement (if any)
                    if change['start'] > current_pos:
                        before_text = new_text[current_pos:change['start']]
                        if before_text:
                            # Find original formatting for this position
                            orig_formatting = None
                            for pos in range(current_pos, change['start']):
                                if pos in char_to_formatting:
                                    orig_formatting = char_to_formatting[pos]
                                    break
                            
                            run = paragraph.add_run(before_text)
                            if orig_formatting:
                                self._apply_preserved_formatting(run, orig_formatting)
                    
                    # Add the replacement text with its formatting
                    if change['text']:
                        run = paragraph.add_run(change['text'])
                        if change['formatting']:
                            self._apply_preserved_formatting(run, change['formatting'])
                    
                    current_pos = change['end']
                
                # Add any remaining text
                if current_pos < len(new_text):
                    remaining_text = new_text[current_pos:]
                    if remaining_text:
                        # Find formatting for remaining text
                        orig_formatting = None
                        for pos in range(current_pos, len(full_text)):
                            if pos in char_to_formatting:
                                orig_formatting = char_to_formatting[pos]
                                break
                        
                        run = paragraph.add_run(remaining_text)
                        if orig_formatting:
                            self._apply_preserved_formatting(run, orig_formatting)
            
            return True
            
        except Exception as e:
            print(f"Detailed formatting rebuild failed: {e}")
            return False
    
    def _apply_preserved_formatting(self, run, formatting_info):
        """Apply preserved formatting to a run"""
        try:
            if formatting_info.get('bold') is not None:
                run.bold = formatting_info['bold']
            if formatting_info.get('italic') is not None:
                run.italic = formatting_info['italic']
            if formatting_info.get('underline') is not None:
                run.underline = formatting_info['underline']
            if formatting_info.get('font_name'):
                run.font.name = formatting_info['font_name']
            if formatting_info.get('font_size'):
                run.font.size = formatting_info['font_size']
            if formatting_info.get('font_color'):
                run.font.color.rgb = formatting_info['font_color']
        except Exception as e:
            print(f"Warning: Could not apply formatting: {e}")

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
        """Generate a single Word document with all records - preserving Word styles"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            from docx.enum.text import WD_BREAK
            from copy import deepcopy
            
            # Start with the template document to preserve all styles and formatting
            merged_doc = Document(self.template_path)
            
            # Replace fields for the first record
            if self.data:
                merged_doc = self.replace_merge_fields(merged_doc, self.data[0])
            
            # Add remaining records while preserving all Word styles
            for row_data in self.data[1:]:
                # Add a page break
                page_break_para = merged_doc.add_paragraph()
                run = page_break_para.add_run()
                run.add_break(WD_BREAK.PAGE)
                
                # Load template and process it
                template_doc = Document(self.template_path)
                processed_doc = self.replace_merge_fields(template_doc, row_data)
                
                # Method 1: Try copying styles and document parts
                try:
                    # Copy styles from template to merged document if not already present
                    for style in processed_doc.styles:
                        try:
                            if style.name not in [s.name for s in merged_doc.styles]:
                                # Add the style to merged document
                                merged_doc.styles.add_style(style.name, style.type)
                                new_style = merged_doc.styles[style.name]
                                # Copy style properties
                                if hasattr(style, 'font'):
                                    new_style.font.name = style.font.name
                                    new_style.font.size = style.font.size
                                    new_style.font.bold = style.font.bold
                                    new_style.font.italic = style.font.italic
                                    new_style.font.underline = style.font.underline
                                    if style.font.color.rgb:
                                        new_style.font.color.rgb = style.font.color.rgb
                        except:
                            pass  # Style might already exist or be built-in
                except Exception as e:
                    print(f"Warning: Could not copy styles: {e}")
                
                # Method 2: Copy content with style preservation
                try:
                    # Copy paragraphs with their complete style information
                    for paragraph in processed_doc.paragraphs:
                        # Create new paragraph
                        new_para = merged_doc.add_paragraph()
                        
                        # Copy paragraph style by name (this preserves Word styles like "Title")
                        try:
                            if paragraph.style.name:
                                # Find the style in the merged document
                                style_found = False
                                for style in merged_doc.styles:
                                    if style.name == paragraph.style.name:
                                        new_para.style = style
                                        style_found = True
                                        break
                                
                                if not style_found:
                                    # If style not found, try to use the original style
                                    new_para.style = paragraph.style
                        except Exception as e:
                            print(f"Warning: Could not copy paragraph style: {e}")
                        
                        # Copy paragraph formatting
                        try:
                            if paragraph.paragraph_format:
                                new_para.paragraph_format.alignment = paragraph.paragraph_format.alignment
                                new_para.paragraph_format.space_before = paragraph.paragraph_format.space_before
                                new_para.paragraph_format.space_after = paragraph.paragraph_format.space_after
                                new_para.paragraph_format.left_indent = paragraph.paragraph_format.left_indent
                                new_para.paragraph_format.right_indent = paragraph.paragraph_format.right_indent
                        except:
                            pass
                        
                        # Copy runs with formatting
                        for run in paragraph.runs:
                            new_run = new_para.add_run(run.text)
                            try:
                                # Copy character formatting
                                new_run.bold = run.bold
                                new_run.italic = run.italic
                                new_run.underline = run.underline
                                
                                # Copy font properties
                                if run.font.name:
                                    new_run.font.name = run.font.name
                                if run.font.size:
                                    new_run.font.size = run.font.size
                                if run.font.color.rgb:
                                    new_run.font.color.rgb = run.font.color.rgb
                                
                                # Copy additional font formatting
                                new_run.font.bold = run.font.bold
                                new_run.font.italic = run.font.italic
                                new_run.font.underline = run.font.underline
                                
                            except Exception as e:
                                print(f"Warning: Could not copy run formatting: {e}")
                    
                    # Copy tables with styles
                    for table in processed_doc.tables:
                        try:
                            rows = len(table.rows)
                            cols = len(table.columns) if rows > 0 else 1
                            new_table = merged_doc.add_table(rows=rows, cols=cols)
                            
                            # Copy table style
                            try:
                                if table.style:
                                    new_table.style = table.style
                            except:
                                pass
                            
                            # Copy table content
                            for i, row in enumerate(table.rows):
                                for j, cell in enumerate(row.cells):
                                    if i < len(new_table.rows) and j < len(new_table.rows[i].cells):
                                        new_cell = new_table.rows[i].cells[j]
                                        new_cell.text = ""
                                        
                                        # Copy cell paragraphs with styles
                                        for para_idx, paragraph in enumerate(cell.paragraphs):
                                            if para_idx == 0:
                                                cell_para = new_cell.paragraphs[0]
                                            else:
                                                cell_para = new_cell.add_paragraph()
                                            
                                            # Copy paragraph style
                                            try:
                                                cell_para.style = paragraph.style
                                            except:
                                                pass
                                            
                                            # Copy runs
                                            for run in paragraph.runs:
                                                cell_run = cell_para.add_run(run.text)
                                                try:
                                                    cell_run.bold = run.bold
                                                    cell_run.italic = run.italic
                                                    cell_run.underline = run.underline
                                                    if run.font.name:
                                                        cell_run.font.name = run.font.name
                                                    if run.font.size:
                                                        cell_run.font.size = run.font.size
                                                    if run.font.color.rgb:
                                                        cell_run.font.color.rgb = run.font.color.rgb
                                                except:
                                                    pass
                        except Exception as e:
                            print(f"Warning: Could not copy table: {e}")
                
                except Exception as e:
                    print(f"Warning: Could not copy content with styles: {e}")
                    # Fallback to original method
                    for element in processed_doc.element.body:
                        merged_doc.element.body.append(element)
            
            merged_doc.save(output_path)
            print(f"Generated single Word document with {len(self.data)} records")
            return True
            
        except Exception as e:
            print(f"Error creating single Word document: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def generate_multiple_word(self, output_dir: str) -> bool:
        """Generate multiple Word documents (one per record)"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            os.makedirs(output_dir, exist_ok=True)
            
            for index, row_data in enumerate(self.data):
                # Load template
                doc = Document(self.template_path)
                
                # Replace merge fields
                processed_doc = self.replace_merge_fields(doc, row_data)
                
                # Generate filename (use first field value or index)
                first_value = list(row_data.values())[0] if row_data else f"record_{index+1}"
                # Clean filename
                safe_filename = re.sub(r'[<>:"/\\|?*]', '_', str(first_value))
                output_path = os.path.join(output_dir, f"{safe_filename}.docx")
                
                processed_doc.save(output_path)
            
            return True
            
        except Exception as e:
            print(f"Error creating multiple Word files: {str(e)}")
            return False
    
    def generate_single_pdf(self, output_path: str) -> bool:
        """Generate a single PDF document by first creating Word doc, then converting"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            print(f"Starting single PDF generation for {len(self.data)} records")
            
            # First, generate a Word document (this works perfectly)
            temp_word_path = output_path.replace('.pdf', '_temp.docx')
            
            success = self.generate_single_word(temp_word_path)
            if not success:
                print("Failed to generate temporary Word document")
                return False
            
            print(f"Temporary Word document created: {temp_word_path}")
            
            # Now convert the Word document to PDF
            success = self._convert_docx_to_pdf(temp_word_path, output_path)
            
            # Clean up temporary file
            try:
                os.remove(temp_word_path)
            except:
                pass
            
            return success
            
        except Exception as e:
            print(f"Error creating single PDF document: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def generate_multiple_pdf(self, output_dir: str) -> bool:
        """Generate multiple PDF documents by first creating Word docs, then converting"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            print(f"Starting multiple PDF generation for {len(self.data)} records")
            
            # First, generate Word documents (this works perfectly)
            temp_word_dir = os.path.join(output_dir, 'temp_word')
            os.makedirs(temp_word_dir, exist_ok=True)
            
            success = self.generate_multiple_word(temp_word_dir)
            if not success:
                print("Failed to generate temporary Word documents")
                return False
            
            print("Temporary Word documents created successfully")
            
            # Now convert each Word document to PDF
            word_files = [f for f in os.listdir(temp_word_dir) if f.endswith('.docx')]
            
            for word_file in word_files:
                word_path = os.path.join(temp_word_dir, word_file)
                pdf_file = word_file.replace('.docx', '.pdf')
                pdf_path = os.path.join(output_dir, pdf_file)
                
                success = self._convert_docx_to_pdf(word_path, pdf_path)
                if not success:
                    print(f"Failed to convert {word_file} to PDF")
                    return False
                
                print(f"Successfully converted: {word_file} -> {pdf_file}")
            
            # Clean up temporary directory
            try:
                import shutil
                shutil.rmtree(temp_word_dir)
            except:
                pass
            
            print(f"Successfully generated {len(word_files)} PDF files")
            return True
            
        except Exception as e:
            print(f"Error creating multiple PDF files: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def _convert_docx_to_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """Convert Word document to PDF using available conversion method"""
        try:
            print(f"Converting Word document to PDF...")
            print(f"Input: {docx_path}")
            print(f"Output: {pdf_path}")
            
            # Try Word's native conversion first (Windows/macOS)
            try:
                from docx2pdf import convert
                convert(docx_path, pdf_path)
                print(f"Successfully converted to PDF using Word: {pdf_path}")
                return True
            except Exception as word_error:
                print(f"Word conversion failed: {word_error}")
                print("Trying LibreOffice conversion...")
            
            # Fallback to LibreOffice (Linux/Windows with LibreOffice installed)
            import subprocess
            import platform
            
            try:
                # Get the directory for output
                output_dir = os.path.dirname(pdf_path)
                
                # Different LibreOffice commands for different OS
                if platform.system() == "Windows":
                    # Try common Windows LibreOffice paths
                    libreoffice_paths = [
                        r"C:\Program Files\LibreOffice\program\soffice.exe",
                        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
                        "soffice.exe",  # If in PATH
                        "libreoffice"   # If in PATH
                    ]
                    
                    libre_cmd = None
                    for path in libreoffice_paths:
                        try:
                            result = subprocess.run([path, "--version"], 
                                                  capture_output=True, text=True, timeout=5)
                            if result.returncode == 0:
                                libre_cmd = path
                                break
                        except:
                            continue
                    
                    if not libre_cmd:
                        raise FileNotFoundError("LibreOffice not found on Windows")
                        
                    cmd = [libre_cmd, "--headless", "--convert-to", "pdf", 
                           "--outdir", output_dir, docx_path]
                else:
                    # Linux/macOS - Enhanced LibreOffice detection for Render
                    cmd = ["libreoffice", "--headless", "--convert-to", "pdf",
                           "--outdir", output_dir, docx_path]
                
                print(f"Running LibreOffice command: {' '.join(cmd)}")
                
                # Add environment variables for headless operation
                import os
                env = os.environ.copy()
                env['DISPLAY'] = ':0'  # Required for headless operation on some systems
                
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, env=env)
                
                if result.returncode == 0:
                    # LibreOffice creates PDF with same name as input file
                    expected_pdf = os.path.join(output_dir, os.path.basename(docx_path).replace('.docx', '.pdf'))
                    
                    if os.path.exists(expected_pdf):
                        # Move to desired location if different
                        if expected_pdf != pdf_path:
                            os.rename(expected_pdf, pdf_path)
                        
                        print(f"Successfully converted to PDF using LibreOffice: {pdf_path}")
                        return True
                    else:
                        print(f"LibreOffice conversion succeeded but PDF not found at: {expected_pdf}")
                        return False
                else:
                    print(f"LibreOffice conversion failed. Return code: {result.returncode}")
                    print(f"stdout: {result.stdout}")
                    print(f"stderr: {result.stderr}")
                    return False
                    
            except subprocess.TimeoutExpired:
                print("LibreOffice conversion timed out")
                return False
            except FileNotFoundError:
                print("LibreOffice not found. Trying basic PDF generation...")
                return self._generate_basic_pdf(docx_path, pdf_path)
            except Exception as libre_error:
                print(f"LibreOffice conversion error: {libre_error}")
                print("Trying basic PDF generation...")
                return self._generate_basic_pdf(docx_path, pdf_path)
            
        except Exception as e:
            print(f"Error converting Word document to PDF: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def _generate_basic_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """Basic PDF generation as final fallback using ReportLab"""
        try:
            print("Using basic PDF generation (ReportLab) as fallback...")
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
            from reportlab.lib.styles import getSampleStyleSheet
            
            # Load the Word document
            doc = Document(docx_path)
            
            # Create PDF document
            pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []
            
            # Convert paragraphs to simple text
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    # Simple text conversion - won't preserve all formatting but works
                    text = paragraph.text.strip()
                    para = Paragraph(text, styles['Normal'])
                    story.append(para)
                    story.append(Spacer(1, 6))
            
            # Build PDF
            pdf_doc.build(story)
            print(f"Basic PDF generated successfully: {pdf_path}")
            return True
            
        except Exception as e:
            print(f"Basic PDF generation failed: {str(e)}")
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
    
    session_id = session['session_id']
    
    if session_id not in processors:
        processors[session_id] = MailMergeProcessor(session_id)
    
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
            return jsonify({
                'success': True,
                'message': f'Template uploaded successfully: {file.filename}',
                'filepath': filepath,
                'filename': file.filename
            })
        else:
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
    """Check current upload status"""
    try:
        processor = get_processor()
        return jsonify({
            'template_loaded': processor.template_path is not None,
            'data_loaded': processor.data_path is not None and len(processor.data) > 0,
            'template_path': processor.template_path,
            'data_records': len(processor.data) if processor.data else 0
        })
    except Exception as e:
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
    return jsonify({'status': 'healthy', 'service': 'Mail Merge SaaS - Word & PDF Support'})

if __name__ == '__main__':
    # Diagnostic information for debugging Render deployment
    print("\n" + "="*50)
    print("Mail Merge SaaS - Startup Diagnostics")
    print("="*50)
    
    # Check PDF conversion capabilities
    check_pdf_conversion_capabilities()
    
    # Additional environment info for Render debugging
    import platform
    import sys
    print(f"Python version: {sys.version}")
    print(f"Platform: {platform.platform()}")
    print(f"Architecture: {platform.architecture()}")
    
    # Check critical imports
    critical_imports = [
        ('flask', 'Flask web framework'),
        ('docx', 'python-docx for Word processing'),
        ('openpyxl', 'Excel file processing'),
        ('reportlab', 'PDF generation fallback')
    ]
    
    print("\nCritical imports check:")
    for module, description in critical_imports:
        try:
            __import__(module)
            print(f"OK {module}: {description}")
        except ImportError as e:
            print(f"FAILED {module}: {e}")
    
    # Additional system checks for Render
    print("\nSystem capabilities check:")
    try:
        import subprocess
        # Check if LibreOffice is available
        try:
            result = subprocess.run(['libreoffice', '--version'], 
                                  capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                print(f"OK LibreOffice: {result.stdout.strip()}")
            else:
                print(f"FAILED LibreOffice: Command failed")
        except FileNotFoundError:
            print("FAILED LibreOffice: Command not found")
        except Exception as e:
            print(f"FAILED LibreOffice: {e}")
            
        # Check available fonts (important for PDF generation)
        try:
            result = subprocess.run(['fc-list'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                font_count = len(result.stdout.split('\n'))
                print(f"OK Fonts: {font_count} fonts available")
            else:
                print("FAILED Fonts: fc-list not available")
        except:
            print("FAILED Fonts: Cannot check font availability")
            
    except Exception as e:
        print(f"System check failed: {e}")
    
    print("="*50)
    
    # Development server
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)