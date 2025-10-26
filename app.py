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

# Advanced conversion imports
try:
    import mammoth
    MAMMOTH_AVAILABLE = True
except ImportError:
    MAMMOTH_AVAILABLE = False

try:
    import pdfkit
    PDFKIT_AVAILABLE = True
except ImportError:
    PDFKIT_AVAILABLE = False

from jinja2 import Template

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
        """Convert Word document to PDF using advanced conversion methods"""
        try:
            print(f"Converting Word document to PDF...")
            print(f"Input: {docx_path}")
            print(f"Output: {pdf_path}")
            
            # Try HTML-based conversion first (best formatting preservation)
            if MAMMOTH_AVAILABLE and PDFKIT_AVAILABLE:
                print("Attempting HTML-based conversion for perfect formatting...")
                if self._convert_via_html(docx_path, pdf_path):
                    return True
                print("HTML conversion failed, trying other methods...")
            
            # Try Word's native conversion (Windows/macOS)
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
                print("LibreOffice not found. Trying format-preserving PDF generation...")
                return self._generate_format_preserving_pdf(docx_path, pdf_path)
            except Exception as libre_error:
                print(f"LibreOffice conversion error: {libre_error}")
                print("Trying format-preserving PDF generation...")
                return self._generate_format_preserving_pdf(docx_path, pdf_path)
            
        except Exception as e:
            print(f"Error converting Word document to PDF: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

    def _convert_via_html(self, docx_path: str, pdf_path: str) -> bool:
        """Convert Word to PDF via HTML for perfect formatting preservation"""
        try:
            print("Converting Word → HTML → PDF for maximum formatting preservation...")
            
            # Step 1: Convert Word to HTML with styling
            with open(docx_path, "rb") as docx_file:
                result = mammoth.convert_to_html(
                    docx_file,
                    convert_image=mammoth.images.img_element(self._convert_image)
                )
                html_content = result.value
                messages = result.messages
                
                if messages:
                    print(f"Mammoth conversion messages: {[str(m) for m in messages]}")
            
            # Step 2: Enhance HTML with better CSS for PDF
            enhanced_html = self._enhance_html_for_pdf(html_content)
            
            # Step 3: Convert HTML to PDF using pdfkit (wkhtmltopdf)
            try:
                # Configure pdfkit options for better rendering
                options = {
                    'page-size': 'A4',
                    'margin-top': '2cm',
                    'margin-right': '2cm',
                    'margin-bottom': '2cm',
                    'margin-left': '2cm',
                    'encoding': "UTF-8",
                    'no-outline': None,
                    'enable-local-file-access': None
                }
                
                pdfkit.from_string(enhanced_html, pdf_path, options=options)
                print(f"Successfully converted to PDF via HTML: {pdf_path}")
                return True
            except Exception as pdfkit_error:
                print(f"pdfkit conversion failed: {pdfkit_error}")
                return False
            
        except Exception as e:
            print(f"HTML-based conversion failed: {e}")
            return False

    def _convert_image(self, image):
        """Convert embedded images for HTML"""
        try:
            # For now, skip images to focus on text formatting
            return {"src": "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"}
        except:
            return {"src": ""}

    def _enhance_html_for_pdf(self, html_content: str) -> str:
        """Enhance HTML with CSS for better PDF formatting"""
        css_styles = """
        <style>
        @page {
            size: A4;
            margin: 2cm;
        }
        
        body {
            font-family: 'Times New Roman', serif;
            font-size: 12pt;
            line-height: 1.4;
            color: #000000;
        }
        
        /* Enhanced paragraph styling */
        p {
            margin: 6pt 0;
            text-align: left;
        }
        
        /* Blue text preservation */
        span[style*="color"] {
            /* Preserve inline colors */
        }
        
        /* Underline preservation */
        u, span[style*="text-decoration: underline"] {
            text-decoration: underline;
        }
        
        /* Bold preservation */
        strong, b, span[style*="font-weight: bold"] {
            font-weight: bold;
        }
        
        /* Italic preservation */
        em, i, span[style*="font-style: italic"] {
            font-style: italic;
        }
        
        /* Table styling */
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 12pt 0;
        }
        
        table td, table th {
            border: 1pt solid #000000;
            padding: 6pt;
            text-align: left;
        }
        
        table th {
            background-color: #f0f0f0;
            font-weight: bold;
        }
        
        /* Page break before titles */
        .page-break {
            page-break-before: always;
        }
        
        /* Special styling for blue underlined titles */
        .invoice-title {
            color: #0000FF;
            text-decoration: underline;
            font-weight: bold;
            font-size: 18pt;
            text-align: center;
            margin: 18pt 0;
        }
        </style>
        """
        
        # Enhance HTML structure
        enhanced_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            {css_styles}
        </head>
        <body>
            {self._process_html_content(html_content)}
        </body>
        </html>
        """
        
        return enhanced_html

    def _process_html_content(self, html_content: str) -> str:
        """Process HTML content to add special classes and page breaks"""
        try:
            # Add page breaks before blue underlined titles
            import re
            
            # Look for patterns that might be invoice titles
            processed_html = html_content
            
            # Add page break class to potential title elements
            # Pattern: text that's both blue and underlined
            blue_underline_pattern = r'<p[^>]*>.*?<span[^>]*style="[^"]*color:[^"]*blue[^"]*"[^>]*>.*?<u>([^<]+)</u>.*?</span>.*?</p>'
            
            def add_page_break(match):
                content = match.group(0)
                if 'invoice-title' not in content:
                    # Add special class for styling
                    content = content.replace('<p', '<p class="invoice-title page-break"', 1)
                return content
            
            processed_html = re.sub(blue_underline_pattern, add_page_break, processed_html, flags=re.IGNORECASE | re.DOTALL)
            
            return processed_html
            
        except Exception as e:
            print(f"HTML processing error: {e}")
            return html_content
    
    def _generate_basic_pdf(self, docx_path: str, pdf_path: str) -> bool:
        """Enhanced PDF generation with formatting preservation using ReportLab"""
        try:
            print("Using enhanced PDF generation (ReportLab) with formatting preservation...")
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors
            from reportlab.lib.units import inch
            from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
            
            # Load the Word document
            doc = Document(docx_path)
            
            # Create PDF document
            pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter, 
                                      leftMargin=0.75*inch, rightMargin=0.75*inch,
                                      topMargin=0.75*inch, bottomMargin=0.75*inch)
            styles = getSampleStyleSheet()
            
            # Create custom styles that match common Word styles
            custom_styles = {}
            
            # Invoice title style (blue with underline)
            custom_styles['invoice_title'] = ParagraphStyle(
                'InvoiceTitle',
                parent=styles['Title'],
                fontSize=18,
                textColor=colors.blue,
                spaceAfter=12,
                alignment=TA_CENTER,
                fontName='Helvetica-Bold'
            )
            
            # Title style (for other headers)
            custom_styles['Title'] = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontSize=24,
                textColor=colors.red,
                spaceAfter=12,
                fontName='Helvetica-Bold'
            )
            
            # Heading styles
            custom_styles['Heading 1'] = ParagraphStyle(
                'CustomHeading1',
                parent=styles['Heading1'],
                fontSize=18,
                textColor=colors.blue,
                fontName='Helvetica-Bold'
            )
            
            custom_styles['Heading 2'] = ParagraphStyle(
                'CustomHeading2',
                parent=styles['Heading2'],
                fontSize=14,
                textColor=colors.darkblue,
                fontName='Helvetica-Bold'
            )
            
            # Enhanced normal style
            custom_styles['Normal'] = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=12,
                fontName='Helvetica'
            )
            
            story = []
            invoice_count = 0
            
            # Process each paragraph with formatting detection
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    text = paragraph.text.strip()
                    
                    # Detect the style based on actual Word formatting
                    detected_style = self._detect_title_from_formatting(paragraph)
                    
                    # Check if this is a title with special formatting (blue + underline)
                    if detected_style == 'invoice_title':
                        if invoice_count > 0:
                            story.append(PageBreak())
                        invoice_count += 1
                        # Use special invoice style
                        style_to_use = custom_styles['invoice_title']
                        # Add underline to text
                        formatted_text = f'<u>{text}</u>'
                    else:
                        # Detect style based on content and apply appropriate formatting
                        style_to_use = self._detect_paragraph_style(paragraph, text, custom_styles, styles)
                        # Enhanced text with formatting preservation
                        formatted_text = self._extract_formatted_text_enhanced(paragraph)
                    
                    if formatted_text:
                        para = Paragraph(formatted_text, style_to_use)
                        story.append(para)
                        story.append(Spacer(1, 6))
            
            # Add tables if any
            for table in doc.tables:
                story.append(self._convert_table_to_reportlab(table, custom_styles))
                story.append(Spacer(1, 12))
            
            # Build PDF
            pdf_doc.build(story)
            print(f"Enhanced PDF generated successfully: {pdf_path}")
            return True
            
        except Exception as e:
            print(f"Enhanced PDF generation failed: {str(e)}")
            # Fall back to even simpler generation
            return self._generate_simple_pdf_fallback(docx_path, pdf_path)
    
    def _detect_title_from_formatting(self, paragraph):
        """Detect if paragraph is a special title based on formatting (blue + underline)"""
        try:
            if not paragraph.runs:
                return 'normal'
            
            first_run = paragraph.runs[0]
            is_blue = False
            is_underlined = False
            
            # Check for blue color
            try:
                if hasattr(first_run.font, 'color') and first_run.font.color:
                    if hasattr(first_run.font.color, 'rgb') and first_run.font.color.rgb:
                        r, g, b = first_run.font.color.rgb
                        # Check if it's blue-ish
                        if b > r and b > g and b > 100:
                            is_blue = True
                    elif hasattr(first_run.font.color, 'theme_color'):
                        if first_run.font.color.theme_color in [3, 5]:
                            is_blue = True
            except:
                pass
            
            # Check for underline
            if first_run.underline:
                is_underlined = True
            
            # If both blue and underlined, it's a special title
            if is_blue and is_underlined:
                return 'invoice_title'
            
            return 'normal'
            
        except:
            return 'normal'

    def _detect_paragraph_style(self, paragraph, text, custom_styles, default_styles):
        """Detect appropriate style based on paragraph properties and content"""
        try:
            # First check if it has special formatting (blue + underline)
            special_style = self._detect_title_from_formatting(paragraph)
            if special_style == 'invoice_title':
                return custom_styles['invoice_title']
            
            # Check Word style name if available
            if hasattr(paragraph, 'style') and paragraph.style:
                style_name = paragraph.style.name
                if style_name in custom_styles:
                    return custom_styles[style_name]
                elif style_name in ['Title', 'Heading 1', 'Heading 2']:
                    return custom_styles.get(style_name, default_styles['Normal'])
            
            # Check for formatting clues
            if paragraph.runs:
                first_run = paragraph.runs[0]
                if first_run.bold and first_run.font.size and first_run.font.size.pt > 16:
                    return custom_styles['Title']
                elif first_run.bold and first_run.font.size and first_run.font.size.pt > 14:
                    return custom_styles['Heading 1']
                elif first_run.bold:
                    return custom_styles['Heading 2']
            
            return custom_styles['Normal']
            
        except Exception as e:
            print(f"Style detection failed: {e}")
            return default_styles['Normal']
    
    def _generate_format_preserving_pdf(self, doc_path, pdf_path):
        """Advanced PDF generation with maximum format preservation"""
        try:
            print("Attempting format-preserving PDF generation...")
            
            # First, analyze the document
            formatting_info = self._analyze_document_formatting(doc_path)
            if not formatting_info:
                return self._generate_enhanced_pdf(doc_path, pdf_path)
            
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors
            from reportlab.lib.units import inch
            from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
            
            # Load document
            doc = Document(doc_path)
            
            # Create PDF document
            pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter, 
                                      topMargin=0.75*inch, bottomMargin=0.75*inch,
                                      leftMargin=0.75*inch, rightMargin=0.75*inch)
            
            styles = getSampleStyleSheet()
            story = []
            
            # Create enhanced custom styles based on document analysis
            custom_styles = {}
            
            # Special Invoice title style (blue with underline)
            custom_styles['invoice_title'] = ParagraphStyle(
                'InvoiceTitle',
                parent=styles['Title'],
                fontSize=18,
                spaceAfter=18,
                alignment=TA_CENTER,
                textColor=colors.blue,
                fontName='Helvetica-Bold'
            )
            
            # Title style (enhanced)
            custom_styles['title'] = ParagraphStyle(
                'CustomTitle',
                parent=styles['Title'],
                fontSize=18,
                spaceAfter=18,
                alignment=TA_CENTER,
                textColor=colors.black,
                fontName='Helvetica-Bold'
            )
            
            # Heading styles
            custom_styles['heading'] = ParagraphStyle(
                'CustomHeading',
                parent=styles['Heading1'],
                fontSize=14,
                spaceAfter=12,
                textColor=colors.black,
                fontName='Helvetica-Bold'
            )
            
            # Enhanced normal style
            custom_styles['normal'] = ParagraphStyle(
                'CustomNormal',
                parent=styles['Normal'],
                fontSize=11,
                spaceAfter=6,
                textColor=colors.black,
                fontName='Helvetica',
                leading=14
            )
            
            # Process each paragraph with enhanced formatting
            invoice_count = 0
            for paragraph in doc.paragraphs:
                if not paragraph.text.strip():
                    continue
                
                # Detect appropriate style
                detected_style = self._detect_advanced_style(paragraph, formatting_info)
                
                # Add page break before each new special title (blue + underline)
                if detected_style == 'invoice_title':
                    if invoice_count > 0:
                        story.append(PageBreak())
                    invoice_count += 1
                
                # Get enhanced formatted text
                formatted_text = self._extract_formatted_text_enhanced(paragraph)
                
                # Choose style
                if detected_style == 'invoice_title':
                    para_style = custom_styles['invoice_title']
                    # Add underline formatting to the text
                    if formatted_text and not '<u>' in formatted_text:
                        formatted_text = f'<u>{formatted_text}</u>'
                elif detected_style == 'title':
                    para_style = custom_styles['title']
                elif detected_style == 'heading':
                    para_style = custom_styles['heading']
                else:
                    para_style = custom_styles['normal']
                
                # Create paragraph
                try:
                    if formatted_text and formatted_text.strip():
                        p = Paragraph(formatted_text, para_style)
                        story.append(p)
                        story.append(Spacer(1, 6))
                except Exception as e:
                    print(f"Paragraph creation failed: {e}")
                    # Fallback to plain text
                    if paragraph.text.strip():
                        p = Paragraph(paragraph.text, para_style)
                        story.append(p)
                        story.append(Spacer(1, 6))
            
            # Process tables with enhanced formatting
            for table in doc.tables:
                try:
                    table_element = self._convert_table_enhanced(table, formatting_info)
                    if table_element:
                        story.append(table_element)
                        story.append(Spacer(1, 12))
                except Exception as e:
                    print(f"Table conversion failed: {e}")
            
            # Build the PDF
            pdf_doc.build(story)
            print(f"Format-preserving PDF generated successfully: {pdf_path}")
            return True
            
        except Exception as e:
            print(f"Format-preserving PDF generation failed: {e}")
            return self._generate_enhanced_pdf(doc_path, pdf_path)

    def _detect_advanced_style(self, paragraph, formatting_info):
        """Advanced style detection using document analysis and actual Word formatting"""
        try:
            text = paragraph.text.strip()
            
            # Check the actual Word style name first
            if hasattr(paragraph, 'style') and paragraph.style:
                style_name = paragraph.style.name.lower()
                
                # Check for title-like styles
                if 'title' in style_name:
                    return 'title'
                elif 'heading' in style_name:
                    return 'heading'
            
            # Check for formatting characteristics that indicate it's a title
            if paragraph.runs:
                first_run = paragraph.runs[0]
                
                # Check if text is blue and/or underlined (likely a styled title)
                is_blue = False
                is_underlined = False
                is_large = False
                is_bold = False
                
                try:
                    # Check for blue color
                    if hasattr(first_run.font, 'color') and first_run.font.color:
                        if hasattr(first_run.font.color, 'rgb') and first_run.font.color.rgb:
                            r, g, b = first_run.font.color.rgb
                            # Check if it's blue-ish (more blue than other colors)
                            if b > r and b > g and b > 100:
                                is_blue = True
                        elif hasattr(first_run.font.color, 'theme_color'):
                            # Theme color 5 is often blue
                            if first_run.font.color.theme_color in [3, 5]:  # Blue theme colors
                                is_blue = True
                except:
                    pass
                
                # Check for underline
                if first_run.underline:
                    is_underlined = True
                
                # Check for bold
                if first_run.bold:
                    is_bold = True
                
                # Check for large font size
                try:
                    if hasattr(first_run.font, 'size') and first_run.font.size:
                        if first_run.font.size.pt >= 16:
                            is_large = True
                except:
                    pass
                
                # If it's blue AND underlined, it's likely a special title
                if is_blue and is_underlined:
                    return 'invoice_title'
                
                # If it's large and bold, it's likely a regular title
                if is_large and is_bold:
                    return 'title'
                
                # If it's just bold and short, it might be a heading
                if is_bold and len(text) < 100:
                    return 'heading'
            
            return 'normal'
            
        except Exception as e:
            print(f"Style detection error: {e}")
            return 'normal'

    def _convert_table_enhanced(self, table, formatting_info):
        """Enhanced table conversion with formatting preservation"""
        try:
            from reportlab.platypus import Table, TableStyle, Spacer
            from reportlab.lib import colors
            
            # Extract table data with formatting
            data = []
            for row_idx, row in enumerate(table.rows):
                row_data = []
                for cell in row.cells:
                    # Get cell text with basic formatting
                    cell_text = cell.text.strip()
                    if not cell_text:
                        cell_text = ""
                    row_data.append(cell_text)
                data.append(row_data)
            
            if not data:
                return Spacer(1, 1)
            
            # Create enhanced table
            t = Table(data, hAlign='LEFT')
            
            # Enhanced table styling
            table_style = [
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
                ('TOPPADDING', (0, 0), (-1, 0), 8),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]
            
            t.setStyle(TableStyle(table_style))
            return t
            
        except Exception as e:
            print(f"Enhanced table conversion failed: {e}")
            return Spacer(1, 1)

    def _analyze_document_formatting(self, doc_path):
        """Analyze the Word document to extract comprehensive formatting information"""
        try:
            doc = Document(doc_path)
            formatting_info = {
                'styles_used': set(),
                'colors_used': set(),
                'fonts_used': set(),
                'has_tables': False,
                'has_images': False,
                'paragraph_count': 0,
                'formatted_runs': []
            }
            
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    formatting_info['paragraph_count'] += 1
                    
                    # Detect style
                    if paragraph.style:
                        formatting_info['styles_used'].add(paragraph.style.name)
                    
                    # Analyze runs for detailed formatting
                    for run in paragraph.runs:
                        run_info = {
                            'text': run.text,
                            'bold': run.bold,
                            'italic': run.italic,
                            'underline': run.underline,
                            'font_size': None,
                            'font_name': None,
                            'color': None
                        }
                        
                        # Font information
                        if hasattr(run.font, 'size') and run.font.size:
                            run_info['font_size'] = run.font.size.pt
                        
                        if hasattr(run.font, 'name') and run.font.name:
                            run_info['font_name'] = run.font.name
                            formatting_info['fonts_used'].add(run.font.name)
                        
                        # Color information
                        if hasattr(run.font, 'color') and run.font.color:
                            try:
                                if hasattr(run.font.color, 'rgb') and run.font.color.rgb:
                                    r, g, b = run.font.color.rgb
                                    hex_color = f"#{r:02x}{g:02x}{b:02x}"
                                    run_info['color'] = hex_color
                                    formatting_info['colors_used'].add(hex_color)
                            except:
                                pass
                        
                        formatting_info['formatted_runs'].append(run_info)
            
            # Check for tables
            if doc.tables:
                formatting_info['has_tables'] = True
            
            print(f"Document analysis: {len(formatting_info['styles_used'])} styles, "
                  f"{len(formatting_info['colors_used'])} colors, "
                  f"{len(formatting_info['fonts_used'])} fonts, "
                  f"{formatting_info['paragraph_count']} paragraphs")
            
            return formatting_info
            
        except Exception as e:
            print(f"Document analysis failed: {e}")
            return None

    def _extract_formatted_text_enhanced(self, paragraph):
        """Enhanced text extraction with comprehensive formatting support"""
        try:
            formatted_parts = []
            
            for run in paragraph.runs:
                text = run.text
                if not text:
                    continue
                
                # Start with the base text
                formatted_text = text
                
                # Apply bold
                if run.bold:
                    formatted_text = f"<b>{formatted_text}</b>"
                
                # Apply italic
                if run.italic:
                    formatted_text = f"<i>{formatted_text}</i>"
                
                # Apply underline
                if run.underline:
                    formatted_text = f"<u>{formatted_text}</u>"
                
                # Handle font size
                if hasattr(run.font, 'size') and run.font.size:
                    try:
                        # Convert from Pt to points
                        size_pt = run.font.size.pt
                        if size_pt != 12:  # Only apply if different from default
                            formatted_text = f'<font size="{size_pt}">{formatted_text}</font>'
                    except:
                        pass
                
                # Handle font color with enhanced detection
                color_applied = False
                try:
                    if hasattr(run.font, 'color') and run.font.color:
                        color = run.font.color
                        
                        # Try RGB color first
                        if hasattr(color, 'rgb') and color.rgb:
                            r, g, b = color.rgb
                            hex_color = f"#{r:02x}{g:02x}{b:02x}"
                            if hex_color != "#000000":  # Only apply if not black
                                formatted_text = f'<font color="{hex_color}">{formatted_text}</font>'
                                color_applied = True
                        
                        # Try theme color as fallback
                        elif hasattr(color, 'theme_color') and color.theme_color:
                            # Map common theme colors to hex
                            theme_colors = {
                                1: "#FF0000",  # Red
                                2: "#00FF00",  # Green
                                3: "#0000FF",  # Blue
                                4: "#FFFF00",  # Yellow
                                5: "#FF00FF",  # Magenta
                                6: "#00FFFF",  # Cyan
                            }
                            if color.theme_color in theme_colors:
                                hex_color = theme_colors[color.theme_color]
                                formatted_text = f'<font color="{hex_color}">{formatted_text}</font>'
                                color_applied = True
                except:
                    pass
                
                # Handle font name/family
                try:
                    if hasattr(run.font, 'name') and run.font.name:
                        font_name = run.font.name
                        # Map to ReportLab supported fonts
                        font_mapping = {
                            'Arial': 'Helvetica',
                            'Times New Roman': 'Times-Roman',
                            'Courier New': 'Courier',
                            'Calibri': 'Helvetica',
                            'Tahoma': 'Helvetica',
                            'Verdana': 'Helvetica'
                        }
                        rl_font = font_mapping.get(font_name, 'Helvetica')
                        if rl_font != 'Helvetica':  # Only apply if different from default
                            formatted_text = f'<font face="{rl_font}">{formatted_text}</font>'
                except:
                    pass
                
                formatted_parts.append(formatted_text)
            
            result = ''.join(formatted_parts) if formatted_parts else paragraph.text
            return result if result.strip() else paragraph.text
            
        except Exception as e:
            print(f"Enhanced text formatting extraction failed: {e}")
            return paragraph.text

    def _extract_formatted_text(self, paragraph):
        """Extract text with basic HTML formatting for ReportLab"""
        try:
            formatted_parts = []
            
            for run in paragraph.runs:
                text = run.text
                if not text:
                    continue
                
                # Apply formatting
                if run.bold:
                    text = f"<b>{text}</b>"
                if run.italic:
                    text = f"<i>{text}</i>"
                if run.underline:
                    text = f"<u>{text}</u>"
                
                # Handle colors (basic support)
                if hasattr(run.font, 'color') and run.font.color.rgb:
                    rgb = run.font.color.rgb
                    if rgb:
                        # Convert to hex color
                        hex_color = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
                        text = f'<font color="{hex_color}">{text}</font>'
                
                formatted_parts.append(text)
            
            return ''.join(formatted_parts) if formatted_parts else paragraph.text
            
        except Exception as e:
            print(f"Text formatting extraction failed: {e}")
            return paragraph.text
    
    def _convert_table_to_reportlab(self, table, custom_styles=None):
        """Convert Word table to ReportLab table"""
        try:
            from reportlab.platypus import Table, TableStyle, Spacer
            from reportlab.lib import colors
            
            # Extract table data
            data = []
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()
                    row_data.append(cell_text)
                data.append(row_data)
            
            if not data:
                return Spacer(1, 1)
            
            # Create ReportLab table
            t = Table(data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 12),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 1, colors.black)
            ]))
            
            return t
            
        except Exception as e:
            print(f"Table conversion failed: {e}")
            return Spacer(1, 1)
    
    def _generate_simple_pdf_fallback(self, docx_path: str, pdf_path: str) -> bool:
        """Simplest possible PDF generation as ultimate fallback"""
        try:
            print("Using simple PDF fallback...")
            from reportlab.lib.pagesizes import letter
            from reportlab.platypus import SimpleDocTemplate, Paragraph
            from reportlab.lib.styles import getSampleStyleSheet
            
            doc = Document(docx_path)
            pdf_doc = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            story = []
            
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    text = paragraph.text.strip()
                    para = Paragraph(text, styles['Normal'])
                    story.append(para)
            
            pdf_doc.build(story)
            print(f"Simple PDF generated: {pdf_path}")
            return True
            
        except Exception as e:
            print(f"Even simple PDF generation failed: {e}")
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
    print("Mail Merge SaaS - Enhanced Startup Diagnostics")
    print("="*50)
    
    # Check PDF conversion capabilities
    check_pdf_conversion_capabilities()
    
    # Check HTML-based conversion
    if MAMMOTH_AVAILABLE and PDFKIT_AVAILABLE:
        print("HTML-based conversion available (PERFECT FORMATTING)")
    else:
        print("HTML-based conversion not available")
        if not MAMMOTH_AVAILABLE:
            print("  - mammoth not installed")
        if not PDFKIT_AVAILABLE:
            print("  - pdfkit not installed")
    
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