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

from flask import Flask, request, jsonify, send_file
from werkzeug.utils import secure_filename
from docx import Document
import openpyxl
import re
from typing import List, Dict, Any, Optional

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Configure folders
UPLOAD_FOLDER = tempfile.mkdtemp()
OUTPUT_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_TEMPLATE_EXTENSIONS = {'docx'}
ALLOWED_DATA_EXTENSIONS = {'xlsx'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

class MailMergeProcessor:
    def __init__(self):
        self.template_path: Optional[str] = None
        self.data_path: Optional[str] = None
        self.data: List[Dict[str, Any]] = []
        
    def load_template(self, template_path: str) -> bool:
        """Load and validate Word template file"""
        try:
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {template_path}")
            
            if not template_path.lower().endswith('.docx'):
                raise ValueError("Template must be a Word document (.docx)")
            
            # Test if file can be opened
            doc = Document(template_path)
            self.template_path = template_path
            return True
            
        except Exception as e:
            print(f"Error loading template: {str(e)}")
            return False
    
    def load_data(self, data_path: str) -> bool:
        """Load and validate Excel data file"""
        try:
            if not os.path.exists(data_path):
                raise FileNotFoundError(f"Data file not found: {data_path}")
            
            if not data_path.lower().endswith('.xlsx'):
                raise ValueError("Data file must be an Excel file (.xlsx)")
            
            # Load Excel data using openpyxl
            workbook = openpyxl.load_workbook(data_path, data_only=True)
            sheet = workbook.active
            
            if sheet.max_row <= 1:
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
            
            if not self.data:
                raise ValueError("No data rows found in Excel file")
            
            self.data_path = data_path
            workbook.close()
            return True
            
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            return False
    
    def replace_merge_fields(self, doc: Document, data_row: Dict[str, Any]) -> Document:
        """Replace merge fields with actual data"""
        # Replace in paragraphs
        for paragraph in doc.paragraphs:
            for field_name, value in data_row.items():
                field_pattern = f"{{{{{field_name}}}}}"
                if field_pattern in paragraph.text:
                    paragraph.text = paragraph.text.replace(field_pattern, str(value))
        
        # Replace in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for field_name, value in data_row.items():
                        field_pattern = f"{{{{{field_name}}}}}"
                        if field_pattern in cell.text:
                            cell.text = cell.text.replace(field_pattern, str(value))
        
        return doc
    
    def generate_single_word(self, output_path: str) -> bool:
        """Generate a single Word document with all records"""
        try:
            if not self.template_path or not self.data:
                raise ValueError("Template and data must be loaded first")
            
            # Create first document from template
            merged_doc = Document(self.template_path)
            
            # Replace fields for first record
            if self.data:
                merged_doc = self.replace_merge_fields(merged_doc, self.data[0])
            
            # Add remaining records
            for row_data in self.data[1:]:
                # Add page break
                merged_doc.add_page_break()
                
                # Load template again for each record
                template_doc = Document(self.template_path)
                processed_doc = self.replace_merge_fields(template_doc, row_data)
                
                # Append content
                for element in processed_doc.element.body:
                    merged_doc.element.body.append(element)
            
            merged_doc.save(output_path)
            return True
            
        except Exception as e:
            print(f"Error creating single Word document: {str(e)}")
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
            else:
                raise ValueError(f"Unsupported output format: {output_format}")
                
        except Exception as e:
            print(f"Error processing mail merge: {str(e)}")
            return False

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

# Global processor instance
processor = MailMergeProcessor()

@app.route('/upload_template', methods=['POST'])
def upload_template():
    """Handle template file upload"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
            return jsonify({'error': 'Invalid file type. Please upload a .docx file'}), 400
        
        # Save file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Load template
        if processor.load_template(filepath):
            return jsonify({
                'success': True,
                'message': f'Template uploaded successfully: {file.filename}',
                'filepath': filepath,
                'filename': file.filename
            })
        else:
            os.remove(filepath)
            return jsonify({'error': 'Invalid template file'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/upload_data', methods=['POST'])
def upload_data():
    """Handle data file upload"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename, ALLOWED_DATA_EXTENSIONS):
            return jsonify({'error': 'Invalid file type. Please upload an Excel file (.xlsx)'}), 400
        
        # Save file
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
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
            os.remove(filepath)
            return jsonify({'error': 'Invalid data file'}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/process_merge', methods=['POST'])
def process_merge():
    """Process the mail merge - Word documents only"""
    try:
        data = request.get_json()
        output_format = data.get('format', 'single-word')
        
        # Only allow Word formats
        if 'pdf' in output_format:
            return jsonify({'error': 'PDF conversion not available on free hosting. Please use Word format - you can convert to PDF locally.'}), 400
        
        if not processor.template_path or not processor.data:
            return jsonify({'error': 'Please upload both template and data files first'}), 400
        
        # Generate unique filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if output_format == 'single-word':
            # Single file output
            output_filename = f"mailmerge_result_{timestamp}.docx"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            
            if processor.process_merge(output_format, output_path):
                return jsonify({
                    'success': True,
                    'message': 'Mail merge completed successfully! Download the Word file and convert to PDF if needed.',
                    'download_url': f'/download/{output_filename}',
                    'filename': output_filename
                })
            else:
                return jsonify({'error': 'Failed to process mail merge'}), 500
                
        else:  # multiple-word
            # Multiple files - create ZIP
            output_dir = os.path.join(OUTPUT_FOLDER, f"mailmerge_{timestamp}")
            zip_filename = f"mailmerge_results_{timestamp}.zip"
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
                    'message': f'Mail merge completed! Generated {len(processor.data)} Word documents. Convert to PDF locally if needed.',
                    'download_url': f'/download/{zip_filename}',
                    'filename': zip_filename
                })
            else:
                return jsonify({'error': 'Failed to process mail merge'}), 500
                
    except Exception as e:
        return jsonify({'error': str(e)}), 500

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
    return jsonify({'status': 'healthy', 'service': 'Mail Merge SaaS - Word Only'})

if __name__ == '__main__':
    # Development server
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)