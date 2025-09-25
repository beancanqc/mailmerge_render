"""
Mail Merge SaaS - Main Application
Flask web server for Render deployment
"""

import os
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from flask import Flask, request, jsonify, send_file, render_template_string, send_from_directory
from werkzeug.utils import secure_filename
import pandas as pd
from docx import Document
from docx2pdf import convert
import re
from typing import List, Dict, Any, Optional

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size

# Configure folders
UPLOAD_FOLDER = tempfile.mkdtemp()
OUTPUT_FOLDER = tempfile.mkdtemp()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

ALLOWED_TEMPLATE_EXTENSIONS = {'docx', 'doc'}
ALLOWED_DATA_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename, allowed_extensions):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions

class MailMergeProcessor:
    def __init__(self):
        self.template_path: Optional[str] = None
        self.data_path: Optional[str] = None
        self.data_df: Optional[pd.DataFrame] = None
        
    def load_template(self, template_path: str) -> bool:
        """Load and validate Word template file"""
        try:
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"Template file not found: {template_path}")
            
            if not template_path.lower().endswith(('.docx', '.doc')):
                raise ValueError("Template must be a Word document (.docx or .doc)")
            
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
            
            if not data_path.lower().endswith(('.xlsx', '.xls')):
                raise ValueError("Data file must be an Excel file (.xlsx or .xls)")
            
            # Load Excel data
            self.data_df = pd.read_excel(data_path)
            
            if self.data_df.empty:
                raise ValueError("Excel file is empty")
            
            self.data_path = data_path
            return True
            
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            return False
    
    def find_merge_fields(self, doc: Document) -> List[str]:
        """Find all merge fields in the document (format: {{field_name}})"""
        merge_fields = set()
        
        # Search in paragraphs
        for paragraph in doc.paragraphs:
            fields = re.findall(r'\{\{(\w+)\}\}', paragraph.text)
            merge_fields.update(fields)
        
        # Search in tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    fields = re.findall(r'\{\{(\w+)\}\}', cell.text)
                    merge_fields.update(fields)
        
        return list(merge_fields)
    
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
        """Generate a single Word document with all merged records"""
        try:
            if not self.template_path or self.data_df is None:
                raise ValueError("Template and data must be loaded first")
            
            # Create output document
            merged_doc = Document()
            
            for index, row in self.data_df.iterrows():
                # Load fresh template for each record
                template_doc = Document(self.template_path)
                
                # Replace merge fields
                processed_doc = self.replace_merge_fields(template_doc, row.to_dict())
                
                # Add content to merged document
                for element in processed_doc.element.body:
                    merged_doc.element.body.append(element)
                
                # Add page break between records (except for the last one)
                if index < len(self.data_df) - 1:
                    merged_doc.add_page_break()
            
            merged_doc.save(output_path)
            return True
            
        except Exception as e:
            print(f"Error creating single Word file: {str(e)}")
            return False
    
    def generate_multiple_word(self, output_dir: str) -> bool:
        """Generate multiple Word documents (one per record)"""
        try:
            if not self.template_path or self.data_df is None:
                raise ValueError("Template and data must be loaded first")
            
            os.makedirs(output_dir, exist_ok=True)
            
            for index, row in self.data_df.iterrows():
                # Load template
                doc = Document(self.template_path)
                
                # Replace merge fields
                processed_doc = self.replace_merge_fields(doc, row.to_dict())
                
                # Generate filename (use first column value or index)
                first_col_value = str(row.iloc[0]) if len(row) > 0 else f"record_{index+1}"
                # Clean filename
                safe_filename = re.sub(r'[<>:"/\\|?*]', '_', first_col_value)
                output_path = os.path.join(output_dir, f"{safe_filename}.docx")
                
                processed_doc.save(output_path)
            
            return True
            
        except Exception as e:
            print(f"Error creating multiple Word files: {str(e)}")
            return False
    
    def generate_single_pdf(self, output_path: str) -> bool:
        """Generate a single PDF with all merged records"""
        try:
            # First create a temporary Word document
            temp_dir = tempfile.mkdtemp()
            temp_word_path = os.path.join(temp_dir, "merged_temp.docx")
            
            if not self.generate_single_word(temp_word_path):
                return False
            
            # Convert to PDF
            convert(temp_word_path, output_path)
            
            # Clean up
            os.remove(temp_word_path)
            os.rmdir(temp_dir)
            
            return True
            
        except Exception as e:
            print(f"Error creating single PDF: {str(e)}")
            return False
    
    def generate_multiple_pdf(self, output_dir: str) -> bool:
        """Generate multiple PDF files (one per record)"""
        try:
            # First create temporary Word documents
            temp_dir = tempfile.mkdtemp()
            word_dir = os.path.join(temp_dir, "word_files")
            
            if not self.generate_multiple_word(word_dir):
                return False
            
            # Convert each Word file to PDF
            os.makedirs(output_dir, exist_ok=True)
            
            for word_file in os.listdir(word_dir):
                if word_file.endswith('.docx'):
                    word_path = os.path.join(word_dir, word_file)
                    pdf_name = word_file.replace('.docx', '.pdf')
                    pdf_path = os.path.join(output_dir, pdf_name)
                    
                    convert(word_path, pdf_path)
            
            # Clean up temporary files
            import shutil
            shutil.rmtree(temp_dir)
            
            return True
            
        except Exception as e:
            print(f"Error creating multiple PDFs: {str(e)}")
            return False
    
    def process_mail_merge(self, output_format: str, output_path: str) -> bool:
        """Main processing function"""
        try:
            if not self.template_path or self.data_df is None:
                raise ValueError("Both template and data files must be loaded")
            
            # Process based on output format
            if output_format == "single-pdf":
                return self.generate_single_pdf(output_path)
            elif output_format == "multiple-pdf":
                return self.generate_multiple_pdf(output_path)
            elif output_format == "single-word":
                return self.generate_single_word(output_path)
            elif output_format == "multiple-word":
                return self.generate_multiple_word(output_path)
            else:
                raise ValueError(f"Unknown output format: {output_format}")
                
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
        return "Application starting...", 200

@app.route('/mailmerge.html')
def mailmerge():
    """Serve the mail merge page"""
    try:
        with open('mailmerge.html', 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        return "Mail merge page not found", 404

@app.route('/style.css')
def serve_css():
    """Serve CSS file"""
    try:
        with open('style.css', 'r', encoding='utf-8') as f:
            css_content = f.read()
        response = app.response_class(css_content, mimetype='text/css')
        return response
    except FileNotFoundError:
        return "CSS not found", 404

@app.route('/mailmerge.js')
def serve_js():
    """Serve JavaScript file"""
    try:
        with open('mailmerge.js', 'r', encoding='utf-8') as f:
            js_content = f.read()
        response = app.response_class(js_content, mimetype='application/javascript')
        return response
    except FileNotFoundError:
        return "JS not found", 404

@app.route('/upload_template', methods=['POST'])
def upload_template():
    """Handle template file upload"""
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename, ALLOWED_TEMPLATE_EXTENSIONS):
            return jsonify({'error': 'Invalid file type. Please upload a .docx or .doc file'}), 400
        
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"template_{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        file.save(filepath)
        
        # Validate the template
        processor = MailMergeProcessor()
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
            return jsonify({'error': 'No file provided'}), 400
        
        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        if not allowed_file(file.filename, ALLOWED_DATA_EXTENSIONS):
            return jsonify({'error': 'Invalid file type. Please upload a .xlsx or .xls file'}), 400
        
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"data_{timestamp}_{filename}"
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        
        file.save(filepath)
        
        # Validate the data
        processor = MailMergeProcessor()
        if processor.load_data(filepath):
            # Return preview of data
            preview_data = processor.data_df.head(3).to_dict('records')
            columns = list(processor.data_df.columns)
            total_rows = len(processor.data_df)
            
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
    """Process the mail merge"""
    try:
        data = request.get_json()
        
        template_path = data.get('template_path')
        data_path = data.get('data_path')
        output_format = data.get('output_format')
        
        if not all([template_path, data_path, output_format]):
            return jsonify({'error': 'Missing required parameters'}), 400
        
        # Create processor
        processor = MailMergeProcessor()
        
        # Load files
        if not processor.load_template(template_path):
            return jsonify({'error': 'Failed to load template'}), 400
        
        if not processor.load_data(data_path):
            return jsonify({'error': 'Failed to load data'}), 400
        
        # Generate output path
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        if output_format in ['single-pdf', 'single-word']:
            extension = '.pdf' if output_format == 'single-pdf' else '.docx'
            output_filename = f"mailmerge_{timestamp}{extension}"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        else:
            output_filename = f"mailmerge_{timestamp}"
            output_path = os.path.join(OUTPUT_FOLDER, output_filename)
            os.makedirs(output_path, exist_ok=True)
        
        # Process mail merge
        success = processor.process_mail_merge(output_format, output_path)
        
        if success:
            if output_format in ['single-pdf', 'single-word']:
                # Single file - return download link
                return jsonify({
                    'success': True,
                    'message': 'Mail merge completed successfully!',
                    'download_url': f'/download/{output_filename}',
                    'filename': output_filename
                })
            else:
                # Multiple files - create zip
                zip_filename = f"{output_filename}.zip"
                zip_path = os.path.join(OUTPUT_FOLDER, zip_filename)
                
                with zipfile.ZipFile(zip_path, 'w') as zipf:
                    for root, dirs, files in os.walk(output_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            zipf.write(file_path, file)
                
                return jsonify({
                    'success': True,
                    'message': f'Mail merge completed! {len(os.listdir(output_path))} files created.',
                    'download_url': f'/download/{zip_filename}',
                    'filename': zip_filename
                })
        else:
            return jsonify({'error': 'Mail merge processing failed'}), 500
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """Download processed files"""
    try:
        file_path = os.path.join(OUTPUT_FOLDER, filename)
        if os.path.exists(file_path):
            return send_file(file_path, as_attachment=True, download_name=filename)
        else:
            return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health')
def health():
    """Health check endpoint for Render"""
    return jsonify({'status': 'healthy', 'service': 'mail-merge-saas'}), 200

if __name__ == '__main__':
    # Render deployment configuration
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)