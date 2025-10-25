#!/usr/bin/env python3
"""
Test script to check PDF pagination issue
"""

import os
from app import MailMergeProcessor

def test_pagination():
    """Test the pagination fix"""
    
    # Sample data for testing
    test_data = [
        {'company': 'Company A', 'amount': '$1000', 'date': '2024-01-01'},
        {'company': 'Company B', 'amount': '$2000', 'date': '2024-01-02'}
    ]
    
    # Initialize processor
    processor = MailMergeProcessor()
    
    # Check if template exists
    template_path = 'index.html'  # or whatever your template is
    if not os.path.exists(template_path):
        print("Template file not found. Please make sure you have a template.")
        return False
    
    # Load template and data
    processor.load_template(template_path)
    processor.data = test_data
    
    # Test single PDF generation
    output_pdf = 'test_pagination.pdf'
    
    print("Testing single PDF generation with pagination fix...")
    success = processor.generate_single_pdf(output_pdf)
    
    if success:
        print(f"✅ PDF generated successfully: {output_pdf}")
        print("Please check if the second record appears on page 2 (not page 3)")
        
        # Check if file exists and get size
        if os.path.exists(output_pdf):
            file_size = os.path.getsize(output_pdf)
            print(f"File size: {file_size} bytes")
        
        return True
    else:
        print("❌ PDF generation failed")
        return False

if __name__ == "__main__":
    test_pagination()