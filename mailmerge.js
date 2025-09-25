/**
 * Mail Merge Frontend JavaScript
 * Handles file uploads, drag & drop, and processing
 */

class MailMergeApp {
    constructor() {
        this.templateFile = null;
        this.dataFile = null;
        this.templatePath = null;
        this.dataPath = null;
        this.selectedFormat = null;

        this.initializeEventListeners();
    }

    initializeEventListeners() {
        // Drop zone event listeners
        this.setupDropZone('template-drop', 'template-input', this.handleTemplateFile.bind(this));
        this.setupDropZone('data-drop', 'data-input', this.handleDataFile.bind(this));

        // Format selection
        document.querySelectorAll('input[name="output-format"]').forEach(radio => {
            radio.addEventListener('change', (e) => {
                this.selectedFormat = e.target.value;
                this.updateMergeButton();
            });
        });

        // Merge button
        const mergeBtn = document.querySelector('.merge-btn');
        if (mergeBtn) {
            mergeBtn.addEventListener('click', this.processMerge.bind(this));
        }
    }

    setupDropZone(dropZoneId, inputId, handleFileCallback) {
        const dropZone = document.getElementById(dropZoneId);
        const fileInput = document.getElementById(inputId);

        if (!dropZone || !fileInput) return;

        // Click to browse
        dropZone.addEventListener('click', () => {
            fileInput.click();
        });

        // File input change
        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileCallback(e.target.files[0]);
            }
        });

        // Drag and drop events
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFileCallback(files[0]);
            }
        });
    }

    async handleTemplateFile(file) {
        // Validate file type
        if (!file.name.toLowerCase().endsWith('.docx') && !file.name.toLowerCase().endsWith('.doc')) {
            this.showError('Please select a Word document (.docx or .doc)');
            return;
        }

        // Update UI
        const dropZone = document.getElementById('template-drop');
        const dropText = dropZone.querySelector('.drop-text');
        const dropIcon = dropZone.querySelector('.drop-icon');

        dropText.textContent = `Uploading ${file.name}...`;
        dropIcon.textContent = '⏳';

        try {
            // Upload file
            const formData = new FormData();
            formData.append('file', file);

            const response = await fetch('/upload_template', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.templateFile = file;
                this.templatePath = result.filepath;

                // Update UI - success
                dropText.textContent = `✓ ${result.filename}`;
                dropIcon.textContent = '📄';
                dropZone.style.borderColor = '#48bb78';
                dropZone.style.backgroundColor = '#f0fff4';

                this.updateMergeButton();
            } else {
                throw new Error(result.error);
            }

        } catch (error) {
            // Update UI - error
            dropText.textContent = 'Upload failed - click to retry';
            dropIcon.textContent = '❌';
            dropZone.style.borderColor = '#e53e3e';

            this.showError(`Template upload failed: ${error.message}`);
        }
    }

    async handleDataFile(file) {
        // Validate file type
        if (!file.name.toLowerCase().endsWith('.xlsx') && !file.name.toLowerCase().endsWith('.xls')) {
            this.showError('Please select an Excel file (.xlsx or .xls)');
            return;
        }

        // Update UI
        const dropZone = document.getElementById('data-drop');
        const dropText = dropZone.querySelector('.drop-text');
        const dropIcon = dropZone.querySelector('.drop-icon');

        dropText.textContent = `Uploading ${file.name}...`;
        dropIcon.textContent = '⏳';

        try {
            // Upload file
            const formData = new FormData();
            formData.append('file', file);

            const response = await fetch('/upload_data', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.dataFile = file;
                this.dataPath = result.filepath;

                // Update UI - success
                dropText.textContent = `✓ ${result.filename} (${result.total_rows} rows)`;
                dropIcon.textContent = '📊';
                dropZone.style.borderColor = '#48bb78';
                dropZone.style.backgroundColor = '#f0fff4';

                // Show data preview
                this.showDataPreview(result.preview, result.columns);

                this.updateMergeButton();
            } else {
                throw new Error(result.error);
            }

        } catch (error) {
            // Update UI - error
            dropText.textContent = 'Upload failed - click to retry';
            dropIcon.textContent = '❌';
            dropZone.style.borderColor = '#e53e3e';

            this.showError(`Data upload failed: ${error.message}`);
        }
    }

    showDataPreview(preview, columns) {
        // Create or update preview section
        let previewSection = document.querySelector('.data-preview');

        if (!previewSection) {
            previewSection = document.createElement('div');
            previewSection.className = 'data-preview';
            const dataDropParent = document.getElementById('data-drop').parentNode;
            dataDropParent.appendChild(previewSection);
        }

        previewSection.innerHTML = `
            <h4 style="margin: 20px 0 10px; color: #2d3748;">Data Preview:</h4>
            <div style="background: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px; font-size: 12px; overflow-x: auto;">
                <strong>Columns:</strong> ${columns.join(', ')}<br><br>
                <strong>Sample rows:</strong>
                <table style="width: 100%; border-collapse: collapse; margin-top: 8px;">
                    <thead>
                        <tr>
                            ${columns.map(col => `<th style="border: 1px solid #cbd5e0; padding: 8px; background: #f7fafc; text-align: left;">${col}</th>`).join('')}
                        </tr>
                    </thead>
                    <tbody>
                        ${preview.map(row =>
            `<tr>${columns.map(col => `<td style="border: 1px solid #cbd5e0; padding: 8px;">${row[col] || ''}</td>`).join('')}</tr>`
        ).join('')}
                    </tbody>
                </table>
            </div>
        `;
    }

    updateMergeButton() {
        const mergeBtn = document.querySelector('.merge-btn');
        if (!mergeBtn) return;

        const canProcess = this.templatePath && this.dataPath && this.selectedFormat;

        mergeBtn.disabled = !canProcess;

        if (canProcess) {
            mergeBtn.textContent = 'Start Mail Merge';
            mergeBtn.style.backgroundColor = '#4299e1';
        } else {
            mergeBtn.textContent = 'Upload files and select format first';
            mergeBtn.style.backgroundColor = '#cbd5e0';
        }
    }

    async processMerge() {
        if (!this.templatePath || !this.dataPath || !this.selectedFormat) {
            this.showError('Please upload both files and select an output format');
            return;
        }

        const mergeBtn = document.querySelector('.merge-btn');
        const originalText = mergeBtn.textContent;

        // Update UI - processing
        mergeBtn.disabled = true;
        mergeBtn.textContent = 'Processing...';
        mergeBtn.style.backgroundColor = '#cbd5e0';

        try {
            const response = await fetch('/process_merge', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    template_path: this.templatePath,
                    data_path: this.dataPath,
                    output_format: this.selectedFormat
                })
            });

            const result = await response.json();

            if (result.success) {
                // Show success and download link
                this.showSuccess(result.message);
                this.showDownloadLink(result.download_url, result.filename);

                mergeBtn.textContent = 'Process Complete!';
                mergeBtn.style.backgroundColor = '#48bb78';
            } else {
                throw new Error(result.error);
            }

        } catch (error) {
            this.showError(`Processing failed: ${error.message}`);

            // Reset button
            mergeBtn.disabled = false;
            mergeBtn.textContent = originalText;
            mergeBtn.style.backgroundColor = '#4299e1';
        }
    }

    showDownloadLink(downloadUrl, filename) {
        // Create or update download section
        let downloadSection = document.querySelector('.download-section');

        if (!downloadSection) {
            downloadSection = document.createElement('div');
            downloadSection.className = 'download-section';
            downloadSection.style.cssText = `
                margin: 30px auto 0;
                text-align: center;
                padding: 20px;
                background: #f0fff4;
                border: 1px solid #48bb78;
                border-radius: 12px;
                max-width: 400px;
            `;
            const container = document.querySelector('.steps-section .container');
            if (container) {
                container.appendChild(downloadSection);
            }
        }

        downloadSection.innerHTML = `
            <h3 style="color: #22543d; margin-bottom: 16px;">✅ Processing Complete!</h3>
            <p style="margin-bottom: 20px; color: #2d3748;">Your merged documents are ready for download.</p>
            <a href="${downloadUrl}" 
               style="display: inline-block; background: #48bb78; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: 500;"
               download="${filename}">
                📥 Download ${filename}
            </a>
        `;
    }

    showSuccess(message) {
        this.showNotification(message, 'success');
    }

    showError(message) {
        this.showNotification(message, 'error');
    }

    showNotification(message, type) {
        // Remove existing notifications
        document.querySelectorAll('.notification').forEach(n => n.remove());

        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.style.cssText = `
            position: fixed;
            top: 20px;
            right: 20px;
            padding: 16px 20px;
            border-radius: 8px;
            color: white;
            font-weight: 500;
            z-index: 1000;
            max-width: 400px;
            background-color: ${type === 'success' ? '#48bb78' : '#e53e3e'};
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        `;

        notification.textContent = message;
        document.body.appendChild(notification);

        // Auto remove after 5 seconds
        setTimeout(() => {
            notification.remove();
        }, 5000);

        // Click to remove
        notification.addEventListener('click', () => {
            notification.remove();
        });
    }
}

// Initialize the app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new MailMergeApp();
});

// Add some additional CSS for better drag & drop visual feedback
const additionalCSS = `
    .drop-zone.dragover {
        border-color: #4299e1 !important;
        background-color: #ebf8ff !important;
        transform: scale(1.02);
    }
    
    .drop-zone {
        transition: all 0.3s ease;
    }
    
    .notification {
        cursor: pointer;
        transition: opacity 0.3s ease;
    }
    
    .notification:hover {
        opacity: 0.9;
    }
`;

// Inject additional CSS
const styleElement = document.createElement('style');
styleElement.textContent = additionalCSS;
document.head.appendChild(styleElement);