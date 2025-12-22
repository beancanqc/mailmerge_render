// Mail Merge SaaS - .NET Edition - Frontend JavaScript
class MailMergeApp {
    constructor() {
        this.selectedOutputType = null;
        this.hasTemplate = false;
        this.hasData = false;
        this.init();
    }

    init() {
        console.log('[MailMerge] Initializing .NET Edition Mail Merge App');
        this.setupFileUploads();
        this.setupOutputSelection();
        this.checkExistingSession();
    }

    setupFileUploads() {
        // Template upload
        const templateZone = document.getElementById('template-zone');
        const templateInput = document.getElementById('template-input');

        this.setupDropZone(templateZone, templateInput, this.handleTemplateUpload.bind(this));
        templateInput.addEventListener('change', (e) => this.handleTemplateUpload(e.target.files[0]));

        // Data upload  
        const dataZone = document.getElementById('data-zone');
        const dataInput = document.getElementById('data-input');

        this.setupDropZone(dataZone, dataInput, this.handleDataUpload.bind(this));
        dataInput.addEventListener('change', (e) => this.handleDataUpload(e.target.files[0]));
    }

    setupDropZone(zone, input, handler) {
        zone.addEventListener('dragover', (e) => {
            e.preventDefault();
            if (!zone.classList.contains('disabled')) {
                zone.classList.add('dragover');
            }
        });

        zone.addEventListener('dragleave', () => {
            zone.classList.remove('dragover');
        });

        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');

            if (zone.classList.contains('disabled')) return;

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handler(files[0]);
            }
        });

        zone.addEventListener('click', () => {
            if (!zone.classList.contains('disabled')) {
                input.click();
            }
        });
    }

    async handleTemplateUpload(file) {
        if (!file) return;

        if (!file.name.toLowerCase().endsWith('.docx')) {
            this.showError('Please select a .docx file for the template.');
            return;
        }

        this.updateStepStatus('step1', 'uploading', 'Uploading...');

        const formData = new FormData();
        formData.append('template', file);

        try {
            const response = await fetch('/MailMerge/UploadTemplate', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.hasTemplate = true;
                this.updateStepStatus('step1', 'completed', 'Template uploaded ✓');
                this.enableStep('step2');
                console.log('[MailMerge] Template uploaded successfully');
            } else {
                this.updateStepStatus('step1', 'error', 'Upload failed');
                this.showError(result.error || 'Failed to upload template');
            }
        } catch (error) {
            console.error('[MailMerge] Template upload error:', error);
            this.updateStepStatus('step1', 'error', 'Upload failed');
            this.showError('Network error occurred');
        }
    }

    async handleDataUpload(file) {
        if (!file) return;

        if (!file.name.toLowerCase().endsWith('.xlsx')) {
            this.showError('Please select an .xlsx file for the data.');
            return;
        }

        this.updateStepStatus('step2', 'uploading', 'Uploading...');

        const formData = new FormData();
        formData.append('data', file);

        try {
            const response = await fetch('/MailMerge/UploadData', {
                method: 'POST',
                body: formData
            });

            const result = await response.json();

            if (result.success) {
                this.hasData = true;
                this.updateStepStatus('step2', 'completed', 'Data uploaded ✓');
                this.displayDataPreview(result.preview);
                this.enableStep('step3');
                console.log('[MailMerge] Data uploaded successfully');
            } else {
                this.updateStepStatus('step2', 'error', 'Upload failed');
                this.showError(result.error || 'Failed to upload data');
            }
        } catch (error) {
            console.error('[MailMerge] Data upload error:', error);
            this.updateStepStatus('step2', 'error', 'Upload failed');
            this.showError('Network error occurred');
        }
    }

    displayDataPreview(preview) {
        const previewDiv = document.getElementById('data-preview');
        const contentDiv = document.getElementById('preview-content');

        if (!preview || !preview.headers || !preview.data) return;

        let html = '<table><thead><tr>';
        preview.headers.forEach(header => {
            html += `<th>${this.escapeHtml(header)}</th>`;
        });
        html += '</tr></thead><tbody>';

        preview.data.forEach(row => {
            html += '<tr>';
            preview.headers.forEach(header => {
                const value = row[header] || '';
                html += `<td>${this.escapeHtml(value)}</td>`;
            });
            html += '</tr>';
        });

        html += '</tbody></table>';

        if (preview.data.length >= 5) {
            html += '<p><small>Showing first 5 rows...</small></p>';
        }

        contentDiv.innerHTML = html;
        previewDiv.style.display = 'block';
    }

    setupOutputSelection() {
        const options = document.querySelectorAll('.output-option');
        options.forEach(option => {
            option.addEventListener('click', () => {
                if (document.getElementById('output-options').classList.contains('disabled')) return;

                // Remove active class from all options
                options.forEach(opt => opt.classList.remove('active'));

                // Add active class to selected option
                option.classList.add('active');

                this.selectedOutputType = option.dataset.type;
                this.updateStepStatus('step3', 'completed', 'Format selected ✓');

                // Start processing
                this.processMailMerge();
            });
        });
    }

    async processMailMerge() {
        if (!this.selectedOutputType) return;

        // Show processing
        document.getElementById('processing').style.display = 'block';
        document.getElementById('processing-message').textContent = 'Processing your mail merge with perfect formatting...';

        const [outputFormat, multiple] = this.selectedOutputType.split('-');
        const isMultiple = multiple === 'multiple';

        try {
            const response = await fetch('/MailMerge/ProcessMerge', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    outputType: outputFormat,
                    multipleFiles: isMultiple
                })
            });

            const result = await response.json();

            if (result.success) {
                this.showDownloads(result.files);
                console.log('[MailMerge] Processing completed successfully');
            } else {
                this.hideProcessing();
                this.showError(result.error || 'Processing failed');
            }
        } catch (error) {
            console.error('[MailMerge] Processing error:', error);
            this.hideProcessing();
            this.showError('Network error occurred during processing');
        }
    }

    showDownloads(files) {
        document.getElementById('processing').style.display = 'none';

        const downloadSection = document.getElementById('download-section');
        const downloadLinks = document.getElementById('download-links');

        let html = '';
        files.forEach(filename => {
            const extension = filename.split('.').pop().toLowerCase();
            const icon = extension === 'pdf' ? '📕' : '📄';

            html += `
                <a href="/MailMerge/Download/${encodeURIComponent(filename)}" class="download-link">
                    <span class="file-icon">${icon}</span>
                    <span class="file-name">${this.escapeHtml(filename)}</span>
                    <span class="download-icon">⬇️</span>
                </a>
            `;
        });

        downloadLinks.innerHTML = html;
        downloadSection.style.display = 'block';
    }

    async checkExistingSession() {
        try {
            const response = await fetch('/MailMerge/CheckStatus');
            const status = await response.json();

            if (status.hasTemplate) {
                this.hasTemplate = true;
                this.updateStepStatus('step1', 'completed', 'Template ready ✓');
                this.enableStep('step2');
            }

            if (status.hasData) {
                this.hasData = true;
                this.updateStepStatus('step2', 'completed', 'Data ready ✓');
                this.enableStep('step3');
            }

            if (status.outputFiles && status.outputFiles.length > 0) {
                this.showDownloads(status.outputFiles);
            }
        } catch (error) {
            console.error('[MailMerge] Session check error:', error);
        }
    }

    updateStepStatus(stepId, type, message) {
        const stepContainer = document.getElementById(stepId);
        const statusElement = document.getElementById(`${stepId}-status`);

        // Reset classes
        stepContainer.classList.remove('completed', 'error');

        // Add appropriate class and update message
        if (type === 'completed') {
            stepContainer.classList.add('completed');
        } else if (type === 'error') {
            stepContainer.classList.add('error');
        }

        statusElement.textContent = message;
    }

    enableStep(stepId) {
        const stepNumber = stepId.replace('step', '');

        if (stepNumber === '2') {
            document.getElementById('data-zone').classList.remove('disabled');
        } else if (stepNumber === '3') {
            document.getElementById('output-options').classList.remove('disabled');
        }
    }

    hideProcessing() {
        document.getElementById('processing').style.display = 'none';
    }

    showError(message) {
        alert(`Error: ${message}`);
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }
}

// Initialize app when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    window.mailMergeApp = new MailMergeApp();
});