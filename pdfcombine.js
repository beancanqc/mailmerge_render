/**
 * PDF Combine Tool - Enhanced JavaScript
 * Office Sigma Document Tools
 */

class PDFCombineTool {
    constructor() {
        this.pdfFiles = [];
        this.currentStep = 1;
        this.isProcessing = false;
        this.maxFiles = 20;
        this.maxTotalSize = 200 * 1024 * 1024; // 200MB
        this.maxFileSize = 50 * 1024 * 1024; // 50MB
        this.uploadProgress = {};

        this.init();
    }

    // Initialize the application
    init() {
        this.bindEvents();
        this.loadExistingPDFs();
        this.updateUI();
    }

    // Bind all event listeners
    bindEvents() {
        const uploadArea = document.getElementById('pdf-upload-area');
        const fileInput = document.getElementById('pdf-file-input');

        // Drag and drop events
        uploadArea.addEventListener('dragover', (e) => this.handleDragOver(e));
        uploadArea.addEventListener('dragleave', (e) => this.handleDragLeave(e));
        uploadArea.addEventListener('drop', (e) => this.handleDrop(e));
        uploadArea.addEventListener('click', () => fileInput.click());

        // File input change
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e));

        // Navigation buttons
        document.getElementById('clear-all-btn')?.addEventListener('click', () => this.clearAllPDFs());
        document.getElementById('proceed-to-combine')?.addEventListener('click', () => this.proceedToArrange());
        document.getElementById('back-to-upload')?.addEventListener('click', () => this.backToUpload());
        document.getElementById('combine-pdfs-btn')?.addEventListener('click', () => this.combinePDFs());
        document.getElementById('start-new')?.addEventListener('click', () => this.startNew());

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => this.handleKeyboardShortcuts(e));

        // Window beforeunload warning
        window.addEventListener('beforeunload', (e) => this.handleBeforeUnload(e));
    }

    // Handle keyboard shortcuts
    handleKeyboardShortcuts(e) {
        if (e.ctrlKey || e.metaKey) {
            switch (e.key.toLowerCase()) {
                case 'u':
                    e.preventDefault();
                    if (this.currentStep === 1) {
                        document.getElementById('pdf-file-input').click();
                    }
                    break;
                case 'enter':
                    e.preventDefault();
                    if (this.currentStep === 1 && this.pdfFiles.length >= 2) {
                        this.proceedToArrange();
                    } else if (this.currentStep === 2) {
                        this.combinePDFs();
                    }
                    break;
                case 'backspace':
                    e.preventDefault();
                    if (this.currentStep === 2) {
                        this.backToUpload();
                    }
                    break;
            }
        }

        // Escape key
        if (e.key === 'Escape') {
            this.hideAllToasts();
        }
    }

    // Handle page unload warning
    handleBeforeUnload(e) {
        if (this.pdfFiles.length > 0 && this.currentStep < 3) {
            e.preventDefault();
            return e.returnValue = 'You have uploaded PDF files. Are you sure you want to leave?';
        }
    }

    // Enhanced drag and drop handlers
    handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();

        const uploadArea = e.target.closest('.upload-area');
        if (uploadArea) {
            uploadArea.classList.add('dragover');

            // Visual feedback for file types
            const items = Array.from(e.dataTransfer.items);
            const hasPDFs = items.some(item => item.type === 'application/pdf');

            if (hasPDFs) {
                uploadArea.classList.add('valid-files');
            } else {
                uploadArea.classList.add('invalid-files');
            }
        }
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();

        const uploadArea = e.target.closest('.upload-area');
        if (uploadArea && !uploadArea.contains(e.relatedTarget)) {
            uploadArea.classList.remove('dragover', 'valid-files', 'invalid-files');
        }
    }

    async handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();

        const uploadArea = e.target.closest('.upload-area');
        if (uploadArea) {
            uploadArea.classList.remove('dragover', 'valid-files', 'invalid-files');
        }

        const files = Array.from(e.dataTransfer.files);
        await this.uploadFiles(files);
    }

    async handleFileSelect(e) {
        const files = Array.from(e.target.files);
        await this.uploadFiles(files);
        e.target.value = ''; // Reset input
    }

    // Enhanced file validation
    validateFiles(files) {
        const errors = [];
        const validFiles = [];

        for (const file of files) {
            // Check file type
            if (file.type !== 'application/pdf') {
                errors.push(`${file.name}: Not a PDF file`);
                continue;
            }

            // Check file size
            if (file.size > this.maxFileSize) {
                const sizeMB = Math.round(file.size / (1024 * 1024));
                const maxMB = Math.round(this.maxFileSize / (1024 * 1024));
                errors.push(`${file.name}: Too large (${sizeMB}MB, max ${maxMB}MB)`);
                continue;
            }

            // Check if it would exceed file limit
            if (this.pdfFiles.length + validFiles.length >= this.maxFiles) {
                errors.push(`${file.name}: Maximum ${this.maxFiles} files allowed`);
                continue;
            }

            // Check if it would exceed total size limit
            const currentSize = this.pdfFiles.reduce((sum, pdf) => sum + pdf.size, 0);
            const newFilesSize = validFiles.reduce((sum, f) => sum + f.size, 0);

            if (currentSize + newFilesSize + file.size > this.maxTotalSize) {
                const totalMB = Math.round(this.maxTotalSize / (1024 * 1024));
                errors.push(`${file.name}: Would exceed total size limit (${totalMB}MB)`);
                continue;
            }

            validFiles.push(file);
        }

        return { validFiles, errors };
    }

    // Enhanced file upload with progress tracking
    async uploadFiles(files) {
        if (this.isProcessing) return;

        const { validFiles, errors } = this.validateFiles(files);

        // Show validation errors
        if (errors.length > 0) {
            this.showError(errors.join('\n'), 'Validation Errors');
        }

        if (validFiles.length === 0) return;

        this.showLoading(`Uploading ${validFiles.length} PDF file(s)...`);

        try {
            const uploadPromises = validFiles.map((file, index) =>
                this.uploadSingleFileWithProgress(file, index)
            );

            await Promise.all(uploadPromises);
            await this.loadExistingPDFs();

            this.hideLoading();
            this.showSuccess(`Successfully uploaded ${validFiles.length} PDF file(s)`);

        } catch (error) {
            this.hideLoading();
            this.showError(error.message || 'Failed to upload files');
        }
    }

    // Upload single file with progress tracking
    async uploadSingleFileWithProgress(file, index) {
        return new Promise((resolve, reject) => {
            const formData = new FormData();
            formData.append('file', file);

            const xhr = new XMLHttpRequest();

            // Track upload progress
            xhr.upload.addEventListener('progress', (e) => {
                if (e.lengthComputable) {
                    const percentComplete = Math.round((e.loaded / e.total) * 100);
                    this.updateUploadProgress(file.name, percentComplete);
                }
            });

            xhr.addEventListener('load', () => {
                try {
                    const result = JSON.parse(xhr.responseText);
                    if (result.success) {
                        resolve(result);
                    } else {
                        reject(new Error(result.error || 'Upload failed'));
                    }
                } catch (error) {
                    reject(new Error('Invalid server response'));
                }
            });

            xhr.addEventListener('error', () => {
                reject(new Error('Network error during upload'));
            });

            xhr.addEventListener('timeout', () => {
                reject(new Error('Upload timeout'));
            });

            xhr.timeout = 30000; // 30 second timeout
            xhr.open('POST', '/upload_pdf');
            xhr.send(formData);
        });
    }

    // Update upload progress display
    updateUploadProgress(filename, percent) {
        this.uploadProgress[filename] = percent;

        // Update loading message with progress
        const avgProgress = Object.values(this.uploadProgress)
            .reduce((sum, p) => sum + p, 0) / Object.keys(this.uploadProgress).length;

        this.updateLoadingMessage(`Uploading files... ${Math.round(avgProgress)}%`);
    }

    // Enhanced PDF list display with animations
    updatePDFList() {
        const listElement = document.getElementById('pdf-list');

        if (this.pdfFiles.length === 0) {
            listElement.innerHTML = '<p class="no-files">No PDF files uploaded yet.</p>';
            return;
        }

        const html = this.pdfFiles.map((pdf, index) => `
            <div class="pdf-item animate-in" data-index="${index}">
                <div class="pdf-icon">📄</div>
                <div class="pdf-info">
                    <div class="pdf-name" title="${pdf.filename}">${this.truncateFilename(pdf.filename)}</div>
                    <div class="pdf-details">
                        ${pdf.pages} page${pdf.pages !== 1 ? 's' : ''} • ${pdf.size_mb} MB
                        ${pdf.title ? ` • ${this.truncateText(pdf.title, 30)}` : ''}
                    </div>
                </div>
                <div class="pdf-actions">
                    <button class="action-btn move-up" onclick="pdfTool.movePDF(${index}, -1)" 
                            ${index === 0 ? 'disabled' : ''} title="Move up">↑</button>
                    <button class="action-btn move-down" onclick="pdfTool.movePDF(${index}, 1)" 
                            ${index === this.pdfFiles.length - 1 ? 'disabled' : ''} title="Move down">↓</button>
                    <button class="remove-btn" onclick="pdfTool.removePDF(${index})" title="Remove">×</button>
                </div>
            </div>
        `).join('');

        listElement.innerHTML = html;

        // Animate new items
        setTimeout(() => {
            listElement.querySelectorAll('.animate-in').forEach(item => {
                item.classList.remove('animate-in');
            });
        }, 50);
    }

    // Enhanced sortable functionality
    initializeSortable() {
        const sortableList = document.getElementById('sortable-pdf-list');

        const html = this.pdfFiles.map((pdf, index) => `
            <div class="sortable-item" data-index="${index}" draggable="true">
                <div class="drag-handle" title="Drag to reorder">⋮⋮</div>
                <div class="pdf-icon">📄</div>
                <div class="pdf-info">
                    <div class="pdf-name" title="${pdf.filename}">${this.truncateFilename(pdf.filename)}</div>
                    <div class="pdf-details">
                        ${pdf.pages} page${pdf.pages !== 1 ? 's' : ''} • ${pdf.size_mb} MB
                        ${pdf.title ? ` • ${this.truncateText(pdf.title, 30)}` : ''}
                    </div>
                </div>
                <div class="order-number">#${index + 1}</div>
                <button class="remove-btn-small" onclick="pdfTool.removePDF(${index})" title="Remove">×</button>
            </div>
        `).join('');

        sortableList.innerHTML = html;
        this.addSortableHandlers();
    }

    // Enhanced sortable drag and drop with better UX
    addSortableHandlers() {
        const items = document.querySelectorAll('.sortable-item');
        let draggedItem = null;
        let draggedOverItem = null;

        items.forEach(item => {
            item.addEventListener('dragstart', (e) => {
                draggedItem = item;
                item.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';

                // Add ghost image
                e.dataTransfer.setData('text/html', item.outerHTML);
            });

            item.addEventListener('dragend', (e) => {
                item.classList.remove('dragging');
                this.clearDragStates();
                draggedItem = null;
                draggedOverItem = null;
            });

            item.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';

                if (draggedItem && draggedItem !== item) {
                    draggedOverItem = item;
                    item.classList.add('drag-over');
                }
            });

            item.addEventListener('dragleave', (e) => {
                if (!item.contains(e.relatedTarget)) {
                    item.classList.remove('drag-over');
                }
            });

            item.addEventListener('drop', (e) => {
                e.preventDefault();

                if (draggedItem && draggedItem !== item) {
                    this.handleSortableDrop(draggedItem, item);
                }

                this.clearDragStates();
            });
        });
    }

    // Handle sortable drop with animation
    handleSortableDrop(draggedItem, targetItem) {
        const container = targetItem.parentNode;
        const items = Array.from(container.children);
        const draggedIndex = items.indexOf(draggedItem);
        const targetIndex = items.indexOf(targetItem);

        if (draggedIndex === targetIndex) return;

        // Animate the move
        targetItem.classList.add('drop-target');

        setTimeout(() => {
            if (draggedIndex > targetIndex) {
                container.insertBefore(draggedItem, targetItem);
            } else {
                container.insertBefore(draggedItem, targetItem.nextSibling);
            }

            this.updateSortOrder();
            targetItem.classList.remove('drop-target');
        }, 150);
    }

    // Clear all drag states
    clearDragStates() {
        document.querySelectorAll('.sortable-item').forEach(item => {
            item.classList.remove('dragging', 'drag-over', 'drop-target');
        });
    }

    // Move PDF up/down in list
    async movePDF(index, direction) {
        if (index + direction < 0 || index + direction >= this.pdfFiles.length) {
            return;
        }

        const newOrder = [...Array(this.pdfFiles.length).keys()];
        [newOrder[index], newOrder[index + direction]] = [newOrder[index + direction], newOrder[index]];

        await this.reorderPDFs(newOrder);
        this.updateUI();
    }

    // Enhanced utility functions
    truncateFilename(filename, maxLength = 40) {
        if (filename.length <= maxLength) return filename;

        const extension = filename.split('.').pop();
        const nameWithoutExt = filename.slice(0, filename.lastIndexOf('.'));
        const availableLength = maxLength - extension.length - 4; // 4 for "..." and "."

        return nameWithoutExt.slice(0, availableLength) + '...' + extension;
    }

    truncateText(text, maxLength = 50) {
        return text.length > maxLength ? text.slice(0, maxLength) + '...' : text;
    }

    // Enhanced error handling with categorization
    showError(message, title = 'Error', category = 'general') {
        const errorToast = document.getElementById('error-toast');
        const errorMessage = document.getElementById('error-message');

        errorMessage.innerHTML = title !== 'Error' ?
            `<strong>${title}</strong><br>${message.replace(/\n/g, '<br>')}` :
            message.replace(/\n/g, '<br>');

        errorToast.className = `error-toast error-${category}`;
        errorToast.style.display = 'block';

        // Auto-hide after delay based on message length
        const hideDelay = Math.max(3000, message.length * 50);
        setTimeout(() => this.hideError(), hideDelay);
    }

    // Enhanced success messages with icons
    showSuccess(message, autoHide = true) {
        const successToast = document.getElementById('success-toast');
        const successMessage = document.getElementById('success-message');

        successMessage.textContent = message;
        successToast.style.display = 'block';

        if (autoHide) {
            setTimeout(() => this.hideSuccess(), 3000);
        }
    }

    // Enhanced loading with progress and cancellation
    showLoading(message = 'Processing...', showProgress = false) {
        document.getElementById('loading-text').textContent = message;
        document.getElementById('loading-overlay').style.display = 'flex';

        // Add cancel button for long operations
        if (message.includes('Combining')) {
            this.addCancelButton();
        }
    }

    updateLoadingMessage(message) {
        const loadingText = document.getElementById('loading-text');
        if (loadingText) {
            loadingText.textContent = message;
        }
    }

    addCancelButton() {
        const overlay = document.getElementById('loading-overlay');
        if (!overlay.querySelector('.cancel-btn')) {
            const cancelBtn = document.createElement('button');
            cancelBtn.className = 'cancel-btn';
            cancelBtn.textContent = 'Cancel';
            cancelBtn.onclick = () => this.cancelOperation();
            overlay.appendChild(cancelBtn);
        }
    }

    cancelOperation() {
        // Implementation for canceling ongoing operations
        this.hideLoading();
        this.showError('Operation cancelled by user');
    }

    hideAllToasts() {
        this.hideError();
        this.hideSuccess();
        this.hideLoading();
    }

    hideError() {
        document.getElementById('error-toast').style.display = 'none';
    }

    hideSuccess() {
        document.getElementById('success-toast').style.display = 'none';
    }

    hideLoading() {
        const overlay = document.getElementById('loading-overlay');
        overlay.style.display = 'none';

        // Remove cancel button
        const cancelBtn = overlay.querySelector('.cancel-btn');
        if (cancelBtn) {
            cancelBtn.remove();
        }

        // Clear progress tracking
        this.uploadProgress = {};
    }

    // API calls with enhanced error handling and retry logic
    async apiCall(url, options = {}, retries = 3) {
        for (let i = 0; i < retries; i++) {
            try {
                const response = await fetch(url, {
                    ...options,
                    headers: {
                        'Cache-Control': 'no-cache',
                        ...options.headers
                    }
                });

                if (!response.ok) {
                    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
                }

                const result = await response.json();

                if (!result.success && result.error) {
                    throw new Error(result.error);
                }

                return result;

            } catch (error) {
                if (i === retries - 1) throw error;

                // Wait before retry
                await new Promise(resolve => setTimeout(resolve, 1000 * (i + 1)));
            }
        }
    }

    // Enhanced load existing PDFs with error recovery
    async loadExistingPDFs() {
        try {
            const result = await this.apiCall('/get_pdf_list');
            this.pdfFiles = result.pdf_list || [];
            this.updateUI();
        } catch (error) {
            console.error('Failed to load PDF list:', error);
            // Don't show error toast for initial load failures
        }
    }

    // All other methods remain the same but use enhanced API calls...
    async removePDF(index) {
        if (this.isProcessing) return;

        this.showLoading('Removing PDF...');

        try {
            const result = await this.apiCall(`/remove_pdf/${index}`, { method: 'DELETE' });
            this.pdfFiles = result.pdf_list || [];
            this.updateUI();
            this.showSuccess('PDF removed successfully');
        } catch (error) {
            this.showError(error.message || 'Failed to remove PDF');
        } finally {
            this.hideLoading();
        }
    }

    async clearAllPDFs() {
        if (this.isProcessing || this.pdfFiles.length === 0) return;

        if (!confirm('Are you sure you want to remove all PDF files?')) {
            return;
        }

        this.showLoading('Clearing all PDFs...');

        try {
            await this.apiCall('/clear_pdf_queue', { method: 'POST' });
            this.pdfFiles = [];
            this.updateUI();
            this.showSuccess('All PDFs cleared successfully');
        } catch (error) {
            this.showError(error.message || 'Failed to clear PDFs');
        } finally {
            this.hideLoading();
        }
    }

    async reorderPDFs(newOrder) {
        try {
            const result = await this.apiCall('/reorder_pdfs', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ order: newOrder })
            });

            this.pdfFiles = result.pdf_list || [];
        } catch (error) {
            console.error('Failed to reorder PDFs:', error);
            this.showError('Failed to save new order');
        }
    }

    async combinePDFs() {
        if (this.isProcessing || this.pdfFiles.length < 2) return;

        this.showLoading('Combining PDF files...', true);
        this.isProcessing = true;

        try {
            const result = await this.apiCall('/combine_pdfs', { method: 'POST' });

            this.currentStep = 3;
            this.showSection('result-section');
            this.updateStepIndicators();

            // Set up download
            document.getElementById('result-message').textContent = result.message;
            document.getElementById('download-btn').onclick = () => {
                window.open(result.download_url, '_blank');
            };

            this.showSuccess('PDFs combined successfully!');

        } catch (error) {
            this.showError(error.message || 'Failed to combine PDFs');
        } finally {
            this.hideLoading();
            this.isProcessing = false;
        }
    }

    // Navigation methods
    proceedToArrange() {
        if (this.pdfFiles.length < 2) {
            this.showError('At least 2 PDF files are required for combination.');
            return;
        }

        this.currentStep = 2;
        this.showSection('arrange-section');
        this.updateStepIndicators();
        this.initializeSortable();
    }

    backToUpload() {
        this.currentStep = 1;
        this.showSection('upload-section');
        this.updateStepIndicators();
    }

    startNew() {
        this.clearAllPDFs();
        this.currentStep = 1;
        this.showSection('upload-section');
        this.updateStepIndicators();
    }

    showSection(sectionId) {
        const sections = ['upload-section', 'arrange-section', 'result-section'];
        sections.forEach(id => {
            const element = document.getElementById(id);
            if (element) {
                element.style.display = id === sectionId ? 'block' : 'none';
            }
        });
    }

    updateStepIndicators() {
        for (let i = 1; i <= 3; i++) {
            const step = document.getElementById(`step${i}`);
            if (step) {
                step.classList.toggle('active', i === this.currentStep);
                step.classList.toggle('completed', i < this.currentStep);
            }
        }
    }

    updateUI() {
        this.updateStats();
        this.updatePDFList();
        this.updateStepIndicators();
    }

    updateStats() {
        const stats = this.calculateStats();

        const fileCountEl = document.getElementById('file-count');
        const totalSizeEl = document.getElementById('total-size');
        const totalPagesEl = document.getElementById('total-pages');

        if (fileCountEl) fileCountEl.textContent = `${stats.fileCount}/${this.maxFiles}`;
        if (totalSizeEl) totalSizeEl.textContent = `${stats.totalSizeMB} MB / ${Math.round(this.maxTotalSize / (1024 * 1024))} MB`;
        if (totalPagesEl) totalPagesEl.textContent = stats.totalPages;

        // Show/hide PDF list container
        const listContainer = document.getElementById('pdf-list-container');
        if (listContainer) {
            listContainer.style.display = stats.fileCount > 0 ? 'block' : 'none';
        }
    }

    calculateStats() {
        const fileCount = this.pdfFiles.length;
        const totalSize = this.pdfFiles.reduce((sum, pdf) => sum + pdf.size, 0);
        const totalPages = this.pdfFiles.reduce((sum, pdf) => sum + pdf.pages, 0);

        return {
            fileCount,
            totalSizeMB: Math.round(totalSize / (1024 * 1024) * 10) / 10,
            totalPages
        };
    }

    updateSortOrder() {
        const items = document.querySelectorAll('.sortable-item');
        const newOrder = Array.from(items).map(item => parseInt(item.dataset.index));

        // Update order numbers visually
        items.forEach((item, index) => {
            const orderEl = item.querySelector('.order-number');
            if (orderEl) {
                orderEl.textContent = `#${index + 1}`;
            }
        });

        // Send new order to server
        this.reorderPDFs(newOrder);
    }
}

// Initialize the tool when DOM is loaded
document.addEventListener('DOMContentLoaded', function () {
    window.pdfTool = new PDFCombineTool();
});

// Global functions for onclick handlers
window.removePDF = function (index) {
    window.pdfTool?.removePDF(index);
};

window.movePDF = function (index, direction) {
    window.pdfTool?.movePDF(index, direction);
};