// Split Word JavaScript functionality

document.addEventListener('DOMContentLoaded', function () {
    // State management
    let state = {
        documentUploaded: false,
        totalPages: 0,
        documentName: '',
        selectedPages: new Set(),
        currentRanges: [],
        processing: false,
        currentMethod: null,
        documentPages: []
    };

    // Get DOM elements
    const documentDrop = document.getElementById('document-drop');
    const documentInput = document.getElementById('document-input');
    const uploadStep = document.getElementById('upload-step');
    const methodStep = document.getElementById('method-step');
    const rangeConfigStep = document.getElementById('range-config-step');
    const pagesConfigStep = document.getElementById('pages-config-step');
    const processingDisplay = document.getElementById('processing-display');
    const downloadReady = document.getElementById('download-ready');
    const documentInfo = document.getElementById('document-info');

    const rangeMethodBtn = document.getElementById('range-method-btn');
    const pagesMethodBtn = document.getElementById('pages-method-btn');

    // Debug function
    function debugLog(message, data) {
        console.log('[SplitWord]', message, data || '');
    }

    // Initialize the application
    init();

    function init() {
        setupDropZone();
        setupMethodButtons();
        setupRangeConfiguration();
        setupPagesConfiguration();
        setupProcessingHandlers();
    }

    function setupDropZone() {
        if (!documentDrop || !documentInput) {
            console.error('Drop zone elements not found');
            return;
        }

        // File drop functionality
        documentDrop.addEventListener('click', () => documentInput.click());

        documentDrop.addEventListener('dragover', (e) => {
            e.preventDefault();
            documentDrop.classList.add('dragover');
        });

        documentDrop.addEventListener('dragleave', () => {
            documentDrop.classList.remove('dragover');
        });

        documentDrop.addEventListener('drop', (e) => {
            e.preventDefault();
            documentDrop.classList.remove('dragover');

            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFileUpload(files[0]);
            }
        });

        documentInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileUpload(e.target.files[0]);
            }
        });
    }

    function handleFileUpload(file) {
        if (!file.name.toLowerCase().endsWith('.docx')) {
            showError('Please select a .docx file');
            return;
        }

        debugLog('Uploading document:', file.name);
        state.documentName = file.name;

        // Update UI for upload state
        documentDrop.classList.add('uploading');
        documentDrop.querySelector('.drop-text').textContent = 'Uploading...';

        const formData = new FormData();
        formData.append('file', file);

        fetch('/upload_document', {
            method: 'POST',
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                documentDrop.classList.remove('uploading');

                if (data.success) {
                    debugLog('Document uploaded successfully');
                    loadDocumentPages();
                } else {
                    throw new Error(data.error);
                }
            })
            .catch(error => {
                documentDrop.classList.remove('uploading');
                documentDrop.querySelector('.drop-text').textContent = 'Drop your .docx file here';
                showError('Upload failed: ' + error.message);
            });
    }

    function loadDocumentPages() {
        fetch('/get_document_pages')
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    state.documentPages = data.pages;
                    state.totalPages = data.total_pages;
                    state.documentUploaded = true;

                    // Update document name from server data if available
                    if (data.pages && data.pages.length > 0 && data.pages[0].filename) {
                        state.documentName = data.pages[0].filename;
                    }

                    // Update UI
                    showDocumentInfo();
                    showMethodSelection();

                } else {
                    throw new Error(data.error);
                }
            })
            .catch(error => {
                showError('Failed to load document: ' + error.message);
            });
    }

    function showDocumentInfo() {
        document.getElementById('total-pages-count').textContent = state.totalPages;
        document.getElementById('document-name').textContent = state.documentName;
        documentInfo.style.display = 'block';
    }

    function showMethodSelection() {
        // Update document drop zone to show success
        documentDrop.classList.add('uploaded');
        documentDrop.querySelector('.drop-text').textContent = 'Document uploaded successfully';
        documentDrop.querySelector('.drop-icon').textContent = '✅';

        // Show method selection
        methodStep.style.display = 'block';
    }

    function setupMethodButtons() {
        if (rangeMethodBtn) {
            rangeMethodBtn.addEventListener('click', () => selectMethod('range'));
        }

        if (pagesMethodBtn) {
            pagesMethodBtn.addEventListener('click', () => selectMethod('pages'));
        }
    }

    function selectMethod(method) {
        state.currentMethod = method;

        // Hide method selection
        methodStep.style.display = 'none';

        if (method === 'range') {
            showRangeConfiguration();
        } else if (method === 'pages') {
            showPagesConfiguration();
        }
    }

    function setupRangeConfiguration() {
        const rangeTypeInputs = document.querySelectorAll('input[name="range-type"]');
        const customContainer = document.getElementById('custom-ranges-container');
        const fixedContainer = document.getElementById('fixed-ranges-container');
        const addRangeBtn = document.getElementById('add-range-btn');
        const generateFixedBtn = document.getElementById('generate-fixed-ranges-btn');
        const splitRangeBtn = document.getElementById('split-by-range-btn');

        // Range type toggle
        rangeTypeInputs.forEach(input => {
            input.addEventListener('change', (e) => {
                if (e.target.value === 'custom') {
                    customContainer.style.display = 'block';
                    fixedContainer.style.display = 'none';
                } else {
                    customContainer.style.display = 'none';
                    fixedContainer.style.display = 'block';
                }
                updateSplitRangeButton();
            });
        });

        // Add custom range
        if (addRangeBtn) {
            addRangeBtn.addEventListener('click', addCustomRange);
        }

        // Generate fixed ranges
        if (generateFixedBtn) {
            generateFixedBtn.addEventListener('click', generateFixedRanges);
        }

        // Split by range button
        if (splitRangeBtn) {
            splitRangeBtn.addEventListener('click', processSplitByRange);
        }
    }

    function showRangeConfiguration() {
        rangeConfigStep.style.display = 'block';
        // Add initial custom range
        if (state.currentRanges.length === 0) {
            addCustomRange();
        }
    }

    function addCustomRange() {
        const rangeId = `range-${Date.now()}`;
        const rangeItem = document.createElement('div');
        rangeItem.className = 'range-item';
        rangeItem.dataset.rangeId = rangeId;

        rangeItem.innerHTML = `
            <div class="range-inputs">
                <label>From page:</label>
                <input type="number" class="range-start" min="1" max="${state.totalPages}" value="1">
                <label>To page:</label>
                <input type="number" class="range-end" min="1" max="${state.totalPages}" value="${Math.min(5, state.totalPages)}">
                <button class="remove-range-btn" onclick="removeRange('${rangeId}')">×</button>
            </div>
            <div class="range-preview" id="preview-${rangeId}"></div>
        `;

        document.getElementById('ranges-list').appendChild(rangeItem);

        // Add event listeners for validation
        const startInput = rangeItem.querySelector('.range-start');
        const endInput = rangeItem.querySelector('.range-end');

        startInput.addEventListener('input', () => validateRange(rangeId));
        endInput.addEventListener('input', () => validateRange(rangeId));

        // Initial validation
        validateRange(rangeId);
        updateSplitRangeButton();
    }

    function removeRange(rangeId) {
        const rangeItem = document.querySelector(`[data-range-id="${rangeId}"]`);
        if (rangeItem) {
            rangeItem.remove();
            updateSplitRangeButton();
        }
    }

    function validateRange(rangeId) {
        const rangeItem = document.querySelector(`[data-range-id="${rangeId}"]`);
        const startInput = rangeItem.querySelector('.range-start');
        const endInput = rangeItem.querySelector('.range-end');
        const preview = rangeItem.querySelector('.range-preview');

        const start = parseInt(startInput.value);
        const end = parseInt(endInput.value);

        let isValid = true;
        let message = '';

        if (start < 1 || start > state.totalPages) {
            isValid = false;
            message = `Start page must be between 1 and ${state.totalPages}`;
        } else if (end < start || end > state.totalPages) {
            isValid = false;
            message = `End page must be between ${start} and ${state.totalPages}`;
        } else {
            message = `Pages ${start}-${end} (${end - start + 1} pages)`;
        }

        preview.textContent = message;
        preview.className = `range-preview ${isValid ? 'valid' : 'invalid'}`;

        rangeItem.dataset.valid = isValid;
        updateSplitRangeButton();
    }

    function generateFixedRanges() {
        const pagesPerRange = parseInt(document.getElementById('pages-per-range').value);
        const preview = document.getElementById('fixed-ranges-preview');

        if (pagesPerRange < 1) {
            preview.innerHTML = '<div class="error">Please enter a valid number of pages per range</div>';
            return;
        }

        const ranges = [];
        for (let start = 1; start <= state.totalPages; start += pagesPerRange) {
            const end = Math.min(start + pagesPerRange - 1, state.totalPages);
            ranges.push({ start, end });
        }

        state.currentRanges = ranges;

        const rangesHtml = ranges.map((range, index) => `
            <div class="fixed-range-item">
                <span class="range-label">Range ${index + 1}:</span>
                <span class="range-text">Pages ${range.start}-${range.end} (${range.end - range.start + 1} pages)</span>
            </div>
        `).join('');

        preview.innerHTML = rangesHtml;
        updateSplitRangeButton();
    }

    function updateSplitRangeButton() {
        const splitBtn = document.getElementById('split-by-range-btn');
        const rangeType = document.querySelector('input[name="range-type"]:checked').value;

        let hasValidRanges = false;

        if (rangeType === 'custom') {
            const rangeItems = document.querySelectorAll('.range-item');
            hasValidRanges = rangeItems.length > 0 &&
                Array.from(rangeItems).every(item => item.dataset.valid === 'true');
        } else {
            hasValidRanges = state.currentRanges.length > 0;
        }

        splitBtn.disabled = !hasValidRanges;
    }

    function setupPagesConfiguration() {
        const selectAllBtn = document.getElementById('select-all-pages');
        const deselectAllBtn = document.getElementById('deselect-all-pages');
        const splitPagesBtn = document.getElementById('split-by-pages-btn');

        if (selectAllBtn) {
            selectAllBtn.addEventListener('click', selectAllPages);
        }

        if (deselectAllBtn) {
            deselectAllBtn.addEventListener('click', deselectAllPages);
        }

        if (splitPagesBtn) {
            splitPagesBtn.addEventListener('click', processSplitByPages);
        }
    }

    function showPagesConfiguration() {
        pagesConfigStep.style.display = 'block';
        renderPagesGrid();
    }

    function renderPagesGrid() {
        const grid = document.getElementById('pages-grid');
        grid.innerHTML = '';

        state.documentPages.forEach(page => {
            const pageItem = document.createElement('div');
            pageItem.className = 'page-item';
            pageItem.dataset.pageNumber = page.page_number;

            pageItem.innerHTML = `
                <div class="page-thumbnail">
                    <div class="page-number">Page ${page.page_number}</div>
                    <div class="page-content-preview">${page.preview_text || 'Page content'}</div>
                    <div class="page-meta">${page.content_blocks} blocks</div>
                </div>
                <div class="page-checkbox">
                    <input type="checkbox" id="page-${page.page_number}" 
                           onchange="togglePage(${page.page_number})">
                    <label for="page-${page.page_number}"></label>
                </div>
            `;

            grid.appendChild(pageItem);
        });

        updatePagesSelectionCount();
    }

    window.togglePage = function (pageNumber) {
        if (state.selectedPages.has(pageNumber)) {
            state.selectedPages.delete(pageNumber);
        } else {
            state.selectedPages.add(pageNumber);
        }

        updatePagesDisplay();
        updateSplitPagesButton();
    };

    function selectAllPages() {
        state.selectedPages.clear();
        state.documentPages.forEach(page => {
            state.selectedPages.add(page.page_number);
        });
        updatePagesDisplay();
        updateSplitPagesButton();
    }

    function deselectAllPages() {
        state.selectedPages.clear();
        updatePagesDisplay();
        updateSplitPagesButton();
    }

    function updatePagesDisplay() {
        document.querySelectorAll('.page-item').forEach(item => {
            const pageNumber = parseInt(item.dataset.pageNumber);
            const checkbox = item.querySelector('input[type="checkbox"]');
            const isSelected = state.selectedPages.has(pageNumber);

            checkbox.checked = isSelected;
            item.classList.toggle('selected', isSelected);
        });

        updatePagesSelectionCount();
    }

    function updatePagesSelectionCount() {
        const count = state.selectedPages.size;
        const countDisplay = document.querySelector('.selection-count');
        if (countDisplay) {
            countDisplay.textContent = `${count} page${count !== 1 ? 's' : ''} selected`;
        }
    }

    function updateSplitPagesButton() {
        const splitBtn = document.getElementById('split-by-pages-btn');
        splitBtn.disabled = state.selectedPages.size === 0;
    }

    function processSplitByRange() {
        const rangeType = document.querySelector('input[name="range-type"]:checked').value;
        const outputType = document.querySelector('input[name="range-output"]:checked').value;

        let ranges = [];

        if (rangeType === 'custom') {
            const rangeItems = document.querySelectorAll('.range-item[data-valid="true"]');
            ranges = Array.from(rangeItems).map(item => {
                const start = parseInt(item.querySelector('.range-start').value);
                const end = parseInt(item.querySelector('.range-end').value);
                return { start, end };
            });
        } else {
            ranges = state.currentRanges;
        }

        if (ranges.length === 0) {
            showError('No valid ranges defined');
            return;
        }

        showProcessing();

        fetch('/split_by_range', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                ranges: ranges,
                output_type: outputType
            })
        })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showDownloadReady(data.download_url, data.filename);
                } else {
                    throw new Error(data.error);
                }
            })
            .catch(error => {
                hideProcessing();
                showError('Split failed: ' + error.message);
            });
    }

    function processSplitByPages() {
        const outputType = document.querySelector('input[name="pages-output"]:checked').value;
        const selectedPagesArray = Array.from(state.selectedPages).sort((a, b) => a - b);

        if (selectedPagesArray.length === 0) {
            showError('No pages selected');
            return;
        }

        showProcessing();

        fetch('/split_by_pages', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                pages: selectedPagesArray,
                output_type: outputType
            })
        })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    showDownloadReady(data.download_url, data.filename);
                } else {
                    throw new Error(data.error);
                }
            })
            .catch(error => {
                hideProcessing();
                showError('Split failed: ' + error.message);
            });
    }

    function setupProcessingHandlers() {
        const restartBtn = document.getElementById('restart-btn');
        if (restartBtn) {
            restartBtn.addEventListener('click', restartProcess);
        }
    }

    function showProcessing() {
        state.processing = true;
        rangeConfigStep.style.display = 'none';
        pagesConfigStep.style.display = 'none';
        processingDisplay.style.display = 'block';
    }

    function hideProcessing() {
        state.processing = false;
        processingDisplay.style.display = 'none';

        // Show appropriate config step
        if (state.currentMethod === 'range') {
            rangeConfigStep.style.display = 'block';
        } else {
            pagesConfigStep.style.display = 'block';
        }
    }

    function showDownloadReady(downloadUrl, filename) {
        processingDisplay.style.display = 'none';
        downloadReady.style.display = 'block';

        const downloadLink = document.getElementById('download-link');
        downloadLink.href = downloadUrl;
        downloadLink.download = filename;
    }

    function restartProcess() {
        // Reset state
        state = {
            documentUploaded: false,
            totalPages: 0,
            documentName: '',
            selectedPages: new Set(),
            currentRanges: [],
            processing: false,
            currentMethod: null,
            documentPages: []
        };

        // Reset UI
        uploadStep.style.display = 'block';
        methodStep.style.display = 'none';
        rangeConfigStep.style.display = 'none';
        pagesConfigStep.style.display = 'none';
        processingDisplay.style.display = 'none';
        downloadReady.style.display = 'none';
        documentInfo.style.display = 'none';

        // Reset document drop zone
        documentDrop.classList.remove('uploaded', 'uploading');
        documentDrop.querySelector('.drop-text').textContent = 'Drop your .docx file here';
        documentDrop.querySelector('.drop-icon').textContent = '📄';

        // Clear file input
        documentInput.value = '';
    }

    function showError(message) {
        // Simple error display - could be enhanced with toast notifications
        alert('Error: ' + message);
        debugLog('Error:', message);
    }

    // Make functions global for onclick handlers
    window.removeRange = removeRange;
    window.togglePage = window.togglePage || function () { };
});