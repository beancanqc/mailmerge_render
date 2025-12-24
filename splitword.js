// Split Word JavaScript functionality

document.addEventListener('DOMContentLoaded', function () {
    // Get DOM elements
    const documentDrop = document.getElementById('document-drop');
    const documentInput = document.getElementById('document-input');
    const methodStep = document.getElementById('method-step');
    const previewStep = document.getElementById('preview-step');
    const rangeConfig = document.getElementById('range-config');
    const pageSelection = document.getElementById('page-selection');
    const processStep = document.getElementById('process-step');

    const methodOptions = document.querySelectorAll('input[name="split-method"]');
    const rangeTabs = document.querySelectorAll('.range-tab');
    const rangeTabData = document.querySelectorAll('.range-panel');

    let documentUploaded = false;
    let currentDocument = null;
    let totalPages = 0;

    // Debug function
    function debugLog(message) {
        console.log('[SplitWord]', message);
    }

    // File upload functionality
    function setupDropZone() {
        if (!documentDrop || !documentInput) {
            console.error('Drop zone elements not found');
            return;
        }

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
                documentInput.files = files;
                handleFileUpload(files[0]);
            }
        });

        documentInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileUpload(e.target.files[0]);
            }
        });
    }

    // Handle file upload
    function handleFileUpload(file) {
        debugLog(`Uploading document: ${file.name}`);

        if (!file.name.toLowerCase().endsWith('.docx')) {
            alert('Please upload a .docx file');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        // Show loading state
        documentDrop.innerHTML = `
            <div class="drop-icon">⏳</div>
            <p class="drop-text">Processing ${file.name}...</p>
        `;

        fetch('/upload_splitdoc', {
            method: 'POST',
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                debugLog('Upload response:', data);

                if (data.success) {
                    documentUploaded = true;
                    currentDocument = data;
                    totalPages = data.page_count;

                    // Update upload UI
                    documentDrop.innerHTML = `
                    <div class="drop-icon">✅</div>
                    <p class="drop-text">${file.name}</p>
                    <p class="drop-subtext">${totalPages} pages</p>
                `;
                    documentDrop.style.borderColor = '#48bb78';
                    documentDrop.style.backgroundColor = '#f0fff4';

                    // Show next steps
                    showMethodSelection();
                    showDocumentPreview();

                } else {
                    // Error
                    documentDrop.innerHTML = `
                    <div class="drop-icon">❌</div>
                    <p class="drop-text">Upload failed</p>
                    <p class="drop-subtext">${data.error}</p>
                `;
                    documentDrop.style.borderColor = '#e53e3e';
                    documentDrop.style.backgroundColor = '#fed7d7';
                }
            })
            .catch(error => {
                debugLog('Upload error:', error);
                documentDrop.innerHTML = `
                <div class="drop-icon">❌</div>
                <p class="drop-text">Upload failed</p>
                <p class="drop-subtext">Network error</p>
            `;
                documentDrop.style.borderColor = '#e53e3e';
                documentDrop.style.backgroundColor = '#fed7d7';
            });
    }

    // Show method selection step
    function showMethodSelection() {
        methodStep.style.display = 'block';
    }

    // Show document preview
    function showDocumentPreview() {
        previewStep.style.display = 'block';
        document.getElementById('total-pages').textContent = totalPages;

        const pagesContainer = document.getElementById('pages-container');
        pagesContainer.innerHTML = '';

        // Create page previews (placeholder for now)
        for (let i = 1; i <= totalPages; i++) {
            const pageDiv = document.createElement('div');
            pageDiv.className = 'page-preview';
            pageDiv.innerHTML = `<span class="page-number">${i}</span>`;
            pageDiv.dataset.page = i;
            pagesContainer.appendChild(pageDiv);
        }
    }

    // Method selection handlers
    methodOptions.forEach(option => {
        option.addEventListener('change', (e) => {
            const method = e.target.value;
            debugLog(`Selected method: ${method}`);

            if (method === 'range') {
                showRangeConfig();
                hidePageSelection();
            } else if (method === 'select') {
                showPageSelection();
                hideRangeConfig();
            }

            showProcessStep();
        });
    });

    // Range configuration
    function showRangeConfig() {
        rangeConfig.style.display = 'block';
        updateFixedRangePreview();
    }

    function hideRangeConfig() {
        rangeConfig.style.display = 'none';
    }

    // Page selection
    function showPageSelection() {
        pageSelection.style.display = 'block';
        createSelectablePages();
    }

    function hidePageSelection() {
        pageSelection.style.display = 'none';
    }

    function createSelectablePages() {
        const selectableContainer = document.getElementById('selectable-pages');
        selectableContainer.innerHTML = '';

        for (let i = 1; i <= totalPages; i++) {
            const pageDiv = document.createElement('div');
            pageDiv.className = 'selectable-page';
            pageDiv.innerHTML = `
                <input type="checkbox" id="page-${i}" value="${i}">
                <label for="page-${i}">
                    <span class="page-number">${i}</span>
                </label>
            `;
            selectableContainer.appendChild(pageDiv);

            pageDiv.querySelector('input').addEventListener('change', updateSelectedCount);
        }
    }

    function updateSelectedCount() {
        const checked = document.querySelectorAll('#selectable-pages input:checked').length;
        document.getElementById('selected-count').textContent = checked;
        updateProcessButton();
    }

    // Range tab functionality
    rangeTabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            const tabType = e.target.dataset.tab;

            // Update tabs
            rangeTabs.forEach(t => t.classList.remove('active'));
            e.target.classList.add('active');

            // Update panels
            document.querySelectorAll('.range-panel').forEach(panel => {
                panel.classList.remove('active');
            });
            document.getElementById(`${tabType}-ranges`).classList.add('active');

            updateProcessButton();
        });
    });

    // Add range functionality
    document.querySelector('.add-range-btn').addEventListener('click', addNewRange);

    function addNewRange() {
        const rangesList = document.getElementById('ranges-list');
        const rangeCount = rangesList.children.length + 1;

        const rangeItem = document.createElement('div');
        rangeItem.className = 'range-item';
        rangeItem.innerHTML = `
            <label>Range ${rangeCount}:</label>
            <input type="number" class="range-start" placeholder="Start page" min="1" max="${totalPages}">
            <span>to</span>
            <input type="number" class="range-end" placeholder="End page" min="1" max="${totalPages}">
            <button class="remove-range">❌</button>
        `;

        rangesList.appendChild(rangeItem);

        // Add remove functionality
        rangeItem.querySelector('.remove-range').addEventListener('click', () => {
            rangeItem.remove();
            updateRangeLabels();
            updateProcessButton();
        });

        // Add change listeners
        rangeItem.querySelectorAll('input').forEach(input => {
            input.addEventListener('input', updateProcessButton);
        });

        updateProcessButton();
    }

    function updateRangeLabels() {
        const rangeItems = document.querySelectorAll('.range-item');
        rangeItems.forEach((item, index) => {
            item.querySelector('label').textContent = `Range ${index + 1}:`;
            const removeBtn = item.querySelector('.remove-range');
            removeBtn.style.display = rangeItems.length > 1 ? 'inline-block' : 'none';
        });
    }

    // Fixed range preview
    document.getElementById('pages-per-split').addEventListener('input', updateFixedRangePreview);

    function updateFixedRangePreview() {
        const pagesPerSplit = parseInt(document.getElementById('pages-per-split').value) || 3;
        const preview = document.getElementById('fixed-preview');

        if (totalPages === 0) return;

        const ranges = [];
        for (let i = 1; i <= totalPages; i += pagesPerSplit) {
            const end = Math.min(i + pagesPerSplit - 1, totalPages);
            ranges.push(`Pages ${i}-${end}`);
        }

        preview.innerHTML = `
            <p>This will create ${ranges.length} files:</p>
            <ul>${ranges.map(range => `<li>${range}</li>`).join('')}</ul>
        `;

        updateProcessButton();
    }

    // Select all/none functionality
    document.querySelector('.select-all-btn').addEventListener('click', () => {
        document.querySelectorAll('#selectable-pages input').forEach(input => {
            input.checked = true;
        });
        updateSelectedCount();
    });

    document.querySelector('.select-none-btn').addEventListener('click', () => {
        document.querySelectorAll('#selectable-pages input').forEach(input => {
            input.checked = false;
        });
        updateSelectedCount();
    });

    // Process step
    function showProcessStep() {
        processStep.style.display = 'block';
        updateProcessButton();
    }

    function updateProcessButton() {
        const processBtn = document.getElementById('process-btn');
        const selectedMethod = document.querySelector('input[name="split-method"]:checked');

        if (!selectedMethod || !documentUploaded) {
            processBtn.disabled = true;
            return;
        }

        if (selectedMethod.value === 'range') {
            const activeTab = document.querySelector('.range-tab.active').dataset.tab;
            if (activeTab === 'custom') {
                const ranges = getCustomRanges();
                processBtn.disabled = ranges.length === 0 || !validateRanges(ranges);
            } else {
                const pagesPerSplit = parseInt(document.getElementById('pages-per-split').value);
                processBtn.disabled = !pagesPerSplit || pagesPerSplit < 1;
            }
        } else if (selectedMethod.value === 'select') {
            const selectedPages = document.querySelectorAll('#selectable-pages input:checked').length;
            processBtn.disabled = selectedPages === 0;
        }
    }

    function getCustomRanges() {
        const ranges = [];
        document.querySelectorAll('.range-item').forEach(item => {
            const start = parseInt(item.querySelector('.range-start').value);
            const end = parseInt(item.querySelector('.range-end').value);
            if (start && end) {
                ranges.push({ start, end });
            }
        });
        return ranges;
    }

    function validateRanges(ranges) {
        return ranges.every(range =>
            range.start >= 1 &&
            range.end <= totalPages &&
            range.start <= range.end
        );
    }

    // Process button click
    document.getElementById('process-btn').addEventListener('click', processDocument);

    function processDocument() {
        const selectedMethod = document.querySelector('input[name="split-method"]:checked').value;
        const processStatus = document.getElementById('process-status');

        processStatus.innerHTML = '<p>Processing...</p>';
        document.getElementById('process-btn').disabled = true;

        let requestData = {
            method: selectedMethod
        };

        if (selectedMethod === 'range') {
            const activeTab = document.querySelector('.range-tab.active').dataset.tab;
            if (activeTab === 'custom') {
                requestData.ranges = getCustomRanges();
                requestData.merge_ranges = document.getElementById('merge-ranges').checked;
            } else {
                requestData.pages_per_split = parseInt(document.getElementById('pages-per-split').value);
            }
        } else if (selectedMethod === 'select') {
            const selectedPages = Array.from(document.querySelectorAll('#selectable-pages input:checked'))
                .map(input => parseInt(input.value));
            requestData.selected_pages = selectedPages;
            requestData.merge_selected = document.getElementById('merge-selected').checked;
        }

        debugLog('Processing with data:', requestData);

        fetch('/process_split', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify(requestData)
        })
            .then(response => response.json())
            .then(data => {
                debugLog('Process response:', data);

                if (data.success) {
                    processStatus.innerHTML = `<p class="success">✅ ${data.message}</p>`;

                    // Trigger download
                    const link = document.createElement('a');
                    link.href = data.download_url;
                    link.download = data.filename;
                    document.body.appendChild(link);
                    link.click();
                    document.body.removeChild(link);
                } else {
                    processStatus.innerHTML = `<p class="error">❌ ${data.error}</p>`;
                }
            })
            .catch(error => {
                debugLog('Process error:', error);
                processStatus.innerHTML = `<p class="error">❌ Network error during processing</p>`;
            })
            .finally(() => {
                document.getElementById('process-btn').disabled = false;
            });
    }

    // Initialize
    setupDropZone();
    updateRangeLabels();

    debugLog('Split Word interface initialized');
});