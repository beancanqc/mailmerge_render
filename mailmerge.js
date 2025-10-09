// Mail Merge JavaScript functionality

document.addEventListener('DOMContentLoaded', function () {
    // Get DOM elements
    const templateDrop = document.getElementById('template-drop');
    const dataDrop = document.getElementById('data-drop');
    const templateInput = document.getElementById('template-input');
    const dataInput = document.getElementById('data-input');
    const mergeBtn = document.querySelector('.merge-btn');
    const formatOptions = document.querySelectorAll('input[name="output-format"]');

    let templateUploaded = false;
    let dataUploaded = false;
    let templateProcessing = false;
    let dataProcessing = false;

    // Debug function
    function debugLog(message) {
        console.log('[MailMerge]', message);
    }

    // Check if elements exist
    if (!templateDrop || !dataDrop) {
        console.error('Drop zones not found in DOM');
        return;
    }

    // Disable PDF options since we only support Word
    const pdfOptions = document.querySelectorAll('input[value*="pdf"]');
    pdfOptions.forEach(option => {
        option.disabled = true;
        const label = option.closest('.format-option');
        if (label) {
            label.style.opacity = '0.5';
            label.style.cursor = 'not-allowed';
            const text = label.querySelector('.format-text');
            if (text) {
                text.innerHTML += ' <small>(Not available)</small>';
            }
        }
    });

    // Set default to single-word
    const singleWordOption = document.getElementById('single-word');
    if (singleWordOption) {
        singleWordOption.checked = true;
    }

    // Check status on page load
    function checkUploadStatus() {
        fetch('/check_status')
            .then(response => response.json())
            .then(data => {
                debugLog('Current status:', data);
                templateUploaded = data.template_loaded;
                dataUploaded = data.data_loaded;
                updateMergeButton();
            })
            .catch(error => {
                debugLog('Status check error:', error);
                // Don't update status if there's an error
            });
    }

    // File drop functionality
    function setupDropZone(dropZone, fileInput, fileType) {
        if (!dropZone || !fileInput) {
            console.error(`Missing elements for ${fileType} drop zone`);
            return;
        }

        dropZone.addEventListener('click', () => fileInput.click());

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
                fileInput.files = files;
                handleFileUpload(files[0], fileType, dropZone);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFileUpload(e.target.files[0], fileType, dropZone);
            }
        });
    }

    // Handle file upload
    function handleFileUpload(file, fileType, dropZone) {
        debugLog(`Uploading ${fileType} file:`, file.name);

        const formData = new FormData();
        formData.append('file', file);

        const endpoint = fileType === 'template' ? '/upload_template' : '/upload_data';

        // Set processing state
        if (fileType === 'template') {
            templateProcessing = true;
            templateUploaded = false;
        } else {
            dataProcessing = true;
            dataUploaded = false;
        }

        updateMergeButton();

        // Show loading state
        dropZone.innerHTML = `
            <div class="drop-icon">⏳</div>
            <p class="drop-text">Uploading ${file.name}...</p>
        `;

        fetch(endpoint, {
            method: 'POST',
            body: formData
        })
            .then(response => {
                debugLog(`Upload response status: ${response.status}`);
                return response.json();
            })
            .then(data => {
                debugLog(`Upload response for ${fileType}:`, data);

                if (data.success !== false && !data.error) {
                    // Success
                    dropZone.innerHTML = `
                    <div class="drop-icon">✅</div>
                    <p class="drop-text">${file.name}</p>
                    <p class="drop-subtext">Uploaded successfully</p>
                `;
                    dropZone.style.borderColor = '#48bb78';
                    dropZone.style.backgroundColor = '#f0fff4';

                    if (fileType === 'template') {
                        templateUploaded = true;
                        templateProcessing = false;
                    } else {
                        dataUploaded = true;
                        dataProcessing = false;
                        // Show data preview if available
                        if (data.preview && data.total_rows) {
                            showDataPreview(data.preview, data.total_rows);
                        }
                    }

                    updateMergeButton();
                } else {
                    // Error
                    const errorMsg = data.error || 'Unknown error';
                    debugLog(`Upload error for ${fileType}:`, errorMsg);

                    dropZone.innerHTML = `
                    <div class="drop-icon">❌</div>
                    <p class="drop-text">Upload failed</p>
                    <p class="drop-subtext">${errorMsg}</p>
                `;
                    dropZone.style.borderColor = '#e53e3e';
                    dropZone.style.backgroundColor = '#fed7d7';

                    // Reset state on error
                    if (fileType === 'template') {
                        templateUploaded = false;
                        templateProcessing = false;
                    } else {
                        dataUploaded = false;
                        dataProcessing = false;
                    }

                    updateMergeButton();
                }
            })
            .catch(error => {
                debugLog(`Network error for ${fileType}:`, error);
                dropZone.innerHTML = `
                <div class="drop-icon">❌</div>
                <p class="drop-text">Upload failed</p>
                <p class="drop-subtext">Network error</p>
            `;
                dropZone.style.borderColor = '#e53e3e';
                dropZone.style.backgroundColor = '#fed7d7';

                // Reset state on error
                if (fileType === 'template') {
                    templateUploaded = false;
                    templateProcessing = false;
                } else {
                    dataUploaded = false;
                    dataProcessing = false;
                }

                updateMergeButton();
            });
    }

    // Show data preview
    function showDataPreview(preview, totalRows) {
        let previewHtml = `<div style="margin-top: 10px; font-size: 12px; color: #666;">`;
        previewHtml += `<strong>Preview (${totalRows} total rows):</strong><br>`;

        if (preview.length > 0) {
            const keys = Object.keys(preview[0]);
            previewHtml += `<strong>Columns:</strong> ${keys.join(', ')}`;
        }

        previewHtml += `</div>`;

        const currentText = dataDrop.querySelector('.drop-subtext');
        if (currentText) {
            currentText.innerHTML = 'Uploaded successfully' + previewHtml;
        }
    }

    // Update merge button state
    function updateMergeButton() {
        const formatSelected = document.querySelector('input[name="output-format"]:checked');

        debugLog('Update button state:', {
            templateUploaded,
            dataUploaded,
            templateProcessing,
            dataProcessing,
            formatSelected: formatSelected ? formatSelected.value : 'none'
        });

        if (mergeBtn) {
            const allReady = templateUploaded && dataUploaded && formatSelected && !templateProcessing && !dataProcessing;

            if (allReady) {
                mergeBtn.disabled = false;
                mergeBtn.textContent = 'Start Mail Merge';
                mergeBtn.style.opacity = '1';
                mergeBtn.style.cursor = 'pointer';
            } else {
                mergeBtn.disabled = true;
                if (templateProcessing || dataProcessing) {
                    mergeBtn.textContent = 'Processing files...';
                } else if (!templateUploaded && !dataUploaded) {
                    mergeBtn.textContent = 'Upload files first';
                } else if (!templateUploaded) {
                    mergeBtn.textContent = 'Upload template first';
                } else if (!dataUploaded) {
                    mergeBtn.textContent = 'Upload data file first';
                } else if (!formatSelected) {
                    mergeBtn.textContent = 'Select output format';
                } else {
                    mergeBtn.textContent = 'Upload files first';
                }
                mergeBtn.style.opacity = '0.6';
                mergeBtn.style.cursor = 'not-allowed';
            }
        }
    }

    // Format selection change
    formatOptions.forEach(option => {
        option.addEventListener('change', updateMergeButton);
    });

    // Merge button click
    if (mergeBtn) {
        mergeBtn.addEventListener('click', function () {
            const selectedFormat = document.querySelector('input[name="output-format"]:checked');

            debugLog('Merge button clicked:', {
                templateUploaded,
                dataUploaded,
                selectedFormat: selectedFormat ? selectedFormat.value : 'none'
            });

            if (!selectedFormat) {
                alert('Please select an output format');
                return;
            }

            if (!templateUploaded || !dataUploaded) {
                alert('Files are still being processed. Please wait a few seconds and try again.');
                return;
            }

            // Disable button and show processing
            mergeBtn.disabled = true;
            mergeBtn.textContent = 'Processing...';

            // Send merge request
            fetch('/process_merge', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    format: selectedFormat.value
                })
            })
                .then(response => {
                    debugLog('Process response status:', response.status);
                    return response.json();
                })
                .then(data => {
                    debugLog('Process response:', data);

                    if (data.success) {
                        // Success - trigger download
                        const link = document.createElement('a');
                        link.href = data.download_url;
                        link.download = data.filename;
                        document.body.appendChild(link);
                        link.click();
                        document.body.removeChild(link);

                        // Show success message
                        alert(data.message);
                    } else {
                        // Error
                        alert(data.error || 'Processing failed');
                    }
                })
                .catch(error => {
                    debugLog('Processing error:', error);
                    alert('Network error during processing');
                })
                .finally(() => {
                    // Re-enable button
                    mergeBtn.disabled = false;
                    updateMergeButton(); // Update based on current state
                });
        });
    }

    // Setup drop zones
    setupDropZone(templateDrop, templateInput, 'template');
    setupDropZone(dataDrop, dataInput, 'data');

    // Initial setup
    updateMergeButton();

    debugLog('Mail merge interface initialized');
});