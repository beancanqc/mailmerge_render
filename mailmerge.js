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

        const btnText = mergeBtn ? mergeBtn.querySelector('.merge-btn-text') : null;

        if (mergeBtn && btnText) {
            if (templateProcessing || dataProcessing) {
                mergeBtn.disabled = true;
                mergeBtn.classList.add('processing');
                btnText.textContent = 'Processing files...';
                mergeBtn.style.opacity = '0.6';
                mergeBtn.style.cursor = 'not-allowed';
            } else if (!formatSelected) {
                mergeBtn.disabled = true;
                mergeBtn.classList.remove('processing');
                btnText.textContent = 'Select output format';
                mergeBtn.style.opacity = '0.6';
                mergeBtn.style.cursor = 'not-allowed';
            } else {
                // Enable button - let the click handler check server status
                mergeBtn.disabled = false;
                mergeBtn.classList.remove('processing');
                btnText.textContent = 'Start Mail Merge';
                mergeBtn.style.opacity = '1';
                mergeBtn.style.cursor = 'pointer';
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

            // Check server status before proceeding
            fetch('/check_status')
                .then(response => response.json())
                .then(statusData => {
                    debugLog('Server status check:', statusData);

                    if (!statusData.template_loaded || !statusData.data_loaded) {
                        alert('Files are still being processed. Please wait a few seconds and try again.');
                        return;
                    }

                    // Disable button and show processing with spinner
                    mergeBtn.disabled = true;
                    mergeBtn.classList.add('processing');
                    const btnText = mergeBtn.querySelector('.merge-btn-text');
                    if (btnText) {
                        btnText.textContent = 'Processing...';
                    }

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
                            // Re-enable button and remove spinner
                            mergeBtn.disabled = false;
                            mergeBtn.classList.remove('processing');
                            updateMergeButton(); // Update based on current state
                        });
                })
                .catch(error => {
                    debugLog('Status check error:', error);
                    alert('Unable to check file status. Please try again.');
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