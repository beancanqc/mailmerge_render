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

    // Disable PDF options since we only support Word
    const pdfOptions = document.querySelectorAll('input[value*="pdf"]');
    pdfOptions.forEach(option => {
        option.disabled = true;
        const label = option.closest('.format-option');
        label.style.opacity = '0.5';
        label.style.cursor = 'not-allowed';
        const text = label.querySelector('.format-text');
        if (text) {
            text.innerHTML += ' <small>(Not available)</small>';
        }
    });

    // Set default to single-word
    document.getElementById('single-word').checked = true;

    // File drop functionality
    function setupDropZone(dropZone, fileInput, fileType) {
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
        const formData = new FormData();
        formData.append('file', file);

        const endpoint = fileType === 'template' ? '/upload_template' : '/upload_data';

        // Show loading state
        dropZone.innerHTML = `
            <div class="drop-icon">⏳</div>
            <p class="drop-text">Uploading...</p>
        `;

        fetch(endpoint, {
            method: 'POST',
            body: formData
        })
            .then(response => response.json())
            .then(data => {
                if (data.success || !data.error) {
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
                    } else {
                        dataUploaded = true;
                        // Show data preview if available
                        if (data.preview && data.total_rows) {
                            showDataPreview(data.preview, data.total_rows);
                        }
                    }

                    updateMergeButton();
                } else {
                    // Error
                    dropZone.innerHTML = `
                    <div class="drop-icon">❌</div>
                    <p class="drop-text">Upload failed</p>
                    <p class="drop-subtext">${data.error || 'Unknown error'}</p>
                `;
                    dropZone.style.borderColor = '#e53e3e';
                    dropZone.style.backgroundColor = '#fed7d7';
                }
            })
            .catch(error => {
                console.error('Upload error:', error);
                dropZone.innerHTML = `
                <div class="drop-icon">❌</div>
                <p class="drop-text">Upload failed</p>
                <p class="drop-subtext">Network error</p>
            `;
                dropZone.style.borderColor = '#e53e3e';
                dropZone.style.backgroundColor = '#fed7d7';
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

        if (templateUploaded && dataUploaded && formatSelected) {
            mergeBtn.disabled = false;
        } else {
            mergeBtn.disabled = true;
        }
    }

    // Format selection change
    formatOptions.forEach(option => {
        option.addEventListener('change', updateMergeButton);
    });

    // Merge button click
    mergeBtn.addEventListener('click', function () {
        const selectedFormat = document.querySelector('input[name="output-format"]:checked');

        if (!selectedFormat) {
            alert('Please select an output format');
            return;
        }

        if (!templateUploaded || !dataUploaded) {
            alert('Please upload both template and data files');
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
            .then(response => response.json())
            .then(data => {
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
                console.error('Processing error:', error);
                alert('Network error during processing');
            })
            .finally(() => {
                // Re-enable button
                mergeBtn.disabled = false;
                mergeBtn.textContent = 'Start Mail Merge';
                updateMergeButton(); // Update based on current state
            });
    });

    // Setup drop zones
    setupDropZone(templateDrop, templateInput, 'template');
    setupDropZone(dataDrop, dataInput, 'data');

    // Initial button state
    updateMergeButton();
});