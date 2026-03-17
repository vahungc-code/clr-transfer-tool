document.addEventListener('DOMContentLoaded', function () {
    const dropZone = document.getElementById('template-drop-zone');
    const fileInput = document.getElementById('template-files');
    const fileList = document.getElementById('template-file-list');
    const transferBtn = document.getElementById('transfer-btn');
    const templateForm = document.getElementById('template-form');

    if (!dropZone || !fileInput) return;

    // Click to browse
    dropZone.addEventListener('click', function (e) {
        if (e.target.closest('.file-remove')) return;
        fileInput.click();
    });

    // File selected via input
    fileInput.addEventListener('change', function () {
        updateFileList();
    });

    // Drag and drop
    dropZone.addEventListener('dragover', function (e) {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', function () {
        dropZone.classList.remove('drag-over');
    });

    dropZone.addEventListener('drop', function (e) {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            // Filter to only valid extensions
            const validFiles = Array.from(e.dataTransfer.files).filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return ext === 'xlsx' || ext === 'xlsm';
            });

            if (validFiles.length === 0) {
                alert('Only .xlsx and .xlsm files are supported.');
                return;
            }

            // Create a new DataTransfer to set files
            const dt = new DataTransfer();
            // Add existing files
            if (fileInput.files) {
                for (const f of fileInput.files) {
                    dt.items.add(f);
                }
            }
            // Add new files
            for (const f of validFiles) {
                dt.items.add(f);
            }
            fileInput.files = dt.files;
            updateFileList();
        }
    });

    // Show processing state on submit
    if (templateForm) {
        templateForm.addEventListener('submit', function () {
            if (transferBtn && !transferBtn.disabled) {
                transferBtn.disabled = true;
                transferBtn.innerHTML = '<span class="spinner" style="width:18px;height:18px;border-width:2px;display:inline-block;vertical-align:middle;"></span> Transferring data...';
            }
        });
    }

    function updateFileList() {
        if (!fileInput.files || fileInput.files.length === 0) {
            fileList.classList.add('d-none');
            transferBtn.disabled = true;
            return;
        }

        fileList.classList.remove('d-none');
        fileList.innerHTML = '';

        for (let i = 0; i < fileInput.files.length; i++) {
            const file = fileInput.files[i];
            const item = document.createElement('div');
            item.className = 'template-file-item';
            item.innerHTML = `
                <div class="d-flex align-items-center gap-2">
                    <i class="bi bi-file-earmark-spreadsheet text-accent"></i>
                    <span class="file-name">${file.name}</span>
                </div>
                <span class="file-size">${formatSize(file.size)}</span>
            `;
            fileList.appendChild(item);
        }

        transferBtn.disabled = false;
    }

    function formatSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }
});
