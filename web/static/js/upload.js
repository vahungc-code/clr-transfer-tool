document.addEventListener('DOMContentLoaded', function () {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('clr-file');
    const fileInfo = document.getElementById('file-info');
    const fileName = document.getElementById('file-name');
    const uploadBtn = document.getElementById('upload-btn');
    const uploadForm = document.getElementById('upload-form');

    if (!dropZone || !fileInput) return;

    // Click to browse
    dropZone.addEventListener('click', function () {
        fileInput.click();
    });

    // File selected via input
    fileInput.addEventListener('change', function () {
        if (fileInput.files.length > 0) {
            showFile(fileInput.files[0]);
        }
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
            const file = e.dataTransfer.files[0];
            const ext = file.name.split('.').pop().toLowerCase();
            if (ext === 'xlsx' || ext === 'xlsm') {
                fileInput.files = e.dataTransfer.files;
                showFile(file);
            } else {
                alert('Only .xlsx and .xlsm files are supported.');
            }
        }
    });

    // Show processing overlay on submit
    if (uploadForm) {
        uploadForm.addEventListener('submit', function () {
            if (uploadBtn && !uploadBtn.disabled) {
                uploadBtn.disabled = true;
                uploadBtn.innerHTML = '<span class="spinner" style="width:18px;height:18px;border-width:2px;display:inline-block;vertical-align:middle;"></span> Analyzing...';
            }
        });
    }

    function showFile(file) {
        fileName.textContent = file.name + ' (' + formatSize(file.size) + ')';
        fileInfo.classList.remove('d-none');
        uploadBtn.disabled = false;
    }

    function formatSize(bytes) {
        if (bytes < 1024) return bytes + ' B';
        if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
        return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
    }
});
