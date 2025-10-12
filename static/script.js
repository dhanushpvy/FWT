let uploadedFile = null;
let cleanedFileBlob = null;
let originalFileName = '';

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileType = document.getElementById('fileType');
const errorMessage = document.getElementById('errorMessage');

uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    
    const file = e.dataTransfer.files[0];
    if (file) {
        handleFileSelection(file);
    }
});

fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
        handleFileSelection(file);
    }
});

function handleFileSelection(file) {
    const allowedExtensions = ['.xlsx', '.docx', '.pdf'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
    
    if (!allowedExtensions.includes(fileExtension)) {
        showError('Please upload an Excel (.xlsx), Word (.docx), or PDF (.pdf) file');
        return;
    }

    uploadedFile = file;
    originalFileName = file.name;
    fileName.textContent = file.name;
    fileType.textContent = getFileTypeDescription(fileExtension);
    fileInfo.classList.remove('hidden');
}

function getFileTypeDescription(fileExtension) {
    const typeMap = {
        '.xlsx': 'Excel Workbook',
        '.docx': 'Word Document',
        '.pdf': 'PDF Document'
    };
    return typeMap[fileExtension] || 'Unknown File Type';
}

function showError(message) {
    errorMessage.textContent = message;
    errorMessage.classList.remove('hidden');
    setTimeout(() => {
        errorMessage.classList.add('hidden');
    }, 5000);
}

function showUploadSection() {
    hideAllSections();
    document.getElementById('upload-section').classList.add('active');
    updateSteps('step-upload');
}

function showConfigSection() {
    if (!uploadedFile) {
        showError('Please upload a file first');
        return;
    }
    hideAllSections();
    document.getElementById('configure-section').classList.add('active');
    updateSteps('step-configure');
    
    // Automatically start processing
    processFile();
}

function showDownloadSection() {
    hideAllSections();
    document.getElementById('download-section').classList.add('active');
    updateSteps('step-download');
}

function resetConfigureSection() {
    // Show processing status
    document.getElementById('processing-status').classList.remove('hidden');
    
    // Hide success message and download button
    document.getElementById('process-success').classList.add('hidden');
    document.getElementById('download-btn').classList.add('hidden');
}

function hideAllSections() {
    const sections = document.querySelectorAll('.section');
    sections.forEach(section => section.classList.remove('active'));
}

function updateSteps(activeStepId) {
    const steps = document.querySelectorAll('.step');
    steps.forEach(step => step.classList.remove('active'));
    document.getElementById(activeStepId).classList.add('active');
}


async function processFile() {
    if (!uploadedFile) {
        showError('Please upload a file first');
        return;
    }

    // Show processing status
    document.getElementById('processing-status').classList.remove('hidden');
    document.getElementById('process-success').classList.add('hidden');
    document.getElementById('download-btn').classList.add('hidden');

    try {
        const formData = new FormData();
        formData.append('file', uploadedFile);

        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        if (response.ok) {
            // Get the cleaned file as blob
            cleanedFileBlob = await response.blob();
            
            setTimeout(() => {
                showDownloadButton();
            }, 1000);
        } else {
            const errorData = await response.json();
            showError(errorData.error || 'Processing failed');
            resetConfigureSection();
        }
    } catch (error) {
        showError('Error processing file: ' + error.message);
        resetConfigureSection();
    }
}

function showDownloadButton() {
    // Hide processing status and show success message with download button
    document.getElementById('processing-status').classList.add('hidden');
    document.getElementById('process-success').classList.remove('hidden');
    document.getElementById('download-btn').classList.remove('hidden');
}