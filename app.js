// Word Document Header/Footer Updater Application
class WordDocumentUpdater {
    constructor() {
        this.templateFile = null;
        this.templateZip = null;
        this.targetFiles = [];
        this.processedFiles = [];
        this.isProcessing = false;
        this.debugMode = false;
        
        this.maxFileSize = 50 * 1024 * 1024; // 50MB
        this.supportedFormat = '.docx';
        
        this.xmlNamespaces = {
            w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        };
        
        this.docxStructure = {
            mainDocument: 'word/document.xml',
            headers: ['word/header1.xml', 'word/header2.xml', 'word/header3.xml'],
            footers: ['word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'],
            relationships: 'word/_rels/document.xml.rels',
            contentTypes: '[Content_Types].xml',
            styleFiles: [
                'word/styles.xml',
                'word/theme/theme1.xml',
                'word/fontTable.xml',
                'word/settings.xml',
                'word/numbering.xml'
            ]
        };
        
        this.errorMessages = {
            invalidFormat: 'Please select a valid .docx file',
            fileTooLarge: 'File size exceeds 50MB limit',
            protectedDocument: 'Document appears to be protected. Please disable Protected View in Word',
            corruptedFile: 'File appears to be corrupted or not a valid Word document',
            noHeadersFooters: 'No headers or footers found in template document',
            processingError: 'An error occurred while processing the document'
        };
        
        this.init();
    }

    init() {
        this.checkBrowserCompatibility();
        this.setupEventListeners();
        this.updateProcessButton();
    }

    checkBrowserCompatibility() {
        const hasRequiredAPIs = 
            typeof FileReader !== 'undefined' &&
            typeof DOMParser !== 'undefined' &&
            typeof XMLSerializer !== 'undefined' &&
            typeof JSZip !== 'undefined' &&
            typeof saveAs !== 'undefined';

        if (!hasRequiredAPIs) {
            document.getElementById('compatibility-warning').classList.remove('hidden');
        }
    }

    setupEventListeners() {
        // Template upload
        const templateZone = document.getElementById('template-upload-zone');
        const templateInput = document.getElementById('template-input');
        const templateBrowse = document.getElementById('template-browse');
        const removeTemplate = document.getElementById('remove-template');

        templateZone.addEventListener('dragover', this.handleDragOver.bind(this));
        templateZone.addEventListener('dragleave', this.handleDragLeave.bind(this));
        templateZone.addEventListener('drop', (e) => this.handleTemplateDrop(e));
        templateZone.addEventListener('click', () => templateInput.click());
        templateBrowse.addEventListener('click', (e) => {
            e.stopPropagation();
            templateInput.click();
        });
        templateInput.addEventListener('change', (e) => this.handleTemplateSelect(e));
        removeTemplate.addEventListener('click', () => this.removeTemplate());

        // Target files upload
        const targetZone = document.getElementById('target-upload-zone');
        const targetInput = document.getElementById('target-input');
        const targetBrowse = document.getElementById('target-browse');
        const clearTargets = document.getElementById('clear-targets');

        targetZone.addEventListener('dragover', this.handleDragOver.bind(this));
        targetZone.addEventListener('dragleave', this.handleDragLeave.bind(this));
        targetZone.addEventListener('drop', (e) => this.handleTargetDrop(e));
        targetZone.addEventListener('click', () => targetInput.click());
        targetBrowse.addEventListener('click', (e) => {
            e.stopPropagation();
            targetInput.click();
        });
        targetInput.addEventListener('change', (e) => this.handleTargetSelect(e));
        clearTargets.addEventListener('click', () => this.clearTargetFiles());

        // Processing
        const processBtn = document.getElementById('process-btn');
        const debugMode = document.getElementById('debug-mode');
        
        processBtn.addEventListener('click', () => this.processDocuments());
        debugMode.addEventListener('change', (e) => {
            this.debugMode = e.target.checked;
        });

        // Download
        const downloadAll = document.getElementById('download-all');
        downloadAll.addEventListener('click', () => this.downloadAllFiles());

        // Error handling
        const clearErrors = document.getElementById('clear-errors');
        clearErrors.addEventListener('click', () => this.clearErrors());

        // Preview Modal
        const previewModalClose = document.getElementById('preview-modal-close');
        previewModalClose.addEventListener('click', () => this.closePreviewModal());
    }

    handleDragOver(e) {
        e.preventDefault();
        e.currentTarget.classList.add('drag-over');
    }

    handleDragLeave(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('drag-over');
    }

    async handleTemplateDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('drag-over');
        
        const files = Array.from(e.dataTransfer.files);
        if (files.length > 0) {
            await this.setTemplateFile(files[0]);
        }
    }

    async handleTemplateSelect(e) {
        const file = e.target.files[0];
        if (file) {
            await this.setTemplateFile(file);
        }
    }

    async handleTargetDrop(e) {
        e.preventDefault();
        e.currentTarget.classList.remove('drag-over');
        
        const files = Array.from(e.dataTransfer.files);
        await this.addTargetFiles(files);
    }

    async handleTargetSelect(e) {
        const files = Array.from(e.target.files);
        await this.addTargetFiles(files);
    }

    async setTemplateFile(file) {
        try {
            this.validateFile(file);
            
            this.templateFile = file;
            this.templateZip = await JSZip.loadAsync(file);
            
            await this.extractAndPreviewTemplate();
            this.showTemplateStatus();
            this.updateProcessButton();
            
        } catch (error) {
            this.logError('Template Error', error.message);
            this.templateFile = null;
            this.templateZip = null;
        }
    }

    async addTargetFiles(files) {
        for (const file of files) {
            try {
                this.validateFile(file);
                
                if (!this.targetFiles.find(f => f.name === file.name)) {
                    this.targetFiles.push(file);
                }
            } catch (error) {
                this.logError(`Target File Error (${file.name})`, error.message);
            }
        }
        
        this.showTargetFiles();
        this.updateProcessButton();
    }

    validateFile(file) {
        if (!file.name.toLowerCase().endsWith(this.supportedFormat)) {
            throw new Error(this.errorMessages.invalidFormat);
        }
        
        if (file.size > this.maxFileSize) {
            throw new Error(this.errorMessages.fileTooLarge);
        }
        
        if (file.size === 0) {
            throw new Error(this.errorMessages.corruptedFile);
        }
    }

    async extractAndPreviewTemplate() {
        try {
            const headerFooterContent = [];
            
            // Extract headers
            for (const headerPath of this.docxStructure.headers) {
                const headerFile = this.templateZip.file(headerPath);
                if (headerFile) {
                    const content = await headerFile.async('text');
                    headerFooterContent.push({
                        type: 'Header',
                        path: headerPath,
                        content: this.formatXmlForPreview(content)
                    });
                }
            }
            
            // Extract footers
            for (const footerPath of this.docxStructure.footers) {
                const footerFile = this.templateZip.file(footerPath);
                if (footerFile) {
                    const content = await footerFile.async('text');
                    headerFooterContent.push({
                        type: 'Footer',
                        path: footerPath,
                        content: this.formatXmlForPreview(content)
                    });
                }
            }
            
            if (headerFooterContent.length === 0) {
                throw new Error(this.errorMessages.noHeadersFooters);
            }
            
            this.showHeaderFooterPreview(headerFooterContent);
            
        } catch (error) {
            throw new Error(`Failed to extract headers/footers: ${error.message}`);
        }
    }

    formatXmlForPreview(xmlContent) {
        try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(xmlContent, 'text/xml');
            
            // Extract text content for preview
            const textNodes = doc.querySelectorAll('w\\:t, t');
            const textContent = Array.from(textNodes).map(node => node.textContent).join(' ');
            
            return textContent.trim() || 'No visible text content';
        } catch (error) {
            return 'Preview not available';
        }
    }

    showTemplateStatus() {
        const statusSection = document.getElementById('template-status');
        const filename = document.getElementById('template-filename');
        
        filename.textContent = this.templateFile.name;
        statusSection.classList.remove('hidden');
    }

    showHeaderFooterPreview(content) {
        const previewSection = document.getElementById('template-preview');
        const previewContent = document.getElementById('preview-content');
        
        let html = '';
        content.forEach(item => {
            html += `<div><strong>${item.type} (${item.path}):</strong><br>${item.content}</div><br>`;
        });
        
        previewContent.innerHTML = html;
        previewSection.classList.remove('hidden');
    }

    showTargetFiles() {
        const filesList = document.getElementById('target-files-list');
        const filesContainer = document.getElementById('target-files');
        
        if (this.targetFiles.length === 0) {
            filesList.classList.add('hidden');
            return;
        }
        
        let html = '';
        this.targetFiles.forEach((file, index) => {
            const fileSize = this.formatFileSize(file.size);
            html += `
                <div class="file-item">
                    <div class="file-info">
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${fileSize}</div>
                    </div>
                    <div class="file-actions">
                        <button type="button" class="btn btn--sm btn--secondary" onclick="app.removeTargetFile(${index})">Remove</button>
                    </div>
                </div>
            `;
        });
        
        filesContainer.innerHTML = html;
        filesList.classList.remove('hidden');
    }

    removeTemplate() {
        this.templateFile = null;
        this.templateZip = null;
        document.getElementById('template-status').classList.add('hidden');
        document.getElementById('template-preview').classList.add('hidden');
        document.getElementById('template-input').value = '';
        this.updateProcessButton();
    }

    removeTargetFile(index) {
        this.targetFiles.splice(index, 1);
        this.showTargetFiles();
        this.updateProcessButton();
    }

    clearTargetFiles() {
        this.targetFiles = [];
        document.getElementById('target-input').value = '';
        this.showTargetFiles();
        this.updateProcessButton();
    }

    updateProcessButton() {
        const processBtn = document.getElementById('process-btn');
        const canProcess = this.templateFile && this.targetFiles.length > 0 && !this.isProcessing;
        
        processBtn.disabled = !canProcess;
        processBtn.textContent = this.isProcessing ? 'Processing...' : 'Process Documents';
        
        if (this.isProcessing) {
            processBtn.classList.add('btn--loading');
        } else {
            processBtn.classList.remove('btn--loading');
        }
    }

    async processDocuments() {
        if (this.isProcessing) return;
        
        this.isProcessing = true;
        this.processedFiles = [];
        this.updateProcessButton();
        this.showProcessingStatus();
        
        try {
            this.log('Starting document processing...', 'info');
            
            for (let i = 0; i < this.targetFiles.length; i++) {
                const file = this.targetFiles[i];
                const progress = ((i + 1) / this.targetFiles.length) * 100;
                
                this.updateProgress(progress, `Processing ${file.name}...`);
                this.log(`Processing: ${file.name}`, 'info');
                
                try {
                    const processedFile = await this.processDocument(file);
                    this.processedFiles.push(processedFile);
                    this.log(`✓ Successfully processed: ${file.name}`, 'success');
                } catch (error) {
                    this.log(`✗ Failed to process ${file.name}: ${error.message}`, 'error');
                    this.logError(`Processing Error (${file.name})`, error.message);
                }
            }
            
            this.updateProgress(100, 'Processing complete!');
            this.log(`Processing complete. ${this.processedFiles.length}/${this.targetFiles.length} files processed successfully.`, 'success');
            
            this.showDownloadSection();
            
        } catch (error) {
            this.log(`Processing failed: ${error.message}`, 'error');
            this.logError('Processing Error', error.message);
        } finally {
            this.isProcessing = false;
            this.updateProcessButton();
        }
    }

    async processDocument(file) {
        try {
            // Load target document
            const targetZip = await JSZip.loadAsync(file);
            
            // IMPROVED LOGIC: Extract only the body content from target and merge it 
            // with the template's document structure, preserving headers/footers
            
            // 1. Get the main document content from both files
            const documentXmlPath = this.docxStructure.mainDocument;
            const targetDocumentXml = await targetZip.file(documentXmlPath)?.async('string');
            
            if (!targetDocumentXml) {
                throw new Error(`Could not find ${documentXmlPath} in the target file: ${file.name}`);
            }

            // 2. Load a fresh copy of the template zip to work with
            const templateBlob = await this.templateFile.arrayBuffer();
            const newTemplateZip = await JSZip.loadAsync(templateBlob);
            
            // 3. Get the template's document.xml to preserve its structure
            const templateDocumentXml = await newTemplateZip.file(documentXmlPath)?.async('string');
            
            if (!templateDocumentXml) {
                throw new Error(`Could not find ${documentXmlPath} in template file`);
            }

            // 4. NEW: Intelligently merge styles.xml, ensuring header/footer styles from template are present
            this.log('Merging styles.xml for exact header/footer formatting...', 'info');
            try {
                const templateStylesXml = await newTemplateZip.file('word/styles.xml')?.async('string');
                const targetStylesXml = await targetZip.file('word/styles.xml')?.async('string');

                // Read all header/footer XML from template
                const headerXmls = [];
                for (const headerPath of this.docxStructure.headers) {
                    const f = newTemplateZip.file(headerPath);
                    if (f) headerXmls.push(await f.async('string'));
                }
                const footerXmls = [];
                for (const footerPath of this.docxStructure.footers) {
                    const f = newTemplateZip.file(footerPath);
                    if (f) footerXmls.push(await f.async('string'));
                }

                if (templateStylesXml && targetStylesXml) {
                    const mergedStylesXml = this.mergeStylesXml(templateStylesXml, targetStylesXml, headerXmls, footerXmls);
                    newTemplateZip.file('word/styles.xml', mergedStylesXml);
                } else {
                    this.log('Could not find styles.xml in both template and target. Skipping merge.', 'warn');
                }
            } catch (error) {
                this.log(`Error during style merge: ${error.message}`, 'error');
            }

            // 5. Copy other critical formatting files from target to the new zip
            this.log(`Copying other formatting files from ${file.name}...`, 'info');
            const otherStyleFiles = this.docxStructure.styleFiles.filter(p => p !== 'word/styles.xml');
            for (const stylePath of otherStyleFiles) {
                const styleFile = targetZip.file(stylePath);
                if (styleFile) {
                    const content = await styleFile.async('blob');
                    newTemplateZip.file(stylePath, content);
                } else if (this.debugMode) {
                    this.log(`Style file not found in target, skipping: ${stylePath}`, 'info');
                }
            }

            // 6. Merge the document body content
            const mergedDocumentXml = this.mergeDocumentContent(templateDocumentXml, targetDocumentXml);
            
            // 7. Update the document.xml in the template with merged content
            newTemplateZip.file(documentXmlPath, mergedDocumentXml);
            
            // 8. Generate processed file from the modified template
            const processedBlob = await newTemplateZip.generateAsync({
                type: 'blob',
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            });
            
            const processedFileName = this.generateProcessedFileName(file.name);
            
            return {
                name: processedFileName,
                originalName: file.name,
                blob: processedBlob,
                size: processedBlob.size
            };
            
        } catch (error) {
            throw new Error(`Failed to process document: ${error.message}`);
        }
    }

    // Helper: Get all styleIds referenced in a given XML string (for header/footer)
    getReferencedStyleIds(xmlString) {
        const ids = new Set();
        const regex = /w:([p|r|tbl]Style)[^>]*w:val="([^"]+)"/g;
        let match;
        while ((match = regex.exec(xmlString)) !== null) {
            ids.add(match[2]);
        }
        return ids;
    }

    // Improved mergeStylesXml: ensure all header/footer styleIds from template are present and unmodified
    mergeStylesXml(templateXml, targetXml, headerXmls, footerXmls) {
        try {
            const parser = new DOMParser();
            const serializer = new XMLSerializer();
            const templateDoc = parser.parseFromString(templateXml, "text/xml");
            const targetDoc = parser.parseFromString(targetXml, "text/xml");
            const templateStyles = templateDoc.querySelector('styles');
            const targetStyles = targetDoc.querySelector('styles');
            if (!templateStyles || !targetStyles) throw new Error('Missing <w:styles>');

            // 1. Collect all styleIds referenced in template header/footer XML
            let referencedIds = new Set();
            for (const xml of [...headerXmls, ...footerXmls]) {
                for (const id of this.getReferencedStyleIds(xml)) referencedIds.add(id);
            }

            // 2. Build a map of styleId -> styleElement for both
            const targetStyleMap = new Map();
            targetStyles.querySelectorAll('style').forEach(s => {
                const id = s.getAttribute('w:styleId');
                if (id) targetStyleMap.set(id, s);
            });
            const templateStyleMap = new Map();
            templateStyles.querySelectorAll('style').forEach(s => {
                const id = s.getAttribute('w:styleId');
                if (id) templateStyleMap.set(id, s);
            });

            // 3. For each referencedId, if not in target or different, add template's style
            let appended = 0;
            referencedIds.forEach(id => {
                const templateStyle = templateStyleMap.get(id);
                const targetStyle = targetStyleMap.get(id);
                if (!targetStyle && templateStyle) {
                    // Not present in target, append
                    const imported = targetDoc.importNode(templateStyle, true);
                    targetStyles.appendChild(imported);
                    appended++;
                } else if (targetStyle && templateStyle && serializer.serializeToString(targetStyle) !== serializer.serializeToString(templateStyle)) {
                    // Present but different, replace
                    targetStyles.replaceChild(targetDoc.importNode(templateStyle, true), targetStyle);
                    appended++;
                }
            });
            if (this.debugMode) this.log(`Ensured ${appended} header/footer styles from template are present in merged styles.xml.`, 'info');
            return serializer.serializeToString(targetDoc);
        } catch (error) {
            this.log(`Error merging styles.xml: ${error.message}`, 'error');
            throw error;
        }
    }

    mergeDocumentContent(templateXml, targetXml) {
        try {
            const parser = new DOMParser();
            const serializer = new XMLSerializer();
            
            // Parse both documents
            const templateDoc = parser.parseFromString(templateXml, 'text/xml');
            const targetDoc = parser.parseFromString(targetXml, 'text/xml');
            
            // Check for parsing errors
            if (templateDoc.documentElement.nodeName === 'parsererror' || 
                targetDoc.documentElement.nodeName === 'parsererror') {
                throw new Error('Failed to parse XML documents');
            }
            
            // Find the body element in both documents
            const templateBody = templateDoc.querySelector('body');
            const targetBody = targetDoc.querySelector('body');
            
            if (!templateBody || !targetBody) {
                throw new Error('Could not find body elements in documents');
            }
            
            // Extract section properties from template (contains header/footer references)
            const templateSectPr = templateBody.querySelector('sectPr');
            
            // Clear the template body content but keep its structure
            while (templateBody.firstChild) {
                templateBody.removeChild(templateBody.firstChild);
            }
            
            // Copy all content from target body to template body (except sectPr)
            const targetChildren = Array.from(targetBody.childNodes);
            targetChildren.forEach(child => {
                // Skip the target's sectPr - we want to keep the template's sectPr
                if (child.nodeType === Node.ELEMENT_NODE && child.nodeName === 'w:sectPr') {
                    return;
                }
                
                // Import and append the node to template body
                const importedNode = templateDoc.importNode(child, true);
                templateBody.appendChild(importedNode);
            });
            
            // Add the template's sectPr at the end (this preserves header/footer references)
            if (templateSectPr) {
                templateBody.appendChild(templateSectPr);
            }
            
            // Serialize the modified template document
            return serializer.serializeToString(templateDoc);
            
        } catch (error) {
            throw new Error(`Failed to merge document content: ${error.message}`);
        }
    }

    generateProcessedFileName(originalName) {
        const nameWithoutExt = originalName.replace(/\.docx$/i, '');
        return `${nameWithoutExt}_updated.docx`;
    }

    showProcessingStatus() {
        document.getElementById('processing-status').classList.remove('hidden');
    }

    updateProgress(percentage, text) {
        const progressFill = document.getElementById('progress-fill');
        const progressText = document.getElementById('progress-text');
        
        progressFill.style.width = `${percentage}%`;
        progressText.textContent = `${Math.round(percentage)}% - ${text}`;
    }

    log(message, type = 'info') {
        const logContainer = document.getElementById('processing-log');
        const logEntry = document.createElement('div');
        logEntry.className = `log-entry log-entry--${type}`;
        logEntry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
        
        logContainer.appendChild(logEntry);
        logContainer.scrollTop = logContainer.scrollHeight;
    }

    showDownloadSection() {
        const downloadSection = document.getElementById('download-section');
        const noFilesSection = document.getElementById('no-processed-files');
        const processedFilesContainer = document.getElementById('processed-files');
        
        if (this.processedFiles.length === 0) {
            downloadSection.classList.add('hidden');
            noFilesSection.classList.remove('hidden');
            return;
        }
        
        let html = '';
        this.processedFiles.forEach((file, index) => {
            const fileSize = this.formatFileSize(file.size);
            html += `
                <div class="processed-file-item">
                    <div class="processed-file-info">
                        <div class="processed-file-name">${file.name}</div>
                        <div class="processed-file-status">Processed from: ${file.originalName} • ${fileSize}</div>
                    </div>
                    <div class="processed-file-actions">
                        <button type="button" class="btn btn--secondary btn--sm" onclick="app.previewFile(${index})">
                            Preview
                        </button>
                        <button type="button" class="btn btn--secondary btn--sm" onclick="app.downloadFile(${index})">
                            Download
                        </button>
                    </div>
                </div>
            `;
        });
        
        processedFilesContainer.innerHTML = html;
        downloadSection.classList.remove('hidden');
        noFilesSection.classList.add('hidden');
    }

    downloadFile(index) {
        const file = this.processedFiles[index];
        saveAs(file.blob, file.name);
    }

    async previewFile(index) {
        const file = this.processedFiles[index];
        if (!file) {
            this.logError('Preview Error', 'File not found.');
            return;
        }

        const previewModal = document.getElementById('preview-modal');
        const previewContainer = document.getElementById('preview-container');
        const previewFilename = document.getElementById('preview-filename');

        if (!previewModal || !previewContainer || !previewFilename) {
            this.logError('UI Error', 'Could not find preview modal elements. Please check index.html');
            return;
        }

        previewFilename.textContent = file.name;
        previewContainer.innerHTML = ''; // Clear previous content
        previewModal.classList.remove('hidden');

        try {
            // Check if docx-preview library is available
            const docxPreview = window.docx || window.docxPreview;
            if (!docxPreview) {
                throw new Error('docx-preview library is not loaded. Please refresh the page and try again.');
            }

            // Use the docx-preview library to render the document
            await docxPreview.renderAsync(file.blob, previewContainer);
        } catch (error) {
            this.logError(`Preview Error (${file.name})`, `Failed to render preview: ${error.message}`);
            previewContainer.innerHTML = `<div class="error-message">
                <p><strong>Preview Failed:</strong> ${error.message}</p>
                <p>You can still download the file to check if it opens correctly in Microsoft Word.</p>
            </div>`;
        }
    }

    closePreviewModal() {
        const previewModal = document.getElementById('preview-modal');
        if (previewModal) {
            previewModal.classList.add('hidden');
        }
        const previewContainer = document.getElementById('preview-container');
        if (previewContainer) {
            previewContainer.innerHTML = ''; // Clean up preview content
        }
    }

    async downloadAllFiles() {
        if (this.processedFiles.length === 0) return;
        
        try {
            const zip = new JSZip();
            
            this.processedFiles.forEach(file => {
                zip.file(file.name, file.blob);
            });
            
            const zipBlob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 }
            });
            
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            saveAs(zipBlob, `processed_documents_${timestamp}.zip`);
            
        } catch (error) {
            this.logError('Download Error', 'Failed to create ZIP file: ' + error.message);
        }
    }

    logError(title, message) {
        const errorLog = document.getElementById('error-log');
        const errorDetails = document.getElementById('error-details');
        
        const timestamp = new Date().toLocaleString();
        const errorText = `[${timestamp}] ${title}: ${message}\n`;
        
        errorDetails.textContent += errorText;
        errorLog.classList.remove('hidden');
    }

    clearErrors() {
        const errorLog = document.getElementById('error-log');
        const errorDetails = document.getElementById('error-details');
        
        errorDetails.textContent = '';
        errorLog.classList.add('hidden');
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }
}

// Initialize the application
let app;
document.addEventListener('DOMContentLoaded', () => {
    app = new WordDocumentUpdater();
});

// Global error handler
window.addEventListener('error', (event) => {
    if (app) {
        app.logError('JavaScript Error', event.error?.message || 'Unknown error occurred');
    }
});

// Handle unhandled promise rejections
window.addEventListener('unhandledrejection', (event) => {
    if (app) {
        app.logError('Promise Rejection', event.reason?.message || 'Unknown promise rejection');
    }
});