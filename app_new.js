// Word Document Header/Footer Updater Application - Conservative Approach
class WordDocumentUpdater {
constructor() {
// State properties
this.templateFile = null;
this.templateZip = null;
this.targetFiles = [];
this.processedFiles = [];
this.isProcessing = false;
this.debugMode = false;
this.insertFlowChart = false;
this.preserveTargetFonts = false;
this.fontOverride = null;
this.extractTitle = true;

// Configuration
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
],
docProps: ['docProps/core.xml', 'docProps/app.xml'],
mediaFolder: 'word/media/'
};

this.errorMessages = {
invalidFormat: 'Please select a valid .docx file',
fileTooLarge: 'File size exceeds 50MB limit',
protectedDocument: 'Document appears to be protected. Please disable Protected View in Word',
corruptedFile: 'File appears to be corrupted or not a valid Word document',
noHeadersFooters: 'No headers or footers found in template document',
processingError: 'An error occurred while processing the document',
titleExtractionFailed: 'Could not extract title from document'
};

this.init();
}

// Initialize the application, check compatibility, and set up event listeners.
init() {
this.checkBrowserCompatibility();
this.setupEventListeners();
this.updateProcessButton();
}

// Check if the browser supports necessary APIs.
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

// Set up all event listeners for the UI.
setupEventListeners() {
// Template upload listeners
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

// Target files upload listeners
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

// Processing listeners
const processBtn = document.getElementById('process-btn');
const debugMode = document.getElementById('debug-mode');
const insertFlowchart = document.getElementById('insert-flowchart');
const preserveTargetFonts = document.getElementById('preserve-target-fonts');
const fontOverride = document.getElementById('font-override');
const extractTitle = document.getElementById('extract-title');

processBtn.addEventListener('click', () => this.processDocuments());
debugMode?.addEventListener('change', (e) => {
this.debugMode = e.target.checked;
});
insertFlowchart?.addEventListener('change', (e) => {
this.insertFlowChart = e.target.checked;
});
preserveTargetFonts?.addEventListener('change', (e) => {
this.preserveTargetFonts = e.target.checked;
});
fontOverride?.addEventListener('input', (e) => {
this.fontOverride = e.target.value.trim() || null;
});
extractTitle?.addEventListener('change', (e) => {
this.extractTitle = e.target.checked;
});

// Download listeners
const downloadAll = document.getElementById('download-all');
downloadAll.addEventListener('click', () => this.downloadAllFiles());

// Error handling listeners
const clearErrors = document.getElementById('clear-errors');
clearErrors?.addEventListener('click', () => this.clearErrors());

// Preview Modal listeners
const previewModalClose = document.getElementById('preview-modal-close');
previewModalClose?.addEventListener('click', () => this.closePreviewModal());
}

// Handle drag-over event for upload zones.
handleDragOver(e) {
e.preventDefault();
e.currentTarget.classList.add('drag-over');
}

// Handle drag-leave event for upload zones.
handleDragLeave(e) {
e.preventDefault();
e.currentTarget.classList.remove('drag-over');
}

// Handle file drop for the template.
async handleTemplateDrop(e) {
e.preventDefault();
e.currentTarget.classList.remove('drag-over');

const files = Array.from(e.dataTransfer.files);
if (files.length > 0) {
await this.setTemplateFile(files[0]);
}
}

// Handle file selection for the template.
async handleTemplateSelect(e) {
const file = e.target.files[0];
if (file) {
await this.setTemplateFile(file);
}
}

// Handle file drop for target documents.
async handleTargetDrop(e) {
e.preventDefault();
e.currentTarget.classList.remove('drag-over');

const files = Array.from(e.dataTransfer.files);
await this.addTargetFiles(files);
}

// Handle file selection for target documents.
async handleTargetSelect(e) {
const files = Array.from(e.target.files);
await this.addTargetFiles(files);
}

// Set and validate the template file.
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

// Add and validate target files.
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

// Validate a file based on format and size.
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

// Enhanced title extraction with multiple strategies
async extractTitleFromDocument(zip, filename) {
this.log("🔍 Extracting title from target document...", 'info');

try {
// Strategy 1: Document Properties
const corePropsXml = await zip.file('docProps/core.xml')?.async('string');
if (corePropsXml) {
const parser = new DOMParser();
const coreDoc = parser.parseFromString(corePropsXml, 'text/xml');
const titleElement = coreDoc.querySelector('title');
if (titleElement && titleElement.textContent.trim()) {
this.log(`✅ Found title in document properties: "${titleElement.textContent.trim()}"`, 'success');
return titleElement.textContent.trim();
}
}

// Strategy 2: Look for existing headers
const headers = await this.findHeaderText(zip);
if (headers.length > 0) {
this.log(`✅ Found title in document headers: "${headers[0]}"`, 'success');
return headers[0].trim();
}

// Strategy 3: Find first heading (H1, H2, etc.)
const headings = await this.findHeadingText(zip);
if (headings.length > 0) {
this.log(`✅ Found title in document headings: "${headings[0]}"`, 'success');
return headings[0].trim();
}

// Strategy 4: First substantial paragraph
const firstParagraph = await this.findFirstSubstantialText(zip);
if (firstParagraph) {
this.log(`✅ Using first substantial text as title: "${firstParagraph}"`, 'success');
return firstParagraph.trim();
}

// Strategy 5: Filename fallback
const fileTitle = filename ? filename.replace(/\.[^/.]+$/, "").replace(/_/g, ' ') : "Untitled Document";
this.log(`⚠️ Using filename as title: "${fileTitle}"`, 'info');
return fileTitle;

} catch (error) {
this.log(`❌ Error extracting title: ${error.message}`, 'error');
return filename ? filename.replace(/\.[^/.]+$/, "").replace(/_/g, ' ') : "Untitled Document";
}
}

async findHeaderText(zip) {
const headers = [];
try {
// Look in header files
for (const headerPath of this.docxStructure.headers) {
const headerFile = zip.file(headerPath);
if (headerFile) {
const content = await headerFile.async('string');
const parser = new DOMParser();
const doc = parser.parseFromString(content, 'text/xml');
const textNodes = doc.querySelectorAll('w\\:t, t');
const rawText = Array.from(textNodes).map(node => node.textContent).join(' ').trim();

if (rawText && rawText.length > 5) {
// Clean the title by removing metadata
const cleanedTitle = this.cleanTitleFromHeaderText(rawText);
if (cleanedTitle && cleanedTitle.length > 3) {
this.log(`Found header text: "${rawText}"`, 'info');
this.log(`Cleaned title: "${cleanedTitle}"`, 'success');
headers.push(cleanedTitle);
}
}
}
}
} catch (error) {
this.log(`Warning: Could not extract header text - ${error.message}`, 'info');
}
return headers;
}

// Helper method to extract clean title from header text containing metadata
cleanTitleFromHeaderText(rawText) {
try {
// Define metadata keywords that indicate where the title ends
const metadataKeywords = [
'Document',
'Revision',
'DCR#',
'DCR',
'Effective Date',
'Rev.',
'Rev ',
'Version',
'Ver.',
'Date:',
/MF\d+/,  // Pattern for document numbers like MF0415
/\d{1,2}\/\d{1,2}\/\d{4}/,  // Date patterns like 12/26/2019
/\d{4}-\d{2}-\d{2}/  // ISO date patterns
];

let cleanText = rawText;

// Split text into words and process
const words = rawText.split(/\s+/);
let titleWords = [];

for (let i = 0; i < words.length; i++) {
const word = words[i];
let foundMetadata = false;

// Check against keyword patterns
for (const keyword of metadataKeywords) {
if (typeof keyword === 'string') {
if (word.includes(keyword)) {
foundMetadata = true;
break;
}
} else if (keyword instanceof RegExp) {
if (keyword.test(word)) {
foundMetadata = true;
break;
}
}
}

if (foundMetadata) {
this.log(`Stopped title extraction at metadata keyword: "${word}"`, 'info');
break;
}

titleWords.push(word);
}

// Join the title words and clean up
cleanText = titleWords.join(' ').trim();

// Additional cleanup
cleanText = cleanText
.replace(/\s+/g, ' ')  // Normalize whitespace
.replace(/[^\w\s&-()]/g, ' ')  // Remove special chars except common ones
.replace(/\s+/g, ' ')  // Normalize again
.trim();

// Validate the extracted title
if (cleanText.length < 3) {
this.log('Extracted title too short, rejecting', 'warning');
return null;
}

if (cleanText.length > 150) {
this.log('Extracted title too long, truncating', 'warning');
cleanText = cleanText.substring(0, 147) + '...';
}

return cleanText;

} catch (error) {
this.log(`Error cleaning title text: ${error.message}`, 'error');
return null;
}
}

async findHeadingText(zip) {
const headings = [];
try {
const docXml = await zip.file(this.docxStructure.mainDocument)?.async('string');
if (docXml) {
const parser = new DOMParser();
const doc = parser.parseFromString(docXml, 'text/xml');

// Look for style-based headings
const paragraphs = doc.querySelectorAll('w\\:p, p');
for (const p of paragraphs) {
const styleNode = p.querySelector('w\\:pStyle, pStyle');
if (styleNode) {
const styleId = styleNode.getAttribute('w:val') || styleNode.getAttribute('val');
if (styleId && (styleId.toLowerCase().includes('heading') || styleId.toLowerCase().includes('title'))) {
const textNodes = p.querySelectorAll('w\\:t, t');
const text = Array.from(textNodes).map(t => t.textContent).join('').trim();
if (text && text.length > 5 && text.length < 100) {
headings.push(text);
break; // Take first heading
}
}
}
}
}
} catch (error) {
this.log(`Warning: Could not extract headings - ${error.message}`, 'info');
}
return headings;
}

async findFirstSubstantialText(zip) {
try {
const docXml = await zip.file(this.docxStructure.mainDocument)?.async('string');
if (docXml) {
const parser = new DOMParser();
const doc = parser.parseFromString(docXml, 'text/xml');
const textNodes = doc.querySelectorAll('w\\:t, t');
const allText = Array.from(textNodes).map(t => t.textContent).join(' ');
const lines = allText.split(/\s+/).filter(word => word.length > 0);

if (lines.length > 0) {
// Take first 10 words as potential title
let firstLine = lines.slice(0, 10).join(' ');
if (firstLine.length > 80) {
firstLine = firstLine.substring(0, 77) + "...";
}
return firstLine;
}
}
} catch (error) {
this.log(`Warning: Could not extract first text - ${error.message}`, 'info');
}
return null;
}

// CONSERVATIVE APPROACH: Replace only headers/footers in target document
async replaceHeadersAndFooters(targetZip, templateZip, extractedTitle) {
this.log("🔄 Conservative approach: Replacing only headers/footers...", 'info');

try {
let headerFooterCount = 0;

// Get template headers/footers with title replacement
for (const headerPath of this.docxStructure.headers) {
const templateHeader = templateZip.file(headerPath);
if (templateHeader) {
let content = await templateHeader.async('string');

// Replace title placeholder if title extraction is enabled
if (this.extractTitle && extractedTitle) {
const placeholder = "{Enter SOP Title}";
if (content.includes(placeholder)) {
content = content.replace(new RegExp(placeholder, 'g'), extractedTitle);
this.log(`✅ Title inserted in ${headerPath}: "${extractedTitle}"`, 'success');
}
}

// Replace in target document
targetZip.file(headerPath, content);
headerFooterCount++;
this.log(`📄 Replaced header: ${headerPath}`, 'info');
}
}

// Get template footers
for (const footerPath of this.docxStructure.footers) {
const templateFooter = templateZip.file(footerPath);
if (templateFooter) {
let content = await templateFooter.async('string');

// Replace title placeholder if title extraction is enabled
if (this.extractTitle && extractedTitle) {
const placeholder = "{Enter SOP Title}";
if (content.includes(placeholder)) {
content = content.replace(new RegExp(placeholder, 'g'), extractedTitle);
this.log(`✅ Title inserted in ${footerPath}: "${extractedTitle}"`, 'success');
}
}

// Replace in target document
targetZip.file(footerPath, content);
headerFooterCount++;
this.log(`📄 Replaced footer: ${footerPath}`, 'info');
}
}

// CRITICAL: Update section properties to reference template headers/footers
await this.updateSectionPropertiesToReferenceTemplateHeaders(targetZip, templateZip);

this.log(`✅ Successfully replaced ${headerFooterCount} header/footer files`, 'success');
return headerFooterCount;

} catch (error) {
this.log(`❌ Error replacing headers/footers: ${error.message}`, 'error');
throw error;
}
}

// NEW: Update section properties to reference template headers/footers
async updateSectionPropertiesToReferenceTemplateHeaders(targetZip, templateZip) {
this.log("🔗 Updating section properties to reference template headers/footers...", 'info');

try {
const docPath = this.docxStructure.mainDocument;

// Get template document to extract section properties
const templateDocXml = await templateZip.file(docPath)?.async('string');
const targetDocXml = await targetZip.file(docPath)?.async('string');

if (!templateDocXml || !targetDocXml) {
throw new Error('Could not find document.xml in template or target');
}

const parser = new DOMParser();
const serializer = new XMLSerializer();

const templateDoc = parser.parseFromString(templateDocXml, 'text/xml');
const targetDoc = parser.parseFromString(targetDocXml, 'text/xml');

// Extract template section properties (header/footer references)
const templateSectPr = templateDoc.querySelector('body sectPr');
const targetSectPr = targetDoc.querySelector('body sectPr');

if (templateSectPr && targetSectPr) {
// Remove old header/footer references from target
const oldHeaderRefs = targetSectPr.querySelectorAll('headerReference');
const oldFooterRefs = targetSectPr.querySelectorAll('footerReference');

oldHeaderRefs.forEach(ref => ref.remove());
oldFooterRefs.forEach(ref => ref.remove());

// Copy template header/footer references
const templateHeaderRefs = templateSectPr.querySelectorAll('headerReference');
const templateFooterRefs = templateSectPr.querySelectorAll('footerReference');

templateHeaderRefs.forEach(ref => {
const clonedRef = targetDoc.importNode(ref, true);
targetSectPr.appendChild(clonedRef);
});

templateFooterRefs.forEach(ref => {
const clonedRef = targetDoc.importNode(ref, true);
targetSectPr.appendChild(clonedRef);
});

// Update target document
const updatedDocXml = serializer.serializeToString(targetDoc);
targetZip.file(docPath, updatedDocXml);

this.log(`✅ Updated section properties with template header/footer references`, 'success');
} else {
this.log(`⚠️ Could not find section properties in template or target document`, 'warning');
}

} catch (error) {
this.log(`❌ Error updating section properties: ${error.message}`, 'error');
throw error;
}
}

// Extract the template's default body font dynamically
async extractTemplateDefaultFont() {
try {
const stylesContent = await this.templateZip.file('word/styles.xml')?.async('string');
if (!stylesContent) return 'Calibri'; // fallback

const parser = new DOMParser();
const stylesDoc = parser.parseFromString(stylesContent, 'text/xml');

// Check document defaults first
const docDefaults = stylesDoc.querySelector('docDefaults rPrDefault rPr rFonts');
if (docDefaults) {
const font = docDefaults.getAttribute('w:ascii') || 
docDefaults.getAttribute('w:hAnsi');
if (font) {
this.log(`Found template font in document defaults: ${font}`, 'info');
return font;
}
}

// Check Normal style as fallback
const normalStyle = stylesDoc.querySelector('style[w\\:styleId="Normal"] rPr rFonts');
if (normalStyle) {
const font = normalStyle.getAttribute('w:ascii') || 
normalStyle.getAttribute('w:hAnsi');
if (font) {
this.log(`Found template font in Normal style: ${font}`, 'info');
return font;
}
}

this.log('Could not extract template font, using Calibri as fallback', 'warning');
return 'Calibri'; // final fallback
} catch (error) {
this.log('Error extracting template font, using Calibri', 'warning');
return 'Calibri';
}
}

// Extract and preview the content of headers and footers from the template.
async extractAndPreviewTemplate() {
try {
const headerFooterContent = [];

// Extract headers
for (const headerPath of this.docxStructure.headers) {
const headerFile = this.templateZip.file(headerPath);
if (headerFile) {
const content = await headerFile.async('text');
const previewText = this.formatXmlForPreview(content);
headerFooterContent.push({
type: 'Header',
path: headerPath,
content: previewText,
hasPlaceholder: content.includes('{Enter SOP Title}')
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
content: this.formatXmlForPreview(content),
hasPlaceholder: content.includes('{Enter SOP Title}')
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

// Format XML content to show only plain text for the preview.
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

// UI update: Show the status of the uploaded template file.
showTemplateStatus() {
const statusSection = document.getElementById('template-status');
const filename = document.getElementById('template-filename');

filename.textContent = this.templateFile.name;
statusSection.classList.remove('hidden');
}

// UI update: Show the preview of extracted headers and footers.
showHeaderFooterPreview(content) {
const previewSection = document.getElementById('template-preview');
const previewContent = document.getElementById('preview-content');

let html = '';
content.forEach(item => {
const placeholderInfo = item.hasPlaceholder ? ' 📝 Contains title placeholder' : '';
html += `<div><strong>${item.type} (${item.path}):</strong>${placeholderInfo}<br>${item.content}</div><br>`;
});

previewContent.innerHTML = html;
previewSection.classList.remove('hidden');
}

// UI update: Show the list of target files.
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

// Remove the selected template file.
removeTemplate() {
this.templateFile = null;
this.templateZip = null;
document.getElementById('template-status').classList.add('hidden');
document.getElementById('template-preview').classList.add('hidden');
document.getElementById('template-input').value = '';
this.updateProcessButton();
}

// Remove a specific target file from the list.
removeTargetFile(index) {
this.targetFiles.splice(index, 1);
this.showTargetFiles();
this.updateProcessButton();
}

// Clear all target files.
clearTargetFiles() {
this.targetFiles = [];
document.getElementById('target-input').value = '';
this.showTargetFiles();
this.updateProcessButton();
}

// Enable or disable the "Process" button based on application state.
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

// Main function to start processing all target documents.
async processDocuments() {
if (this.isProcessing) return;

this.isProcessing = true;
this.processedFiles = [];
this.updateProcessButton();
this.showProcessingStatus();

try {
this.log('Starting conservative document processing (preserving all target content)...', 'info');
this.log('='.repeat(70), 'info');

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
this.logError(`Processing Error (${file.name})`, error.stack);
}
}

this.updateProgress(100, 'Processing complete!');
this.log(`Processing complete. ${this.processedFiles.length}/${this.targetFiles.length} files processed successfully.`, 'success');

this.showDownloadSection();

} catch (error) {
this.log(`Processing failed: ${error.message}`, 'error');
this.logError('Processing Error', error.stack);
} finally {
this.isProcessing = false;
this.updateProcessButton();
}
}

// CONSERVATIVE PROCESSING: Start with target, replace only headers/footers
async processDocument(file) {
try {
this.log(`🏁 CONSERVATIVE PROCESSING: ${file.name}`, 'info');
this.log('Starting with target document to preserve all content including images...', 'info');

// Load target document as base (preserves all images and content)
const targetZip = await JSZip.loadAsync(file);

// Step 1: Extract title from target document
let extractedTitle = '';
if (this.extractTitle) {
extractedTitle = await this.extractTitleFromDocument(targetZip, file.name);
this.log(`📝 Extracted title: "${extractedTitle}"`, 'success');
}

// Step 2: Replace ONLY headers/footers from template
const headerFooterCount = await this.replaceHeadersAndFooters(targetZip, this.templateZip, extractedTitle);

// Step 3: Font handling (optional)
const targetFont = this.fontOverride || 
(this.preserveTargetFonts ? null : await this.extractTemplateDefaultFont());

if (targetFont && !this.preserveTargetFonts) {
this.log(`Applying font: ${targetFont}`, 'info');

// Apply font to styles.xml
const targetStylesXml = await targetZip.file('word/styles.xml')?.async('string');
if (targetStylesXml) {
const modifiedStylesXml = this.forceAllStylesToFont(targetStylesXml, targetFont);
targetZip.file('word/styles.xml', modifiedStylesXml);
} else {
this.log(`Target document is missing styles.xml, cannot apply font.`, 'warning');
}

// Apply template's font table
const templateFontTable = await this.templateZip.file('word/fontTable.xml')?.async('string');
if (templateFontTable) {
targetZip.file('word/fontTable.xml', templateFontTable);
this.log('Applied template font table', 'info');
}
} else if (this.preserveTargetFonts) {
this.log('Preserving target document fonts', 'info');
}

// Step 4: Apply inline font corrections if needed
if (targetFont && !this.preserveTargetFonts) {
let targetDocXml = await targetZip.file(this.docxStructure.mainDocument)?.async('string');
if (targetDocXml) {
targetDocXml = this.correctInlineStyleFonts(targetDocXml, targetFont);

// Apply Process Flow Chart insertion if needed
if (this.insertFlowChart) {
this.log('Attempting to insert "Process Flow Chart" section...', 'info');
const parser = new DOMParser();
const serializer = new XMLSerializer();
let targetDoc = parser.parseFromString(targetDocXml, 'text/xml');
const inserted = this.insertProcessFlowChartSection(targetDoc);
if (inserted) {
targetDocXml = serializer.serializeToString(targetDoc);
}
}

targetZip.file(this.docxStructure.mainDocument, targetDocXml);
}
}

// Step 5: Generate final document
const processedBlob = await targetZip.generateAsync({
type: 'blob',
mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
});

const processedFileName = this.generateProcessedFileName(file.name);

this.log(`✅ Conservative processing completed: ${headerFooterCount} headers/footers replaced`, 'success');
this.log('All original content and images preserved!', 'success');

return {
name: processedFileName,
originalName: file.name,
blob: processedBlob,
size: processedBlob.size,
extractedTitle: extractedTitle,
preservedImages: true
};

} catch (error) {
console.error(`Error processing ${file.name}:`, error);
throw new Error(`Failed to process document: ${error.message}`);
}
}

// Fix fonts in document content inline styles
correctInlineStyleFonts(documentXml, targetFont) {
const parser = new DOMParser();
const serializer = new XMLSerializer();
const doc = parser.parseFromString(documentXml, 'text/xml');

// Find all run properties with font specifications
const rFontsElements = doc.querySelectorAll('rFonts');

rFontsElements.forEach(rFonts => {
// Update all font attributes
if (rFonts.hasAttribute('w:ascii')) rFonts.setAttribute('w:ascii', targetFont);
if (rFonts.hasAttribute('w:hAnsi')) rFonts.setAttribute('w:hAnsi', targetFont);
if (rFonts.hasAttribute('w:cs')) rFonts.setAttribute('w:cs', targetFont);

// Add missing attributes
if (!rFonts.hasAttribute('w:ascii')) rFonts.setAttribute('w:ascii', targetFont);
if (!rFonts.hasAttribute('w:hAnsi')) rFonts.setAttribute('w:hAnsi', targetFont);
});

this.log(`Corrected ${rFontsElements.length} inline font references`, 'info');

return serializer.serializeToString(doc);
}

// Brute-force method to change every style to a specific font.
forceAllStylesToFont(stylesXmlString, fontName) {
const parser = new DOMParser();
const serializer = new XMLSerializer();
const stylesDoc = parser.parseFromString(stylesXmlString, "text/xml");

// Function to set the font on a given rPr (run properties) node.
const setFont = (rPrNode) => {
if (!rPrNode) return;
let fontNode = rPrNode.querySelector('rFonts');
if (!fontNode) {
fontNode = stylesDoc.createElementNS(this.xmlNamespaces.w, 'w:rFonts');
rPrNode.appendChild(fontNode);
}
fontNode.setAttributeNS(this.xmlNamespaces.w, 'w:ascii', fontName);
fontNode.setAttributeNS(this.xmlNamespaces.w, 'w:hAnsi', fontName);
fontNode.setAttributeNS(this.xmlNamespaces.w, 'w:cs', fontName);
};

// 1. Force the document defaults
let rPrDefault = stylesDoc.querySelector('docDefaults > rPrDefault > rPr');
if (rPrDefault) {
setFont(rPrDefault);
}

// 2. Force every single style definition
const allStyles = stylesDoc.querySelectorAll('style');
allStyles.forEach(style => {
let rPrNode = style.querySelector('rPr');
if (!rPrNode) {
rPrNode = stylesDoc.createElementNS(this.xmlNamespaces.w, 'w:rPr');
// Insert it after the name or at the beginning if no name
const nameNode = style.querySelector('name');
if (nameNode) {
nameNode.insertAdjacentElement('afterend', rPrNode);
} else {
style.prepend(rPrNode);
}
}
setFont(rPrNode);
});

if (this.debugMode) this.log(`Forced ALL styles to use ${fontName}.`, 'success');

return serializer.serializeToString(stylesDoc);
}

// REVISED FUNCTION 5.0: Finds "PROCEDURE", clones it, formats it, and inserts an indented "NA" paragraph and a blank line.
insertProcessFlowChartSection(doc) {
const body = doc.querySelector('body');
if (!body) return false;

const paragraphs = body.querySelectorAll('p');
let targetParagraph = null;
let subSectionParagraph = null;

for (let i = 0; i < paragraphs.length; i++) {
const p = paragraphs[i];
const numPrNode = p.querySelector('pPr > numPr');
const textContent = Array.from(p.querySelectorAll('t')).map(t => t.textContent).join('').trim();

if (numPrNode && textContent.toUpperCase().includes('PROCEDURE')) {
targetParagraph = p;
if (i + 1 < paragraphs.length) {
subSectionParagraph = paragraphs[i+1];
}
break; 
}
}

if (targetParagraph) {
if(this.debugMode) this.log('Found "PROCEDURE" heading. Cloning and inserting new section.', 'info');

const newHeading = targetParagraph.cloneNode(true);

const existingRuns = newHeading.querySelectorAll('r');
existingRuns.forEach(run => run.remove());

const newRun = doc.createElementNS(this.xmlNamespaces.w, 'w:r');
const runProps = doc.createElementNS(this.xmlNamespaces.w, 'w:rPr');
const boldTag = doc.createElementNS(this.xmlNamespaces.w, 'w:b');
runProps.appendChild(boldTag);
newRun.appendChild(runProps);

const newText = doc.createElementNS(this.xmlNamespaces.w, 'w:t');
newText.textContent = 'PROCESS FLOW CHART';
newText.setAttribute('xml:space', 'preserve'); 
newRun.appendChild(newText);
newHeading.appendChild(newRun);

const naParagraph = doc.createElementNS(this.xmlNamespaces.w, 'w:p');
if (subSectionParagraph) {
const subSectionProps = subSectionParagraph.querySelector('pPr');
if (subSectionProps) {
const clonedProps = subSectionProps.cloneNode(true);
const numPr = clonedProps.querySelector('numPr');
if (numPr) numPr.remove();
naParagraph.appendChild(clonedProps);
if (this.debugMode) this.log('Cloned sub-section style for NA indentation.', 'info');
}
}
const naRun = doc.createElementNS(this.xmlNamespaces.w, 'w:r');
const naText = doc.createElementNS(this.xmlNamespaces.w, 'w:t');
naText.textContent = 'NA';
naRun.appendChild(naText);
naParagraph.appendChild(naRun);

// Create a blank paragraph for spacing
const blankParagraph = doc.createElementNS(this.xmlNamespaces.w, 'w:p');

// Insert the new heading before the "PROCEDURE" paragraph
targetParagraph.parentNode.insertBefore(newHeading, targetParagraph);
// Insert the "NA" paragraph right after the new heading
targetParagraph.parentNode.insertBefore(naParagraph, targetParagraph);
// Insert the blank line after the "NA" paragraph
targetParagraph.parentNode.insertBefore(blankParagraph, targetParagraph);


this.log('Successfully inserted "PROCESS FLOW CHART" section with content and spacing.', 'success');
return true;

} else {
this.log('Could not find a numbered heading containing "PROCEDURE". Skipping section insertion.', 'warning');
return false; 
}
}

// Generate a new filename for the processed document.
generateProcessedFileName(originalName) {
const nameWithoutExt = originalName.replace(/\.docx$/i, '');
return `${nameWithoutExt}_updated.docx`;
}

// UI update: Show the processing status section.
showProcessingStatus() {
document.getElementById('processing-status').classList.remove('hidden');
}

// UI update: Update the progress bar and text.
updateProgress(percentage, text) {
const progressFill = document.getElementById('progress-fill');
const progressText = document.getElementById('progress-text');

progressFill.style.width = `${percentage}%`;
progressText.textContent = `${Math.round(percentage)}% - ${text}`;
}

// Log a message to the processing log in the UI.
log(message, type = 'info') {
const logContainer = document.getElementById('processing-log');
const logEntry = document.createElement('div');
logEntry.className = `log-entry log-entry--${type}`;
logEntry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;

logContainer.appendChild(logEntry);
logContainer.scrollTop = logContainer.scrollHeight;
}

// UI update: Show the download section with processed files.
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
const titleInfo = file.extractedTitle ? ` • Title: "${file.extractedTitle}"` : '';
const imageInfo = file.preservedImages ? ' • Images: Preserved' : '';
html += `
<div class="processed-file-item">
<div class="processed-file-info">
<div class="processed-file-name">${file.name}</div>
<div class="processed-file-status">Processed from: ${file.originalName} • ${fileSize}${titleInfo}${imageInfo}</div>
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

// Download a single processed file.
downloadFile(index) {
const file = this.processedFiles[index];
saveAs(file.blob, file.name);
}

// Show a preview of a processed file in a modal.
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
previewContainer.innerHTML = 'Loading preview...'; // Clear previous content
previewModal.classList.remove('hidden');

try {
const docx = window.docx;
if (!docx) {
throw new Error('docx-preview library is not loaded.');
}

await docx.renderAsync(file.blob, previewContainer);
} catch (error) {
this.logError(`Preview Error (${file.name})`, `Failed to render preview: ${error.message}`);
previewContainer.innerHTML = `<div class="error-message" style="color:var(--color-error); padding:1rem;">
<p><strong>Preview Failed:</strong> ${error.message}</p>
<p>You can still download the file to check if it opens correctly in Microsoft Word.</p>
</div>`;
}
}

// Close the preview modal.
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

// Download all processed files as a single ZIP archive.
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

// Log an error to the error log in the UI.
logError(title, message) {
const errorLog = document.getElementById('error-log');
const errorDetails = document.getElementById('error-details');

const timestamp = new Date().toLocaleString();
const errorText = `[${timestamp}] ${title}: ${message}\n`;

errorDetails.textContent += errorText;
errorLog.classList.remove('hidden');
}

// Clear the error log.
clearErrors() {
const errorLog = document.getElementById('error-log');
const errorDetails = document.getElementById('error-details');

errorDetails.textContent = '';
errorLog.classList.add('hidden');
}

// Format file size for display.
formatFileSize(bytes) {
if (bytes === 0) return '0 Bytes';

const k = 1024;
const sizes = ['Bytes', 'KB', 'MB', 'GB'];
const i = Math.floor(Math.log(bytes) / Math.log(k));

return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}
}

// Initialize the application once the DOM is loaded.
let app;
document.addEventListener('DOMContentLoaded', () => {
app = new WordDocumentUpdater();
});

// Global error handler for uncaught exceptions.
window.addEventListener('error', (event) => {
if (app) {
app.logError('JavaScript Error', event.error?.stack || event.error?.message || 'Unknown error occurred');
}
});

// Global handler for unhandled promise rejections.
window.addEventListener('unhandledrejection', (event) => {
if (app) {
app.logError('Promise Rejection', event.reason?.stack || event.reason?.message || 'Unknown promise rejection');
}
});
