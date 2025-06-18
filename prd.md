# Product Requirements Document: Word Document Header/Footer Updater

**Version:** 1.0
**Date:** June 18, 2025
**Author:** GitHub Copilot

## 1. Overview

The "Word Document Header/Footer Updater" is a client-side web application designed to solve a common document management problem: applying a consistent header and footer from a template `.docx` file to one or more target `.docx` files. The application operates entirely within the user's browser, ensuring data privacy and security as no files are uploaded to a server.

The primary goal is to automate a tedious manual task, improve document consistency, and save users time, while correctly handling the complex internal structure of `.docx` files to prevent corruption.

## 2. Features

### 2.1. Core Functionality
- **Template Upload:** Users can select a single `.docx` file to serve as the "template." The application extracts the header and footer definitions from this file.
- **Target Document Upload:** Users can select one or multiple `.docx` files to which the template's header and footer will be applied.
- **In-Browser Processing:** The core logic runs entirely on the client-side using JavaScript. It modifies the target documents by merging them with the template's structure.
- **Download Processed Files:**
    - Users can download each processed file individually.
    - Users can download all processed files in a single `.zip` archive for convenience.

### 2.2. User Interface & Experience
- **Drag-and-Drop Zones:** Intuitive drag-and-drop areas for both template and target file uploads.
- **File Previews:**
    - The application displays a text-based preview of the headers and footers extracted from the template file.
    - A full visual preview of the processed document is available in a modal window before downloading.
- **Real-time Feedback:**
    - A progress bar and status log provide real-time updates during the file processing stage.
    - Clear error messages and a dedicated troubleshooting section guide the user.
- **State Management:** The UI dynamically enables/disables the "Process" button based on whether a valid template and at least one target file have been selected.

## 3. Technical Stack

- **Frontend:** HTML5, CSS3, Vanilla JavaScript (ES6+)
- **File Processing:**
    - **JSZip:** For reading and writing the `.docx` file format (which is a zip archive of XML files).
    - **FileSaver.js:** For triggering file downloads from the browser.
- **Document Preview:**
    - **docx-preview.js:** For rendering a visual preview of the processed `.docx` file in the browser.
- **Core APIs:** DOMParser, XMLSerializer, FileReader API.

## 4. Technical Challenges & Solutions

This section documents the critical technical hurdles encountered during development and the solutions that were implemented. This is crucial for avoiding the same pitfalls in future versions.

### 4.1. The Core Problem: `.docx` File Corruption

- **Symptom:** The initial versions of the application produced `.docx` files that could not be opened by Microsoft Word, which reported an "unreadable content" error.
- **Root Cause:** A `.docx` file is not a single entity but a compressed archive of XML files with a strict, interconnected structure. Key files like `[Content_Types].xml` and various relationship files (`_rels/.rels`, `word/_rels/document.xml.rels`) define the document's parts and how they relate to each other. Simply adding or replacing files without updating these relationships and definitions leads to a corrupt file.

### 4.2. Evolution of the Processing Logic

#### Attempt 1: Manual Patching (Failed)
- **Strategy:** Copy the `header*.xml` and `footer*.xml` files from the template zip to the target zip, then manually parse and update `[Content_Types].xml` and `word/_rels/document.xml.rels` to include the new parts and relationships.
- **Result:** This approach was brittle and incomplete. It failed to account for all necessary structural updates, particularly the **section properties (`<w:sectPr>`)** within `word/document.xml`, which actually reference the header/footer parts. The files remained corrupt.

#### Attempt 2: Simple XML Swap (Failed)
- **Strategy:** To simplify, the logic was changed to take the entire `word/document.xml` from the target file and replace the `word/document.xml` in a fresh copy of the template file.
- **Result:** This fixed the file corruption, but it had a major side effect: **the headers and footers disappeared**. This was because the target's `document.xml` contained its own `<w:sectPr>` element (or none at all), which overwrote the template's crucial header/footer references.

#### Attempt 3: Intelligent XML Merge (Successful)
- **Strategy:** This successful approach treats the `document.xml` files as structured data and merges them intelligently.
    1.  Parse both the template's and the target's `word/document.xml` into DOM objects.
    2.  From the **template's** XML, find and preserve the `<w:body>/<w:sectPr>` element. This element contains the vital links to the header and footer files.
    3.  From the **target's** XML, take all child nodes of the `<w:body>` element *except* for its `<w:sectPr>` element. This is the user's content.
    4.  Construct a new `document.xml` in memory: Start with the template's structure, clear its `<w:body>`, append the content nodes from the target, and finally, append the preserved `<w:sectPr>` from the template.
- **Result:** This method successfully combines the target's body content with the template's headers and footers, preserving the necessary relationships and producing a valid, non-corrupt `.docx` file.

### 4.3. UI & Library Issues

- **Previewer Failure:** The document previewer initially failed with a `docx is not defined` error.
    - **Solution:** The CDN link for `docx-preview.js` was found to be unreliable. It was updated to a more stable source (unpkg). Additionally, the `previewFile` function was made more robust by checking if the `window.docx` object exists before attempting to use it, providing a clearer error to the user if the library fails to load.
- **Modal Element Not Found:** The preview modal failed to launch due to a "Could not find preview modal elements" error.
    - **Solution:** This was a simple but important UI bug. The JavaScript was searching for an element with `id="preview-filename"`, but the HTML had `id="preview-modal-title"`. Aligning the IDs in both files fixed the issue.

## 5. Future Improvements

- **Improve Formatting Preservation:** While the current merge logic is effective, subtle formatting differences defined in document-level styles could still be lost. A more advanced merge could involve comparing and merging the style definitions (`styles.xml`) as well.
- **Image and Media Relationship Handling:** The current logic does not explicitly manage relationships for images or other media embedded in the body content. If a target document has images, their `rId`s may conflict or be lost. A future version should parse, remap, and merge relationships from `word/_rels/document.xml.rels`.
- **Support for `.doc`:** Investigate using a server-side component or a WebAssembly-based library (like LibreOffice) to handle the older, binary `.doc` format.
- **User-Selectable Merge Strategy:** Allow users to choose whether to keep the template's formatting or the target's formatting.
