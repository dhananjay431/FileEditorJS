# FileEditorJS

A lightweight JavaScript library to view and edit `.docx` and `.xlsx` files in the browser.

## Features

- **DOCX Viewing**: Uses `docx-preview` for high-fidelity rendering.
- **DOCX Editing**: Uses `mammoth.js` to convert to HTML for editing (ContentEditable).
- **XLSX Viewing & Editing**: Uses `SheetJS` (xlsx) to render spreadsheets as interactive HTML tables.
- **Pure JavaScript**: No framework dependencies.

## Installation

You can include the library and its dependencies via CDN:

```html
<!-- Dependencies -->
<script src="https://unpkg.com/jszip/dist/jszip.min.js"></script>
<script src="https://unpkg.com/mammoth@1.6.0/mammoth.browser.min.js"></script>
<script src="https://unpkg.com/docx-preview/dist/docx-preview.min.js"></script>
<script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>

<!-- Library -->
<script src="path/to/src/FileEditor.js"></script>
```

## Usage

```javascript
// Initialize the editor in a container
const editor = new FileEditor("container-id");

// Load a file (from an input change event or fetch)
fileInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  await editor.loadFile(file);
});
```

## Demo

Check out the `examples/index.html` file to see it in action.
# FileEditorJS
