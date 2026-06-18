/**
 * FileEditorJS - A simple library to view and edit DOCX and XLSX files.
 */
class FileEditor {
  constructor(containerId) {
    this.container = document.getElementById(containerId);
    if (!this.container) {
      throw new Error(`Container with id "${containerId}" not found.`);
    }
    this.currentFile = null;
    this.currentFileType = null; // 'docx' or 'xlsx'
    this.data = null; // Holds the parsed data or original blob
  }

  /**
   * Load a file (Blob or File object)
   * @param {File|Blob} file
   */
  async loadFile(file) {
    this.currentFile = file;
    const extension = file.name.split(".").pop().toLowerCase();

    if (extension === "docx") {
      this.currentFileType = "docx";
      await this.renderDocx(file);
    } else if (extension === "xlsx" || extension === "xls") {
      this.currentFileType = "xlsx";
      await this.renderXlsx(file);
    } else {
      throw new Error(
        "Unsupported file format. Please provide a .docx or .xlsx file.",
      );
    }
  }

  /**
   * Render DOCX using docx-preview
   */
  async renderDocx(file) {
    this.container.innerHTML = '<div id="docx-container"></div>';
    const docxContainer = document.getElementById("docx-container");

    // mammoth can be used for editing (HTML conversion)
    // docx-preview is better for high-fidelity viewing
    if (window.docx) {
      await window.docx.renderAsync(file, docxContainer);
      // Make it editable? docx-preview renders to HTML elements.
      // For real editing, we usually convert to HTML via mammoth and put in a rich text editor.
      this.setupDocxEditing(file);
    } else {
      this.container.innerHTML = "docx-preview library not loaded.";
    }
  }

  async setupDocxEditing(file) {
    if (window.mammoth) {
      const arrayBuffer = await file.arrayBuffer();
      const result = await window.mammoth.convertToHtml({
        arrayBuffer: arrayBuffer,
      });
      this.container.innerHTML = `
                <div style="margin-bottom: 10px;">
                    <button id="save-docx" class="btn">Download Edited DOCX</button>
                    <small> (Editing is experimental: HTML to DOCX conversion)</small>
                </div>
                <div id="editor-content" contenteditable="true" style="border: 1px solid #ccc; padding: 20px; min-height: 400px; background: white;">
                    ${result.value}
                </div>
            `;

      document.getElementById("save-docx").onclick = () => this.saveDocx();
    }
  }

  /**
   * Render XLSX using SheetJS
   */
  async renderXlsx(file) {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: "array" });
    this.data = workbook;

    this.renderSheet(workbook, 0);
  }

  renderSheet(workbook, sheetIndex) {
    const sheetName = workbook.SheetNames[sheetIndex];
    const worksheet = workbook.Sheets[sheetName];
    const html = XLSX.utils.sheet_to_html(worksheet, { editable: true });

    this.container.innerHTML = `
            <div style="margin-bottom: 10px;">
                <select id="sheet-selector">
                    ${workbook.SheetNames.map((name, i) => `<option value="${i}" ${i === sheetIndex ? "selected" : ""}>${name}</option>`).join("")}
                </select>
                <button id="save-xlsx" class="btn">Download Edited XLSX</button>
            </div>
            <div id="xlsx-container" style="overflow: auto;">
                ${html}
            </div>
        `;

    document.getElementById("sheet-selector").onchange = (e) => {
      this.renderSheet(workbook, parseInt(e.target.value));
    };

    document.getElementById("save-xlsx").onclick = () => this.saveXlsx();
  }

  async saveDocx() {
    const content = document.getElementById("editor-content").innerHTML;

    // We use html-docx-js to convert HTML to DOCX blob
    if (window.htmlDocx) {
      const fullHtml = `
            <!DOCTYPE html>
            <html>
                <head>
                    <meta charset="utf-8">
                    <title>Exported DOCX</title>
                </head>
                <body>
                    ${content}
                </body>
            </html>
        `;
      const blob = window.htmlDocx.asBlob(fullHtml);
      this.downloadBlob(blob, "edited.docx");
    } else {
      alert("html-docx-js library not loaded. Falling back to HTML download.");
      const blob = new Blob([content], { type: "text/html" });
      this.downloadBlob(blob, "edited.html");
    }
  }

  async saveXlsx() {
    const table = this.container.querySelector("table");
    const newWorksheet = XLSX.utils.table_to_sheet(table);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sheet1");

    const wbout = XLSX.write(newWorkbook, { bookType: "xlsx", type: "binary" });

    function s2ab(s) {
      const buf = new ArrayBuffer(s.length);
      const view = new Uint8Array(buf);
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xff;
      return buf;
    }

    this.downloadBlob(
      new Blob([s2ab(wbout)], { type: "application/octet-stream" }),
      "edited.xlsx",
    );
  }

  downloadBlob(blob, fileName) {
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    a.click();
    window.URL.revokeObjectURL(url);
  }
}
