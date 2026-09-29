const vscode = require('vscode');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

class ExcelDocument {
  constructor(uri, workbook) {
    this._uri = uri;
    this._workbook = workbook;
    this._onDidChange = new vscode.EventEmitter();
    this.onDidChange = this._onDidChange.event;
    this._edits = [];
    this._savedEdits = [];
  }

  get uri() {
    return this._uri;
  }

  get workbook() {
    return this._workbook;
  }

  getSheetNames() {
    return this._workbook.SheetNames;
  }

  getSheetData(sheetName) {
    const sheet = this._workbook.Sheets[sheetName];
    if (!sheet) return { headers: [], rows: [], range: null };

    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    const data = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const row = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellAddress = XLSX.utils.encode_cell({ r, c });
        const cell = sheet[cellAddress];
        row.push({
          value: cell ? cell.v : '',
          formatted: cell ? XLSX.utils.format_cell(cell) : '',
          type: cell ? cell.t : 's',
          address: cellAddress,
        });
      }
      data.push(row);
    }

    return {
      data,
      range: {
        startRow: range.s.r,
        endRow: range.e.r,
        startCol: range.s.c,
        endCol: range.e.c,
      },
      merges: sheet['!merges'] || [],
      colWidths: sheet['!cols'] || [],
    };
  }

  applyEdit(edit) {
    const { sheetName, address, value } = edit;
    const sheet = this._workbook.Sheets[sheetName];
    if (!sheet) return;

    const cellRef = address;
    if (value === '' || value === null || value === undefined) {
      delete sheet[cellRef];
    } else {
      // Try to parse as number
      const numValue = Number(value);
      if (!isNaN(numValue) && value.trim() !== '') {
        sheet[cellRef] = { t: 'n', v: numValue };
      } else {
        sheet[cellRef] = { t: 's', v: value };
      }
    }

    // Update the range if needed
    const cell = XLSX.utils.decode_cell(cellRef);
    const currentRange = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    if (cell.r > currentRange.e.r) currentRange.e.r = cell.r;
    if (cell.c > currentRange.e.c) currentRange.e.c = cell.c;
    sheet['!ref'] = XLSX.utils.encode_range(currentRange);

    this._edits.push(edit);
    this._onDidChange.fire();
  }

  async save() {
    const buffer = XLSX.write(this._workbook, {
      type: 'buffer',
      bookType: 'xlsx',
    });
    await vscode.workspace.fs.writeFile(this._uri, buffer);
    this._savedEdits = [...this._edits];
  }

  async saveAs(targetUri) {
    const buffer = XLSX.write(this._workbook, {
      type: 'buffer',
      bookType: 'xlsx',
    });
    await vscode.workspace.fs.writeFile(targetUri, buffer);
  }

  dispose() {
    this._onDidChange.dispose();
  }
}

class ExcelEditorProvider {
  constructor(context) {
    this._context = context;
    this._onDidChangeCustomDocument = new vscode.EventEmitter();
    this.onDidChangeCustomDocument = this._onDidChangeCustomDocument.event;
  }

  async openCustomDocument(uri, openContext, token) {
    const fileData = await vscode.workspace.fs.readFile(uri);
    const workbook = XLSX.read(fileData, { type: 'array', cellStyles: true });
    const document = new ExcelDocument(uri, workbook);

    document.onDidChange(() => {
      this._onDidChangeCustomDocument.fire({
        document,
        undo: async () => {},
        redo: async () => {},
      });
    });

    return document;
  }

  async resolveCustomEditor(document, webviewPanel, token) {
    webviewPanel.webview.options = {
      enableScripts: true,
    };

    webviewPanel.webview.html = this._getHtmlForWebview(
      webviewPanel.webview,
      document
    );

    // Handle messages from the webview
    webviewPanel.webview.onDidReceiveMessage((message) => {
      switch (message.type) {
        case 'ready':
          this._sendDataToWebview(webviewPanel.webview, document);
          break;

        case 'edit':
          document.applyEdit({
            sheetName: message.sheetName,
            address: message.address,
            value: message.value,
          });
          break;

        case 'requestSheetData':
          this._sendSheetData(
            webviewPanel.webview,
            document,
            message.sheetName
          );
          break;

        case 'exportCSV': {
          const sheet =
            document.workbook.Sheets[message.sheetName];
          if (sheet) {
            const csv = XLSX.utils.sheet_to_csv(sheet);
            this._saveExport(document.uri, csv, 'csv', message.sheetName);
          }
          break;
        }

        case 'exportJSON': {
          const jsonSheet =
            document.workbook.Sheets[message.sheetName];
          if (jsonSheet) {
            const json = XLSX.utils.sheet_to_json(jsonSheet);
            this._saveExport(
              document.uri,
              JSON.stringify(json, null, 2),
              'json',
              message.sheetName
            );
          }
          break;
        }

        case 'addRow':
          this._addRow(webviewPanel.webview, document, message.sheetName, message.afterRow);
          break;

        case 'addColumn':
          this._addColumn(webviewPanel.webview, document, message.sheetName, message.afterCol);
          break;
      }
    });
  }

  _sendDataToWebview(webview, document) {
    const sheetNames = document.getSheetNames();
    const activeSheet = sheetNames[0];
    const sheetData = document.getSheetData(activeSheet);

    webview.postMessage({
      type: 'init',
      sheetNames,
      activeSheet,
      sheetData,
    });
  }

  _sendSheetData(webview, document, sheetName) {
    const sheetData = document.getSheetData(sheetName);
    webview.postMessage({
      type: 'sheetData',
      sheetName,
      sheetData,
    });
  }

  _addRow(webview, document, sheetName, afterRow) {
    const sheet = document.workbook.Sheets[sheetName];
    if (!sheet) return;
    // Shift rows down from afterRow
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    for (let r = range.e.r; r > afterRow; r--) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const fromAddr = XLSX.utils.encode_cell({ r, c });
        const toAddr = XLSX.utils.encode_cell({ r: r + 1, c });
        if (sheet[fromAddr]) {
          sheet[toAddr] = { ...sheet[fromAddr] };
        } else {
          delete sheet[toAddr];
        }
      }
    }
    // Clear the new row
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r: afterRow + 1, c });
      delete sheet[addr];
    }
    range.e.r++;
    sheet['!ref'] = XLSX.utils.encode_range(range);
    document._onDidChange.fire();
    this._sendSheetData(webview, document, sheetName);
  }

  _addColumn(webview, document, sheetName, afterCol) {
    const sheet = document.workbook.Sheets[sheetName];
    if (!sheet) return;
    const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    for (let c = range.e.c; c > afterCol; c--) {
      for (let r = range.s.r; r <= range.e.r; r++) {
        const fromAddr = XLSX.utils.encode_cell({ r, c });
        const toAddr = XLSX.utils.encode_cell({ r, c: c + 1 });
        if (sheet[fromAddr]) {
          sheet[toAddr] = { ...sheet[fromAddr] };
        } else {
          delete sheet[toAddr];
        }
      }
    }
    for (let r = range.s.r; r <= range.e.r; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: afterCol + 1 });
      delete sheet[addr];
    }
    range.e.c++;
    sheet['!ref'] = XLSX.utils.encode_range(range);
    document._onDidChange.fire();
    this._sendSheetData(webview, document, sheetName);
  }

  async _saveExport(originalUri, content, format, sheetName) {
    const originalPath = originalUri.fsPath;
    const baseName = path.basename(originalPath, path.extname(originalPath));
    const defaultUri = vscode.Uri.file(
      path.join(
        path.dirname(originalPath),
        `${baseName}_${sheetName}.${format}`
      )
    );

    const saveUri = await vscode.window.showSaveDialog({
      defaultUri,
      filters: {
        [format.toUpperCase()]: [format],
      },
    });

    if (saveUri) {
      await vscode.workspace.fs.writeFile(
        saveUri,
        Buffer.from(content, 'utf-8')
      );
      vscode.window.showInformationMessage(
        `Exported ${sheetName} to ${saveUri.fsPath}`
      );
    }
  }

  async saveCustomDocument(document, cancellation) {
    await document.save();
  }

  async saveCustomDocumentAs(document, destination, cancellation) {
    await document.saveAs(destination);
  }

  async revertCustomDocument(document, cancellation) {
    const fileData = await vscode.workspace.fs.readFile(document.uri);
    const workbook = XLSX.read(fileData, { type: 'array', cellStyles: true });
    document._workbook = workbook;
    document._edits = [];
    document._savedEdits = [];
  }

  async backupCustomDocument(document, context, cancellation) {
    return {
      id: context.destination.toString(),
      delete: async () => {},
    };
  }

  _getHtmlForWebview(webview, document) {
    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Excel Viewer Pro</title>
  <style>
    ${this._getStyles()}
  </style>
</head>
<body>
  <div id="app">
    <div class="toolbar">
      <div class="toolbar-left">
        <div class="file-info">
          <svg class="file-icon" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <path d="M14 2H6C4.9 2 4 2.9 4 4V20C4 21.1 4.9 22 6 22H18C19.1 22 20 21.1 20 20V8L14 2Z" fill="#217346"/>
            <path d="M14 2V8H20" fill="#33A867"/>
            <text x="12" y="17" text-anchor="middle" fill="white" font-size="7" font-weight="bold">XL</text>
          </svg>
          <span id="fileName" class="file-name"></span>
        </div>
      </div>
      <div class="toolbar-center">
        <div class="cell-reference">
          <input type="text" id="cellRef" class="cell-ref-input" readonly />
        </div>
        <div class="formula-bar">
          <span class="fx-label">fx</span>
          <input type="text" id="formulaBar" class="formula-input" placeholder="Cell value" />
        </div>
      </div>
      <div class="toolbar-right">
        <div class="toolbar-actions">
          <button class="toolbar-btn" id="btnSearch" title="Search (Ctrl+F)">
            <svg viewBox="0 0 16 16" fill="currentColor"><path d="M11.742 10.344a6.5 6.5 0 1 0-1.397 1.398h-.001l3.85 3.85a1 1 0 0 0 1.415-1.414l-3.85-3.85zm-5.442.156a4.5 4.5 0 1 1 0-9 4.5 4.5 0 0 1 0 9z"/></svg>
          </button>
          <div class="dropdown" id="exportDropdown">
            <button class="toolbar-btn" id="btnExport" title="Export">
              <svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a.5.5 0 0 1 .5.5v11.793l3.146-3.147a.5.5 0 0 1 .708.708l-4 4a.5.5 0 0 1-.708 0l-4-4a.5.5 0 0 1 .708-.708L7.5 13.293V1.5A.5.5 0 0 1 8 1z"/><path d="M2 14.5a.5.5 0 0 1 .5-.5h11a.5.5 0 0 1 0 1h-11a.5.5 0 0 1-.5-.5z"/></svg>
            </button>
            <div class="dropdown-menu" id="exportMenu">
              <button class="dropdown-item" id="exportCSV">Export as CSV</button>
              <button class="dropdown-item" id="exportJSON">Export as JSON</button>
            </div>
          </div>
          <button class="toolbar-btn" id="btnAddRow" title="Add Row">
            <svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2z"/></svg>
            <span class="btn-label">Row</span>
          </button>
          <button class="toolbar-btn" id="btnAddCol" title="Add Column">
            <svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 2a.5.5 0 0 1 .5.5v5h5a.5.5 0 0 1 0 1h-5v5a.5.5 0 0 1-1 0v-5h-5a.5.5 0 0 1 0-1h5v-5A.5.5 0 0 1 8 2z"/></svg>
            <span class="btn-label">Col</span>
          </button>
        </div>
      </div>
    </div>

    <!-- Search Panel -->
    <div id="searchPanel" class="search-panel hidden">
      <div class="search-container">
        <input type="text" id="searchInput" class="search-input" placeholder="Find..." />
        <span id="searchCount" class="search-count"></span>
        <button class="search-nav-btn" id="searchPrev" title="Previous">&#9650;</button>
        <button class="search-nav-btn" id="searchNext" title="Next">&#9660;</button>
        <button class="search-close-btn" id="searchClose">&times;</button>
      </div>
    </div>

    <!-- Sheet tabs -->
    <div class="sheet-tabs-container">
      <div class="sheet-tabs" id="sheetTabs"></div>
    </div>

    <!-- Table container -->
    <div class="table-wrapper" id="tableWrapper">
      <div class="loading" id="loading">
        <div class="spinner"></div>
        <span>Loading spreadsheet...</span>
      </div>
      <table class="spreadsheet" id="spreadsheet" style="display:none;">
        <thead id="tableHead"></thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>

    <!-- Status bar -->
    <div class="status-bar">
      <div class="status-left">
        <span id="statusInfo">Ready</span>
      </div>
      <div class="status-right">
        <span id="cellCount"></span>
        <span id="selectionInfo"></span>
      </div>
    </div>
  </div>

  <script>
    ${this._getScript()}
  </script>
</body>
</html>`;
  }

  _getStyles() {
    return `
      * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
      }

      :root {
        --bg-primary: #ffffff;
        --bg-secondary: #f3f3f3;
        --bg-header: #e7e6e6;
        --bg-selected: #d6e4f0;
        --bg-hover: #e8f0fe;
        --bg-active-tab: #ffffff;
        --bg-inactive-tab: #e7e6e6;
        --border-color: #d4d4d4;
        --border-strong: #b4b4b4;
        --text-primary: #1e1e1e;
        --text-secondary: #616161;
        --text-header: #333333;
        --accent: #217346;
        --accent-light: #33a867;
        --accent-hover: #185c37;
        --selection-bg: #b4d6a4;
        --selection-border: #217346;
        --toolbar-bg: #f8f8f8;
        --statusbar-bg: #217346;
        --statusbar-text: #ffffff;
        --search-highlight: #fff2cc;
        --search-active: #ff9632;
        --scrollbar-thumb: #c1c1c1;
        --scrollbar-track: #f1f1f1;
        --dropdown-shadow: 0 4px 12px rgba(0,0,0,0.15);
      }

      @media (prefers-color-scheme: dark) {
        :root {
          --bg-primary: #1e1e1e;
          --bg-secondary: #252526;
          --bg-header: #2d2d2d;
          --bg-selected: #264f78;
          --bg-hover: #2a2d2e;
          --bg-active-tab: #1e1e1e;
          --bg-inactive-tab: #2d2d2d;
          --border-color: #3c3c3c;
          --border-strong: #555555;
          --text-primary: #cccccc;
          --text-secondary: #999999;
          --text-header: #cccccc;
          --accent: #217346;
          --accent-light: #33a867;
          --accent-hover: #2ea85c;
          --selection-bg: #264f78;
          --selection-border: #007acc;
          --toolbar-bg: #252526;
          --statusbar-bg: #217346;
          --statusbar-text: #ffffff;
          --search-highlight: #5c4b00;
          --search-active: #ff9632;
          --scrollbar-thumb: #424242;
          --scrollbar-track: #1e1e1e;
          --dropdown-shadow: 0 4px 12px rgba(0,0,0,0.4);
        }
      }

      body.vscode-dark {
        --bg-primary: #1e1e1e;
        --bg-secondary: #252526;
        --bg-header: #2d2d2d;
        --bg-selected: #264f78;
        --bg-hover: #2a2d2e;
        --bg-active-tab: #1e1e1e;
        --bg-inactive-tab: #2d2d2d;
        --border-color: #3c3c3c;
        --border-strong: #555555;
        --text-primary: #cccccc;
        --text-secondary: #999999;
        --text-header: #cccccc;
        --accent: #217346;
        --accent-light: #33a867;
        --accent-hover: #2ea85c;
        --selection-bg: #264f78;
        --selection-border: #007acc;
        --toolbar-bg: #252526;
        --statusbar-bg: #217346;
        --statusbar-text: #ffffff;
        --search-highlight: #5c4b00;
        --search-active: #ff9632;
        --scrollbar-thumb: #424242;
        --scrollbar-track: #1e1e1e;
        --dropdown-shadow: 0 4px 12px rgba(0,0,0,0.4);
      }

      body.vscode-high-contrast {
        --bg-primary: #000000;
        --bg-secondary: #000000;
        --bg-header: #1a1a1a;
        --bg-selected: #0078d4;
        --bg-hover: #1a1a1a;
        --bg-active-tab: #000000;
        --bg-inactive-tab: #1a1a1a;
        --border-color: #6fc3df;
        --border-strong: #6fc3df;
        --text-primary: #ffffff;
        --text-secondary: #d4d4d4;
        --text-header: #ffffff;
        --selection-border: #f38518;
      }

      html, body {
        width: 100%;
        height: 100%;
        overflow: hidden;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
        font-size: 13px;
        background: var(--bg-primary);
        color: var(--text-primary);
      }

      #app {
        display: flex;
        flex-direction: column;
        height: 100vh;
        width: 100%;
      }

      /* ---- Toolbar ---- */
      .toolbar {
        display: flex;
        align-items: center;
        padding: 4px 8px;
        background: var(--toolbar-bg);
        border-bottom: 1px solid var(--border-color);
        gap: 8px;
        min-height: 36px;
        flex-shrink: 0;
      }

      .toolbar-left {
        display: flex;
        align-items: center;
        gap: 8px;
        min-width: 0;
      }

      .file-info {
        display: flex;
        align-items: center;
        gap: 6px;
      }

      .file-icon {
        width: 20px;
        height: 20px;
        flex-shrink: 0;
      }

      .file-name {
        font-size: 12px;
        font-weight: 600;
        color: var(--text-primary);
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 200px;
      }

      .toolbar-center {
        display: flex;
        align-items: center;
        flex: 1;
        gap: 4px;
        min-width: 0;
      }

      .cell-reference {
        flex-shrink: 0;
      }

      .cell-ref-input {
        width: 70px;
        padding: 3px 6px;
        border: 1px solid var(--border-color);
        background: var(--bg-primary);
        color: var(--text-primary);
        font-size: 12px;
        font-family: 'Consolas', 'Courier New', monospace;
        text-align: center;
        border-radius: 2px;
      }

      .formula-bar {
        display: flex;
        align-items: center;
        flex: 1;
        border: 1px solid var(--border-color);
        border-radius: 2px;
        overflow: hidden;
        min-width: 0;
      }

      .fx-label {
        padding: 3px 8px;
        font-style: italic;
        font-weight: bold;
        color: var(--text-secondary);
        border-right: 1px solid var(--border-color);
        background: var(--bg-secondary);
        font-size: 12px;
        user-select: none;
      }

      .formula-input {
        flex: 1;
        padding: 3px 8px;
        border: none;
        background: var(--bg-primary);
        color: var(--text-primary);
        font-size: 12px;
        font-family: 'Consolas', 'Courier New', monospace;
        outline: none;
        min-width: 0;
      }

      .formula-input:focus {
        background: var(--bg-primary);
      }

      .toolbar-right {
        display: flex;
        align-items: center;
        flex-shrink: 0;
      }

      .toolbar-actions {
        display: flex;
        align-items: center;
        gap: 2px;
      }

      .toolbar-btn {
        display: flex;
        align-items: center;
        gap: 4px;
        padding: 4px 8px;
        border: none;
        background: transparent;
        color: var(--text-primary);
        cursor: pointer;
        border-radius: 3px;
        font-size: 11px;
        transition: background 0.15s;
      }

      .toolbar-btn:hover {
        background: var(--bg-hover);
      }

      .toolbar-btn svg {
        width: 14px;
        height: 14px;
      }

      .btn-label {
        font-size: 11px;
      }

      /* Dropdown */
      .dropdown {
        position: relative;
      }

      .dropdown-menu {
        display: none;
        position: absolute;
        top: 100%;
        right: 0;
        background: var(--bg-primary);
        border: 1px solid var(--border-color);
        border-radius: 4px;
        box-shadow: var(--dropdown-shadow);
        z-index: 100;
        min-width: 160px;
        padding: 4px 0;
      }

      .dropdown-menu.show {
        display: block;
      }

      .dropdown-item {
        display: block;
        width: 100%;
        padding: 6px 16px;
        border: none;
        background: transparent;
        color: var(--text-primary);
        font-size: 12px;
        text-align: left;
        cursor: pointer;
        transition: background 0.15s;
      }

      .dropdown-item:hover {
        background: var(--bg-hover);
      }

      /* ---- Search Panel ---- */
      .search-panel {
        padding: 4px 8px;
        background: var(--toolbar-bg);
        border-bottom: 1px solid var(--border-color);
        flex-shrink: 0;
      }

      .search-panel.hidden {
        display: none;
      }

      .search-container {
        display: flex;
        align-items: center;
        gap: 6px;
      }

      .search-input {
        flex: 1;
        max-width: 300px;
        padding: 3px 8px;
        border: 1px solid var(--border-color);
        background: var(--bg-primary);
        color: var(--text-primary);
        font-size: 12px;
        border-radius: 2px;
        outline: none;
      }

      .search-input:focus {
        border-color: var(--accent);
      }

      .search-count {
        font-size: 11px;
        color: var(--text-secondary);
        min-width: 60px;
      }

      .search-nav-btn, .search-close-btn {
        padding: 2px 6px;
        border: none;
        background: transparent;
        color: var(--text-primary);
        cursor: pointer;
        border-radius: 3px;
        font-size: 12px;
      }

      .search-nav-btn:hover, .search-close-btn:hover {
        background: var(--bg-hover);
      }

      /* ---- Sheet Tabs ---- */
      .sheet-tabs-container {
        display: flex;
        align-items: center;
        border-bottom: 1px solid var(--border-color);
        background: var(--bg-secondary);
        flex-shrink: 0;
        overflow-x: auto;
      }

      .sheet-tabs-container::-webkit-scrollbar {
        height: 3px;
      }

      .sheet-tabs-container::-webkit-scrollbar-thumb {
        background: var(--scrollbar-thumb);
        border-radius: 3px;
      }

      .sheet-tabs {
        display: flex;
        align-items: flex-end;
        gap: 0;
        padding: 0 4px;
      }

      .sheet-tab {
        padding: 6px 16px;
        border: 1px solid transparent;
        border-bottom: none;
        background: var(--bg-inactive-tab);
        color: var(--text-secondary);
        cursor: pointer;
        font-size: 12px;
        white-space: nowrap;
        transition: all 0.15s;
        border-radius: 4px 4px 0 0;
        position: relative;
      }

      .sheet-tab:hover {
        background: var(--bg-hover);
        color: var(--text-primary);
      }

      .sheet-tab.active {
        background: var(--bg-active-tab);
        color: var(--text-primary);
        border-color: var(--border-color);
        font-weight: 600;
      }

      .sheet-tab.active::after {
        content: '';
        position: absolute;
        bottom: -1px;
        left: 0;
        right: 0;
        height: 2px;
        background: var(--accent);
      }

      /* ---- Table ---- */
      .table-wrapper {
        flex: 1;
        overflow: auto;
        position: relative;
      }

      .table-wrapper::-webkit-scrollbar {
        width: 12px;
        height: 12px;
      }

      .table-wrapper::-webkit-scrollbar-thumb {
        background: var(--scrollbar-thumb);
        border-radius: 6px;
        border: 3px solid transparent;
        background-clip: padding-box;
      }

      .table-wrapper::-webkit-scrollbar-track {
        background: var(--scrollbar-track);
      }

      .table-wrapper::-webkit-scrollbar-corner {
        background: var(--scrollbar-track);
      }

      .spreadsheet {
        border-collapse: collapse;
        table-layout: fixed;
        min-width: 100%;
      }

      .spreadsheet th,
      .spreadsheet td {
        border: 1px solid var(--border-color);
        padding: 0;
        height: 22px;
        min-width: 80px;
        max-width: 300px;
        font-size: 12px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        position: relative;
      }

      .spreadsheet th {
        background: var(--bg-header);
        color: var(--text-header);
        font-weight: 600;
        text-align: center;
        padding: 0 4px;
        position: sticky;
        user-select: none;
        cursor: default;
      }

      .spreadsheet thead th {
        top: 0;
        z-index: 10;
      }

      .spreadsheet th.row-header {
        width: 50px;
        min-width: 50px;
        max-width: 50px;
        left: 0;
        z-index: 11;
        text-align: center;
        color: var(--text-secondary);
        font-weight: normal;
        font-size: 11px;
      }

      .spreadsheet thead th.corner-header {
        left: 0;
        z-index: 12;
        width: 50px;
        min-width: 50px;
        max-width: 50px;
      }

      .spreadsheet td {
        padding: 0 4px;
        cursor: cell;
        color: var(--text-primary);
        background: var(--bg-primary);
      }

      .spreadsheet td:hover {
        background: var(--bg-hover);
      }

      .spreadsheet td.selected {
        outline: 2px solid var(--selection-border);
        outline-offset: -1px;
        background: var(--bg-selected);
        z-index: 5;
      }

      .spreadsheet td.in-range {
        background: var(--selection-bg);
      }

      .spreadsheet td.search-match {
        background: var(--search-highlight) !important;
      }

      .spreadsheet td.search-active {
        background: var(--search-active) !important;
        outline: 2px solid var(--search-active);
      }

      .spreadsheet td.number-cell {
        text-align: right;
      }

      .spreadsheet td.editing {
        padding: 0;
        overflow: visible;
      }

      .spreadsheet td.editing input {
        width: 100%;
        height: 100%;
        border: none;
        padding: 0 4px;
        font-size: 12px;
        font-family: inherit;
        background: var(--bg-primary);
        color: var(--text-primary);
        outline: 2px solid var(--selection-border);
        outline-offset: -1px;
        box-sizing: border-box;
      }

      .spreadsheet th.col-selected {
        background: var(--bg-selected);
      }

      .spreadsheet th.row-selected {
        background: var(--bg-selected);
      }

      /* Column resize handle */
      .col-resize-handle {
        position: absolute;
        right: 0;
        top: 0;
        bottom: 0;
        width: 4px;
        cursor: col-resize;
        user-select: none;
      }

      .col-resize-handle:hover {
        background: var(--accent);
      }

      /* ---- Loading ---- */
      .loading {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 12px;
        height: 100%;
        color: var(--text-secondary);
        font-size: 14px;
      }

      .spinner {
        width: 24px;
        height: 24px;
        border: 3px solid var(--border-color);
        border-top-color: var(--accent);
        border-radius: 50%;
        animation: spin 0.8s linear infinite;
      }

      @keyframes spin {
        to { transform: rotate(360deg); }
      }

      /* ---- Status Bar ---- */
      .status-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 2px 10px;
        background: var(--statusbar-bg);
        color: var(--statusbar-text);
        font-size: 11px;
        flex-shrink: 0;
        min-height: 22px;
      }

      .status-left, .status-right {
        display: flex;
        align-items: center;
        gap: 16px;
      }

      /* Context menu */
      .context-menu {
        position: fixed;
        background: var(--bg-primary);
        border: 1px solid var(--border-color);
        border-radius: 4px;
        box-shadow: var(--dropdown-shadow);
        z-index: 1000;
        min-width: 180px;
        padding: 4px 0;
      }

      .context-menu-item {
        display: block;
        width: 100%;
        padding: 6px 16px;
        border: none;
        background: transparent;
        color: var(--text-primary);
        font-size: 12px;
        text-align: left;
        cursor: pointer;
        transition: background 0.12s;
      }

      .context-menu-item:hover {
        background: var(--bg-hover);
      }

      .context-menu-separator {
        height: 1px;
        background: var(--border-color);
        margin: 4px 0;
      }
    `;
  }

  _getScript() {
    return `
      (function() {
        const vscode = acquireVsCodeApi();

        let state = {
          sheetNames: [],
          activeSheet: '',
          sheetData: null,
          selectedCell: null,
          editingCell: null,
          searchMatches: [],
          currentMatchIndex: -1,
          isResizing: false,
        };

        // DOM Elements
        const $ = (sel) => document.querySelector(sel);
        const $$ = (sel) => document.querySelectorAll(sel);

        const elements = {
          sheetTabs: $('#sheetTabs'),
          tableHead: $('#tableHead'),
          tableBody: $('#tableBody'),
          spreadsheet: $('#spreadsheet'),
          loading: $('#loading'),
          cellRef: $('#cellRef'),
          formulaBar: $('#formulaBar'),
          statusInfo: $('#statusInfo'),
          cellCount: $('#cellCount'),
          selectionInfo: $('#selectionInfo'),
          searchPanel: $('#searchPanel'),
          searchInput: $('#searchInput'),
          searchCount: $('#searchCount'),
          fileName: $('#fileName'),
          tableWrapper: $('#tableWrapper'),
        };

        // Initialize
        vscode.postMessage({ type: 'ready' });

        // Listen for messages from extension
        window.addEventListener('message', (event) => {
          const msg = event.data;
          switch (msg.type) {
            case 'init':
              state.sheetNames = msg.sheetNames;
              state.activeSheet = msg.activeSheet;
              state.sheetData = msg.sheetData;
              renderAll();
              break;
            case 'sheetData':
              state.activeSheet = msg.sheetName;
              state.sheetData = msg.sheetData;
              state.selectedCell = null;
              renderTable();
              updateStatus();
              break;
          }
        });

        function renderAll() {
          renderSheetTabs();
          renderTable();
          updateStatus();
          elements.loading.style.display = 'none';
          elements.spreadsheet.style.display = 'table';
        }

        function renderSheetTabs() {
          elements.sheetTabs.innerHTML = '';
          state.sheetNames.forEach((name) => {
            const tab = document.createElement('button');
            tab.className = 'sheet-tab' + (name === state.activeSheet ? ' active' : '');
            tab.textContent = name;
            tab.onclick = () => switchSheet(name);
            elements.sheetTabs.appendChild(tab);
          });
        }

        function switchSheet(name) {
          if (name === state.activeSheet) return;
          state.activeSheet = name;
          state.selectedCell = null;
          state.editingCell = null;
          // Update tab styles
          $$('.sheet-tab').forEach((tab) => {
            tab.classList.toggle('active', tab.textContent === name);
          });
          vscode.postMessage({ type: 'requestSheetData', sheetName: name });
          elements.statusInfo.textContent = 'Loading...';
        }

        function renderTable() {
          const { data, range } = state.sheetData;
          if (!data || data.length === 0) {
            elements.tableHead.innerHTML = '<tr><th class="corner-header"></th></tr>';
            elements.tableBody.innerHTML = '<tr><td colspan="100" style="text-align:center;padding:20px;color:var(--text-secondary)">Empty sheet</td></tr>';
            return;
          }

          const numCols = range.endCol - range.startCol + 1;

          // Render header
          let headHtml = '<tr><th class="corner-header"></th>';
          for (let c = 0; c < numCols; c++) {
            const colLetter = getColLetter(range.startCol + c);
            headHtml += '<th>' + colLetter + '<div class="col-resize-handle" data-col="' + c + '"></div></th>';
          }
          headHtml += '</tr>';
          elements.tableHead.innerHTML = headHtml;

          // Render body
          let bodyHtml = '';
          for (let r = 0; r < data.length; r++) {
            const rowNum = range.startRow + r + 1;
            bodyHtml += '<tr><th class="row-header">' + rowNum + '</th>';
            for (let c = 0; c < data[r].length; c++) {
              const cell = data[r][c];
              const isNumber = cell.type === 'n';
              const displayVal = escapeHtml(cell.formatted || String(cell.value || ''));
              bodyHtml += '<td data-row="' + r + '" data-col="' + c + '" data-address="' + cell.address + '"'
                + (isNumber ? ' class="number-cell"' : '')
                + ' title="' + escapeHtml(String(cell.value || '')) + '">'
                + displayVal + '</td>';
            }
            bodyHtml += '</tr>';
          }
          elements.tableBody.innerHTML = bodyHtml;

          // Attach click handlers
          attachCellHandlers();
          attachResizeHandlers();
        }

        function attachCellHandlers() {
          elements.tableBody.querySelectorAll('td').forEach((td) => {
            td.addEventListener('click', (e) => {
              if (state.editingCell) return;
              selectCell(td);
            });
            td.addEventListener('dblclick', (e) => {
              startEditing(td);
            });
          });
        }

        function selectCell(td) {
          // Remove previous selection
          $$('td.selected').forEach((el) => el.classList.remove('selected'));
          $$('th.col-selected').forEach((el) => el.classList.remove('col-selected'));
          $$('th.row-selected').forEach((el) => el.classList.remove('row-selected'));

          td.classList.add('selected');
          state.selectedCell = td;

          const address = td.dataset.address;
          const row = parseInt(td.dataset.row);
          const col = parseInt(td.dataset.col);

          elements.cellRef.value = address;

          // Update formula bar
          const cellData = state.sheetData.data[row][col];
          elements.formulaBar.value = cellData.value !== undefined && cellData.value !== null
            ? String(cellData.value) : '';

          // Highlight column and row headers
          const headerCells = elements.tableHead.querySelectorAll('th');
          if (headerCells[col + 1]) headerCells[col + 1].classList.add('col-selected');
          const rowHeader = td.parentElement.querySelector('th.row-header');
          if (rowHeader) rowHeader.classList.add('row-selected');

          updateSelectionInfo(address, cellData);
        }

        function startEditing(td) {
          if (state.editingCell) {
            finishEditing();
          }

          const row = parseInt(td.dataset.row);
          const col = parseInt(td.dataset.col);
          const cellData = state.sheetData.data[row][col];
          const currentValue = cellData.value !== undefined && cellData.value !== null
            ? String(cellData.value) : '';

          td.classList.add('editing');
          const input = document.createElement('input');
          input.type = 'text';
          input.value = currentValue;
          td.textContent = '';
          td.appendChild(input);
          input.focus();
          input.select();

          state.editingCell = { td, row, col, originalValue: currentValue };

          input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              finishEditing();
              // Move to next row
              const nextRow = row + 1;
              if (nextRow < state.sheetData.data.length) {
                const nextTd = elements.tableBody.querySelector(
                  'td[data-row="' + nextRow + '"][data-col="' + col + '"]'
                );
                if (nextTd) selectCell(nextTd);
              }
            } else if (e.key === 'Tab') {
              e.preventDefault();
              finishEditing();
              // Move to next column
              const nextCol = e.shiftKey ? col - 1 : col + 1;
              if (nextCol >= 0 && nextCol < state.sheetData.data[row].length) {
                const nextTd = elements.tableBody.querySelector(
                  'td[data-row="' + row + '"][data-col="' + nextCol + '"]'
                );
                if (nextTd) selectCell(nextTd);
              }
            } else if (e.key === 'Escape') {
              cancelEditing();
            }
          });

          input.addEventListener('blur', () => {
            if (state.editingCell) {
              finishEditing();
            }
          });

          // Update formula bar while editing
          input.addEventListener('input', () => {
            elements.formulaBar.value = input.value;
          });
        }

        function finishEditing() {
          if (!state.editingCell) return;

          const { td, row, col, originalValue } = state.editingCell;
          const input = td.querySelector('input');
          const newValue = input ? input.value : originalValue;

          td.classList.remove('editing');

          // Update display
          const isNumber = !isNaN(Number(newValue)) && newValue.trim() !== '';
          td.textContent = newValue;
          td.className = isNumber ? 'number-cell' : '';

          // Update local data
          state.sheetData.data[row][col].value = newValue;
          state.sheetData.data[row][col].formatted = newValue;
          state.sheetData.data[row][col].type = isNumber ? 'n' : 's';
          td.setAttribute('title', newValue);

          // Send edit to extension if value changed
          if (newValue !== originalValue) {
            const address = td.dataset.address;
            vscode.postMessage({
              type: 'edit',
              sheetName: state.activeSheet,
              address: address,
              value: newValue,
            });
            elements.statusInfo.textContent = 'Cell ' + address + ' modified';
          }

          state.editingCell = null;
          selectCell(td);
        }

        function cancelEditing() {
          if (!state.editingCell) return;

          const { td, row, col, originalValue } = state.editingCell;
          td.classList.remove('editing');
          td.textContent = originalValue;

          state.editingCell = null;
          selectCell(td);
        }

        // Formula bar editing
        elements.formulaBar.addEventListener('keydown', (e) => {
          if (e.key === 'Enter' && state.selectedCell) {
            e.preventDefault();
            const newValue = elements.formulaBar.value;
            const row = parseInt(state.selectedCell.dataset.row);
            const col = parseInt(state.selectedCell.dataset.col);
            const address = state.selectedCell.dataset.address;

            // Update display
            const isNumber = !isNaN(Number(newValue)) && newValue.trim() !== '';
            state.selectedCell.textContent = newValue;
            state.selectedCell.className = isNumber ? 'number-cell' : '';

            // Update local data
            state.sheetData.data[row][col].value = newValue;
            state.sheetData.data[row][col].formatted = newValue;
            state.sheetData.data[row][col].type = isNumber ? 'n' : 's';

            // Send edit
            vscode.postMessage({
              type: 'edit',
              sheetName: state.activeSheet,
              address: address,
              value: newValue,
            });

            elements.statusInfo.textContent = 'Cell ' + address + ' modified';
            state.selectedCell.focus();
          }
        });

        // Column resize
        function attachResizeHandlers() {
          $$('.col-resize-handle').forEach((handle) => {
            handle.addEventListener('mousedown', (e) => {
              e.preventDefault();
              e.stopPropagation();

              const colIndex = parseInt(handle.dataset.col);
              const th = handle.parentElement;
              const startX = e.clientX;
              const startWidth = th.offsetWidth;
              state.isResizing = true;

              const onMouseMove = (me) => {
                const diff = me.clientX - startX;
                const newWidth = Math.max(40, startWidth + diff);
                th.style.width = newWidth + 'px';
                th.style.minWidth = newWidth + 'px';
                // Resize all cells in that column
                elements.tableBody.querySelectorAll('td[data-col="' + colIndex + '"]').forEach((td) => {
                  td.style.width = newWidth + 'px';
                  td.style.minWidth = newWidth + 'px';
                });
              };

              const onMouseUp = () => {
                state.isResizing = false;
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
              };

              document.addEventListener('mousemove', onMouseMove);
              document.addEventListener('mouseup', onMouseUp);
            });
          });
        }

        // Search
        $('#btnSearch').addEventListener('click', toggleSearch);

        function toggleSearch() {
          elements.searchPanel.classList.toggle('hidden');
          if (!elements.searchPanel.classList.contains('hidden')) {
            elements.searchInput.focus();
          } else {
            clearSearch();
          }
        }

        elements.searchInput.addEventListener('input', performSearch);
        $('#searchClose').addEventListener('click', () => {
          elements.searchPanel.classList.add('hidden');
          clearSearch();
        });

        $('#searchPrev').addEventListener('click', () => navigateSearch(-1));
        $('#searchNext').addEventListener('click', () => navigateSearch(1));

        function performSearch() {
          clearSearchHighlights();
          const query = elements.searchInput.value.toLowerCase().trim();
          if (!query) {
            elements.searchCount.textContent = '';
            state.searchMatches = [];
            state.currentMatchIndex = -1;
            return;
          }

          state.searchMatches = [];
          elements.tableBody.querySelectorAll('td').forEach((td) => {
            const text = td.textContent.toLowerCase();
            if (text.includes(query)) {
              td.classList.add('search-match');
              state.searchMatches.push(td);
            }
          });

          if (state.searchMatches.length > 0) {
            state.currentMatchIndex = 0;
            highlightCurrentMatch();
          }

          elements.searchCount.textContent = state.searchMatches.length + ' found';
        }

        function navigateSearch(direction) {
          if (state.searchMatches.length === 0) return;
          state.searchMatches[state.currentMatchIndex]?.classList.remove('search-active');
          state.currentMatchIndex = (state.currentMatchIndex + direction + state.searchMatches.length) % state.searchMatches.length;
          highlightCurrentMatch();
        }

        function highlightCurrentMatch() {
          const match = state.searchMatches[state.currentMatchIndex];
          if (match) {
            match.classList.add('search-active');
            match.scrollIntoView({ behavior: 'smooth', block: 'center' });
            selectCell(match);
          }
          elements.searchCount.textContent = (state.currentMatchIndex + 1) + '/' + state.searchMatches.length;
        }

        function clearSearch() {
          clearSearchHighlights();
          elements.searchInput.value = '';
          elements.searchCount.textContent = '';
          state.searchMatches = [];
          state.currentMatchIndex = -1;
        }

        function clearSearchHighlights() {
          $$('td.search-match').forEach((td) => td.classList.remove('search-match'));
          $$('td.search-active').forEach((td) => td.classList.remove('search-active'));
        }

        // Export
        $('#btnExport').addEventListener('click', () => {
          $('#exportMenu').classList.toggle('show');
        });

        $('#exportCSV').addEventListener('click', () => {
          vscode.postMessage({ type: 'exportCSV', sheetName: state.activeSheet });
          $('#exportMenu').classList.remove('show');
        });

        $('#exportJSON').addEventListener('click', () => {
          vscode.postMessage({ type: 'exportJSON', sheetName: state.activeSheet });
          $('#exportMenu').classList.remove('show');
        });

        // Add Row / Column
        $('#btnAddRow').addEventListener('click', () => {
          const afterRow = state.selectedCell
            ? parseInt(state.selectedCell.dataset.row)
            : (state.sheetData?.data?.length || 1) - 1;
          vscode.postMessage({ type: 'addRow', sheetName: state.activeSheet, afterRow });
        });

        $('#btnAddCol').addEventListener('click', () => {
          const afterCol = state.selectedCell
            ? parseInt(state.selectedCell.dataset.col)
            : (state.sheetData?.data?.[0]?.length || 1) - 1;
          vscode.postMessage({ type: 'addColumn', sheetName: state.activeSheet, afterCol });
        });

        // Close dropdown on outside click
        document.addEventListener('click', (e) => {
          if (!e.target.closest('#exportDropdown')) {
            $('#exportMenu').classList.remove('show');
          }
        });

        // Keyboard shortcuts
        document.addEventListener('keydown', (e) => {
          // Ctrl+F: Search
          if ((e.ctrlKey || e.metaKey) && e.key === 'f') {
            e.preventDefault();
            toggleSearch();
            return;
          }

          // Escape: close search or cancel editing
          if (e.key === 'Escape') {
            if (!elements.searchPanel.classList.contains('hidden')) {
              elements.searchPanel.classList.add('hidden');
              clearSearch();
              return;
            }
            if (state.editingCell) {
              cancelEditing();
              return;
            }
          }

          // If editing, don't handle arrow keys
          if (state.editingCell) return;

          // Arrow key navigation
          if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && state.selectedCell) {
            e.preventDefault();
            const row = parseInt(state.selectedCell.dataset.row);
            const col = parseInt(state.selectedCell.dataset.col);
            let newRow = row, newCol = col;

            switch (e.key) {
              case 'ArrowUp': newRow = Math.max(0, row - 1); break;
              case 'ArrowDown': newRow = Math.min(state.sheetData.data.length - 1, row + 1); break;
              case 'ArrowLeft': newCol = Math.max(0, col - 1); break;
              case 'ArrowRight': newCol = Math.min(state.sheetData.data[0].length - 1, col + 1); break;
            }

            const nextTd = elements.tableBody.querySelector(
              'td[data-row="' + newRow + '"][data-col="' + newCol + '"]'
            );
            if (nextTd) {
              selectCell(nextTd);
              nextTd.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'nearest' });
            }
          }

          // Enter or F2 to start editing
          if ((e.key === 'Enter' || e.key === 'F2') && state.selectedCell && !state.editingCell) {
            e.preventDefault();
            startEditing(state.selectedCell);
          }

          // Delete to clear cell
          if (e.key === 'Delete' && state.selectedCell && !state.editingCell) {
            e.preventDefault();
            const row = parseInt(state.selectedCell.dataset.row);
            const col = parseInt(state.selectedCell.dataset.col);
            const address = state.selectedCell.dataset.address;

            state.selectedCell.textContent = '';
            state.sheetData.data[row][col].value = '';
            state.sheetData.data[row][col].formatted = '';

            vscode.postMessage({
              type: 'edit',
              sheetName: state.activeSheet,
              address: address,
              value: '',
            });
            elements.statusInfo.textContent = 'Cell ' + address + ' cleared';
          }

          // Type to start editing (alphanumeric)
          if (state.selectedCell && !state.editingCell && !e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
            startEditing(state.selectedCell);
            // The character will be typed into the input
          }
        });

        // Context menu
        elements.tableBody.addEventListener('contextmenu', (e) => {
          e.preventDefault();
          const td = e.target.closest('td');
          if (!td) return;

          selectCell(td);
          showContextMenu(e.clientX, e.clientY, td);
        });

        function showContextMenu(x, y, td) {
          // Remove existing context menu
          const existing = document.querySelector('.context-menu');
          if (existing) existing.remove();

          const menu = document.createElement('div');
          menu.className = 'context-menu';
          menu.style.left = x + 'px';
          menu.style.top = y + 'px';

          const row = parseInt(td.dataset.row);
          const col = parseInt(td.dataset.col);

          menu.innerHTML = [
            '<button class="context-menu-item" data-action="copy">Copy</button>',
            '<button class="context-menu-item" data-action="paste">Paste</button>',
            '<div class="context-menu-separator"></div>',
            '<button class="context-menu-item" data-action="insert-row-above">Insert Row Above</button>',
            '<button class="context-menu-item" data-action="insert-row-below">Insert Row Below</button>',
            '<button class="context-menu-item" data-action="insert-col-left">Insert Column Left</button>',
            '<button class="context-menu-item" data-action="insert-col-right">Insert Column Right</button>',
            '<div class="context-menu-separator"></div>',
            '<button class="context-menu-item" data-action="clear">Clear Cell</button>',
          ].join('');

          document.body.appendChild(menu);

          // Position adjustments
          const rect = menu.getBoundingClientRect();
          if (rect.right > window.innerWidth) menu.style.left = (x - rect.width) + 'px';
          if (rect.bottom > window.innerHeight) menu.style.top = (y - rect.height) + 'px';

          menu.addEventListener('click', (e) => {
            const action = e.target.dataset.action;
            if (!action) return;

            switch (action) {
              case 'copy':
                navigator.clipboard.writeText(td.textContent);
                break;
              case 'paste':
                navigator.clipboard.readText().then((text) => {
                  const address = td.dataset.address;
                  td.textContent = text;
                  state.sheetData.data[row][col].value = text;
                  state.sheetData.data[row][col].formatted = text;
                  vscode.postMessage({
                    type: 'edit',
                    sheetName: state.activeSheet,
                    address: address,
                    value: text,
                  });
                });
                break;
              case 'insert-row-above':
                vscode.postMessage({ type: 'addRow', sheetName: state.activeSheet, afterRow: row - 1 });
                break;
              case 'insert-row-below':
                vscode.postMessage({ type: 'addRow', sheetName: state.activeSheet, afterRow: row });
                break;
              case 'insert-col-left':
                vscode.postMessage({ type: 'addColumn', sheetName: state.activeSheet, afterCol: col - 1 });
                break;
              case 'insert-col-right':
                vscode.postMessage({ type: 'addColumn', sheetName: state.activeSheet, afterCol: col });
                break;
              case 'clear':
                td.textContent = '';
                state.sheetData.data[row][col].value = '';
                state.sheetData.data[row][col].formatted = '';
                vscode.postMessage({
                  type: 'edit',
                  sheetName: state.activeSheet,
                  address: td.dataset.address,
                  value: '',
                });
                break;
            }
            menu.remove();
          });

          // Close on outside click
          setTimeout(() => {
            document.addEventListener('click', function closeMenu() {
              menu.remove();
              document.removeEventListener('click', closeMenu);
            });
          }, 0);
        }

        function updateStatus() {
          if (!state.sheetData) return;
          const { data, range } = state.sheetData;
          const totalRows = data.length;
          const totalCols = data[0] ? data[0].length : 0;
          elements.cellCount.textContent = totalRows + ' rows × ' + totalCols + ' cols';
          elements.statusInfo.textContent = 'Sheet: ' + state.activeSheet;
        }

        function updateSelectionInfo(address, cellData) {
          let info = address;
          if (cellData.type === 'n') {
            info += ' | Number: ' + cellData.value;
          }
          elements.selectionInfo.textContent = info;
        }

        function getColLetter(colIndex) {
          let result = '';
          let n = colIndex;
          while (n >= 0) {
            result = String.fromCharCode(65 + (n % 26)) + result;
            n = Math.floor(n / 26) - 1;
          }
          return result;
        }

        function escapeHtml(str) {
          const div = document.createElement('div');
          div.textContent = str;
          return div.innerHTML;
        }
      })();
    `;
  }
}

module.exports = ExcelEditorProvider;
