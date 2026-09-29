# Excel Viewer Pro

[![VS Code Marketplace](https://img.shields.io/visual-studio-marketplace/v/youcode.excel-viewer-pro)](https://marketplace.visualstudio.com/items?itemName=youcode.excel-viewer-pro)
[![Downloads](https://img.shields.io/visual-studio-marketplace/d/youcode.excel-viewer-pro)](https://marketplace.visualstudio.com/items?itemName=youcode.excel-viewer-pro)

**View and edit Excel files directly in VS Code** with a beautiful, Excel-like spreadsheet interface.

## ✨ Features

### 📊 View Excel Files
- Open `.xlsx`, `.xls`, and `.xlsm` files directly in VS Code
- Beautiful spreadsheet-style grid with Excel-like appearance
- Automatic number formatting and type detection

### 📑 Multiple Sheet Support
- Navigate between worksheets with tabs
- Quick switching between sheets

### ✏️ Cell Editing
- Double-click any cell to edit
- Formula bar for editing cell values
- Supports Enter, Tab, and Escape keyboard shortcuts
- Auto-detection of number vs text values

### 🔍 Search
- Press `Ctrl+F` to search within the current sheet
- Navigate between search results with arrow buttons
- Highlights all matches with active match indicator

### 📤 Export
- Export current sheet as **CSV**
- Export current sheet as **JSON**
- Choose save location

### ⌨️ Keyboard Shortcuts
| Shortcut | Action |
|----------|--------|
| `Arrow Keys` | Navigate between cells |
| `Enter` / `F2` | Start editing selected cell |
| `Tab` | Move to next cell |
| `Shift+Tab` | Move to previous cell |
| `Escape` | Cancel editing / Close search |
| `Delete` | Clear selected cell |
| `Ctrl+F` | Open search |

### 🎨 Theme Support
- Automatically adapts to VS Code themes (Light, Dark, High Contrast)
- Respects your VS Code color preferences

### ➕ Row & Column Management
- Insert rows above or below
- Insert columns left or right
- Right-click context menu for quick actions

### 📐 Column Resize
- Drag column borders to resize
- Customizable column widths

## 📦 Installation

1. Open VS Code
2. Go to Extensions (`Ctrl+Shift+X`)
3. Search for "Excel Viewer Pro"
4. Click **Install**

Or install from the command line:
```bash
code --install-extension youcode.excel-viewer-pro
```

## 🚀 Usage

Simply open any `.xlsx`, `.xls`, or `.xlsm` file in VS Code. The extension will automatically display it in a spreadsheet view.

You can also use the command palette (`Ctrl+Shift+P`) and search for:
- **Excel Viewer Pro: Open Excel File** - Open a file dialog to select an Excel file
- **Excel Viewer Pro: Export Current Sheet as CSV**
- **Excel Viewer Pro: Export Current Sheet as JSON**

## 🛠️ Requirements

- VS Code 1.80.0 or higher
- No additional software required

## 📋 Supported File Formats

| Format | Extension | Support |
|--------|-----------|---------|
| Excel Workbook | `.xlsx` | ✅ Full |
| Excel 97-2003 | `.xls` | ✅ Full |
| Excel Macro-Enabled | `.xlsm` | ✅ Full |

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This extension is licensed under the [MIT License](LICENSE).

## 🐛 Known Issues

- Very large files (>100MB) may take a moment to load
- Complex formulas are displayed as their calculated values

## 📝 Release Notes

### 1.0.0
- Initial release
- Excel file viewing with spreadsheet grid
- Cell editing support
- Multiple sheet tabs
- Search functionality
- CSV and JSON export
- Dark/Light theme support
- Column resize
- Context menu actions
