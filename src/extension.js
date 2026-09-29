const vscode = require('vscode');
const ExcelEditorProvider = require('./excelEditorProvider');

/**
 * @param {vscode.ExtensionContext} context
 */
function activate(context) {
  // Register the custom editor provider for Excel files
  const provider = new ExcelEditorProvider(context);

  context.subscriptions.push(
    vscode.window.registerCustomEditorProvider(
      'excelViewerPro.spreadsheet',
      provider,
      {
        webviewOptions: {
          retainContextWhenHidden: true,
        },
        supportsMultipleEditorsPerDocument: false,
      }
    )
  );

  // Register command: Open Excel File
  context.subscriptions.push(
    vscode.commands.registerCommand('excelViewerPro.openFile', async () => {
      const fileUri = await vscode.window.showOpenDialog({
        canSelectMany: false,
        filters: {
          'Excel Files': ['xlsx', 'xls', 'xlsm'],
        },
      });
      if (fileUri && fileUri[0]) {
        await vscode.commands.executeCommand(
          'vscode.openWith',
          fileUri[0],
          'excelViewerPro.spreadsheet'
        );
      }
    })
  );

  // Register command: Export as CSV
  context.subscriptions.push(
    vscode.commands.registerCommand('excelViewerPro.exportCSV', () => {
      vscode.window.showInformationMessage(
        'Use the export button in the Excel Viewer toolbar to export as CSV.'
      );
    })
  );

  // Register command: Export as JSON
  context.subscriptions.push(
    vscode.commands.registerCommand('excelViewerPro.exportJSON', () => {
      vscode.window.showInformationMessage(
        'Use the export button in the Excel Viewer toolbar to export as JSON.'
      );
    })
  );
}

function deactivate() {}

module.exports = {
  activate,
  deactivate,
};
