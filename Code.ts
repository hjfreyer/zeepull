
/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

var SIDEBAR_TITLE = 'ZeeMaps Update';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Upload CSVs', 'showSidebar')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Returns the value in the active cell.
 *
 * @return {String} The value of the active cell.
 */
function getActiveValue() {
  // Retrieve and return the information requested by the sidebar.
  var cell = SpreadsheetApp.getActiveSheet().getActiveCell();
  return cell.getValue();
}

type FileMeta = {
  filename: string
  contents: string
}

/**
 * Replaces the active cell value with the given value.
 *
 * @param {Number} value A reference number to replace with.
 */
function doUpdates(csvFiles: FileMeta[]) {
  const numberToFilenames = indexFilesByValuesInColumn(csvFiles, "number");

  const sheetNumberColumn = getColumnDataByName("number");

  const newColumn = sheetNumberColumn.map((number, rowIdx) => {
    if (rowIdx === 0) {
      return 'In Files'
    }
    const numberStr = '' + number;
    return (numberToFilenames[numberStr] || []).join(",");
  });

  addColumnLeft(newColumn);
}

/**
 * Parse a collection of CSV files with a column named `columnName`, and build
 * an index from values appearing in that column to the set of files it appears
 * in.
 *
 * Throws an exception if any of the CSV files don't have a matching column.
 */
function indexFilesByValuesInColumn(csvFiles: FileMeta[], columnName: string): Record<string, string[]> {
  const res : Record<string, string[]> = {};
  for (const csvFile of csvFiles) {
    const csv = Utilities.parseCsv(csvFile.contents);
    const numberColumnIdx = csv[0].indexOf(columnName);
    if (numberColumnIdx === -1) {
      throw new Error(`No column named '${columnName}' in uploaded file ${csvFile.filename}`);
    }

    for (const row of csv.slice(1)) {
      const number = row[numberColumnIdx];
      if (!(number in res)) {
        res[number] = [];
      }
      res[number].push(csvFile.filename);
    }
  }
  return res;
}

function getColumnDataByName(columnName: string): unknown[] {
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = currentSheet.getDataRange().getValues();

  const colIdx = data[0].indexOf(columnName);
  if (colIdx === -1) {
    throw new Error(`No column named '${columnName}' in the current sheet`);
  }

  return data.map(row => row[colIdx]);
}

function addColumnLeft(columnData: string[]): void {
  const currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  currentSheet.insertColumns(1);
  const newColumnRange = currentSheet.getRange(1, 1, columnData.length, 1);
  newColumnRange.setValues(columnData.map(value => [value]));
  currentSheet.autoResizeColumn(1);
}
