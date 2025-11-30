/******************************************************************
 * Creates a static snapshot of the active sheet (tab).
 * It copies the active sheet, renames the copy with a timestamp,
 * and then converts all formulas in the copy to their final 
 * calculated values (text/numbers).
 ******************************************************************/
function createSnapshot() {
  //---------------------------------------------------------------
  // 1. Get the active spreadsheet and the currently active sheet
  //---------------------------------------------------------------
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = spreadsheet.getActiveSheet();
  
  //---------------------------------------------------------------
  // 2. Create the new name for the snapshot sheet
  // if the new sheetname exists, delete it first
  //---------------------------------------------------------------
  const timestamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), "yyyyMMdd");
  const newSheetName = activeSheet.getName() + "_" + timestamp;
  const existingSheet = spreadsheet.getSheetByName(newSheetName);
  if (existingSheet) {
    // If a sheet with the exact snapshot name already exists, delete it.
    spreadsheet.deleteSheet(existingSheet);
    Logger.log(`Deleted existing sheet: ${newSheetName}`);
  }
  
  //---------------------------------------------------------------
  // 3. Duplicate the active sheet and place it at the end
  //---------------------------------------------------------------
  const newSheet = activeSheet.copyTo(spreadsheet);
  newSheet.setName(newSheetName);
  spreadsheet.setActiveSheet(newSheet); // Makes the new sheet the active one

  //---------------------------------------------------------------
  // 4. Get the range of data in the new sheet (A1 to the last row/column)
  // This handles sparse data by using the sheet dimensions.
  //---------------------------------------------------------------
  const range = newSheet.getRange(1, 1, newSheet.getMaxRows(), newSheet.getMaxColumns());

  //---------------------------------------------------------------
  // 5. Convert all formulas to static values (the core requirement)
  // getValues() gets the raw data (formulas as strings), while 
  // getDisplayValues() gets the calculated results as displayed (strings).
  //---------------------------------------------------------------
  const values = range.getValues();
  
  //---------------------------------------------------------------
  // Paste the calculated values back into the range, overwriting formulas
  //---------------------------------------------------------------
  range.setValues(values);

  //---------------------------------------------------------------
  // Optional: Clean up empty rows/columns in the snapshot for a tidy look
  //---------------------------------------------------------------
  const lastRow = newSheet.getLastRow();
  const maxRows = newSheet.getMaxRows();
  if (lastRow < maxRows) {
    newSheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

  const lastColumn = newSheet.getLastColumn();
  const maxColumns = newSheet.getMaxColumns();
  if (lastColumn < maxColumns) {
    newSheet.deleteColumns(lastColumn + 1, maxColumns - lastColumn);
  }

  spreadsheet.setActiveSheet(activeSheet); // Makes the original sheet the active one
}
