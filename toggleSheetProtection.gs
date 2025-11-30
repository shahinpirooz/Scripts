/******************************************************************
 * Toggles the protections on the active sheet on or off
 ******************************************************************/
function toggleSheetProtection() {

  var sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getProtections().length > 0) {

    sheet.getProtections()[0].remove();

  } else {

    sheet.protect();

  }

}
