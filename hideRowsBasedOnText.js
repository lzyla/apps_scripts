function hideRowsBasedOnText() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Get the data range of the active sheet
  var values = range.getValues(); // Get values of all cells in the data range
  var textsToHide = ["Mail Delivery Subsystem", "Mail Delivery System"]; // Add all texts you want to hide

  // Loop through all rows in the range
  for (var i = 0; i < values.length; i++) {
    // Check if the current cell's value (in the first column) is in the textsToHide array
    if (textsToHide.includes(values[i][0])) {
      sheet.hideRows(i + 1); // Hides the row if the condition is met. Rows are 1-indexed
    }
  }
}
