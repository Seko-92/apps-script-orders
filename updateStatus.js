function updateStatus(rowNumber, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('All orders');
  
  // STATUS is column F (6th column) - adjust if different
  const statusCol = 6;
  
  sheet.getRange(rowNumber, statusCol).setValue(status);
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    rowNumber: rowNumber,
    status: status
  })).setMimeType(ContentService.MimeType.JSON);
}