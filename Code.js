/*
*/
function sheetName() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
}

function updateInvoiceData() {
  clearInvoiceLineItems(sheetName());
  getInvoicesWithLineItems(sheetName());
}

function updateTransactionData() {
  clearTransctionLineItems(sheetName());
  getTransactionsWithLineItems(sheetName());
}