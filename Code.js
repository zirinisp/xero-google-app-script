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
  clearTransactionLineItems(sheetName());
  getTransactionsWithLineItems(sheetName());
}


function getSalePeriod(saleDate, periods) {
  for (period in periods) {
    
  }
}