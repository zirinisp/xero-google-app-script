function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💢 Paz Lab')
  .addItem('Update Tracking Categories', 'getTrackingCategory')
  .addItem('Authorization', 'showAuthUrl')
  .addSeparator()
  .addItem('Xero Invoice Reset Current', 'updateInvoiceData')  
  .addItem('Xero Invoice Reset', 'xeroInvoiceReset')  
  .addSeparator()
  .addItem('Xero Transacrion Reset Current', 'updateTransactionData')  
  .addItem('Xero Transaction Reset', 'xeroTransactionsReset')  
  .addSeparator()
  .addItem('Xero Payments Reset', 'xeroPaymentsReset')  
  .addSeparator()
  .addItem('Update Day Totals', 'updateDayTotalsTable')  
  .addSeparator()
  .addItem('Xero Update Last Year', 'xeroUpdateLastYear')  
  .addItem('Xero Update Last 2 Years', 'xeroUpdateLast2Years')  
  .addItem('Xero Reset', 'xeroReset')  
  .addToUi();
}

function xeroReset() {
  xeroInvoiceReset();
  xeroTransactionsReset();
  xeroPaymentsReset();
  Utilities.sleep(15000);
  updateDayTotalsTable();
}

function xeroUpdateLast2Years() {
  clearInvoiceLineItems(sheetInvoices2024);
  getInvoicesWithLineItems(sheetInvoices2024);
  clearInvoiceLineItems(sheetInvoices2025);
  getInvoicesWithLineItems(sheetInvoices2025);

  clearTransactionLineItems(sheetTransactions2024);
  getTransactionsWithLineItems(sheetTransactions2024);
  clearTransactionLineItems(sheetTransactions2025);
  getTransactionsWithLineItems(sheetTransactions2025);

  xeroPaymentsReset();

  Utilities.sleep(15000);
  updateDayTotalsTable();
}

function xeroUpdateLastYear() {
  clearInvoiceLineItems(sheetInvoices2025);
  getInvoicesWithLineItems(sheetInvoices2025);

  clearTransactionLineItems(sheetTransactions2025);
  getTransactionsWithLineItems(sheetTransactions2025);

  xeroPaymentsReset();

  Utilities.sleep(15000);
  updateDayTotalsTable();
}


///////////////////////////////////////////////////////////////////
function log(p){return(Logger.log(p))}