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
  .addItem('Update Day Totals', 'updateDayTotals')  
  .addSeparator()
  .addItem('Xero Reset', 'xeroReset')  
  .addToUi();
}

function xeroReset() {
  xeroInvoiceReset();
  xeroTransactionsReset();
  sleep(15000);
  updateSales();
}
///////////////////////////////////////////////////////////////////
function log(p){return(Logger.log(p))}