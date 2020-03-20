function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💢 Paz Lab')
  .addItem('Update Tracking Categories', 'getTrackingCategory')
  .addItem('Authorization', 'showAuthUrl')
  .addSeparator()
  .addItem('Xero Invoice Reset', 'xeroInvoiceReset')  
  .addSeparator()
  .addItem('Xero Transaction Reset', 'xeroTransactionsReset')  
  .addSeparator()
  .addItem('Xero Reset', 'xeroReset')  
  .addToUi();
}

function xeroReset() {
  xeroInvoiceReset();
  xeroTransactionsReset();
}
///////////////////////////////////////////////////////////////////
function log(p){return(Logger.log(p))}