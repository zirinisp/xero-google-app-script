function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('💢 Paz Lab')
  .addItem('Update Tracking Categories', 'getTrackingCategory')
  .addSeparator()
  .addItem('Xero Update', 'getInvoicesWithLineItems')
  .addItem('Xero Reset', 'xeroReset')  
  .addToUi();
}
///////////////////////////////////////////////////////////////////
function log(p){return(Logger.log(p))}