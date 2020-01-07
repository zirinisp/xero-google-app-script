function sandbox() {
  
  // Get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0 ; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetID) {
      var sheet = sheets[i];
      break;
    }
  }
  
  // Get query
  var queryInput = sheet.getRange(2, 3, 1, 2).getValues();
  var complexQuery = queryInput[0][0] + "=" + encodeURIComponent(queryInput[0][1]);
  var filterInput = sheet.getRange(4, 3).getValue();
  if (filterInput === "") {
    var filter = "Get all";
  } else {
    filter = filterInput.split(",");
  }
  var pageNo = sheet.getRange(2, 9).getValue();
  if (pageNo === "") {
    pageNo = 1;
  }
  
  // Get invoices
  var invoices = getInvoices_(complexQuery, pageNo);
  Logger.log(invoices);
  linebyline(invoices);
}

function linebyline(invoices) {
  // Initialise
  var invoice, key, val;
  var lineItems, lineItem, lineKey, lineVal, acCode, itemDate, lItemID, flag;
  
  var keys = ['InvoiceID', 'InvoiceNumber', 'DateString', 'Type', 'Status', 'Total'];
  var lineKeys = ['LineItemID', 'AccountCode', 'Description', 'Quantity', 'LineAmount'];
  var output = [];
  
  for (var i = 0; i < invoices.length; i++) {
    invoice = invoices[i];
    lineItems = invoice.LineItems;
    
    if (lineItems.length > 0) {
      
      var invoiceData = [];
      
      for (var k = 0; k < keys.length; k++) {
        key = keys[k];
        val = invoice[key];
        if (key == 'DateString') {
          val = val.split("T")[0];
          itemDate = val;
        }
        invoiceData.push(val);
      }
      
      for (var li = 0; li < lineItems.length; li++) {      
        lineItem = lineItems[li];
        if (lineItem.Tracking.length > 0) {
          if (lineItem.Tracking[0].Option == 'KM') {
            var lineData = [];
            
            // Check if lineItem is already on the sheet
            lItemID = lineItem.LineItemID;        
            acCode = lineItem.AccountCode;
            itemDate = new Date(itemDate);
            itemDate = itemDate.getTime();        
            for (var lik = 0; lik < lineKeys.length; lik++) {
              lineKey = lineKeys[lik];
              lineVal = lineItem[lineKey];
              lineData.push(lineVal);
            }
            lineData.push(lineItem.Tracking[0].Option)
            lineData = invoiceData.concat(lineData);
            output.push(lineData);
          }
        }
      }
    }
  }
  Logger.log(output);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sand');
  sheet.getRange(sheet.getLastRow() + 1, 1, output.length, output[0].length).setValues(output); 
}