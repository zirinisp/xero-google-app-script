// Global variable

var batchSize = 100; // Get these many pages and then stop
var sheetID = '1269043913';

//-------------------------------------------------------

function xeroReset() {
  clearLineItems();
  getInvoicesWithLineItems();
}

function clearLineItems() {
  
  // Get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0 ; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetID) {
      var sheet = sheets[i];
      break;
    }
  }

  // Clear sheet
  if (sheet.getLastRow() > 9) {
    sheet.getRange(10, 1, sheet.getLastRow() - 9, sheet.getLastColumn()).clearContent();
  }
  
  // Clear page number
  sheet.getRange(2, 9, 2, 1).clearContent();

  // Clear date
  sheet.getRange(6, 9).clearContent();
}

function getInvoicesWithLineItems() {
  
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
  var complexQuery = {};
  
  // Get Status
  var statuses = sheet.getRange("D2").getValue();  
  if(statuses != "")
    complexQuery["Statuses"] = statuses;

  // Get Start and End Dates
  var startDate = sheet.getRange(6, 3).getValue();
  var endDate = sheet.getRange(6,6).getValue();
  var dateQuery = "";
  if (startDate) {
    
    var startDateInput = ">="+convertDateToXero(startDate);
    var startDateQuery = "Date" + startDateInput;
    dateQuery = startDateQuery;
  }

  if (endDate) {
    var endDateInput = "<="+convertDateToXero(endDate);
    endDateQuery = "Date" + endDateInput;
    if (dateQuery.length > 0) {
      dateQuery += "&&";
    }
    dateQuery += endDateQuery;
  }
  if (dateQuery.length > 0) {
    complexQuery["where"] = dateQuery;
  }
  
  // Get account code filter
  var filterInput = sheet.getRange(4, 3).getValue();
  if (filterInput === "") {
    var filter = "Get all";
  } else {
    filter = filterInput.split(",");
  }
  
  // Get page number
  var pageNo = sheet.getRange(2, 9).getValue();
  if (pageNo === "") {
    pageNo = 1;
  }
  
  complexQuery["page"] = pageNo;
  // Get invoices
  var invoices = getInvoices_(complexQuery);
  
  if (invoices.length > 0) {
    // Get line items
    var lineItems = getLineItems_(invoices, filter);
    
    // Paste values
    if(lineItems.length > 0)
      sheet.getRange(sheet.getLastRow() + 1, 1, lineItems.length, lineItems[0].length).setValues(lineItems);
    
    // Paste date of latest query
    sheet.getRange(6, 9).setValue(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
    
  } else {
    SpreadsheetApp.getActive().toast("No new invoices");
  }
  
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function getInvoices_(complexQuery){
  var pageNo = complexQuery["page"],
      invoices = [];
  // Get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0 ; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetID) {
      var sheet = sheets[i];
      break;
    }
  }
  
  var currentPageCell = sheet.getRange(3, 9);
  
  var moreFlag = true;
  var uptoPage = parseInt(pageNo) + batchSize; // But not including
  
  while (moreFlag && pageNo < uptoPage ) {
    var invoicesPage = getXeroInvoices(complexQuery)
    if (invoicesPage.length > 0 ) {
      invoices = invoices.concat(invoicesPage);
      currentPageCell.setValue(pageNo);
      SpreadsheetApp.flush();
    } else { // No more data
      moreFlag = false;
    }
    complexQuery["page"] = ++pageNo; // next page  
  }
  return invoices;
}

function _getInvoices_(complexQuery, pageNo) {
  
  // Get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0 ; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetID) {
      var sheet = sheets[i];
      break;
    }
  }
  
  var currentPageCell = sheet.getRange(3, 9);
  
  // Initialise
  var moreFlag = true;
  var uptoPage = parseInt(pageNo) + batchSize; // But not including
  var output = [];
  var method = 'GET';
  var requestURL = API_END_POINT + INVOICES_END_POINT ;
  var oauth_signature_method = 'RSA-SHA1';
  var oauth_version = '1.0';    
  var rsa = new RSAKey();
  rsa.readPrivateKeyFromPEMString(PEM_KEY);
  var hashAlg = "sha1";
  
  var invoices = [];
  
  while (moreFlag && pageNo < uptoPage ) {
    var oauth_nonce = createGuid();
    var oauth_timestamp = (new Date().getTime()/1000).toFixed();
    var signBase = 'GET' + '&' + encodeURIComponent(requestURL) + '&' +
      encodeURIComponent( complexQuery + '&oauth_consumer_key=' + CONSUMER_KEY + '&oauth_nonce=' + 
                         oauth_nonce + '&oauth_signature_method=' + oauth_signature_method + '&oauth_timestamp=' +
                         oauth_timestamp + '&oauth_token=' + CONSUMER_KEY + '&oauth_version=' + oauth_version + '&page=' + pageNo);
    
    var hSig = rsa.signString(signBase, hashAlg);
    
    var oauth_signature = encodeURIComponent(hextob64(hSig));  
    var authHeader = "OAuth oauth_token=\"" + CONSUMER_KEY + "\",oauth_nonce=\"" + oauth_nonce + 
      "\",oauth_consumer_key=\"" + CONSUMER_KEY + "\",oauth_signature_method=\"RSA-SHA1\",oauth_timestamp=\"" +
        oauth_timestamp + "\",oauth_version=\"1.0\",oauth_signature=\"" + oauth_signature + "\"";
    
    var dateDisplayValue = sheet.getRange(6, 9).getDisplayValue();
    if (dateDisplayValue === "") {
      dateDisplayValue = "1990-01-01";
    }
    var utcDate = new Date(dateDisplayValue);
    utcDate = utcDate.toUTCString();
    
    var headers = { "Authorization": authHeader, "Accept": "application/json", "If-Modified-Since": utcDate};
    var options = { 'muteHttpExceptions': true, "headers": headers}; 
    var url = requestURL + '?' + complexQuery + '&page=' + pageNo;
    var response = UrlFetchApp.fetch(url, options);
    var json = response.getContentText();
    var data2 = JSON.parse(json);
    var invoicesPage = data2.Invoices;
    if (invoicesPage.length > 0 ) {
      invoices = invoices.concat(invoicesPage);
      currentPageCell.setValue(pageNo);
      SpreadsheetApp.flush();
    } else { // No more data
      moreFlag = false;
    }
    pageNo++; // next page  
  }
  
  return invoices;
}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getLineItems_(invoices, filter) {
  
  // Get sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0 ; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetID) {
      var sheet = sheets[i];
      break;
    }
  }
  // Get lineItems already fetched
  var oldData = sheet.getDataRange().getValues();
  var oldLineItemsArray = [];
  for (var x = 9; x < oldData.length; x++) {
    oldLineItemsArray.push(oldData[x][6]);
  }
  
  
  // Get date filters
  var startDate = sheet.getRange(6, 3).getValue();
  if (startDate == "") {
    startDate = 1;
  }
  startDate = new Date(startDate);
  startDate = startDate.getTime();
  
  var endDate = sheet.getRange(6, 6).getValue();
  if (endDate == "") {
    endDate = new Date();
  } else {
    endDate = new Date(endDate);
  }
  endDate = endDate.getTime();
  
  // Get Tracking Category Filter
  // var tCat = sheet.getRange(4, 9).getValue();
  
  // Initialise
  var invoice, key, val;
  var lineItems, lineItem, lineKey, lineVal, acCode, itemDate, lItemID, flag;
  
  var keys = ['InvoiceID', 'InvoiceNumber', 'DateString', 'Type', 'Status', 'Total'];
  var lineKeys = ['LineItemID', 'AccountCode', 'Description', 'Quantity', 'LineAmount', 'TaxAmount'];
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
        var lineData = [];
        lineItem = lineItems[li];
        //if (lineItem.Tracking[0].Option == tCat || tCat === '') {
          // Check if lineItem is already on the sheet
          lItemID = lineItem.LineItemID;
          if ( oldLineItemsArray.indexOf(lItemID) == -1) {
            flag = true; // new Line Item
          } else {
            flag = false; // old Line Item
          }
          
          acCode = lineItem.AccountCode;
          itemDate = new Date(itemDate);
          itemDate = itemDate.getTime();
          
          // Filter ==> match index code      OR  no filter applied  AND   date after start date AND date before end date
          if ( (filter.indexOf(acCode) !== -1 || filter == "Get all" ) && ( itemDate >=  startDate && itemDate <= endDate) ) {
            for (var lik = 0; lik < lineKeys.length; lik++) {
              lineKey = lineKeys[lik];
              lineVal = lineItem[lineKey];
              lineData.push(lineVal);
            }
            if ( lineItem.Tracking.length > 0 ) {
              lineData.push(lineItem.Tracking[0].Option);
            } else {
              lineData.push('');            
            }
            lineData = invoiceData.concat(lineData);
            
            if (flag) {
              
              // New Line Item. Add to array to be pasted later.
              output.push(lineData);
            } else {
              
              // Old Line Item. Replace row
              sheet.getRange(oldLineItemsArray.indexOf(lItemID) + 10, 1, 1, keys.length + lineKeys.length + 1).setValues([lineData]);    
            }
          }
        //}
      }
    }
  }
  return output;
}

// Convert Date to Xero-DateTime
function convertDateToXero(date) {
  var xeroDate = "DateTime("+date.getYear()+","+(date.getMonth()+1)+","+date.getDate()+")";
  return xeroDate;
}