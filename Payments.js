// Global variable

var batchSize = 200; // Get these many pages and then stop

var sheetUberPayments = 'Uber-Payments';
var sheetDeliverooPayments = 'Deliveroo-Payments';
var sheetZettlePayments = 'Zettle-Payments';
var sheetApiPayments = 'Api-Payments';
var sheetFeedrPayments = 'Feedr-Payments';
var sheetJEBPayments = 'JustEat Business-Payments';
var sheetJEPayments = 'JustEat-Payments';

//-------------------------------------------------------

function testDate() {
  var dateString1 = 1631059200000;
  var dateString2 = "/Date(1631059200000+0000)/";
  var dateString2 = dateString2.slice(1,-1);

  var date1 = Date.parse(dateString1);
  var date1n = new Date(dateString1);
  var date4 = Date.parse(dateString2);

  var date3 = Date.now();

}



function xeroPaymentsReset() {
  var payments = [
    sheetUberPayments,
    sheetDeliverooPayments,
    sheetZettlePayments,
    sheetApiPayments,
    sheetFeedrPayments,
    sheetJEBPayments,
    sheetJEPayments
  ]
  payments.forEach(element => {
    clearPaymentsLineItems(element);
    getPaymentsWithLineItems(element);      
  });
}

function testPayments() {
  clearPaymentsLineItems(sheetUberPayments);
  getPaymentsWithLineItems(sheetUberPayments);
}

function clearPaymentsLineItems(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Clearing "+sheetName);

  // Get sheet
  var sheet = ss.getSheetByName(sheetName);

  // Clear sheet
  if (sheet.getLastRow() > 9) {
    sheet.getRange(10, 1, sheet.getLastRow() - 9, sheet.getLastColumn()).clearContent();
  }
  
  // Clear page number
  sheet.getRange(2, 9, 2, 1).clearContent();

  // Clear date
  sheet.getRange(6, 9).clearContent();
}


function getPaymentsWithLineItems(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("Getting "+sheetName);

  // Get sheet
  var sheet = ss.getSheetByName(sheetName);
  
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
  if (filterInput !== "") {
    complexQuery["where"] += "&&Reference=\""+filterInput+"\"";
  }
  
  // Get page number
  var pageNo = sheet.getRange(2, 9).getValue();
  if (pageNo === "") {
    pageNo = 1;
  }
  
  complexQuery["page"] = pageNo;
  // Get invoices
  var payments = getPayments_(sheetName, complexQuery);
  
  if (payments.length > 0) {
    // Get line items
    var lineItems = getPaymentsLineItems_(sheetName, payments);
    
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
function getPayments_(sheetName, complexQuery){
  var pageNo = complexQuery["page"],
      payments = [];
  // Get sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  var currentPageCell = sheet.getRange(3, 9);
  
  var moreFlag = true;
  var uptoPage = parseInt(pageNo) + batchSize; // But not including
  
  while (moreFlag && pageNo < uptoPage ) {
    var invoicesPage = getXeroPayments(complexQuery)
    if (invoicesPage.length > 0 ) {
      payments = payments.concat(invoicesPage);
      currentPageCell.setValue(pageNo);
      SpreadsheetApp.flush();
    } else { // No more data
      moreFlag = false;
    }
    complexQuery["page"] = ++pageNo; // next page  
  }
  return payments;
}



////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function getPaymentsLineItems_(sheetName, payments) {
  
  // Initialise
  var output = [];
  
  for (var i = 0; i < payments.length; i++) {

    var payment = payments[i];
    var lineData = [];

    var currencyRate;
    if (payment.CurrencyCode == "GBP") {
      currencyRate = 1.0;
    } else {
      currencyRate = payment.CurrencyRate;
    }
    if ((currencyRate === undefined) || (currencyRate == 0.0)) {
      currencyRate = 1.0;
    }

    var paymentDateString = payment.Date;
    paymentDateString = parseInt(paymentDateString.slice(6,-7), 10);
    var paymentDate = new Date(paymentDateString);
    var sheetDate = Utilities.formatDate(paymentDate, "GMT+0", "dd/MM/yyyy")

    var paymentAmount = -payment.Amount;

    lineData.push(payment.PaymentID);
    lineData.push(payment.Reference);
    lineData.push(payment.Account.Code);
    lineData.push(sheetDate);
    lineData.push(payment.PaymentType); //TYPE
    lineData.push(""); //Contact Name
    lineData.push(payment.Status);
    lineData.push(paymentAmount); // Total
    lineData.push(""); // Line Amount Types
    lineData.push(payment.CurrencyCode);// Currency code
    lineData.push(currencyRate);// Currency rate
    
    lineData.push(""); // Line Item Id
    lineData.push(payment.Account.AccountID);
    lineData.push(""); // Description
    lineData.push(1); // Quantity
    lineData.push(paymentAmount); // Line Amounbt
    lineData.push(0); // Tax Amounbt
    lineData.push(""); // Tracking
    
          

    var gbpAmountNoTax = paymentAmount;
    var gbpTax = 0;
    var gbpAmountWithTax = gbpAmountNoTax;
    if (currencyRate != 1.0) {
      gbpAmountNoTax = gbpAmountNoTax / currencyRate;
      gbpAmountWithTax = gbpAmountWithTax / currencyRate;
    }
    lineData.push(gbpAmountNoTax);
    lineData.push(gbpTax);
    lineData.push(gbpAmountWithTax);

    lineData.push(payment.PaymentID);
    output.push(lineData);
  }
  return output;
}
