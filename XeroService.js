// https://github.com/gsuitedevs/apps-script-oauth2
function getXeroService() {
  // Create a new service with the given name. The name will be used when
  // persisting the authorized token, so ensure it is unique within the
  // scope of the property store.
  return OAuth2.createService("xero")
  
  // Set the endpoint URLs.
  .setAuthorizationBaseUrl('https://login.xero.com/identity/connect/authorize')
  .setTokenUrl('https://identity.xero.com/connect/token')
  
  // Set the client ID and secret
  .setClientId('9175F8AFA3AD439E9E72DB9D1762D89D')
  .setClientSecret('gr1tWCDmMqMIiy7krTO-hV47D2b4TTpHtd1THr2go33uVbGi')
  
  // Set the name of the callback function in the script referenced
  // above that should be invoked to complete the OAuth flow.
  .setCallbackFunction('authCallback')
  
  // Set the property store where authorized tokens should be persisted.
  .setPropertyStore(PropertiesService.getDocumentProperties())
  
  // Set the scopes to request
  .setScope('accounting.transactions.read offline_access')
  
  // Requests offline access.
  .setParam('access_type', 'offline')
  
  // xero specific
  .setParam('redirect_uri','https://script.google.com/macros/d/1YXcJ3q9Uq9sj4jooJDsWKG4mSdO09FtbhldtPQ5_e7Wqw40VbwuNzQdE/usercallback')
  .setParam('client_id','9175F8AFA3AD439E9E72DB9D1762D89D')
  .setParam('response_type','code')  
  
}

function xeroLogout(){
  var service = getXeroService();
  service.reset();
}

function authCallback(request) {
  var xeroService = getXeroService();
  var isAuthorized = xeroService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('Denied. You can close this tab');
  }
}

function getTenantId() {
  var user_props = PropertiesService.getDocumentProperties(),
      XUPK = "oauth2.xero", // Xero User Properties Key
      tenantId;  
  var oauth2_xero = JSON.parse(user_props.getProperty(XUPK));
  
  // return if tenantId already saved for more information on tenant Ids visit https://developer.xero.com/documentation/oauth2/auth-flow
  if(oauth2_xero["tenantId"] && oauth2_xero["tenantId"] != "") return oauth2_xero["tenantId"];
    
  var xeroService = getXeroService();
  var response = UrlFetchApp.fetch('https://api.xero.com/connections', {
    headers: {
      'Authorization': 'Bearer ' + xeroService.getAccessToken(),
      'Content-Type': 'application/json'
    }
  });
  
  tenantId = JSON.parse(response.getContentText())[0]["tenantId"];
  oauth2_xero["tenantId"] = tenantId;
  user_props.setProperty(XUPK, JSON.stringify(oauth2_xero));
  
  return tenantId;  
}

/**
 * Converts key value pair to percent encoded param string
 * @param {object} obj String object
 */
String.prototype.addQuery = function(obj,ignore_keys) {
  return this + Object.keys(obj).reduce(function(p, e, i) {
    return p + (i == 0 ? "?" : "&") +
      (Array.isArray(obj[e]) ? obj[e].reduce(function(str, f, j) {
        return str + e + "=" + encodeURIComponent(f) + (j != obj[e].length - 1 ? "&" : "")
      },"") : e + "=" + (Array.isArray(ignore_keys) ? (ignore_keys.indexOf(e) === -1 ? encodeURIComponent(obj[e]) : obj[e]) : encodeURIComponent(obj[e])));
  },"");
}

function getXeroHeaders(){
    return {
        'Authorization': 'Bearer ' + getXeroService().getAccessToken(),
        'Content-Type': 'application/json',
        'Accept' : 'application/json',
        'Xero-tenant-id' : getTenantId()
    } 
}

/**
 * Gets the invoices according to a where clause
 * @param {number} page An integer
 * @param {string} filter A valid where clause for more information visit https://developer.xero.com/documentation/api/requests-and-responses#get-modified
 */
function getXeroInvoices(complexQuery){   
  var url = "https://api.xero.com/api.xro/2.0/Invoices";
  
  // see addQuery prototype for details
  var end_point = url.addQuery(complexQuery,["Statuses"]);
  
  var response = UrlFetchApp.fetch(end_point, {
    headers : getXeroHeaders()
  })
  var responseText = response.getContentText();
  var json = JSON.parse(responseText);

  return json.Invoices;
}

/**
 * Gets the invoices according to a where clause
 * @param {number} page An integer
 * @param {string} filter A valid where clause for more information visit https://developer.xero.com/documentation/api/requests-and-responses#get-modified
 */
function getXeroTransactions(complexQuery){   
  var url = "https://api.xero.com/api.xro/2.0/banktransactions";
  
  // see addQuery prototype for details
  var end_point = url.addQuery(complexQuery,["Statuses"]);
  
  var response = UrlFetchApp.fetch(end_point, {
    headers : getXeroHeaders()
  })
  var responseText = response.getContentText();
  var json = JSON.parse(responseText);
  return json.BankTransactions;
}

/**
 * Gets the invoices according to a where clause
 * @param {number} page An integer
 * @param {string} filter A valid where clause for more information visit https://developer.xero.com/documentation/api/requests-and-responses#get-modified
 */
 function getXeroPayments(complexQuery){   
  var url = "https://api.xero.com/api.xro/2.0/Payments";
  
  // see addQuery prototype for details
  var end_point = url.addQuery(complexQuery,["Statuses"]);
  
  var response = UrlFetchApp.fetch(end_point, {
    headers : getXeroHeaders()
  })
  var responseText = response.getContentText();
  var json = JSON.parse(responseText);

  return json.Payments;
}


function showAuthUrl() {
  var xeroService = getXeroService();
  if (!xeroService.hasAccess()) {
    var authorizationUrl = xeroService.getAuthorizationUrl();
    var template = HtmlService.createTemplate(
        '<a href="<?= authorizationUrl ?>" target="_blank">Authorize</a>.');
    template.authorizationUrl = authorizationUrl;
    var page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    var html = '<button onclick="google.script.run.xeroLogout(); return false;">Log Out</button>';
    
    SpreadsheetApp.getUi()
    .showSidebar(HtmlService.createHtmlOutput(html));    
  }
}

function logauthURL(){
  var xeroService = getXeroService();
  if (!xeroService.hasAccess())
    Logger.log(xeroService.getAuthorizationUrl());
}

function logUserProperties(){
  Logger.log(JSON.stringify(PropertiesService.getDocumentProperties().getProperties()));
}

function getTenantIdTest(){
  Logger.log(getTenantId());
}

function getXeroInvoiceTest(){
  var complexQuery = {};
  complexQuery["Statuses"] = "PAID";
  complexQuery["where"] = "Date>=DateTime(2019,3,1)&&Date<=DateTime(2019,3,10)";
  complexQuery["page"] = 1;
  Logger.log(getXeroInvoices(complexQuery)); // year month date
  // Statuses=PAID
}

function getXeroTransactionsTest(){
  var complexQuery = {};
  complexQuery["Statuses"] = "AUTHORIZED";
  complexQuery["where"] = "Date>=DateTime(2019,3,1)&&Date<=DateTime(2019,3,10)";
  complexQuery["page"] = 1;

  Logger.log(getXeroTransactions(complexQuery)); // year month date
  // Statuses=PAID
}

// Convert Date to Xero-DateTime
function convertDateToXero(date) {
  var xeroDate = "DateTime("+date.getFullYear()+","+(date.getMonth()+1)+","+date.getDate()+")";
  return xeroDate;
}