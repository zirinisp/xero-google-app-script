// https://github.com/csi-lk/google-app-script-xero-api

var API_END_POINT = 'https://api.xero.com/api.xro/2.0';
var INVOICES_END_POINT = '/Invoices';

function getAuthHeader(method, url, query) {
  var oauth_nonce = createGuid();
  var oauth_timestamp = (new Date().valueOf()/1000).toFixed(0);
  
  var signatureBase = method + "&"
      + encodeURIComponent(url);
  
  var signatureBaseMid = "";
  if (query) {
    signatureBaseMid += query+"&";
  }
  signatureBaseMid += "oauth_consumer_key=" + CONSUMER_KEY +
                      "&oauth_nonce=" + oauth_nonce + 
                      "&oauth_signature_method=RSA-SHA1&oauth_timestamp=" + oauth_timestamp + 
                      "&oauth_token=" + CONSUMER_KEY + 
                      "&oauth_version=1.0"
                           
  signatureBase += "&"+encodeURIComponent(signatureBaseMid);

  var rsa = new RSAKey();
  rsa.readPrivateKeyFromPEMString(PEM_KEY);
  var hashAlg = "sha1";
  var hSig = rsa.signString(signatureBase, hashAlg);

  var oauth_signature = encodeURIComponent(hextob64(hSig));

  var authHeader = "OAuth oauth_token=\"" + CONSUMER_KEY + "\",oauth_nonce=\"" + oauth_nonce +
    "\",oauth_consumer_key=\"" + CONSUMER_KEY + "\",oauth_signature_method=\"RSA-SHA1\",oauth_timestamp=\"" +
      oauth_timestamp + "\",oauth_version=\"1.0\",oauth_signature=\"" + oauth_signature + "\"";
  return authHeader;
}

// Working method to send a request with a payload
function sendRequest(endpoint, method, payload) {
  var url = 'https://api.xero.com/api.xro/2.0' + endpoint
  var authHeader = getAuthHeader(method, url);
  var headers = {
    "Authorization": authHeader,
    "Accept": "application/json"
  };
  var options = {
    "headers": headers,
    'method' : method,
    'payload' : payload,
    'muteHttpExceptions': true,
  };

  var response = UrlFetchApp.fetch(url, options);
  return response;
}

function createGuid() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random()*16|0, v = c == 'x' ? r : (r&0x3|0x8);
    return v.toString(16)
      });
}

function RSAKey(){this.n=null;this.e=0;this.d=null;this.p=null;this.q=null;this.dmp1=null;this.dmq1=null;this.coeff=null}