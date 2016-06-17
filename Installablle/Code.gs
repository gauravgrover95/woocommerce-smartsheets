
/**
* Author: Gaurav Grover (gauravgrover95@gmail.com)
* Date: 14/06/2016
* Purpose: To set every necessary default variable and defaults for every dataType(Orders, Products, Customers etc...)
* File: Defaults.gs
*/


// fetching User, Script and Document specific properties..
var scriptProperties = PropertiesService.getScriptProperties();
var userProperties = PropertiesService.getUserProperties();
var documentProperties = PropertiesService.getDocumentProperties();
var currentSheet = SpreadsheetApp.getActiveSheet();

// fetching consumer key and consumer properties from the user variables (These needs to be set by user using "User Settings")
var consumer_key = userProperties.getProperty('CONSUMER_KEY');
var consumer_secret = userProperties.getProperty('CONSUMER_SECRET');

// fetching Site Url property from the document variables (This needs to be set by user using "User Settings")
var siteUrl = documentProperties.getProperty("SITE_URL");

var dataType = "orders";

// Declaring BaseURL containg full address of the api from where data is to be imported
var baseUrl = siteUrl + "/wc-api/v3/";
baseUrl += dataType;
baseUrl += "?consumer_key=" + consumer_key;
baseUrl += "&consumer_secret=" + consumer_secret;


// Setting Default Import Settings for Orders in Document Specific variable.
var defaultOrdersImportSettings = JSON.stringify({
  "Id": "id",
  "Order Id": "order_number",
  "Total Amount": "total",
  "Customer First Name": "customer,first_name",
  "Customer Last Name": "customer,last_name",
  "Pincode": "shipping_address,postcode",
  "Payment Mode": "payment_details,method_title"
});
documentProperties.setProperty('DEFAULT_ORDERS_IMPORT_SETTINGS', defaultOrdersImportSettings);


/*******************************************************************************************************************/


/**
* Author: Gaurav Grover (gauravgrover95@gmail.com)
* Date: 14/06/2016
* Purpose: To set basic required functions for the script
* File: Code.gs
*/


/**
* Simple Trigger Function
* Triggers let Apps Script run a function automatically when a certain event occurs.
* This function executes automatically on opening of the document
* Creates Custom Menu for the Spreadsheet 
*/
function onOpen() {  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Orders')
      .addItem('Reimport', 'reimport')
      .addItem('Clear', 'clear')
      .addItem('Import', 'import')
      .addItem('Import Settings', 'showImportSettingsDialog')
      .addItem('User Settings', 'showUserSettingsDialog')
      // .addItem('Settings', 'showSidebar')
      .addToUi();
      
}

/**
* Shows Dialog Box for User Settings option in the menu
*/
function showUserSettingsDialog() {
  
  var html = HtmlService.createHtmlOutputFromFile('User_Settings')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(300)
      .setHeight(250);      
  SpreadsheetApp.getUi() 
      .showModalDialog(html, 'User Settings');
}

/**
* Saves the user preferences and authentication variables
* @param {url}             string   url of the site from where data is to be imported
* @param {consumer_key}    string   consumer key of the wc-api provided to the user received from User_Settings.html
* @param {consumer_secret} string   consumer secret of the wc-api provided to the user received from User_Settings.html
*/
function saveUserSetttings(url, consumer_key, consumer_secret) {
  documentProperties.setProperty('SITE_URL',url);
  userProperties.setProperty('CONSUMER_KEY', consumer_key)
  userProperties.setProperty('CONSUMER_SECRET', consumer_secret);
  SpreadsheetApp.flush();
  Utilities.sleep(5000);  
}

/**
* Opens Dialog Box for Import Settings option in the menu
* Opens interface coded in Import_Settings.html
*/
function showImportSettingsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Import_Settings')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(420)
      .setHeight(320);      
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Import Settings');
}


/**
* Function used to refresh the data and sync new data from the API
*/
function reimport() {
  clear();
  import();
}

/**
* Clears the whole worksheet
*/
function clear() {
  currentSheet.clear();
}

/**
* Shows Sidebar for Detailed Session Settings.
* Opens interface coded in Settings.html
*/
//function showSidebar() {
//  var html = HtmlService.createHtmlOutputFromFile('Settings')
//      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
//      .setTitle('Orders')
//      .setWidth(100);
//  
//  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
//      .showSidebar(html);
//}


/**************************************************************************************************************/


/**
* Author: Gaurav Grover (gauravgrover95@gmail.com)
* Date: 14/06/2016
* Purpose: Collect All the code relavent to import at one place
* File: Import.gs
*/

// fetching default and user settings
var defaultImportSettings = JSON.parse(documentProperties.getProperty("DEFAULT_ORDERS_IMPORT_SETTINGS"));
var userImportSettings = JSON.parse(userProperties.getProperty("USER_ORDERS_IMPORT_SETTINGS"));


/**
* Saving user preferences when user hit on save button through GUI.
* @param  {obj}   object   import field settings received from import_settings.html send() func
*/
function setImportSettings(obj) {
  userProperties.setProperty('USER_ORDERS_IMPORT_SETTINGS', obj);
  SpreadsheetApp.flush();
  Utilities.sleep(5000);
}


/**
* delete user preference when user hits "Use Defaults" Button in import_settings.html page
*/
function deleteUserOrdersImportSettings() {
  userProperties.deleteProperty('USER_ORDERS_IMPORT_SETTINGS');
  SpreadsheetApp.flush();
  Utilities.sleep(5000);
}


/**
* getting total number of pages of data from received headers from wc-api
* per page 10 orders objects are available in the API response
* This is important to set limit to iterator in import()
*/
function getNumPages() {
  var response = UrlFetchApp.fetch(baseUrl);
  var data = response.getHeaders();
  return data["X-WC-TotalPages"];
}

/**
* fetch and parse data from wc-api
* @param  {page}   page from which data is to be fetched
*/
function fetchData(page) { 
 var url = baseUrl + "&page=" + page;
 var response = UrlFetchApp.fetch(url ,{'muteHttpExceptions': true});
 var data = response.getContentText();

 var obj = JSON.parse(data);
 var orders = obj.orders;
 return orders;
}

/**
* Creates Header of the Spreadsheet
* changing the style in sheet
* Writing Header Data
* @param  {options}   object    Header string and json path of field value as key-value pairs
*/
function createHeader(options) {
  var sheet = currentSheet;
  var range = sheet.getRange("A1:Z1");
  range.setBackground("#61B450");
  range.setFontColor("white");
  range.setFontWeight('bold');
  sheet.setFrozenRows(1);
    
  var objKeys = Object.keys(options);
  sheet.appendRow(objKeys);
}


/**
* Writing Data to the Spreadsheet Body
* options is a object (data in key-value pairs)
* Keys are the headings of the table
* Corresponding values are the JSON address of the fields according to schema in comma(,) seperated list
* @param {orders}   array of objects   orders data imported from the api
* @param {options}  object             preference data of fields to import in key-value pairs 
*/
function appendData(orders, options) {
  var sheet = currentSheet;
  /* append data with value of each field = value of orders[options["key"]] */
    orders.forEach(function(item) {
    var fields = [];
    for(var key in options) {
      // simplifying JSON address and fetching corresponding data for a particular order
      var addr = options[key].split(",");
      var field = (addr.length == 1) ? item[addr[0]] : (addr.length == 2)? item[addr[0]][addr[1]] : (addr.length == 3)? item[addr[0]][addr[1]][addr[2]] : item[addr[0]][addr[1]][addr[2]][addr[3]];
      fields.push(field);
    }
    sheet.appendRow(fields);
  });

}

/**
* Main function that uses all sub-routines to call the whole import process
* clears the sheet for fresh import
* Decides what configuration to use
* creates freezed header for data
* iterating over all pages in api to import data
*/

function import() {

  clear();
  
  var options = {};
  if(userImportSettings) {
    options = userImportSettings;
  } else {
    options = defaultImportSettings;
  }

  createHeader(options);

  var sheet = currentSheet;
  var pages = getNumPages();
  
  for(var i=1; i<=pages; ++i) {
   var orders = fetchData(i);
   appendData(orders, options);
  }
}