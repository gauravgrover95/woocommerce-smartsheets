/**
* Author: Gaurav Grover (gauravgrover95@gmail.com)
* Date: 14/06/2016
* Purpose: Collect All the code relavent to import at one place
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