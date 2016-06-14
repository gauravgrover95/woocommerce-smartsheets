/**
* Author: Gaurav Grover (gauravgrover95@gmail.com)
* Date: 14/06/2016
* Purpose: To set every necessary default variable and defaults for every dataType(Orders, Products, Customers etc...)
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
