/**
* Author: Gaurav Grover (gauravgrover95@gmail.com)
* Date: 14/06/2016
* Purpose: To set basic required functions for the script
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