/*====================================================================================================================================*
  WooCommerce Smartsheets by Gaurav Grover (@Grozip)
  ====================================================================================================================================
  Version:      1.0.0
  Project Page: 
  Copyright:    (c) 2016 by Gaurav Grover
  License:      The MIT License (MIT) 
                https://opensource.org/licenses/MIT
  ------------------------------------------------------------------------------------------------------------------------------------
  Add-On for importing Orders data from WC-API into Google spreadsheets. Functions include:
     import                Import data from a URL by stating preferences 
     fetchData             Import a JSON feed from a URL using GET parameters
     getNumPages           Get total number of pages of data available in API from headers
     createHeader          Create Header containg headings of data fields and style it
     appendData            Write data to the worksheet body
     Refresh               Sync latest data and do fresh import
     clear                 Clears the Active Worksheet

  Future enhancements may include:
   - Refresh Algorithm needs improvements.
   - Make it scalable to import any data. Example Products, Customers, Orders etc.
   - Detection of empty field and write N/A instead of showing that field empty
   - Making it Real Time if possible
   - Add User friendly controls of settings of the script with various customization options like currency option, refresh rate, etc
   - Decide the default fields(name, id, address) of every data option(Products, Customers etc.)
   - Create import wizard with import constraint according to user like vendors should be given option only to import their own products.
   - Improvement for automatic detection of scalability in function send() in client side js code in import_settings.html file.
   - Algorithm to improve in import(), Currently it is goin O(n-cube).
   - three ternary operators in funtion appendDummyData is against scalability. better way-out of that
   - Detecting Asynchronous calls and make sure that it completes before making a call to another function which depends on output of it.
     For Now I just made the script to sleep for 5 sec so that it get enough time to set its settings
  ------------------------------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.1 - Changed Refresh menu and function name to Reimport
  1.0.0 - Initial release
 *====================================================================================================================================*/

