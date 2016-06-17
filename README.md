WooCommerce SmartSheets  
=======================
![Google Sheets Logo](https://www.gstatic.com/images/icons/material/product/1x/sheets_64dp.png) *Google Sheets add-on to import your WooCommerce supported E-Commerce website data to Google Sheets.* 

No more writing data manually to Spreadsheets for your E-Commerce Website. This Google Sheets add-on will make your worksheets smart enough to import data from your E-Commerce website for you in just few clicks.

** Currently just supported for Orders. More Coming Soon...


## Why one would do that? ##
Spreadsheets does a great job at handling the data, making it well structured and organized. Not only that, it also supports various very powerful tools (like charts, graphs, functions) which make data visualization much more fun. "Wait! doesn't E-Commerce Solutions like Woo-Commerce itself already have these things?" Well, Sure they do, but they don't give you the full liberty to view, modify, ammend and analyse the data whatever you want and  however you want. 

Also Google Sheets' support over devices, of all screen sizes, is phenomenal. Why wouldn't you want to keep all your data in your pocket to stay up to date with it.


## Installation ##

Just a few clicks away...

## Method 1 (Direct Install) (Coming Soon) ##

 1. Open a new SpreadSheet
 2. Click on "Add-Ons" from Menu Bar
 3. Click on "Get add-ons"
 4. Search for WooCommerce Smartsheets
 5. Click install


************************************ OR **************************************

## Method 2 (Code Duplication) ##

 1. Create/Open a Spreadsheet where you want to install the script.
 2. Click on *Tools* > *Script Editor* from Menu Bar.
 3. Make three files *Code.gs*, *Import_Settings.html* & *User_Settings.html* from *File > New* option in *Script Editor*.
 4. Open the folder *Installable* from Github repositiory. 
 5. Duplicate the code from Github repository to corresponding files you made.
 6. Save all the files and refresh the spreadsheet.
    
\* (make sure file names are same as given).

and its done...  :)


## Lets take a Demo ##
Here i am leaving some screenshots to guide you through the demo for the usage of the add-n. 

For Screenshots click [here](http://imgur.com/a/mWfSm)

*"A picture is worth the million words"*


## File Structure ##

 - **Code.gs** : To set basic code required functions for the script 
 - **Defaults.gs** : To set every necessary default variable and defaults for every dataType(Orders, Products, Customers etc...)
 - **Import.gs** :  All the code relavent to import mechanism at one place
 - **Documentation.gs** : Basic Documentation for the code
 - **Import_Settings.html** : GUI for the "Import Setting" option in the custom menu
 - **User_Settings.html** : GUI for the "User Setting" option in the custom menu
 - **Settings.html** : GUI for the "Setting" option in the custom menu, opens sidebar of settings
