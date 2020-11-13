var DIALOG_TITLE = 'Dialog';
var SIDEBAR_TITLE = 'Sidebar';

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Show sidebar', 'showSidebar')
      .addItem('Convert to CSV', 'showDialog')
      .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showSidebar() {
  var ui = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle(SIDEBAR_TITLE)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showDialog() {
  printDebugMsg("Show dialog");
  var ui = HtmlService.createTemplateFromFile('Dialog')
      .evaluate()
      .setWidth(400)
      .setHeight(190)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}

/**
 * Close the existing
 */
function hideSideBar() {
  var p = PropertiesService.getScriptProperties();
  if (p.getProperty("sidebar") == "open") {
    var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>"); //Closes the current dialog or sidebar.
    SpreadsheetApp.getUi().showSidebar(html); //As 2 sidebars cannot be opened, simultaneously we overwritte the existing one by a temporal sidebar
    p.setProperty("sidebar", "close");
    }
}

function printDebugMsg(msgText) {

  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = "debug";
  if (ss.getSheetByName(mySheet) == null){
    ss.insertSheet('debug'); //Create Debug sheet if non exist
    }
  var sheet = ss.getSheetByName(mySheet);
  var cell = sheet.getRange("A1"); 
  var previousLog = cell.getValue();
  cell.setValue(previousLog + "\r" + msgText);
  var sheet = ss.getSheetByName("in");
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .getContent();
}

function debugMsg() {
  printDebugMsg("Debug message");
}

function getDebugLog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var mySheet = "debug";
  var sheet = ss.getSheetByName(mySheet);
  var cell = sheet.getRange("A1"); 
  var log = cell.getValue();
  return log;
}

function setValue(data){
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("debug");
  var cell = sheet.getRange("A1"); 
  var log = cell.setValue(data);
}

function convertToCsv(delimiter) {
  try {
  printDebugMsg("Call convert to CSV with "+delimiter);
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName('in');
  // create a folder from the name of the spreadsheet
  var folder = getParentFolder(ss); 
  var folderId = DriveApp.getFolderById(folder);
  var newFolder = folderId.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  // append ".csv" extension to the sheet name
  var fileName = ss.getName() + "_Modified.csv";
  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile(fileName, sheet, delimiter);
  // create a file in the Docs List with the given name and the csv data
  var file = newFolder.createFile(fileName, csvFile);
  //File downlaod
  var downloadURL = file.getDownloadUrl().slice(0, -8);
  printDebugMsg(downloadURL);
  showurl(downloadURL);
}
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

function convertRangeToCsvFile(csvFileName, sheet, delimiter) {
  printDebugMsg("Processing data..."); 
  // get available data range in the spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getSheetByName("in");
  var activeRange = ss.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;
    
    var column = sheet.getRange("F2:G");
    column.setNumberFormat("YYYY-MM-DD HH:MM:SS");
    
    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(delimiter) + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    printDebugMsg("Data processed using "+delimiter);
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}

function createSubFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var folder = getParentFolder(ss); 
  var folderId = DriveApp.getFolderById(folder);
  var newFolder = folderId.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  printDebugMsg("The "+newFolder.getName()+" has been created");
  newFolder = newFolder.getId();
  printDebugMsg("New folder ID: "+newFolder);
  return newFolder;
}

function getParentFolder(spreadsheet){
  var file = DriveApp.getFileById(spreadsheet.getId());
  var folders = file.getParents();
  while (folders.hasNext()){
    var parentFolder = folders.next().getId();
  };
  printDebugMsg("Current file folder: "+parentFolder);
  return parentFolder;
}
function showurl(downloadURL) {
  //Change what the download button says here
  var link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Click here to download</a>');
  SpreadsheetApp.getUi().showModalDialog(link, 'Your CSV file is ready!');
}

function convert(action){
if (action == "comma") {
    convertToCsv(",");
  } else if (action == "tab") {
    convertToCsv("\t");
  } else if (action == "semicolon") {
    convertToCsv(";");
  }
}