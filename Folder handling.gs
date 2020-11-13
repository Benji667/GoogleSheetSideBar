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