//Create archive/backup of SCCM and Casper datafeeds in Google Sheets
//3-2-18 new archive sheet: https://docs.google.com/spreadsheets/d/1Wd-7A90uc-IE5mj_PITDA7ppOhWLrj-3T-8gYqiUe7A/edit#gid=523443885
//Changed so it wouldn't clutter up the main sheet as bad
//******************************************************************************************************************************************************
//******************************************************************************************************************************************************
//******************************************************************************************************************************************************
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Import newest SCCM Sheet', functionName: 'importSCCM'},
    {name: 'Import newest JAMF Sheet', functionName: 'importJAMF'},
    {name: 'Backup SCCM Sheet', functionName: 'clone_SG_GoogleSheet'},
    {name: 'Backup JAMF Sheet', functionName: 'clone_Casper_GoogleSheet'},     
  ];
    spreadsheet.addMenu('Generate Sheets', menuItems);
    }
    
    //******************************************************************************************************************************************************
    //******************************************************************************************************************************************************
    //******************************************************************************************************************************************************
    function importSCCM(){
    
    id = "1OzR5vCoLhnKddZoWn3cD25x98upd4QO8"; //SCCM Excel file with datafeed
    var currentDate = new Date();
    var file = DriveApp.getFileById(id);
    Logger.log(file.getMimeType());
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SCCM (live import)");
    
    // Is the attachment an Excel file?  
    if (file.getMimeType() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
    Logger.log("Excel sheet found, importing...");
    //convert to Google Sheet
    var convertedFile = {
    title: file.getName()+"_"+currentDate,
    parents: [{ id: "1QudHKu0hpIK20lenJYLCG0Ent67SB-jY" }]
  };
convertedFile = Drive.Files.insert(convertedFile,file, {
  convert:true
});

//convertedFile.parents

//import Google Sheet
var SSSheets = SpreadsheetApp.openById(convertedFile.id)
// Get full range of data
var sheetRange = SSSheets.getDataRange();
// get the data values in range
var sheetData = sheetRange.getValues();
sheet.clear();
sheet.getRange(1, 1, SSSheets.getLastRow(), SSSheets.getLastColumn()).setValues(sheetData);     

//Update A1 with note featuring the date of the import
sheet.getRange("A1").setNote("Imported " + currentDate);

} else {
  Logger.log("Not an Excel file");
}

//Add warranty check for serial numbers
hyperlink(sheet.getRange("G2:G"));

}
//******************************************************************************************************************************************************
//******************************************************************************************************************************************************
//******************************************************************************************************************************************************
function importJAMF(){
  
  //id = "17r2ViM2qQj-E800ep20Rr_xHMb2t5wjG"; //JAMF Excel xml file with datafeed, not working as of Jan 19
  id = "1EgQLzzIQNbq7RfxHuoBsatDckREeplKM"; //JAMF Excel query file with datafeed
  var currentDate = new Date();
  var file = DriveApp.getFileById(id);
  Logger.log(file.getMimeType());
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Casper (live import)");
  
  // Is the attachment an Excel file?  
  if (file.getMimeType() == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"){
    Logger.log("Excel sheet found, importing...");
    //convert to Google Sheet
    var convertedFile = {
      title: file.getName()+"_"+currentDate,
      parents: [{ id: "1QudHKu0hpIK20lenJYLCG0Ent67SB-jY" }]
  };
  convertedFile = Drive.Files.insert(convertedFile,file, {
    convert:true
  });
  
  //convertedFile.parents
  
  //import Google Sheet
  var SSSheets = SpreadsheetApp.openById(convertedFile.id)
  // Get full range of data
  var sheetRange = SSSheets.getDataRange();
  // get the data values in range
  var sheetData = sheetRange.getValues();
  sheet.clear();
  sheet.getRange(1, 1, SSSheets.getLastRow(), SSSheets.getLastColumn()).setValues(sheetData); 
  
  //Update A1 with note featuring the date of the import
  sheet.getRange("A1").setNote("Imported " + currentDate);
  
} else {
  Logger.log("Not an Excel file");
}

//Add warranty check for serial numbers
hyperlink(sheet.getRange("C2:C"));
//Add direct link to Casper
hyperlinkCasper(sheet.getRange("A2:A"), sheet.getRange("M2:M"));
}

//******************************************************************************************************************************************************
//******************************************************************************************************************************************************
//******************************************************************************************************************************************************

function clone_SG_GoogleSheet() {
  
  var date = new Date();
  var name = 'SCCM_'+ date;
  var folderId = '1QudHKu0hpIK20lenJYLCG0Ent67SB-jY';
  var resource = {
    title: name,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  }
  var fileJson = Drive.Files.insert(resource);
  var fileId = fileJson.id;
  
  var source = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = source.getSheetByName('SCCM (live import)');
  removeEmptyColumns();
  removeEmptyRows();
  
  var destination = SpreadsheetApp.openById(fileId);
  
  sheet.copyTo(destination);
  
  var sheet1 = destination.getSheetByName('Sheet1');
  destination.deleteSheet(sheet1);
  
  removeEmptyColumns();
  removeEmptyRows();
  
  //******************************************************************************************************************************************************
  //******************************************************************************************************************************************************
  //******************************************************************************************************************************************************
  
}

function clone_Casper_GoogleSheet() {
  
  var date = new Date();
  var name = 'Casper_'+ date;
  var folderId = '1QudHKu0hpIK20lenJYLCG0Ent67SB-jY';
  var resource = {
    title: name,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  }
  var fileJson = Drive.Files.insert(resource);
  var fileId = fileJson.id;
  
  var source = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = source.getSheetByName('Casper (live import)');
  
  var destination = SpreadsheetApp.openById(fileId);
  
  sheet.copyTo(destination);
  
  var sheet1 = destination.getSheetByName('Sheet1');
  destination.deleteSheet(sheet1);
  
  removeEmptyColumns();
  removeEmptyRows();
  
}

//https://productforums.google.com/forum/#!msg/docs/-2mGCzmUIkY/TwBwiX4OT4QJ
function sheetnames() { 
  var out = new Array();
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i=0 ; i<sheets.length ; i++) out.push( [ sheets[i].getName() ] );
  return out;  
  
}

//*******************************************************************************************************************
//*******************************************************************************************************************
//Delete empty columns

function removeEmptyColumns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s];
    var maxColumns = sheet.getMaxColumns(); 
    var lastColumn = sheet.getLastColumn();
    if (maxColumns-lastColumn != 0){
      sheet.deleteColumns(lastColumn+1, maxColumns-lastColumn);
    }
  }
}

//Delete empty rows
function removeEmptyRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allsheets = ss.getSheets();
  for (var s in allsheets){
    var sheet=allsheets[s];
    var maxRows = sheet.getMaxRows(); 
    var lastRow = sheet.getLastRow();
    if (maxRows-lastRow > 1){
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
    }
  }
}

//*******************************************************************************************************************
//*******************************************************************************************************************
//*******************************************************************************************************************
