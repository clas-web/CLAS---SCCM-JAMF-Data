/**
* Add serial number check hyperlink to computers
* @param range The address of the cell(s) to update (optional, if not included the selected range will be the range)
*/

function hyperlink(range){
  var selected = range || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange();
  var sheetName = selected.getSheet().getName();
  var values = selected.getDisplayValues();
  var arr = [values];
  
  for (var i = 0; i < values.length; i++) {
    arr[i] = [values[i]];
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] != ""){
        if (sheetName == "SCCM (live import)") {
          arr[i][j] = '=HYPERLINK' + '("http://www.dell.com/support/home/us/en/04/product-support/servicetag/' + 
            values[i][j]+'/warranty?ref=captchaseen","'+values[i][j]+'")';
        } else if (sheetName == "Casper (live import)") {
          arr[i][j] = '=HYPERLINK' + '("https://checkcoverage.apple.com/us/en/?sn=' + 
            values[i][j]+'","'+values[i][j]+'")';
        }        
      } else {
        arr[i][j] = "";
      }      
    }
  }  
  
  //Set hyperlink   
  selected.setValues(arr);      
}

/**
* Add Casper hyperlink to Macs
* @param range1 The address of the cell(s) to update (optional, if not included the selected range will be the range)
* @param range2 The address of the computer ID(s)
*/

function hyperlinkCasper(range1, range2){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var selected = range1 || ss.getActiveSheet().getActiveRange();
  range2 = range2 || ss.getSheetByName("Casper (live import)").getRange("M2:M");
  var sheetName = selected.getSheet().getName();
  var values = selected.getDisplayValues();
  var arr = [values];
  var arr2 = range2.getValues();
  
  for (var i = 0; i < values.length; i++) {
    arr[i] = [values[i]];
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] != ""){
        if (sheetName == "Casper (live import)") {
          arr[i][j] = '=HYPERLINK' + '("https://uncc-casper.uncc.edu:8443/computers.html?id=' + 
            arr2[i][j] + '&o=r","'+values[i][j]+'")';
        }     
      } else {
        arr[i][j] = "";
      }      
    }
  }  
  
  //Set hyperlink   
  selected.setValues(arr);      
}
