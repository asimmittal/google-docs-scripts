/*
 * create a simple container that shows the string argument
 * in a text box
*/

function dumpDataIntoUI(data){
  
  //the white container is the app
  var app = UiApp.createApplication();
  app.setHeight(480);
  app.setWidth(640);
  
  //we want a nice label with the title of the script
  var label = app.createLabel("All Rows to JSON");
  label.setStyleAttribute("font-size","24");
  label.setStyleAttribute("font-weight","bold");
  label.setStyleAttribute("padding-bottom", "25px");
  
  //next, we will add the text box and fill it with the string argument  
  var text = app.createTextArea();
  text.setHeight(400);
  text.setWidth(640);
  text.setValue(data);
  text.setSelectionRange(0, data.length);
  
  //add all the widgets into the app and show it
  app.add(label);
  app.add(text);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.show(app);
}


/*
 * function to actually grab all the rows and put them
 * into a JSON string.
*/

function rowsToJson() {
  
  //grab the rows from the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  //the first row contains the titles or json keys
  var firstRowTitles = values[0];
  var allDataJSONString = "";
  
  //this loop converts all the values starting from the second row to the last
  //into one big JSON string
  for(var i = 1; i < numRows; i++){
    var row = values[i];
    var rowData = "{\n";
    for(var j = 0; j < row.length; j++){
      var comma = (j == row.length - 1)?"":",";
      rowData += '"' + firstRowTitles[j] + '"' + ":" + '"' + row[j] + '"' + comma + "\n";
    }
    
    var comma = (i == numRows-1)?"":","
    rowData+="}" + comma + "\n";
    allDataJSONString += rowData;      
  }  
  
  //add delimiters to that string and dump into the UI 
  dumpDataIntoUI("[\n" + allDataJSONString.toLowerCase() + "\n]");  
};

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "All Rows to JSON",
      functionName : "rowsToJson"
    }    
  ];
  
  sheet.addMenu("My Scripts", entries);
};