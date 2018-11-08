/*

The function "update" is triggered when the form is submitted.
The argument to update (e) is the data from the form.

*/


//**************************************************************
// Global Variables
//**************************************************************
var TYPE_TABLE = {
    "General Service Project": "*G",
    "Large Group Service Project": "*G",
    "Non-APO Project": "*N",
    "Donation Hours": "*D",
    "Soup Kitchen": "*S",
    "Publicity": "*P"
  };

var TITLE;
var TYPE;   
var NAMES_B;
var NAMES_P;
var HOST;
var POINTS;
var DEBUG_COUNT = 1;

var BROTHER_SHEET_ID = "1vuZu2PUo6T9d_bpYgaGLh5z8pIT3FcrZPYwV4G56j04"; //TODO: Add brother spreadsheet id here
var PLEDGE_SHEET_ID  = "1dwLlcjGLE6lLPNzTrQOF2GD1tZYKo7YmC_2E0ALmtn8"; //TODO: Add pledge spreadsheet id here

//**************************************************************
// UPDATE STARTS HERE
//**************************************************************

// Triggered when the form is submitted. Updates the spreadsheets
function update(form) {
  
  Logger.log(form);
  
  // Read in form values
  TITLE = form.values[1];  
  POINTS = form.values[2];
  TYPE = form.values[3];
  HOST = [];
  HOST[0] = form.values[4];
  var attendeesB = form.values[5];
  var attendeesP = form.values[6];
  var attendeesBadB = form.values[7];
  var attendeesBadP = form.values[8];
  
  Logger.log(TITLE);
  Logger.log(POINTS);
  Logger.log("Type: " + TYPE);
  Logger.log("Host: " + HOST[0]);
  Logger.log("Attendees: " + attendeesB);
  Logger.log("Pledge Attendees: " + attendeesP);
  Logger.log("Bad Attendees: " + attendeesBadB);
  Logger.log("Bad Pledge Attendees: " + attendeesBadP);
  
  // Set up spreadsheet variables
  if (TYPE == "General Service Project" || TYPE == "Large Group Service Project"){
    var serviceSheetB = SpreadsheetApp.openById(BROTHER_SHEET_ID).getSheetByName("Service"); // brother service sheet
    var serviceSheetP = SpreadsheetApp.openById(PLEDGE_SHEET_ID).getSheetByName("Service Details"); // pledge service sheet
  } else if(TYPE == "Publicity"){
    var serviceSheetB = SpreadsheetApp.openById(BROTHER_SHEET_ID).getSheetByName("Publicity"); // brother service sheet
    var serviceSheetP = SpreadsheetApp.openById(PLEDGE_SHEET_ID).getSheetByName("Service Details"); // pledge service sheet
  } else{
    var serviceSheetB = SpreadsheetApp.openById(BROTHER_SHEET_ID).getSheetByName("SoupNonAPO"); // changing sheet for soup kitchens
    var serviceSheetP = SpreadsheetApp.openById(PLEDGE_SHEET_ID).getSheetByName("Service Details II"); // changing sheet for soup kitchens
  }
    
  // insert a new column for the project and get the column number
  var columnNumB = insertProjectColumn("2:2", attendeesB, serviceSheetB);
  var columnNumP = insertProjectColumn("2:2", attendeesP, serviceSheetP);
    
  // get properly formatted list of attendees, format: "last; first"
  NAMES_B = parseNames(attendeesB);
  NAMES_P = parseNames(attendeesP);
  // namesBad = parseNames(attendeesBadB);
  // namesBadP = parseNames(attendeesBadP);   

  // write points for all attendees to the spreadsheet
  writePoints(NAMES_B, serviceSheetB, columnNumB, POINTS);
  writePoints(NAMES_P, serviceSheetP, columnNumP, POINTS);

  //writePoints(namesBad, serviceSheet, columnNumB, -1);
  //writePoints(namesBadP, serviceSheetP, columnNumP, -1);
  
}

// Insert the new project as a column into the spread sheet
// Returns the number of the column that was just inserted
function insertProjectColumn(rowRangeString, attendeesList, serviceSheet) {
  // gives entire row of service project names
  var projectNamesRowRange = serviceSheet.getRange(rowRangeString); 
  
  // get the marker indicating the last column in the spreadsheet
  var marker = TYPE_TABLE[TYPE];
  var columnNumToInsert = findColumnNum(projectNamesRowRange, marker);
  
  //check to see if the column is already there, make columnNum that number if it is
  if (findColumnNum(projectNamesRowRange, TITLE) != -1) {
    columnNumToInsert = findColumnNum(projectNamesRowRange, TITLE);
  } else {
    //  only add column if there are people that attended
    if (attendeesList != "") {                                                
      serviceSheet.insertColumnBefore(columnNumToInsert);
      // isolates 1 cell
      var writeTitleRangeB = serviceSheet.getRange(2, columnNumToInsert, 1, 1); 
      writeTitleRangeB.setValue(TITLE);
    }
  }
  
  return columnNumToInsert;
}

// Takes in list of names, googlesheet, columnNumber, number of points, and whether or not they are the host.
// Writes the points onto the spreadsheet.
function writePoints(list, serviceSheet, columnNum, argPoints) {
  // if there was at least one person 
  if (list.length > 0) {                                   
    var firstNameColumnRange = serviceSheet.getRange("B:B");
    for (var i in list) {                                   
      var semiColon = list[i].indexOf(";");
      // get the first and last name
      var lastName = list[i].substring(0,semiColon).trim();               
      var firstName = list[i].substring(semiColon+2).trim();              
      
      var start = 1;
      var writePointsRange;
      while (findRowNum(firstNameColumnRange,firstName,start) != -1) {
        // find row number of the first name
        var rowNum = findRowNum(firstNameColumnRange,firstName,start);           
        // check to make sure the lastname also matches
        if (serviceSheet.getRange(rowNum, 1, 1, 1).getValue().trim() == lastName) { 
          // row = person, column = project
          writePointsRange = serviceSheet.getRange(rowNum, columnNum, 1, 1);            
          writePointsRange.setValue(argPoints);
          
          // breaks out of the loop to make sure hours are not doubled
          break;                                                                             
        } else {
          start = rowNum+1; // procedes to next person with that first name
        }
      }
    }
  }
}

// Give hosting credit, CURRENTLY NOT IN USE
function writeHost(serviceSheet, isPledgeSheet){
  for (var h in HOST) {
    if (TYPE == "General Service Project" || TYPE == "Large Group Service Project" || TYPE == "Soup Kitchen") {
      var firstNameColumnRange = serviceSheet.getRange("B:B");
      
      var semi = HOST[h].indexOf(";");
      var lastName = HOST[h].substring(0,semi).trim();
      var firstName = HOST[h].substring(semi+2).trim();
      
      var start = 1;
      var writeHostRange;

      while(findRowNum(firstNameColumnRange,firstName,start) != -1) {
        //finds row number of the first name
        var rowNum = findRowNum(firstNameColumnRange,firstName,start);      
        //checks to make sure the lastname also matches
        if (serviceSheet.getRange(rowNum, 1, 1, 1).getValue().trim() == lastName) {             
          if (isPledgeSheet) {
            // goes to hosting column
            writeHostRange = serviceSheet.getRange(rowNum, 11, 1, 1);                     
            //writeHostRange.setValue("YES");
          } else {
            // goes to hosting column
            writeHostRange = serviceSheet.getRange(rowNum, 8, 1, 1);                     
            //writeHostRange.setValue("YES");
          }
          break;                                                                           
        } else {
          start = rowNum+1; 
        }
      }
    }
  }
}

//**************************************************************
// Helpers
//**************************************************************

// returns the column number of a value (searchKey) in a single row of values (range)
function findColumnNum(range,searchKey) {
  var data = range.getValues();
  for (var j = 0; j < data[0].length; j++) {
    if (searchKey == data[0][j]) {
        return j+1;
    }
  }
  return -1;
}

// returns the row number of a value (searchKey) in a single column of values (range) starting at a specified index (start)
function findRowNum(range,searchKey,start) {
  var data = range.getValues();
  for( var i = start-1; i < data.length; i++ )
    if( searchKey == data[i][0] )
      return i+1;
  return -1;
}


// takes in a string containing a list of names, in the form: "last1, first1; last2, first2"
// returns an array of names, in the form: ["last1, first1", "last2, first2"]
function parseNames(list){
  var namesInAnArray = [];
  var count = 0;
  var comma;
                                                        // example: attendees = "Smith; John, Horton; Frank"
  while (list.indexOf(",") != -1){                      // puts all the names into an array called "names"
    comma = list.indexOf(",");                          // comma = 10
    namesInAnArray[count] = list.substring(0,comma);    // names[0] = Smith; John
    list = list.substring(comma+2);                     // attendees = Horton; Frank 
    count++;
  }
  
  if(list != ""){
    namesInAnArray[count] = list;
  }
  
  return namesInAnArray
}


// for debugging
function print(text){
  var cellNum = "A";
  var debugCell; 
  
  MembershipSheet = SpreadsheetApp.openById(BROTHER_SHEET_ID).getSheetByName("debug(DoNotDelete!)");
  debugCell = cellNum + DEBUG_COUNT;
  MembershipSheet.getRange(debugCell).setValue(text);
  
  DEBUG_COUNT = DEBUG_COUNT + 1;
}


/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function readRows() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  for (var i = 0; i <= numRows - 1; i++) {
    var row = values[i];
  }
  
  // Resubmit the form with the data from the selected row
  var currentData = sheet.getActiveRange();
  update({values: currentData.getValues()[0]});
}

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
  var entries = [{
    name : "Submit Data",
    functionName : "readRows"
  }];
  sheet.addMenu("Script Center Menu", entries);
}
