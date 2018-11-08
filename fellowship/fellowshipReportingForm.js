/*

the funciton "update" is triggered when the form is submitted.
the argument to update (e) is the data from the form.

*/


var type;
var survey;
var names;
var namesP;
var host;
var numPeople;
var points;

brotherSheetID = ""; //TODO: use sheet id for brother spreadsheet here (the random string between /d and /edit in the UR:)
pledgeSheetID  = ""; //TODO: use sheet id for pledge spreadsheet here


// returns the column number of a value (searchKey) in a single row of values (range)
function findColumnNum(range,searchKey) {
  var data = range.getValues();
  for( var j = 0; j < data[0].length; j++ )
    if( searchKey == data[0][j] )
      return j+1;
  return -1;
}

// returns the row number of a value (searchKey) in a single column of values (range) starting at a specified index (start)
function findRowNum(range,searchKey,start) {
  var data = range.getValues();
  for( var i = start-1; i < data.length; i++ )
    if( searchKey.trim() == data[i][0].trim() )
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


// takes in an array of names, removes the hosts from the array if present
// returns the array without hosts.
function removeHost(array){
  var h = array.indexOf(host[0]);
  if(h != -1) {
    array.splice(h, 1);
  }
  var h2 = array.indexOf(host[1]);
  if(h2 != -1) {
    array.splice(h2, 1);
  }
  return array;
}

// determines the amount of points that attendees and hosts should get from the fellowship
// modifies points[]
function calculatePointsAndDetermineHosting(){
  points = [];
  
  switch (type)
  {
    case "Regular":
      points[0] = 1;
      points[1] = 1.5;
      break;
    case "Recurring":
      points[0] = 1;
      points[1] = 1;
      break;
    case "Adventure":
      points[0] = 2;
      points[1] = 2.5;
      break;
    case "Random":
      points[0] = 0.5;
      points[1] = 0.5;
      // add hosts to regular array, remove both names from host array
      names = names.concat(host);
      namesP = namesP.concat(host);
      host.splice(0,2);
      break;
    case "Service Planned":
      if(survey=="Yes"){
        points[0] = 1;
        points[1] = 1;
      }
      else{
        points[0] = 1;
        points[1] = 1;
      }
      break;
    case "Service Random":
    case "Leadership Random":
      if(survey=="Yes"){
        points[0] = 0.5;
        points[1] = 0.5;
      }
      else{
        points[0] = 0.5;
        points[1] = 0.5;
      }
      // add hosts to regular array, remove both names from host array
      names = names.concat(host);
      namesP = namesP.concat(host);
      host.splice(0,2);
      break;
    default:
      points[0] = 0;
      points[1] = 0;
  }
  
  // Super Fellowship!
  if( ( numPeople >= 15) && (type.indexOf("Random") == -1)){
    if(type == "Recurring"){
      points[0] = points[0]+1;
      points[1] = points[1]+1;
    }
    else{
      points[0] = points[0]+1;
      points[1] = points[1]+1;
    }
  }
  // Half Fellowship
  else if( ((numPeople == 3) || (numPeople == 4)) &&  type != "Service Planned" && type != "Service Random" && type != "Leadership Random"){
    if(type.indexOf("Random") == -1){
      points[0] = 0.5;
      points[1] = 0.5;
    }
    else{
      points[0] = 0;
      points[1] = 0;
    }
    // add hosts to regular array, remove both names from host array
    names = names.concat(host);
    namesP = namesP.concat(host);
    host.splice(0, 2);
  }
  // No Fellowship
  else if( numPeople < 3){
    points[0] = 0;
    points[1] = 0;
    names = names.concat(host);
    namesP = namesP.concat(host);
    host.splice(0, 2);
  }
  
}


// takes in list of names, googlesheet, columnNumber, number of points, and whether or not they are the host.
// writes the points onto the spreadsheet. 

function writePoints(list, fellowshipSheet, columnNum, argPoints, isHost){
  if (list.length > 0) {                                   // if there was at least one person 
    var firstColumnRange = fellowshipSheet.getRange("A:A");  // last name column
    var secondColumnRange = fellowshipSheet.getRange("B:B"); // first name column
    var numPoints;
    for (var i in list) {                                     // cool notation, bro
      if(list[i] == "None"){
        continue;
      }
      numPoints = argPoints;
      var semi = list[i].indexOf(";");
      var lastName = list[i].substring(0,semi).trim();               // isolates last name
      var firstName = list[i].substring(semi+2).trim();              // isolates first name
      
      var start = 1;
      var rowNum;
      var writePointsRange;
      while(findRowNum(secondColumnRange,firstName,start) != -1) {
        rowNum = findRowNum(secondColumnRange,firstName,start);                              //finds row number of the first name
        if(fellowshipSheet.getRange(rowNum, 1, 1, 1).getValue().trim() == lastName) {             //checks to make sure the lastname also matches
          writePointsRange = fellowshipSheet.getRange(rowNum, columnNum, 1, 1); // row = person, column = project
          // this is to give recurring fellowship hosts 1.5 points for their first recurring fellowship
          if(type == "Recurring" && isHost == 1 && fellowshipSheet.getRange(rowNum, columnNum+2, 1, 1).getValue() == ""){
            if(numPeople < 15){
              numPoints = numPoints + 0.5;
            }
            fellowshipSheet.getRange(rowNum, columnNum+2, 1, 1).setValue("hosted");
          }
          writePointsRange.setValue(numPoints);
          break;                                                                             //breaks out of the loop to make sure hours are not doubled
        }
        start = rowNum+1; // procedes to next person with that first name
      }
    }
  }
}

function writeHost(fellowshipSheet, isPledgeSheet){
  var secondRowRange = fellowshipSheet.getRange("2:2"); // gives entire row of service project names  
  for (var h in host){
    
    if (type=="Regular" || type=="Recurring" || type=="Adventure" || type=="Service Planned"){
      var firstColumnRange = fellowshipSheet.getRange("A:A");
      var secondColumnRange = fellowshipSheet.getRange("B:B");
      
      var semi = host[h].indexOf(";");
      var lastName = host[h].substring(0,semi).trim();
      var firstName = host[h].substring(semi+2).trim();
      
      var start = 1;
      var rowNum;
      var writeHostRange;
      while(findRowNum(secondColumnRange,firstName,start) != -1) {
        rowNum = findRowNum(secondColumnRange,firstName,start);                              //finds row number of the first name
        if(fellowshipSheet.getRange(rowNum, 1, 1, 1).getValue().trim() == lastName) {             //checks to make sure the lastname also matches
          writeHostRange = fellowshipSheet.getRange(rowNum, 7, 1, 1);                     // goes to hosting column
          if(isPledgeSheet == 0){
            writeHostRange.setValue("YES");
          }
          else{
            writeHostRange.setValue("=TRUE()");
          }
          break;                                                                             //breaks out of the loop to make sure hours are not doubled
        }
        start = rowNum+1; // continues to next name if last name doesn't match
      }
    }
    
  }
}

function print(text){
  fellowshipSheet = SpreadsheetApp.openById(brotherSheetID).getSheetByName("Fellowship");
  fellowshipSheet.getRange('C1').setValue(text);
}

///////////////////////////////////////////////////////////////////////////////////////////
// UPDATE STARTS HERE
///////////////////////////////////////////////////////////////////////////////////////////



// updates the spreadsheet given a form submission
function update(e) {
  
  console.log(e.values);
  
  var pointsHost;
  var typeTable = {
    "Regular": "*R",
    "Recurring": "*RE",
    "Adventure": "*AD",
    "Random": "*RA",
    "Service Planned": "*SE",
    "Service Random": "*SE",
    "Leadership Random": "*LE"
  };
  
  
  //---------------------------//
  // read in form values       //
  //---------------------------//
  
  var title = e.values[2];
  var attendees = e.values[5];
  var attendeesP = e.values[6];
  
  type = e.values[1];
  survey = "No";
  
  host = [];
  host[0] = e.values[3];
  if(e.values[4] != ""){
    host[1] = e.values[4];
  }
  
  
  //------------------------------------//
  // Set up the spreadsheet variables   //
  //------------------------------------//
  
  if(type == "Regular"){
    var fellowshipSheet = SpreadsheetApp.openById(brotherSheetID).getSheetByName("Fellowship"); // brother service sheet
    var fellowshipSheetP = SpreadsheetApp.openById(pledgeSheetID).getSheetByName("Fellowship Details"); // pledge service sheet
  }
  else{
    var fellowshipSheet = SpreadsheetApp.openById(brotherSheetID).getSheetByName("OtherFellowships"); // changing sheet for soup kitchens
    var fellowshipSheetP = SpreadsheetApp.openById(pledgeSheetID).getSheetByName("Fellowship Details II"); // changing sheet for soup kitchens
  }
    
  var secondRowRange = fellowshipSheet.getRange("2:2"); // gives entire row of service project names
  var secondRowRangeP = fellowshipSheetP.getRange("2:2"); // gives entire row of service project names
  
  var marker = typeTable[type];
  var columnNum = findColumnNum(secondRowRange, marker);
  var columnNumP = findColumnNum(secondRowRangeP, marker);
 
  
  
  
  //-------------------------------------//
  // has determined what kind of project //
  // and isolated last column in section //
  //-------------------------------------//
  
  
  //create new column and fill in the name of the service project
  
  if (attendees != "") {                                                // if there are people
    fellowshipSheet.insertColumnBefore(columnNum);
    var writetitleRange = fellowshipSheet.getRange(2, columnNum, 1, 1); // isolates 1 cell
    writetitleRange.setValue(title);
  }
  
  
  if (attendeesP != "") {
    fellowshipSheetP.insertColumnBefore(columnNumP);
    var writetitleRangeP = fellowshipSheetP.getRange(2, columnNumP, 1, 1);
    writetitleRangeP.setValue(title);
  }
  
  
  //-----------------------------------//
  // created new column for project    //
  //-----------------------------------//
    
  names = parseNames(attendees);
  Logger.log("Names: " + names);
  namesP = parseNames(attendeesP);
  Logger.log("NamesP: " + namesP);

  
  //-----------------------------------//
  // names[] contains all names of     //
  // attendees, format: "last; first"  //
  //-----------------------------------//
  
  
  names = removeHost(names);
  namesP = removeHost(namesP);
  
  
  //-----------------------------------//
  // names[] contains all names of     //
  // attendees (without hosts)         //
  // format: "last; first"             //
  //-----------------------------------//
  
  
  numPeople = names.length + namesP.length + host.length;
  
  
  if(type == "Service Planned" || type == "Service Random" || type == "Leadership Random"){
    var namesSurvey = parseNames("");
    var hostSurvey = [];
    
    // take host names (if any) out of namesSurvey and put into hostSurvey
    var hs1 = namesSurvey.indexOf(host[0]);
    if(hs1 != -1) {
      hostSurvey = hostSurvey.concat(namesSurvey.splice(hs1, 1))
    }
    
    var hs2 = namesSurvey.indexOf(host[1]);
    if(hs2 != -1) {
      hostSurvey = hostSurvey.concat(namesSurvey.splice(hs2, 1));
    }
    
    
    // input points for everyone as if no one took the survey
    survey = "No";
    calculatePointsAndDetermineHosting();
    var regPoints = points[0];
    var hostPoints = points[1];
    
    writePoints(names, fellowshipSheet, columnNum, regPoints, 0);
    writePoints(namesP, fellowshipSheetP, columnNumP, regPoints, 0);
    writePoints(host, fellowshipSheet, columnNum, hostPoints, 1);
    writePoints(host, fellowshipSheetP, columnNumP, hostPoints, 1);
    
    
    // change the points of the people who did take the survey
    survey = "Yes";
    calculatePointsAndDetermineHosting();
    regPoints = points[0];
    hostPoints = points[1];
    
    writePoints(namesSurvey, fellowshipSheet, columnNum, regPoints, 0);
    writePoints(namesSurvey, fellowshipSheetP, columnNumP, regPoints, 0);
    writePoints(hostSurvey, fellowshipSheet, columnNum, hostPoints, 1);
    writePoints(host, fellowshipSheetP, columnNumP, hostPoints, 1);
  }
  else{
    calculatePointsAndDetermineHosting();
    var regPoints = points[0];
    var hostPoints = points[1];
    
    writePoints(names, fellowshipSheet, columnNum, regPoints, 0);
    writePoints(namesP, fellowshipSheetP, columnNumP, regPoints, 0);
    writePoints(host, fellowshipSheet, columnNum, hostPoints, 1);
    writePoints(host, fellowshipSheetP, columnNumP, hostPoints, 1);
  }
  
  
  
  
  //set the host column value to YES
  fellowshipSheet = SpreadsheetApp.openById(brotherSheetID).getSheetByName("Fellowship");
  fellowshipSheetP = SpreadsheetApp.openById(pledgeSheetID).getSheetByName("Fellowship Details"); // pledge fellowship sheet

  
  writeHost(fellowshipSheet, 0);
  writeHost(fellowshipSheetP, 1);
    
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
    //Logger.log(row);
  }
  
  // Resubmit the form with the data from the selected row
  var currentData = sheet.getActiveRange();
  update({values: currentData.getValues()[0]});
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
  var entries = [{
    name : "Submit Data",
    functionName : "readRows"
  }];
  sheet.addMenu("Script Center Menu", entries);
};
