/*
 * Big-Little Reporting Form Script
 * Last Updated: 3/6/2018 by Will Glisson
 *
 * The function "update" is triggered when the form is submitted. (See Current project's triggers in form menu)
 * The argument to update (e) is the data from the form.
 */

// This is the spreadsheet for the current semester
var bigLittleSpreadsheet = SpreadsheetApp.openById(""); //TODO: Add in big little spreadsheet id here

// This table maps the submitted event type to the tab in the Big-Little spreadsheet
var sheetTable = {
  "REQUIRED: Fellowship Co-Hosting (2 pts)" : "FellowshipHosting",
  "REQUIRED: Lineage Hangout... ≥3 lineage members (1.5 pts)" : "LineageHangout",
  "REQUIRED: First February Private Hangout (2 pts)" : "PrivateHangout",
  "REQUIRED: First March Private Hangout (2 pts)" : "PrivateHangout",
  "REQUIRED: Service Co-Attendance (2 pts)" : "ServiceCoAttendance",
  "Extra Private Hangout (1 pt/day)" : "PrivateHangout",
  "Fellowship attendance (1 pt)": "Misc",
  "APO Event... fundraiser, leadership event, etc. (2 pts)": "Misc",
  "Miscellaneous, ie. extraordinary act of support for Big/Little  (pts vary)": "Misc"
};


// Returns the row number of a value (key) in a single column of values (range)
function findRowNum(range, key) {
  var data = range.getValues();
  for (var i = 0; i < data.length; i++) {
    Logger.log("Row: " + data[i][0]);
    
    // Return i+1 because sheet ranges are 1-indexed while JS arrays are 0-indexed
    if (key == data[i][0])
      return i + 1;
  }
  return -1;
}

function addServiceAttendance(sheet, rowNum, date, desc) {
  if (sheet.getRange(rowNum, 6, 1, 1).getValue() == "") {
    sheet.getRange(rowNum, 6, 1, 1).setValue(date);
    sheet.getRange(rowNum, 7, 1, 1).setValue(desc);
  } else {
    var extra = sheet.getRange(rowNum, 8, 1, 1);
    if (extra.getValue() == "") {
      extra.setValue("1");
    } else {
      extra.setValue(parseInt(extra.getValue()) + 1);
    }
  }
}

function addFellowshipHosting(sheet, rowNum, fellowshipDate, fellowshipDesc) {
  Logger.log("row num: " + rowNum);
  if (sheet.getRange(rowNum, 6, 1, 1).getValue() == "") {
    sheet.getRange(rowNum, 6, 1, 1).setValue(fellowshipDate);
    sheet.getRange(rowNum, 7, 1, 1).setValue(fellowshipDesc);
  } else {
    var extraFellowships = sheet.getRange(rowNum, 8, 1, 1);
    if (extraFellowships.getValue() == "") {
      extraFellowships.setValue("1");
    } else {
      extraFellowships.setValue(parseInt(extraFellowships.getValue()) + 1);
    }
  }
}

function addLineageHangouts(sheet, rowNum, linHangDate, linHangDesc) {
  if (sheet.getRange(rowNum, 6, 1, 1).getValue() == "") {
    sheet.getRange(rowNum, 6, 1, 1).setValue(linHangDate);
    sheet.getRange(rowNum, 7, 1, 1).setValue(linHangDesc);
  } else {
    var extraLinHangouts = sheet.getRange(rowNum, 8, 1, 1);
    if (extraLinHangouts.getValue() == "") {
      extraLinHangouts.setValue("1");
    } else {
      extraLinHangouts.setValue(parseInt(extraLinHangouts.getValue()) + 1);
    }
  }
}

function addPrivateHangouts(sheet, rowNum, privHangDate, privHangDesc, privHangMonth) {
  // These index vars can be updated depending on the format of the sheet
  var febStartCol = 8;
  var marStartCol = 10;
  var extraStartCol = 12;
  
  if ((privHangMonth == "Feb") && (sheet.getRange(rowNum, febStartCol, 1, 1).getValue() == "")) {
    sheet.getRange(rowNum, febStartCol, 1, 1).setValue(privHangDate);
    sheet.getRange(rowNum, febStartCol + 1, 1, 1).setValue(privHangDesc);
  } else if ((privHangMonth == "Mar") && (sheet.getRange(rowNum, marStartCol, 1, 1).getValue() == "")) {
    sheet.getRange(rowNum, marStartCol, 1, 1).setValue(privHangDate);
    sheet.getRange(rowNum, marStartCol + 1, 1, 1).setValue(privHangDesc);
  } else {
    var extraPrivateHangouts = sheet.getRange(rowNum, extraStartCol, 1, 1);
    if (extraPrivateHangouts.getValue() == "") {
      extraPrivateHangouts.setValue("1");
    } else {
      extraPrivateHangouts.setValue(parseInt(extraPrivateHangouts.getValue()) + 1);
    }
  }
}

function addFellowshipAttendance(sheet, rowNum, fellowshipDate, fellowshipDesc) {
  var colNum = sheet.getDataRange().getValues()[0].indexOf("Fellowship Attendance") + 1;
  var fellowshipCell = sheet.getRange(rowNum, colNum, 1, 1);
  if (fellowshipCell.getValue() == "") {
    fellowshipCell.setValue(fellowshipDate + " - " + fellowshipDesc);
  } else {
    fellowshipCell.setValue(fellowshipCell.getValue() + "\n" + fellowshipDate + " - " + fellowshipDesc);
  }
  colNum = sheet.getDataRange().getValues()[0].indexOf("FellowshipAttendance") + 1;
  fellowshipCell = sheet.getRange(rowNum, colNum, 1, 1);
  fellowshipCell.setValue(parseInt(fellowshipCell.getValue()) + 1);
}

function addAPOEvent(sheet, rowNum, apoEventDate, apoEventDesc) {
  var colNum = sheet.getDataRange().getValues()[0].indexOf("APO Event") + 1;
  var apoEventCell = sheet.getRange(rowNum, colNum, 1, 1);
  if (apoEventCell.getValue() == "") {
    apoEventCell.setValue(apoEventDate + " - " + apoEventDesc);
  } else {
    apoEventCell.setValue(apoEventCell.getValue() + "\n" + apoEventDate + " - " + apoEventDesc);
  }
  colNum = sheet.getDataRange().getValues()[0].indexOf("APOEvent") + 1;
  apoEventCell = sheet.getRange(rowNum, colNum, 1, 1);
  apoEventCell.setValue(parseInt(apoEventCell.getValue()) + 1);
}

// This function is run every time the Big-Little Reporting Form is submitted.
function update(e) {
  var pair = e.values[1]; // 1: Big-Little pair
  var type = e.values[2]; // 2: Event type
  var description = e.values[3];
  var date = e.values[4];
  
  // get the appropriate sheet based on the type of event
  var sheet = bigLittleSpreadsheet.getSheetByName(sheetTable[type]);
  Logger.log("type: " + sheetTable[type]);
  
  // get row number of current pair using first column of sheet (names of pairs)
  var rowNum = findRowNum(sheet.getRange("A:A"), pair);
  Logger.log("pair: " + pair)
  Logger.log("rownum: " + rowNum);
  
  switch (type) {
    case "REQUIRED: Fellowship Co-Hosting (2 pts)":
      Logger.log("fellowship co-host case");
      addFellowshipHosting(sheet, rowNum, date, description);
      break;
    case "REQUIRED: Lineage Hangout... ≥3 lineage members (1.5 pts)":
      Logger.log("lin case");
      addLineageHangouts(sheet, rowNum, date, description);
      break;
    case "REQUIRED: Service Co-Attendance (2 pts)":
      Logger.log("First service case");
      addServiceAttendance(sheet, rowNum, date, description);
      break;
    case "REQUIRED: First February Private Hangout (2 pts)":
      Logger.log("First February Private Hangout case");
      addPrivateHangouts(sheet, rowNum, date, description, "Feb");
      break;
    case "REQUIRED: First March Private Hangout (2 pts)":
      Logger.log("First March Private Hangout case");
      addPrivateHangouts(sheet, rowNum, date, description, "Mar");
      break;
    case "Extra Private Hangout (1 pt/day)":
      Logger.log("extra private hangout case");
      addPrivateHangouts(sheet, rowNum, date, description, "none");
      break;
    case "Fellowship attendance (1 pt)":
      Logger.log("fellowship attendance case");
      addFellowshipAttendance(sheet, rowNum, date, description);
      break;
    case "APO Event... fundraiser, leadership event, etc. (2 pts)":
      Logger.log("APO Event case");
      addAPOEvent(sheet, rowNum, date, description);
      break;
    default:
      Logger.log(type);
      break;
  }
}

// TODO: This has no trigger, determine if it's necessary
// This function runs whenever a value is changed on the spreadsheet
function onEdit(e) {
  if (e.range.getNumColumns() != 1 || e.range.getColumn() != 1) {
    return;
  }
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var editRowNum = e.range.getRow();
  var pair = activeSheet.getRange(editRowNum, 5, 1, 1).getValue();
  var type = activeSheet.getRange(editRowNum, 6, 1, 1).getValue();
  
  // get the appropriate sheet based on the type of event
  var sheet = bigLittleSpreadsheet.getSheetByName(sheetTable[type]);
  
  // get row number of current pair using first column of sheet (names of pairs)
  var rowNum = findRowNum(sheet.getRange("A:A"), pair);
  
  if ((e.oldValue == "No" || e.oldValue == "Pending..." || e.oldValue == "Duplicate") && e.value == "Yes") {
    Logger.log("in if");
    switch (type) {
      /* case "2 REQUIRED: Big-Little Event (3 pts)":
      case "Extra Big-Little Event (5 pts)":
        // 5: Big-little event
        addBigLittleEvents(sheet, rowNum, activeSheet.getRange(editRowNum, 9, 1, 1).getValue());
        break; */
      case "REQUIRED: Fellowship Co-Hosting (2 pts)":
        // 13: Date of fellowship, 7: Description of fellowship
        addFellowshipHosting(sheet, rowNum, activeSheet.getRange(editRowNum, 8, 1, 1).getValue(),
                             activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
        break;
      case "REQUIRED: Lineage Hangout... ≥3 lineage members (1.5 pts)":
        // 14: Date of lineage hangout, 11: Description of lineage hangout
        addLineageHangouts(sheet, rowNum, activeSheet.getRange(editRowNum, 8, 1, 1).getValue(),
                           activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
        break;
      case "REQUIRED: First February Private Hangout (2 pts)":
        addPrivateHangouts(sheet, rowNum, activeSheet.getRange(editRowNum, 8, 1, 1).getValue(),
                           activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
      case "Extra Private Hangout (1 pt/day)":
        // 12: Date of private hangout, 3: Description of private hangout
        addPrivateHangouts(sheet, rowNum, activeSheet.getRange(editRowNum, 8, 1, 1).getValue(),
                           activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
        break;
      case "Fellowship attendance (1 pt)":
        // 13: Date of fellowship, 7: Description of fellowship
        var fellowshipDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        fellowshipDate = (parseInt(1 + fellowshipDate.getMonth()) + "/" + parseInt(fellowshipDate.getDate()) + "/" + parseInt(fellowshipDate.getYear()));
        addFellowshipAttendance(sheet, rowNum, fellowshipDate, activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
        break;
      case "APO Event... fundraiser, leadership event, etc. (2 pts )":
        // 15: Date of APO event, 6: Description of APO event
        var apoEventDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        apoEventDate = (parseInt(1 + apoEventDate.getMonth()) + "/" + parseInt(apoEventDate.getDate()) + "/" + parseInt(apoEventDate.getYear()));
        addAPOEvent(sheet, rowNum, apoEventDate,
                    activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
        break;
      case "Miscellaneous, ie. extraordinary act of support for Big/Little  (pts vary)":
        e.range.setValue("Yes (Manual)");
        break;
      default:
        break;
    }
  /*}  else if (e.oldValue == "Yes" && e.value == "Duplicate" &&
             (type == "2 REQUIRED: Big-Little Event (3 pts)" || type == "Extra Big-Little Event (5 pts)")) {
    return; */
  } else if (e.oldValue == "Yes" && (e.value == "No" || e.value == "Pending..." || e.value == "Duplicate")) {
    Logger.log("in else if");
    switch (type) {
      /* case "2 REQUIRED: Big-Little Event (3 pts)":
      case "Extra Big-Little Event (5 pts)":
        // 5: Big-little event
        var blEvent = activeSheet.getRange(editRowNum, 9, 1, 1).getValue();
        if (blEvent == "Hunt for the Infinity Stones") {
          sheet.getRange(rowNum, 6, 1, 1).setValue("");
        } else if (blEvent == "Heroes vs. Villains") {
          sheet.getRange(rowNum, 7, 1, 1).setValue("");
        } else if (blEvent == "Souper Cook-off") {
          sheet.getRange(rowNum, 8, 1, 1).setValue("");
        }
        break; */
      case "REQUIRED: Fellowship Co-Hosting (2 pts)":
        // 13: Date of fellowship, 7: Description of fellowship
        var fellowshipDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        var fellowshipDesc = activeSheet.getRange(editRowNum, 7, 1, 1).getValue();
        var extraFellowships = sheet.getRange(rowNum, 8, 1, 1);
        if (sheet.getRange(rowNum, 7, 1, 1).getValue() == fellowshipDesc) {
          sheet.getRange(rowNum, 6, 1, 1).setValue("");
          sheet.getRange(rowNum, 7, 1, 1).setValue("");
          if (parseInt(extraFellowships.getValue()) > 0) {
            var data = activeSheet.getDataRange().getValues();
            for (var i = 0; i < data.length; i++) {
              if (data[i][2] == pair &&
                  data[i][3] == "REQUIRED: Fellowship Co-Hosting (2 pts)" &&
                  data[i][4] != fellowshipDesc) {
                sheet.getRange(rowNum, 6, 1, 1).setValue(data[i][5]);
                sheet.getRange(rowNum, 7, 1, 1).setValue(data[i][4]);
                extraFellowships.setValue(parseInt(extraFellowships.getValue() - 1));
                if (extraFellowships.getValue() == "0") { extraFellowships.setValue(""); }
                return;
              }     
            }
          }
        } else {
          extraFellowships.setValue(parseInt(extraFellowships.getValue()) - 1);
          if (parseInt(extraFellowships.getValue()) <= 0) { extraFellowships.setValue(""); }
        }
        break;
      case "REQUIRED: Lineage Hangout... ≥3 lineage members (1.5 pts)":
        // 14: Date of lineage hangout, 11: Description of lineage hangout
        var linHangDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        var linHangDesc = activeSheet.getRange(editRowNum, 7, 1, 1).getValue();
        var extraLinHangs = sheet.getRange(rowNum, 8, 1, 1);
        if (sheet.getRange(rowNum, 7, 1, 1).getValue() == linHangDesc) {
          sheet.getRange(rowNum, 6, 1, 1).setValue("");
          sheet.getRange(rowNum, 7, 1, 1).setValue("");
          if (parseInt(extraLinHangs.getValue()) > 0) {
            var data = activeSheet.getDataRange().getValues();
            for (var i = 0; i < data.length; i++) {
              if (data[i][2] == pair &&
                  data[i][3] == "REQUIRED: Lineage Hangout... ≥3 lineage members (1.5 pts)" &&
                  data[i][4] != linHangDesc) {
                sheet.getRange(rowNum, 6, 1, 1).setValue(data[i][5]);
                sheet.getRange(rowNum, 7, 1, 1).setValue(data[i][4]);
                extraLinHangs.setValue(parseInt(extraLinHangs.getValue() - 1));
                if (extraLinHangs.getValue() == "0") { extraLinHangs.setValue(""); }
                return;
              }     
            }
          }
        } else {
          extraLinHangs.setValue(parseInt(extraLinHangs.getValue()) - 1);
          if (parseInt(extraLinHangs.getValue()) <= 0) { extraLinHangs.setValue(""); }
        }
        break;
      case "REQUIRED: First February Private Hangout (2 pts)":
        // 12: Date of private hangout, 3: Description of private hangout
        var privHangDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        var privHangDesc = activeSheet.getRange(editRowNum, 7, 1, 1).getValue();
        var extraPrivHangs = sheet.getRange(rowNum, 8, 1, 1);
        if (sheet.getRange(rowNum, 7, 1, 1).getValue() == privHangDesc) {
          sheet.getRange(rowNum, 6, 1, 1).setValue("");
          sheet.getRange(rowNum, 7, 1, 1).setValue("");
          if (parseInt(extraPrivHangs.getValue()) > 0) {
            var data = activeSheet.getDataRange().getValues();
            for (var i = 0; i < data.length; i++) {
              if (data[i][2] == pair &&
                  (data[i][3] == "REQUIRED: First February Private Hangout (2 pts)" || data[i][3] == "Extra Private Hangout (1 pt/day)") &&
                  data[i][4] != privHangDesc) {
                sheet.getRange(rowNum, 6, 1, 1).setValue(data[i][5]);
                sheet.getRange(rowNum, 7, 1, 1).setValue(data[i][4]);
                extraPrivHangs.setValue(parseInt(extraPrivHangs.getValue() - 1));
                if (extraPrivHangs.getValue() == "0") { extraPrivHangs.setValue(""); }
                return;
              }     
            }
          }
        } else {
          extraPrivHangs.setValue(parseInt(extraPrivHangs.getValue()) - 1);
          if (parseInt(extraPrivHangs.getValue()) <= 0) { extraPrivHangs.setValue(""); }
        }
      case "Extra Private Hangout (1 pt/day)":
        // 12: Date of private hangout, 3: Description of private hangout
        var privHangDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        var privHangDesc = activeSheet.getRange(editRowNum, 7, 1, 1).getValue();
        var extraPrivHangs = sheet.getRange(rowNum, 8, 1, 1);
        if (sheet.getRange(rowNum, 7, 1, 1).getValue() == privHangDesc) {
          sheet.getRange(rowNum, 6, 1, 1).setValue("");
          sheet.getRange(rowNum, 7, 1, 1).setValue("");
          if (parseInt(extraPrivHangs.getValue()) > 0) {
            var data = activeSheet.getDataRange().getValues();
            for (var i = 0; i < data.length; i++) {
              if (data[i][2] == pair &&
                  (data[i][3] == "REQUIRED: First February Private Hangout (2 pts)" || data[i][3] == "Extra Private Hangout (1 pt/day)") &&
                  data[i][4] != privHangDesc) {
                sheet.getRange(rowNum, 6, 1, 1).setValue(data[i][5]);
                sheet.getRange(rowNum, 7, 1, 1).setValue(data[i][4]);
                extraPrivHangs.setValue(parseInt(extraPrivHangs.getValue() - 1));
                if (extraPrivHangs.getValue() == "0") { extraPrivHangs.setValue(""); }
                return;
              }     
            }
          }
        } else {
          extraPrivHangs.setValue(parseInt(extraPrivHangs.getValue()) - 1);
          if (parseInt(extraPrivHangs.getValue()) <= 0) { extraPrivHangs.setValue(""); }
        }
        break;
      case "Fellowship attendance (1 pt)":
        // 13: Date of fellowship, 7: Description of fellowship
        var fellowshipDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        fellowshipDate = (parseInt(1 + fellowshipDate.getMonth()) + "/" + parseInt(fellowshipDate.getDate()) + "/" + parseInt(fellowshipDate.getYear()));
        var fellowshipDesc = activeSheet.getRange(editRowNum, 7, 1, 1).getValue();
        var colNum = sheet.getDataRange().getValues()[0].indexOf("Fellowship Attendance") + 1;
        var fellowshipCell = sheet.getRange(rowNum, colNum, 1, 1);
        var allFellowships = fellowshipCell.getValue().toString();
        var fellowshipString = fellowshipDate + " - " + fellowshipDesc;
        if (allFellowships == fellowshipString) {
          fellowshipCell.setValue("");
        } else if (allFellowships.indexOf(fellowshipString) == 0) {
          fellowshipCell.setValue(allFellowships.replace(fellowshipString + "\n", ""));
        } else {
          fellowshipCell.setValue(allFellowships.replace("\n" + fellowshipString, ""));
        }
        colNum = sheet.getDataRange().getValues()[0].indexOf("FellowshipAttendance") + 1;
        fellowshipCell = sheet.getRange(rowNum, colNum, 1, 1);
        fellowshipCell.setValue(parseInt(fellowshipCell.getValue()) - 1);
        break;
      case "APO Event... fundraiser, leadership event, etc. (2 pts)":
        // 15: Date of APO event, 6: Description of APO event
        addAPOEvent(sheet, rowNum, activeSheet.getRange(editRowNum, 8, 1, 1).getValue(),
                    activeSheet.getRange(editRowNum, 7, 1, 1).getValue());
        var apoEventDate = activeSheet.getRange(editRowNum, 8, 1, 1).getValue();
        apoEventDate = (parseInt(1 + apoEventDate.getMonth()) + "/" + parseInt(apoEventDate.getDate()) + "/" + parseInt(apoEventDate.getYear()));
        var apoEventDesc = activeSheet.getRange(editRowNum, 7, 1, 1).getValue();
        var colNum = sheet.getDataRange().getValues()[0].indexOf("APO Event") + 1;
        var apoEventCell = sheet.getRange(rowNum, colNum, 1, 1);
        var allAPOEvents = apoEventCell.getValue().toString();
        var apoEventString = apoEventDate + " - " + apoEventDesc;
        if (allAPOEvents == apoEventString) {
          apoEventCell.setValue("");
        } else if (allAPOEvents.indexOf(apoEventString) == 0) {
          apoEventCell.setValue(allAPOEvents.replace(apoEventString + "\n", ""));
        } else {
          apoEventCell.setValue(allAPOEvents.replace("\n" + apoEventString, ""));
        }
        colNum = sheet.getDataRange().getValues()[0].indexOf("APOEvent") + 1;
        apoEventCell = sheet.getRange(rowNum, colNum, 1, 1);
        apoEventCell.setValue(parseInt(apoEventCell.getValue()) + 1);
        break;
      default:
        break;
    }
  }
}