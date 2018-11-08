//TODO: These indices may have to change based on the form you create
var namePosition = 1;
var startDatePosition = 2;
var endDatePosition = 3;
var startTimePosition = 4;
var endTimePosition = 5;
var descriptionPosition = 6;
var contactInformationPosition = 7;
var hostNamesPosition = 8;
var repeatsPosition = 9;
var locationPosition = 10;
var calendarId=""; //TODO: Use fellowship calendar url here
var description;

/* 
START HERE
This function is called using the triggers above every time the form is submitted,
and creates a calendar event with the submitted data
*/
function createFellowship(data) {
  var formData = data.values;
  Logger.log("Form data is: " + formData);
  description = getDescription(formData);
  createEvent(formData);
  Logger.log("Successfully created fellowship " + formData[namePosition]);
}

//Computes the full event description based on the form data
function getDescription(formData) {
  var desc = formData[descriptionPosition];
  var contactInfo = formData[contactInformationPosition];
  var hostNames = formData[hostNamesPosition];
  
  return desc + "\n \n<strong>Host(s):</strong> \n" + hostNames + "\n \n <strong>Contact Information:</strong> \n" + contactInfo;
}

//Creates the calendar event
function createEvent(formData) {
  var calendar = CalendarApp.getCalendarById(calendarId);
  var startDate = getDate(formData[startDatePosition], formData[startTimePosition]);
  var endDate = getDate(formData[endDatePosition], formData[endTimePosition]);
  var numRepeats = 4; //limit to 4 to make sure hosts don't forget and to prevent infinite repetition
  
  //no repeat
  if (formData[repeatsPosition] == "No") {
    var event = calendar.createEvent(formData[namePosition], startDate, endDate);
    event.setDescription(description);
    event.setLocation(formData[locationPosition]);
    Logger.log("Created event on " + startDate);
    
  //monthly repeat
  } else if (formData[repeatsPosition].indexOf("Monthly") > 0) {
    var monthlyRecurrence = CalendarApp.newRecurrence().addMonthlyRule().times(numRepeats);
    var events = calendar.createEventSeries(formData[namePosition], startDate, endDate, monthlyRecurrence);
    events.setDescription(description);
    events.setLocation(formData[locationPosition]);
    Logger.log("Created " + numRepeats + " monthly repeating events starting on " + startDate);
    
  //biweekly repeat
  } else if (formData[repeatsPosition].indexOf("Every other week") > 0) {
    var biweeklyRecurrence = CalendarApp.newRecurrence().addWeeklyRule().times(numRepeats).addWeeklyExclusion().interval(2);
    var events = calendar.createEventSeries(formData[namePosition], startDate, endDate, biweeklyRecurrence);
    events.setDescription(description);
    events.setLocation(formData[locationPosition]);
    Logger.log("Created " + numRepeats + " biweekly repeating events starting on " + startDate);
    
  //weekly repeat
  } else {
    var weeklyRecurrence = CalendarApp.newRecurrence().addWeeklyRule().times(numRepeats);
    var events = calendar.createEventSeries(formData[namePosition], startDate, endDate, weeklyRecurrence);
    events.setDescription(description);
    events.setLocation(formData[locationPosition]);
    Logger.log("Created " + numRepeats + " weekly repeating events starting on " + startDate);
  }
}

//Gets the start and end dates in the format required by google apps script
function getDate(date, time) {
  var dateObj = new Date(date);
  var timeData = getTimeData(time);
  dateObj.setHours(timeData[0]);
  dateObj.setMinutes(timeData[1]);
  return dateObj;
}

//Parses the time string to get the hour component in 24 hour time
function getTimeData(time) {
  var split = time.split(":");
  var hour = split[0];
  var minute = split[1];
  //if pm, add 12
  if (time.indexOf("PM") > 0) {
    hour = parseInt(hour) + 12;
  }
  return [hour, minute];
}
