var signupURL;
var signupResponsesId;
var attendeesURL;
var calendarURL;
var folder;

function onSubmit(form) {
  //TODO: These indices may have to change based on the form you create
  Logger.log(form.values);
  createSignupForm(form.values[1], form.values[2]);
  createListOfAttendees(form.values[1], form.values[2], form.values[3], form.values[4], form.values[5], form.values[6], form.values[8]);
  createCalendarEvent(form.values[1], form.values[2], form.values[3], form.values[4], form.values[5], form.values[6], form.values[8], form.values[9], form.values[10], form.values[7]);
}

function createSignupForm(projectName, date) {
  var signupForm = FormApp.create(projectName + ' - ' + date);
  signupForm.addTextItem().setTitle('Name').setRequired(true);
  signupForm.addTextItem().setTitle('Email address').setRequired(true);
  signupForm.addTextItem().setTitle('Phone Number').setRequired(true);
  signupForm.addMultipleChoiceItem().setTitle('Would you like to be the host?').setChoiceValues(['Yes', 'Yes, If Necessary', 'No']).setRequired(true);
  
  signupURL = signupForm.getPublishedUrl();
  Logger.log(signupForm.getSummaryUrl());
  
  var newSheet = SpreadsheetApp.create(projectName + ' (responses) - ' + date);
  signupForm.setDestination(FormApp.DestinationType.SPREADSHEET, newSheet.getId());
  
  Logger.log('sheet url: ' + newSheet.getUrl());
  Logger.log('sheet id: ' + newSheet.getId());
  signupResponsesId = newSheet.getUrl().split('/')[5];
}

function createListOfAttendees(projectName, date, start, end, location, numVols) {
  var attendeesDraft = SpreadsheetApp.openById('1DpviVgvcNMVRdgi5vuKHKXsu54feBXUkVe6I0UySsaw');
  var attendeesFinal = attendeesDraft.copy(projectName + ' - ' + date);
  var attendeesSheet = attendeesFinal.getActiveSheet();
  attendeesSheet.getRange(1,4,1,1).setValue(projectName + ' - ' + date);
  attendeesSheet.getRange(2,4,1,1).setValue(start);
  attendeesSheet.getRange(2,5,1,1).setValue(end);
  attendeesSheet.getRange(3,4,1,1).setValue(location);
  attendeesSheet.getRange(4,4,1,1).setValue(numVols);
  attendeesSheet.getRange(10,9,1,1).setValue(signupResponsesId);
  
  attendeesURL = attendeesFinal.getUrl();
}

function createCalendarEvent(projectName, date, start, end, location, numVols, logistics, transportation, about, type) {
  var formattedStartDate = formatDate(date, start);
  var formattedEndDate = formatDate(date, end);
  var startDate = new Date(formattedStartDate);
  var endDate = new Date(formattedEndDate);
  var cal = CalendarApp.getCalendarById('') //TODO: Add in service calendar Id here
  var event = cal.createEvent(projectName + ' - ' + date, startDate, endDate);
  event.setDescription(getDescription(type, numVols, logistics, transportation, about, signupURL, attendeesURL));
  event.setLocation(location);
}

function getDescription(type, numVols, logistics, transportation, about, sLink, aLink) {
  var signup = "<p> <a href=\"" + sLink + "\">Signup form</a> </p>";
  var attendees = "<p> <a href=\"" + aLink + "\">List of Attendees</a> </p>";
  var end = "<p> ___________ <br> If you need to withdraw from a service project, you can use the form below as long as it is at least 48 hours before the time of the service project. If you do withdraw from the project, please find someone to take your place. Otherwise, contact the host or vpservice@upennapo.org. <br> Withdrawal form: bitly.com/servicewithdraw </p>"
  if (type === "Feast Incarnate" || type === "Books through Bars" || type === "UCHC") {
    return getRecurringDescription(type) + signup + attendees + end;
  }
  var vols = "<p> <strong> Volunteers: " + numVols + "</p>";
  var lSection = "<p> <strong>Logistics: </strong> <br> " + logistics + "</p>";
  var tSection = "<p> <strong>Transportation: <br> " + transportation + "</strong> </p>";
  var aSection = "<p> <strong>About: </strong> <br> " + about + "</p>";
  return vols + lSection + tSection + aSection + signup + attendees + end;
}

//TODO: This function can maybe be exported to a spreadsheet in the future if a lot of projects end up getting added
function getRecurringDescription(project) {
  if (project.equals("Feast Incarnate")) {
    return "<b>LOGISTICS:</b> <br>" + 
           "Show up at 5:30 PM to help with food preparation, meal distribution, and set-up/clean-up. Do not forget to sign in with the host.<br>" +
           "<b>Volunteers: 2</b><br><b>Travel Time: 10 minutes by walking</b><br>" + 
           "When you arrive at the site, find the entrance and then the kitchen (it’s not hard to find at each site, but often you’ll have to go " +
           "around the church and use a back door) and there will be other volunteers waiting to guide you. Just introduce yourself and where you’re from and get to work<br>" +
           "<b>ABOUT:</b><br>Feast Incarnate, a weekly hospitality ministry of University Lutheran Church of the Incarnation, started in November of 1988. It was a direct response " +
           "to the situation of two members of the parish at that time, who had been diagnosed with HIV, and who had been rejected by friends, families and partners. The evening was planned as a safe environment, free from prejudice and discrimination. " +
           "An invitation was issued to the HIV/AIDS community to come and share an evening of food and fellowship. Every Tuesday since, University Lutheran has served a meal for the HIV/AIDS community of University City in Philadelphia.<br>"+
           "As of mid-2015, Feast Incarnate is an independent 501(c)(3) non-profit organization and operates under the direction of an executive board. Feast Incarnate has expanded its reach to all those in need of a meal within our community. "+
           "The organization employs the help of volunteers from around the region, including service fraternities, local congregations, and members of the church, who help prepare meals, greet visitors, serve food, and clean up the church every week.";

  } else if (project.equals("UCHC")) {
    return "<b>LOGISTICS:</b><br>Show up at 5:30 PM to help with food preparation, meal distribution, and set-up/clean-up. Do not forget to sign in with the host.<br>"+
           "<b>Volunteers: 3</b><br>When you arrive at the site, find the entrance and then the kitchen (it’s not hard to find at each site, but often you’ll have to go around the church and use a back door) and there will be other volunteers waiting to guide you. Just introduce yourself and where you’re from and get to work.<br>"+
           "<b>ABOUT:</b>The University City Hospitality Coalition (UCHC) not only provides a hot meal to the homeless and hungry five nights a week, but it strives to assist people in as many ways as possible. &gt;Refer guests to additional housing, medical, legal, and job training services, and provide a mailing address for those without one. &gt;Don’t forget, you’re there to not only serve the meals but also to interact with the community members you are serving. <br>The full meal schedule is listed here: http://uchc.phillycharities.org/";
  } else {
    return "<b>Description of Event:</b><br>Books Through Bars is an all-volunteer, non-profit organization dedicated to educating prisoners through the donation and distribution of books to prisoners in seven states. Each month they send out approximately 2100 books to 700 people. They aim to reverse the devastating effects that injustice and incarceration has on individuals, families, and communities. <br>"+
           "You will be helping out by reading the letters sent by the prisoners, selecting appropriate books based on prisoners&#39; interests, and packaging the books. If it is your first time at Books Through Bars PLEASE TELL THEM so they can tell you more about the organization and help you out. <br>"+
           "<a href=\"https://drive.google.com/file/d/0B0IjNxyzqN0_cUVQeW9oVTdFdUk/view\" target=\"_blank\">More information</a><br><b>Logistics:</b><br>Volunteers: 5 <br>Travel time: 20 minutes walking <br>If it is your first time, you will have to participate in a 30-minute volunteer training session on how to pick books or pack books.<br><b> NOTE: You must be 18 or older to volunteer at this project!</b>";
  }
}

function formatDate(date, time) {
  var splitDate = date.split('/');
  var months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var month = months[parseInt(splitDate[0]) - 1];
  var day = splitDate[1];
  var year = splitDate[2];
  return month + ', ' + day + ' ' + year + ' ' + time;
}

function setFolder() {
  folder = DriveApp.getFoldersByName("Service Projects").next();
}
