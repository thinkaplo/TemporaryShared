// ---------------------------------------------------------------|
// PROJECT:         APLO                                          |
// PROCESS:         Workshop booking                              |
// SCOPE:           Automate event booking based in form answer   |
// CREATOR:         think.aplo@gmail.com                          |
// VERSION:         1.0                                           |
// PUBLISHED:       28/01/2024                                    |
// ---------------------------------------------------------------|

function getLastNonEmptyRow() {
// Change the name of the sheet as needed
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form responses 1');

// Get all values in the sheet
  var values = sheet.getDataRange().getValues();

// Iterate through the values starting from the bottom of the sheet
  for (var i = values.length - 1; i >= 0; i--) {
    for (var j = 0; j < values[i].length; j++) {
  // Check if the cell is not empty
      if (values[i][j] !== "") {
    // Return the index number (add 1 if you want to refer to row number)
        return i;
      }
    }
  }

// If empty row is found, return 0 meaning the file is completely empty
  return 0;
}


function sendInvitation() {
// Get the form responses. Change range when needed.
  var responses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form responses 1");
  var responses = responses.getRange("A:F").getValues();

// Set the calendar ID and event ID. Change calendarId and eventId when needed
  var calendarId = '{your calendar Id}';
  var eventId = '{your event Id}';

// Get the calendar event
  var calendar = CalendarApp.getCalendarById(calendarId);
  var event = calendar.getEventById(eventId);

  
// Get the email address from the form responses
  var row = getLastNonEmptyRow();
  var emailAddress = responses[row][3]; // replace with the index of the email field in your form

// Add the email as a guest to the calendar event. keep in mind addGuest() method only works with Gmail accounts.
  event.addGuest(emailAddress);

// Send a calendar invitation. Develop this method if email recipient is different from Gmail.
    
// Send an email confirmation. Change subject and message when needed.
  var subject = '{your message subject}.';
  var message = '{your message body}';

  MailApp.sendEmail(emailAddress, subject, message);
  
}
