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
        // Return the row number (add 1 since arrays are zero-indexed)
        return i;
      }
    }
  }

  // If empty row is found, return 0 meaning the file is completely empty
  return 0;
}


function sendInvitation() {
  // Get the form responses
  var responses = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form responses 1");
  var responses = responses.getRange("A:F").getValues();

  // Set the calendar ID and event ID
  var calendarId = 'think.aplo@gmail.com';
  var eventId = '7o34lo5tq1d4ecmi1noghgoo5e';

  // Get the calendar event
  var calendar = CalendarApp.getCalendarById(calendarId);
  var event = calendar.getEventById(eventId);

  

    // Get the email address from the form responses
  var row = getLastNonEmptyRow();
  var emailAddress = responses[row][3]; // replace with the index of the email field in your form

    // Add the email as a guest to the calendar event
  event.addGuest(emailAddress);

    // Send a calendar invitation
    

    // Send an email confirmation
  var subject = 'Beyond the Basics: Mastering Google Sheets & Workspace Tools.';
  var message = 'Thank you for completing the subscription form. \nKindly confirm your attendance by accepting the invitation we have sent to your calendar. \nWe look forward to seeing you on Saturday with a positive and eager-to-learn mindset!';

  MailApp.sendEmail(emailAddress, subject, message);
 
  
}
