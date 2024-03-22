// Indexes need to match whatever sheet you are using
var emailIndex = 0; /* The email of the individual you are managing 1:1s with */
var cadenceIndex = 6; /* The desired cadence, see cadenceDictionary below */
var lastOneonOnIndex = 9; /* The date you had your last one on one */
var lastOneonOneThreshold = 10; /* The column where the script will flag whether that date is "Good", "Close", "Overdue", 
                                governed by the offset periods in cadenceDictionary */
var lastMeetingIndex = 12; /* The date of any last meeting in which you and the individual were in today */
var lastMeetingTitleIndex = 13; /* The title of that last meeting */
var emailOfCalendarOwner = 'your@email.com'; /* The owner of the calendar */
 
// Total Period Days From Today represents the entire duration of the cadence, Yellow Offset from End of Period, governs "Close"
var cadenceDictionary = {
  "Weekly": [7, 0],
  "Bi-Weekly": [14, 7],
  "Monthly": [30, 7],
  "3 Weeks": [21, 7],
  "6 Weeks": [42, 14],
  "Every 2 Months": [60, 14],
  "Quarterly": [90, 14],
  "Semi-Annually": [180, 21]
};
 
// This is the function you put into your sheet, attached to an image. 
function updateMeetingData() {
  // Reference active sheet or...
  var sheet = SpreadsheetApp.getActiveSheet();

  // For debugging, can point to a specific sheet
  // var sheet = SpreadsheetApp.openByUrl('https://docs.google.com/spreadsheets/<YOUR SHEET>');
  
  // Pull in the data from the sheet
  var data = sheet.getDataRange();
  var values = data.getValues();
  var lastRowIndex = sheet.getLastRow();
 
  var rowDictionaryByEmail = {};
 
  // Loop through all the rows, starting with the second one since I have a header role, creating the dictionary by email
  for (var rowIndex = 1; rowIndex < lastRowIndex; rowIndex++) {
    var currentEmail = (values[rowIndex][emailIndex]).toLowerCase();
    rowDictionaryByEmail[currentEmail] = values[rowIndex];
  }
 
  // Process meetings from the last 6 months on the calendar, can adjust for shorter/longer depending on what you want
  var now = new Date();
  var sinceDate = new Date(now.getTime() - (180 * 24 * 60 * 60 * 1000));
  var events = CalendarApp.getDefaultCalendar().getEvents(sinceDate, now);
 
  // Go through all the events
  for (var eventId = 0; eventId < events.length; eventId++) {
    var currentEvent = events[eventId];
    // Get the guests
    var guests = currentEvent.getGuestList(true);
    // Remove owner 
    for (var guestId = 0; guestId < guests.length; guestId++) {
      if (guests[guestId].getEmail().toLowerCase() == emailOfCalendarOwner) {
        guests.splice(guestId, 1);
      }
    }
 
    // Save off the date of the event
    var meetingDate = currentEvent.getStartTime();
    var meetingTitle = currentEvent.getTitle();
 
    // Loop through each guest
    for (var guestId = 0; guestId < guests.length; guestId++) {
      var guestEmail = guests[guestId].getEmail().toLowerCase();
      var guestStatus = guests[guestId].getGuestStatus();
     
      // Get the row corresponding to this guest
      var guestEmailRow = rowDictionaryByEmail[guestEmail];
      if (guestEmailRow == null) {
        // if the guest isn't someone we are comparing in our 1:1 sheet we just move to the next
        continue;
      }
     
      // If we have a row, let's get the currently stored last meeting dates
      var currentLastOneOnOne = guestEmailRow[lastOneonOnIndex];
      var currentLastMeetingNotOneOnOne = guestEmailRow[lastMeetingIndex];
      // First process 1:1, which means only 1 guest
      if (guests.length == 1 && (currentLastOneOnOne == null || currentLastOneOnOne < meetingDate)) {
          guestEmailRow[lastOneonOnIndex] = meetingDate;
      }
 
      // Next process meetings with more than 1 person, but the guest needs to have accepted
      if (guests.length > 1 && (guestStatus == CalendarApp.GuestStatus.YES || guestStatus == CalendarApp.GuestStatus.OWNER) && (currentLastMeetingNotOneOnOne == null || currentLastMeetingNotOneOnOne < meetingDate)) {
          guestEmailRow[lastMeetingIndex] = meetingDate;
          guestEmailRow[lastMeetingTitleIndex] = meetingTitle;
      }
    }
  }
 
  // Commit updates to the sheet
  data.setValues(values);
 
  // Now we want to process the color coordinated visual column
  for (var rowIndex = 1; rowIndex < lastRowIndex; rowIndex++) {
    var cadence = values[rowIndex][cadenceIndex];
    var lastOneOnOneDate = values[rowIndex][lastOneonOnIndex];

    // If there is no meeting data for this person, clear the column
    if (lastOneOnOneDate == null || lastOneOnOneDate == "" || cadence == null || cadence == "") {
      // Clear
      data.getCell(rowIndex + 1, lastOneonOnIndex + 1).setBackgroundRGB(255, 255, 255);
      data.getCell(rowIndex + 1, lastOneonOneThreshold + 1).setValue("n/a");
      continue;
    }

    // Otherwise, figure out what state the meeting is in, Good, Close or Overdue
    var totalPeriodFromToday = cadenceDictionary[cadence][0];
    var yellowOffest = cadenceDictionary[cadence][1];
    var lastDateOfPeriod = new Date(now.getTime() - (totalPeriodFromToday * 24 * 60 * 60 * 1000));
    var firstDateForYellow = new Date(now.getTime() - ((totalPeriodFromToday - yellowOffest) * 24 * 60 * 60 * 1000));
    
    // getCell is not 0-based
    if (lastOneOnOneDate >= firstDateForYellow) {
      // Green
      data.getCell(rowIndex + 1, lastOneonOnIndex + 1).setBackgroundRGB(212, 250, 200);
      data.getCell(rowIndex + 1, lastOneonOneThreshold + 1).setValue("Good");
    }
    else if (lastOneOnOneDate > lastDateOfPeriod) {
      // Yellow
      data.getCell(rowIndex + 1, lastOneonOnIndex + 1).setBackgroundRGB(232, 235, 52);
      data.getCell(rowIndex + 1, lastOneonOneThreshold + 1).setValue("Close");
    }
    else
    {
      // Red
      data.getCell(rowIndex + 1, lastOneonOnIndex + 1).setBackgroundRGB(230, 160, 145);
      data.getCell(rowIndex + 1, lastOneonOneThreshold + 1).setValue("Overdue");
    }
  }
 
  return;
}
 