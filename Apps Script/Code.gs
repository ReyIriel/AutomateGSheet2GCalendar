/**
 * Initial function to add bar-dropdown menu into spreadsheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Tools')
      .addItem('Calendar Sync', 'createCalendarEvents')
      .addToUi();
}

/**
 * Main function.
 * Coloumn index array glossary:
 * [0] Date info (Coloumn A: DATE)
 * [1] Start time (Coloumn B: START)
 * [2] End time (Coloumn C: END)
 * [3] Reminder (Coloumn D: REMINDER - "Yes" / "No")
 * [4] Reminder time (Coloumn E: REMINDER TIME)
 * [5] Loop type (Coloumn F: REPEAT - "Daily", "Weekly", "Monthly")
 * [6] Loop occurrence (Coloumn G: TIME REPEAT)
 * [7] Event title (Coloumn H: TITLE)
 * [8] Event description (Coloumn I: DESCRIPTION)
 * [9] Status (Coloumn J: STATUS)
 * 
 * Note : This script won't work if there's no data in column A and there's data in column J
 */
function createCalendarEvents() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
    
  // Skip fist coloumn for header
  // Get data from column A (index 1) to J (index 10), 10 columns in total.
  var dataRange = sheet.getRange(2, 1, lastRow - 1, 10); 
  var data = dataRange.getValues();

  var calendar = CalendarApp.getDefaultCalendar(); // Use the user's default calendar

  // Loop through each row of data
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowIndex = i + 2; // Actual row in the spreadsheet (since we start from row 2)

    // Read data from each column based on array index (0-indexed)
    var startDateStr = row[0];		// Column A: DATE
    var startTimeStr = row[1];		// Column B: START
    var endTimeStr = row[2];		// Column C: END
    var reminderOption = row[3];	// Column D: REMINDER
    var reminderMinutes = row[4];	// Column E: REMINDER TIME
    var recurrenceType = row[5];	// Column F: REPEAT
    var recurrenceTime = row[6];	// Column G: TIME REPEAT
    var title = row[7];				// Column H: TITLE
    var description = row[8];		// Column I: DESCRIPTION
    var status = row[9];			// Column J: STATUS

    // Stop processing if Start Date column (A) is empty
    if (!startDateStr) {
      Logger.log('Column A (Start Date) is empty in row ' + rowIndex + '. Stopping processing.');
      break;
    }

    // Skip row if status (Column J) is already filled (already processed)
    if (status && status !== '') {
      continue;
    }

    try {
      // Combine date and time to create Date objects
      // Using ISO 8601 format with GMT+7 (+07:00) time offset
      // Custom to your timezone by editing +07:00 below
      var startDateTime = new Date(startDateStr + 'T' + startTimeStr + ':00+07:00');
      var endDateTime = new Date(startDateStr + 'T' + endTimeStr + ':00+07:00');
      
      // Validate date and time
      if (isNaN(startDateTime.getTime()) || isNaN(endDateTime.getTime())) {
        sheet.getRange(rowIndex, 10).setValue('Failed: Format Date/Time format is invalid');
        continue;
      }
      
      // Automate to adding date +1 when the event end in the next day
      if (endDateTime <= startDateTime) {
        endDateTime.setDate(endDateTime.getDate() + 1);
      }

      var event;
      var options = {};

      // Add description to event options
      if (description) {
        options.description = description;
      }

      // Determine recurrence
      if (recurrenceType) {
        // Create a new recurrence object before adding rules
        var recurrence = CalendarApp.newRecurrence(); 
        switch (recurrenceType.toLowerCase()) {
          case 'daily':
            recurrence = recurrence.addDailyRule();
            break;
          case 'weekly':
            recurrence = recurrence.addWeeklyRule();
            break;
          case 'monthly':
            recurrence = recurrence.addMonthlyRule();
            break;
          default:
            recurrence = null; // No recurrence if type is unknown/blank
        }

        if (recurrence) {
          // Set number of loop
          var timesToRepeat = parseInt(recurrenceTime);
          if (!isNaN(timesToRepeat) && timesToRepeat > 0) {
            Logger.log('Type of recurrence before times(): ' + typeof recurrence);
            Logger.log('Is recurrence an object? ' + (typeof recurrence === 'object' && recurrence !== null));
            Logger.log('Does recurrence have times method? ' + (typeof recurrence.times === 'function'));
            recurrence.times(timesToRepeat); 
          } else {
            // If the value is <= 0 then the event will repeat indefinitely
            Logger.log('Invalid number of repetitions in row ' + rowIndex + '. Event will repeat indefinitely.');
          }

          // Adding additional title for recurrence event
          var eventTitle = title || recurrenceType;
          event = calendar.createEventSeries(eventTitle, startDateTime, endDateTime, recurrence, options);
        } else {
          // Single event if no recurrence or invalid type
          var eventTitle = title || recurrenceType;
          event = calendar.createEvent(eventTitle, startDateTime, endDateTime, options);
        }
      } else {
        // Single event if no recurrence type
        var eventTitle = title;
        event = calendar.createEvent(eventTitle, startDateTime, endDateTime, options);
      }

      // Add reminder if "Yes" option is selected
      if (reminderOption && reminderOption.toLowerCase() === 'yes') {
        var minutes = parseInt(reminderMinutes);
        if (isNaN(minutes) || minutes <= 0) {
          minutes = 30; // Default 30 minutes if empty or invalid. Adjust this value as your preference
        }
        event.addPopupReminder(minutes); // Add pop-up reminder
      }

      // Update column J (index 10) with success status
      sheet.getRange(rowIndex, 10).setValue('Done');

    } catch (e) {
      // Handle errors and update column J (index 10) with error message
      sheet.getRange(rowIndex, 10).setValue('Failed: ' + e.message);
      Logger.log('Error in row ' + rowIndex + ': ' + e.message);
    }
  }
  SpreadsheetApp.getUi().alert('Schedule Synced');
}
