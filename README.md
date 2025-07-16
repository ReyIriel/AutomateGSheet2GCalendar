# AutomateGSheet2GCalendar
Powered by Google Apps Script to integrate program

# Installation
1. Create a new Google Spreadsheet and set up your columns as shown below. Table formatting is not mandatory.

![spreadsheet](/src/Main_Sheet.png)

2. Navigate to **Extensions > Apps Script**. This will open a new tab with a blank, untitled project. You can rename the project to whatever you like.
3. Go to **Project Settings** (the ⚙️ icon) and check the box for "**Show 'appsscript.json' manifest file in editor**". This is required to configure the necessary permissions for the script.

![settings](/src/Settings.png)

4. An example of the `appsscript.json` file can be found in the Apps Script directory of this repository. The mandatory section is `oauthScopes`, and ensure the `runtimeVersion` is set to V8.
5. Copy the code in `Code.gs` from this repository to your `Code.gs` in Apps 	Script.
6. Save everything, then run the code for the first time. It will ask needed permissions to your Google Account.
7. Back to spreadsheet. Sometimes the custom menu is not automatically shown, but you can refresh the page, then make sure the '**Tools**' bar is shown.

![tools](/src/Tools.png)

8. Click into **Calendar Sync** everytime after you add your schedule into spreadsheet. Whenever it successfully added into Google Calendar it will shown 'Done' in status coloumn.

# Formatting and Settings
We need to formates the spreadsheet such as :
- Column **A, B, C, and E** (Date, Start, End, and Reminder Time coloumn) must be formatted as '**Plain text**'.
- The **Date** should be filled with `yyyy-mm-dd` format. The code will not work if you set the format into anything else.
- Add drop down into **coloumn D / Reminder** (Yes, No). You can leave this blank.
- The value of **coloumn F (Repeat)** is `Daily, Weekly, Monthly`. You can leave this blank.
- The value of **Time Repeat** is > 0. If you fill with 0 it will repeat the event/schedule for indefinetely. You can leave this blank.
- **Title and Description** coloumn is a mandatory. You can not leave this blank.
- By default the code timezone is formatted as **GMT+7** you can change it from Code.gs and navigate into line 65. Change ':00+07:00' into your timezone. Example your timezone is UTC then set it into ':00+00:00'
```
    try {
      // Combine date and time to create Date objects
      // Using ISO 8601 format with GMT+7 (+07:00) time offset
      // Custom to your timezone by editing +07:00 below
      var startDateTime = new Date(startDateStr + 'T' + startTimeStr + ':00+07:00');
      var endDateTime = new Date(startDateStr + 'T' + endTimeStr + ':00+07:00');
      
```