# AutomateGSheet2GCalendar
Powered by Google Apps Script to integrate program

# Installation
1. Create new spreadsheet and make table or coloumn like bellow. Table is not a mandatory.

![spreadsheet](/src/Main_Sheet.png)

2. After you create table navigate into Extension bar and choose Apps Script, then it will open new tab with blank untitled project. You can customize the name of the project.
3. Navigate to Project settings and add check mark into Show 'appsscript.json' manifest file in editor. This will configure the needed permission to execute the script.

![settings](/src/Settings.png)

4. The example of appsscript.json can be found in Apps Script directory of this repository. The mandatory section is only on oauthScopes and make sure runtimeVersion is set to V8.
5. Copy the code the file Code.gs from this repository to your Code.gs in Apps 	Script.
6. Save everything, then run the code for the first time. It will ask needed permissions to your Google Account.
7. Back to spreadsheet. Sometimes the button is not automatically shown, but you can refresh the page, then make sure the 'Tools' bar is shown.

![tools](/src/Tools.png)

8. Click into Calendar Sync everytime after you add your schedule into spreadsheet. Whenever it successfully added into GoogleCalendar it will shown 'Done' in status coloumn.

# Formatting and settings
We need to formate the spreadsheet such as :
- Column A-C and E (Date, Start, End, and Reminder Time coloumn) will be formatted as 'text'
- Add drop down into coloumn D / Reminder (Yes, No). You can leave this blank.
- The value of coloumn F (Repeat) is Daily, Weekly, Monthly. You can leave this blank.
- The value of Time Repeat is > 0. If you fill with 0 it will repeat the event/schedule for indefinetely. You can leave this blank.
- Title and Description coloumn is a mandatory. You can not leave this blank.
- By default the code timezone is formatted as GMT+7 you can change it from Gode.gs and navigate into line 65. Change ':00+07:00' into your timezone. Example your timezone is UTC then set it into ':00+00:00'
```
    try {
      // Combine date and time to create Date objects
      // Using ISO 8601 format with GMT+7 (+07:00) time offset
      // Custom to your timezone by editing +07:00 below
      var startDateTime = new Date(startDateStr + 'T' + startTimeStr + ':00+07:00');
      var endDateTime = new Date(startDateStr + 'T' + endTimeStr + ':00+07:00');
      
```