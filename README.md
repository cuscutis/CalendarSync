# CalendarSync
One way sync from Outlook to Google Calendar

This is a combination of the Google Calendar API [.NET Quickstart](https://developers.google.com/google-apps/calendar/quickstart/dotnet) and some example code on [Stack Overflow](http://stackoverflow.com/a/92184/2277) with some code of my own to merge items from Outlook to Google Calendar.

This doesn't ask for email addresses from Outlook so it doesn't get trigger a security dialog to allow permission to accesss the information. It only reads the following from Outlook:
* Message Id
* Title
* Location
* Start time
* Duration
 
You will have to get your own API key from Google following the instructions in the quickstart.
