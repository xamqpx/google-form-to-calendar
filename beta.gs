// this program adds vehicle request form submissions as calendar events, notifies requestors.

// features:
// accounts for daylight savings
// rejects overlaps

function makeRequest() {

  // call up spreadsheet:
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1AHBXB9W_zI-EEQRUpb7Rj9N0YGtpJ56zvFeBAnHZRyE/edit?usp=sharing");

  // call up calendar:
  var eventCal = CalendarApp.getCalendarById("c_e12uc461u76nmm8ovd6ipgkh1s@group.calendar.google.com");

  // get values:
  var range = spreadsheet.getSheetByName("master").getRange("X3:AC")
  var signups = range.getValues();  // actual values used for making events. change range as needed

  // for loop going through each form entry:
  for(x=0; x<signups.length; x++) {
    var shift = signups[x]

    // setting variables for columns:
    var checkOutTime = shift[0]  
    var returnTime = shift[1]
    var driverName = shift[2]
    var driverEmail = shift[3]
    var vehicleName = shift[4]
    var eventStatus = shift[5]
    var ccEmail = "libsec@usc.edu, montesst@usc.edu"

    Logger.log(driverName + ": " + vehicleName) // keeps track of which entry code is processing

    // converts Excel date value to JavaScript:
    var jsCheckOut = new Date((checkOutTime - (25567 + 1.5+1/6))*86400*1000)
    var jsReturn = new Date((returnTime - (25567 + 1.5+1/6))*86400*1000)

    // converts to daylight savings:
    var jsCheckOutDST = new Date(jsCheckOut.getTime() - 3600000)
    var jsCheckOutDSTString = Utilities.formatDate(jsCheckOutDST, spreadsheet.getSpreadsheetTimeZone(), "MMMM dd, yyyy hh:mm a")
    var jsReturnDST = new Date(jsReturn.getTime() - 3600000)
    var jsReturnDSTString = Utilities.formatDate(jsReturnDST, spreadsheet.getSpreadsheetTimeZone(), "MMMM dd, yyyy hh:mm a")

    // converts to PST from EST when displaying to user:
    var jsCheckOutPSTString = Utilities.formatDate(jsCheckOut, spreadsheet.getSpreadsheetTimeZone(), "MMMM dd, yyyy hh:mm a")
    var jsReturnPSTString = Utilities.formatDate(jsReturn, spreadsheet.getSpreadsheetTimeZone(), "MMMM dd, yyyy hh:mm a")

    // pre-formatting email with variables
    var linkChange = '<pre style="font-family: times-new-roman; font-size: 16px">\n=\nThe link to the requests calendar has changed: tinyurl.com/usclvr . We apologize for the inconvenience.</pre>'
    var nonDstDisclaimer = '<pre style="font-family: times-new-roman; font-size: 12px;"> \n\n(Times listed above do not reflect Daylight Savings. If the date you selected is within Daylight Savings Time, convert your times accordingly and confirm that your times do not overlap with any other requests via the calendar.You do not need to make a new request.)</pre>'
    var theDstDisclaimer = '<pre style="font-family: times-new-roman; font-size: 12px;"> \n\n(Times are adjusted according to Daylight Savings. If the date you selected is not within DST, convert your times accordingly and confirm that your times do not overlap with any other requests via the calendar. You do not need to make a new request.)</pre>'
    var cancel = '<pre style="font-family: calibri; font-size: 14px; color: red">\n=\nIf you did not request this reservation, or you would like to cancel your reservation, please contact Facilities at 213-764-4135.</pre>'
    var links = '<pre style="font-family: times-new-roman; font-size: 12px">\n=\nTo submit another request, click <a href="https://tinyurl.com/vehiclerequests">here</a>. \n\nTo access the Google Calendar, click <a href="https://tinyurl.com/usclvr">here</a>. Click on each event for more details. \n\nTo view this calendar in Outlook, follow the linked instructions to <a href="https://support.microsoft.com/en-us/office/see-your-google-calendar-in-outlook-c1dab514-0ad4-4811-824a-7d02c5e77126">subscribe to a Google Calendar on Outlook</a>. The iCal address is: \n\nhttps://calendar.google.com/calendar/ical/c_phca7kne28nksioddqnco88ibc%40group.calendar.google.com/public/basic.ics \n\nNOTE: Updates to the Google Calendar can take up to 10-20 minutes to be reflected on Outlook. Changes to the Outlook calendar will not affect the Google Calendar.\n\nPlease contact Facilities at 213-764-4135 if you have any questions or concerns. </pre>'

    // prewritten email messages
    var confirmedDST = {
      to: driverEmail,  
      subject: "Confirmed: USC Libraries Vehicle Request for " + driverName, 
      htmlBody: "Dear " + driverName
        + ", <br><br>Your request to reserve: <br><br>" 
        + vehicleName 
        + "<br><br>from " 
        + jsCheckOutDSTString 
        + " to " 
        + jsReturnDSTString 
        + "<br><br>is approved." 
        + linkChange + cancel + links + theDstDisclaimer,
      cc: ccEmail}
    
    var deniedDST = {
      to: driverEmail, 
      subject: "Denied: USC Libraries Vehicle Request for " + driverName,
      htmlBody: "Dear " + driverName
        + ", <br><br>The vehicle you requested:<br><br>"
        + vehicleName
        + "<br><br>is not available at the times you selected: "
        + jsCheckOutDSTString 
        + " to " 
        + jsReturnDSTString + "."
        + linkChange + cancel + links + theDstDisclaimer,
      cc: ccEmail}

    var confirmedNON = {
      to: driverEmail,
      subject: "Confirmed: USC Libraries Vehicle Request for " + driverName, 
      htmlBody: "Dear " + driverName
        + ", <br><br>Your request to reserve: <br><br>" 
        + vehicleName 
        + "<br><br>from " 
        + jsCheckOutPSTString 
        + " to " 
        + jsReturnPSTString 
        + "<br><br>is approved." 
        + linkChange + cancel + links + nonDstDisclaimer,
      cc: ccEmail}

    var deniedNON = {
      to: driverEmail, 
      subject: "Denied: USC Libraries Vehicle Request for " + driverName,
      htmlBody: "Dear " + driverName
        + ", <br><br>The vehicle you requested: <br><br>"
        + vehicleName
        + "<br><br>is not available at the times you selected: "
        + jsCheckOutPSTString 
        + " to " 
        + jsReturnPSTString + "."
        + linkChange + cancel + links + nonDstDisclaimer,
      cc: ccEmail}

    var invalidDST = {
      to: driverEmail,
      subject: "Invalid times: USC Libraries Vehicle Request for " + driverName, 
      htmlBody: "Dear " + driverName
        + ", <br><br>Your request to reserve: <br><br>" 
        + vehicleName 
        + "<br><br>from " 
        + jsCheckOutDSTString 
        + " to " 
        + jsReturnDSTString 
        + "<br><br>cannot be approved due to invalid times. Check that check-out time is before return time."
        + "<br><br>This error often occurs with incorrect AM/PM entries. " 
        + cancel + links + nonDstDisclaimer,
      cc: ccEmail}

    var invalidNON = {
      to: driverEmail,
      subject: "Invalid times: USC Libraries Vehicle Request for " + driverName, 
      htmlBody: "Dear " + driverName
        + ", <br><br>Your request to reserve: <br><br>" 
        + vehicleName 
        + "<br><br>from " 
        + jsCheckOutPSTString 
        + " to " 
        + jsReturnPSTString 
        + "<br><br>cannot be approved due to invalid times. Check that check-out time is before return time."
        + "<br><br>This error often occurs with incorrect AM/PM entries. " 
        + cancel + links + nonDstDisclaimer,
      cc: ccEmail}

    // creates variable tracking days since New Year, to track Daylight Savings:
    // Daylight Savings usually occurs between the 73rd and 310th days of the year.
    // (This is not always the case, but nonetheless averages out.)
    var yearStart = new Date(jsCheckOut.getFullYear(), 0, 0);
    var diff = jsCheckOut - yearStart;
    var oneDay = 1000 * 60 * 60 * 24;
    var daysSinceNewYear = Math.floor(diff / oneDay);
    console.log('Day of year: ' + daysSinceNewYear);

    // tracks any conflicting events within the requested time period:
    var arrayEvents = eventCal.getEvents(jsCheckOut, jsReturn)
    var arrayEventsDST = eventCal.getEvents(jsCheckOutDST, jsReturnDST)

    // if invalid
    if(eventStatus == "") { // check if not done already
      if(jsCheckOut > jsReturn) {
      Logger.log("INVALID TIME: ")
      if(daysSinceNewYear >= 73 && daysSinceNewYear <= 310) {
        Logger.log("jsCheckOut = " + jsCheckOutDSTString)
        Logger.log("jsReturn = " + jsReturnDSTString)
        MailApp.sendEmail(invalidDST)
      } else {
        Logger.log("jsCheckOut = " + jsCheckOutPSTString)
        Logger.log("jsReturn = " + jsReturnPSTString)
        MailApp.sendEmail(invalidNON)
      }
      break
    }
    }

    // IF NOT COMPLETED ALREADY:
    if(eventStatus == "") {
      // IF DAYLIGHT SAVINGS
      if(daysSinceNewYear >= 73 && daysSinceNewYear <= 310) {
      Logger.log("Daylight Savings! Adjusting...")
        if(arrayEventsDST.length == 0) {
       
          // MAKE EVENT
            eventCal.createEvent(driverName + ": " + vehicleName, jsCheckOutDST, jsReturnDST);
            Logger.log("1 new request. ")

          // SEND EMAIL TO REQUESTOR    
          MailApp.sendEmail(confirmedDST)
          
        } else {
        var conflict = 0
        // CHECK EVENTS TO SEE IF REQUESTING SAME VEHICLE
        for(i=0; i<arrayEventsDST.length; i++) {
          var vehicleMatchDST = arrayEventsDST[i].getTitle()
        // IF REQUESTING RESERVED VEHICLE. BLOCK
          if(vehicleMatchDST.includes(vehicleName)) {
            conflict = 1
            MailApp.sendEmail(deniedDST)

            Logger.log("BLOCKED: " + vehicleMatchDST)
            break
          } else {
            if(!vehicleMatchDST.includes(vehicleName)) {
              Logger.log(vehicleMatchDST + ": vehicleName not taken for this event.")}
          }
        }
        if(conflict == 0) {
          // MAKE EVENT
          eventCal.createEvent(driverName + ": " + vehicleName, jsCheckOutDST, jsReturnDST);
            Logger.log("1 new request. ")

          // SEND EMAIL TO REQUESTOR    
          MailApp.sendEmail(confirmedDST)
        }
      }} else {
      // IF NOT DAYLIGHT SAVINGS
      if(daysSinceNewYear <= 73 || daysSinceNewYear >= 310) {
        Logger.log("Not Daylight Savings.")
        if(arrayEvents.length == 0) {
          if(returnTime < checkOutTime) { 
            // MAKE EVENT
              eventCal.createEvent(driverName + ": " + vehicleName, jsCheckOut, jsReturn);
              Logger.log("1 new request. ")

            // SEND EMAIL TO REQUESTOR    
            MailApp.sendEmail(confirmedNON)
      } else {
        var conflict = 0
        // CHECK EVENTS TO SEE IF REQUESTING SAME VEHICLE
        for(i=0; i<arrayEvents.length; i++) {
          var vehicleMatch = arrayEvents[i].getTitle()
        // IF REQUESTING RESERVED VEHICLE. BLOCK
          if(vehicleMatch.includes(vehicleName)) {
            conflict = 1
            MailApp.sendEmail(deniedNON)

            Logger.log("BLOCKED by " + vehicleMatch)
            break
          } else {
            if(!vehicleMatch.includes(vehicleName)) {
              Logger.log(vehicleMatch + ": checked, not taken for this event.")}
          }
        }
      if(conflict == 0) {
        // MAKE EVENT
        eventCal.createEvent(driverName + ": " + vehicleName, jsCheckOut, jsReturn);
          Logger.log("1 new request. ")

        // SEND EMAIL TO REQUESTOR    
        MailApp.sendEmail(confirmedNON)
      }
    }
      }
    } else {
      if(eventStatus == "done") {
        Logger.log("No new requests.")
        break
        }
      }
      
  }
}   }}
