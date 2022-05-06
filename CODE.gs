// generate list of sheets within spreadsheet
// pros:
// sheets can therefore be accessed by referring to position in sheets_list array e.g. sheets_list[0].
// readily accepts name changes.
// cons: code becomes dependent on order of sheets in spreadsheet. master must remain first spreadsheet; reqNumGenerator must remain last.
function getSheets(spreadsheet) {
  var sheets = spreadsheet.getSheets()
  var sheets_list = []
  for(i=0; i<sheets.length; i++) {
    sheets_list.push(sheets[i].getName())
  }
  return sheets_list
}

// extract data from all sheets and combine into single master spreadsheet
function combineSheets(spreadsheet, sheets_list) {

  // for each sheet:

  // get data.
  // first, the master form responses, which encompass all vehicles. 
  var range = spreadsheet.getSheetByName(sheets_list[1]).getRange("A2:G")
  var signups = range.getValues()

  // assign data to variables
  for(g=0; g<signups.length; g++) {
    var shift = signups[g]
    
    var timestamp = shift[0]
    var email = shift[1]
    var name = shift[2]
    var vehicle = shift[3]
    var checkout_day = shift[4]
    var checkout_time = shift[5]
    var return_time = shift[6]

    // convert format of checkout and return times
    var cot_datetime = new Date(checkout_time)
    var hours = cot_datetime.getHours();
    var mins = cot_datetime.getMinutes();
    var secs = cot_datetime.getSeconds();
    var cot_convert = hours + ":" + mins + ":" + secs

    var rt_datetime = new Date(return_time)
    var rthours = rt_datetime.getHours();
    var rtmins = rt_datetime.getMinutes();
    var rtsecs = rt_datetime.getSeconds();
    var rt_convert = rthours + ":" + rtmins + ":" + rtsecs

    // stop when reached empty row
    if(timestamp == "") {
      break
    }

    // collect each entry as an array, will later add each value in array to corresponding column.
    // create 1D and 2D arrays, for ease of access.
    // setValues only accepts 2D arrays.
    var entries = new Array(timestamp, email, name, vehicle, checkout_day, cot_convert, rt_convert)
    var entries2d = [entries]

    // need to compare with existing log of events to check if events have already been added.
    // this will be done by comparing timestamps, which are least likely to overlap.
    var master_range = spreadsheet.getSheetByName(sheets_list[0]).getRange("A2:G")
    var master_signups = master_range.getValues()

    var existing = new Array()

    // read existing log. collect timestamps and add to 'existing' Array.
    for(y=0; y<master_signups.length; y++) {
      var master_shift = master_signups[y]

      var master_timestamp = master_shift[0]

      if(master_timestamp == "") {
        break
      }

      existing.push(String(master_timestamp))
    }

    // compare timestamps. if timestamp already exists, 
    if(existing.includes(String(timestamp))) {
        continue
      } else {
        // find input range, aka first empty row
        for(z=0; z<master_signups.length; z++) {
          var master_shift = master_signups[z]

          var master_timestamp = master_shift[0]

          if(master_timestamp == "") {
            var range_input = spreadsheet.getSheetByName(sheets_list[0]).getRange(z+2, 1, 1, 7)
            break
          }
        }

        // add values to first empty row
        range_input.setValues(entries2d)
      }
  }
  
  // now process all other form responses, which are each exclusive to one vehicle.
  for(i=2; i<=7; i++) {
    var open_range = spreadsheet.getSheetByName(sheets_list[i]).getRange("A3:J")
    var signups = open_range.getValues()

    // step 1. create entries array
    for(x=0; x<signups.length; x++) {

      var shift = signups[x]

      var timestamp = shift[0]
      var email = shift[1]
      var name = shift[2]
      var checkout_day = shift[3]
      var checkout_time = shift[4]
      var return_time = shift[5]

      var vehicle = spreadsheet.getSheetByName(sheets_list[i]).getRange("A1").getValue()

      var cot_datetime = new Date(checkout_time)
      var hours = cot_datetime.getHours();
      var mins = cot_datetime.getMinutes();
      var secs = cot_datetime.getSeconds();
      var cot_convert = hours + ":" + mins + ":" + secs

      var rt_datetime = new Date(return_time)
      var rthours = rt_datetime.getHours();
      var rtmins = rt_datetime.getMinutes();
      var rtsecs = rt_datetime.getSeconds();
      var rt_convert = rthours + ":" + rtmins + ":" + rtsecs

      // negate blank entries
      if(timestamp == "") {
        break
      }
      
      // 1d and 2d arrays, for ease of access
      var entries = new Array(timestamp, email, name, vehicle, checkout_day, cot_convert, rt_convert)
      var entries2d = [entries]

      // step 2. compare to existing log of entries
      var master_range = spreadsheet.getSheetByName(sheets_list[0]).getRange("A2:G")
      var master_signups = master_range.getValues()

      var existing = new Array()

      // read existing log. collect timestamps.
      for(y=0; y<master_signups.length; y++) {
        var master_shift = master_signups[y]

        var master_timestamp = master_shift[0]

        if(master_timestamp == "") {
          break
        }

        existing.push(String(master_timestamp))


      }
      
      // if timestamp already exists, then move on to next entry
      if(existing.includes(String(timestamp))) {
        continue
      } else {
        // find input range
        for(z=0; z<master_signups.length; z++) {
          var master_shift = master_signups[z]

          var master_timestamp = master_shift[0]

          if(master_timestamp == "") {
            var range_input = spreadsheet.getSheetByName(sheets_list[0]).getRange(z+2, 1, 1, 7)
            break
          }
        }
        range_input.setValues(entries2d)
      }
    }
  }
}

// generate a random 10 digit number that is unique and permanently assigned to an event.
function makeReqNum(spreadsheet, sheets_list) {
  
  conflict = true

  // loops until random number is unique.
  while (conflict == true) {
    var conflictQuant = 0

    // updates cell in reqNumGenerator sheet, generating new number using Google Sheets formula.
    var cell = spreadsheet.getSheetByName(sheets_list[8]).getRange("A1")
    cell.setValue("=randbetween(1111111111,9999999999)")
    var cell = spreadsheet.getSheetByName(sheets_list[8]).getRange("A1")
    var value = String(cell.getValue())
    
    // check existing request numbers.
    var existing_range = spreadsheet.getSheetByName(sheets_list[0]).getRange("J2:J")
    var existing = existing_range.getValues()

    // if generated number already exists, reset while loop
    for(i=0; i<existing.length; i++) {
      var shift = existing[i]
      var req = shift[0]

      if(req == "") {
        break
      }
      
      if(value == req) {
        conflictQuant += 1
        break
      }

    }
    if(conflictQuant != 0) {
      conflict = true
    } else {
      if(conflictQuant == 0) {
        conflict = false
      }
    }
  }
  return value
}

// add generated req num to appropriate cell in master sheet
function addReqNum(spreadsheet, sheets_list) {
  var range = spreadsheet.getSheetByName(sheets_list[0]).getRange("A2:J")
  var signups = range.getValues()

  for(x=0; x<signups.length; x++) {
    var shift = signups[x]
    
    var timestamp = shift[0]
    var existing_num = shift[9]

    // only add if event has already been added to master, but no request number
    if(timestamp != "" && existing_num == "") {
      var reqNum = makeReqNum(spreadsheet, sheets_list)
      var input_range = spreadsheet.getSheetByName(sheets_list[0]).getRange(x+2, 10, 1, 1)
      input_range.setValue(reqNum)
    }
  }

}

// make calendar events
function makeCalendarEvents(spreadsheet, sheets_list, eventCal) {
  var range = spreadsheet.getSheetByName(sheets_list[0]).getRange("A2:K")
  var signups = range.getValues()

  for(x=0; x<signups.length; x++) {
    var shift = signups[x]

    var timestamp = shift[0]
    var email = shift[1]
    var name = shift[2]
    var vehicle = shift[3]
    var cot = shift[7]
    var rt = shift[8]
    var reqNum = shift[9]
    var done = shift[10]

    // if empty row
    if(timestamp == "") {
      break
    }

    // column K tracks if events have already been created.
    // check if value in col K says "y" aka completed
    if(done == "y") {
      continue
    }

    // input "y" in col K cell.
    var range_input = spreadsheet.getSheetByName(sheets_list[0]).getRange(x+2, 11, 1, 1)
    range_input.setValue("y")

    // if checkout time is after return time, status = invalid and exit the whole loop.
    // prevents termination by error
    var cotdate = new Date(cot)
    var cotvalue = cotdate.valueOf()
    var rtdate = new Date(rt)
    var rtvalue = rtdate.valueOf()

    if(cotvalue > rtvalue) {
      status = "invalid"
      return status
    }

    // if times are valid, check for existing events within requested times.
    var arrayEvents = eventCal.getEvents(new Date(cot), new Date(rt))
    
    // if any events exist, check if vehicle has already been reserved.
    if(arrayEvents.length > 0) {

      conflict = 0

      for(i=0; i<arrayEvents.length; i++) {
        var title = arrayEvents[i].getTitle()

        if(title.includes(vehicle)) {
          conflict += 1
        }

      }

      // if vehicle has been reserved for any of the conflicting events, then return denied.
      // otherwise, create the event and return success.
      if(conflict > 0) {
        var status = "denied"
        return status
      } else {
        eventCal.createEvent(
          name + ": " + vehicle,
          new Date(cot),
          new Date(rt),
          {description: "Request #: " + reqNum}
        )
        var status = "success"
        return status
      }
    } else {
      eventCal.createEvent(
        name + ": " + vehicle,
        new Date(cot),
        new Date(rt),
        {description: "Request #: " + reqNum}
      )
      var status = "success"
      return status
    }
  }
}

// send emails to user notifying status.
function sendEmails(spreadsheet, sheets_list, status, eventCal) {

  // formatted messages including helpful information and instructions. to be sent as footnote regardless of status.
  var cancel = '<pre style="font-family: calibri; font-size: 14px; color: red">\n=\nIf you did not request this reservation, or you would like to cancel your reservation, please contact Facilities at 213-764-4135.</pre>'
  var links = '<pre style="font-family: times-new-roman; font-size: 12px">\n=\nTo submit another request, click <a href="https://tinyurl.com/vehiclerequests">here</a>. \n\nTo access the Google Calendar, click <a href="https://tinyurl.com/usclvr">here</a>. Click on each event for more details. \n\nTo view this calendar in Outlook, follow the linked instructions to <a href="https://support.microsoft.com/en-us/office/see-your-google-calendar-in-outlook-c1dab514-0ad4-4811-824a-7d02c5e77126">subscribe to a Google Calendar on Outlook</a>. The iCal address is: \n\nhttps://calendar.google.com/calendar/ical/c_phca7kne28nksioddqnco88ibc%40group.calendar.google.com/public/basic.ics \n\nNOTE: Updates to the Google Calendar can take up to 10-20 minutes to be reflected on Outlook. Changes to the Outlook calendar will not affect the Google Calendar.\n\nPlease contact Facilities at 213-764-4135 if you have any questions or concerns. </pre>'
  
  // get data and assign to variables again, including new col L tracking whether email was sent.
  var range = spreadsheet.getSheetByName(sheets_list[0]).getRange("A2:L")
  var signups = range.getValues()

  for(x=0; x<signups.length; x++) {
    var shift = signups[x]

    var timestamp = shift[0]
    var email = shift[1]
    var name = shift[2]
    var vehicle = shift[3]
    var cot = shift[7]
    var cotdate = new Date(cot)
    var cotstring = cotdate.toLocaleString('en-US', {timeZone: "America/Los_Angeles", weekday:"long", year:"numeric", month:"short", day:"numeric", hour: 'numeric', minute: 'numeric', hour12: true})
    var rt = shift[8]
    var rtdate = new Date(rt)
    var rtstring = rtdate.toLocaleString('en-US', {timeZone: "America/Los_Angeles", weekday:"long", year:"numeric", month:"short", day:"numeric", hour: 'numeric', minute: 'numeric', hour12: true})
    var reqNum = shift[9]
    var done = shift[10]
    var emaildone = shift[11]

    // if empty row
    if(timestamp == "") {
      break
    }

    // if event has been processed but email not sent yet:    
    if(done == "y" && emaildone == "") {

      // if denied, aka conflicting events
      if(status == "denied") {
        
        // list the conflicting events for the user.

        // first create an empty list to store events.
        var conflict_list = []

        // get titles of events occurring within requested time.
        var arrayEvents = eventCal.getEvents(cotdate, rtdate)
        for(i=0; i<arrayEvents.length; i++) {
          var title = arrayEvents[i].getTitle()

          // if vehicle is reserved for any of those events,
          if(title.includes(vehicle)) {

            // get start and end times for relevant events
            var start = arrayEvents[i].getStartTime().toLocaleString('en-US', {timeZone: "America/Los_Angeles", weekday:"long", year:"numeric", month:"short", day:"numeric", hour: 'numeric', minute: 'numeric', hour12: true})
            var end = arrayEvents[i].getEndTime().toLocaleString('en-US', {timeZone: "America/Los_Angeles", weekday:"long", year:"numeric", month:"short", day:"numeric", hour: 'numeric', minute: 'numeric', hour12: true})

            // format events as strings, add strings to list
            conflict_list.push(title + " | " + start + " to " + end)
          }
        }

        Logger.log(conflict_list)

        // convert list to string
        var conflicting_dates = String(conflict_list.join("<br>"))

        // send email
        MailApp.sendEmail(
          {to: email,
          subject: "Denied: USC Libraries Vehicle Request for " + name,
          htmlBody: "Dear " + name + ",<br><br>"
          + "REQ#: " + reqNum + "<br><br>"
          + "The vehicle you requested:<br><br>"
          + vehicle + "<br><br>is not available at the times you selected."
          + "<br><br>Please note the following conflicts:<br><br>"
          + conflicting_dates + links,
          cc: "montesst@usc.edu, libsec@usc.edu, jpark802@usc.edu"}
        )
      } else {
        if(status == "success") {

          // send email
          MailApp.sendEmail(
            {to: email,
            subject: "Confirmed: USC Libraries Vehicle Request for " + name,
            htmlBody: "Dear " + name + ",<br><br>"
            + "REQ#: " + reqNum + "<br><br>"
            + "This is to confirm your reservation for:<br><br>"
            + vehicle + "<br><br>from " + cotstring + " to " + rtstring + "."
            + cancel + links,
            cc: "montesst@usc.edu, libsec@usc.edu, jpark802@usc.edu"})
          
          }

          if(status == "invalid") {

            // send email
            MailApp.sendEmail(
              {to: email,
              subject: "Invalid times: USC Libraries Vehicle Request for " + name,
              htmlBody: "Dear " + name + ",<br><br>"
              + "REQ#: " + reqNum + "<br><br>"
              + "Your reservation is invalid as you entered invalid times.<br><br>Here are the times you entered:<br><br>" 
              + "Checkout: " + cotstring
              + "<br><br>Return: " + rtstring
              + "<br><br>Please confirm that checkout time is before return time. This usually occurs when entering inaccurate AM/PM values."
              + links,
              cc: "montesst@usc.edu, libsec@usc.edu, jpark802@usc.edu"}
            )
          }        
      }
    
    // set email done as "y"
    var range_input = spreadsheet.getSheetByName(sheets_list[0]).getRange(x+2, 12, 1, 1)
    range_input.setValue("y")
    }
  }
}

// execute
function main() {
  var spreadsheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1oiC11wB7hq8I0feH7lEMXtLTCUsurKnJNy2EjBhvMQ8/edit?usp=sharing");
  var eventCal = CalendarApp.getCalendarById("c_p5nepsmb6v3d2qv1h44n9tbc28@group.calendar.google.com");
  sheets_list = getSheets(spreadsheet)
  combineSheets(spreadsheet, sheets_list)
  Logger.log("Sheets combined.")
  addReqNum(spreadsheet, sheets_list)
  Logger.log("Request numbers generated.")
  status = makeCalendarEvents(spreadsheet, sheets_list, eventCal)
  Logger.log(status)
  sendEmails(spreadsheet, sheets_list, status, eventCal)
  Logger.log("Emails sent.")

}
