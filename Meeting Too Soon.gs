function getVIP_(sheetID) {
  var ss = SpreadsheetApp.openById(sheetID);
  var sheet = ss.getSheets()[0];
  var VIP = sheet.getRange("A:A").getValues().flat().filter(r => r != "");
  return VIP;
}
function getConfig_(sheetID) {
  var ss = SpreadsheetApp.openById(sheetID)
  var sheet = ss.getSheets()[0];
  var hours = sheet.getRange("B1").getValues().flat().filter(r => r != "");
  return hours[0].toString();
}

function sendReason_(sender, subject, hours) {
  Logger.log("notify sender:  " + sender);
  var message = "You attempted to schedule a meeting without adequate notice.\n\nPlease schedule during my normal work hours and with at least " + hours + " hours notice.\n\nI understand emergencies occure, if so please contact by instant message prior to scheduling and I will make every effort to assist you, or get the help you need.\n";
  var mailSubject = "Auto-Declined meeting for short notice:  " + subject;
  MailApp.sendEmail(sender, mailSubject, message);
}


function MeetingTooSoon() {

//
Update SheetID variable below needs to contain the document ID in your google drive.   
This sheet should be in your personal google drive -  an .xlsx is included as a sample

A trigger needs to be created to call this script when the calendar is updated.
//



  var sheetID = "< __ EDI THIS VALUE PER THE INSTRUCTION PROVIDED ___>";   // OBTAINED MY URL FOR SHEET CONTINAING CONFIG.   Column A contains VIP emails which are excluded from auto-decline.  B1 containes # hours that are declined.
  var VIP = getVIP_(sheetID);
  var hours = getConfig_(sheetID);
  var now = new Date();
  var hoursFromNow = new Date(now.getTime() + (hours * 60 * 60 * 1000));
  var events = CalendarApp.getDefaultCalendar().getEvents(now, hoursFromNow);
  Logger.log('Number of events being evaluated: ' + events.length);

  for (var i = 0; i < events.length; i++) {
    var created = events[i].getDateCreated();
    var NormCreated = created.valueOf();
    var NormNow = now.valueOf();
    var NormDiff = ((NormNow - NormCreated) / 1000).toFixed(0);  // convert to Seconds
    Logger.log("Created "+i+" created by " + events[i].getCreators() + " subject: '" + events[i].getTitle() + "' Created:  " + NormCreated + " Now:  " + NormNow + " Diff " + NormDiff);


    // determine if meeting created by a VIP, 'me', or has already been accepted/declined, if so stop.
    // else will reject the message.
    // VIp.IndexOf function will return -1 for NON VIP.  >= 0 are indexes of spreadsheet

    if (VIP.indexOf(events[i].getCreators().toString()) != -1 ||
      events[i].getCreators() == Session.getEffectiveUser().getEmail() ||
      events[i].getMyStatus() != "INVITED") {

        Logger.log("Event "+i+" has been determined to be from a VIP, my own event, or already accepted/declined");

        } else {
            if (NormDiff < (hours * 3600)) {
              var meetingID = events[i].getId()
              events[i].setMyStatus(CalendarApp.GuestStatus.NO);
              var sender = events[i].getCreators().toString();
              Logger.log("Declined : " + events[i].getTitle() + " from: " + sender);
              sendReason_(sender, events[i].getTitle(), hours);
            } else { 
              Logger.log("Event "+i+" meets acceptable time window.");
      }
    }

  }
}