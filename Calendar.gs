//push new events to calendar
function pushToCalendar() {
   
  //spreadsheet variables
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow(); 
  var range = sheet.getRange(3,1,lastRow,25);
  var values = range.getValues();
  var installColor = "#0D7813"
  var serviceCallColor = "#AB8B00"
   
  //calendar variables
  var calendarObj = CalendarApp.getCalendarById('CALENDAR_ID')
      
  for (var i = 0; i < values.length; i++) {     
    //check to see if name and type are filled out - date is left off because length is "undefined"
    if ((values[i][5].length > 0) && (values[i][6].length > 0)) {
       
      //check if it's been entered before         
      if (values[i][0] == '') {                       
         
        //create event https://developers.google.com/apps-scriptredni/class_calendarapp#createEvent
        var newEventTitle = values[i][1] + ': ' + values[i][2] + ' - ' + values[i][3] + ' - ' + values[i][4] + ': ' + values[i][9];
        
        var datesplit = values[i][5].split("/")
        var month = datesplit[0] 
        var day = datesplit[1]
        var year = datesplit[2]
        
        var time = values[i][6]
        var timesplit = values[i][6].split(':')
        
        var timeZone = values[i][8]
        var timeZoneOffset = getOffset(timeZone)
        
        var starttime = new Date(year, month-1, day, timesplit[0], timesplit[1], timesplit[2])

        // get the timezones of the calendar and the session
        var calOffset = Utilities.formatDate(new Date(), timeZone, "Z");
        var scriptOffset = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "Z");

        // remove the 0s
        var re = /0/gi;
        calOffset = parseInt(calOffset.replace(re,''));
        scriptOffset = parseInt(scriptOffset.replace(re,''));

        var offsetDif =   calOffset - scriptOffset;
        starttime.setHours(starttime.getHours() - offsetDif);
        
        var endtime = new Date(starttime.getTime() + values[i][7] * 3600000)
        
        var description = "Vendor: " + values[i][11] + "\n\nContact: " + values[i][12] + "\n\nCDT: " + values[i][13] + "\n\nScope: " + values[i][14]

        var options, color;
        
        if (( values[i][9] == "Install") || (values[i][9] == "Revisit")) {
          color = installColor
        }else {
          color = serviceCallColor
        }
        
        //Find out all the options for the event 
        if ( values[i][10] != '' ) {
          options = {description: description, guests: values[i][10], sendInvites: 'True' }
        }else {
          options = {description: description}
        }
        
        // Create the event
        var newEvent = calendarObj.createEvent(newEventTitle, starttime, endtime, options)
        
        //get ID
        var newEventId = newEvent.getId();
         
        //mark as entered, enter ID
        sheet.getRange(i+3,1).setValue(newEventId);
         
      } 
    }
  }
 
}

function getOffset( timezone ) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var optionsSheet = ss.getSheetByName('Options')
  var dataRange = optionsSheet.getDataRange();
  var values = dataRange.getValues();
  
  for (var i = 0; i < values.length; i++) {
    var row = "";
    if (values[i][0] == timezone ) { 
      row = values[i][1].valueOf();
      return row;
    }    
  }  
  
  return null
}
 
//add a menu when the spreadsheet is opened
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];  
  menuEntries.push({name: "Update Calendar", functionName: "pushToCalendar"}); 
  sheet.addMenu("Installation Calendar", menuEntries);  
}
 
