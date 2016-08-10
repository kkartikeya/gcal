//push new events to calendar
function pushToCalendar() {
   
  //spreadsheet variables
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow(); 
  var range = sheet.getRange(3,1,lastRow,25);
  var values = range.getValues();
   
  //calendar variables
  var calendarObj = CalendarApp.getCalendarById('CALENDAR_ID')
  var calendarTZ = calendarObj.getTimeZone()
      
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
        
        // Set the timezone of the calendar to that of the event.
        var cal=calendarObj.setTimeZone(timeZone)

        // get the timezones of the calendar and the session
        var calOffset = Utilities.formatDate(new Date(), timeZone, "Z");
        var textdate=year+'-'+month+'-'+day+'T'+timesplit[0]+':'+timesplit[1]+':'+timesplit[2]+calOffset
        var startdate = new Date(getDateFromIso(textdate))
        
        var endtime = new Date(startdate.getTime() + values[i][7] * 3600000)
        
        var description = "Vendor: " + values[i][11] + "\n\nContact: " + values[i][12] + "\n\nCDT: " + values[i][13] + "\n\nScope: " + values[i][14] + "\n\nSOW: "+ values[i][15] + "\n\nGoogle Drive: "+ values[i][16] + "\n\nEquipment + IP Addresses: " + values[i][17] + "\n\nNotes: " + values[i][18] + "\n\nAdditional Information: " + values[i][19] + "\n\nFloorplan: " + values[i][20]

        var options;
                
        //Find out all the options for the event 
        if ( values[i][10] != '' ) {
          options = {description: description, guests: values[i][10], sendInvites: 'True' }
        }else {
          options = {description: description}
        }
        
        // Create the event
        var newEvent = cal.createEvent(newEventTitle, startdate, endtime, options)
        
        //get ID
        var newEventId = newEvent.getId();
         
        //mark as entered, enter ID
        sheet.getRange(i+3,1).setValue(newEventId);
         
      } 
    }
  }
  
  calendarObj.setTimeZone(calendarTZ)
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
 
function getDateFromIso(string) {
  try{
    var aDate = new Date();
    var regexp = "([0-9]{4})(-([0-9]{2})(-([0-9]{2})" +
        "(T([0-9]{2}):([0-9]{2})(:([0-9]{2}))?" +
        "(Z|(([-+])([0-9]{2})([0-9]{2}))?)?)?)?)?";
    var d = string.match(new RegExp(regexp));

    var offset = 0;
    var date = new Date(d[1], 0, 1);

    if (d[3]) { date.setMonth(d[3] - 1); }
    if (d[5]) { date.setDate(d[5]); }
    if (d[7]) { date.setHours(d[7]); }
    if (d[8]) { date.setMinutes(d[8]); }
    if (d[10]) { date.setSeconds(d[10]); }
    if (d[11]) {
      offset = (Number(d[14]) * 60) + Number(d[15]);
      offset *= ((d[13] == '-') ? 1 : -1);
    }

    offset -= date.getTimezoneOffset();
    time = (Number(date) + (offset * 60 * 1000));
    return aDate.setTime(Number(time));
  } catch(e){
    return;
  }
}
