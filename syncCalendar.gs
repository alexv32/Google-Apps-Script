function syncCalendar() {
  /*Try catch to catch and show a notification for any errors on run */
  try{ 
  
  // Creating a New calendar if no calendar with that name exists and getting the calendar object
  
  var eventCal
  if(!CalendarApp.getCalendarsByName("Shifts")[0]){
    CalendarApp.createCalendar("Shifts")
    eventCal=CalendarApp.getCalendarsByName("Shifts")[0]
  }
  // If the Calendar exists, just get the object of the calendar
  else{
  eventCal=CalendarApp.getCalendarsByName("Shifts")[0]
  }
    
  //Getting the first sheet in the Spreadsheet used to manage the shifts  
  var agentShifts=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName())
  
  //Getting the email of the current user looking at the shifts and getting the row of the user and all shifts assigned to that user
  var userEmail = Session.getActiveUser().getEmail()
  var agentRow = agentShifts.createTextFinder(userEmail).findNext().getRow()
  
  var shiftsRow= agentShifts.getRange("C"+agentRow+":"+"AG"+agentRow).getValues()[0]
  
  //Date data to be used later, the current year and the Month mentioned in the sheet name, for example the sheet is called "December 2022"
  var dayOfmonth=1
  var thisMonth = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName().split(" ")[0]
  var thisYear= new Date().getFullYear().toString()

  //Only add shifts is there are no shifts added at all, to refrain from double assignment on the same day
  if((eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+', '+thisYear))).length==0){
    
    /*Three for Loops for every shift we have at our office, M for Morning, Mid for Mid Shift, and E for Evening
      This way the correct events are created for each shift with the right times and Title
      Run on the array of shifts we got and increment the day of month each time to match the day of the month, Friday and Saturday are left empty*/
    for(var i=0;i<shiftsRow.length;i=i+1){
      if(shiftsRow[i]=="M"){
        eventCal.createEvent("Morning Shift",
        new Date(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'06:45:00 GMT+2'),
        new Date(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'15:15:00 GMT+2'))
        dayOfmonth=dayOfmonth+1
      }
      else if(shiftsRow[i]=="Mid"){
        eventCal.createEvent("Mid Shift",
        new Date(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'11:00:00 GMT+2'),
        new Date(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'19:30:00 GMT+2'))
        dayOfmonth=dayOfmonth+1
      }
      else if(shiftsRow[i]=="E"){
        eventCal.createEvent("Evening Shift",
        new Date(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'13:45:00 GMT+2'),
        new Date(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'22:15:00 GMT+2'))
        dayOfmonth=dayOfmonth+1
      }
      else dayOfmonth=dayOfmonth+1
    
    }
  }
  }catch(error) {
    SpreadsheetApp.getActive().toast(error) //Catch and show a toast of the error
  }
}

/*Resync the calendar for each worker based on the changes made for each shift */
function resyncCalendar(){
  
  /*Same process of data pulling as the previous function, get the shifts the email and row and every shift for that worker */
  var eventCal=CalendarApp.getCalendarsByName("Tier 1 Shifts")[0]

  var agentShifts=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName())
  var userEmail = Session.getActiveUser().getEmail()

  var agentRow = agentShifts.createTextFinder(userEmail).findNext().getRow()
  
  var shiftsRow= agentShifts.getRange("C"+agentRow+":"+"AG"+agentRow).getValues()[0]
  
  //Again set data for date handling in the function
  var dayOfmonth=1
  var thisMonth = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName().split(" ")[0]
  var thisYear= new Date().getFullYear().toString()
 
  /*For loop as made previously in the function above */
  for(var i=0;i<shiftsRow.length;i=i+1){
    
    /*If statement for each shift, check what shift is set in the sheet and compare it to the Event title set in the first Calendar Sync
      If the shift is M for example and the Title is not Morning Shift 
      so the event in the calendar is updated with the new start and end times and the correct title for that event*/
    
    if((shiftsRow[i]=="M") && (eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'06:45:00 GMT+2'))[0].getTitle() != "Morning Shift")){
     
      eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'00:00:00 GMT+2'))[0].setTime(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'06:45:00 GMT+2'),new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'15:15:00 GMT+2')).setTitle("Morning Shift")
      dayOfmonth=dayOfmonth+1

    }
    else if((shiftsRow[i]=="Mid") && (eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'06:45:00 GMT+2'))[0].getTitle() != "Mid shift")){

      eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'00:00:00 GMT+2'))[0].setTime(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'11:00:00 GMT+2'),new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'19:30:00 GMT+2')).setTitle("Mid Shift")
      dayOfmonth=dayOfmonth+1

    }
    else if((shiftsRow[i]=="E") && (eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'06:45:00 GMT+2'))[0].getTitle() != "Evening Shift")){

      eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'00:00:00 GMT+2'))[0].setTime(new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'13:45:00 GMT+2'),new Date(thisMonth+' '+dayOfmonth+','+' '+ thisYear+' '+'22:15:00 GMT+2')).setTitle("Evening Shift")
      dayOfmonth=dayOfmonth+1

    }
    else dayOfmonth=dayOfmonth+1
  
  }


}

//A fucntion to add the current functions as Custom menu option to click in the sheet
function onOpen() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Sync Calendar",
    functionName : "syncCalendar"
  },{name : "Resync Changes", functionName : "resyncCalendar"}];
  activeSheet.addMenu("Custom Menu",entries);
}
