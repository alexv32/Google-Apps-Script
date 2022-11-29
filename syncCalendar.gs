function syncCalendar() {
  try{
  
  var eventCal
  if(!CalendarApp.getCalendarsByName("Tier 1 Shifts")[0]){
    CalendarApp.createCalendar("Tier 1 Shifts")
    eventCal=CalendarApp.getCalendarsByName("Tier 1 Shifts")[0]
  }
  else{
  eventCal=CalendarApp.getCalendarsByName("Tier 1 Shifts")[0]
  }
  var agentShifts=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName())
  var userEmail = Session.getActiveUser().getEmail()

  var agentRow = agentShifts.createTextFinder(userEmail).findNext().getRow()
  
  var shiftsRow= agentShifts.getRange("C"+agentRow+":"+"AG"+agentRow).getValues()[0]
  
  var dayOfmonth=1
  var thisMonth = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName().split(" ")[0]
  var thisYear= new Date().getFullYear().toString()

  //Logger.log(thisMonth+' '+dayOfmonth+', '+thisYear+' '+'06:45:00 GMT+2')
  if((eventCal.getEventsForDay(new Date(thisMonth+' '+dayOfmonth+', '+thisYear))).length==0){

    Logger.log("Condition works")
  
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
    SpreadsheetApp.getActive().toast(error)
    Logger.log(error)
  }
}

function resyncCalendar(){
 
  var eventCal=CalendarApp.getCalendarsByName("Tier 1 Shifts")[0]

  var agentShifts=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName())
  var userEmail = Session.getActiveUser().getEmail()

  var agentRow = agentShifts.createTextFinder(userEmail).findNext().getRow()
  
  var shiftsRow= agentShifts.getRange("C"+agentRow+":"+"AG"+agentRow).getValues()[0]
  
  var dayOfmonth=1
  var thisMonth = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].getSheetName().split(" ")[0]
  var thisYear= new Date().getFullYear().toString()
 
  for(var i=0;i<shiftsRow.length;i=i+1){
    
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

function onOpen() {
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Sync Calendar",
    functionName : "syncCalendar"
  },{name : "Resync Changes", functionName : "resyncCalendar"}];
  activeSheet.addMenu("Custom Menu",entries);
}
