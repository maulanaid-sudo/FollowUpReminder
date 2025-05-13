
function FollowUpReminder() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var calendarId = sheet.getRange("I3").getValue();
  var calendar = CalendarApp.getCalendarById(calendarId);
  if (!calendar) {
    Logger.log("Calender Not Found");
    return;
  }


  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Normalisasi waktu

  for (var i = 1; i < data.length; i++) {
    var companyName = data[i][1];
    var firstFu = data[i][4];
    var secondFu = data[i][5];

    // ------- Follow Up 1 -------
    if (firstFu instanceof Date) {
      var followUp = "Follow Up 1: " + companyName;
      //var lateFU = "Terlewat Follow - " + companyName;

      if (firstFu < today) {
        delEvent(calendar, followUp, firstFu);
        if (!checkEvent(calendar, followUp, firstFu)) {
          var event = calendar.createAllDayEvent(followUp, firstFu);
          event.setColor(CalendarApp.EventColor.GRAY);
        }
      } else {
        if (!checkEvent(calendar, followUp, firstFu)) {
          calendar.createAllDayEvent(followUp, firstFu);
        }
      }
    }

    // ------- Follow Up 2 -------
    if (secondFu instanceof Date) {
      var followUp = "Follow Up 2: " + companyName;
      //var lateFU = "Terlewat: Follow Up 2 - " + companyName;

      if (secondFu < today) {
        delEvent(calendar, followUp, secondFu);
        if (!checkEvent(calendar, followUp, secondFu)) {
          var event = calendar.createAllDayEvent(followUp, secondFu);
          event.setColor(CalendarApp.EventColor.GRAY);
        }
      } else {
        if (!checkEvent(calendar, followUp, secondFu)) {
          calendar.createAllDayEvent(followUp, secondFu);
        }
      }
    }
  }
}

// Mengecek apakah event dengan nama & tanggal tertentu sudah ada
function checkEvent(calendar, namaEvent, tanggal) {
  var events = calendar.getEventsForDay(tanggal);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === namaEvent) {
      return true;
    }
  }
  return false;
}

// Menghapus event yang sesuai nama & tanggal
function delEvent(calendar, namaEvent, tanggal) {  
  var events = calendar.getEventsForDay(tanggal);
  for (var i = 0; i < events.length; i++) {
    if (events[i].getTitle() === namaEvent) {
      events[i].deleteEvent();
    }
  }
}
