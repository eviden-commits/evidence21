/* =========================================================================
   Schedule.gs
   주간일정 조회/저장 및 공휴일 조회.
   ========================================================================= */

function getScheduleData_(ss) {
  var schSheet = ss.getSheetByName('주간일정');
  var locationMap = {};
  if (schSheet) {
    var schData = schSheet.getDataRange().getValues();
    schData.forEach(function(row, idx) {
      if (idx === 0 || !row[0]) return;
      var d = (row[0] instanceof Date) ? row[0] : new Date(row[0]);
      var key = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
      locationMap[key] = row[1];
    });
  }

  var holidayMap = getHolidayMap_();

  var DAY_NAMES = ['SUN', 'MON', 'TUE', 'WED', 'THU', 'FRI', 'SAT'];
  var now = new Date();
  var schedule = [];
  for (var d = 0; d < 4; d++) {
    var target = new Date(now.getFullYear(), now.getMonth(), now.getDate() + d);
    var key = Utilities.formatDate(target, "GMT+9", "yyyy-MM-dd");
    schedule.push({
      id: DAY_NAMES[target.getDay()],
      date: Utilities.formatDate(target, "GMT+9", "dd"),
      fullDate: key,
      loc: locationMap[key] || '—',
      isHoliday: !!holidayMap[key],
      holiday: holidayMap[key] || ''
    });
  }
  return schedule;
}

function getHolidayMap_() {
  var holidayMap = {};
  try {
    var cal = CalendarApp.getCalendarById('ko.south_korea#holiday@group.v.calendar.google.com');
    if (cal) {
      var now = new Date();
      var rangeEnd = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 4);
      cal.getEvents(now, rangeEnd).forEach(function(ev) {
        var hKey = Utilities.formatDate(ev.getStartTime(), "GMT+9", "yyyy-MM-dd");
        holidayMap[hKey] = ev.getTitle();
      });
    }
  } catch (e) {
    console.log("Holiday API Error: " + e);
  }
  return holidayMap;
}

function saveSchedule_(ss, data) {
  var schSheet = ss.getSheetByName('주간일정');
  if (!schSheet) {
    schSheet = ss.insertSheet('주간일정');
    schSheet.appendRow(['날짜', '장소/내용']);
    schSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  var entries = data.entries; // [{date, loc}, ...]
  entries.forEach(function(entry) {
    var targetDate = entry.date;
    var loc = entry.loc || '';
    var schData = schSheet.getLastRow() > 1 ? schSheet.getDataRange().getValues() : [['날짜', '장소/내용']];
    var found = false;
    for (var i = 1; i < schData.length; i++) {
      var rowDate = schData[i][0] instanceof Date
        ? Utilities.formatDate(schData[i][0], "GMT+9", "yyyy-MM-dd")
        : String(schData[i][0]);
      if (rowDate === targetDate) {
        schSheet.getRange(i + 1, 2).setValue(loc);
        found = true;
        break;
      }
    }
    if (!found) {
      schSheet.appendRow([targetDate, loc]);
    }
  });
}
