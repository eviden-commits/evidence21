/* =========================================================================
   Personnel.gs
   출력인원 시트 조회/저장.
   ========================================================================= */

function getTodayPersonnel_(ss) {
  var todayKey = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
  var personnel = { hitech: 0, arch: 0 };
  var persSheet = ss.getSheetByName('출력인원');
  if (!persSheet || persSheet.getLastRow() <= 1) return personnel;

  var persData = persSheet.getDataRange().getValues();
  for (var i = 1; i < persData.length; i++) {
    var rowDate = persData[i][0] instanceof Date
      ? Utilities.formatDate(persData[i][0], "GMT+9", "yyyy-MM-dd")
      : String(persData[i][0]);
    if (rowDate === todayKey) {
      personnel = { hitech: Number(persData[i][1]) || 0, arch: Number(persData[i][2]) || 0 };
      break;
    }
  }
  return personnel;
}

function savePersonnel_(ss, data) {
  var persSheet = ss.getSheetByName('출력인원');
  if (!persSheet) {
    persSheet = ss.insertSheet('출력인원');
    persSheet.appendRow(['날짜', '하이테크', '건축기계']);
    persSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  var targetDate = data.date;
  var hitech = Number(data.hitech) || 0;
  var arch = Number(data.arch) || 0;
  var persData = persSheet.getLastRow() > 1 ? persSheet.getDataRange().getValues() : [['날짜', '하이테크', '건축기계']];
  var found = false;
  for (var i = 1; i < persData.length; i++) {
    var rowDate = persData[i][0] instanceof Date
      ? Utilities.formatDate(persData[i][0], "GMT+9", "yyyy-MM-dd")
      : String(persData[i][0]);
    if (rowDate === targetDate) {
      persSheet.getRange(i + 1, 2).setValue(hitech);
      persSheet.getRange(i + 1, 3).setValue(arch);
      found = true;
      break;
    }
  }
  if (!found) {
    persSheet.appendRow([targetDate, hitech, arch]);
  }
}
