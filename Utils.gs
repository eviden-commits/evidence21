/* =========================================================================
   Utils.gs
   공용 헬퍼. 날짜/시각 셀은 경로에 따라 Date 객체(SpreadsheetApp) 또는
   일련번호(Sheets API batchGet, UNFORMATTED_VALUE)로 들어올 수 있어 한
   곳에서 처리. 일련번호는 스프레드시트 타임존(Asia/Seoul, GMT+9)의
   "벽시계" 값이므로 UTC로 잘못 해석되지 않도록 9시간을 보정합니다.
   ========================================================================= */

function serialToDate_(serial) {
  var ms = Math.round((serial - 25569) * 86400000) - (9 * 3600000); // Sheets epoch → Unix epoch, GMT+9 보정
  return new Date(ms);
}

function cellToDateKey_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd");
  if (typeof val === 'number') return Utilities.formatDate(serialToDate_(val), "GMT+9", "yyyy-MM-dd");
  return val ? String(val) : '';
}

function cellToDateTimeKey_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd HH:mm");
  if (typeof val === 'number') return Utilities.formatDate(serialToDate_(val), "GMT+9", "yyyy-MM-dd HH:mm");
  return val ? String(val) : '';
}
