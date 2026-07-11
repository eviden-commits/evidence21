/* =========================================================================
   Utils.gs
   공용 헬퍼. 날짜 셀은 경로에 따라 Date 객체(SpreadsheetApp) 또는 일련번호
   (Sheets API batchGet, UNFORMATTED_VALUE)로 들어올 수 있어 한 곳에서 처리.
   ========================================================================= */

function cellToDateKey_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, "GMT+9", "yyyy-MM-dd");
  if (typeof val === 'number') {
    var ms = Math.round((val - 25569) * 86400000); // Sheets epoch(1899-12-30) → Unix epoch
    return Utilities.formatDate(new Date(ms), "GMT+9", "yyyy-MM-dd");
  }
  return val ? String(val) : '';
}
