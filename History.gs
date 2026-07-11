/* =========================================================================
   History.gs
   업무 등록/수정/삭제/상태변경 이력 기록. '변경이력' 시트에 누적되며,
   doGet은 최신 항목 일부만 반환합니다.
   ========================================================================= */

function logChange_(ss, action, category, id, title, detail) {
  var sheet = ss.getSheetByName('변경이력');
  if (!sheet) {
    sheet = ss.insertSheet('변경이력');
    sheet.appendRow(['시각', '동작', '분류', '업무ID', '제목', '상세']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  sheet.appendRow([
    Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm"),
    action, category || '', id || '', title || '', detail || ''
  ]);
}

function getRecentHistory_(ss, limit, preloadedRows) {
  var sheet = ss.getSheetByName('변경이력');
  if (!sheet) return [];
  var data = preloadedRows || (sheet.getLastRow() <= 1 ? [] : sheet.getDataRange().getValues());
  if (data.length <= 1) return [];
  var rows = data.slice(1).map(function(row) {
    return { time: row[0], action: row[1], category: row[2], id: row[3], title: row[4], detail: row[5] };
  });
  rows.reverse();
  return rows.slice(0, limit);
}

function buildTaskDiff_(before, after) {
  var labels = { category: '분류', title: '제목', content: '내용', duedate: '마감일', assignee: '담당자' };
  var parts = [];
  Object.keys(labels).forEach(function(key) {
    var b = before[key] || '';
    var a = after[key] || '';
    if (String(b) !== String(a)) parts.push(labels[key] + ': ' + (b || '(없음)') + ' → ' + (a || '(없음)'));
  });
  return parts.join(' / ');
}
