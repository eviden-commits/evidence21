/* =========================================================================
   Categories.gs
   업무 분류 목록 조회/추가/비활성화. '카테고리' 시트(이름, 활성)에 저장되며,
   없으면 DEFAULT_CATEGORIES로 시딩합니다. 비활성 분류는 목록/드롭다운에서만
   숨겨지고 기존 업무 데이터는 그대로 남습니다.
   ========================================================================= */

function getAllCategories_(ss, preloadedRows) {
  var sheet = ss.getSheetByName('카테고리');
  if (!sheet) {
    sheet = ss.insertSheet('카테고리');
    sheet.appendRow(['이름', '활성']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    DEFAULT_CATEGORIES.forEach(function(name) { sheet.appendRow([name, true]); });
    preloadedRows = null; // 방금 새로 만들었으니 미리 읽어온 값은 무시
  }
  var data = preloadedRows;
  if (!data) {
    if (!sheet.getRange(1, 2).getValue()) sheet.getRange(1, 2).setValue('활성');
    data = sheet.getDataRange().getValues();
  }

  var categories = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var flag = data[i][1];
    var active = !(flag === false || flag === 'FALSE');
    categories.push({ name: String(data[i][0]), active: active });
  }
  return categories;
}

function getCategories_(ss, preloadedRows) {
  return getAllCategories_(ss, preloadedRows).filter(function(c) { return c.active; }).map(function(c) { return c.name; });
}

function addCategory_(ss, data) {
  var name = String(data.name || '').trim();
  if (!name) throw new Error('카테고리 이름을 입력하세요.');
  var existing = getAllCategories_(ss).map(function(c) { return c.name; });
  if (existing.indexOf(name) !== -1) throw new Error('이미 존재하는 카테고리입니다.');
  ss.getSheetByName('카테고리').appendRow([name, true]);
}

function toggleCategoryActive_(ss, data) {
  var name = String(data.name || '').trim();
  var sheet = ss.getSheetByName('카테고리');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === name) {
      var flag = values[i][1];
      var current = !(flag === false || flag === 'FALSE');
      sheet.getRange(i + 1, 2).setValue(!current);
      return;
    }
  }
  throw new Error('카테고리를 찾을 수 없습니다.');
}
