/* =========================================================================
   Categories.gs
   업무 분류 목록 조회/추가. '카테고리' 시트에 저장되며, 없으면
   DEFAULT_CATEGORIES로 시딩합니다.
   ========================================================================= */

function getCategories_(ss) {
  var sheet = ss.getSheetByName('카테고리');
  if (!sheet) {
    sheet = ss.insertSheet('카테고리');
    sheet.appendRow(['이름']);
    sheet.getRange(1, 1, 1, 1).setFontWeight('bold');
    DEFAULT_CATEGORIES.forEach(function(name) { sheet.appendRow([name]); });
  }
  var data = sheet.getDataRange().getValues();
  var categories = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) categories.push(String(data[i][0]));
  }
  return categories;
}

function addCategory_(ss, data) {
  var name = String(data.name || '').trim();
  if (!name) throw new Error('카테고리 이름을 입력하세요.');
  var existing = getCategories_(ss);
  if (existing.indexOf(name) !== -1) throw new Error('이미 존재하는 카테고리입니다.');
  ss.getSheetByName('카테고리').appendRow([name]);
}
