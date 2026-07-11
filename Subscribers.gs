/* =========================================================================
   Subscribers.gs
   리마인드 이메일 수신자 목록. '구독자' 시트(이름, 이메일)에 저장되며,
   업무의 '담당자' 이름과 정확히 일치하는 경우에만 리마인드가 발송됩니다.
   ========================================================================= */

function getSubscribers_(ss) {
  var sheet = ss.getSheetByName('구독자');
  if (!sheet) {
    sheet = ss.insertSheet('구독자');
    sheet.appendRow(['이름', '이메일']);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  }
  var data = sheet.getDataRange().getValues();
  var subs = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) subs.push({ name: String(data[i][0]), email: String(data[i][1]) });
  }
  return subs;
}

function addSubscriber_(ss, data) {
  var name = String(data.name || '').trim();
  var email = String(data.email || '').trim();
  if (!name) throw new Error('이름을 입력하세요.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('올바른 이메일 주소를 입력하세요.');

  var sheet = ss.getSheetByName('구독자');
  if (!sheet) { getSubscribers_(ss); sheet = ss.getSheetByName('구독자'); }
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === name) {
      sheet.getRange(i + 1, 2).setValue(email); // 같은 이름이면 이메일 갱신
      return;
    }
  }
  sheet.appendRow([name, email]);
}

function findSubscriberEmail_(subscribers, assigneeName) {
  if (!assigneeName) return '';
  var found = subscribers.find(function(s) { return s.name === assigneeName; });
  return found ? found.email : '';
}
