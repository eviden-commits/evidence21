/* =========================================================================
   Reminders.gs
   1) 마감일 기준 D-7/D-3/D-day에 미완료 업무 담당자에게 리마인드 메일 발송
      (체크포인트당 1회, M열에 발송 이력을 남겨 중복 발송 방지)
   2) 최근 24시간 내 완료된 업무를 모아 하루 1회 마스터 이메일로 다이제스트 발송
   매일 정해진 시각에 실행되는 시간 트리거에서 호출됩니다.
   ========================================================================= */

function runDailyDigestAndReminders_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName('업무데이터');
  if (!sheet) return;

  var subscribers = getSubscribers_(ss);
  var today = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
  var now = new Date();
  var values = sheet.getDataRange().getValues();
  var completedItems = [];

  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var status = row[6] || 'active';
    var category = row[1], title = row[3], duedateRaw = row[5];

    if (status === 'done') {
      var completedAt = row[11];
      if (completedAt && hoursSince_(completedAt, now) <= 24) {
        completedItems.push({ category: category, title: title, assignee: row[10] || '', completedAt: cellToDateTimeKey_(completedAt) });
      }
      continue;
    }

    if (!duedateRaw) continue;
    var duedateStr = Utilities.formatDate(new Date(duedateRaw), "GMT+9", "yyyy-MM-dd");
    var daysLeft = daysBetween_(today, duedateStr);
    if (REMINDER_DAYS_BEFORE.indexOf(daysLeft) === -1) continue;

    var sentRaw = String(row[12] || '');
    var sentList = sentRaw ? sentRaw.split(',') : [];
    if (sentList.indexOf(String(daysLeft)) !== -1) continue; // 이미 이 체크포인트 발송함

    var assignee = row[10] || '';
    var email = findSubscriberEmail_(subscribers, assignee);
    if (email) {
      sendReminderEmail_(email, { category: category, title: title, content: row[4] || '', duedate: duedateStr, daysLeft: daysLeft, assignee: assignee });
    }
    sentList.push(String(daysLeft));
    sheet.getRange(i + 1, 13).setValue(sentList.join(','));
  }

  if (completedItems.length > 0) sendCompletionDigest_(completedItems);
}

function daysBetween_(fromStr, toStr) {
  var a = new Date(fromStr + 'T00:00:00');
  var b = new Date(toStr + 'T00:00:00');
  return Math.round((b - a) / 86400000);
}

function hoursSince_(cellVal, now) {
  if (!cellVal) return Infinity;
  var t = (cellVal instanceof Date) ? cellVal : new Date(String(cellVal).replace(' ', 'T') + '+09:00');
  if (isNaN(t.getTime())) return Infinity;
  return (now - t) / 3600000;
}

function sendReminderEmail_(toEmail, task) {
  var statusLabel = task.daysLeft === 0 ? '오늘 마감입니다' : task.daysLeft + '일 남았습니다';
  var subject = '[업무창고] ' + task.title + ' — 업무가 마감되지 않았습니다';
  var html =
    '<div style="font-family:sans-serif;font-size:14px;color:#1D1D1F;line-height:1.6;">' +
    '<p><b>' + escapeHtml_(task.title) + '</b> 업무가 아직 완료되지 않았습니다.</p>' +
    '<table style="border-collapse:collapse;margin:12px 0;">' +
    '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;">분류</td><td>' + escapeHtml_(task.category) + '</td></tr>' +
    '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;">담당자</td><td>' + escapeHtml_(task.assignee) + '</td></tr>' +
    '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;">마감일</td><td>' + escapeHtml_(task.duedate) + ' (' + statusLabel + ')</td></tr>' +
    (task.content ? '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;vertical-align:top;">내용</td><td>' + escapeHtml_(task.content).replace(/\n/g, '<br>') + '</td></tr>' : '') +
    '</table>' +
    '<p><a href="' + APP_URL + '" style="color:#007AFF;">업무창고에서 확인하기 →</a></p>' +
    '</div>';
  MailApp.sendEmail({ to: toEmail, subject: subject, htmlBody: html });
}

// 신규 등록 또는 담당자 변경(재배정 포함) 시 즉시 발송
function notifyAssignment_(ss, task) {
  if (!task.assignee) return;
  var email = findSubscriberEmail_(getSubscribers_(ss), task.assignee);
  if (!email) return;
  var subject = '[업무창고] 업무가 배정되었습니다 — ' + task.title;
  var html =
    '<div style="font-family:sans-serif;font-size:14px;color:#1D1D1F;line-height:1.6;">' +
    '<p><b>' + escapeHtml_(task.title) + '</b> 업무가 담당자로 배정되었습니다.</p>' +
    '<table style="border-collapse:collapse;margin:12px 0;">' +
    '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;">분류</td><td>' + escapeHtml_(task.category) + '</td></tr>' +
    (task.duedate ? '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;">마감일</td><td>' + escapeHtml_(task.duedate) + '</td></tr>' : '') +
    (task.content ? '<tr><td style="color:#6C6C70;padding:2px 12px 2px 0;vertical-align:top;">내용</td><td>' + escapeHtml_(task.content).replace(/\n/g, '<br>') + '</td></tr>' : '') +
    '</table>' +
    '<p><a href="' + APP_URL + '" style="color:#007AFF;">업무창고에서 확인하기 →</a></p>' +
    '</div>';
  MailApp.sendEmail({ to: email, subject: subject, htmlBody: html });
}

function sendCompletionDigest_(items) {
  var rows = items.map(function(t) {
    return '<tr><td style="padding:6px 12px 6px 0;color:#6C6C70;">' + escapeHtml_(t.category) + '</td>' +
      '<td style="padding:6px 12px 6px 0;font-weight:600;">' + escapeHtml_(t.title) + '</td>' +
      '<td style="padding:6px 12px 6px 0;">' + escapeHtml_(t.assignee) + '</td>' +
      '<td style="padding:6px 0;color:#6C6C70;">' + escapeHtml_(t.completedAt) + '</td></tr>';
  }).join('');
  var html =
    '<div style="font-family:sans-serif;font-size:14px;color:#1D1D1F;">' +
    '<p>최근 24시간 동안 완료 처리된 업무 ' + items.length + '건입니다.</p>' +
    '<table style="border-collapse:collapse;">' +
    '<tr style="font-size:11px;color:#AEAEB2;text-transform:uppercase;"><th style="text-align:left;padding-right:12px;">분류</th><th style="text-align:left;padding-right:12px;">제목</th><th style="text-align:left;padding-right:12px;">담당자</th><th style="text-align:left;">완료일시</th></tr>' +
    rows +
    '</table>' +
    '<p style="margin-top:16px;"><a href="' + APP_URL + '" style="color:#007AFF;">업무창고에서 확인하기 →</a></p>' +
    '</div>';
  MailApp.sendEmail({ to: MASTER_EMAIL, subject: '[업무창고] 최근 24시간 완료 업무 ' + items.length + '건', htmlBody: html });
}

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/* ── 최초 1회 수동 실행: 메일 발송 권한(Mail scope) 승인용. 실제 메일은 보내지 않음 ── */
function authorizeMailScope() {
  MailApp.getRemainingDailyQuota();
}

/* ── 최초 1회 수동 실행: 매일 오전 9시 트리거 설치 ── */
function installDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyDigestAndReminders_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('runDailyDigestAndReminders_').timeBased().everyDays(1).atHour(9).create();
}
