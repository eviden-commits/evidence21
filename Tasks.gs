/* =========================================================================
   Tasks.gs
   업무데이터 시트 CRUD (조회/추가/수정/삭제/상태토글), 첨부파일 업로드,
   변경 이력 기록.

   컬럼: A id, B 분류, C 등록일, D 제목, E 내용, F 마감일, G 상태,
        H 파일(여러 개면 줄바꿈으로 구분), I 위험도, J 태그, K 담당자,
        L 완료일시, M 리마인드발송(내부용)
   ========================================================================= */

function getTasks_(ss, categoryNames, preloadedRows) {
  var tasks = {};
  categoryNames.forEach(function(cat) { tasks[cat] = []; });

  var taskSheet = ss.getSheetByName('업무데이터');
  if (!taskSheet) return tasks;

  var taskData = preloadedRows || taskSheet.getDataRange().getValues();
  for (var i = 1; i < taskData.length; i++) {
    var row = taskData[i];
    var cat = row[1];
    if (tasks[cat]) {
      tasks[cat].push({
        id: row[0],
        date: cellToDateKey_(row[2]),
        title: row[3],
        content: row[4] || '',
        duedate: cellToDateKey_(row[5]),
        status: row[6] || 'active',
        fileUrls: splitFileUrls_(row[7]),
        riskLevel: row[8] || '',
        tags: row[9] || '',
        assignee: row[10] || '',
        completedAt: cellToDateTimeKey_(row[11])
      });
    }
  }
  return tasks;
}

function splitFileUrls_(cell) {
  return cell ? String(cell).split('\n').filter(Boolean) : [];
}

function uploadTaskFiles_(files) {
  if (!files || !files.length) return [];
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  return files.map(function(f) {
    var contentBytes = Utilities.base64Decode(f.fileData);
    var blob = Utilities.newBlob(contentBytes, f.fileMime, f.fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  });
}

function addTask_(ss, data) {
  var sheet = ss.getSheetByName('업무데이터');
  var urls = uploadTaskFiles_(data.files);
  sheet.appendRow([data.id, data.category, data.date, data.title, data.content || '', data.duedate || '', 'active', urls.join('\n'), data.riskLevel || '', data.tags || '', data.assignee || '', '', '']);
  logChange_(ss, '신규 등록', data.category, data.id, data.title, '');
  notifyAssignment_(ss, { title: data.title, category: data.category, duedate: data.duedate, content: data.content, assignee: data.assignee });
}

function updateTask_(ss, data) {
  var sheet = ss.getSheetByName('업무데이터');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      var rowNum = i + 1;
      var before = { category: values[i][1], title: values[i][3], content: values[i][4], duedate: cellToDateKey_(values[i][5]), assignee: values[i][10] };
      sheet.getRange(rowNum, 2).setValue(data.category);
      sheet.getRange(rowNum, 3).setValue(data.date);
      sheet.getRange(rowNum, 4).setValue(data.title);
      sheet.getRange(rowNum, 5).setValue(data.content || '');
      sheet.getRange(rowNum, 6).setValue(data.duedate || '');
      sheet.getRange(rowNum, 9).setValue(data.riskLevel || '');
      sheet.getRange(rowNum, 10).setValue(data.tags || '');
      sheet.getRange(rowNum, 11).setValue(data.assignee || '');
      var newUrls = (data.existingFileUrls || []).concat(uploadTaskFiles_(data.files));
      sheet.getRange(rowNum, 8).setValue(newUrls.join('\n'));
      // 마감일이 바뀌면 이전에 보낸 리마인드 체크포인트를 초기화해서 새 마감일 기준으로 다시 보내도록 함
      if (before.duedate !== String(data.duedate || '')) sheet.getRange(rowNum, 13).setValue('');
      var after = { category: data.category, title: data.title, content: data.content, duedate: data.duedate, assignee: data.assignee };
      logChange_(ss, '수정', data.category, data.id, data.title, buildTaskDiff_(before, after));
      if (data.assignee && String(data.assignee) !== String(before.assignee || '')) {
        notifyAssignment_(ss, { title: data.title, category: data.category, duedate: data.duedate, content: data.content, assignee: data.assignee });
      }
      break;
    }
  }
}

function deleteOrToggleTask_(ss, data) {
  var sheet = ss.getSheetByName('업무데이터');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      var cat = values[i][1], title = values[i][3];
      if (data.action === 'deleteTask') {
        sheet.deleteRow(i + 1);
        logChange_(ss, '삭제', cat, data.id, title, '');
      } else {
        var current = values[i][6] || 'active';
        var next = current === 'done' ? 'active' : 'done';
        sheet.getRange(i + 1, 7).setValue(next);
        if (next === 'done') {
          sheet.getRange(i + 1, 12).setValue(Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm"));
        } else {
          sheet.getRange(i + 1, 12).setValue('');
          sheet.getRange(i + 1, 13).setValue('');
        }
        logChange_(ss, next === 'done' ? '완료 처리' : '진행중으로 변경', cat, data.id, title, '');
      }
      break;
    }
  }
}
