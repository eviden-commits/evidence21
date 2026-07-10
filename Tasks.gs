/* =========================================================================
   Tasks.gs
   업무데이터 시트 CRUD (조회/추가/수정/삭제/상태토글) 및 첨부파일 업로드.
   ========================================================================= */

function getTasks_(ss, categoryNames) {
  var tasks = {};
  categoryNames.forEach(function(cat) { tasks[cat] = []; });

  var taskSheet = ss.getSheetByName('업무데이터');
  if (!taskSheet) return tasks;

  var taskData = taskSheet.getDataRange().getValues();
  for (var i = 1; i < taskData.length; i++) {
    var row = taskData[i];
    var cat = row[1];
    if (tasks[cat]) {
      tasks[cat].push({
        id: row[0],
        date: row[2] ? Utilities.formatDate(new Date(row[2]), "GMT+9", "yyyy-MM-dd") : '',
        title: row[3],
        content: row[4] || '',
        duedate: row[5] ? Utilities.formatDate(new Date(row[5]), "GMT+9", "yyyy-MM-dd") : '',
        status: row[6] || 'active',
        fileUrl: row[7] || ''
      });
    }
  }
  return tasks;
}

function uploadTaskFile_(data) {
  if (!data.fileData || !data.fileName) return '';
  var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  var contentBytes = Utilities.base64Decode(data.fileData);
  var blob = Utilities.newBlob(contentBytes, data.fileMime, data.fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function addTask_(ss, data) {
  var sheet = ss.getSheetByName('업무데이터');
  var fileUrl = uploadTaskFile_(data);
  sheet.appendRow([data.id, data.category, data.date, data.title, data.content || '', data.duedate || '', 'active', fileUrl]);
}

function updateTask_(ss, data) {
  var sheet = ss.getSheetByName('업무데이터');
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(data.id)) {
      var rowNum = i + 1;
      sheet.getRange(rowNum, 2).setValue(data.category);
      sheet.getRange(rowNum, 3).setValue(data.date);
      sheet.getRange(rowNum, 4).setValue(data.title);
      sheet.getRange(rowNum, 5).setValue(data.content || '');
      sheet.getRange(rowNum, 6).setValue(data.duedate || '');
      if (data.fileData && data.fileName) {
        sheet.getRange(rowNum, 8).setValue(uploadTaskFile_(data));
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
      if (data.action === 'deleteTask') {
        sheet.deleteRow(i + 1);
      } else {
        var current = values[i][6] || 'active';
        sheet.getRange(i + 1, 7).setValue(current === 'done' ? 'active' : 'done');
      }
      break;
    }
  }
}
