/* =========================================================================
   Code.gs
   진입점(doGet/doPost) 라우팅 전용. 실제 로직은 다른 파일에 있습니다.
   ========================================================================= */

function doGet(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var categories = getCategories_(ss);
    var payload = {
      categories: categories,
      allCategories: getAllCategories_(ss),
      tasks: getTasks_(ss, categories),
      schedule: getScheduleData_(ss),
      personnel: getTodayPersonnel_(ss),
      history: getRecentHistory_(ss, 50),
      lastSync: Utilities.formatDate(new Date(), "GMT+9", "yy-MM-dd HH:mm")
    };
    return jsonOutput_(payload);
  } catch (error) {
    return jsonOutput_({ error: error.toString() });
  }
}

function doPost(e) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var data = JSON.parse(e.postData.contents);

    if (!checkPassword_(data)) {
      return jsonOutput_({ status: "auth_error", message: "인증 코드가 올바르지 않습니다." });
    }

    switch (data.action) {
      case 'savePersonnel':
        savePersonnel_(ss, data);
        break;
      case 'addTask':
        addTask_(ss, data);
        break;
      case 'updateTask':
        updateTask_(ss, data);
        break;
      case 'saveSchedule':
        saveSchedule_(ss, data);
        break;
      case 'deleteTask':
      case 'toggleStatus':
        deleteOrToggleTask_(ss, data);
        break;
      case 'addCategory':
        addCategory_(ss, data);
        break;
      case 'toggleCategoryActive':
        toggleCategoryActive_(ss, data);
        break;
    }

    return jsonOutput_({ status: "ok" });
  } catch (error) {
    return jsonOutput_({ status: "error", message: error.toString() });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
