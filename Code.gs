/* =========================================================================
   Code.gs
   진입점(doGet/doPost) 라우팅 전용. 실제 로직은 다른 파일에 있습니다.
   ========================================================================= */

var DASHBOARD_CACHE_KEY = 'dashboard_payload';
var DASHBOARD_CACHE_SECONDS = 60;

function doGet(e) {
  try {
    var pwd = (e && e.parameter && e.parameter.pwd) || '';
    if (!checkPassword_({ password: pwd })) {
      return jsonOutput_({ authError: true, error: "인증 코드가 올바르지 않습니다." });
    }

    var cache = CacheService.getScriptCache();
    var cached = cache.get(DASHBOARD_CACHE_KEY);
    if (cached) return jsonOutput_(JSON.parse(cached));

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var batched = fetchBatchedSheetValues_(ss);

    var allCategories = getAllCategories_(ss, batched['카테고리']);
    var categories = allCategories.filter(function(c) { return c.active; }).map(function(c) { return c.name; });

    var payload = {
      categories: categories,
      allCategories: allCategories,
      tasks: getTasks_(ss, categories, batched['업무데이터']),
      schedule: getScheduleData_(ss, batched['주간일정']),
      personnel: getTodayPersonnel_(ss, batched['출력인원']),
      history: getRecentHistory_(ss, 50, batched['변경이력']),
      lastSync: Utilities.formatDate(new Date(), "GMT+9", "yy-MM-dd HH:mm")
    };
    cache.put(DASHBOARD_CACHE_KEY, JSON.stringify(payload), DASHBOARD_CACHE_SECONDS);
    return jsonOutput_(payload);
  } catch (error) {
    return jsonOutput_({ error: error.toString() });
  }
}

// 시트가 이미 존재하는 경우에 한해 Sheets API로 여러 시트를 한 번에 읽어옵니다.
// 실패하거나 아직 시트가 없으면 빈 결과를 반환하고, 각 조회 함수가 자체적으로
// 개별 읽기(및 필요 시 자동 생성)로 폴백합니다.
function fetchBatchedSheetValues_(ss) {
  var wanted = ['업무데이터', '카테고리', '출력인원', '주간일정', '변경이력'];
  var result = {};
  try {
    var existingNames = ss.getSheets().map(function(s) { return s.getName(); });
    var present = wanted.filter(function(n) { return existingNames.indexOf(n) !== -1; });
    if (present.length === 0) return result;

    var ranges = present.map(function(n) { return "'" + n.replace(/'/g, "''") + "'"; });
    var resp = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET_ID, {
      ranges: ranges,
      valueRenderOption: 'UNFORMATTED_VALUE'
    });
    present.forEach(function(n, idx) {
      result[n] = (resp.valueRanges[idx] && resp.valueRanges[idx].values) || [];
    });
  } catch (e) {
    console.log('batchGet fallback: ' + e);
  }
  return result;
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
      case 'addSubscriber':
        addSubscriber_(ss, data);
        break;
      case 'setAdminPassword':
        setAdminPassword_(data);
        break;
    }

    CacheService.getScriptCache().remove(DASHBOARD_CACHE_KEY);
    return jsonOutput_({ status: "ok" });
  } catch (error) {
    return jsonOutput_({ status: "error", message: error.toString() });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
