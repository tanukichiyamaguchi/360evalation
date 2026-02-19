// =============================================================
// 360° スタッフ評価システム - バックエンド (Code.gs)
// 同時書き込み問題をLockServiceで解決
// =============================================================

// スプレッドシートのIDとシート名
const SPREADSHEET_ID = '1lCK7YD8kGhXDmr4kbkYXRsGyKKu9cfNslVknRy34GzA';
const EVALUATOR_MASTER_SHEET_NAME = '評価者マスタ';
const EVALUATED_MASTER_SHEET_NAME = '被評価者マスタ';
const ITEM_MASTER_SHEET_NAME = '評価項目マスタ';
const TEMP_SAVE_SHEET_NAME = '一時保存データ';
const USER_AUTH_SHEET_NAME = 'ユーザー認証情報';
const SETTINGS_SHEET_NAME = '設定マスタ';
const PERIOD_MASTER_SHEET_NAME = '評価期間マスタ';
const COMPARISON_SHEET_NAME = '期間比較';
const PRINCIPAL_NAME = "院長 太郎";
const MAX_TOTAL_SCORE = 300;

// =============================================================
// ロック付き書き込みヘルパー
// =============================================================

/**
 * LockServiceを使用して排他制御付きでシートに書き込む
 */
function withLock(callback) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var result = callback();
    SpreadsheetApp.flush();
    return result;
  } catch (e) {
    throw e;
  } finally {
    lock.releaseLock();
  }
}

// =============================================================
// 評価期間管理
// =============================================================

function getPeriods() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var periodSheet = ss.getSheetByName(PERIOD_MASTER_SHEET_NAME);
    if (!periodSheet) return [];
    var lastRow = periodSheet.getLastRow();
    if (lastRow < 2) return [];
    var data = periodSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    return data
      .filter(function(row) {
        var id = row[0];
        return id &&
               id.toString().trim() !== '' &&
               !id.toString().startsWith('※') &&
               !id.toString().startsWith('=');
      })
      .map(function(row) {
        return {
          id: row[0],
          name: row[1],
          startDate: row[2] ? formatDateForWeb(row[2]) : '',
          endDate: row[3] ? formatDateForWeb(row[3]) : '',
          status: row[4]
        };
      });
  } catch (e) {
    console.error('評価期間取得エラー:', e);
    return [];
  }
}

function formatDateForWeb(date) {
  if (!date) return '';
  if (typeof date === 'string') return date;
  try {
    var d = new Date(date);
    var year = d.getFullYear();
    var month = ('0' + (d.getMonth() + 1)).slice(-2);
    var day = ('0' + d.getDate()).slice(-2);
    return year + '-' + month + '-' + day;
  } catch (e) {
    return '';
  }
}

function getActivePeriod() {
  var periods = getPeriods();
  var activePeriod = null;
  for (var i = 0; i < periods.length; i++) {
    if (periods[i].status === 'active') {
      activePeriod = periods[i];
      break;
    }
  }
  if (!activePeriod && periods.length > 0) {
    return periods[0];
  }
  return activePeriod || null;
}

function createNewPeriod(periodId, periodName, startDate, endDate) {
  return withLock(function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var periodSheet = ss.getSheetByName(PERIOD_MASTER_SHEET_NAME);
      if (!periodSheet) {
        periodSheet = ss.insertSheet(PERIOD_MASTER_SHEET_NAME);
        periodSheet.getRange('A1:E1').setValues([['期間ID', '期間名', '開始日', '終了日', 'ステータス']]).setFontWeight('bold');
      }
      var lastRow = periodSheet.getLastRow();
      if (lastRow > 1) {
        var statusRange = periodSheet.getRange(2, 5, lastRow - 1, 1);
        var statuses = statusRange.getValues();
        var updatedStatuses = statuses.map(function() { return ['closed']; });
        statusRange.setValues(updatedStatuses);
      }
      periodSheet.appendRow([periodId, periodName, startDate, endDate, 'active']);
      createPeriodSheets(ss, periodId);
      return { status: 'success', message: '評価期間「' + periodName + '」を作成しました。' };
    } catch (e) {
      return { status: 'error', message: 'エラー: ' + e.message };
    }
  });
}

function createPeriodSheets(ss, periodId) {
  var dataSheetName = 'データ_' + periodId;
  if (!ss.getSheetByName(dataSheetName)) {
    var dataSheet = ss.insertSheet(dataSheetName);
    var itemTemplates = getEvaluationItemsFromSheet(ss);
    var itemHeaders = itemTemplates.map(function(item) {
      return item.id + '.' + item.title;
    });
    var headers = [
      'タイムスタンプ', '被評価者', '評価者', '評価者の立場'
    ].concat(itemHeaders).concat([
      'A項目小計', 'B項目小計', '総合得点', '総合評価ランク',
      '良い点_1位', '良い点_2位', '良い点_3位',
      '改善点_1位', '改善点_2位', '改善点_3位', '具体的なコメント'
    ]);
    dataSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff');
  }
}

function getDataSheetName(periodId) {
  return periodId ? 'データ_' + periodId : 'データ';
}

function getSummarySheetName(periodId) {
  return periodId ? '総合評価_' + periodId : '総合評価';
}

// =============================================================
// 設定管理
// =============================================================

function loadSettings() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    if (!settingsSheet) return getDefaultSettings();
    var data = settingsSheet.getRange('A2:B20').getValues();
    var settings = {};
    data.forEach(function(row) {
      var key = row[0];
      var value = row[1];
      if (key && value !== '') {
        settings[key] = parseFloat(value);
      }
    });
    var defaultSettings = getDefaultSettings();
    Object.keys(defaultSettings).forEach(function(key) {
      if (settings[key] === undefined || settings[key] === null || isNaN(settings[key])) {
        settings[key] = defaultSettings[key];
      }
    });
    return settings;
  } catch (e) {
    console.error('設定読み込みエラー:', e);
    return getDefaultSettings();
  }
}

function getDefaultSettings() {
  return {
    'A項目重み': 60,
    'B項目重み': 40,
    'よくできる重み': 1.0,
    'できる重み': 0.6,
    '少しできる重み': 0.3,
    'できていない重み': 0,
    '優秀閾値': 240,
    '良好閾値': 180,
    '要努力閾値': 140,
    '本人評価重み': 0.1,
    '同僚評価重み': 0.4,
    '上司評価重み': 0.5,
    'ログイン機能有効': 1
  };
}

function getScoreWeight(settings, level) {
  var levelMap = {
    '10': 'よくできる重み',
    '6': 'できる重み',
    '3': '少しできる重み',
    '0': 'できていない重み'
  };
  var settingKey = levelMap[level.toString()];
  return settingKey ? settings[settingKey] : 0;
}

// =============================================================
// ユーザー認証
// =============================================================

function registerUser(evaluatorName, email, password) {
  return withLock(function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var authSheet = ss.getSheetByName(USER_AUTH_SHEET_NAME);
      if (!authSheet) {
        authSheet = ss.insertSheet(USER_AUTH_SHEET_NAME);
        authSheet.getRange('A1:D1').setValues([['評価者名', 'メールアドレス（ログインID）', 'パスワード', '登録日時']]).setFontWeight('bold');
      }
      var lastRow = authSheet.getLastRow();
      if (lastRow > 1) {
        var emails = authSheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
        var emailLower = email.trim().toLowerCase();
        for (var i = 0; i < emails.length; i++) {
          if (emails[i] && emails[i].toString().trim().toLowerCase() === emailLower) {
            return { status: 'error', message: 'このメールアドレスは既に登録されています。' };
          }
        }
      }
      var saveData = [evaluatorName, email.trim(), "'" + password.trim(), new Date()];
      authSheet.appendRow(saveData);
      try {
        var subject = 'スタッフ評価システム - 登録完了のお知らせ';
        var body = evaluatorName + ' 様\n\n' +
                   'スタッフ評価システムへの登録が完了しました。\n\n' +
                   '以下の情報でログインしてください：\n' +
                   '----------------------------------------\n' +
                   'ログインID（メールアドレス）: ' + email + '\n' +
                   'パスワード: ' + password + '\n' +
                   '----------------------------------------\n\n' +
                   '※このメールは大切に保管してください。\n' +
                   '※パスワードは他の人に知られないようご注意ください。\n\n' +
                   'スタッフ評価システム';
        MailApp.sendEmail(email.trim(), subject, body);
      } catch (mailError) {
        return { status: 'success', message: '登録が完了しました。（メール送信はエラーになりました）' };
      }
      return { status: 'success', message: '登録が完了しました。メールアドレスに登録情報を送信しました。' };
    } catch (e) {
      console.error('登録エラー:', e);
      return { status: 'error', message: '登録エラー: ' + e.message };
    }
  });
}

function authenticate(email, password) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settings = loadSettings();
    if (settings['ログイン機能有効'] === 0) {
      var evaluatorName = email;
      var evaluators = getListFromSheet(ss, EVALUATOR_MASTER_SHEET_NAME);
      if (!evaluators.includes(evaluatorName)) {
        return { status: 'error', message: '評価者マスタに登録されていない名前です。' };
      }
      var activePeriod = getActivePeriod();
      return {
        status: 'success',
        data: {
          evaluatorName: evaluatorName,
          evaluators: evaluators,
          evaluated: getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME),
          items: getEvaluationItemsFromSheet(ss),
          settings: settings,
          periods: getPeriods(),
          activePeriod: activePeriod,
          savedData: loadTempData(evaluatorName),
          isAdmin: isAdminUser(evaluatorName)
        }
      };
    }
    var authSheet = ss.getSheetByName(USER_AUTH_SHEET_NAME);
    if (!authSheet) {
      return { status: 'error', message: 'ユーザー認証情報シートが見つかりません。' };
    }
    var lastRow = authSheet.getLastRow();
    if (lastRow < 2) {
      return { status: 'error', message: '登録されているユーザーがいません。' };
    }
    var userData = authSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < userData.length; i++) {
      var storedEmail = userData[i][1] ? userData[i][1].trim() : '';
      var storedPassword = userData[i][2] ? userData[i][2].toString().trim() : '';
      if (storedPassword.startsWith("'")) {
        storedPassword = storedPassword.substring(1);
      }
      if (storedEmail.toLowerCase() === email.trim().toLowerCase() &&
          storedPassword === password.trim()) {
        var evalName = userData[i][0];
        var ap = getActivePeriod();
        return {
          status: 'success',
          data: {
            evaluatorName: evalName,
            evaluators: getListFromSheet(ss, EVALUATOR_MASTER_SHEET_NAME),
            evaluated: getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME),
            items: getEvaluationItemsFromSheet(ss),
            settings: settings,
            periods: getPeriods(),
            activePeriod: ap,
            savedData: loadTempData(evalName),
            isAdmin: isAdminUser(evalName)
          }
        };
      }
    }
    return { status: 'error', message: 'メールアドレスまたはパスワードが正しくありません。' };
  } catch (e) {
    console.error('認証エラー:', e);
    return { status: 'error', message: '認証エラー: ' + e.message };
  }
}

/**
 * 管理者かどうかを判定（院長 or 上司フラグ）
 */
function isAdminUser(evaluatorName) {
  try {
    if (evaluatorName === PRINCIPAL_NAME) return true;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var details = getEvaluatorDetails(ss);
    return details[evaluatorName] && details[evaluatorName].isSupervisor;
  } catch (e) {
    return false;
  }
}

function checkUserRegistration(evaluatorName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var authSheet = ss.getSheetByName(USER_AUTH_SHEET_NAME);
    if (!authSheet) return { isRegistered: false };
    var lastRow = authSheet.getLastRow();
    if (lastRow < 2) return { isRegistered: false };
    var evaluatorNames = authSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    return { isRegistered: evaluatorNames.includes(evaluatorName) };
  } catch (e) {
    return { isRegistered: false, error: e.message };
  }
}

function getRegisteredUsers() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var authSheet = ss.getSheetByName(USER_AUTH_SHEET_NAME);
    if (!authSheet || authSheet.getLastRow() < 2) return { users: [] };
    var userData = authSheet.getRange(2, 1, authSheet.getLastRow() - 1, 3).getValues();
    var users = userData.map(function(row) {
      return { name: row[0], email: row[1], hasPassword: row[2] ? true : false };
    });
    return { users: users };
  } catch (e) {
    return { users: [], error: e.message };
  }
}

// =============================================================
// 一時保存（LockService使用）
// =============================================================

function saveTempData(evaluatorName, data) {
  return withLock(function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var tempSheet = ss.getSheetByName(TEMP_SAVE_SHEET_NAME);
      if (!tempSheet) {
        tempSheet = ss.insertSheet(TEMP_SAVE_SHEET_NAME);
        tempSheet.getRange('A1:C1').setValues([['評価者', '保存日時', 'データ']]).setFontWeight('bold');
      }
      var lastRow = tempSheet.getLastRow();
      var rowToUpdate = -1;
      if (lastRow > 1) {
        var evaluatorColumn = tempSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < evaluatorColumn.length; i++) {
          if (evaluatorColumn[i][0] === evaluatorName) {
            rowToUpdate = i + 2;
            break;
          }
        }
      }
      var saveData = [evaluatorName, new Date(), JSON.stringify(data)];
      if (rowToUpdate > 0) {
        tempSheet.getRange(rowToUpdate, 1, 1, 3).setValues([saveData]);
      } else {
        tempSheet.appendRow(saveData);
      }
      return { status: 'success', message: '自動保存しました' };
    } catch (e) {
      return { status: 'error', message: '保存エラー: ' + e.message };
    }
  });
}

function loadTempData(evaluatorName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var tempSheet = ss.getSheetByName(TEMP_SAVE_SHEET_NAME);
    if (!tempSheet) return null;
    var lastRow = tempSheet.getLastRow();
    if (lastRow < 2) return null;
    var data = tempSheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === evaluatorName) {
        try { return JSON.parse(data[i][2]); } catch (e) { return null; }
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function clearTempData(evaluatorName) {
  return withLock(function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var tempSheet = ss.getSheetByName(TEMP_SAVE_SHEET_NAME);
      if (!tempSheet) return { status: 'success' };
      var lastRow = tempSheet.getLastRow();
      if (lastRow < 2) return { status: 'success' };
      var evaluatorColumn = tempSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = evaluatorColumn.length - 1; i >= 0; i--) {
        if (evaluatorColumn[i][0] === evaluatorName) {
          tempSheet.deleteRow(i + 2);
        }
      }
      return { status: 'success', message: '一時保存データを削除しました' };
    } catch (e) {
      return { status: 'error', message: '削除エラー: ' + e.message };
    }
  });
}

// =============================================================
// 評価データ保存（LockService使用 - 同時書き込み対策）
// =============================================================

function saveMultipleEvaluations(evaluations, periodId) {
  return withLock(function() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      if (!periodId) {
        var activePeriod = getActivePeriod();
        periodId = activePeriod ? activePeriod.id : null;
      }
      var dataSheetName = getDataSheetName(periodId);
      var sheet = ss.getSheetByName(dataSheetName);
      if (!sheet) {
        return { status: 'error', message: 'データシート「' + dataSheetName + '」が見つかりません。' };
      }
      var itemTemplates = getEvaluationItemsFromSheet(ss);
      var evaluatorDetails = getEvaluatorDetails(ss);
      var settings = loadSettings();
      var dataSheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var headerMap = {};
      dataSheetHeaders.forEach(function(header, i) { headerMap[header] = i; });

      // 全評価データを一括で配列に収集
      var allNewRows = [];

      evaluations.forEach(function(formData) {
        var scores = {};
        var totalA = 0;
        var totalB = 0;

        itemTemplates.forEach(function(item, index) {
          var subItemScores = formData.checks[index] || [];
          var mainItemScore = 0;
          subItemScores.forEach(function(score) {
            var numScore = Number(score);
            if (!isNaN(numScore)) {
              mainItemScore += numScore;
            }
          });
          scores['score' + item.id] = mainItemScore;
          if (item.group === 'A') {
            totalA += mainItemScore;
          } else if (item.group === 'B') {
            totalB += mainItemScore;
          }
        });

        var totalScore = totalA + totalB;
        var rank = '';
        if (totalScore >= settings['優秀閾値']) rank = '優秀';
        else if (totalScore >= settings['良好閾値']) rank = '良好';
        else if (totalScore >= settings['要努力閾値']) rank = '要努力';
        else rank = '要指導';

        var evaluatorRole = '';
        if (formData.evaluator === formData.evaluatedPerson) {
          evaluatorRole = '本人';
        } else if (formData.evaluator === PRINCIPAL_NAME) {
          evaluatorRole = '院長';
        } else if (evaluatorDetails[formData.evaluator] && evaluatorDetails[formData.evaluator].isSupervisor) {
          evaluatorRole = '上司';
        } else {
          evaluatorRole = '同僚';
        }

        var newRow = new Array(dataSheetHeaders.length).fill('');
        newRow[headerMap['タイムスタンプ']] = new Date();
        newRow[headerMap['被評価者']] = formData.evaluatedPerson;
        newRow[headerMap['評価者']] = formData.evaluator;
        newRow[headerMap['評価者の立場']] = evaluatorRole;

        itemTemplates.forEach(function(item) {
          var headerName = item.id + '.' + item.title;
          if (headerMap.hasOwnProperty(headerName)) {
            newRow[headerMap[headerName]] = scores['score' + item.id] || 0;
          }
        });

        newRow[headerMap['A項目小計']] = totalA;
        newRow[headerMap['B項目小計']] = totalB;
        newRow[headerMap['総合得点']] = totalScore;
        newRow[headerMap['総合評価ランク']] = rank;

        var goodPoints = formData.goodPoints || {};
        newRow[headerMap['良い点_1位']] = goodPoints['1位'] || '';
        newRow[headerMap['良い点_2位']] = goodPoints['2位'] || '';
        newRow[headerMap['良い点_3位']] = goodPoints['3位'] || '';
        var improvementPoints = formData.improvementPoints || {};
        newRow[headerMap['改善点_1位']] = improvementPoints['1位'] || '';
        newRow[headerMap['改善点_2位']] = improvementPoints['2位'] || '';
        newRow[headerMap['改善点_3位']] = improvementPoints['3位'] || '';
        newRow[headerMap['具体的なコメント']] = formData.comments;

        allNewRows.push(newRow);
      });

      // 一括書き込み（appendRowの繰り返しではなくsetValuesで高速化）
      if (allNewRows.length > 0) {
        var lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, allNewRows.length, dataSheetHeaders.length).setValues(allNewRows);
      }

      if (evaluations.length > 0 && evaluations[0].evaluator) {
        // ロック内なので直接削除処理
        var tempSheet = ss.getSheetByName(TEMP_SAVE_SHEET_NAME);
        if (tempSheet) {
          var tempLastRow = tempSheet.getLastRow();
          if (tempLastRow > 1) {
            var evalCol = tempSheet.getRange(2, 1, tempLastRow - 1, 1).getValues();
            for (var ti = evalCol.length - 1; ti >= 0; ti--) {
              if (evalCol[ti][0] === evaluations[0].evaluator) {
                tempSheet.deleteRow(ti + 2);
              }
            }
          }
        }
      }

      updateSummarySheet(periodId);

      return { status: 'success', message: 'すべての評価データが正常に送信されました。' };
    } catch (e) {
      Logger.log('エラー発生: ' + e.message);
      Logger.log('スタック: ' + e.stack);
      return { status: 'error', message: 'エラーが発生しました: ' + e.message };
    }
  });
}

// =============================================================
// 評価進捗管理
// =============================================================

/**
 * 評価進捗を取得（管理者画面用）
 */
function getEvaluationProgress(periodId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!periodId) {
      var activePeriod = getActivePeriod();
      periodId = activePeriod ? activePeriod.id : null;
    }
    var dataSheetName = getDataSheetName(periodId);
    var dataSheet = ss.getSheetByName(dataSheetName);
    var evaluatedList = getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME);
    var evaluatorList = getListFromSheet(ss, EVALUATOR_MASTER_SHEET_NAME);

    if (!dataSheet || dataSheet.getLastRow() < 2) {
      return {
        status: 'success',
        data: {
          evaluatedList: evaluatedList,
          evaluatorList: evaluatorList,
          progress: {},
          totalExpected: evaluatedList.length * evaluatorList.length,
          totalCompleted: 0
        }
      };
    }

    var dataRange = dataSheet.getDataRange();
    var dataValues = dataRange.getValues();
    var headers = dataValues[0];
    var evaluatedIdx = headers.indexOf('被評価者');
    var evaluatorIdx = headers.indexOf('評価者');

    var progress = {};
    evaluatedList.forEach(function(person) {
      progress[person] = {
        evaluators: [],
        total: evaluatorList.length,
        completed: 0
      };
    });

    for (var i = 1; i < dataValues.length; i++) {
      var evaluated = dataValues[i][evaluatedIdx];
      var evaluator = dataValues[i][evaluatorIdx];
      if (progress[evaluated] && !progress[evaluated].evaluators.includes(evaluator)) {
        progress[evaluated].evaluators.push(evaluator);
        progress[evaluated].completed++;
      }
    }

    var totalCompleted = 0;
    Object.keys(progress).forEach(function(key) {
      totalCompleted += progress[key].completed;
    });

    return {
      status: 'success',
      data: {
        evaluatedList: evaluatedList,
        evaluatorList: evaluatorList,
        progress: progress,
        totalExpected: evaluatedList.length * evaluatorList.length,
        totalCompleted: totalCompleted
      }
    };
  } catch (e) {
    return { status: 'error', message: 'エラー: ' + e.message };
  }
}

// =============================================================
// 評価シート生成（HTML）
// =============================================================

/**
 * 評価完了処理 - 全スタッフの評価シートHTMLを生成
 */
function completeEvaluationAndGenerateSheets(periodId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settings = loadSettings();

    if (!periodId) {
      var activePeriod = getActivePeriod();
      periodId = activePeriod ? activePeriod.id : null;
    }

    // 総合評価シートを最新に更新
    updateSummarySheet(periodId);

    var dataSheetName = getDataSheetName(periodId);
    var dataSheet = ss.getSheetByName(dataSheetName);

    if (!dataSheet || dataSheet.getLastRow() < 2) {
      return { status: 'error', message: '評価データが見つかりません。' };
    }

    var dataValues = dataSheet.getDataRange().getValues();
    var headers = dataValues[0];
    var headerMap = {};
    headers.forEach(function(h, i) { headerMap[h] = i; });

    var evaluatedList = getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME);
    var itemTemplates = getEvaluationItemsFromSheet(ss);

    var sheets = [];

    evaluatedList.forEach(function(person) {
      var personEvals = [];
      for (var i = 1; i < dataValues.length; i++) {
        if (dataValues[i][headerMap['被評価者']] === person) {
          personEvals.push(dataValues[i]);
        }
      }

      if (personEvals.length === 0) return;

      var selfEvals = personEvals.filter(function(r) { return r[headerMap['評価者の立場']] === '本人'; });
      var colleagueEvals = personEvals.filter(function(r) { return r[headerMap['評価者の立場']] === '同僚'; });
      var supervisorEvals = personEvals.filter(function(r) {
        return r[headerMap['評価者の立場']] === '上司' || r[headerMap['評価者の立場']] === '院長';
      });

      var WEIGHT_SELF = settings['本人評価重み'];
      var WEIGHT_COLLEAGUE = settings['同僚評価重み'];
      var WEIGHT_SUPERVISOR = settings['上司評価重み'];

      var itemResults = itemTemplates.map(function(item) {
        var colName = item.id + '.' + item.title;
        var colIndex = headerMap[colName];
        if (colIndex === undefined) return null;

        var selfAvg = calculateAverageForItem(selfEvals, colIndex);
        var colleagueAvg = calculateAverageForItem(colleagueEvals, colIndex);
        var supervisorAvg = calculateAverageForItem(supervisorEvals, colIndex);
        var weightedAvg = (selfAvg * WEIGHT_SELF) + (colleagueAvg * WEIGHT_COLLEAGUE) + (supervisorAvg * WEIGHT_SUPERVISOR);

        return {
          id: item.id,
          title: item.title,
          group: item.group,
          checks: item.checks,
          selfScore: Math.round(selfAvg * 10) / 10,
          colleagueScore: Math.round(colleagueAvg * 10) / 10,
          supervisorScore: Math.round(supervisorAvg * 10) / 10,
          weightedScore: Math.round(weightedAvg * 10) / 10
        };
      }).filter(function(r) { return r !== null; });

      // 立場別の合計
      var selfATotal = calculateAverageForItem(selfEvals, headerMap['A項目小計']);
      var selfBTotal = calculateAverageForItem(selfEvals, headerMap['B項目小計']);
      var colleagueATotal = calculateAverageForItem(colleagueEvals, headerMap['A項目小計']);
      var colleagueBTotal = calculateAverageForItem(colleagueEvals, headerMap['B項目小計']);
      var supervisorATotal = calculateAverageForItem(supervisorEvals, headerMap['A項目小計']);
      var supervisorBTotal = calculateAverageForItem(supervisorEvals, headerMap['B項目小計']);

      var weightedA = (selfATotal * WEIGHT_SELF) + (colleagueATotal * WEIGHT_COLLEAGUE) + (supervisorATotal * WEIGHT_SUPERVISOR);
      var weightedB = (selfBTotal * WEIGHT_SELF) + (colleagueBTotal * WEIGHT_COLLEAGUE) + (supervisorBTotal * WEIGHT_SUPERVISOR);
      var totalWeighted = weightedA + weightedB;

      var rank = '';
      if (totalWeighted >= settings['優秀閾値']) rank = '優秀';
      else if (totalWeighted >= settings['良好閾値']) rank = '良好';
      else if (totalWeighted >= settings['要努力閾値']) rank = '要努力';
      else rank = '要指導';

      // コメント収集
      var comments = [];
      personEvals.forEach(function(row) {
        var comment = row[headerMap['具体的なコメント']];
        var role = row[headerMap['評価者の立場']];
        if (comment && comment.toString().trim() !== '') {
          comments.push({ text: comment.toString().trim(), role: role });
        }
      });

      // 良い点・改善点の集計
      var goodPointsCounts = {};
      var improvementPointsCounts = {};
      personEvals.forEach(function(row) {
        ['良い点_1位', '良い点_2位', '良い点_3位'].forEach(function(key) {
          var val = row[headerMap[key]];
          if (val && val.toString().trim()) {
            var v = val.toString().trim();
            goodPointsCounts[v] = (goodPointsCounts[v] || 0) + 1;
          }
        });
        ['改善点_1位', '改善点_2位', '改善点_3位'].forEach(function(key) {
          var val = row[headerMap[key]];
          if (val && val.toString().trim()) {
            var v = val.toString().trim();
            improvementPointsCounts[v] = (improvementPointsCounts[v] || 0) + 1;
          }
        });
      });

      // ソートして上位3つ
      var sortedGoodPoints = Object.keys(goodPointsCounts).sort(function(a, b) {
        return goodPointsCounts[b] - goodPointsCounts[a];
      }).slice(0, 3);

      var sortedImprovementPoints = Object.keys(improvementPointsCounts).sort(function(a, b) {
        return improvementPointsCounts[b] - improvementPointsCounts[a];
      }).slice(0, 3);

      sheets.push({
        personName: person,
        items: itemResults,
        totalA: Math.round(weightedA * 10) / 10,
        totalB: Math.round(weightedB * 10) / 10,
        totalScore: Math.round(totalWeighted * 10) / 10,
        rank: rank,
        maxScore: MAX_TOTAL_SCORE,
        evaluatorCount: personEvals.length,
        selfCount: selfEvals.length,
        colleagueCount: colleagueEvals.length,
        supervisorCount: supervisorEvals.length,
        comments: comments,
        goodPoints: sortedGoodPoints,
        improvementPoints: sortedImprovementPoints,
        settings: {
          weightSelf: WEIGHT_SELF,
          weightColleague: WEIGHT_COLLEAGUE,
          weightSupervisor: WEIGHT_SUPERVISOR
        }
      });
    });

    // 期間のステータスをclosedに変更
    var periodSheet = ss.getSheetByName(PERIOD_MASTER_SHEET_NAME);
    if (periodSheet) {
      var pLastRow = periodSheet.getLastRow();
      if (pLastRow > 1) {
        var pData = periodSheet.getRange(2, 1, pLastRow - 1, 5).getValues();
        for (var pi = 0; pi < pData.length; pi++) {
          if (pData[pi][0] == periodId) {
            periodSheet.getRange(pi + 2, 5).setValue('completed');
            break;
          }
        }
      }
    }

    var periodInfo = null;
    var periods = getPeriods();
    for (var pp = 0; pp < periods.length; pp++) {
      if (periods[pp].id == periodId) {
        periodInfo = periods[pp];
        break;
      }
    }

    return {
      status: 'success',
      data: {
        sheets: sheets,
        periodInfo: periodInfo,
        generatedAt: new Date().toISOString()
      }
    };
  } catch (e) {
    console.error('評価シート生成エラー:', e);
    return { status: 'error', message: 'エラー: ' + e.message };
  }
}

// =============================================================
// 総合評価シート更新
// =============================================================

function updateSummarySheet(periodId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settings = loadSettings();
    if (!periodId) {
      var activePeriod = getActivePeriod();
      periodId = activePeriod ? activePeriod.id : null;
    }
    var dataSheetName = getDataSheetName(periodId);
    var summarySheetName = getSummarySheetName(periodId);
    var dataSheet = ss.getSheetByName(dataSheetName);
    if (!dataSheet || dataSheet.getLastRow() < 2) {
      console.log('評価データが見つかりません: ' + dataSheetName);
      return;
    }
    var dataRange = dataSheet.getDataRange();
    var dataValues = dataRange.getValues();
    var dataHeaders = dataValues[0];
    var evaluationData = dataValues.slice(1);
    var itemHeadersInData = [];
    var headerMap = {};
    dataHeaders.forEach(function(header, index) {
      headerMap[header] = index;
      if (header && /^\d+\./.test(header)) {
        itemHeadersInData.push(header);
      }
    });
    var currentItemTemplates = getEvaluationItemsFromSheet(ss);
    var validItems = currentItemTemplates.filter(function(item) {
      var expectedHeader = item.id + '.' + item.title;
      return itemHeadersInData.includes(expectedHeader) || headerMap.hasOwnProperty(expectedHeader);
    });
    if (validItems.length === 0) {
      console.log('有効な評価項目が見つかりません');
      return;
    }
    var itemHeadersForSummary = [];
    validItems.forEach(function(item) {
      itemHeadersForSummary.push(item.title + '(本人)');
      itemHeadersForSummary.push(item.title + '(同僚)');
      itemHeadersForSummary.push(item.title + '(上司)');
      itemHeadersForSummary.push(item.title + '(総合)');
    });

    var headers = ['被評価者', '総合得点(A)', '総合得点(B)', '総合得点(全体)']
      .concat(itemHeadersForSummary)
      .concat(['具体的なコメント', '最終更新日時']);

    var summarySheet = ss.getSheetByName(summarySheetName);
    if (summarySheet) ss.deleteSheet(summarySheet);
    summarySheet = ss.insertSheet(summarySheetName, 0);
    summarySheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight('bold')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff');
    summarySheet.setFrozenRows(1);
    summarySheet.setFrozenColumns(1);

    var evaluatedPersonsSet = {};
    evaluationData.forEach(function(row) {
      var p = row[headerMap['被評価者']];
      if (p) evaluatedPersonsSet[p] = true;
    });
    var evaluatedPersons = Object.keys(evaluatedPersonsSet);

    var WEIGHT_SELF = settings['本人評価重み'];
    var WEIGHT_COLLEAGUE = settings['同僚評価重み'];
    var WEIGHT_SUPERVISOR = settings['上司評価重み'];

    var summaryRows = [];
    evaluatedPersons.forEach(function(person) {
      var personEvaluations = evaluationData.filter(function(row) {
        return row[headerMap['被評価者']] === person;
      });
      if (personEvaluations.length === 0) return;

      var selfEvals = personEvaluations.filter(function(r) { return r[headerMap['評価者の立場']] === '本人'; });
      var colleagueEvals = personEvaluations.filter(function(r) { return r[headerMap['評価者の立場']] === '同僚'; });
      var supervisorEvals = personEvaluations.filter(function(r) {
        return r[headerMap['評価者の立場']] === '上司' || r[headerMap['評価者の立場']] === '院長';
      });

      var itemScores = [];
      validItems.forEach(function(item) {
        var colName = item.id + '.' + item.title;
        var colIndex = headerMap[colName];
        if (colIndex === undefined) {
          itemScores.push(0, 0, 0, 0);
          return;
        }
        var selfAvg = calculateAverageForItem(selfEvals, colIndex);
        var colleagueAvg = calculateAverageForItem(colleagueEvals, colIndex);
        var supervisorAvg = calculateAverageForItem(supervisorEvals, colIndex);
        var weightedAvg = (selfAvg * WEIGHT_SELF) + (colleagueAvg * WEIGHT_COLLEAGUE) + (supervisorAvg * WEIGHT_SUPERVISOR);
        itemScores.push(selfAvg, colleagueAvg, supervisorAvg, weightedAvg);
      });

      var selfAScore = calculateAverageForItem(selfEvals, headerMap['A項目小計']);
      var selfBScore = calculateAverageForItem(selfEvals, headerMap['B項目小計']);
      var colleagueAScore = calculateAverageForItem(colleagueEvals, headerMap['A項目小計']);
      var colleagueBScore = calculateAverageForItem(colleagueEvals, headerMap['B項目小計']);
      var supervisorAScore = calculateAverageForItem(supervisorEvals, headerMap['A項目小計']);
      var supervisorBScore = calculateAverageForItem(supervisorEvals, headerMap['B項目小計']);

      var weightedAScore = (selfAScore * WEIGHT_SELF) + (colleagueAScore * WEIGHT_COLLEAGUE) + (supervisorAScore * WEIGHT_SUPERVISOR);
      var weightedBScore = (selfBScore * WEIGHT_SELF) + (colleagueBScore * WEIGHT_COLLEAGUE) + (supervisorBScore * WEIGHT_SUPERVISOR);
      var totalWeightedScore = weightedAScore + weightedBScore;

      var aggregatedComments = [];
      personEvaluations.forEach(function(row) {
        var comment = row[headerMap['具体的なコメント']];
        var role = row[headerMap['評価者の立場']];
        if (comment && comment.toString().trim() !== '') {
          var roleSuffix = '';
          if (role === '本人') roleSuffix = '(本人)';
          else if (role === '同僚') roleSuffix = '(同僚)';
          else if (role === '上司' || role === '院長') roleSuffix = '(上司)';
          aggregatedComments.push(comment.toString().trim() + ' ' + roleSuffix);
        }
      });

      var row = [
        person,
        Math.round(weightedAScore * 10) / 10,
        Math.round(weightedBScore * 10) / 10,
        Math.round(totalWeightedScore * 10) / 10
      ];
      itemScores.forEach(function(score) {
        row.push(Math.round(score * 10) / 10);
      });
      row.push(aggregatedComments.join('\n'));
      row.push(new Date());
      summaryRows.push(row);
    });

    if (summaryRows.length > 0) {
      summarySheet.getRange(2, 1, summaryRows.length, headers.length).setValues(summaryRows);
      var numScoreCols = validItems.length * 4 + 3;
      summarySheet.getRange(2, 2, summaryRows.length, numScoreCols).setNumberFormat('0.0');
      var sortRange = summarySheet.getRange(2, 1, summaryRows.length, headers.length);
      sortRange.sort({ column: 4, ascending: false });
    }

    applyConditionalFormatting(summarySheet, validItems.length, settings);
    console.log('総合評価シートが更新されました: ' + summarySheetName + ' (' + validItems.length + '項目)');
  } catch (e) {
    console.error('総合評価シート更新エラー:', e.message, e.stack);
    throw e;
  }
}

function calculateAverageForItem(evaluations, colIndex) {
  if (!evaluations || evaluations.length === 0 || colIndex === undefined) return 0;
  var scores = evaluations
    .map(function(row) { return row[colIndex]; })
    .filter(function(score) { return score !== null && score !== undefined && !isNaN(score); })
    .map(Number);
  if (scores.length === 0) return 0;
  var total = scores.reduce(function(sum, score) { return sum + score; }, 0);
  return total / scores.length;
}

function applyConditionalFormatting(sheet, numItems, settings) {
  try {
    sheet.setColumnWidth(1, 120);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 110);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    var totalScoreRange = sheet.getRange(2, 4, lastRow - 1, 1);
    var rules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThanOrEqualTo(settings['優秀閾値'])
        .setBackground('#d5e8d4').setFontColor('#000000')
        .setRanges([totalScoreRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(settings['良好閾値'], settings['優秀閾値'] - 0.1)
        .setBackground('#d4e1f5').setFontColor('#000000')
        .setRanges([totalScoreRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberBetween(settings['要努力閾値'], settings['良好閾値'] - 0.1)
        .setBackground('#fff2cc').setFontColor('#000000')
        .setRanges([totalScoreRange]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThan(settings['要努力閾値'])
        .setBackground('#f4cccc').setFontColor('#000000')
        .setRanges([totalScoreRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);

    var itemTemplates = getEvaluationItemsFromSheet(SpreadsheetApp.openById(SPREADSHEET_ID));
    var itemStartCol = 5;
    var groupAColor = '#fce8b2';
    var groupBColor = '#c9daf8';
    for (var i = 0; i < numItems && i < itemTemplates.length; i++) {
      var baseCol = itemStartCol + (i * 4);
      var bgColor = itemTemplates[i].group === 'A' ? groupAColor : groupBColor;
      if (baseCol + 3 <= sheet.getMaxColumns()) {
        sheet.getRange(1, baseCol, lastRow, 4).setBackground(bgColor);
        sheet.setColumnWidths(baseCol, 4, 85);
        sheet.getRange(2, baseCol + 3, lastRow - 1, 1).setFontWeight('bold');
      }
    }
    var commentCol = itemStartCol + (numItems * 4);
    if (commentCol <= sheet.getMaxColumns()) {
      sheet.getRange(1, commentCol, lastRow, 1).setBackground('#e6e6e6');
      sheet.setColumnWidth(commentCol, 400);
      sheet.getRange(2, commentCol, lastRow - 1, 1).setWrap(true).setVerticalAlignment('top');
      if (commentCol + 1 <= sheet.getMaxColumns()) {
        sheet.setColumnWidth(commentCol + 1, 150);
        sheet.getRange(2, commentCol + 1, lastRow - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm:ss');
      }
    }
  } catch (e) {
    console.error('条件付き書式適用エラー:', e.message);
  }
}

// =============================================================
// 期間比較
// =============================================================

function createComparisonSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settings = loadSettings();
    var periods = getPeriods().filter(function(p) { return p.status !== 'planned'; });
    if (periods.length < 2) {
      SpreadsheetApp.getUi().alert('比較するには2つ以上の評価期間が必要です。');
      return;
    }
    var compSheet = ss.getSheetByName(COMPARISON_SHEET_NAME);
    if (compSheet) ss.deleteSheet(compSheet);
    compSheet = ss.insertSheet(COMPARISON_SHEET_NAME, 0);

    var periodHeaders = [];
    periods.forEach(function(p) {
      periodHeaders.push(p.name + '得点');
      periodHeaders.push(p.name + 'ランク');
    });
    var headers = ['被評価者'].concat(periodHeaders).concat(['推移']);
    compSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');

    var evaluatedList = getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME);
    var comparisonData = evaluatedList.map(function(person) {
      var row = [person];
      var scores = [];
      periods.forEach(function(period) {
        var summarySheetName = getSummarySheetName(period.id);
        var summarySheet = ss.getSheetByName(summarySheetName);
        if (summarySheet && summarySheet.getLastRow() > 1) {
          var data = summarySheet.getDataRange().getValues();
          var headerRow = data[0];
          var personRow = null;
          for (var d = 1; d < data.length; d++) {
            if (data[d][0] === person) { personRow = data[d]; break; }
          }
          if (personRow) {
            var scoreIndex = headerRow.indexOf('総合得点(全体)');
            var score = personRow[scoreIndex] || 0;
            scores.push(score);
            var rank = '';
            if (score >= settings['優秀閾値']) rank = '優秀';
            else if (score >= settings['良好閾値']) rank = '良好';
            else if (score >= settings['要努力閾値']) rank = '要努力';
            else rank = '要指導';
            row.push(score, rank);
          } else {
            row.push('', '未評価');
            scores.push(0);
          }
        } else {
          row.push('', '未評価');
          scores.push(0);
        }
      });
      if (scores.length >= 2 && scores.every(function(s) { return s > 0; })) {
        var diff = scores[scores.length - 1] - scores[0];
        var trend = '';
        if (diff > 10) trend = '大幅改善';
        else if (diff > 0) trend = '改善';
        else if (diff === 0) trend = '横ばい';
        else if (diff > -10) trend = '低下';
        else trend = '大幅低下';
        row.push(trend);
      } else {
        row.push('-');
      }
      return row;
    });

    if (comparisonData.length > 0) {
      compSheet.getRange(2, 1, comparisonData.length, headers.length).setValues(comparisonData);
      for (var ci = 0; ci < periods.length; ci++) {
        var scoreCol = 2 + (ci * 2);
        compSheet.getRange(2, scoreCol, comparisonData.length, 1).setNumberFormat('0.0');
      }
      compSheet.setColumnWidth(1, 120);
      compSheet.setColumnWidth(headers.length, 150);
    }
    SpreadsheetApp.getUi().alert('期間比較シートを作成しました。\n比較期間数: ' + periods.length);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

// =============================================================
// データシートヘッダー更新
// =============================================================

function updateDataSheetHeaders(periodId) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!periodId) {
      var activePeriod = getActivePeriod();
      periodId = activePeriod ? activePeriod.id : null;
    }
    var dataSheetName = getDataSheetName(periodId);
    var dataSheet = ss.getSheetByName(dataSheetName);
    var itemTemplates = getEvaluationItemsFromSheet(ss);
    if (!dataSheet) {
      SpreadsheetApp.getUi().alert('データシート「' + dataSheetName + '」が見つかりません。');
      return;
    }
    var lastRow = dataSheet.getLastRow();
    if (lastRow > 1) {
      var result = SpreadsheetApp.getUi().alert(
        '警告',
        '既存のデータがあります。ヘッダーを更新すると、データとの整合性が失われる可能性があります。\n続行しますか？',
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
      );
      if (result !== SpreadsheetApp.getUi().Button.OK) return;
    }
    var itemHeaders = itemTemplates.map(function(item) { return item.id + '.' + item.title; });
    var headers = ['タイムスタンプ', '被評価者', '評価者', '評価者の立場']
      .concat(itemHeaders)
      .concat([
        'A項目小計', 'B項目小計', '総合得点', '総合評価ランク',
        '良い点_1位', '良い点_2位', '良い点_3位',
        '改善点_1位', '改善点_2位', '改善点_3位', '具体的なコメント'
      ]);
    dataSheet.getRange(1, 1, 1, headers.length)
      .setValues([headers]).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    SpreadsheetApp.getUi().alert('データシート「' + dataSheetName + '」のヘッダーを更新しました。\n項目数: ' + itemTemplates.length);
  } catch (e) {
    SpreadsheetApp.getUi().alert('エラー: ' + e.message);
  }
}

// =============================================================
// トリガー・メニュー
// =============================================================

function setupDataChangesTrigger() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(function(trigger) {
      if (trigger.getHandlerFunction() === 'onDataSheetChange') {
        ScriptApp.deleteTrigger(trigger);
      }
    });
    ScriptApp.newTrigger('onDataSheetChange').forSpreadsheet(ss).onChange().create();
    SpreadsheetApp.getUi().alert('総合評価の自動更新トリガーが設定されました。');
  } catch (e) {
    SpreadsheetApp.getUi().alert('トリガー設定エラー: ' + e.message);
  }
}

function onDataSheetChange(e) {
  var changedSheetName = e.source.getActiveSheet().getName();
  if (changedSheetName.startsWith('データ_')) {
    try {
      var periodId = changedSheetName.replace('データ_', '');
      updateSummarySheet(periodId);
    } catch (err) {
      console.error('データ変更時の総合評価更新エラー:', err.stack);
    }
  }
}

function manualUpdateSummary() {
  try {
    var activePeriod = getActivePeriod();
    if (!activePeriod) {
      SpreadsheetApp.getUi().alert('アクティブな評価期間がありません。');
      return;
    }
    updateSummarySheet(activePeriod.id);
    SpreadsheetApp.getUi().alert('総合評価シート「' + getSummarySheetName(activePeriod.id) + '」を手動更新しました。');
  } catch (e) {
    SpreadsheetApp.getUi().alert('更新エラー: ' + e.message);
  }
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('評価システム管理')
    .addItem('初期設定（全シート作成）', 'setupSpreadsheet')
    .addSeparator()
    .addItem('システム診断', 'diagnoseSystem')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('評価期間管理')
      .addItem('新しい評価期間を作成', 'showCreatePeriodDialog')
      .addItem('期間比較シートを作成', 'createComparisonSheet'))
    .addSeparator()
    .addItem('データシートのヘッダーを更新', 'updateDataSheetHeadersDialog')
    .addItem('総合評価シートを手動更新', 'manualUpdateSummary')
    .addItem('総合評価の自動更新を設定(onChange)', 'setupDataChangesTrigger')
    .addToUi();
}

function showCreatePeriodDialog() {
  var ui = SpreadsheetApp.getUi();
  var periodIdResponse = ui.prompt('評価期間の作成', '期間IDを入力してください:', ui.ButtonSet.OK_CANCEL);
  if (periodIdResponse.getSelectedButton() !== ui.Button.OK) return;
  var periodId = periodIdResponse.getResponseText().trim();
  if (!periodId) { ui.alert('期間IDを入力してください。'); return; }
  var periodNameResponse = ui.prompt('評価期間の作成', '期間名を入力してください:', ui.ButtonSet.OK_CANCEL);
  if (periodNameResponse.getSelectedButton() !== ui.Button.OK) return;
  var periodName = periodNameResponse.getResponseText().trim();
  if (!periodName) { ui.alert('期間名を入力してください。'); return; }
  var startDateResponse = ui.prompt('評価期間の作成', '開始日（例: 2024/01/01）を入力してください:', ui.ButtonSet.OK_CANCEL);
  if (startDateResponse.getSelectedButton() !== ui.Button.OK) return;
  var startDate = startDateResponse.getResponseText().trim();
  var endDateResponse = ui.prompt('評価期間の作成', '終了日（例: 2024/03/31）を入力してください:', ui.ButtonSet.OK_CANCEL);
  if (endDateResponse.getSelectedButton() !== ui.Button.OK) return;
  var endDate = endDateResponse.getResponseText().trim();
  var result = createNewPeriod(periodId, periodName, startDate, endDate);
  ui.alert(result.message);
}

function updateDataSheetHeadersDialog() {
  var activePeriod = getActivePeriod();
  if (activePeriod) {
    updateDataSheetHeaders(activePeriod.id);
  } else {
    SpreadsheetApp.getUi().alert('アクティブな評価期間がありません。');
  }
}

// =============================================================
// 初期設定
// =============================================================

function setupSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var messages = [];

  try {
    if (!ss.getSheetByName(SETTINGS_SHEET_NAME)) {
      var settingsSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
      settingsSheet.getRange('A1:B1').setValues([['設定項目', '値']]).setFontWeight('bold');
      var defaultSettings = getDefaultSettings();
      var settingsData = [
        ['A項目重み', defaultSettings['A項目重み']],
        ['B項目重み', defaultSettings['B項目重み']],
        ['', ''],
        ['=== 評価レベル重み ===', ''],
        ['よくできる重み', defaultSettings['よくできる重み']],
        ['できる重み', defaultSettings['できる重み']],
        ['少しできる重み', defaultSettings['少しできる重み']],
        ['できていない重み', defaultSettings['できていない重み']],
        ['', ''],
        ['=== 評価ランク閾値(300点満点) ===', ''],
        ['優秀閾値', defaultSettings['優秀閾値']],
        ['良好閾値', defaultSettings['良好閾値']],
        ['要努力閾値', defaultSettings['要努力閾値']],
        ['', ''],
        ['=== 立場別評価重み(合計1.0) ===', ''],
        ['本人評価重み', defaultSettings['本人評価重み']],
        ['同僚評価重み', defaultSettings['同僚評価重み']],
        ['上司評価重み', defaultSettings['上司評価重み']],
        ['', ''],
        ['=== システム設定 ===', ''],
        ['ログイン機能有効', defaultSettings['ログイン機能有効']]
      ];
      settingsSheet.getRange(2, 1, settingsData.length, 2).setValues(settingsData);
      settingsSheet.setColumnWidth(1, 250);
      settingsSheet.setColumnWidth(2, 100);
      messages.push('「設定マスタ」シートを作成しました。');
    }

    if (!ss.getSheetByName(PERIOD_MASTER_SHEET_NAME)) {
      var periodSheet = ss.insertSheet(PERIOD_MASTER_SHEET_NAME);
      periodSheet.getRange('A1:E1').setValues([['期間ID', '期間名', '開始日', '終了日', 'ステータス']]).setFontWeight('bold');
      periodSheet.appendRow(['2024Q1', '2024年第1四半期 (サンプル)', '2024-01-01', '2024-03-31', 'active']);
      periodSheet.setColumnWidth(1, 120);
      periodSheet.setColumnWidth(2, 200);
      messages.push('「評価期間マスタ」シートを作成しました。');
    }

    if (!ss.getSheetByName(EVALUATOR_MASTER_SHEET_NAME)) {
      var newSheet = ss.insertSheet(EVALUATOR_MASTER_SHEET_NAME);
      newSheet.getRange('A1:B1').setValues([['評価者氏名', '上司フラグ']]).setFontWeight('bold');
      newSheet.getRange('A2:B4').setValues([
        ['山田 太郎 (例)', ''],
        [PRINCIPAL_NAME + ' (例)', '○'],
        ['佐藤 次郎 (例)', '○']
      ]);
      newSheet.setColumnWidth(1, 200).setColumnWidth(2, 100);
      messages.push('「評価者マスタ」シートを作成しました。');
    }

    if (!ss.getSheetByName(EVALUATED_MASTER_SHEET_NAME)) {
      var newSheet2 = ss.insertSheet(EVALUATED_MASTER_SHEET_NAME);
      newSheet2.getRange('A1').setValue('被評価者氏名').setFontWeight('bold');
      newSheet2.getRange('A2:A3').setValues([['山田 太郎 (例)'], ['鈴木 花子 (例)']]);
      messages.push('「被評価者マスタ」シートを作成しました。');
    }

    if (!ss.getSheetByName(ITEM_MASTER_SHEET_NAME)) {
      var newSheet3 = ss.insertSheet(ITEM_MASTER_SHEET_NAME);
      var itemHeaders = ['ID', '評価項目タイトル', 'グループ', 'チェック項目1', 'チェック項目2', 'チェック項目3', 'チェック項目4', 'チェック項目5'];
      newSheet3.getRange(1, 1, 1, itemHeaders.length).setValues([itemHeaders]).setFontWeight('bold');
      var exampleItems = [
        [1, 'スピード感', 'A', '急ぎの指示にすぐ動く','患者様を待たせない','次の準備ができている','忙しい時は速く動く','診療の流れを理解'],
        [2, '言葉遣い', 'A', '診療中の私語なし','敬語で話す','声の大きさ適切','プロらしい会話','患者様に聞かれてもOK'],
        [3, '整理整頓', 'A', '使ったら必ず戻す','診療台きれい','ゴミすぐ捨てる','共有場所も整理','次の人への配慮'],
        [4, '報告連絡相談', 'A', '変化を必ず報告','ミスを隠さない','分からない時は聞く','タイミング良い','重要度を判断'],
        [5, '医院貢献', 'A', '医院の為に考える','他の人を手伝う','残業も協力的','医院の評判大切に','マイナス発言をしない'],
        [6, '患者対応', 'B', '明るい挨拶','笑顔','丁寧な対応','気配り', ''],
        [7, '向上心', 'B', '勉強会参加','練習する','質問する','実践に活かす', ''],
        [8, '勤怠', 'B', '遅刻なし','欠勤なし','体調管理','予防接種', ''],
        [9, '協調性', 'B', '全員と仲良く','スタッフ間で挨拶をする','感謝を伝える','思いやり', ''],
        [10, '正確性', 'B', 'ミスが少ない','確認する','最後まで責任','確実な仕事', '']
      ];
      newSheet3.getRange(2, 1, exampleItems.length, exampleItems[0].length).setValues(exampleItems);
      messages.push('「評価項目マスタ」シートを作成しました。');
    }

    var activePeriod = getActivePeriod();
    if (activePeriod) {
      createPeriodSheets(ss, activePeriod.id);
      messages.push('評価期間「' + activePeriod.name + '」用のシートを作成しました。');
    }

    if (!ss.getSheetByName(TEMP_SAVE_SHEET_NAME)) {
      var newSheet4 = ss.insertSheet(TEMP_SAVE_SHEET_NAME);
      newSheet4.getRange('A1:C1').setValues([['評価者', '保存日時', 'データ']]).setFontWeight('bold');
      messages.push('「一時保存データ」シートを作成しました。');
    }

    if (!ss.getSheetByName(USER_AUTH_SHEET_NAME)) {
      var newSheet5 = ss.insertSheet(USER_AUTH_SHEET_NAME);
      newSheet5.getRange('A1:D1').setValues([['評価者名', 'メールアドレス（ログインID）', 'パスワード', '登録日時']]).setFontWeight('bold');
      messages.push('「ユーザー認証情報」シートを作成しました。');
    }

    if (messages.length > 0) {
      ui.alert('設定が完了しました。\n\n' + messages.join('\n'));
    } else {
      ui.alert('必要なシートはすべて準備済みです。');
    }
  } catch (e) {
    ui.alert('エラーが発生しました: ' + e.message);
  }
}

// =============================================================
// データ取得ヘルパー
// =============================================================

function getInitialData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var settings = loadSettings();
    var periods = getPeriods();
    var activePeriod = getActivePeriod();
    var evaluators = getListFromSheet(ss, EVALUATOR_MASTER_SHEET_NAME);
    var evaluated = getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME);
    var items = getEvaluationItemsFromSheet(ss);
    return {
      evaluators: evaluators,
      evaluated: evaluated,
      items: items,
      settings: settings,
      periods: periods,
      activePeriod: activePeriod
    };
  } catch (e) {
    console.error('getInitialData エラー:', e.message);
    return {
      evaluators: [],
      evaluated: [],
      items: [],
      settings: getDefaultSettings(),
      periods: [],
      activePeriod: null,
      error: e.message
    };
  }
}

function getListFromSheet(ss, sheetName) {
  try {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var range = sheet.getRange('A2:A' + lastRow);
    return range.getValues().map(function(row) { return row[0]; }).filter(function(name) { return name; });
  } catch (e) {
    return [];
  }
}

function getEvaluatorDetails(ss) {
  try {
    var sheet = ss.getSheetByName(EVALUATOR_MASTER_SHEET_NAME);
    if (!sheet) return {};
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return {};
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    var evaluatorDetails = {};
    data.forEach(function(row) {
      var name = row[0];
      var isSupervisor = row[1] && (row[1].toString().trim() === '○' || row[1].toString().trim() === '上司' || row[1].toString().trim() === 'TRUE');
      if (name) {
        evaluatorDetails[name] = { isSupervisor: isSupervisor };
      }
    });
    return evaluatorDetails;
  } catch (e) {
    return {};
  }
}

function getEvaluationItemsFromSheet(ss) {
  try {
    var sheet = ss.getSheetByName(ITEM_MASTER_SHEET_NAME);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    var data = sheet.getRange('A2:Z' + lastRow).getValues();
    return data.map(function(row) {
      var id = row[0];
      var title = row[1];
      var group = row[2];
      var checks = row.slice(3).filter(function(item) { return item; });
      if (id && title && group && checks.length > 0) {
        return { id: id, title: title, group: group, checks: checks };
      }
      return null;
    }).filter(function(item) { return item; });
  } catch (e) {
    return [];
  }
}

// =============================================================
// システム診断
// =============================================================

function diagnoseSystem() {
  var results = [];
  try {
    results.push('=== システム診断 ===');
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    results.push('スプレッドシート接続成功: ' + ss.getName());
    var requiredSheets = [
      SETTINGS_SHEET_NAME, PERIOD_MASTER_SHEET_NAME,
      EVALUATOR_MASTER_SHEET_NAME, EVALUATED_MASTER_SHEET_NAME,
      ITEM_MASTER_SHEET_NAME, TEMP_SAVE_SHEET_NAME, USER_AUTH_SHEET_NAME
    ];
    requiredSheets.forEach(function(sheetName) {
      var sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        results.push('OK シート存在: ' + sheetName + ' (行数: ' + sheet.getLastRow() + ')');
      } else {
        results.push('NG シート不在: ' + sheetName);
      }
    });
    var periods = getPeriods();
    results.push('評価期間数: ' + periods.length);
    var activePeriod = getActivePeriod();
    if (activePeriod) {
      results.push('アクティブ期間: ' + activePeriod.name);
    } else {
      results.push('アクティブ期間なし');
    }
    var evaluators = getListFromSheet(ss, EVALUATOR_MASTER_SHEET_NAME);
    results.push('評価者数: ' + evaluators.length);
    var evaluated = getListFromSheet(ss, EVALUATED_MASTER_SHEET_NAME);
    results.push('被評価者数: ' + evaluated.length);
    var items = getEvaluationItemsFromSheet(ss);
    results.push('評価項目数: ' + items.length);
    results.push('=== 診断完了 ===');
  } catch (e) {
    results.push('診断エラー: ' + e.message);
  }
  var message = results.join('\n');
  try { SpreadsheetApp.getUi().alert(message); } catch (e) {}
  return message;
}

// =============================================================
// Web App エントリポイント
// =============================================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('360° スタッフ評価システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
