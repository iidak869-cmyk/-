/**
 * customform.jp のアンケート回答を自動でGoogleスプレッドシートに反映するスクリプト。
 *
 * 事前準備:
 * 1. スクリプトプロパティに以下を設定する（プロジェクトの設定 > スクリプト プロパティ）
 *      CUSTOMFORM_ACCOUNT    customform.jp のログインID
 *      CUSTOMFORM_PASSPHRASE customform.jp のログインパスワード
 * 2. setup() を一度だけ実行する（新しいスプレッドシートの作成、定期実行トリガーの登録、初回同期を行う）
 *
 * 以降は syncAll() が定期実行トリガーにより自動で回答を反映する。
 */

var FORMS = [
  { name: 'Cytekiサポート中間アンケート', id: 241254, sheetName: 'Cyteki中間_回答一覧' },
  { name: 'Cytekiサポート最終アンケート', id: 237126, sheetName: 'Cyteki最終_回答一覧' },
  { name: '納品後サポートアンケート', id: 231931, sheetName: '納品後サポート_回答一覧' }
];

var BASE_URL = 'https://customform.jp';

function setup() {
  var props = PropertiesService.getScriptProperties();
  var spreadsheetId = props.getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    var ss = SpreadsheetApp.create('アンケート回答自動反映');
    spreadsheetId = ss.getId();
    props.setProperty('SPREADSHEET_ID', spreadsheetId);
    Logger.log('スプレッドシートを新規作成しました: ' + ss.getUrl());
  } else {
    Logger.log('既存のスプレッドシートを使用します: ' + SpreadsheetApp.openById(spreadsheetId).getUrl());
  }

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t);
  });
  var weekdays = [
    ScriptApp.WeekDay.MONDAY,
    ScriptApp.WeekDay.TUESDAY,
    ScriptApp.WeekDay.WEDNESDAY,
    ScriptApp.WeekDay.THURSDAY,
    ScriptApp.WeekDay.FRIDAY
  ];
  weekdays.forEach(function (day) {
    ScriptApp.newTrigger('syncAll').timeBased().everyWeeks(1).onWeekDay(day).atHour(19).create();
  });
  Logger.log('平日(月〜金)19時の自動実行トリガーを登録しました。');

  syncAll();
}

function syncAll() {
  var props = PropertiesService.getScriptProperties();
  var account = props.getProperty('CUSTOMFORM_ACCOUNT');
  var passphrase = props.getProperty('CUSTOMFORM_PASSPHRASE');
  var spreadsheetId = props.getProperty('SPREADSHEET_ID');

  if (!account || !passphrase) {
    throw new Error('スクリプトプロパティに CUSTOMFORM_ACCOUNT / CUSTOMFORM_PASSPHRASE を設定してください。');
  }
  if (!spreadsheetId) {
    throw new Error('先に setup() を実行してスプレッドシートを作成してください。');
  }

  var auth = customformLogin(account, passphrase);
  var deleteAfterSync = props.getProperty('DELETE_AFTER_SYNC') === 'true';
  var ss = SpreadsheetApp.openById(spreadsheetId);

  FORMS.forEach(function (formConf) {
    try {
      var addedCount = syncFormToSheet(ss, formConf, auth, deleteAfterSync);
      Logger.log(formConf.name + ': ' + addedCount + ' 件追加');
    } catch (e) {
      Logger.log('【エラー】' + formConf.name + ': ' + e);
    }
  });
}

function customformLogin(account, passphrase) {
  var pre = UrlFetchApp.fetch(BASE_URL + '/signin', { muteHttpExceptions: true });
  var preCookies = getCookiesFromResponse(pre);
  var xsrfToken = decodeURIComponent(preCookies['XSRF-TOKEN']);

  var loginRes = UrlFetchApp.fetch(BASE_URL + '/api/signin', {
    method: 'post',
    payload: { account: account, passphrase: passphrase },
    headers: {
      Cookie: buildCookieHeader(preCookies),
      'X-XSRF-TOKEN': xsrfToken,
      Accept: 'application/json',
      'X-Requested-With': 'XMLHttpRequest'
    },
    muteHttpExceptions: true
  });

  var result;
  try {
    result = JSON.parse(loginRes.getContentText());
  } catch (e) {
    throw new Error('customform.jp へのログインに失敗しました（想定外のレスポンス）: ' + loginRes.getContentText().substring(0, 300));
  }
  if (!result.result || !result.result.login_check) {
    throw new Error('customform.jp へのログインに失敗しました。IDまたはパスワードを確認してください。');
  }

  var loginCookies = getCookiesFromResponse(loginRes);
  var mergedCookies = Object.assign({}, preCookies, loginCookies);
  return {
    cookieHeader: buildCookieHeader(mergedCookies),
    xsrfToken: decodeURIComponent(mergedCookies['XSRF-TOKEN'])
  };
}

var MAX_DETAIL_FETCH_PER_RUN = 200;

var TRAILING_TITLE_MARKER = 'ホームページ';

var MARK = {
  DATE: '__DATE__',
  COMPANY: '__COMPANY__',
  PERSON: '__PERSON__',
  TOTAL: '__TOTAL__',
  AVERAGE: '__AVERAGE__',
  ANSWER_ID: '__ANSWER_ID__'
};

var RENAME_MAP = {
  241254: { // Cytekiサポート中間アンケート
    2566947: '①担当者\n満足度',
    2620889: '②特によかったサポート内容',
    2566948: '③現在の不安感',
    2566954: '④導入して変わりそうと感じている部分',
    2566950: '⑤業務への\n役立ち',
    2566952: '⑥今後のサポートで知りたいこと',
    2620890: 'HPで紹介',
    2641012: 'HPで紹介'
  },
  237126: { // Cytekiサポート最終アンケート
    2514815: '①担当者\n　 満足度',
    2514846: '②理解',
    2620891: '③②の理由',
    2514858: '④サポート全体\nの満足度',
    2620892: '⑤④の理由',
    2514892: '⑥特に良かった点',
    2620897: '⑦他社比較',
    2620894: 'HPで紹介'
  },
  231931: { // 納品後サポートアンケート
    2455228: '①満足度',
    2455230: '②活用',
    2455232: '③記述',
    2455233: '④紹介',
    2455234: '⑤記述',
    2455235: 'その他'
  }
};

function getDisplayTitle(formId, questionId, originalTitle) {
  var map = RENAME_MAP[formId];
  if (map && map.hasOwnProperty(questionId)) return map[questionId];
  return originalTitle;
}

function syncFormToSheet(ss, formConf, auth, deleteAfterSync) {
  var cookieHeader = auth.cookieHeader;
  var headings = fetchQuestionHeadings(formConf.id, cookieHeader);
  if (headings.length < 2) return 0;

  var nameHeading = headings[0];
  var personHeading = headings[1];
  var restHeadings = headings.slice(2);

  var normalHeadings = restHeadings.filter(function (q) { return q.title.indexOf(TRAILING_TITLE_MARKER) === -1; });
  var trailingHeadings = restHeadings.filter(function (q) { return q.title.indexOf(TRAILING_TITLE_MARKER) !== -1; });

  var normalEntries = uniqueHeadingEntries(normalHeadings, formConf.id);
  var trailingEntries = uniqueHeadingEntries(trailingHeadings, formConf.id);
  var normalMarkers = normalEntries.map(function (e) { return e.marker; });
  var normalDisplays = normalEntries.map(function (e) { return e.display; });
  var trailingMarkers = trailingEntries.map(function (e) { return e.marker; });
  var trailingDisplays = trailingEntries.map(function (e) { return e.display; });

  var numericQuestionIds = restHeadings
    .filter(function (q) { return q.input_type === 1; })
    .map(function (q) { return q.question_id; });

  var answerMetas = fetchAllAnswerMetas(formConf.id, cookieHeader);
  if (answerMetas.length === 0) return 0;

  var sheet = ss.getSheetByName(formConf.sheetName);
  if (!sheet) sheet = ss.insertSheet(formConf.sheetName);

  var headerRow;
  var markerRow;
  if (sheet.getLastRow() === 0) {
    headerRow = ['日付', '社名', '担当者']
      .concat(normalDisplays)
      .concat(['合計点', '平均'])
      .concat(trailingDisplays)
      .concat(['回答ID']);
    markerRow = [MARK.DATE, MARK.COMPANY, MARK.PERSON]
      .concat(normalMarkers)
      .concat([MARK.TOTAL, MARK.AVERAGE])
      .concat(trailingMarkers)
      .concat([MARK.ANSWER_ID]);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
    setHeaderNotes(sheet, markerRow);
    sheet.setFrozenRows(1);
  } else {
    var lastCol = sheet.getLastColumn();
    headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    markerRow = sheet.getRange(1, 1, 1, lastCol).getNotes()[0];

    // ノート機能の導入前に作られた列（ノートが空）を、表示テキストから推測して補完する。
    var allEntries = normalEntries.concat(trailingEntries);
    for (var i = 0; i < headerRow.length; i++) {
      if (markerRow[i]) continue;
      var text = headerRow[i];
      if (text === '日付') markerRow[i] = MARK.DATE;
      else if (text === '社名') markerRow[i] = MARK.COMPANY;
      else if (text === '担当者') markerRow[i] = MARK.PERSON;
      else if (text === '合計点') markerRow[i] = MARK.TOTAL;
      else if (text === '平均') markerRow[i] = MARK.AVERAGE;
      else if (text === '回答ID') markerRow[i] = MARK.ANSWER_ID;
      else {
        var match = allEntries.filter(function (e) { return e.display === text; })[0];
        if (match) markerRow[i] = match.marker;
      }
    }

    var normalInsertPos = markerRow.indexOf(MARK.TOTAL);
    if (normalInsertPos === -1) normalInsertPos = headerRow.length;
    normalEntries.forEach(function (e) {
      if (markerRow.indexOf(e.marker) === -1) {
        headerRow.splice(normalInsertPos, 0, e.display);
        markerRow.splice(normalInsertPos, 0, e.marker);
        normalInsertPos++;
      }
    });
    var trailingInsertPos = markerRow.indexOf(MARK.ANSWER_ID);
    if (trailingInsertPos === -1) trailingInsertPos = headerRow.length;
    trailingEntries.forEach(function (e) {
      if (markerRow.indexOf(e.marker) === -1) {
        headerRow.splice(trailingInsertPos, 0, e.display);
        markerRow.splice(trailingInsertPos, 0, e.marker);
        trailingInsertPos++;
      }
    });
    if (markerRow.indexOf(MARK.ANSWER_ID) === -1) {
      headerRow.push('回答ID');
      markerRow.push(MARK.ANSWER_ID);
    }
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    setHeaderNotes(sheet, markerRow);
  }

  var idColIndex = markerRow.indexOf(MARK.ANSWER_ID);
  sheet.hideColumns(idColIndex + 1);

  var existingIds = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues().forEach(function (r) {
      existingIds[String(r[0])] = true;
    });
  }

  var targets = answerMetas.filter(function (m) {
    return !m.del_flg && !existingIds[String(m.answer_id)];
  });
  targets.sort(function (a, b) { return new Date(a.create_time) - new Date(b.create_time); });

  var limited = targets.slice(0, MAX_DETAIL_FETCH_PER_RUN);
  var newRows = [];
  var syncedMetas = [];
  limited.forEach(function (meta) {
    var detail = fetchAnswerDetail(formConf.id, meta.answer_id, cookieHeader);
    if (!detail) return;

    var qTitleToValue = {};
    restHeadings.forEach(function (q) {
      var vals = detail.answers[String(q.question_id)];
      var v = vals ? vals.join('、') : '';
      if (v || !qTitleToValue.hasOwnProperty(q.title)) qTitleToValue[q.title] = v;
    });

    var sum = 0;
    var count = 0;
    numericQuestionIds.forEach(function (qid) {
      var vals = detail.answers[String(qid)];
      var n = vals && vals[0] !== '' ? parseFloat(vals[0]) : NaN;
      if (!isNaN(n)) {
        sum += n;
        count++;
      }
    });

    var nameVal = (detail.answers[String(nameHeading.question_id)] || []).join('、');
    var personVal = (detail.answers[String(personHeading.question_id)] || []).join('、');
    var createDate = new Date(meta.create_time.replace(' ', 'T'));

    var row = markerRow.map(function (marker) {
      if (marker === MARK.DATE) return createDate;
      if (marker === MARK.COMPANY) return nameVal;
      if (marker === MARK.PERSON) return personVal;
      if (marker === MARK.TOTAL) return count > 0 ? sum : '';
      if (marker === MARK.AVERAGE) return count > 0 ? sum / count : '';
      if (marker === MARK.ANSWER_ID) return meta.answer_id;
      return qTitleToValue.hasOwnProperty(marker) ? qTitleToValue[marker] : '';
    });
    newRows.push(row);
    syncedMetas.push(meta);
  });

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headerRow.length).setValues(newRows);
  }

  if (targets.length > limited.length) {
    Logger.log(formConf.name + ': 未取得の回答が残り ' + (targets.length - limited.length) + ' 件あります。次回の自動実行で続きを取り込みます。');
  }

  if (deleteAfterSync) {
    syncedMetas.forEach(function (meta) {
      try {
        deleteAnswerFromCustomform(formConf.id, meta.answer_id, auth);
      } catch (e) {
        Logger.log('【削除エラー】' + formConf.name + ' answer_id=' + meta.answer_id + ': ' + e);
      }
    });
  }

  return newRows.length;
}

function setHeaderNotes(sheet, markerRow) {
  var current = sheet.getRange(1, 1, 1, markerRow.length).getNotes()[0];
  markerRow.forEach(function (marker, i) {
    if (current[i] !== marker) {
      sheet.getRange(1, i + 1).setNote(marker);
    }
  });
}

function uniqueHeadingEntries(headings, formId) {
  var entries = [];
  var seen = {};
  headings.forEach(function (q) {
    if (!seen[q.title]) {
      seen[q.title] = true;
      entries.push({ marker: q.title, display: getDisplayTitle(formId, q.question_id, q.title) });
    }
  });
  return entries;
}

function fetchQuestionHeadings(formId, cookieHeader) {
  var res = UrlFetchApp.fetch(BASE_URL + '/api/manage/form/question?customform_id=' + formId, {
    headers: { Cookie: cookieHeader, Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return [];
  var json = JSON.parse(res.getContentText());
  if (json.status !== 'OK') return [];
  return json.result.slice().sort(function (a, b) { return a.order_id - b.order_id; });
}

function deleteAnswerFromCustomform(formId, answerId, auth) {
  var res = UrlFetchApp.fetch(BASE_URL + '/api/manage/form/answer/remove', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ customform_id: formId, answer_id: answerId }),
    headers: {
      Cookie: auth.cookieHeader,
      'X-XSRF-TOKEN': auth.xsrfToken,
      Accept: 'application/json',
      'X-Requested-With': 'XMLHttpRequest'
    },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error('削除リクエストが失敗しました（status=' + res.getResponseCode() + '）');
  }
}

function fetchAnswerDetail(formId, answerId, cookieHeader) {
  var url = BASE_URL + '/api/manage/form/answer/info/' + answerId + '?customform_id=' + formId;
  var res = UrlFetchApp.fetch(url, {
    headers: { Cookie: cookieHeader, Accept: 'application/json', 'X-Requested-With': 'XMLHttpRequest' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return null;
  var json = JSON.parse(res.getContentText());
  if (json.status !== 'OK') return null;
  return json.result;
}

function fetchAllAnswerMetas(formId, cookieHeader) {
  var all = [];
  var page = 1;
  var MAX_PAGES = 500;

  while (page <= MAX_PAGES) {
    var url = BASE_URL + '/manage/form/' + formId + '/answer?page=' + page;
    var res = UrlFetchApp.fetch(url, {
      headers: { Cookie: cookieHeader },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) break;

    var metas = extractAnswerMetas(res.getContentText());
    if (metas.length === 0) break;

    all = all.concat(metas);
    page++;
  }

  return all;
}

function extractAnswerMetas(html) {
  var tagRe = /<answer-info-component\b([^>]*)>/g;
  var results = [];
  var tagMatch;
  while ((tagMatch = tagRe.exec(html)) !== null) {
    var raw = extractAttr(tagMatch[1], 'answer');
    if (!raw) continue;
    var answer = JSON.parse(htmlDecode(raw));
    results.push({ answer_id: answer.answer_id, create_time: answer.create_time, del_flg: answer.del_flg });
  }
  return results;
}

function extractAttr(attrString, attrName) {
  var re = new RegExp(':' + attrName + '="([^"]*)"');
  var m = re.exec(attrString);
  return m ? m[1] : null;
}

function htmlDecode(str) {
  return str
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function getCookiesFromResponse(res) {
  var headers = res.getAllHeaders();
  var setCookies = headers['Set-Cookie'] || headers['set-cookie'] || [];
  if (typeof setCookies === 'string') setCookies = [setCookies];
  var cookies = {};
  setCookies.forEach(function (c) {
    var pair = c.split(';')[0];
    var idx = pair.indexOf('=');
    cookies[pair.substring(0, idx)] = pair.substring(idx + 1);
  });
  return cookies;
}

function buildCookieHeader(cookieObj) {
  return Object.keys(cookieObj)
    .map(function (k) { return k + '=' + cookieObj[k]; })
    .join('; ');
}
