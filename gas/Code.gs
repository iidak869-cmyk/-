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

// 同期対象のフォーム一覧（customform.jp の「フォーム一覧」ページから取得したID）
var FORMS = [
  { name: 'Cytekiサポート中間アンケート', id: 241254, sheetName: 'Cyteki中間_回答一覧' },
  { name: 'Cytekiサポート最終アンケート', id: 237126, sheetName: 'Cyteki最終_回答一覧' },
  { name: '納品後サポートアンケート', id: 231931, sheetName: '納品後サポート_回答一覧' },
  { name: '納品/公開前アンケート', id: 170650, sheetName: '納品公開前_回答一覧' }
];

var BASE_URL = 'https://customform.jp';

/**
 * 初回セットアップ用。新しいスプレッドシートを作成し、定期実行トリガーを登録して初回同期を行う。
 * 2回目以降実行すると、スプレッドシートは作り直さず、トリガーだけ再登録する。
 */
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
  ScriptApp.newTrigger('syncAll').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  Logger.log('毎週月曜6時の自動実行トリガーを登録しました。');

  syncAll();
}

/**
 * 全フォームの回答を同期するメイン処理。定期実行トリガーから呼ばれる。
 */
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

  var cookieHeader = customformLogin(account, passphrase);
  var ss = SpreadsheetApp.openById(spreadsheetId);

  FORMS.forEach(function (formConf) {
    try {
      var addedCount = syncFormToSheet(ss, formConf, cookieHeader);
      Logger.log(formConf.name + ': ' + addedCount + ' 件追加');
    } catch (e) {
      Logger.log('【エラー】' + formConf.name + ': ' + e);
    }
  });
}

/**
 * customform.jp にログインし、以降のAPI/ページ取得に使うCookieヘッダー文字列を返す。
 */
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
  return buildCookieHeader(mergedCookies);
}

// 1回の実行で新規に詳細取得する回答の最大件数（実行時間の上限対策。超えた分は次回実行で続きを取得する）
var MAX_DETAIL_FETCH_PER_RUN = 200;

/**
 * 指定フォームの回答一覧ページ（全ページ分）を取得し、スプレッドシートに未反映の回答だけ追記する。
 * 戻り値は追加した件数。
 *
 * 回答一覧ページのHTMLに埋め込まれる <answer-info-component> は「社名・担当者名」など
 * ごく一部の質問しか含まないため、回答ID・日時の一覧だけをそこから取得し、
 * 質問文は /api/manage/form/question、各回答の全項目は /api/manage/form/answer/info/<ID>
 * から個別に取得する。
 *
 * シートの列は「日付・社名・担当者・（各質問）・合計点・平均・回答ID」の順。
 * 先頭2つの質問（★御社名・★担当者名）は社名／担当者列に割り当て、
 * ラジオボタン形式（input_type=1、10点満点評価など）の質問だけを合計点・平均の対象にする。
 * 同じ質問文が複数の質問IDに重複登録されている場合は1列にまとめる。
 */
function syncFormToSheet(ss, formConf, cookieHeader) {
  var headings = fetchQuestionHeadings(formConf.id, cookieHeader);
  if (headings.length < 2) return 0;

  var nameHeading = headings[0];
  var personHeading = headings[1];
  var restHeadings = headings.slice(2); // ★御社名・★担当者名以外の質問（重複タイトルを含む場合あり）

  var restTitles = []; // 表示順を保った重複なしの質問文一覧
  var seenTitles = {};
  restHeadings.forEach(function (q) {
    if (!seenTitles[q.title]) {
      seenTitles[q.title] = true;
      restTitles.push(q.title);
    }
  });
  var numericQuestionIds = restHeadings
    .filter(function (q) { return q.input_type === 1; })
    .map(function (q) { return q.question_id; });

  var answerMetas = fetchAllAnswerMetas(formConf.id, cookieHeader);
  if (answerMetas.length === 0) return 0;

  var sheet = ss.getSheetByName(formConf.sheetName);
  if (!sheet) sheet = ss.insertSheet(formConf.sheetName);

  var headerRow;
  if (sheet.getLastRow() === 0) {
    headerRow = ['日付', '社名', '担当者'].concat(restTitles).concat(['合計点', '平均', '回答ID']);
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // 新しい質問が増えていた場合は「回答ID」の手前に列を追加する
    var insertPos = headerRow.indexOf('回答ID');
    if (insertPos === -1) insertPos = headerRow.length;
    restTitles.forEach(function (title) {
      if (headerRow.indexOf(title) === -1) {
        headerRow.splice(insertPos, 0, title);
        insertPos++;
      }
    });
    if (headerRow.indexOf('回答ID') === -1) headerRow.push('回答ID');
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
  }

  var idColIndex = headerRow.indexOf('回答ID');
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

    var row = headerRow.map(function (col) {
      if (col === '日付') return createDate;
      if (col === '社名') return nameVal;
      if (col === '担当者') return personVal;
      if (col === '合計点') return count > 0 ? sum : '';
      if (col === '平均') return count > 0 ? sum / count : '';
      if (col === '回答ID') return meta.answer_id;
      return qTitleToValue.hasOwnProperty(col) ? qTitleToValue[col] : '';
    });
    newRows.push(row);
  });

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headerRow.length).setValues(newRows);
  }

  if (targets.length > limited.length) {
    Logger.log(formConf.name + ': 未取得の回答が残り ' + (targets.length - limited.length) + ' 件あります。次回の自動実行で続きを取り込みます。');
  }

  return newRows.length;
}

/**
 * フォームの質問一覧（質問ID・タイトル・表示順）を取得する。
 */
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

/**
 * 指定した回答の全質問への回答内容を取得する。{ answers: { question_id: [値, ...] } } を返す。
 */
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

/**
 * 指定フォームの回答一覧を全ページ分取得し、各回答の {回答ID, 回答日時, 削除フラグ} の一覧を返す。
 */
function fetchAllAnswerMetas(formId, cookieHeader) {
  var all = [];
  var page = 1;
  var MAX_PAGES = 500; // 安全のための上限

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

/**
 * HTML中の <answer-info-component :answer="..."> から回答ID・回答日時・削除フラグだけを抜き出す。
 */
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
