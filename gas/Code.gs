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
  ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();
  Logger.log('1時間ごとの自動実行トリガーを登録しました。');

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

/**
 * 指定フォームの回答一覧ページ（全ページ分）を取得し、スプレッドシートに未反映の回答だけ追記する。
 * 戻り値は追加した件数。
 */
function syncFormToSheet(ss, formConf, cookieHeader) {
  var comps = fetchAllAnswerComponents(formConf.id, cookieHeader);
  if (comps.length === 0) return 0;

  var headings = JSON.parse(htmlDecode(comps[0].question_headings));
  headings.sort(function (a, b) { return a.order_id - b.order_id; });

  var sheet = ss.getSheetByName(formConf.sheetName);
  if (!sheet) sheet = ss.insertSheet(formConf.sheetName);

  var headerRow;
  if (sheet.getLastRow() === 0) {
    headerRow = ['回答ID', '回答日時'].concat(headings.map(function (q) { return q.title; }));
    sheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // 新しい質問が増えていた場合は列を追加する
    headings.forEach(function (q) {
      if (headerRow.indexOf(q.title) === -1) {
        headerRow.push(q.title);
        sheet.getRange(1, headerRow.length).setValue(q.title).setFontWeight('bold');
      }
    });
  }

  var existingIds = {};
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(function (r) {
      existingIds[String(r[0])] = true;
    });
  }

  var newRows = [];
  comps.forEach(function (comp) {
    var answer = JSON.parse(htmlDecode(comp.answer));
    if (answer.del_flg) return; // 削除済みの回答は無視
    if (existingIds[String(answer.answer_id)]) return; // 既に反映済み

    var options = JSON.parse(htmlDecode(comp.answer_options));
    var qTitleToValue = {};
    headings.forEach(function (q) {
      var vals = options[String(q.question_id)];
      qTitleToValue[q.title] = vals ? vals.join('、') : '';
    });

    var createDate = new Date(answer.create_time.replace(' ', 'T'));
    var row = headerRow.map(function (col) {
      if (col === '回答ID') return answer.answer_id;
      if (col === '回答日時') return createDate;
      return qTitleToValue.hasOwnProperty(col) ? qTitleToValue[col] : '';
    });

    newRows.push({ time: createDate.getTime(), row: row });
  });

  newRows.sort(function (a, b) { return a.time - b.time; });

  if (newRows.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, headerRow.length)
      .setValues(newRows.map(function (r) { return r.row; }));
  }

  return newRows.length;
}

/**
 * 指定フォームの回答一覧を全ページ分取得し、各回答の <answer-info-component> の属性一覧を返す。
 */
function fetchAllAnswerComponents(formId, cookieHeader) {
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

    var comps = extractAnswerComponents(res.getContentText());
    if (comps.length === 0) break;

    all = all.concat(comps);
    page++;
  }

  return all;
}

/**
 * HTML中の <answer-info-component :answer="..." :answer_options="..." :question_headings="..."> を
 * すべて抜き出し、各属性の生文字列（HTMLエンティティのまま）を配列で返す。
 */
function extractAnswerComponents(html) {
  var tagRe = /<answer-info-component\b([^>]*)>/g;
  var results = [];
  var tagMatch;
  while ((tagMatch = tagRe.exec(html)) !== null) {
    var attrs = tagMatch[1];
    var answer = extractAttr(attrs, 'answer');
    var answerOptions = extractAttr(attrs, 'answer_options');
    var questionHeadings = extractAttr(attrs, 'question_headings');
    if (answer && answerOptions && questionHeadings) {
      results.push({ answer: answer, answer_options: answerOptions, question_headings: questionHeadings });
    }
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
