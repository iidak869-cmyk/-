/**
 * 自動化タスクの完了内容を記録するWebhook。
 * POSTされた { secret, task, summary, link } を検証のうえ、
 * Googleスプレッドシート「自動化 対応履歴」に1行追記する。
 *
 * 事前準備:
 * 1. スクリプトプロパティに LOG_SECRET を設定する（推測されにくい適当な文字列）
 *      プロジェクトの設定 > スクリプト プロパティ
 * 2. エディタ上部「デプロイ」>「新しいデプロイ」> 種類:ウェブアプリ
 *      実行するユーザー: 自分
 *      アクセスできるユーザー: 全員
 *    でデプロイし、発行されたウェブアプリのURLを控える
 * 3. 発行されたURLと LOG_SECRET を Claude Code 環境の環境変数
 *    LOG_WEBHOOK_URL / LOG_WEBHOOK_SECRET に設定する
 */
function doPost(e) {
  var props = PropertiesService.getScriptProperties();
  var secret = props.getProperty('LOG_SECRET');

  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput({ ok: false, error: 'invalid_json' });
  }

  if (!secret || body.secret !== secret) {
    return jsonOutput({ ok: false, error: 'unauthorized' });
  }

  var spreadsheetId = props.getProperty('SPREADSHEET_ID');
  var ss;
  if (!spreadsheetId) {
    ss = SpreadsheetApp.create('自動化 対応履歴');
    spreadsheetId = ss.getId();
    props.setProperty('SPREADSHEET_ID', spreadsheetId);
  } else {
    ss = SpreadsheetApp.openById(spreadsheetId);
  }

  var sheet = ss.getSheetByName('履歴');
  if (!sheet) {
    sheet = ss.insertSheet('履歴');
    sheet.appendRow(['完了日時', 'タスク名', '対応内容', '関連リンク']);
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    body.task || '',
    body.summary || '',
    body.link || ''
  ]);

  return jsonOutput({ ok: true, url: ss.getUrl() });
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
