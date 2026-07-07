# アンケートフォーム自動反映（customform.jp → Googleスプレッドシート）

customform.jp で回答を受け付けている以下の4つのアンケートフォームの回答を、
Google Apps Script (GAS) で定期的に取得し、新規作成したGoogleスプレッドシートに自動反映します。

- Cytekiサポート中間アンケート
- Cytekiサポート最終アンケート
- 納品後サポートアンケート
- 納品/公開前アンケート

customform.jp にはWebhookやCSV/API連携機能が無いため、GASが customform.jp にログインして
回答一覧ページを取得し、解析して反映する方式です。

## セットアップ手順

1. https://script.google.com で新しいプロジェクトを作成する
2. `gas/Code.gs` の内容をコピーして、プロジェクトのコードとして貼り付ける
3. 左側の「プロジェクトの設定」(歯車アイコン) → 「スクリプト プロパティ」で以下を追加する
   - `CUSTOMFORM_ACCOUNT` : customform.jp のログインID
   - `CUSTOMFORM_PASSPHRASE` : customform.jp のログインパスワード
4. エディタ上部の関数選択で `setup` を選び、実行する
   - 初回実行時はGoogleアカウントへの権限承認が必要です
   - 実行すると新しいスプレッドシート「アンケート回答自動反映」が作成され、1時間ごとの自動実行トリガーが登録され、初回の同期が実行されます
5. 実行完了後、実行ログ（表示 > 実行数）に表示されるスプレッドシートのURLを開いて確認する

以降は1時間ごとに `syncAll` が自動実行され、新しい回答だけがスプレッドシートに追記されます。

## 同期の仕組み

- 各フォームごとに1つのシート（例: `Cyteki中間_回答一覧`）が作成されます
- 1行目は質問文をそのまま列見出しとして使用します（`回答ID` と `回答日時` が先頭2列）
- 既に取り込んだ回答（`回答ID`列で判定）はスキップされ、新しい回答だけが末尾に追記されます
- 質問がフォーム側で追加された場合、既存シートの右側に新しい列が自動追加されます
- 削除済みの回答は取り込みません

## 同期間隔の変更

`gas/Code.gs` の `setup()` 内、`ScriptApp.newTrigger('syncAll').timeBased().everyHours(1).create();`
の `everyHours(1)` を `everyMinutes(30)` などに変更し、`setup()` を再実行するとトリガーが更新されます。

## 対象フォームの追加・変更

`gas/Code.gs` 先頭の `FORMS` 配列に `{ name, id, sheetName }` を追加/削除してください。
`id` は customform.jp の管理画面でフォームを開いたときのURL（`/manage/form/<id>`）に含まれる数字です。
