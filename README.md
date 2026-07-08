# アンケートフォーム自動反映（customform.jp → Googleスプレッドシート）

customform.jp で回答を受け付けている以下の3つのアンケートフォームの回答を、
Google Apps Script (GAS) で定期的に取得し、新規作成したGoogleスプレッドシートに自動反映します。

- Cytekiサポート中間アンケート
- Cytekiサポート最終アンケート
- 納品後サポートアンケート

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
   - 実行すると新しいスプレッドシート「アンケート回答自動反映」が作成され、毎週月曜6時の自動実行トリガーが登録され、初回の同期が実行されます
5. 実行完了後、実行ログ（表示 > 実行数）に表示されるスプレッドシートのURLを開いて確認する

以降は毎週月曜6時に `syncAll` が自動実行され、新しい回答だけがスプレッドシートに追記されます。

## 同期の仕組み

- 各フォームごとに1つのシート（例: `Cyteki中間_回答一覧`）が作成されます
- 列は「日付・社名・担当者・（各質問）・合計点・平均」の順（`合計点`/`平均`はラジオボタン形式の質問だけを対象に自動計算）
- HPやSNSでの紹介可否など「ホームページ」という文言を含む質問は、合計点・平均より後ろの末尾に配置されます
- 重複チェック用の`回答ID`列を末尾に持ちますが非表示になっています
- 既に取り込んだ回答（`回答ID`列で判定）はスキップされ、新しい回答だけが末尾に追記されます
- 質問がフォーム側で増えた場合、既存シートに新しい列が自動追加されます
- 削除済みの回答は取り込みません

## customform.jp側の回答を反映後に削除する

スクリプトプロパティに `DELETE_AFTER_SYNC` を `true` にすると、スプレッドシートへの反映に成功した回答を
customform.jp側でも削除します（「削除済み回答一覧」に移動するだけで、そこから復元可能です）。
未設定または `true` 以外の値の場合は削除しません（デフォルトは削除しない安全側の動作）。

## 同期間隔の変更

`gas/Code.gs` の `setup()` 内、
`ScriptApp.newTrigger('syncAll').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();`
の曜日・時間・頻度を変更し、`setup()` を再実行するとトリガーが更新されます。

## 対象フォームの追加・変更

`gas/Code.gs` 先頭の `FORMS` 配列に `{ name, id, sheetName }` を追加/削除してください。
`id` は customform.jp の管理画面でフォームを開いたときのURL（`/manage/form/<id>`）に含まれる数字です。

## 対応履歴ログ（completion-log.csv）

Claudeが自動化タスクに対応完了するたびに、`completion-log.csv` に1行追記してpushする。
列は「完了日時・タスク名・対応内容・関連リンク」の順。Excelやスプレッドシートソフトでそのまま開ける。

### gas-log/ について（現在未使用）

当初はGoogleスプレッドシートに直接記録するための独立したGAS Webhook（`gas-log/Code.gs`）を
用意したが、Google Workspace（weeare.co.jp）側の設定により匿名でのウェブアプリアクセスが
ブロックされ、外部からPOSTできなかったため未使用。Workspace管理者側でApps Scriptウェブアプリの
匿名アクセスを許可できれば、`gas-log/Code.gs` を使う方式に戻すことも可能。
