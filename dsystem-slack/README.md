# Dシステム カレンダー → Slack 自動投稿（CS課スケジュール）

社内グループウェア「Dシステム」のスケジュールページ
`http://win2012sv-aws/DSystem/Schedule.aspx` から CS課メンバーの予定を抜き出し、
Slack のチャンネル **#dsystem-cs-schedule** に投稿するツールです。

## 重要な前提

- `win2012sv-aws` は **社内ネットワーク専用のサーバー** のため、クラウドや GAS からは接続できません。
- そのため、このスクリプトは **社内の Windows PC（またはサーバー）上で実行** します。
  PowerShell 3.0 以降（Windows 8 / Server 2012 以降の標準搭載）で動作します。
- Slack への送信は Incoming Webhook を使用します（下記手順で発行）。

## セットアップ手順

### 1. Slack Incoming Webhook を発行する

1. https://api.slack.com/apps で「Create New App」→「From scratch」でアプリを作成
   （名前は「Dシステムスケジュール」など任意、ワークスペースは自社のもの）
2. 左メニュー「Incoming Webhooks」→ 有効化（On）
3. 「Add New Webhook to Workspace」→ 投稿先に **#dsystem-cs-schedule** を選択
4. 発行された `https://hooks.slack.com/services/...` のURLを控える

### 2. スクリプトを社内PCに配置して動作確認する

`Export-DSystemSchedule.ps1` を社内PCの任意のフォルダ（例: `C:\Tools\dsystem-slack\`）にコピーし、
PowerShell で以下の順に実行します。

```powershell
# (1) まずページのHTMLを保存して、スケジュール表が取れているか確認
.\Export-DSystemSchedule.ps1 -DumpOnly

# (2) 解析結果をSlackに送らずコンソールで確認
.\Export-DSystemSchedule.ps1 -NoPost

# (3) 問題なければ実際に投稿
.\Export-DSystemSchedule.ps1 -WebhookUrl 'https://hooks.slack.com/services/XXX/YYY/ZZZ'
```

> **実行ポリシーのエラーが出る場合**は、次のように起動してください:
> `powershell -NoProfile -ExecutionPolicy Bypass -File .\Export-DSystemSchedule.ps1 -NoPost`

> **解析に失敗した場合**（「スケジュール表を自動検出できませんでした」と表示された場合）は、
> 同じフォルダに保存される `Schedule_dump.html` を共有してください。
> Dシステムの実際のページ構造に合わせてパーサーを調整します。
> ※ 表示するグループの切り替えがURLのクエリではなく画面操作（ポストバック）で行われる場合も、
> ダンプがあれば対応方法を判断できます。

### 3. 毎日自動で投稿する（タスクスケジューラ登録）

管理者権限のPowerShell/コマンドプロンプトで以下を実行すると、毎朝8:30に自動投稿されます。

```bat
schtasks /Create /TN "DSystemScheduleToSlack" /SC DAILY /ST 08:30 ^
  /TR "powershell -NoProfile -ExecutionPolicy Bypass -File C:\Tools\dsystem-slack\Export-DSystemSchedule.ps1 -WebhookUrl https://hooks.slack.com/services/XXX/YYY/ZZZ"
```

- 時刻は `/ST` を変更してください。平日のみにする場合は `/SC WEEKLY /D MON,TUE,WED,THU,FRI` を使います。
- PCがスリープしていると実行されないため、常時起動しているPCかサーバーでの登録を推奨します。

## パラメータ一覧

| パラメータ | 既定値 | 説明 |
|---|---|---|
| `-ScheduleUrl` | `http://win2012sv-aws/DSystem/Schedule.aspx` | 取得対象のURL。グループ指定のクエリが必要な場合はここに含める |
| `-GroupName` | `CS課` | 投稿メッセージの見出しに使うグループ名 |
| `-WebhookUrl` | 環境変数 `SLACK_WEBHOOK_URL` | Slack Incoming Webhook のURL |
| `-DumpOnly` | - | ページHTMLを保存するだけで終了（構造確認用） |
| `-DumpPath` | スクリプトと同じフォルダの `Schedule_dump.html` | ダンプ保存先 |
| `-NoPost` | - | Slackに投稿せず整形結果をコンソール表示 |
| `-AllDays` | - | 表示中の全日付分を投稿（既定は今日の分のみ） |

## 動作の仕組み

1. `Invoke-WebRequest -UseDefaultCredentials` でスケジュールページを取得
   （社内ASP.NETサイトで一般的なWindows統合認証を想定。実行ユーザーの資格情報で認証されます）
2. HTML内から「ヘッダー行に日付が並び、先頭列にメンバー名がある表」を自動検出
3. 今日の日付の列（`-AllDays` 指定時は全列）をメンバーごとに整形
4. Incoming Webhook 経由で #dsystem-cs-schedule に投稿（長文は自動分割）

## 投稿イメージ

```
【CS課 スケジュール】2026/07/09 (木)
山田 太郎
  ・09:00-10:00 定例会議
  ・13:00-15:00 顧客訪問
佐藤 花子
  ・(予定なし)
```
