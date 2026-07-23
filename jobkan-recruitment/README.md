# ジョブカン採用 作業自動化（AirWORK / doda 応募者CSV取得）

「ジョブカン採用 作業手順」書の手順1〜22をPlaywright（Node.js）で自動化するプロジェクトです。
ローカルPC（Windows）上で実行し、Windowsタスクスケジューラで定期実行することを想定しています。

## 現在の実装状況

| 手順 | 内容 | 状況 |
|---|---|---|
| 1〜7 | AirWORKログイン〜応募者CSVダウンロード・格納・ログアウト | 実装済み（要セレクタ確認） |
| 8〜17 | dodaログイン〜応募者CSVダウンロード・格納・ログアウト | 実装済み（要セレクタ確認） |
| 18〜21 | 反映用リスト.xlsxへの転記（応募経路付き） | **未実装**（CSV/xlsxのサンプルファイル待ち） |
| 22 | 共有ドライブへのコピー | 実装済み |

**重要:** 1〜17のコードは実際の画面を見ずに作成した仮のセレクタ（`TODO(要確認)`コメント箇所）を含みます。
初回は必ず `HEADLESS=false` で実行し、目視しながら動作確認・調整してください。

## セットアップ

### 1. 前提

- Node.js 18以上がインストール済みであること
- 対象PCに Google Chrome や既存ブラウザは不要（Playwrightが専用のChromiumを自動インストールします）

### 2. インストール

```bat
cd jobkan-recruitment
npm install
```

（`npm install` 後に自動でPlaywright用Chromiumがダウンロードされます）

### 3. .env の作成

`.env.example` を `.env` にコピーし、値を埋めてください。

```bat
copy .env.example .env
```

- `AIRWORK_LOGIN_URL` / `DODA_LOGIN_URL`: 実際のログインページURLに変更
- `AIRWORK_ID` / `AIRWORK_PASSWORD` / `DODA_ID` / `DODA_PASSWORD`: ログイン情報
- `APPLICANT_LIST_DIR`: 「応募者リスト」フォルダの実パス（未作成の場合は仮パスのままでOK。指定フォルダが無ければ自動作成されます）
- `REFLECT_EXCEL_PATH` / `APPLICANT_CSV_ROOT_DIR` / `SHARED_DRIVE_DEST_DIR`: 手順18以降で使用（現時点では未使用）

`.env` には認証情報が入るため、Gitには絶対にコミットしないでください（`.gitignore` 済み）。

### 4. 動作確認（手動実行）

```bat
npm start
```

`HEADLESS=false` の場合、Chromiumウィンドウが開いて実際の操作が目に見える形で進みます。
`TODO(要確認)` のセレクタが実画面とズレている場合、そこでエラーになりエラースクリーンショットが
`work/error-screenshots/` に保存されます。それを見て `src/airwork.js` / `src/doda.js` の該当箇所を修正してください。

### 5. セレクタの確認・修正方法（推奨）

推測で書いたセレクタを実際の画面に合わせて直す一番確実な方法は、Playwrightの記録機能を使うことです。

```bat
npx playwright codegen https://実際のログインURL
```

ブラウザが開き、実際にクリック・入力した操作がコードとして自動生成されます。生成されたコードの
`page.getByRole(...)` や `page.locator(...)` 部分を、`src/airwork.js` / `src/doda.js` の該当箇所に
コピーして置き換えてください。

## Windowsタスクスケジューラでの自動実行

1. `run.bat` の中身を確認（このプロジェクトフォルダで `node src\index.js` を実行し、ログを `work\run.log` に出力します）
2. タスクスケジューラを開く（Win+R → `taskschd.msc`）
3. 「タスクの作成」を選択
   - **全般タブ**: 名前「ジョブカン採用_応募者CSV自動取得」、「ユーザーがログオンしているかどうかにかかわらず実行する」にチェック
   - **トリガータブ**: 新規 → 毎日、実行したい時刻を指定
   - **操作タブ**: 新規 → プログラム/スクリプトに `run.bat` のフルパスを指定、開始（オプション）に `jobkan-recruitment` フォルダのフルパスを指定
   - **条件タブ**: 「AC電源で実行時のみ」等、必要に応じて調整
4. 保存時にWindowsのログオンパスワードを求められた場合は入力

コマンドラインから登録する場合の例（管理者権限のコマンドプロンプトで実行）:

```bat
schtasks /create /tn "ジョブカン採用_応募者CSV自動取得" /tr "C:\path\to\jobkan-recruitment\run.bat" /sc daily /st 09:00 /ru "%USERNAME%"
```

## フォルダ構成

```
jobkan-recruitment/
├── package.json
├── .env.example        # コピーして .env を作成
├── run.bat             # タスクスケジューラから呼び出すバッチ
├── src/
│   ├── config.js       # 環境変数・パス設定の読み込み
│   ├── airwork.js       # 手順1〜7
│   ├── doda.js           # 手順8〜17
│   ├── excel.js          # 手順18〜21（未実装・サンプル待ち）
│   ├── copyToSharedDrive.js  # 手順22
│   └── index.js          # 全体のオーケストレーション
└── work/                # 実行時に自動生成（応募者リスト・ログ・エラースクショの仮置き場）
```

`APPLICANT_LIST_DIR` 等の実フォルダが用意でき次第、`.env` のパスを本番パスに書き換えるだけで
そのまま本番フォルダに保存されるようになります。

## クラウド実行（タスクスケジューラを使わない場合）について

Windowsタスクスケジューラの代わりにClaude Codeのルーチン機能でクラウド実行することも可能ですが、
その場合 `G:\共有ドライブ\...` のようなローカルネットワークドライブには直接アクセスできないため、
手順22のコピー処理をGoogle Drive API経由のアップロードなどに置き換える必要があります。
現在はローカルPC + タスクスケジューラ方式で進めているため未対応です。
