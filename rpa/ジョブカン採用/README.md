# ジョブカン採用 RPA

AirWORK と doda の 2 サイトから応募者 CSV を取得し、管理用 Excel に前日分を転記します。

## セットアップ

```powershell
cd C:\Users\iida869\Desktop\rpa-playwright\rpa\ジョブカン採用

# config.json を作成
Copy-Item config.example.json config.json
# config.json を開いて認証情報・パスを確認・編集する

# 依存パッケージをインストール（初回のみ）
npm install exceljs
```

## 実行

```powershell
cd C:\Users\iida869\Desktop\rpa-playwright
node rpa\ジョブカン採用\jobcan-recruit.js
```

## 処理フロー

| # | 処理 | 備考 |
|---|------|------|
| 1 | AirWORK ログイン | セッションを airwork_session.json に保存（2回目以降はスキップ） |
| 2 | 応募者一覧 CSV ダウンロード | `応募者リスト(AirWORK).csv` にリネームして格納 |
| 3 | doda ログイン | セッションを doda_session.json に保存 |
| 4 | 応募者情報 CSV ダウンロード | 当日日付で検索 → `応募者リスト(doda).csv` にリネームして格納 |
| 5 | 反映用 Excel に前日分を追記 | `反映用リストてすと.xlsx` の末尾に追加 |
| 6 | Excel を日付フォルダにコピー | `RPA運用-採用課連携\YYYYMMDD\` フォルダに保存 |

## 列マッピング

| Excel 列 | AirWORK CSV | doda CSV |
|----------|-------------|----------|
| 氏名 | 応募者名 | 応募者名 |
| 氏名（かな） | ふりがな | ふりがな |
| 性別 | 性別 | 性別 |
| メールアドレス | メールアドレス | メールアドレス |
| 電話番号 | 電話番号 | 電話番号 |
| お電話希望時間 | （CSV になし・空欄） | （CSV になし・空欄） |
| 希望勤務地 | 応募勤務地 | 勤務地 |
| 経路 | 応募経路 | 経路 |
| 入社種類 | AirWORK（固定） | doda（固定） |
| 応募日 | 応募日時（日付部分） | 応募日時（日付部分） |
| RPA稼働時刻 | スクリプト実行時刻 | スクリプト実行時刻 |

## 注意事項

- `config.json` は Git にコミットしないこと（.gitignore に含まれています）
- ログインセッションが切れた場合は `*_session.json` を削除して再実行
- 初回実行時は手動ログインが必要（ブラウザが開くので操作してください）
- doda の画面は仕様変更が多い。セレクターが合わない場合は `jobcan-recruit.js` の該当行を修正する
