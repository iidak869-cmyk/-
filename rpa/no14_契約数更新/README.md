# No.14 契約数更新（ほうこっくん → プロダクト日報 → スプシ）

ほうこっくんで「前日/次日 × 新規/リピート」の4パターンを検索し、
契約一覧の各報告内容をプロダクト日報Excelへ転記（重複チェック付き）、
追記分をCSVに出力してスプシ転記に使うPlaywrightスクリプトです。

## セットアップ（ローカルWindows側）

```powershell
cd C:\Users\iida869\Desktop\rpa-playwright
# このフォルダ(no14_契約数更新)一式をコピーして配置

npm install exceljs

# config.example.json をコピーして config.json を作成し、TODO箇所を埋める
copy .\no14_契約数更新\config.example.json .\no14_契約数更新\config.json

# ほうこっくんのセッションを保存（Googleスプシ_ログイン保存.js と同じ方式）
node .\no14_契約数更新\ほうこっくん_ログイン保存.js
```

## 実行方法

| コマンド | 動作 |
|---|---|
| `node .\契約数更新.js --探索` | 画面を開いてPlaywright Inspectorを起動（セレクタ調査用） |
| `node .\契約数更新.js --dry-run` | Excelに書き込まず抽出結果の表示だけ |
| `node .\契約数更新.js` | 本番実行 |

まず `--探索` でセレクタを確認 → `--dry-run` で抽出確認 → 本番、の順を推奨。

## 完成までに必要な情報（未確定分）

業務詳細シートに記載がなかったため、以下が埋まると完成します。

1. **ほうこっくんの管理画面URL** → `config.json` の `houkokkun.url`
2. **転記する項目の一覧**（顧客名、契約日、金額…など）→ `契約数更新.js` の `FIELDS`
3. **プロダクト日報Excelのパスとシート名** → `config.json` の `productNippo`
4. **転記先スプシのURLと貼り付け先** → `config.json` の `spreadsheet.url`
5. **画面のセレクタ**（検索結果テーブル・詳細リンク・プルダウン等）
   → `--探索` モードで起動し、Inspectorの「Pick locator」で調べて `SELECTORS` を更新

## 重複チェックの仕組み

`config.json` の `productNippo.dedupeColumns`（既定: 日付+顧客名）の組み合わせで
既存行と照合し、一致する行はスキップします。
