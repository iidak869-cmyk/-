// No.14 契約数更新
//
// ほうこっくんで「前日/次日 × 新規/リピート」の4パターンを検索し、
// 契約一覧の各報告内容からプロダクト日報Excelへ転記、
// 追記分をCSVに出力してスプシ転記用に使う。
//
// 事前準備:
//   1. config.example.json をコピーして config.json を作り、TODO箇所を埋める
//   2. npm install exceljs   （playwrightは既存の rpa-playwright と共用）
//   3. node .\ほうこっくん_ログイン保存.js でセッションを保存
//
// 実行:
//   node .\契約数更新.js            … 本番実行
//   node .\契約数更新.js --探索     … 画面を開いてPlaywright Inspectorを起動（セレクタ調査用）
//   node .\契約数更新.js --dry-run  … Excelに書き込まず、抽出結果の表示だけ行う

const { chromium } = require("playwright");
const fs = require("fs");
const ExcelJS = require("exceljs");

const config = require("./config.json");

// ============================================================
// ここから要調整：ほうこっくんの実際の画面に合わせて修正する
// （--探索 モードで起動し、Inspectorの「Pick locator」で確認する）
// ============================================================
const SELECTORS = {
  // メニュー画面左上の「前日」「次日」ボタン
  dayButton: (label) => `text=${label}`,
  // 項目一覧の「営業」
  salesItem: `text=営業`,
  // 報告内容のプルダウン
  reportTypeSelect: `select`, // TODO: 実際のセレクタに変更
  // プルダウン右側の「検索」ボタン
  searchButton: `text=検索`,
  // 検索結果（契約一覧）の行
  resultRows: `table tbody tr`, // TODO: 実際のセレクタに変更
  // 各行の報告内容を開くリンク/ボタン
  detailLink: `a`, // TODO: 実際のセレクタに変更
  // 詳細画面を閉じて一覧に戻る操作（戻るボタン等）
  backToList: `text=戻る`, // TODO: 実際のセレクタに変更
};

// 報告詳細画面から転記する項目。
// キー = プロダクト日報の列見出し、値 = 詳細画面上のセレクタ
// TODO: 転記項目の一覧が判明したら追記する
const FIELDS = {
  "顧客名": "TODO_セレクタ",
  "契約日": "TODO_セレクタ",
  // 例) "担当者": "#tantosha",
};
// ============================================================

const isExplore = process.argv.includes("--探索");
const isDryRun = process.argv.includes("--dry-run");

async function extractPattern(page, pattern) {
  // 前日/次日 → 営業 → 報告内容プルダウン → 検索
  await page.click(SELECTORS.dayButton(pattern.day));
  await page.click(SELECTORS.salesItem);
  await page.selectOption(SELECTORS.reportTypeSelect, { label: pattern.type });
  await page.click(SELECTORS.searchButton);
  await page.waitForLoadState("networkidle");

  const records = [];
  const rowCount = await page.locator(SELECTORS.resultRows).count();
  console.log(`  ${pattern.day}×${pattern.type}: ${rowCount}件`);

  for (let i = 0; i < rowCount; i++) {
    // 詳細画面を開いて項目を抽出
    await page.locator(SELECTORS.resultRows).nth(i).locator(SELECTORS.detailLink).first().click();
    await page.waitForLoadState("networkidle");

    const record = { 区分: `${pattern.day}×${pattern.type}` };
    for (const [label, selector] of Object.entries(FIELDS)) {
      record[label] = (await page.locator(selector).first().textContent().catch(() => "")) ?? "";
      record[label] = record[label].trim();
    }
    records.push(record);

    await page.click(SELECTORS.backToList);
    await page.waitForLoadState("networkidle");
  }
  return records;
}

// プロダクト日報Excelに追記（重複チェック付き）
async function appendToExcel(records) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(config.productNippo.excelPath);
  const ws = wb.getWorksheet(config.productNippo.sheetName);
  if (!ws) throw new Error(`シートが見つかりません: ${config.productNippo.sheetName}`);

  // 1行目を見出しとして列位置を特定
  const headerRow = ws.getRow(1);
  const colIndex = {};
  headerRow.eachCell((cell, col) => {
    colIndex[String(cell.value).trim()] = col;
  });

  // 既存データから重複チェック用キーを収集
  const dedupeCols = config.productNippo.dedupeColumns;
  const existingKeys = new Set();
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const key = dedupeCols.map((c) => String(row.getCell(colIndex[c] ?? 0).value ?? "").trim()).join("|");
    if (key.replace(/\|/g, "")) existingKeys.add(key);
  });

  const appended = [];
  for (const record of records) {
    const key = dedupeCols.map((c) => String(record[c] ?? "").trim()).join("|");
    if (existingKeys.has(key)) {
      console.log(`  スキップ(重複): ${key}`);
      continue;
    }
    const newRow = ws.addRow([]);
    for (const [label, value] of Object.entries(record)) {
      if (colIndex[label]) newRow.getCell(colIndex[label]).value = value;
    }
    existingKeys.add(key);
    appended.push(record);
  }

  if (appended.length > 0) {
    await wb.xlsx.writeFile(config.productNippo.excelPath);
  }
  return appended;
}

// スプシ貼り付け用CSVを出力
function writeCsv(records) {
  if (records.length === 0) return;
  const headers = Object.keys(records[0]);
  const lines = [headers.join(",")];
  for (const r of records) {
    lines.push(headers.map((h) => `"${String(r[h] ?? "").replace(/"/g, '""')}"`).join(","));
  }
  fs.writeFileSync(config.spreadsheet.csvOutput, "﻿" + lines.join("\r\n"), "utf8");
  console.log(`追記分CSVを出力しました: ${config.spreadsheet.csvOutput}`);
}

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    storageState: fs.existsSync(config.houkokkun.sessionFile) ? config.houkokkun.sessionFile : undefined,
  });
  const page = await context.newPage();
  await page.goto(config.houkokkun.url);

  if (isExplore) {
    // セレクタ調査用: Inspectorを開いたまま止める
    await page.pause();
    await browser.close();
    return;
  }

  const allRecords = [];
  for (const pattern of config.patterns) {
    allRecords.push(...(await extractPattern(page, pattern)));
  }
  await browser.close();

  console.log(`抽出合計: ${allRecords.length}件`);
  if (isDryRun) {
    console.table(allRecords);
    return;
  }

  const appended = await appendToExcel(allRecords);
  console.log(`プロダクト日報へ追記: ${appended.length}件`);

  writeCsv(appended);
  console.log(`スプシへの転記: ${config.spreadsheet.url} を開いて ${config.spreadsheet.csvOutput} の内容を貼り付けてください`);
})();
