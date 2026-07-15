// No.14 契約数更新
//
// ほうこっくんで「前日/次日 × 新規/リピート」の4パターンを検索し、
// 契約一覧の各報告内容からプロダクト日報Excelへ転記、
// 追記分をCSVに出力してスプシ転記用に使う。
//
// 事前準備:
//   1. rpa-playwright 直下で npm install exceljs（playwrightは既存環境と共用）
//   2. node .\ほうこっくん_ログイン保存.js でセッションを保存
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
// キー = プロダクト日報「入六」シートの列見出し、値 = ほうこっくん詳細画面上のセレクタ
// TODO: --探索 モードで詳細画面のセレクタを確認して埋める。
//       ほうこっくん側に存在しない項目（手作業で埋める列）は削除してOK
const FIELDS = {
  "契約日": "TODO_セレクタ",
  "管理番号": "TODO_セレクタ",
  "顧客名": "TODO_セレクタ",
  "営業担当者": "TODO_セレクタ",
  "営業部署": "TODO_セレクタ",
  "会社所在地": "TODO_セレクタ",
  "一括orリース": "TODO_セレクタ",
  "業種": "TODO_セレクタ",
  "業種カテゴリ": "TODO_セレクタ",
  "営業報告の添付画像": "TODO_セレクタ",
  "格納": "TODO_セレクタ",
  "Cyteki": "TODO_セレクタ",
  "売上（グロス）": "TODO_セレクタ",
  "売上（ネット②）": "TODO_セレクタ",
  "クレカの有無": "TODO_セレクタ",
};

// スプシ（テスト反映先）の列順。A列からこの順に貼り付ける
const SHEET_COLUMNS = [
  "契約日", "管理番号", "顧客名", "営業担当者", "営業部署", "会社所在地",
  "一括orリース", "業種", "業種カテゴリ", "営業報告の添付画像", "格納", "Cyteki",
  "売上（グロス）", "売上（ネット②）", "クレカの有無",
];
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

  // 見出し行（config.productNippo.headerRow）から列位置を特定
  const headerRowNum = config.productNippo.headerRow ?? 1;
  const headerRow = ws.getRow(headerRowNum);
  const colIndex = {};
  headerRow.eachCell((cell, col) => {
    colIndex[String(cell.value ?? "").replace(/\s+/g, "")] = col;
  });

  const norm = (s) => String(s ?? "").replace(/\s+/g, "");

  // 既存データから重複チェック用キーを収集
  const dedupeCols = config.productNippo.dedupeColumns;
  const existingKeys = new Set();
  ws.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowNum) return;
    const key = dedupeCols.map((c) => String(row.getCell(colIndex[norm(c)] ?? 0).value ?? "").trim()).join("|");
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
      if (colIndex[norm(label)]) newRow.getCell(colIndex[norm(label)]).value = value;
    }
    existingKeys.add(key);
    appended.push(record);
  }

  if (appended.length > 0) {
    await wb.xlsx.writeFile(config.productNippo.excelPath);
  }
  return appended;
}

// テスト反映先スプシへ追記分を貼り付け（gsheet_session.json のセッションを使用）
async function appendToSpreadsheet(records) {
  if (records.length === 0) return;

  const tsv = records
    .map((r) => SHEET_COLUMNS.map((c) => String(r[c] ?? "").replace(/[\t\r\n]+/g, " ")).join("\t"))
    .join("\n");

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({ storageState: config.spreadsheet.sessionFile });
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  const page = await context.newPage();
  await page.goto(config.spreadsheet.url);
  await page.waitForLoadState("networkidle");
  await page.waitForTimeout(3000); // シートの描画待ち

  await page.evaluate((text) => navigator.clipboard.writeText(text), tsv);

  // データ末尾へ移動 → 行頭(A列) → 1行下の空行へ → 貼り付け
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Control+V");
  await page.waitForTimeout(5000); // 貼り付け＆自動保存待ち

  await browser.close();
  console.log(`スプシへ追記しました: ${config.spreadsheet.url}`);
}

// スプシ貼り付け用CSVを出力（自動貼り付けに失敗した場合の予備）
function writeCsv(records) {
  if (records.length === 0) return;
  const headers = Object.keys(records[0]);
  const lines = [headers.join(",")];
  for (const r of records) {
    lines.push(headers.map((h) => `"${String(r[h] ?? "").replace(/"/g, '""')}"`).join(","));
  }
  fs.writeFileSync(config.spreadsheet.csvOutput, "\uFEFF" + lines.join("\r\n"), "utf8");
  console.log(`追記分CSVを出力しました: ${config.spreadsheet.csvOutput}`);
}

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    storageState: fs.existsSync(config.houkokkun.sessionFile) ? config.houkokkun.sessionFile : undefined,
  });

  // ほうこっくんは http://…/top/ へリダイレクトするがポート80が閉じているため、
  // http を https に強制的に置き換える（通常ブラウザの自動https格上げ相当）
  await context.route("http://houkoku.access-mgr.biz/**", (route) => {
    const url = route.request().url().replace(/^http:/, "https:");
    return route.fulfill({ status: 302, headers: { location: url } });
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
  await appendToSpreadsheet(appended);
})();
