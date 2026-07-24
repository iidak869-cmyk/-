/**
 * ジョブカン採用 RPA スクリプト
 *
 * 処理概要:
 *   1. AirWORK にログイン → 応募者一覧 CSV をダウンロード
 *   2. doda にログイン → 応募者情報 CSV をダウンロード
 *   3. 両 CSV の「前日」応募分を反映用リスト.xlsx に追記
 *   4. 完成ファイルを採用課連携フォルダにコピー
 *
 * 実行方法:
 *   cd C:\Users\iida869\Desktop\rpa-playwright
 *   node rpa\ジョブカン採用\jobcan-recruit.js
 *
 * 事前準備:
 *   同フォルダの config.json に認証情報を記入してください。
 */

const { chromium } = require('playwright');
const fs   = require('fs');
const path = require('path');
const readline = require('readline');

// ──────────────────────────────────────────────
// 設定読み込み
// ──────────────────────────────────────────────
const CONFIG_PATH = path.join(__dirname, 'config.json');
if (!fs.existsSync(CONFIG_PATH)) {
  console.error('❌ config.json が見つかりません。config.example.json をコピーして作成してください。');
  process.exit(1);
}
const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));

// パス設定（テスト環境）
const PATHS = {
  // ダウンロード先（Windows のデフォルトダウンロードフォルダ）
  downloadDir: cfg.downloadDir || path.join(process.env.USERPROFILE || 'C:\\Users\\iida869', 'Downloads'),

  // 応募者リスト格納フォルダ
  applicantListDir: cfg.applicantListDir ||
    'G:\\共有ドライブ\\RPA運用-技術課連携\\RPA検証用_テストデータ\\ジョブカン採用\\応募者リスト',

  // 反映先 Excel（テスト環境）
  targetExcel: cfg.targetExcel ||
    'G:\\共有ドライブ\\RPA運用-技術課連携\\RPA検証用_テストデータ\\ジョブカン採用\\反映用リストてすと.xlsx',

  // コピー先フォルダ（テスト環境）
  copyDestDir: cfg.copyDestDir ||
    'G:\\共有ドライブ\\RPA運用-技術課連携\\RPA検証用_テストデータ\\ジョブカン採用\\RPA運用-採用課連携',

  // AirWORK セッション保存先
  airworkSession: path.join(__dirname, 'airwork_session.json'),
  // doda セッション保存先
  dodaSession: path.join(__dirname, 'doda_session.json'),
};

const AIRWORK_URL = 'https://aw.airwork.net/';
const DODA_URL    = 'https://directhr.doda.jp/';

// ──────────────────────────────────────────────
// ユーティリティ
// ──────────────────────────────────────────────
function waitEnter(msg) {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(msg, () => { rl.close(); resolve(); });
  });
}

/** ダウンロードフォルダから最新の CSV を取得する */
function getLatestCsv(dir) {
  const files = fs.readdirSync(dir)
    .filter(f => f.endsWith('.csv'))
    .map(f => ({ name: f, mtime: fs.statSync(path.join(dir, f)).mtime }))
    .sort((a, b) => b.mtime - a.mtime);
  if (files.length === 0) throw new Error(`CSV が見つかりません: ${dir}`);
  return path.join(dir, files[0].name);
}

/** ファイルを安全にコピー（コピー先フォルダがなければ作成） */
function safeCopy(src, destDir, destName) {
  if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });
  const dest = path.join(destDir, destName);
  fs.copyFileSync(src, dest);
  console.log(`📂 コピー完了: ${dest}`);
  return dest;
}

/** 前日の日付文字列を返す (YYYY/MM/DD 形式) */
function yesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

/** 今日の日付文字列を返す (YYYY/MM/DD 形式) */
function today() {
  const d = new Date();
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

/** 今日の日付文字列を返す (YYYYMMDD 形式、フォルダ名用) */
function todayFolderName() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

/**
 * CSV 文字列を行 × 列の 2D 配列に変換する（ダブルクォート・改行対応）
 */
function parseCsv(text) {
  const rows = [];
  let cur = '', inQ = false, row = [];
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (inQ && text[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === ',' && !inQ) {
      row.push(cur); cur = '';
    } else if ((ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) && !inQ) {
      if (ch === '\r') i++;
      row.push(cur); cur = '';
      rows.push(row); row = [];
    } else {
      cur += ch;
    }
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

/**
 * CSV から前日の応募者データを抽出する
 * @param {string} csvPath  CSV ファイルパス
 * @param {string} source   'airwork' | 'doda'
 * @returns {{ name,kana,gender,email,phone,callTime,location,route,joinType,applyDate }[]}
 */
function extractPrevDayApplicants(csvPath, source) {
  const raw = fs.readFileSync(csvPath, 'utf8').replace(/^\uFEFF/, ''); // BOM 除去
  const rows = parseCsv(raw);
  if (rows.length < 2) return [];

  const headers = rows[0].map(h => h.trim());
  const yest    = yesterday();

  const col = name => headers.indexOf(name);

  const results = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(c => !c)) continue;

    const get = name => {
      const idx = col(name);
      return idx >= 0 && idx < r.length ? (r[idx] || '').trim() : '';
    };

    let applyDate = '';
    if (source === 'airwork') {
      // AirWORK: 「応募日時」列 → "YYYY/MM/DD HH:mm:ss" 形式
      const raw = get('応募日時');
      applyDate = raw.substring(0, 10).replace(/-/g, '/'); // 日付部分だけ
    } else {
      // doda: 「応募日時」列 → 同形式想定
      const raw = get('応募日時') || get('応募日');
      applyDate = raw.substring(0, 10).replace(/-/g, '/');
    }

    if (applyDate !== yest) continue; // 前日分のみ

    results.push({
      name:      get('応募者名'),
      kana:      get('ふりがな'),
      gender:    get('性別'),
      email:     get('メールアドレス'),
      phone:     get('電話番号'),
      callTime:  '', // CSV に対応列なし
      location:  get('応募勤務地') || get('勤務地'),
      route:     get('応募経路') || get('経路'),
      joinType:  get('入社種類') || (source === 'airwork' ? 'AirWORK' : 'doda'),
      applyDate,
    });
  }
  return results;
}

// ──────────────────────────────────────────────
// STEP 1: AirWORK からダウンロード
// ──────────────────────────────────────────────
async function downloadAirWork(browser) {
  console.log('\n📥 [AirWORK] ダウンロード開始...');

  const contextOpts = fs.existsSync(PATHS.airworkSession)
    ? { storageState: PATHS.airworkSession }
    : {};
  const context = await browser.newContext({ ...contextOpts, acceptDownloads: true });
  const page    = await context.newPage();

  await page.goto(AIRWORK_URL);

  // セッションが切れていたらログイン
  if (page.url().includes('login') || page.url().includes('signin')) {
    console.log('  ログイン画面を検出。手動でログインしてください。');
    await waitEnter('  ブラウザでログインが完了したら Enter キーを押してください...');
    await context.storageState({ path: PATHS.airworkSession });
    console.log('  セッションを保存しました。');
  }

  // 「応募者」メニューへ
  await page.click('text=応募者').catch(() => page.click('[href*="applicant"]'));
  await page.waitForLoadState('networkidle');

  // ページ最下部にある「応募者一覧をダウンロード」ボタン
  const downloadPromise = page.waitForEvent('download');
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
  await page.click('text=応募者一覧をダウンロード').catch(() =>
    page.click('button:has-text("ダウンロード")'));
  const download = await downloadPromise;
  const tmpPath  = await download.path();

  // リネームして応募者リストフォルダへ
  const destName = '応募者リスト(AirWORK).csv';
  if (!fs.existsSync(PATHS.applicantListDir)) fs.mkdirSync(PATHS.applicantListDir, { recursive: true });
  const destPath = path.join(PATHS.applicantListDir, destName);
  fs.copyFileSync(tmpPath, destPath);
  console.log(`  ✅ AirWORK CSV 保存: ${destPath}`);

  await context.close();
  return destPath;
}

// ──────────────────────────────────────────────
// STEP 2: doda からダウンロード
// ──────────────────────────────────────────────
async function downloadDoda(browser) {
  console.log('\n📥 [doda] ダウンロード開始...');

  const contextOpts = fs.existsSync(PATHS.dodaSession)
    ? { storageState: PATHS.dodaSession }
    : {};
  const context = await browser.newContext({ ...contextOpts, acceptDownloads: true });
  const page    = await context.newPage();

  await page.goto(DODA_URL);

  if (page.url().includes('login') || page.url().includes('signin')) {
    console.log('  ログイン画面を検出。手動でログインしてください。');
    await waitEnter('  ブラウザでログインが完了したら Enter キーを押してください...');
    await context.storageState({ path: PATHS.dodaSession });
    console.log('  セッションを保存しました。');
  }

  // 「doda 求人情報」メニューへ
  await page.click('text=doda 求人情報').catch(() =>
    page.click('[href*="job"]')).catch(() =>
    page.click('nav >> text=求人'));
  await page.waitForLoadState('networkidle');

  // 日付フィールドを当日に変更（開始日・終了日の両方）
  const todayStr = today(); // YYYY/MM/DD
  const dateInputs = await page.locator('input[type="date"], input[placeholder*="日付"], input[placeholder*="date"]').all();
  for (const inp of dateInputs) {
    await inp.fill(todayStr).catch(() => {});
  }

  // 「この条件で検索」
  await page.click('text=この条件で検索').catch(() =>
    page.click('button:has-text("検索")'));
  await page.waitForLoadState('networkidle');

  // 全選択チェックボックス
  await page.locator('thead input[type="checkbox"], th input[type="checkbox"]').first()
    .check().catch(() => page.locator('input[type="checkbox"]').first().check());

  // プルダウンで「応募者情報をCSV出力する」→ 実行
  const select = page.locator('select').first();
  await select.selectOption({ label: /応募者情報をCSV出力/ }).catch(() =>
    select.selectOption({ value: 'csv' }));
  const downloadPromise = page.waitForEvent('download');
  await page.click('button:has-text("実行"), input[type="submit"]:has-text("実行"), button:has-text("OK")')
    .catch(() => page.keyboard.press('Enter'));
  const download = await downloadPromise;
  const tmpPath  = await download.path();

  const destName = '応募者リスト(doda).csv';
  if (!fs.existsSync(PATHS.applicantListDir)) fs.mkdirSync(PATHS.applicantListDir, { recursive: true });
  const destPath = path.join(PATHS.applicantListDir, destName);
  fs.copyFileSync(tmpPath, destPath);
  console.log(`  ✅ doda CSV 保存: ${destPath}`);

  await context.close();
  return destPath;
}

// ──────────────────────────────────────────────
// STEP 3 & 4: Excel に書き込み → コピー
// ──────────────────────────────────────────────
async function writeToExcel(airworkCsv, dodaCsv) {
  console.log('\n📝 Excel への書き込み開始...');

  // ExcelJS を使用（Node.js でパスワードなし xlsx を操作）
  let ExcelJS;
  try {
    ExcelJS = require('exceljs');
  } catch {
    console.log('  exceljs をインストール中...');
    require('child_process').execSync('npm install exceljs', { cwd: path.join(__dirname, '../..'), stdio: 'inherit' });
    ExcelJS = require('exceljs');
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(PATHS.targetExcel);
  const sheet = workbook.worksheets[0];

  // ヘッダー行から列インデックスを解決
  const headerRow = sheet.getRow(1).values; // 1-indexed, [0] は undefined
  const colIndex = name => headerRow.findIndex(h => h === name);

  const C = {
    name:      colIndex('氏名'),
    kana:      colIndex('氏名（かな）'),
    gender:    colIndex('性別'),
    email:     colIndex('メールアドレス'),
    phone:     colIndex('電話番号'),
    callTime:  colIndex('お電話希望時間'),
    location:  colIndex('希望勤務地'),
    route:     colIndex('経路'),
    joinType:  colIndex('入社種類'),
    applyDate: colIndex('応募日'),
    rpaTime:   colIndex('RPA稼働時刻'),
  };

  console.log('  列マッピング:', C);

  const now = new Date().toLocaleString('ja-JP');

  // AirWORK 抽出
  const airworkApplicants = extractPrevDayApplicants(airworkCsv, 'airwork');
  console.log(`  AirWORK 前日応募者: ${airworkApplicants.length} 件`);

  // doda 抽出
  const dodaApplicants = extractPrevDayApplicants(dodaCsv, 'doda');
  console.log(`  doda 前日応募者: ${dodaApplicants.length} 件`);

  const allApplicants = [...airworkApplicants, ...dodaApplicants];

  if (allApplicants.length === 0) {
    console.log('  ℹ️ 追記対象の応募者データが 0 件です。');
  }

  // 最終行の次から書き込む
  let nextRow = sheet.lastRow ? sheet.lastRow.number + 1 : 2;
  for (const a of allApplicants) {
    const row = sheet.getRow(nextRow);
    if (C.name      > 0) row.getCell(C.name).value      = a.name;
    if (C.kana      > 0) row.getCell(C.kana).value      = a.kana;
    if (C.gender    > 0) row.getCell(C.gender).value    = a.gender;
    if (C.email     > 0) row.getCell(C.email).value     = a.email;
    if (C.phone     > 0) row.getCell(C.phone).value     = a.phone;
    if (C.callTime  > 0) row.getCell(C.callTime).value  = a.callTime;
    if (C.location  > 0) row.getCell(C.location).value  = a.location;
    if (C.route     > 0) row.getCell(C.route).value     = a.route;
    if (C.joinType  > 0) row.getCell(C.joinType).value  = a.joinType;
    if (C.applyDate > 0) row.getCell(C.applyDate).value = a.applyDate;
    if (C.rpaTime   > 0) row.getCell(C.rpaTime).value   = now;
    row.commit();
    nextRow++;
  }

  await workbook.xlsx.writeFile(PATHS.targetExcel);
  console.log(`  ✅ Excel 更新完了 (${allApplicants.length} 件追記): ${PATHS.targetExcel}`);

  // コピー先: YYYYMMDD フォルダを作成してコピー
  const folderName = todayFolderName();
  const destDir    = path.join(PATHS.copyDestDir, folderName);
  const destName   = path.basename(PATHS.targetExcel);
  safeCopy(PATHS.targetExcel, destDir, destName);
}

// ──────────────────────────────────────────────
// メイン
// ──────────────────────────────────────────────
(async () => {
  const browser = await chromium.launch({ headless: false });

  try {
    // 1. AirWORK CSV ダウンロード
    const airworkCsv = await downloadAirWork(browser);

    // 2. doda CSV ダウンロード
    const dodaCsv = await downloadDoda(browser);

    // 3. & 4. Excel 書き込み & コピー
    await writeToExcel(airworkCsv, dodaCsv);

    console.log('\n✅ すべての処理が完了しました。');
  } catch (err) {
    console.error('\n❌ エラーが発生しました:', err.message);
    console.error(err.stack);
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
