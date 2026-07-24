const { chromium } = require('playwright');
const fs   = require('fs');
const path = require('path');
const readline = require('readline');

const CONFIG_PATH = path.join(__dirname, 'config.json');
if (!fs.existsSync(CONFIG_PATH)) { console.error('config.json が見つかりません。'); process.exit(1); }
const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8').replace(/^﻿/, ''));

const PATHS = {
  downloadDir:      cfg.downloadDir || path.join(process.env.USERPROFILE || '', 'Downloads'),
  applicantListDir: cfg.applicantListDir,
  targetExcel:      cfg.targetExcel,
  copyDestDir:      cfg.copyDestDir,
  airworkSession:   path.join(__dirname, 'airwork_session.json'),
  dodaSession:      path.join(__dirname, 'doda_session.json'),
};

const AIRWORK_URL = 'https://ats.rct.airwork.net/';
const DODA_URL    = 'https://assist.doda.jp/';

function waitEnter(msg) {
  return new Promise(resolve => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(msg, () => { rl.close(); resolve(); });
  });
}

function getLatestCsv(dir) {
  const files = fs.readdirSync(dir)
    .filter(f => f.toLowerCase().endsWith('.csv'))
    .map(f => ({ name: f, mtime: fs.statSync(path.join(dir, f)).mtime.getTime() }))
    .sort((a, b) => b.mtime - a.mtime);
  if (files.length === 0) throw new Error('CSVが見つかりません: ' + dir);
  return path.join(dir, files[0].name);
}

function safeCopy(src, destDir, destName) {
  if (!fs.existsSync(destDir)) fs.mkdirSync(destDir, { recursive: true });
  const dest = path.join(destDir, destName);
  fs.copyFileSync(src, dest);
  console.log('コピー完了: ' + dest);
  return dest;
}

function yesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  return d.getFullYear() + '/' + String(d.getMonth()+1).padStart(2,'0') + '/' + String(d.getDate()).padStart(2,'0');
}

function todayFolderName() {
  const d = new Date();
  return '' + d.getFullYear() + String(d.getMonth()+1).padStart(2,'0') + String(d.getDate()).padStart(2,'0');
}

function parseCsv(text) {
  const rows = []; let cur = '', inQ = false, row = [];
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '"') { if (inQ && text[i+1] === '"') { cur += '"'; i++; } else inQ = !inQ; }
    else if (ch === ',' && !inQ) { row.push(cur); cur = ''; }
    else if ((ch === '\n' || (ch === '\r' && text[i+1] === '\n')) && !inQ) { if (ch === '\r') i++; row.push(cur); cur = ''; rows.push(row); row = []; }
    else { cur += ch; }
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

function extractPrevDayApplicants(csvPath, source) {
  const raw = fs.readFileSync(csvPath).toString('utf8').replace(/^﻿/, '');
  const rows = parseCsv(raw);
  if (rows.length < 2) return [];
  const headers = rows[0].map(h => h.trim());
  const yest = yesterday();
  const get = (r, name) => { const idx = headers.indexOf(name); return idx >= 0 && idx < r.length ? (r[idx]||'').trim() : ''; };
  const results = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || r.every(c => !c)) continue;
    const rawDate = get(r,'応募日時') || get(r,'応募日');
    const applyDate = rawDate.substring(0,10).replace(/-/g,'/');
    if (applyDate !== yest) continue;
    results.push({
      name:      get(r,'応募者名'),
      kana:      get(r,'ふりがな'),
      gender:    get(r,'性別'),
      email:     get(r,'メールアドレス'),
      phone:     get(r,'電話番号'),
      callTime:  '',
      location:  get(r,'応募勤務地') || get(r,'勤務地'),
      route:     get(r,'応募経路')   || get(r,'経路'),
      joinType:  source === 'airwork' ? 'AirWORK' : 'doda',
      applyDate,
    });
  }
  return results;
}

async function downloadManual(browser, url, sessionPath, siteName) {
  const contextOpts = fs.existsSync(sessionPath) ? { storageState: sessionPath } : {};
  const context = await browser.newContext({ ...contextOpts, acceptDownloads: true });
  const page = await context.newPage();
  await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });
  const before = fs.readdirSync(PATHS.downloadDir).filter(f => f.toLowerCase().endsWith('.csv'));
  console.log('');
  console.log('========================================');
  console.log('[' + siteName + '] ブラウザで操作してください:');
  console.log('  1. ログイン（未ログインの場合）');
  console.log('  2. 応募者一覧のCSVをダウンロード');
  console.log('  3. ダウンロード完了後、ここに戻る');
  console.log('========================================');
  await waitEnter('ダウンロード完了後、Enterを押してください...');
  await context.storageState({ path: sessionPath });
  await context.close();
  const after = fs.readdirSync(PATHS.downloadDir).filter(f => f.toLowerCase().endsWith('.csv'));
  const newFiles = after.filter(f => !before.includes(f));
  const csvPath = newFiles.length > 0 ? path.join(PATHS.downloadDir, newFiles[0]) : getLatestCsv(PATHS.downloadDir);
  console.log('  取得CSV: ' + csvPath);
  return csvPath;
}

async function writeToExcel(airworkCsv, dodaCsv) {
  console.log('');
  console.log('Excel書き込み開始...');
  const ExcelJS = require('exceljs');
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(PATHS.targetExcel);
  const sheet = workbook.worksheets[0];
  const headerRow = sheet.getRow(1).values;
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
  const now = new Date().toLocaleString('ja-JP');
  if (!fs.existsSync(PATHS.applicantListDir)) fs.mkdirSync(PATHS.applicantListDir, { recursive: true });
  const awDest = path.join(PATHS.applicantListDir, '応募者リスト(AirWORK).csv');
  const ddDest = path.join(PATHS.applicantListDir, '応募者リスト(doda).csv');
  fs.copyFileSync(airworkCsv, awDest);
  fs.copyFileSync(dodaCsv, ddDest);
  const awA = extractPrevDayApplicants(awDest, 'airwork');
  const ddA = extractPrevDayApplicants(ddDest, 'doda');
  console.log('  AirWORK前日分: ' + awA.length + '件 / doda前日分: ' + ddA.length + '件');
  const all = [...awA, ...ddA];
  if (all.length === 0) console.log('  ※前日の応募者データが0件でした');
  let nextRow = sheet.lastRow ? sheet.lastRow.number + 1 : 2;
  for (const a of all) {
    const row = sheet.getRow(nextRow);
    if (C.name>0)      row.getCell(C.name).value      = a.name;
    if (C.kana>0)      row.getCell(C.kana).value      = a.kana;
    if (C.gender>0)    row.getCell(C.gender).value    = a.gender;
    if (C.email>0)     row.getCell(C.email).value     = a.email;
    if (C.phone>0)     row.getCell(C.phone).value     = a.phone;
    if (C.callTime>0)  row.getCell(C.callTime).value  = a.callTime;
    if (C.location>0)  row.getCell(C.location).value  = a.location;
    if (C.route>0)     row.getCell(C.route).value     = a.route;
    if (C.joinType>0)  row.getCell(C.joinType).value  = a.joinType;
    if (C.applyDate>0) row.getCell(C.applyDate).value = a.applyDate;
    if (C.rpaTime>0)   row.getCell(C.rpaTime).value   = now;
    row.commit(); nextRow++;
  }
  await workbook.xlsx.writeFile(PATHS.targetExcel);
  console.log('  Excel更新完了 (' + all.length + '件追記)');
  const destDir = path.join(PATHS.copyDestDir, todayFolderName());
  safeCopy(PATHS.targetExcel, destDir, path.basename(PATHS.targetExcel));
}

(async () => {
  const browser = await chromium.launch({ headless: false });
  try {
    const airworkCsv = await downloadManual(browser, AIRWORK_URL, PATHS.airworkSession, 'AirWORK');
    const dodaCsv    = await downloadManual(browser, DODA_URL,    PATHS.dodaSession,    'doda');
    await writeToExcel(airworkCsv, dodaCsv);
    console.log('');
    console.log('すべての処理が完了しました。');
  } catch (err) {
    console.error('エラー:', err.message);
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
