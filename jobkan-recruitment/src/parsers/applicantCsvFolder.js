const fs = require('fs');
const path = require('path');
const iconv = require('iconv-lite');
const ExcelJS = require('exceljs');
const { parse } = require('csv-parse/sync');

// 応募者数CSVフォルダ内の各ファイル（媒体ごとの応募者エクスポート。LINE/Instagram等の
// マーケティングツールからの出力で、AirWORK/dodaとは別形式）で使う固定の列名。
// 「入社種類」「希望勤務地」「お電話希望時間」はツール上のカスタム項目のため、
// ファイルによって存在しない/位置が違うことがある前提で、名前で検索する。
const FIELD = {
  NAME: 'お名前',
  KANA: 'フリガナ',
  EMAIL: 'メールアドレス',
  PHONE: '電話番号',
  PHONE_TIME: 'お電話希望時間',
  WORK_LOCATION: '希望勤務地',
  EMPLOYMENT_TYPE: '入社種類',
  REGISTERED_AT: '登録日',
  SCENARIO: 'シナリオ名（購入商品）',
};

function buildHeaderIndex(header) {
  const map = {};
  header.forEach((h, i) => {
    if (h && !(h in map)) map[h] = i; // 同名ヘッダーが複数ある場合は先頭を優先
  });
  return map;
}

function getField(row, headerIndex, name) {
  const i = headerIndex[name];
  return i === undefined ? '' : row[i] || '';
}

function parseRegisteredAt(value) {
  // "2026-07-03 12:03:48" -> Date
  const m = (value || '').match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/);
  if (!m) return null;
  const [, y, mo, d, hh, mm, ss] = m.map(Number);
  return new Date(y, mo - 1, d, hh, mm, ss);
}

// 「経路」はシナリオ名（例: "★稼働中【WeeAre公式】Instagram"）の末尾（最後の "】" より後ろ）を
// 媒体名として抽出する。カッコ内のタグ（"WeeAre公式" "えの" 等）が異なる複数サンプルで確認済み。
function extractRouteFromScenario(scenarioName, fallback) {
  if (!scenarioName) return fallback;
  const idx = scenarioName.lastIndexOf('】');
  const extracted = idx >= 0 ? scenarioName.slice(idx + 1).trim() : scenarioName.trim();
  return extracted || fallback;
}

function rowsToRecords(header, rows, fallbackRoute) {
  const headerIndex = buildHeaderIndex(header);
  return rows
    .filter((row) => row.some((cell) => cell && cell.trim()))
    .map((row) => ({
      route: extractRouteFromScenario(getField(row, headerIndex, FIELD.SCENARIO), fallbackRoute),
      name: getField(row, headerIndex, FIELD.NAME),
      kana: getField(row, headerIndex, FIELD.KANA),
      gender: '', // このツールのエクスポートには性別項目が無いため常に空欄
      email: getField(row, headerIndex, FIELD.EMAIL),
      phone: getField(row, headerIndex, FIELD.PHONE),
      phoneTime: getField(row, headerIndex, FIELD.PHONE_TIME),
      workLocation: getField(row, headerIndex, FIELD.WORK_LOCATION),
      employmentType: getField(row, headerIndex, FIELD.EMPLOYMENT_TYPE),
      applicationDate: parseRegisteredAt(getField(row, headerIndex, FIELD.REGISTERED_AT)),
    }));
}

function parseCsvFile(filePath, fallbackRoute) {
  const buffer = fs.readFileSync(filePath);
  // BOM付きUTF-8とCP932(Shift_JIS)の両方に対応
  const isUtf8Bom = buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
  const raw = isUtf8Bom ? buffer.toString('utf8').replace(/^﻿/, '') : iconv.decode(buffer, 'cp932');
  const rows = parse(raw, { skip_empty_lines: true });
  const [header, ...dataRows] = rows;
  return rowsToRecords(header, dataRows, fallbackRoute);
}

async function parseExcelFile(filePath, fallbackRoute) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];
  const rows = [];
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    // exceljsの row.values は1始まりの配列で先頭がnull
    rows.push(row.values.slice(1).map((v) => (v == null ? '' : String(v))));
  });
  const [header, ...dataRows] = rows;
  return rowsToRecords(header, dataRows, fallbackRoute);
}

/**
 * 応募者数CSVフォルダ内の1ファイル（.csv または .xlsx）を読み込み、
 * 正規化した応募者レコードの配列を返す。
 */
async function parseApplicantExportFile(filePath) {
  const fallbackRoute = path.basename(filePath, path.extname(filePath));
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.xlsx' || ext === '.xlsm') {
    return parseExcelFile(filePath, fallbackRoute);
  }
  return parseCsvFile(filePath, fallbackRoute);
}

/**
 * folderDir 直下にある「現在日時のフォルダ」を1つ選ぶ。
 * TODO(要確認): サブフォルダの命名規則が未確認のため、現状は
 * 「更新日時が最も新しいサブフォルダ」を採用している。
 */
function findLatestDatedSubfolder(rootDir) {
  if (!fs.existsSync(rootDir)) return null;
  const entries = fs
    .readdirSync(rootDir, { withFileTypes: true })
    .filter((e) => e.isDirectory())
    .map((e) => {
      const fullPath = path.join(rootDir, e.name);
      return { fullPath, mtime: fs.statSync(fullPath).mtimeMs };
    })
    .sort((a, b) => b.mtime - a.mtime);
  return entries.length > 0 ? entries[0].fullPath : null;
}

/**
 * 対象フォルダ内の応募者エクスポートファイル（.csv/.xlsx/.xlsm）一覧を返す。
 */
function listApplicantExportFiles(folderDir) {
  if (!fs.existsSync(folderDir)) return [];
  return fs
    .readdirSync(folderDir)
    .filter((name) => ['.csv', '.xlsx', '.xlsm'].includes(path.extname(name).toLowerCase()))
    .map((name) => path.join(folderDir, name));
}

module.exports = {
  parseApplicantExportFile,
  findLatestDatedSubfolder,
  listApplicantExportFiles,
};
