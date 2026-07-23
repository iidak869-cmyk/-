const path = require('path');
const ExcelJS = require('exceljs');
const config = require('./config');
const { parseAirworkCsv } = require('./parsers/airworkCsv');
const { parseDodaCsv } = require('./parsers/dodaCsv');

function isSameDate(a, b) {
  return (
    a.getFullYear() === b.getFullYear() &&
    a.getMonth() === b.getMonth() &&
    a.getDate() === b.getDate()
  );
}

function getYesterday(referenceDate) {
  const d = new Date(referenceDate);
  d.setDate(d.getDate() - 1);
  return d;
}

// 手順19,20: 「前日」＝スクリプト実行日の1日前、として判定する
// TODO(要確認): 「前日」の定義（実行日の1日前で確定でよいか）は要確認
function filterPreviousDay(records, referenceDate = new Date()) {
  const yesterday = getYesterday(referenceDate);
  return records.filter(
    (r) => r.applicationDate && isSameDate(r.applicationDate, yesterday)
  );
}

function formatDate(d) {
  if (!d) return '';
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${y}/${m}/${day}`;
}

function formatDateTime(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  const ss = String(d.getSeconds()).padStart(2, '0');
  return `${y}/${m}/${day} ${hh}:${mm}:${ss}`;
}

// A列（氏名）に値がある最後の行番号を返す。
// worksheet.addRow()は書式だけが設定された空行もカウントしてしまい、実データの
// 末尾より後ろに追記されることがあるため、実データの有無で判定する。
function findLastDataRow(worksheet, keyColumnIndex = 1) {
  let lastRow = 1; // ヘッダー行
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (row.getCell(keyColumnIndex).value) {
      lastRow = rowNumber;
    }
  });
  return lastRow;
}

// 反映用リスト.xlsx の列順: 氏名,氏名（かな）,性別,メールアドレス,電話番号,
// お電話希望時間,希望勤務地,経路,入社種類,応募日,RPA稼働時刻
async function appendApplicantsToReflectExcel(excelPath, records) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelPath);
  const worksheet = workbook.worksheets[0];

  const rpaTimestamp = formatDateTime(new Date());
  let nextRowNumber = findLastDataRow(worksheet) + 1;

  records.forEach((r) => {
    worksheet.getRow(nextRowNumber).values = [
      r.name,
      r.kana,
      r.gender,
      r.email,
      r.phone,
      r.phoneTime,
      r.workLocation,
      r.route,
      r.employmentType,
      formatDate(r.applicationDate),
      rpaTimestamp,
    ];
    nextRowNumber += 1;
  });

  await workbook.xlsx.writeFile(excelPath);
  return records.length;
}

/**
 * 手順19,20: AirWORK/dodaのCSVから前日応募者を抽出し、反映用リスト.xlsxに追記する
 */
async function reflectAirworkAndDoda({
  airworkCsvPath = path.join(config.applicantListDir, '応募者リスト(AirWORK).csv'),
  dodaCsvPath = path.join(config.applicantListDir, '応募者リスト(doda).csv'),
  excelPath = config.reflectExcelPath,
  referenceDate = new Date(),
} = {}) {
  const airworkRecords = filterPreviousDay(parseAirworkCsv(airworkCsvPath), referenceDate);
  const dodaRecords = filterPreviousDay(parseDodaCsv(dodaCsvPath), referenceDate);

  const airworkCount = await appendApplicantsToReflectExcel(excelPath, airworkRecords);
  console.log(`[反映] AirWORKの前日応募者 ${airworkCount}件を追記しました`);

  const dodaCount = await appendApplicantsToReflectExcel(excelPath, dodaRecords);
  console.log(`[反映] dodaの前日応募者 ${dodaCount}件を追記しました`);

  return { airworkCount, dodaCount };
}

/**
 * 手順21: 応募者数CSVフォルダ内の各Excelファイルから前日応募者を抽出し、反映用リスト.xlsxに追記する
 * TODO: サンプルファイル受領後に実装予定（未実装）
 */
async function reflectApplicantCsvFolder() {
  throw new Error('reflectApplicantCsvFolder is not implemented yet (サンプルファイル待ち)');
}

module.exports = {
  filterPreviousDay,
  appendApplicantsToReflectExcel,
  reflectAirworkAndDoda,
  reflectApplicantCsvFolder,
};
