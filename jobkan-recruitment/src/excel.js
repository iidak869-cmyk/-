const path = require('path');
const ExcelJS = require('exceljs');
const config = require('./config');
const { parseAirworkCsv } = require('./parsers/airworkCsv');
const { parseDodaCsv } = require('./parsers/dodaCsv');
const {
  parseApplicantExportFile,
  findLatestDatedSubfolder,
  listApplicantExportFiles,
} = require('./parsers/applicantCsvFolder');

const ONE_DAY_MS = 24 * 60 * 60 * 1000;

// 手順19,20: 厳密な「前日」判定はせず、実行時刻から遡って24時間以内に
// 応募があったものを抽出する。毎日同じ間隔で実行し続ける限り、これで
// 「前日分」が漏れなく（重複なく）取得できる。
function filterWithinLast24Hours(records, referenceDate = new Date()) {
  const cutoff = new Date(referenceDate.getTime() - ONE_DAY_MS);
  return records.filter(
    (r) => r.applicationDate && r.applicationDate >= cutoff && r.applicationDate <= referenceDate
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
 * 手順19,20: AirWORK/dodaのCSVから直近24時間以内の応募者を抽出し、反映用リスト.xlsxに追記する
 */
async function reflectAirworkAndDoda({
  airworkCsvPath = path.join(config.applicantListDir, '応募者リスト(AirWORK).csv'),
  dodaCsvPath = path.join(config.applicantListDir, '応募者リスト(doda).csv'),
  excelPath = config.reflectExcelPath,
  referenceDate = new Date(),
} = {}) {
  const airworkRecords = filterWithinLast24Hours(parseAirworkCsv(airworkCsvPath), referenceDate);
  const dodaRecords = filterWithinLast24Hours(parseDodaCsv(dodaCsvPath), referenceDate);

  const airworkCount = await appendApplicantsToReflectExcel(excelPath, airworkRecords);
  console.log(`[反映] AirWORKの直近24時間の応募者 ${airworkCount}件を追記しました`);

  const dodaCount = await appendApplicantsToReflectExcel(excelPath, dodaRecords);
  console.log(`[反映] dodaの直近24時間の応募者 ${dodaCount}件を追記しました`);

  return { airworkCount, dodaCount };
}

/**
 * 手順21: 応募者数CSVフォルダ内にある「現在日時のフォルダ」内の各ファイル
 * （LINE/Instagram等、媒体ごとのマーケティングツールからのエクスポート）から
 * 直近24時間以内の応募者を抽出し、反映用リスト.xlsxに追記する。
 *
 * TODO(要確認): サブフォルダの命名規則が未確認のため、現状は
 * 「応募者数CSVフォルダ直下で更新日時が最も新しいサブフォルダ」を対象にしている。
 */
async function reflectApplicantCsvFolder({
  rootDir = config.applicantCsvRootDir,
  excelPath = config.reflectExcelPath,
  referenceDate = new Date(),
} = {}) {
  const targetDir = findLatestDatedSubfolder(rootDir);
  if (!targetDir) {
    console.log(`[反映] 応募者数CSVフォルダに対象サブフォルダが見つかりませんでした: ${rootDir}`);
    return { fileCount: 0, recordCount: 0 };
  }

  const files = listApplicantExportFiles(targetDir);
  let recordCount = 0;

  for (const filePath of files) {
    const records = filterWithinLast24Hours(
      await parseApplicantExportFile(filePath),
      referenceDate
    );
    const count = await appendApplicantsToReflectExcel(excelPath, records);
    recordCount += count;
    console.log(`[反映] ${path.basename(filePath)} の直近24時間の応募者 ${count}件を追記しました`);
  }

  return { fileCount: files.length, recordCount };
}

module.exports = {
  filterWithinLast24Hours,
  appendApplicantsToReflectExcel,
  reflectAirworkAndDoda,
  reflectApplicantCsvFolder,
};
