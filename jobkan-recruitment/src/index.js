const fs = require('fs');
const path = require('path');
const { chromium } = require('playwright');
const config = require('./config');
const { downloadAirworkApplicants } = require('./airwork');
const { downloadDodaApplicants } = require('./doda');
const { reflectAirworkAndDoda, reflectApplicantCsvFolder } = require('./excel');
const { copyReflectExcelToSharedDrive } = require('./copyToSharedDrive');

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

async function withDebugOnError(page, label, fn) {
  try {
    return await fn();
  } catch (err) {
    const debugDir = path.join(__dirname, '..', 'work', 'error-screenshots');
    ensureDir(debugDir);
    const shotPath = path.join(debugDir, `${label}-${Date.now()}.png`);
    try {
      await page.screenshot({ path: shotPath, fullPage: true });
      console.error(`[${label}] エラー発生。スクリーンショットを保存しました: ${shotPath}`);
    } catch (shotErr) {
      console.error(`[${label}] スクリーンショット保存にも失敗しました: ${shotErr.message}`);
    }
    throw err;
  }
}

async function main() {
  ensureDir(config.applicantListDir);

  const browser = await chromium.launch({
    headless: config.headless,
    slowMo: config.slowMoMs,
  });

  try {
    // --- AirWORK（手順1〜7） ---
    const airworkContext = await browser.newContext({ acceptDownloads: true });
    // サイト側の反応が遅い日があるため、個別のlocatorごとにtimeoutを指定するのではなく
    // コンテキート全体の既定タイムアウトを底上げしておく
    airworkContext.setDefaultTimeout(60000);
    const airworkPage = await airworkContext.newPage();
    await withDebugOnError(airworkPage, 'airwork', () =>
      downloadAirworkApplicants(airworkPage)
    );
    await airworkContext.close();

    // --- doda（手順8〜17） ---
    const dodaContext = await browser.newContext({ acceptDownloads: true });
    dodaContext.setDefaultTimeout(60000);
    const dodaPage = await dodaContext.newPage();
    await withDebugOnError(dodaPage, 'doda', () => downloadDodaApplicants(dodaPage));
    await dodaContext.close();

    console.log('CSVダウンロードが完了しました（手順1〜17）。');

    // --- 反映用リスト.xlsxへの転記（手順19〜20） ---
    await reflectAirworkAndDoda();

    // --- 応募者数CSVフォルダからの転記（手順21） ---
    await reflectApplicantCsvFolder();

    // --- 共有ドライブへのコピー（手順22） ---
    await copyReflectExcelToSharedDrive();
  } finally {
    await browser.close();
  }
}

main().catch((err) => {
  console.error('処理中にエラーが発生しました:', err);
  process.exitCode = 1;
});
