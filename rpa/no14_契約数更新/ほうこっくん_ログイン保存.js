// ほうこっくんに手動ログインしてセッションを保存するスクリプト
// （Googleスプシ_ログイン保存.js と同じ方式）
//
// 実行: node .\ほうこっくん_ログイン保存.js
const { chromium } = require("playwright");
const readline = require("readline");

const config = require("./config.json");

function waitForEnter(message) {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(message, () => {
      rl.close();
      resolve();
    });
  });
}

(async () => {
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();

  // ほうこっくんは http://…/top/ へリダイレクトするがポート80が閉じているため、
  // http を https に強制的に置き換える（通常ブラウザの自動https格上げ相当）
  await context.route("http://houkoku.access-mgr.biz/**", (route) => {
    const url = route.request().url().replace(/^http:/, "https:");
    return route.fulfill({ status: 302, headers: { location: url } });
  });

  const page = await context.newPage();

  await page.goto(config.houkokkun.url);

  await waitForEnter("ブラウザ上でほうこっくんに手動ログインし、メニュー画面が開けたことを確認したら、ここでEnterキーを押してください...");

  await context.storageState({ path: config.houkokkun.sessionFile });
  console.log(`セッションを保存しました：${config.houkokkun.sessionFile}`);

  await browser.close();
})();
