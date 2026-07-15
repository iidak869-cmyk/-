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
  const page = await context.newPage();

  await page.goto(config.houkokkun.url);

  await waitForEnter("ブラウザ上でほうこっくんに手動ログインし、メニュー画面が開けたことを確認したら、ここでEnterキーを押してください...");

  await context.storageState({ path: config.houkokkun.sessionFile });
  console.log(`セッションを保存しました：${config.houkokkun.sessionFile}`);

  await browser.close();
})();
