const path = require('path');
const config = require('./config');

/**
 * AirWORK: ログイン → 応募者一覧CSVダウンロード → ログアウト（手順1〜7）
 *
 * 手順5〜7（ダウンロードフォルダから取得→リネーム→移動）は、Playwrightの
 * download イベントを直接 saveAs() で目的のパス（応募者リスト(AirWORK).csv）に
 * 保存することで一手順にまとめている。動作結果は手順通り。
 *
 * NOTE: 下記の locator は実画面を見ずに書いた仮のものです（TODO(要確認)の箇所）。
 * 初回はHEADLESS=falseで実行し、画面を見ながら selector を実際のものに合わせてください。
 * ズレている箇所は `npx playwright codegen <URL>` で実際の操作を記録し、
 * 出力されたコードの locator をここに置き換えるのが最も確実です。
 */
async function downloadAirworkApplicants(page) {
  console.log('[AirWORK] ログインページへ移動します');
  await page.goto(config.airwork.loginUrl);

  // 実画面で確認済み: AirIDログイン画面はlabelが無くplaceholderのみで入力欄を識別する
  await page.getByPlaceholder('AirIDまたはメールアドレス').fill(config.airwork.id());
  await page.getByPlaceholder('パスワード').fill(config.airwork.password());
  await page.getByRole('button', { name: 'ログイン' }).click();

  await page.waitForLoadState('networkidle');
  console.log('[AirWORK] ログイン完了');

  // 手順2: メニューバーの「応募者」を選択
  // 実画面で確認済み: exact指定が無いと「応募者を管理する」等の他リンクにも
  // 部分一致してしまうため、exact: true で完全一致のメニューリンクに絞る
  await page.getByRole('link', { name: '応募者', exact: true }).click();
  await page.waitForLoadState('networkidle');
  console.log('[AirWORK] 応募者ページへ遷移しました');

  // 手順3: 最下部の「応募者一覧をダウンロード」ボタンを選択し、ダウンロードを捕捉
  const downloadButton = page.getByRole('button', { name: '応募者一覧をダウンロード' });
  await downloadButton.scrollIntoViewIfNeeded();

  const [download] = await Promise.all([
    page.waitForEvent('download'),
    downloadButton.click(),
  ]);

  // 手順5〜7: ダウンロードフォルダ経由ではなく、直接目的のファイル名・格納先に保存
  const targetPath = path.join(config.applicantListDir, '応募者リスト(AirWORK).csv');
  await download.saveAs(targetPath);
  console.log(`[AirWORK] CSVを保存しました: ${targetPath}`);

  // 手順4: ログアウト
  // TODO(要確認): ログアウト導線（ヘッダーのアカウントメニュー配下など）は実画面で要確認
  await page.getByRole('button', { name: /アカウント|メニュー/ }).click();
  await page.getByRole('link', { name: 'ログアウト' }).click();
  console.log('[AirWORK] ログアウトしました');

  return targetPath;
}

module.exports = { downloadAirworkApplicants };
