const path = require('path');
const config = require('./config');

function formatToday() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  // TODO(要確認): 画面側の日付フォーマットが YYYY/MM/DD 以外の場合はここを調整
  return `${y}/${m}/${d}`;
}

/**
 * doda: ログイン → 求人情報検索 → 応募者情報CSV出力 → ログアウト（手順8〜17）
 *
 * 手順15〜17（ダウンロードフォルダから取得→リネーム→移動）は、Playwrightの
 * download イベントを直接 saveAs() で目的のパス（応募者リスト(doda).csv）に
 * 保存することで一手順にまとめている。
 *
 * NOTE: airwork.js と同様、locator は実画面未確認の仮のものです（TODO(要確認)）。
 * HEADLESS=false で実行し、実際の画面に合わせて調整してください。
 */
async function downloadDodaApplicants(page) {
  console.log('[doda] ログインページへ移動します');
  // 'load'（既定値）はSSOのリダイレクトページ等で完了しないことがあるため
  // 'domcontentloaded' を使う。要素の待機は各locatorの自動待機に任せる。
  await page.goto(config.doda.loginUrl, { waitUntil: 'domcontentloaded' });

  // 実画面で確認済み: doda CONNECTのログイン画面は「メールアドレス」「パスワード」
  // ラベル、ボタンは「同意してログイン」
  await page.getByLabel('メールアドレス').fill(config.doda.id());
  // exact指定が無いと「パスワードを表示する」ボタンにも部分一致してしまうため exact: true
  await page.getByLabel('パスワード', { exact: true }).fill(config.doda.password());
  await page.getByRole('button', { name: '同意してログイン' }).click();
  console.log('[doda] ログイン完了');

  // 手順9: 「利用サービス」一覧内の「doda 求人情報」を選択
  // 実画面で確認済み: カード全体がリンクになっており、アクセシブルネームが
  // 説明文まで含んで長くなるため、見出しテキストで直接指定する
  await page.getByText('doda 求人情報', { exact: true }).click();
  console.log('[doda] 求人情報ページへ遷移しました');

  // 手順10: 日付を両方とも当日に変更
  const today = formatToday();
  // TODO(要確認): 実際の日付入力欄が2つとも同じlabel/placeholderかどうか要確認
  const dateInputs = page.locator('input[type="text"][name*="date" i], input[placeholder*="日付"]');
  await dateInputs.nth(0).fill(today);
  await dateInputs.nth(1).fill(today);
  console.log(`[doda] 日付を本日(${today})に設定しました`);

  // 手順11: 「この条件で検索」をクリック
  await page.getByRole('button', { name: 'この条件で検索' }).click();
  console.log('[doda] 検索を実行しました');

  // 手順12: 検索結果上部のチェックボックスを選択（全選択チェックボックスを想定）
  // TODO(要確認): 「全選択」チェックボックスの実際の位置・name属性を要確認
  await page.locator('thead input[type="checkbox"], th input[type="checkbox"]').first().check();

  // 手順13: プルダウンを「応募者情報をCSV出力する」にして実行
  await page.getByRole('combobox').selectOption({ label: '応募者情報をCSV出力する' });

  const [download] = await Promise.all([
    page.waitForEvent('download'),
    // TODO(要確認): 「実行」ボタンの実際の名称・位置を要確認
    page.getByRole('button', { name: '実行' }).click(),
  ]);

  // 手順15〜17: 目的のファイル名・格納先に直接保存
  const targetPath = path.join(config.applicantListDir, '応募者リスト(doda).csv');
  await download.saveAs(targetPath);
  console.log(`[doda] CSVを保存しました: ${targetPath}`);

  // 手順14: ログアウト
  // TODO(要確認): ログアウト導線は実画面で要確認
  await page.getByRole('button', { name: /アカウント|メニュー/ }).click();
  await page.getByRole('link', { name: 'ログアウト' }).click();
  console.log('[doda] ログアウトしました');

  return targetPath;
}

module.exports = { downloadDodaApplicants };
