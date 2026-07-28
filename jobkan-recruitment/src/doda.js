const path = require('path');
const config = require('./config');

/**
 * doda: ログイン → 求人情報検索 → 応募者情報CSV出力 → ログアウト（手順8〜17）
 *
 * 手順15〜17（ダウンロードフォルダから取得→リネーム→移動）は、Playwrightの
 * download イベントを直接 saveAs() で目的のパス（応募者リスト(doda).csv）に
 * 保存することで一手順にまとめている。
 *
 * NOTE: 手順9〜13は `npx playwright codegen` で実際の操作を記録し、その結果に
 * 基づいて実装している（doda CONNECTは独自のカレンダー/チェックボックス/
 * ドロップダウンウィジェットを使っており、ネイティブ要素ではない）。
 * 手順14（ログアウト）のみ未検証（TODO(要確認)）。
 *
 * 本日の応募が0件の日もあり得るため、その場合はCSV出力をスキップして
 * null を返す（ログアウトは行う）。
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

  // 手順9: 「doda 求人情報」サービスを選択
  // 実画面のcodegen記録で確認済み: サービス選択パネルをクリックした後、
  // 「求人情報サービス」という画像リンクをクリックする2段階の操作。
  // ログイン直後は #loadingOverlay がクリックを妨げ、リトライ+遷移待ちの
  // 合計がデフォルトの30秒を超えることがあるため、オーバーレイが消えるのを
  // 待ってからクリックし、タイムアウトも伸ばす。
  await page
    .locator('#loadingOverlay')
    .waitFor({ state: 'hidden', timeout: 30000 })
    .catch(() => {});
  await page
    .locator(
      '.servicePanel_box.servicePanel_box_SelectServiceMode_Panel_Mode1 > .servicePanel_box_flame > .servicePanel_detail'
    )
    .click({ timeout: 60000 });
  await page.getByRole('img', { name: '求人情報サービス' }).click({ timeout: 60000 });
  console.log('[doda] 求人情報ページへ遷移しました');

  // 手順10: 日付を両方とも当日に変更
  // 実画面のcodegen記録で確認済み: 独自のカレンダーウィジェット。
  // 入力欄をクリックするとカレンダーが開き、当日の日付リンクをクリックする。
  // TODO(要確認): 1桁の日付(1〜9日)がカレンダー上でゼロ埋め表示されるかは未確認
  const todayDay = String(new Date().getDate());
  await page.locator('#applicationFromDateInputId').click();
  await page.getByRole('link', { name: todayDay, exact: true }).click();
  await page.locator('#applicationToDateInputId').click();
  await page.getByRole('link', { name: todayDay, exact: true }).click();
  console.log(`[doda] 日付を本日(${todayDay}日)に設定しました`);

  // 手順11: 「この条件で検索」をクリック（buttonではなくlink）
  await page.getByRole('link', { name: 'この条件で検索' }).click();
  console.log('[doda] 検索を実行しました');

  // 本日の応募が0件の日もあり得るため、件数表示（例:「0件中 0件を表示」）で判定する。
  // ヘッダーのチェックボックスは0件でも常に表示されてしまうため判定材料にできない。
  const countLocator = page.getByText(/\d+件中\s*\d+件を表示/).first();
  await countLocator.waitFor({ state: 'visible', timeout: 20000 });
  const countText = ((await countLocator.textContent()) || '').trim();
  const isEmpty = /^0件中/.test(countText);

  let targetPath = null;

  if (isEmpty) {
    console.log(`[doda] 本日分の応募者は0件でした（表示: ${countText}）。CSV出力をスキップします。`);
  } else {
    // 手順12: 検索結果上部のチェックボックスを選択
    // 実画面で確認済み: Infragistics製の独自チェックボックスウィジェット
    await page.locator('.ui-igcheckbox-normal-off').first().click();

    // 手順13: プルダウンを「応募者情報をCSV出力する」にして実行
    // 実画面で確認済み: ネイティブのselectではなく、クリックで開く独自ドロップダウン
    await page.getByText('一括操作 キャリアシートを見る 応募者情報をCSV').first().click();
    await page.getByText('応募者情報をCSV出力する').first().click();

    const [download] = await Promise.all([
      page.waitForEvent('download'),
      // 「実行」もbuttonではなくlink
      page.getByRole('link', { name: '実行' }).first().click(),
    ]);

    // 手順15〜17: 目的のファイル名・格納先に直接保存
    targetPath = path.join(config.applicantListDir, '応募者リスト(doda).csv');
    await download.saveAs(targetPath);
    console.log(`[doda] CSVを保存しました: ${targetPath}`);
  }

  // 手順14: ログアウト
  // 実画面で確認済み: 右上の会社名（株式会社Weプラス）をクリックするとメニューが開き、
  // 「ログアウト」が表示される。button/link roleではない可能性があるためテキストで指定。
  // 全角/半角の表記ゆれに対応するため正規表現で「We」部分をワイルドカードにする
  await page.getByText(/株式会社.*プラス/).first().click();
  await page.getByText('ログアウト').first().click();
  console.log('[doda] ログアウトしました');

  return targetPath;
}

module.exports = { downloadDodaApplicants };
