// No.14 契約数更新（完成版）
//
// ほうこっくんの「前日(id=-1)」「当日(id=0)」×「新規/リピート」の4パターンを検索し、
// 営業報告の詳細から項目を抽出して プロダクト日報Excel へ転記（重複チェック付き）、
// 追記分をテスト用スプシへ自動貼り付けする。
//
// 実行:
//   node .\契約数更新.js            … 本番実行
//   node .\契約数更新.js --dry-run  … Excel・スプシに書き込まず、抽出結果の表示だけ行う
//   node .\契約数更新.js --探索     … 画面を開いてPlaywright Inspectorを起動（調査用）
//
// 補足:
//   ・「7/15の報告一覧」には 営業日付7/15の契約 と 7/15に報告された前日営業分 が混在する仕様。
//     前日と当日の両方を回し、契約日+会社名の重複チェックで二重転記を防ぐ（手作業と同じ運用）。

const { chromium } = require("playwright");
const fs = require("fs");
const ExcelJS = require("exceljs");

const config = require("./config.json");

const BASE = "https://houkoku.access-mgr.biz";

// 「入力」シートの書き込み先（列レター → 転記する項目）。他の列は空欄のまま
const COLUMN_MAP = {
  A: "契約日",
  C: "顧客名",
  D: "営業担当者",
  E: "営業部署",
  F: "会社所在地",
  H: "業種",
  J: "画像",            // 添付画像なしのとき「無」、あるときは空欄
  M: "金額",            // 月額（グロス）を数値で
  O: "クレカの有無",    // 報告内容の「クレカ(所持/未所持)」から 有/無
  AD: "挨拶日",         // 報告内容の「挨拶zoom 日時」の日付
  AE: "挨拶時間",       // 同・時間
  CH: "納品物件",       // Cyteki / DegiOne
};

// 列レター → 列番号 (A=1)
function colNum(letter) {
  let n = 0;
  for (const ch of letter) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

const isExplore = process.argv.includes("--探索");
const isDryRun = process.argv.includes("--dry-run");

// ほうこっくんにログイン済みの状態を保証する。
// セッションが切れていたら config.json の id / password で自動ログインする
async function ensureLoggedIn(page) {
  await page.goto(config.houkokkun.topUrl || config.houkokkun.url, { waitUntil: "domcontentloaded" });
  await page.waitForLoadState("domcontentloaded");

  const needLogin =
    page.url().includes("login.php") ||
    (await page.locator('input[placeholder="PASSWORD"]').count()) > 0;
  if (!needLogin) return;

  const { id, password } = config.houkokkun;
  if (!id || !password) {
    throw new Error(
      "ほうこっくんのセッションが切れています。config.json の houkokkun.id / houkokkun.password を設定してください"
    );
  }
  await page.fill('input[placeholder="ID"]', id);
  await page.fill('input[placeholder="PASSWORD"]', password);

  // ボタンクリックだと http://…/top/ へのリダイレクト追跡でポート80に接続して固まるため、
  // fetch(redirect: manual) でPOSTだけ送り、その後httpsのトップへ自分で移動する
  await page.evaluate(async () => {
    const form = document.querySelector("form");
    const data = new FormData(form);
    for (const btn of form.querySelectorAll('input[type="submit"]')) {
      if (btn.name) data.append(btn.name, btn.value);
    }
    await fetch(form.getAttribute("action") || location.href, {
      method: "POST",
      body: data,
      redirect: "manual",
    });
  });

  await page.goto(config.houkokkun.topUrl || config.houkokkun.url, { waitUntil: "domcontentloaded" });
  await page.waitForLoadState("domcontentloaded");

  if (page.url().includes("login.php")) {
    throw new Error("ほうこっくんへのログインに失敗しました。ID / PASSWORD を確認してください");
  }
}

// 検索直後にその一覧から詳細を開いて抽出する。
// ほうこっくんは「直前に表示した一覧に載っている報告」しか詳細を開かせない
// （それ以外は「不正なアクセスです」になる）ため、検索→即詳細の順で処理する
async function collectAndExtract(page) {
  const seen = new Set();
  const records = [];

  for (const dayOffset of [-1, 0]) {
    const dayLabel = dayOffset === -1 ? "前日" : "当日";
    for (const type of ["新規", "リピート"]) {
      const value = type === "新規" ? "s,1" : "s,3";
      const listUrl = `${BASE}/top/index.php?id=${dayOffset}`;

      await page.goto(listUrl, { waitUntil: "domcontentloaded" });
      await page.click('label[for="tab1_1"]'); // 営業タブ
      await page.selectOption('select[name="sales_type"]', value);
      await page.click('input[name="sales_search"]');
      await page.waitForLoadState("domcontentloaded");
      await page.waitForTimeout(800);

      const rows = await page.$$eval('tr[data-href*="sales_product_reportdetail"]', (trs) =>
        trs.map((tr) => {
          const tds = [...tr.querySelectorAll("td")].map((td) =>
            td.textContent.replace(/\s+/g, " ").trim()
          );
          return {
            href: tr.dataset.href,
            更新日時: tds[0],
            契約日: tds[1], // 営業日付
            報告種別: tds[2],
            営業担当: tds[3],
            アポ担当: tds[4],
            金額: tds[5],
            会社名: tds[6],
          };
        })
      );
      console.log(`  ${dayLabel}×${type}: ${rows.length}件`);

      const targets = [];
      for (const row of rows) {
        if (row.報告種別 !== type) continue; // 絞り込み前の行が混ざった場合の保険
        const id = (row.href.match(/id=(\d+)/) || [])[1];
        if (!id || seen.has(id)) continue; // 前日/当日の一覧に同じ報告が出るケースを除去
        seen.add(id);
        targets.push({ ...row, id, 区分: `${dayLabel}×${type}` });
      }

      // この一覧に載っている報告の詳細を、一覧の状態が生きているうちに開く
      for (const report of targets) {
        records.push(await extractDetail(page, report, listUrl));
      }
    }
  }
  return records;
}

// 営業報告の詳細ページから転記項目を抽出する
async function extractDetail(page, report, listUrl) {
  const detailUrl = new URL(report.href, `${BASE}/top/`).href;

  for (let attempt = 1; attempt <= 2; attempt++) {
    await page.goto(detailUrl, { waitUntil: "domcontentloaded", referer: listUrl });
    await page.waitForSelector("main table tr th", { timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(300);

    // 「不正なアクセスです」と拒否されたら、一覧を開き直して1回だけ再試行
    if ((await page.locator("text=不正なアクセス").count()) > 0 && attempt === 1) {
      console.log(`  id=${report.id}: 不正なアクセス表示 → 一覧を開き直して再試行`);
      await page.goto(listUrl, { waitUntil: "domcontentloaded" });
      await page.waitForTimeout(500);
      continue;
    }
    break;
  }

  // 【ラベル】→ 値 のマップを作る（画面のth/td構造を利用）
  const map = await page.$$eval("main table tr", (trs) => {
    const m = {};
    for (const tr of trs) {
      const th = tr.querySelector("th");
      const td = tr.querySelector("td");
      if (!th || !td) continue;
      const key = th.textContent.replace(/[【】\s]/g, "");
      if (key && !(key in m)) m[key] = td.textContent.replace(/\s+/g, " ").trim();
    }
    return m;
  });

  console.log(`  詳細取得: id=${report.id} 項目数=${Object.keys(map).length} url=${page.url()}`);

  // 調査用: 項目数0が起きた最初の1件だけ、実際に見えているHTMLを保存する
  if (!extractDetail.dumped && Object.keys(map).length === 0) {
    extractDetail.dumped = true;
    const dumpPath =
      "G:\\共有ドライブ\\RPA運用-技術課連携\\RPA検証用_テストデータ\\契約数更新\\html_dump\\99_run_detail_dump.html";
    try {
      fs.writeFileSync(dumpPath, await page.content(), "utf8");
      console.log(`  (調査用HTMLを保存しました: ${dumpPath})`);
    } catch (e) {
      console.log(`  (調査用HTMLの保存に失敗: ${e.message})`);
    }
  }

  const kingaku = map["金額（税込）"] || "";
  const grossStr = (kingaku.match(/月額[^¥]*¥\s*([\d,]+)/) || [])[1] || "";
  const gross = grossStr ? Number(grossStr.replace(/,/g, "")) : "";
  const houkoku = map["報告内容"] || "";
  const kureka = (houkoku.match(/クレカ[^】]*】\s*(未所持|所持)/) || [])[1] || "";
  // 報告内容の「【挨拶zoom】日時：7/20 10:00」から日付と時間を取り出す
  const zoom = houkoku.match(/挨拶zoom[^日]*日時[：:]\s*([0-9\/]+)\s*([0-9:：]+)/i) || [];
  const hasImage = (await page.locator('a[data-lightbox="attach"]').count()) > 0;

  return {
    "契約日": report.契約日,
    "顧客名": map["会社名"] || report.会社名,
    "営業担当者": map["営業担当"] || report.営業担当,
    "営業部署": map["営業担当所属部署"] || "",
    "会社所在地": map["会社所在地"] || "",
    "業種": map["業種"] || "",
    "画像": hasImage ? "" : "無",
    "金額": gross,
    "クレカの有無": kureka ? (kureka === "所持" ? "有" : "無") : "",
    "挨拶日": zoom[1] || "",
    "挨拶時間": (zoom[2] || "").replace(/：/g, ":"),
    "納品物件": map["納品物件"] || "",
    "区分": report.区分,
    "報告ID": report.id,
  };
}

// Excelセル値を比較用の文字列にする（日付セルは yyyy/mm/dd に揃える）
function cellStr(v) {
  if (v instanceof Date) {
    const p = (n) => String(n).padStart(2, "0");
    return `${v.getFullYear()}/${p(v.getMonth() + 1)}/${p(v.getDate())}`;
  }
  if (v && typeof v === "object" && "text" in v) return String(v.text).trim();
  return String(v ?? "").trim();
}

// プロダクト日報「入力」シートに追記（COLUMN_MAPの固定列へ、重複チェック付き）
async function appendToExcel(records) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(config.productNippo.excelPath);
  let ws = wb.getWorksheet(config.productNippo.sheetName);
  if (!ws) {
    ws = wb.worksheets.find(
      (w) => w.name.replace(/\s+/g, "") === String(config.productNippo.sheetName).replace(/\s+/g, "")
    );
  }
  if (!ws) {
    throw new Error(
      `シートが見つかりません: ${config.productNippo.sheetName}（このExcelにあるシート: ${wb.worksheets.map((w) => w.name).join(" / ")}）`
    );
  }

  // 既存データから重複チェック用キー（A列:契約日 + C列:顧客名）を収集
  const existingKeys = new Set();
  let lastDataRow = 1;
  ws.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    const key = `${cellStr(row.getCell(colNum("A")).value)}|${cellStr(row.getCell(colNum("C")).value)}`;
    if (key !== "|") {
      existingKeys.add(key);
      lastDataRow = Math.max(lastDataRow, rowNumber);
    }
  });

  const appended = [];
  for (const record of records) {
    const key = `${String(record["契約日"] ?? "").trim()}|${String(record["顧客名"] ?? "").trim()}`;
    if (existingKeys.has(key)) {
      console.log(`  スキップ(転記済み): ${key}`);
      continue;
    }
    lastDataRow += 1;
    const row = ws.getRow(lastDataRow);
    for (const [letter, field] of Object.entries(COLUMN_MAP)) {
      let value = record[field];
      if (value === "" || value === undefined || value === null) continue;
      // 契約日はExcelの日付として書き込む
      if (letter === "A" && /^\d{4}\/\d{1,2}\/\d{1,2}$/.test(String(value))) {
        const [y, m, d] = String(value).split("/").map(Number);
        value = new Date(y, m - 1, d);
      }
      row.getCell(colNum(letter)).value = value;
    }
    row.commit();
    existingKeys.add(key);
    appended.push(record);
  }

  if (appended.length > 0) {
    await wb.xlsx.writeFile(config.productNippo.excelPath);
  }
  return appended;
}

// テスト反映先スプシへ追記分を貼り付け（gsheet_session.json のセッションを使用）
async function appendToSpreadsheet(records) {
  if (records.length === 0) return;

  // 「入力」シートと同じ列位置（A〜CH）に値を置いたTSVを作る
  const width = Math.max(...Object.keys(COLUMN_MAP).map(colNum));
  const tsv = records
    .map((r) => {
      const cells = new Array(width).fill("");
      for (const [letter, field] of Object.entries(COLUMN_MAP)) {
        cells[colNum(letter) - 1] = String(r[field] ?? "").replace(/[\t\r\n]+/g, " ");
      }
      return cells.join("\t");
    })
    .join("\n");

  const browser = await chromium.launch({ headless: false, channel: config.browserChannel || undefined });
  const context = await browser.newContext({ storageState: config.spreadsheet.sessionFile });
  await context.grantPermissions(["clipboard-read", "clipboard-write"]);
  context.setDefaultNavigationTimeout(90000);
  const page = await context.newPage();
  await page.goto(config.spreadsheet.url, { waitUntil: "domcontentloaded" });
  await page.waitForLoadState("domcontentloaded");
  await page.waitForTimeout(3000); // シートの描画待ち

  await page.evaluate((text) => navigator.clipboard.writeText(text), tsv);

  // データ末尾へ移動 → 行頭(A列) → 1行下の空行へ → 貼り付け
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+End");
  await page.keyboard.press("Home");
  await page.keyboard.press("ArrowDown");
  await page.keyboard.press("Control+V");
  await page.waitForTimeout(5000); // 貼り付け＆自動保存待ち

  await browser.close();
  console.log(`スプシへ追記しました: ${config.spreadsheet.url}`);
}

// スプシ貼り付け用CSVを出力（自動貼り付けに失敗した場合の予備 兼 実行ログ）
function writeCsv(records) {
  if (records.length === 0) return;
  const headers = Object.keys(records[0]);
  const lines = [headers.join(",")];
  for (const r of records) {
    lines.push(headers.map((h) => `"${String(r[h] ?? "").replace(/"/g, '""')}"`).join(","));
  }
  fs.writeFileSync(config.spreadsheet.csvOutput, "\uFEFF" + lines.join("\r\n"), "utf8");
  console.log(`追記分CSVを出力しました: ${config.spreadsheet.csvOutput}`);
}

(async () => {
  const browser = await chromium.launch({ headless: false, channel: config.browserChannel || undefined });
  const context = await browser.newContext({
    storageState: fs.existsSync(config.houkokkun.sessionFile) ? config.houkokkun.sessionFile : undefined,
  });

  // ほうこっくんは http://…/top/ へリダイレクトするがポート80が閉じているため、
  // http を https に強制的に置き換える（通常ブラウザの自動https格上げ相当）
  await context.route("http://houkoku.access-mgr.biz/**", (route) => {
    const url = route.request().url().replace(/^http:/, "https:");
    return route.fulfill({ status: 302, headers: { location: url } });
  });

  context.setDefaultNavigationTimeout(90000);
  const page = await context.newPage();

  if (isExplore) {
    await ensureLoggedIn(page).catch((e) => console.log(String(e.message ?? e)));
    await page.pause();
    await browser.close();
    return;
  }

  await ensureLoggedIn(page);

  const records = await collectAndExtract(page);
  console.log(`抽出合計: ${records.length}件（同一報告の重複は除去済み）`);
  await browser.close();

  if (isDryRun) {
    console.table(records);
    console.log("dry-runのためExcel・スプシへは書き込みませんでした");
    return;
  }

  const appended = await appendToExcel(records);
  console.log(`プロダクト日報へ追記: ${appended.length}件`);

  writeCsv(appended);
  await appendToSpreadsheet(appended);
})();
