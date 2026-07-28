"""
更新代行反映 CMS自動化スクリプト
対象ファイル: 更新代行反映CMS1.xlsx / 更新代行反映CMS5.xlsx
フォルダ: G:\共有ドライブ\RPA運用-品質保証課連携\
"""

import os
import sys
import time
import asyncio
import openpyxl
from pathlib import Path
from playwright.async_api import async_playwright, Page

# ===== 設定 =====
FOLDER_PATH = Path(r"G:\共有ドライブ\RPA運用-技術課連携\RPA運用-品質保証課連携テスト")
TARGET_FILES = [
    "更新代行反映CMS1.xlsx",
    "更新代行反映CMS5.xlsx",
]

# Excelの列名定義（1行目がヘッダーであることを前提）
COL_URL    = "管理画面URL"
COL_ID     = "ID"
COL_PW     = "PW"
COL_TITLE  = "ブログタイトル"
COL_BODY   = "ブログ内容"

HEADLESS = False  # True にすると画面表示なし（バックグラウンド実行）


# ===== Excel読み込み =====
def load_excel(file_path: Path) -> list[dict]:
    """Excelファイルを読み込み、行データのリストを返す"""
    wb = openpyxl.load_workbook(file_path)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):  # 空行スキップ
            continue
        record = dict(zip(headers, row))
        # 必須カラムがすべて揃っている行のみ対象
        if all(record.get(c) for c in [COL_URL, COL_ID, COL_PW, COL_TITLE, COL_BODY]):
            rows.append(record)
    wb.close()
    return rows


# ===== ブラウザ操作 =====
async def process_row(page: Page, record: dict, row_num: int):
    """1行分のCMS投稿処理"""
    url    = str(record[COL_URL]).strip()
    user   = str(record[COL_ID]).strip()
    pw     = str(record[COL_PW]).strip()
    title  = str(record[COL_TITLE]).strip()
    body   = str(record[COL_BODY]).strip()

    print(f"  [行{row_num}] アクセス中: {url}")

    # ---- STEP 4: 管理画面URLにアクセス ----
    # Basic認証がある場合はURL埋め込み方式を試みる
    parsed = url
    if "://" in url and "@" not in url:
        scheme, rest = url.split("://", 1)
        parsed = f"{scheme}://{user}:{pw}@{rest}"

    try:
        await page.goto(parsed, wait_until="domcontentloaded", timeout=30000)
    except Exception:
        # Basic認証なし → 通常アクセスしてログインフォームを探す
        await page.goto(url, wait_until="domcontentloaded", timeout=30000)

    # ログインフォームが表示される場合（ID/PW入力欄があれば自動入力）
    await _try_login(page, user, pw)

    # ---- STEP 5: 「お知らせ」ボタンをクリック ----
    print(f"  [行{row_num}] 「お知らせ」クリック")
    await _click_text(page, "お知らせ")

    # ---- STEP 6: 「項目追加」をクリック ----
    print(f"  [行{row_num}] 「項目追加」クリック")
    await _click_text(page, "項目追加")

    # ---- STEP 7: タイトル入力 ----
    print(f"  [行{row_num}] タイトル入力: {title}")
    await _fill_field(page, "タイトル", title)

    # ---- STEP 8: 本文入力 ----
    print(f"  [行{row_num}] 本文入力")
    await _fill_field(page, "本文", body)

    # ---- STEP 9: 登録ボタンクリック ----
    print(f"  [行{row_num}] 「登録」クリック")
    await _click_text(page, "登録")

    # ---- STEP 10: OKボタンクリック（確認ダイアログ） ----
    print(f"  [行{row_num}] 「OK」クリック")
    await _handle_ok(page)

    print(f"  [行{row_num}] 完了")


async def _try_login(page: Page, user: str, pw: str):
    """ログインフォームが存在すれば入力してsubmitする"""
    try:
        # よくあるID/PW入力欄のセレクター候補
        id_selectors = [
            'input[type="text"][name*="user" i]',
            'input[type="text"][name*="id" i]',
            'input[type="text"][name*="login" i]',
            'input[name*="user" i]',
            'input[name*="id" i]',
            '#username', '#user_id', '#login_id',
        ]
        pw_selectors = [
            'input[type="password"]',
        ]

        id_input = None
        for sel in id_selectors:
            el = page.locator(sel).first
            if await el.count() > 0:
                id_input = el
                break

        pw_input = page.locator(pw_selectors[0]).first
        has_pw = await pw_input.count() > 0

        if id_input and has_pw:
            await id_input.fill(user)
            await pw_input.fill(pw)
            # Submitボタンまたはログインボタンを探してクリック
            submit_candidates = [
                'input[type="submit"]',
                'button[type="submit"]',
                'button:has-text("ログイン")',
                'button:has-text("サインイン")',
                'button:has-text("login")',
            ]
            for sel in submit_candidates:
                el = page.locator(sel).first
                if await el.count() > 0:
                    await el.click()
                    await page.wait_for_load_state("domcontentloaded", timeout=15000)
                    break
    except Exception as e:
        print(f"    ログイン処理をスキップ (理由: {e})")


async def _click_text(page: Page, text: str):
    """テキストを含むリンク・ボタン・メニューをクリック"""
    candidates = [
        f'a:has-text("{text}")',
        f'button:has-text("{text}")',
        f'input[value="{text}"]',
        f'li:has-text("{text}")',
        f'span:has-text("{text}")',
        f'td:has-text("{text}")',
        f'div:has-text("{text}")',
    ]
    for sel in candidates:
        el = page.locator(sel).first
        if await el.count() > 0:
            await el.click()
            await page.wait_for_load_state("domcontentloaded", timeout=15000)
            return
    raise RuntimeError(f'クリック対象が見つかりません: 「{text}」')


async def _fill_field(page: Page, label: str, value: str):
    """ラベルに対応する入力欄を探して値を入力する"""
    # ラベルテキストの隣にある input/textarea を探す
    candidates = [
        f'label:has-text("{label}") >> .. >> input',
        f'label:has-text("{label}") >> .. >> textarea',
        f'input[placeholder*="{label}"]',
        f'textarea[placeholder*="{label}"]',
        f'input[name*="{label}"]',
        f'textarea[name*="{label}"]',
        # th/td テーブル形式の場合
        f'th:has-text("{label}") >> .. >> td >> input',
        f'th:has-text("{label}") >> .. >> td >> textarea',
        f'td:has-text("{label}") >> .. >> td >> input',
        f'td:has-text("{label}") >> .. >> td >> textarea',
    ]
    for sel in candidates:
        try:
            el = page.locator(sel).first
            if await el.count() > 0:
                await el.clear()
                await el.fill(value)
                return
        except Exception:
            continue
    raise RuntimeError(f'入力欄が見つかりません: 「{label}」')


async def _handle_ok(page: Page):
    """登録後のOK/確認ダイアログを処理する"""
    try:
        # ブラウザネイティブダイアログ（alert/confirm）の場合
        page.on("dialog", lambda d: asyncio.ensure_future(d.accept()))
        await asyncio.sleep(1)
    except Exception:
        pass

    # モーダル内のOKボタン
    ok_candidates = [
        'button:has-text("OK")',
        'button:has-text("ok")',
        'input[value="OK"]',
        'a:has-text("OK")',
    ]
    for sel in ok_candidates:
        el = page.locator(sel).first
        if await el.count() > 0:
            await el.click()
            await page.wait_for_load_state("domcontentloaded", timeout=10000)
            return


# ===== メイン処理 =====
async def main():
    processed_files = []

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=HEADLESS)
        context = await browser.new_context()
        page = await context.new_page()

        for filename in TARGET_FILES:
            file_path = FOLDER_PATH / filename

            # ---- STEP 1/2: ファイル存在確認 ----
            if not file_path.exists():
                print(f"[スキップ] ファイルが存在しません: {file_path}")
                continue

            print(f"\n{'='*50}")
            print(f"[処理開始] {filename}")
            print(f"{'='*50}")

            # ---- STEP 2-a/3: Excelからデータ抽出 ----
            rows = load_excel(file_path)
            print(f"  抽出行数: {len(rows)} 件")

            if not rows:
                print("  処理対象データがありません。スキップします。")
                continue

            # ---- STEP 4〜11: 各行を処理 ----
            for i, record in enumerate(rows, start=1):
                try:
                    await process_row(page, record, i)
                except Exception as e:
                    print(f"  [行{i}] エラー発生: {e}")
                    # エラーがあっても次の行に進む
                    continue

            processed_files.append(file_path)
            print(f"[完了] {filename}")

        await browser.close()

    # ---- STEP 13: 使用したExcelファイルを削除 ----
    if processed_files:
        print(f"\n{'='*50}")
        print("処理済みファイルを削除します...")
        for fp in processed_files:
            try:
                fp.unlink()
                print(f"  削除完了: {fp.name}")
            except Exception as e:
                print(f"  削除失敗: {fp.name} ({e})")

    print("\nすべての処理が完了しました。")


if __name__ == "__main__":
    asyncio.run(main())
