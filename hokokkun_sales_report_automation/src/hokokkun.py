"""ほうこっくんの画面操作。"""
from __future__ import annotations

from urllib.parse import urljoin

from playwright.sync_api import Page

from .config import Config

TOP_PATH = "/top/"
LOGIN_URL_MARKER = "login.php"


class HokokkunError(Exception):
    """ほうこっくんの操作に失敗した場合のエラー。"""


class LoginError(HokokkunError):
    """ログインに失敗した場合のエラー（全体処理を停止する）。"""


class UnauthorizedAccessError(HokokkunError):
    """詳細画面で「不正なアクセスです。」が表示された場合のエラー。"""


def login_to_hokokkun(page: Page, config: Config) -> None:
    """ほうこっくんへログインする。"""
    if not config.hokokkun_url or not config.hokokkun_login_id or not config.hokokkun_password:
        raise LoginError("HOKOKKUN_URL / HOKOKKUN_LOGIN_ID / HOKOKKUN_PASSWORD が .env に設定されていません。")

    page.on("dialog", lambda dialog: dialog.accept())

    page.goto(config.hokokkun_url, timeout=90000)
    page.get_by_role("textbox", name="ID").fill(config.hokokkun_login_id)
    page.get_by_role("textbox", name="PASSWORD").fill(config.hokokkun_password)

    # クリック直後の画面遷移はサーバー応答が遅く detach/timeout しやすいため、
    # click自体の遷移待ちはせず、少し待ってからトップページへ明示的に移動する。
    page.get_by_role("button", name="SIGN IN").click(no_wait_after=True)
    page.wait_for_timeout(3000)

    top_url = urljoin(config.hokokkun_url, TOP_PATH)
    page.goto(top_url, timeout=90000)

    if LOGIN_URL_MARKER in page.url:
        raise LoginError("ほうこっくんへのログインに失敗しました。IDまたはパスワードを確認してください。")


def open_sales_list(page: Page, target_day: str) -> None:
    """前日または当日の営業一覧を表示する。target_day は "前日" または "当日"。"""
    if target_day == "前日":
        page.get_by_role("link", name="前日").click()
    elif target_day == "当日":
        page.get_by_role("link", name="次日").click()
    else:
        raise ValueError(f'target_day は "前日" か "当日" を指定してください: {target_day}')

    page.get_by_text("営業", exact=True).click()
    page.wait_for_timeout(1500)


# 一覧の列順（0始まり）。5.4節・一覧画面のスクリーンショットに対応。
LIST_COL_UPDATED_AT = 0     # 更新日時
LIST_COL_COMPANY_NAME = 6  # 会社名


def get_sales_list_items(page: Page) -> list[dict]:
    """営業一覧に表示された報告を取得する。"""
    rows = page.get_by_role("row")
    items: list[dict] = []
    for row_index in range(rows.count()):
        cells = rows.nth(row_index).get_by_role("cell")
        if cells.count() == 0:
            # 見出し行(th)には role="cell" が無いためスキップされる。
            continue
        updated_at = cells.nth(LIST_COL_UPDATED_AT).inner_text().strip()
        company_name = cells.nth(LIST_COL_COMPANY_NAME).inner_text().strip()
        if not updated_at:
            continue
        items.append({
            "row_index": row_index,
            "updated_at": updated_at,
            "company_name": company_name,
        })
    return items


def open_sales_detail(page: Page, item: dict) -> None:
    """営業一覧から詳細画面を開く。"""
    row = page.get_by_role("row").nth(item["row_index"])
    row.get_by_role("cell").nth(LIST_COL_UPDATED_AT).click()
    page.wait_for_timeout(1500)


def extract_sales_detail(page: Page) -> dict:
    """詳細画面から必要項目を取得する。"""
    raise NotImplementedError("次の記録セッションで実装します。")


def return_to_sales_list(page: Page) -> None:
    """詳細画面から営業一覧へ戻る。"""
    page.get_by_role("button", name="一覧へ戻る").click()
    page.wait_for_timeout(1500)
