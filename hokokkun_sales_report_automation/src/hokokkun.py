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

    page.goto(config.hokokkun_url)
    page.get_by_role("textbox", name="ID").fill(config.hokokkun_login_id)
    page.get_by_role("textbox", name="PASSWORD").fill(config.hokokkun_password)
    page.get_by_role("button", name="SIGN IN").click()
    page.wait_for_load_state("networkidle")

    if LOGIN_URL_MARKER in page.url:
        raise LoginError("ほうこっくんへのログインに失敗しました。IDまたはパスワードを確認してください。")

    top_url = urljoin(config.hokokkun_url, TOP_PATH)
    page.goto(top_url)
    page.wait_for_load_state("networkidle")


def open_sales_list(page: Page, target_day: str) -> None:
    """前日または当日の営業一覧を表示する。target_day は "前日" または "当日"。"""
    raise NotImplementedError("次の記録セッションで実装します。")


def get_sales_list_items(page: Page):
    """営業一覧に表示された報告を取得する。"""
    raise NotImplementedError("次の記録セッションで実装します。")


def open_sales_detail(page: Page, item) -> None:
    """営業一覧から詳細画面を開く。"""
    raise NotImplementedError("次の記録セッションで実装します。")


def extract_sales_detail(page: Page) -> dict:
    """詳細画面から必要項目を取得する。"""
    raise NotImplementedError("次の記録セッションで実装します。")


def return_to_sales_list(page: Page) -> None:
    """詳細画面から営業一覧へ戻る。"""
    raise NotImplementedError("次の記録セッションで実装します。")
