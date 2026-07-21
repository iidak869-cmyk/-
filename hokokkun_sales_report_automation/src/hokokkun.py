"""ほうこっくんの画面操作。"""
from __future__ import annotations

import re
from urllib.parse import urljoin

from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

from .config import Config

TOP_PATH = "/top/"
LOGIN_URL_MARKER = "login.php"


class HokokkunError(Exception):
    """ほうこっくんの操作に失敗した場合のエラー。"""


class LoginError(HokokkunError):
    """ログインに失敗した場合のエラー（全体処理を停止する）。"""


class UnauthorizedAccessError(HokokkunError):
    """詳細画面で「不正なアクセスです。」が表示された場合のエラー。"""


class CompanyNameNotFoundError(HokokkunError):
    """詳細画面から会社名を取得できなかった場合のエラー（この報告は登録せずスキップする）。"""


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


REPORT_LIST_HEADING_RE = re.compile("の報告一覧")
DETAIL_HEADING_TEXT = "詳細表示"
NAVIGATION_TIMEOUT = 60000


def open_sales_list(page: Page, target_day: str) -> None:
    """前日または当日の営業一覧を表示する。target_day は "前日" または "当日"。"""
    if target_day == "前日":
        page.get_by_role("link", name="前日").click(no_wait_after=True)
    elif target_day == "当日":
        page.get_by_role("link", name="次日").click(no_wait_after=True)
    else:
        raise ValueError(f'target_day は "前日" か "当日" を指定してください: {target_day}')
    page.get_by_text(REPORT_LIST_HEADING_RE).first.wait_for(timeout=NAVIGATION_TIMEOUT)

    page.get_by_text("営業", exact=True).click(no_wait_after=True)
    page.get_by_text(REPORT_LIST_HEADING_RE).first.wait_for(timeout=NAVIGATION_TIMEOUT)
    page.wait_for_timeout(1000)


# 一覧の列順（0始まり）。一覧画面のスクリーンショットに対応。
LIST_COL_UPDATED_AT = 0      # 更新日時
LIST_COL_MEETING_DATE = 1   # 営業日付
LIST_COL_AMOUNT = 5          # 金額
LIST_COL_COMPANY_NAME = 6   # 会社名
LIST_COLUMN_COUNT = 8        # 更新日時〜チェック欄の8列


def get_sales_list_items(page: Page) -> list[dict]:
    """営業一覧に表示された報告を取得する。"""
    rows = page.get_by_role("row")
    items: list[dict] = []
    for row_index in range(rows.count()):
        cells = rows.nth(row_index).get_by_role("cell")
        if cells.count() != LIST_COLUMN_COUNT:
            # 見出し行や、画面上部のカレンダー等の別の表はここで除外する。
            continue
        updated_at = cells.nth(LIST_COL_UPDATED_AT).inner_text().strip()
        if not updated_at:
            continue
        items.append({
            "row_index": row_index,
            "updated_at": updated_at,
            "meeting_date": cells.nth(LIST_COL_MEETING_DATE).inner_text().strip(),
            "amount": cells.nth(LIST_COL_AMOUNT).inner_text().strip(),
            "company_name": cells.nth(LIST_COL_COMPANY_NAME).inner_text().strip(),
        })
    return items


def open_sales_detail(page: Page, item: dict) -> None:
    """営業一覧から詳細画面を開く。"""
    row = page.get_by_role("row").nth(item["row_index"])
    row.get_by_role("cell").nth(LIST_COL_UPDATED_AT).click(no_wait_after=True)
    # 「不正なアクセスです。」の場合は「詳細表示」が出ないため、ここでは待つだけで
    # 例外にはしない。不正アクセスかどうかの判定は呼び出し側で行う。
    try:
        page.get_by_text(DETAIL_HEADING_TEXT).wait_for(timeout=NAVIGATION_TIMEOUT)
    except PlaywrightTimeoutError:
        pass
    page.wait_for_timeout(1000)


# 詳細画面の項目ラベル→取得結果のキーの対応。「現在の状況」列の値をそのまま使う項目のみ。
DETAIL_LABEL_TO_FIELD = {
    "営業方法": "sales_method",
    "結果": "result",
    "営業担当": "sales_staff",
    "同行": "accompany",
    "種類": "type",
    "更新日時": "updated_at_raw",
    "営業日時": "meeting_datetime_raw",
    "会社名": "company_name",
    "会社所在地": "company_address",
    "代表者名": "representative",
    "業種": "industry",
    "アポ担当": "appointment_staff",
    "納品物件": "delivery_item",
    "NG理由": "ng_reason",
}

_MEETING_TIME_RE = re.compile(r"(\d{1,2}:\d{2}\s*[〜~\-]\s*\d{1,2}:\d{2})")
_CREDIT_CARD_RE = re.compile(r"クレカ[^\n】]*】\s*(持|未所持)")
_INITIAL_COLLECTION_RE = re.compile(r"初期[^\n】]*】\s*①期日[:：]\s*([^\n]*)")
_YEAR_RE = re.compile(r"(\d{4})年")
_MONTH_DAY_RE = re.compile(r"(\d{1,2})月(\d{1,2})日")


def extract_sales_detail(page: Page) -> dict:
    """詳細画面から必要項目を取得する。

    項目の多くは「項目名 | 現在の状況 | (初期登録状況)」という単純な行だが、
    「報告内容」だけは多数のサブ項目（クレカ所持・初期回収 等）を含む
    自由記述の1項目としてまとめて入っている。そのためクレカ所持・初期回収は
    「報告内容」の文章から正規表現で拾う。
    """
    detail: dict = {}
    report_content = ""

    rows = page.get_by_role("row")
    for row_index in range(rows.count()):
        cells = rows.nth(row_index).get_by_role("cell")
        if cells.count() < 2:
            continue
        label = cells.nth(0).inner_text().strip().strip("【】")
        current_value = cells.nth(1).inner_text().strip()

        if label == "報告内容":
            report_content = current_value
            continue

        if label in DETAIL_LABEL_TO_FIELD:
            detail[DETAIL_LABEL_TO_FIELD[label]] = current_value

    meeting_datetime_raw = detail.pop("meeting_datetime_raw", "")
    meeting_time_match = _MEETING_TIME_RE.search(meeting_datetime_raw)
    detail["meeting_time"] = meeting_time_match.group(1).replace(" ", "") if meeting_time_match else ""

    year_match = _YEAR_RE.search(detail.pop("updated_at_raw", ""))
    month_day_match = _MONTH_DAY_RE.search(meeting_datetime_raw)
    if year_match and month_day_match:
        year = year_match.group(1)
        month = int(month_day_match.group(1))
        day = int(month_day_match.group(2))
        detail["contract_date"] = f"{year}/{month:02d}/{day:02d}"
    else:
        detail["contract_date"] = ""

    credit_card_match = _CREDIT_CARD_RE.search(report_content)
    detail["credit_card"] = "有" if credit_card_match and credit_card_match.group(1) == "持" else ""

    initial_collection_match = _INITIAL_COLLECTION_RE.search(report_content)
    initial_collection_raw = initial_collection_match.group(1).strip() if initial_collection_match else ""
    detail["initial_collection"] = "完了" if initial_collection_raw.startswith("無") else ""

    # 契約形態は納品物件の値をそのまま使う。
    detail["contract_type"] = detail.get("delivery_item", "")

    if not detail.get("company_name", "").strip():
        raise CompanyNameNotFoundError("詳細画面から会社名を取得できませんでした。")

    return detail


def return_to_sales_list(page: Page) -> None:
    """詳細画面から営業一覧へ戻る。"""
    page.get_by_role("button", name="一覧へ戻る").click(no_wait_after=True)
    page.get_by_text(REPORT_LIST_HEADING_RE).first.wait_for(timeout=NAVIGATION_TIMEOUT)
    page.wait_for_timeout(1000)


UNAUTHORIZED_ACCESS_TEXT = "不正なアクセスです。"


def is_unauthorized_access(page: Page) -> bool:
    """詳細画面で「不正なアクセスです。」が表示されているか確認する。"""
    return page.get_by_text(UNAUTHORIZED_ACCESS_TEXT).count() > 0
