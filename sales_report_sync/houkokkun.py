# -*- coding: utf-8 -*-
"""ほうこっくんから営業報告を取得するSeleniumクライアント。

画面遷移は以下の要件どおりに行う:
  ログイン → 「営業」を選択 → 「前日」を選択 → 一覧の詳細を1件ずつ開いて抽出
  → 「次日」で当日へ切り替え → 同じ処理

詳細画面で「不正なアクセスです。」と表示された場合は、
同じ日の営業一覧を再表示して1回だけ再試行する。
1件の取得失敗では全体を止めず、ログに残して次の件へ進む。

※ 画面のHTML構造は環境によって異なるため、ログインフォームや
   一覧のリンクが見つからない場合は下の SELECTORS を実際の画面に
   合わせて調整すること(README「セレクタの調整」参照)。
"""

import logging
import time
import unicodedata

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.common.exceptions import WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from fields import FIELDS, label_to_field

logger = logging.getLogger(__name__)

INVALID_ACCESS_TEXT = "不正なアクセス"

# 実際の画面に合わせて調整するセレクタ類 -------------------------------------
SELECTORS = {
    # ログイン画面の入力欄(name属性)。合わない場合は実画面のname/idに変更する。
    "login_id_names": ["login_id", "loginId", "user_id", "userid", "id", "account"],
    "login_pw_names": ["password", "passwd", "login_pw", "pass"],
    # ログインボタン
    "login_submit_xpath": (
        "//button[@type='submit'] | //input[@type='submit'] "
        "| //button[contains(., 'ログイン')] | //input[@type='button' and contains(@value, 'ログイン')]"
    ),
    # 項目一覧の「営業」メニュー
    "sales_menu_xpath": "//a[normalize-space()='営業'] | //button[normalize-space()='営業']",
    # 日付切り替え
    "prev_day_xpath": (
        "//a[normalize-space()='前日'] | //button[normalize-space()='前日'] "
        "| //input[@type='button' and @value='前日'] | //input[@type='submit' and @value='前日']"
    ),
    "next_day_xpath": (
        "//a[normalize-space()='次日'] | //button[normalize-space()='次日'] "
        "| //input[@type='button' and @value='次日'] | //input[@type='submit' and @value='次日'] "
        "| //a[normalize-space()='翌日'] | //button[normalize-space()='翌日']"
    ),
    # 営業報告一覧の各行から詳細画面へ飛ぶリンク。
    # 既定では「一覧テーブルの行内のリンク」をすべて対象にする。
    "detail_link_xpath": "//table//tr//a[@href]",
}
# ---------------------------------------------------------------------------


def _clean_text(text: str) -> str:
    if text is None:
        return ""
    text = unicodedata.normalize("NFKC", text)
    return " ".join(text.split())


def parse_detail_html(html: str) -> dict:
    """詳細画面のHTMLから抽出項目を取り出す。

    th/td・dt/dd のラベルと値の組をすべて拾い、FIELDS に該当する
    ものだけを採用する。取得できない項目は空欄のままにする。
    """
    soup = BeautifulSoup(html, "html.parser")
    pairs = {}

    for th in soup.find_all("th"):
        td = th.find_next_sibling("td")
        if td is not None:
            pairs.setdefault(_clean_text(th.get_text()), _clean_text(td.get_text()))

    for dt in soup.find_all("dt"):
        dd = dt.find_next_sibling("dd")
        if dd is not None:
            pairs.setdefault(_clean_text(dt.get_text()), _clean_text(dd.get_text()))

    record = {field: "" for field in FIELDS}
    for label, value in pairs.items():
        field = label_to_field(label)
        if field:
            record[field] = value
    return record


class HoukokkunClient:
    def __init__(self, config):
        self.base_url = config["houkokkun"]["url"].rstrip("/")
        self.login_id = config["houkokkun"]["login_id"]
        self.password = config["houkokkun"]["password"]
        self.wait_seconds = int(config["houkokkun"].get("wait_seconds", "15"))
        self.headless = config["houkokkun"].get("headless", "true").lower() == "true"
        self.driver = None
        # 一覧を再表示するために「当日から何回前日を押したか」を記録する
        self._day_offset = 0

    # -- ライフサイクル -----------------------------------------------------
    def __enter__(self):
        options = webdriver.ChromeOptions()
        if self.headless:
            options.add_argument("--headless=new")
        options.add_argument("--window-size=1400,1000")
        options.add_argument("--disable-gpu")
        self.driver = webdriver.Chrome(options=options)
        self.driver.set_page_load_timeout(60)
        return self

    def __exit__(self, exc_type, exc, tb):
        if self.driver is not None:
            try:
                self.driver.quit()
            except WebDriverException:
                pass
        return False

    # -- 基本操作 -----------------------------------------------------------
    def _wait(self):
        return WebDriverWait(self.driver, self.wait_seconds)

    def _find_first(self, xpath):
        elements = self.driver.find_elements(By.XPATH, xpath)
        for element in elements:
            if element.is_displayed():
                return element
        return elements[0] if elements else None

    def _click(self, xpath, description):
        element = self._find_first(xpath)
        if element is None:
            raise RuntimeError(
                f"{description} が見つかりません。houkokkun.py の SELECTORS を実画面に合わせて調整してください。"
            )
        element.click()
        time.sleep(1.0)

    # -- ログイン -----------------------------------------------------------
    def login(self):
        logger.info("ほうこっくんへログインします: %s", self.base_url)
        self.driver.get(self.base_url)
        time.sleep(1.0)

        id_input = self._find_input(SELECTORS["login_id_names"], "text")
        pw_input = self._find_input(SELECTORS["login_pw_names"], "password")
        if id_input is None or pw_input is None:
            raise RuntimeError(
                "ログインフォームが見つかりません。SELECTORS の login_id_names / login_pw_names を調整してください。"
            )
        id_input.clear()
        id_input.send_keys(self.login_id)
        pw_input.clear()
        pw_input.send_keys(self.password)
        self._click(SELECTORS["login_submit_xpath"], "ログインボタン")
        time.sleep(1.5)

        if self._find_input(SELECTORS["login_pw_names"], "password") is not None:
            raise RuntimeError("ログインに失敗しました。ID/パスワードを確認してください。")
        logger.info("ログインに成功しました。")

    def _find_input(self, names, fallback_type):
        for name in names:
            elements = self.driver.find_elements(By.NAME, name)
            for element in elements:
                if element.is_displayed():
                    return element
        # name属性で見つからない場合はtype属性で探す
        for element in self.driver.find_elements(By.XPATH, f"//input[@type='{fallback_type}']"):
            if element.is_displayed():
                return element
        return None

    # -- 一覧表示 -----------------------------------------------------------
    def open_sales_list(self):
        """項目一覧から「営業」を選択して営業報告一覧(当日)を表示する。"""
        self._click(SELECTORS["sales_menu_xpath"], "項目一覧の「営業」")
        self._day_offset = 0

    def go_previous_day(self):
        self._click(SELECTORS["prev_day_xpath"], "「前日」ボタン")
        self._day_offset += 1

    def go_next_day(self):
        self._click(SELECTORS["next_day_xpath"], "「次日」ボタン")
        self._day_offset = max(0, self._day_offset - 1)

    def _refresh_list(self):
        """同じ日の営業一覧を再表示する(不正アクセス時・エラー復帰用)。

        メニューの「営業」から入り直し、記録済みの回数だけ「前日」を押して
        同じ日に戻る。
        """
        offset = self._day_offset
        self.open_sales_list()
        for _ in range(offset):
            self.go_previous_day()

    # -- 詳細取得 -----------------------------------------------------------
    def _detail_links(self):
        return [
            element
            for element in self.driver.find_elements(By.XPATH, SELECTORS["detail_link_xpath"])
            if element.is_displayed()
        ]

    def fetch_reports(self, stats, day_label):
        """表示中の日の営業報告一覧を1件ずつ開き、抽出結果のリストを返す。

        要件どおり、一覧を表示した直後にその一覧の詳細を順番に開く。
        1件の失敗では止めず、ログに残して続行する。
        """
        count = len(self._detail_links())
        logger.info("[%s] 営業報告一覧: %d 件", day_label, count)
        reports = []
        for index in range(count):
            try:
                record = self._open_and_extract(index, day_label)
            except Exception:
                logger.exception("[%s] %d件目の取得中にエラーが発生しました。次の件へ進みます。", day_label, index + 1)
                stats.failed += 1
                try:
                    self._refresh_list()
                except Exception:
                    logger.exception("[%s] 一覧の再表示に失敗しました。この日の処理を打ち切ります。", day_label)
                    break
                continue
            if record is None:
                stats.failed += 1
            else:
                reports.append(record)
        return reports

    def _open_and_extract(self, index, day_label):
        """一覧のindex番目の詳細を開いて抽出する。

        「不正なアクセスです。」が表示された場合は同じ日の一覧を
        再表示して1回だけ再試行する。失敗した場合は None を返す。
        """
        for attempt in (1, 2):
            links = self._detail_links()
            if index >= len(links):
                logger.error("[%s] %d件目: 一覧のリンク数が変わり詳細を開けません。", day_label, index + 1)
                return None
            links[index].click()
            time.sleep(1.0)

            html = self.driver.page_source
            if INVALID_ACCESS_TEXT in html:
                logger.warning(
                    "[%s] %d件目: 「不正なアクセスです。」が表示されました(%d回目)。一覧を再表示します。",
                    day_label, index + 1, attempt,
                )
                self._refresh_list()
                continue

            record = parse_detail_html(html)
            self.driver.back()
            time.sleep(1.0)
            if not record.get("会社名"):
                logger.warning("[%s] %d件目: 会社名が取得できませんでした。空欄項目のまま処理します。", day_label, index + 1)
            return record

        logger.error("[%s] %d件目: 再試行しても詳細を開けませんでした。スキップします。", day_label, index + 1)
        return None
