"""ほうこっくんから取得したデータをExcelの列値へ変換する。"""
from __future__ import annotations

from . import excel_io

DELIVERY_MARK = "○"

# 一覧の「金額」列で、金額が無いことを表す記号（見た目は似ているが別々のUnicode文字）。
_DASH_PLACEHOLDERS = {"-", "ー", "－", "‐", "−", "ｰ"}


def clean_amount(amount_text: str) -> str:
    """一覧の「金額」列の値を整える。

    - 2行目以降（金額が無いことを示す「ー」など）は使わず、1行目だけを使う。
    - 1行目が「ー」等の記号のみの場合は空欄にする。
    """
    first_line = (amount_text or "").splitlines()[0].strip() if amount_text else ""
    if first_line in _DASH_PLACEHOLDERS:
        return ""
    return first_line

# 納品物件の文字列に含まれるキーワード→対応するExcel列。
_DELIVERY_ITEM_KEYWORDS: dict[int, tuple[str, ...]] = {
    excel_io.COL_NET1: ("ネット1",),
    excel_io.COL_NET2: ("ネット2",),
    excel_io.COL_HP: ("HP", "ホームページ"),
    excel_io.COL_VIDEO_SITE: ("動画サイト",),
    excel_io.COL_LP: ("LP", "ランディングページ"),
    excel_io.COL_LINE: ("LINE",),
    excel_io.COL_VIDEO_CR: ("動画CR", "動画クリエイティブ"),
}


def determine_delivery_marks(delivery_item_text: str) -> dict[int, str]:
    """納品物件の列に入れる「○」を決める。

    - 商品が Cyteki の場合: Z列（Cyteki）のみに「○」
    - 商品が Degion(DegiOne) の場合: 該当する納品物件（ネット1・ネット2・HP・
      動画サイト・LP・LINE・動画CR）の列にそれぞれ「○」
    """
    text = (delivery_item_text or "").strip()
    if not text:
        return {}

    if "cyteki" in text.lower():
        return {excel_io.COL_CYTEKI: DELIVERY_MARK}

    marks: dict[int, str] = {}
    lowered = text.lower()
    for col, keywords in _DELIVERY_ITEM_KEYWORDS.items():
        if any(keyword.lower() in lowered for keyword in keywords):
            marks[col] = DELIVERY_MARK
    return marks


def build_excel_row_values(detail: dict) -> dict[int, object]:
    """詳細画面から取得したデータを、Excelの列番号→値の辞書に変換する。"""
    values: dict[int, object] = {
        excel_io.COL_CONTRACT_DATE: detail.get("contract_date", ""),
        excel_io.COL_MEETING_TIME: detail.get("meeting_time", ""),
        excel_io.COL_SALES_METHOD: detail.get("sales_method", ""),
        excel_io.COL_COMPANY_NAME: detail.get("company_name", ""),
        excel_io.COL_RESULT: detail.get("result", ""),
        excel_io.COL_NG_REASON_1: detail.get("ng_reason", ""),
        excel_io.COL_COMPANY_ADDRESS: detail.get("company_address", ""),
        excel_io.COL_REPRESENTATIVE: detail.get("representative", ""),
        excel_io.COL_TYPE: detail.get("type", ""),
        excel_io.COL_SALES_STAFF: detail.get("sales_staff", ""),
        excel_io.COL_ACCOMPANY: detail.get("accompany", ""),
        excel_io.COL_APPOINTMENT_STAFF: detail.get("appointment_staff", ""),
        excel_io.COL_INDUSTRY: detail.get("industry", ""),
        excel_io.COL_CONTRACT_TYPE: detail.get("contract_type", ""),
        excel_io.COL_AMOUNT: detail.get("amount", ""),
        excel_io.COL_CREDIT_CARD: detail.get("credit_card", ""),
        excel_io.COL_INITIAL_COLLECTION: detail.get("initial_collection", ""),
    }
    values.update(determine_delivery_marks(detail.get("delivery_item", "")))
    return values


def build_sheet_row(detail: dict) -> list:
    """Googleスプレッドシート用に、Excelと同じ列構成（A〜AB列）の1行分リストを作る。

    転記先のテスト用スプレッドシートは、営業報告データベースExcelと
    まったく同じ列構成（契約日〜初期回収）のため、Excel用の値をそのまま
    列の並び順に変換するだけでよい。
    """
    values = build_excel_row_values(detail)
    row = [""] * excel_io.COL_INITIAL_COLLECTION
    for col, value in values.items():
        row[col - 1] = value
    return row
