"""ほうこっくんから取得したデータを、Googleスプレッドシート
「営業活動報告フォーム」シートの行データへ変換する。

このシートは1つのフォームに複数の報告カテゴリー（契約報告・NG報告・
電話対応報告 等）が混在しており、カテゴリーごとに使う列（ブロック）が
異なる。契約報告はA〜BK列、NG報告はBL〜CN列を使う。
"""
from __future__ import annotations

import re

from openpyxl.utils import column_index_from_string

CATEGORY_CONTRACT = "契約報告"
CATEGORY_NG = "NG報告"

# 共通列（カテゴリーに関係なく1回だけ使う）。
COL_TIMESTAMP = "A"
COL_CATEGORY = "C"

# 契約報告ブロック（A〜BK列）。
CONTRACT_BLOCK = {
    "sales_method": "E",
    "type": "F",
    "start_datetime": "G",
    "end_datetime": "H",
    "sales_staff": "J",
    "sales_staff_dept": "K",
    "accompany": "L",
    "company_name": "M",
    "company_name_kana": "N",
    "company_address": "O",
    "representative": "P",
    "representative_kana": "Q",
    "industry": "R",
    "email1": "S",
    "email2": "T",
    "sales_location": "U",
    "appointment_staff": "V",
    "appointment_staff_dept": "W",
    "lease_amount": "X",
    "lump_sum_amount": "EQ",
    "initial_fee": "Y",
    "advance_payment": "Z",
    "sms_application_status": "AA",
    "sms_member_info": "AB",
    "credit_card": "AD",
    "proposed_product": "AE",
    "proposal_content": "AF",
    "delivery_item": "AG",
    "entertainment_expense": "AH",
    "incoming_call1": "AJ",
    "incoming_call2": "AK",
    "outgoing_call1": "AL",
    "outgoing_call2": "AM",
    "secret": "AN",
    "confirmation_phone": "AP",
    "identification": "AQ",
    "co_owner": "AR",
    "greeting_visit": "AT",
    "greeting_zoom": "AV",
    "cyteki_support_accompany": "AW",
    "personality": "AX",
    "owner_business": "AY",
    "appeal_points": "AZ",
    "target": "BA",
    "business_name_reason": "BB",
    "other_number_told": "BC",
    "call_restricted_numbers": "BD",
}

# NG報告ブロック（BL〜CN列）。契約報告と違い、金額・納品物件・クレカ等は無い。
NG_BLOCK = {
    "sales_method": "BM",
    "type": "BN",
    "start_datetime": "BO",
    "end_datetime": "BP",
    "re_visit": "BQ",
    "sales_staff": "BR",
    "sales_staff_dept": "BS",
    "accompany": "BT",
    "company_name": "BU",
    "company_name_kana": "BV",
    "company_address": "BW",
    "representative": "BX",
    "representative_kana": "BY",
    "industry": "BZ",
    "email1": "CA",
    "email2": "CB",
    "sales_location": "CC",
    "appointment_staff": "CD",
    "appointment_staff_dept": "CE",
    "proposed_product": "CF",
    "proposal_content": "CG",
    "entertainment_expense": "CH",
    "incoming_call1": "CJ",
    "incoming_call2": "CK",
    "outgoing_call1": "CL",
    "outgoing_call2": "CM",
    "ng_reason": "CN",
}

# EQ列（一括金額）が最も右側の使用列。
SHEET_TOTAL_COLUMNS = column_index_from_string("EQ")

_LEASE_AMOUNT_RE = re.compile(r"([\d,]+)")
_ZERO_INSTALLMENT_RE = re.compile(r"×\s*0\s*回")
_TIME_RANGE_RE = re.compile(r"([^\s〜~\-]+)\s*[〜~\-]\s*([^\s〜~\-]+)")


def parse_amount(amount_text: str) -> tuple[str, str]:
    """金額欄の文字列から (リース月額, 一括金額) を求める。

    「¥25,000×60回」のような通常の分割契約なら、先頭の単価をリース月額として返す。
    「¥1,930,000＋¥0×0回」のように「×0回」が含まれる一括契約なら、
    先頭の金額を一括金額として返し、リース月額は空欄にする。
    """
    text = amount_text or ""
    match = _LEASE_AMOUNT_RE.search(text)
    first_number = match.group(1).replace(",", "") if match else ""
    if not first_number:
        return "", ""
    if _ZERO_INSTALLMENT_RE.search(text):
        return "", first_number
    return first_number, ""


def combine_start_end_datetime(contract_date: str, meeting_time: str) -> tuple[str, str]:
    """契約日＋商談時間から、開始・終了それぞれの日時文字列を作る。"""
    if not contract_date or not meeting_time:
        return "", ""
    match = _TIME_RANGE_RE.search(meeting_time)
    if not match:
        return "", ""
    start_time, end_time = match.group(1), match.group(2)
    return f"{contract_date} {start_time}", f"{contract_date} {end_time}"


def build_sheet_row(detail: dict) -> list:
    """詳細情報から、シートへ1行分（A〜CN列）のリストを組み立てる。"""
    row = [""] * SHEET_TOTAL_COLUMNS

    def set_value(col_letter: str, value) -> None:
        row[column_index_from_string(col_letter) - 1] = value

    is_ng = detail.get("result", "") == "NG"
    category = CATEGORY_NG if is_ng else CATEGORY_CONTRACT
    set_value(COL_CATEGORY, category)

    start_dt, end_dt = combine_start_end_datetime(detail.get("contract_date", ""), detail.get("meeting_time", ""))
    lease_amount, lump_sum_amount = parse_amount(detail.get("amount", ""))

    block = NG_BLOCK if is_ng else CONTRACT_BLOCK
    for field, col_letter in block.items():
        if field == "start_datetime":
            set_value(col_letter, start_dt)
        elif field == "end_datetime":
            set_value(col_letter, end_dt)
        elif field == "lease_amount":
            set_value(col_letter, lease_amount)
        elif field == "lump_sum_amount":
            set_value(col_letter, lump_sum_amount)
        else:
            set_value(col_letter, detail.get(field, ""))

    return row
