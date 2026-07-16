# -*- coding: utf-8 -*-
"""今回Excelへ新規追加したデータのみをGoogleスプレッドシートへ転記する。

サービスアカウント(JSONキー)で認証する。対象スプレッドシートを
サービスアカウントのメールアドレスへ共有しておくこと。
"""

import logging

import gspread
from google.oauth2.service_account import Credentials

from fields import FIELDS

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def transfer_new_rows(config, records):
    """新規追加分の records をスプレッドシートの末尾に追記する。"""
    if not records:
        logger.info("スプレッドシートへ転記する新規データはありません。")
        return

    credentials = Credentials.from_service_account_file(
        config["google_sheets"]["service_account_json"], scopes=SCOPES
    )
    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(config["google_sheets"]["spreadsheet_id"])

    worksheet_name = config["google_sheets"].get("worksheet_name", "").strip()
    worksheet = spreadsheet.worksheet(worksheet_name) if worksheet_name else spreadsheet.sheet1

    # ヘッダーが未設定なら1行目に項目名を書き込む
    header = worksheet.row_values(1)
    if not any(header):
        worksheet.update("A1", [FIELDS])
        header = FIELDS

    # シート側のヘッダー順に合わせて値を並べる(FIELDS以外の列は空欄)
    columns = ["".join(str(h).split()) for h in header]
    rows = [[record.get(column, "") for column in columns] for record in records]

    worksheet.append_rows(rows, value_input_option="USER_ENTERED")
    logger.info("スプレッドシートへ %d 件転記しました。", len(rows))
