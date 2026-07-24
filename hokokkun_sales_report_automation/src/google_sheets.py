"""Googleスプレッドシートへの転記。"""
from __future__ import annotations

import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class GoogleSheetsError(Exception):
    """Googleスプレッドシートへの転記に失敗した場合のエラー。"""


def append_rows(spreadsheet_id: str, sheet_name: str, credentials_path: str, rows: list[list]) -> int:
    """rows をシートの末尾に追記する。追記できた件数を返す。"""
    if not rows:
        return 0

    try:
        credentials = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
        client = gspread.authorize(credentials)
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(sheet_name)
        worksheet.append_rows(rows, value_input_option="USER_ENTERED")
    except Exception as e:
        raise GoogleSheetsError(f"Googleスプレッドシートへの転記に失敗しました: {e}") from e

    return len(rows)
