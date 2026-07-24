"""Googleスプレッドシートへの転記。

社内のセキュリティ設定でサービスアカウントとの外部共有ができなかったため、
自分自身のGoogleアカウントでログインするOAuth方式を使う。
初回だけブラウザでログイン・許可が必要で、以降は認証情報がファイルに
保存され、自動的に使い回される。
"""
from __future__ import annotations

import gspread

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class GoogleSheetsError(Exception):
    """Googleスプレッドシートへの転記に失敗した場合のエラー。"""


def append_rows(
    spreadsheet_id: str,
    sheet_name: str,
    oauth_client_path: str,
    oauth_token_path: str,
    rows: list[list],
) -> int:
    """rows をシートの末尾に追記する。追記できた件数を返す。"""
    if not rows:
        return 0

    try:
        client = gspread.oauth(
            scopes=SCOPES,
            credentials_filename=oauth_client_path,
            authorized_user_filename=oauth_token_path,
        )
        spreadsheet = client.open_by_key(spreadsheet_id)
        worksheet = spreadsheet.worksheet(sheet_name)
        worksheet.append_rows(rows, value_input_option="USER_ENTERED")
    except Exception as e:
        raise GoogleSheetsError(f"Googleスプレッドシートへの転記に失敗しました: {e}") from e

    return len(rows)
