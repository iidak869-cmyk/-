"""Googleスプレッドシートへの転記だけを再実行するスクリプト。

Excelへの保存は成功したがGoogleスプレッドシートへの転記だけ失敗した場合に、
temp フォルダに残っている一時JSON（added_details_*.json）を使って
転記だけをやり直す。

実行方法（プロジェクトフォルダで）:
    python scripts\\retry_sheets.py
"""
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src import google_sheets
from src.config import ConfigError, load_config
from src.sheet_mapping import build_sheet_row


def main() -> int:
    try:
        config = load_config()
    except ConfigError as e:
        print(f"[NG] 設定エラー: {e}")
        return 1

    if not config.enable_google_sheets:
        print("[NG] .env の ENABLE_GOOGLE_SHEETS が false のままです。true にしてから実行してください。")
        return 1

    temp_files = sorted(config.temp_directory.glob("added_details_*.json"))
    if not temp_files:
        print("転記対象の一時JSONが見つかりませんでした（temp フォルダ）。再実行の必要はありません。")
        return 0

    for temp_path in temp_files:
        with open(temp_path, encoding="utf-8") as f:
            added_details = json.load(f)

        print(f"{temp_path.name} ({len(added_details)}件) を転記します...")
        try:
            sheet_rows = [build_sheet_row(detail) for detail in added_details]
            transferred = google_sheets.append_rows(
                config.google_spreadsheet_id,
                config.google_sheet_name,
                config.google_credentials_path,
                config.google_oauth_token_path,
                sheet_rows,
            )
            print(f"[OK] {transferred}件を転記しました。")
            temp_path.unlink()
        except google_sheets.GoogleSheetsError as e:
            print(f"[NG] {e}")
            return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
