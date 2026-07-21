""".env から設定を読み込む。"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"


class ConfigError(Exception):
    """設定ファイルが存在しない、または必須項目が不足している場合のエラー。"""


@dataclass
class Config:
    hokokkun_url: str
    hokokkun_login_id: str
    hokokkun_password: str
    excel_file_path: str
    excel_sheet_name: str
    enable_google_sheets: bool
    google_spreadsheet_id: str
    google_sheet_name: str
    google_credentials_path: str
    log_directory: Path
    temp_directory: Path
    browser_headless: bool


def load_config() -> Config:
    if not ENV_PATH.exists():
        raise ConfigError(
            f".envファイルが見つかりません: {ENV_PATH}\n"
            ".env.example をコピーして .env を作成し、値を設定してください。"
        )

    load_dotenv(ENV_PATH)

    excel_file_path = os.getenv("EXCEL_FILE_PATH", "").strip()
    excel_sheet_name = os.getenv("EXCEL_SHEET_NAME", "").strip()
    if not excel_file_path or not excel_sheet_name:
        raise ConfigError(
            "EXCEL_FILE_PATH または EXCEL_SHEET_NAME が .env に設定されていません。"
        )

    log_directory = PROJECT_ROOT / os.getenv("LOG_DIRECTORY", "logs")
    temp_directory = PROJECT_ROOT / os.getenv("TEMP_DIRECTORY", "temp")
    log_directory.mkdir(parents=True, exist_ok=True)
    temp_directory.mkdir(parents=True, exist_ok=True)

    return Config(
        hokokkun_url=os.getenv("HOKOKKUN_URL", "").strip(),
        hokokkun_login_id=os.getenv("HOKOKKUN_LOGIN_ID", "").strip(),
        hokokkun_password=os.getenv("HOKOKKUN_PASSWORD", "").strip(),
        excel_file_path=excel_file_path,
        excel_sheet_name=excel_sheet_name,
        enable_google_sheets=os.getenv("ENABLE_GOOGLE_SHEETS", "false").strip().lower() == "true",
        google_spreadsheet_id=os.getenv("GOOGLE_SPREADSHEET_ID", "").strip(),
        google_sheet_name=os.getenv("GOOGLE_SHEET_NAME", "転記データ").strip(),
        google_credentials_path=os.getenv("GOOGLE_CREDENTIALS_PATH", "").strip(),
        log_directory=log_directory,
        temp_directory=temp_directory,
        browser_headless=os.getenv("BROWSER_HEADLESS", "true").strip().lower() != "false",
    )
