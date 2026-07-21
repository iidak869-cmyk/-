"""ほうこっくん→Excel の一連の処理をまとめる。"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright

from . import excel_io, hokokkun
from .config import Config
from .logger import (
    RESULT_FAILED,
    RESULT_PARTIAL_SUCCESS,
    RESULT_SUCCESS,
    RunLog,
)
from .mapping import build_excel_row_values
from .normalize import normalize_company_name

DAYS_TO_PROCESS = ["前日", "当日"]


def open_detail_with_retry(page, target_day: str, item: dict, run_log: RunLog) -> None:
    """一覧から詳細を開く。不正アクセスエラーの場合は1回だけ再試行する。"""
    hokokkun.open_sales_detail(page, item)
    if not hokokkun.is_unauthorized_access(page):
        return

    company_name = item.get("company_name", "")
    run_log.add_retry(f"不正アクセスエラーを検知したため再試行します: {company_name}")

    hokokkun.open_sales_list(page, target_day)
    items = hokokkun.get_sales_list_items(page)
    retry_item = items[item["row_index"]] if item["row_index"] < len(items) else item
    hokokkun.open_sales_detail(page, retry_item)

    if hokokkun.is_unauthorized_access(page):
        run_log.add_retry(f"再試行しましたが不正アクセスエラーが再度表示されました: {company_name}")
        hokokkun.open_sales_list(page, target_day)
        raise hokokkun.UnauthorizedAccessError(
            f"不正アクセスエラーが再試行後も表示されました: {company_name}"
        )

    run_log.add_retry(f"再試行に成功しました: {company_name}")


def process_day(page, target_day: str, known_companies: set[str], run_log: RunLog) -> list[dict]:
    """1日分（前日 or 当日）の営業報告を処理し、新規追加分の行データを返す。"""
    hokokkun.open_sales_list(page, target_day)
    items = hokokkun.get_sales_list_items(page)

    if target_day == "前日":
        run_log.previous_day_count = len(items)
    else:
        run_log.current_day_count = len(items)

    new_rows: list[dict] = []

    for item in items:
        company_name_hint = item.get("company_name", "")
        try:
            open_detail_with_retry(page, target_day, item, run_log)
        except hokokkun.UnauthorizedAccessError as e:
            run_log.detail_failed_count += 1
            run_log.add_error(str(e), company_name_hint)
            continue
        except Exception as e:
            run_log.detail_failed_count += 1
            run_log.add_error(f"詳細画面を開けませんでした: {e}", company_name_hint)
            continue

        try:
            detail = hokokkun.extract_sales_detail(page)
        except hokokkun.CompanyNameNotFoundError as e:
            run_log.company_name_failed_count += 1
            run_log.add_error(str(e), company_name_hint)
            hokokkun.return_to_sales_list(page)
            continue
        except Exception as e:
            run_log.detail_failed_count += 1
            run_log.add_error(f"詳細取得に失敗しました: {e}", company_name_hint)
            hokokkun.return_to_sales_list(page)
            continue

        normalized = normalize_company_name(detail.get("company_name", ""))
        if normalized in known_companies:
            run_log.duplicate_skipped_count += 1
            hokokkun.return_to_sales_list(page)
            continue

        amount = item.get("amount", "")
        detail["amount"] = "" if amount == "-" else amount

        row_values = build_excel_row_values(detail)
        new_rows.append(row_values)
        known_companies.add(normalized)

        hokokkun.return_to_sales_list(page)

    return new_rows


def save_temp_added_rows(temp_directory: Path, added_rows: list[dict]) -> Path | None:
    if not added_rows:
        return None
    temp_directory.mkdir(parents=True, exist_ok=True)
    path = temp_directory / f"added_rows_{datetime.now():%Y%m%d_%H%M%S}.json"
    serializable = [{str(col): value for col, value in row.items()} for row in added_rows]
    with open(path, "w", encoding="utf-8") as f:
        json.dump(serializable, f, ensure_ascii=False, indent=2)
    return path


def run(config: Config) -> str:
    run_log = RunLog(log_directory=config.log_directory)

    try:
        wb = excel_io.load_workbook_safe(config.excel_file_path)
        ws = excel_io.get_sheet(wb, config.excel_sheet_name)
    except excel_io.ExcelLoadError as e:
        run_log.add_error(str(e))
        run_log.write(RESULT_FAILED)
        return RESULT_FAILED

    last_data_row = excel_io.find_last_data_row(ws)
    known_companies = excel_io.get_existing_company_names(ws, last_data_row)

    all_new_rows: list[dict] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        try:
            hokokkun.login_to_hokokkun(page, config)
        except hokokkun.LoginError as e:
            run_log.add_error(str(e))
            browser.close()
            run_log.write(RESULT_FAILED)
            return RESULT_FAILED

        for target_day in DAYS_TO_PROCESS:
            try:
                day_rows = process_day(page, target_day, known_companies, run_log)
                all_new_rows.extend(day_rows)
            except Exception as e:
                run_log.add_error(f"{target_day}の営業一覧を表示できませんでした: {e}")
                browser.close()
                run_log.write(RESULT_FAILED)
                return RESULT_FAILED

        browser.close()

    style_source_row = last_data_row if last_data_row > excel_io.HEADER_ROW else None
    next_row = last_data_row + 1
    for row_values in all_new_rows:
        excel_io.append_row(ws, next_row, row_values, style_source_row)
        style_source_row = next_row
        next_row += 1

    try:
        excel_io.save_workbook_safe(wb, config.excel_file_path)
    except excel_io.ExcelSaveError as e:
        run_log.excel_save_result = f"失敗: {e}"
        run_log.add_error(str(e))
        run_log.write(RESULT_FAILED)
        return RESULT_FAILED

    run_log.excel_save_result = "成功"
    run_log.excel_added_count = len(all_new_rows)

    save_temp_added_rows(config.temp_directory, all_new_rows)

    if config.enable_google_sheets:
        run_log.sheet_result = "未実装（第4段階で実装予定）"
    else:
        run_log.sheet_result = "無効化（第4段階で有効化予定）"

    if run_log.detail_failed_count or run_log.company_name_failed_count:
        overall_result = RESULT_PARTIAL_SUCCESS
    else:
        overall_result = RESULT_SUCCESS

    run_log.write(overall_result)
    return overall_result
