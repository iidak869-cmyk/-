"""実行ログの記録。"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

RESULT_SUCCESS = "SUCCESS"
RESULT_PARTIAL_SUCCESS = "PARTIAL_SUCCESS"
RESULT_EXCEL_SUCCESS_SHEET_FAILED = "EXCEL_SUCCESS_SHEET_FAILED"
RESULT_FAILED = "FAILED"


@dataclass
class RunLog:
    log_directory: Path
    start_time: dt.datetime = field(default_factory=dt.datetime.now)
    previous_day_count: int = 0
    current_day_count: int = 0
    excel_added_count: int = 0
    sheet_transferred_count: int = 0
    duplicate_skipped_count: int = 0
    detail_failed_count: int = 0
    company_name_failed_count: int = 0
    excel_save_result: str = "未実施"
    sheet_result: str = "未実施（第4段階で有効化予定）"
    errors: list[str] = field(default_factory=list)
    error_companies: list[str] = field(default_factory=list)
    retries: list[str] = field(default_factory=list)

    def add_error(self, message: str, company_name: str = "") -> None:
        self.errors.append(message)
        if company_name:
            self.error_companies.append(company_name)

    def add_retry(self, message: str) -> None:
        self.retries.append(message)

    def write(self, overall_result: str) -> Path:
        end_time = dt.datetime.now()
        self.log_directory.mkdir(parents=True, exist_ok=True)
        log_path = self.log_directory / f"contract_update_{self.start_time:%Y%m%d}.log"

        lines = [
            f"実行結果: {overall_result}",
            f"実行開始日時: {self.start_time:%Y-%m-%d %H:%M:%S}",
            f"実行終了日時: {end_time:%Y-%m-%d %H:%M:%S}",
            f"前日分の一覧件数: {self.previous_day_count}",
            f"当日分の一覧件数: {self.current_day_count}",
            f"Excelへの追加件数: {self.excel_added_count}",
            f"Googleスプレッドシートへの転記件数: {self.sheet_transferred_count}",
            f"重複スキップ件数: {self.duplicate_skipped_count}",
            f"詳細取得失敗件数: {self.detail_failed_count}",
            f"会社名取得失敗件数: {self.company_name_failed_count}",
            f"Excel保存結果: {self.excel_save_result}",
            f"Googleスプレッドシート転記結果: {self.sheet_result}",
            "再試行:",
        ]
        if self.retries:
            lines.extend(f"  - {r}" for r in self.retries)
        else:
            lines.append("  (なし)")
        lines.append("エラー:")
        if self.errors:
            lines.extend(f"  - {e}" for e in self.errors)
        else:
            lines.append("  (なし)")
        lines.append("エラーが発生した会社名:")
        if self.error_companies:
            lines.extend(f"  - {c}" for c in self.error_companies)
        else:
            lines.append("  (なし)")

        with open(log_path, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n" + ("=" * 40) + "\n")

        return log_path
