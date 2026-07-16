# -*- coding: utf-8 -*-
"""契約数更新(NG含む)自動化 メイン処理。

処理手順(要件どおり):
  1. ほうこっくんへログイン
  2. 営業報告データベースExcelを読み込み
  3. 「前日」を選択
  4. 項目一覧から「営業」を選択して一覧を取得
  5. 一覧を1件ずつ開いて必要項目を抽出
  6. Excel内に同じ会社名が登録済みか確認(重複判定は会社名のみ)
  7. 未登録の会社のみExcelへ追記(結果がNGでも転記対象)
  8. 「次日」で当日に切り替えて同じ処理
  9. Excelを上書き保存
 10. 今回新規追加したデータのみGoogleスプレッドシートへ転記

実行結果・追加件数・重複スキップ件数・取得失敗件数を logs/ へ保存する。
Windowsタスクスケジューラから run_sync.bat 経由で毎日9:30に起動する。
"""

import configparser
import logging
import os
import sys
from dataclasses import dataclass
from datetime import datetime

from excel_db import ExcelDatabase, ExcelSaveLockedError
from houkokkun import HoukokkunClient
from sheets_transfer import transfer_new_rows

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
logger = logging.getLogger(__name__)


@dataclass
class SyncStats:
    added: int = 0
    skipped: int = 0
    failed: int = 0


def setup_logging():
    log_dir = os.path.join(BASE_DIR, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"sync_{datetime.now():%Y%m%d_%H%M%S}.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    return log_path


def load_config():
    path = os.path.join(BASE_DIR, "config.ini")
    if not os.path.exists(path):
        raise RuntimeError(
            "config.ini が見つかりません。config.example.ini をコピーして設定してください。"
        )
    config = configparser.ConfigParser()
    config.read(path, encoding="utf-8")
    return config


def process_day(client, excel, stats, added_records, day_label):
    """表示中の日の一覧を取得し、重複を除いてExcelへ追記する。"""
    reports = client.fetch_reports(stats, day_label)
    for record in reports:
        company = record.get("会社名", "")
        if not company:
            # 会社名が取れないと重複判定できないため取得失敗として記録する
            logger.error("[%s] 会社名が空のためExcelへ追記できません: %s", day_label, record)
            stats.failed += 1
            continue
        if excel.has_company(company):
            logger.info("[%s] 重複スキップ: %s", day_label, company)
            stats.skipped += 1
            continue
        excel.append(record)
        added_records.append(record)
        stats.added += 1
        logger.info("[%s] Excelへ追記: %s (結果: %s)", day_label, company, record.get("結果", ""))


def run():
    log_path = setup_logging()
    logger.info("契約数更新(NG含む)処理を開始します。")
    stats = SyncStats()
    added_records = []
    exit_code = 0

    try:
        config = load_config()

        # 2. Excel読み込み(保存できない状態を早期検知するため先に読み込む)
        excel = ExcelDatabase(
            config["excel"]["path"],
            config["excel"].get("sheet_name", "").strip(),
        )

        # 1. ログイン → 3-8. 前日・当日の取得と追記
        with HoukokkunClient(config) as client:
            client.login()
            client.open_sales_list()
            client.go_previous_day()
            process_day(client, excel, stats, added_records, "前日")
            client.go_next_day()
            process_day(client, excel, stats, added_records, "当日")

        # 9. Excel上書き保存
        try:
            excel.save()
        except ExcelSaveLockedError:
            logger.exception("Excelの保存に失敗しました。スプレッドシートへの転記も中止します。")
            return 1

        # 10. 今回新規追加したデータのみスプレッドシートへ転記
        try:
            transfer_new_rows(config, added_records)
        except Exception:
            logger.exception("スプレッドシートへの転記に失敗しました。Excelへの追記は完了しています。")
            exit_code = 1

    except Exception:
        logger.exception("処理中に致命的なエラーが発生しました。")
        exit_code = 1
    finally:
        logger.info(
            "実行結果: %s / 追加 %d 件 / 重複スキップ %d 件 / 取得失敗 %d 件",
            "正常終了" if exit_code == 0 else "エラーあり",
            stats.added,
            stats.skipped,
            stats.failed,
        )
        logger.info("ログ: %s", log_path)

    return exit_code


if __name__ == "__main__":
    sys.exit(run())
