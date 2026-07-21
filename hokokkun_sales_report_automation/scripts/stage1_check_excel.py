"""第1段階：Excelが正しく読み込めるかを確認するスクリプト。

実行方法（プロジェクトフォルダで）:
    python scripts\\stage1_check_excel.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import ConfigError, load_config
from src.excel_io import (
    DELIVERY_ITEM_COLUMNS,
    ExcelLoadError,
    find_last_data_row,
    get_existing_company_names,
    get_sheet,
    load_workbook_safe,
)

EXPECTED_HEADERS = [
    "契約日", "商談時間", None, None, "営業方法", "会社名", "結果", "NG理由①", "NG理由②",
    "会社所在地", "代表者名", "種類", "営業担当", "同行", "アポ担当", "業種", "契約形態", "金額",
    "ネット1", "ネット2", "HP", "動画サイト", "LP", "LINE", "動画CR", "Cyteki", "クレカ所持", "初期回収",
]

DELIVERY_ITEM_NAMES = ["ネット1", "ネット2", "HP", "動画サイト", "LP", "LINE", "動画CR", "Cyteki"]


def main() -> int:
    print("=== 第1段階: Excel読み込み確認 ===")
    try:
        config = load_config()
    except ConfigError as e:
        print(f"[NG] 設定エラー: {e}")
        return 1

    print(f"対象ファイル: {config.excel_file_path}")
    print(f"対象シート  : {config.excel_sheet_name}")

    try:
        wb = load_workbook_safe(config.excel_file_path)
        ws = get_sheet(wb, config.excel_sheet_name)
    except ExcelLoadError as e:
        print(f"[NG] {e}")
        return 1

    print(f"[OK] Excelを読み込めました。シート一覧: {wb.sheetnames}")

    headers = [ws.cell(row=1, column=c).value for c in range(1, len(EXPECTED_HEADERS) + 1)]
    if headers == EXPECTED_HEADERS:
        print("[OK] 見出し行(1行目)は想定通りです。")
    else:
        print("[注意] 見出し行が想定と異なります。")
        print(f"  想定: {EXPECTED_HEADERS}")
        print(f"  実際: {headers}")

    last_data_row = find_last_data_row(ws)
    data_row_count = max(0, last_data_row - 1)
    print(f"最終データ行: {last_data_row}行目（データ件数: {data_row_count}件）")
    print(f"次に追記する行: {last_data_row + 1}行目")

    company_names = get_existing_company_names(ws, last_data_row)
    print(f"登録済み会社名（重複判定用に正規化後）の件数: {len(company_names)}件")

    if data_row_count == 0:
        print()
        print("[確認事項] データベースシートにデータが1件もありません。")
        print("  そのため、納品物件(S列〜Z列: ネット1/ネット2/HP/動画サイト/LP/LINE/動画CR/Cyteki)を")
        print("  どのように入力するか（例: 該当列に「○」を入れる／商品名をそのまま入れる 等）を")
        print("  既存データから判断できませんでした。実データが入っているExcelがあれば、")
        print("  該当行を数件教えてください。")
    else:
        print()
        print("納品物件列(S〜Z列)のサンプル（末尾から最大10件）:")
        sample_rows = list(range(max(2, last_data_row - 9), last_data_row + 1))
        for row in sample_rows:
            values = [ws.cell(row=row, column=c).value for c in DELIVERY_ITEM_COLUMNS]
            pairs = ", ".join(f"{name}={val!r}" for name, val in zip(DELIVERY_ITEM_NAMES, values))
            print(f"  {row}行目: {pairs}")

    print()
    print("=== 確認終了 ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
