"""詳細画面から実際に読み取れる内容を確認するための診断スクリプト。

実行方法（プロジェクトフォルダで）:
    python scripts\\stage3_debug_detail.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from playwright.sync_api import sync_playwright

from src.config import load_config
from src.hokokkun import get_sales_list_items, login_to_hokokkun, open_sales_detail, open_sales_list

OUTPUT_PATH = Path(__file__).resolve().parent.parent / "temp" / "detail_debug.txt"


def main() -> int:
    config = load_config()
    lines: list[str] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        login_to_hokokkun(page, config)
        open_sales_list(page, "当日")
        items = get_sales_list_items(page)
        lines.append(f"一覧件数: {len(items)}")
        if not items:
            lines.append("一覧が0件でした。")
        else:
            open_sales_detail(page, items[0])

            lines.append(f"現在のURL: {page.url}")
            lines.append(f"page.get_by_role('row').count() = {page.get_by_role('row').count()}")
            lines.append(f"page.get_by_role('cell').count() = {page.get_by_role('cell').count()}")
            lines.append(f"iframeの数: {len(page.frames)}")
            for i, frame in enumerate(page.frames):
                lines.append(f"  frame[{i}] url={frame.url} rows={frame.get_by_role('row').count()}")

            lines.append("---- page.get_by_role('row') の中身 ----")
            rows = page.get_by_role("row")
            for row_index in range(rows.count()):
                cells = rows.nth(row_index).get_by_role("cell")
                cell_texts = [cells.nth(c).inner_text().strip().replace("\n", "\\n") for c in range(cells.count())]
                lines.append(f"row[{row_index}] cells={cells.count()} : {cell_texts}")

            lines.append("---- page.inner_text('body') の先頭2000文字 ----")
            lines.append(page.locator("body").inner_text()[:2000])

        input("確認できたら、このウィンドウでEnterキーを押してください...")
        browser.close()

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text("\n".join(lines), encoding="utf-8")
    print(f"結果を書き出しました: {OUTPUT_PATH}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
