"""第2段階：login_to_hokokkun関数が正しく動くか確認するスクリプト。

実行方法（プロジェクトフォルダで）:
    python scripts\\stage2_test_login.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from playwright.sync_api import sync_playwright

from src.config import ConfigError, load_config
from src.hokokkun import HokokkunError, login_to_hokokkun


def main() -> int:
    print("=== 第2段階: ログイン確認 ===")
    try:
        config = load_config()
    except ConfigError as e:
        print(f"[NG] 設定エラー: {e}")
        return 1

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()
        try:
            login_to_hokokkun(page, config)
        except HokokkunError as e:
            print(f"[NG] {e}")
            browser.close()
            return 1

        print(f"[OK] ログインに成功しました。現在のURL: {page.url}")
        print("ブラウザに表示されている画面が、いつものログイン後トップ画面と同じか確認してください。")
        input("確認できたら、このウィンドウで Enter キーを押すとブラウザを閉じます...")
        browser.close()

    print("=== 確認終了 ===")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
