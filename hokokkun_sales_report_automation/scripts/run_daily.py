"""日次実行スクリプト（第5段階でWindowsタスクスケジューラへ登録する）。

実行方法（プロジェクトフォルダで）:
    python scripts\\run_daily.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.config import ConfigError, load_config
from src.pipeline import run


def main() -> int:
    try:
        config = load_config()
    except ConfigError as e:
        print(f"[NG] 設定エラー: {e}")
        return 1

    result = run(config)
    print(f"実行結果: {result}")
    return 0 if result in ("SUCCESS", "PARTIAL_SUCCESS") else 1


if __name__ == "__main__":
    raise SystemExit(main())
