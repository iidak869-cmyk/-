"""会社名の重複判定用の正規化処理。

- 前後の空白を削除
- 全角スペースと半角スペースを統一
- 連続した空白を1つに統一
法人格（株式会社など）は削除しない。
"""
import re

_WHITESPACE_RE = re.compile(r"[　\s]+")


def normalize_company_name(name: str) -> str:
    if name is None:
        return ""
    return _WHITESPACE_RE.sub(" ", str(name)).strip()
