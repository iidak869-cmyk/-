# -*- coding: utf-8 -*-
"""営業報告の抽出項目定義。

Excelヘッダー・スプレッドシートヘッダー・詳細画面のラベルの
対応をここで一元管理する。
"""

# Excel / スプレッドシートへ転記する項目(この順に列を並べる)
FIELDS = [
    "契約日",
    "商談時間",
    "営業方法",
    "会社名",
    "結果",
    "NG理由①",
    "会社所在地",
    "代表者名",
    "種類",
    "営業担当",
    "同行",
    "アポ担当",
    "業種",
    "契約形態",
    "金額",
    "納品物件",
    "クレカ所持",
    "初期回収",
]

# ほうこっくんの詳細画面に表示されるラベル → FIELDS の項目名 の読み替え表。
# ほうこっくん上のNG理由は1項目のみなので「NG理由①」へ転記する。
# (「NG理由②」は使用しないため転記対象に含めない)
# 画面上のラベルが FIELDS と一致しない場合はここに追記して調整する。
LABEL_ALIASES = {
    "NG理由": "NG理由①",
    "NG理由1": "NG理由①",
    "NG理由１": "NG理由①",
    "所在地": "会社所在地",
    "住所": "会社所在地",
    "代表者": "代表者名",
    "担当": "営業担当",
    "営業": "営業担当",
    "アポ": "アポ担当",
}


def normalize_label(label: str) -> str:
    """画面ラベルの表記ゆれ(空白・コロンなど)を吸収する。"""
    if label is None:
        return ""
    cleaned = "".join(label.split())
    return cleaned.rstrip(":：").strip()


def label_to_field(label: str):
    """詳細画面のラベルを FIELDS の項目名へ変換する。対象外なら None。"""
    key = normalize_label(label)
    field = LABEL_ALIASES.get(key, key)
    return field if field in FIELDS else None
