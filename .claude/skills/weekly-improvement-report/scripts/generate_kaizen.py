"""
業務改善指示書 Excel生成スクリプト

テンプレート（template_kaizen.xlsx）をコピーして内容を差し替える方式。
フォーマット・テーマカラーはテンプレートをそのまま継承する。

Usage:
    python generate_kaizen.py --date "2026年8月6日" \
        --incident "..." --risk "..." --cause "..." \
        --improvement "..." --action "..." \
        --output "業務改善指示書_20260806.xlsx"
"""

import argparse
import shutil
import os
from datetime import date, datetime
import openpyxl
from openpyxl.styles import Font, Alignment

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SCRIPT_DIR, 'template_kaizen.xlsx')

FONT_NAME = 'メイリオ'

# Content cells in the template (top-left of each merged range)
# ①  A4:J10  – 発生事象
# ②  A12:J17 – 想定リスク（ガイド質問込み）
# ③  A19:J26 – 問題点
# ④  A28:J34 – 改善策
# ⑤  A36:J39 – 対策・行動
CONTENT_CELLS = {
    'incident':    'A4',
    'risk':        'A12',
    'cause':       'A19',
    'improvement': 'A28',
    'action':      'A36',
}

# ②〜⑤ にはガイド質問を自動付加する（テンプレートの書式に合わせた定型文）
GUIDE_QUESTIONS = {
    'risk': (
        "【今回の事例によって、実際にどのような影響や問題が発生しましたか？"
        "もしくは今回の事例が放置されていた場合、どのような影響や問題が発生していた可能性がありますか？】\n"
        "{risk_answer}\n\n"
        "【今回の対応でリスクは解消されましたか？　今後も同様のリスクが発生する可能性は残っていますか？】\n"
        "{risk_future}"
    ),
    'cause': (
        "【どのような仕組み・ルール・体制の不足が背景にあったと考えられますか？】\n"
        "{cause_structure}\n\n"
        "【情報共有・教育・確認体制などに課題はありませんでしたか？】\n"
        "{cause_info}"
    ),
    'improvement': (
        "【具体的にどのような運用手順や仕組みの見直しが考えられますか？】\n"
        "{improvement_operation}\n\n"
        "【現場で実施可能な再発防止策は何がありますか？】\n"
        "{improvement_prevention}"
    ),
    'action': (
        "【自分では判断が難しい課題や、他部署の協力が必要なことはありますか？】\n"
        "{action_external}\n\n"
        "【長期的に仕組みや方針を見直すべき部分はありますか？】\n"
        "{action_longterm}"
    ),
}


def _set_content(ws, cell_ref, text, size=9):
    cell = ws[cell_ref]
    cell.value = text
    cell.font = Font(name=FONT_NAME, size=size)
    cell.alignment = Alignment(wrap_text=True, horizontal='left', vertical='top')


def create_kaizen_excel(creation_date, incident,
                        risk_answer, risk_future,
                        cause_structure, cause_info,
                        improvement_operation, improvement_prevention,
                        action_external, action_longterm,
                        output_path):

    # テンプレートをコピー
    shutil.copy2(TEMPLATE_PATH, output_path)

    wb = openpyxl.load_workbook(output_path)
    ws = wb.active

    # 作成日
    ws['B2'].value = creation_date
    ws['B2'].font = Font(name=FONT_NAME, size=9)
    ws['B2'].alignment = Alignment(horizontal='center', vertical='center')

    # ① 発生事象（ガイドなし、ユーザー文章をそのまま）
    _set_content(ws, 'A4', incident)

    # ② 想定リスク（ガイド付き2問）
    risk_text = (
        "【今回の事例によって、実際にどのような影響や問題が発生しましたか？"
        "もしくは今回の事例が放置されていた場合、どのような影響や問題が発生していた可能性がありますか？】\n"
        f"{risk_answer}\n\n"
        "【今回の対応でリスクは解消されましたか？　今後も同様のリスクが発生する可能性は残っていますか？】\n"
        f"{risk_future}"
    )
    _set_content(ws, 'A12', risk_text)

    # ③ 問題点（ガイド付き2問）
    cause_text = (
        "【どのような仕組み・ルール・体制の不足が背景にあったと考えられますか？】\n"
        f"{cause_structure}\n\n"
        "【情報共有・教育・確認体制などに課題はありませんでしたか？】\n"
        f"{cause_info}"
    )
    _set_content(ws, 'A19', cause_text)

    # ④ 改善策（ガイド付き2問）
    improvement_text = (
        "【具体的にどのような運用手順や仕組みの見直しが考えられますか？】\n"
        f"{improvement_operation}\n\n"
        "【現場で実施可能な再発防止策は何がありますか？】\n"
        f"{improvement_prevention}"
    )
    _set_content(ws, 'A28', improvement_text)

    # ⑤ 対策・行動（ガイド付き2問）
    action_text = (
        "【自分では判断が難しい課題や、他部署の協力が必要なことはありますか？】\n"
        f"{action_external}\n\n"
        "【長期的に仕組みや方針を見直すべき部分はありますか？】\n"
        f"{action_longterm}"
    )
    _set_content(ws, 'A36', action_text)

    wb.save(output_path)
    print(f"Saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='業務改善指示書 Excel生成（新フォーマット）')
    parser.add_argument('--date',                   required=True)
    parser.add_argument('--incident',               required=True, help='①発生事象（構造化テキスト）')
    parser.add_argument('--risk-answer',            required=True, help='②影響・リスク内容')
    parser.add_argument('--risk-future',            required=True, help='②リスク解消・再発可能性')
    parser.add_argument('--cause-structure',        required=True, help='③仕組み・体制の不足')
    parser.add_argument('--cause-info',             required=True, help='③情報共有・教育の課題')
    parser.add_argument('--improvement-operation',  required=True, help='④運用手順の見直し')
    parser.add_argument('--improvement-prevention', required=True, help='④現場での再発防止策')
    parser.add_argument('--action-external',        required=True, help='⑤他部署協力が必要なこと')
    parser.add_argument('--action-longterm',        required=True, help='⑤長期的な方針見直し')
    parser.add_argument('--output',                 required=True)
    args = parser.parse_args()

    create_kaizen_excel(
        creation_date=args.date,
        incident=args.incident,
        risk_answer=args.risk_answer,
        risk_future=args.risk_future,
        cause_structure=args.cause_structure,
        cause_info=args.cause_info,
        improvement_operation=args.improvement_operation,
        improvement_prevention=args.improvement_prevention,
        action_external=args.action_external,
        action_longterm=args.action_longterm,
        output_path=args.output,
    )


if __name__ == '__main__':
    main()
